# -*- coding: utf-8 -*-
"""
Moteur de requêtes Ponctualité
=======================================================
Un seul moteur, deux usages : l'application Streamlit (app_requetes.py)
et le notebook Colab (requetes_ponctualite.ipynb).

ENTRÉE : 2 fichiers Excel (ordre indifférent, reconnus par leurs colonnes)
  - Pointages   : DATE_POINTAGE, MATRICULE, Conducteur, Agent
  - Ponctualité : Date, Ligne, Sens, Point d'arrêt, Conducteur (= code agent/service),
                  Heure_départ_planifiée, Heure_départ_mesurée, ...

MODÈLE : tout est ramené à UNE table plate `passages` (vue filtrée `p`) :
  jour, mois, semaine, jour_num, jour_sem, heure, tranche, ligne, sens, arret,
  course, service, vehicule, conducteur, heure_prevue, heure_mesuree, retard (min)

  retard = départ mesuré - départ planifié (minutes, correction du passage de minuit)
  À l'heure = retard entre `avance` (-1) et `seuil` (+5) ; Retard > seuil ; Avance < avance.
"""
import io
import math
import re
import sqlite3
import threading
import unicodedata
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

SEUIL_RETARD = 5.0
SEUIL_AVANCE = -1.0
JOURS = {1: "Lundi", 2: "Mardi", 3: "Mercredi", 4: "Jeudi", 5: "Vendredi", 6: "Samedi", 7: "Dimanche"}
ORDRE_JOURS = list(JOURS.values())


# ----------------------------------------------------------------------------
# 1. LECTURE ET PRÉPARATION DES 2 FICHIERS
# ----------------------------------------------------------------------------
def _cle(nom):
    """Clé de comparaison de colonne : sans accents, sans crochets/espaces, minuscule."""
    s = unicodedata.normalize("NFKD", str(nom))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]", "", s.lower())


def _col(df, *noms):
    cles = {_cle(c): c for c in df.columns}
    for n in noms:
        if _cle(n) in cles:
            return cles[_cle(n)]
    return None


def _lire(nom, src):
    if isinstance(src, (bytes, bytearray)):
        src = io.BytesIO(src)
    if str(nom).lower().endswith(".csv"):
        return pd.read_csv(src, sep=None, engine="python")
    return pd.read_excel(src)


def _type_fichier(df):
    cles = {_cle(c) for c in df.columns}
    if "datepointage" in cles and "agent" in cles:
        return "pointages"
    if "heuredepartmesuree" in cles or "ponctualite" in cles:
        return "ponctualite"
    return None


def _code(s):
    """Code agent/service normalisé : texte, sans espaces, sans '.0' final, minuscule."""
    return s.astype(str).str.strip().str.lower().str.replace(r"\.0$", "", regex=True)


def _txt(s):
    return s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True).replace({"nan": np.nan, "None": np.nan, "": np.nan})


def _minutes(s):
    """Heure ('HH:MM[:SS]', time ou datetime Excel) -> minutes depuis minuit."""
    t = s.astype(str).str.extract(r"(\d{1,2}:\d{2}(?::\d{2})?)")[0]
    t = t.where(t.str.count(":") != 1, t + ":00")
    return pd.to_timedelta(t, errors="coerce").dt.total_seconds() / 60


def _hhmm(minutes):
    m = (minutes % 1440).fillna(-1).astype(int)
    return m.map(lambda x: np.nan if x < 0 else f"{x // 60:02d}:{x % 60:02d}")


def _tranche(h):
    bornes = [(0, 6, "00-06h"), (6, 9, "06-09h"), (9, 12, "09-12h"), (12, 14, "12-14h"),
              (14, 16, "14-16h"), (16, 19, "16-19h"), (19, 24, "19-24h")]
    out = pd.Series(np.nan, index=h.index, dtype=object)
    for a, b, lib in bornes:
        out[(h >= a) & (h < b)] = lib
    return out


def preparer(pointages, ponctualite, anonymiser=False):
    """Retourne (table `passages`, DataFrame de contrôle du rapprochement)."""
    # --- Pointages : qui conduisait quel service (agent) quel jour
    c_date, c_cond, c_agent = (_col(pointages, "DATE_POINTAGE"), _col(pointages, "Conducteur"),
                               _col(pointages, "Agent"))
    if not (c_date and c_cond and c_agent):
        raise ValueError("Fichier de pointages : colonnes DATE_POINTAGE, Conducteur et Agent attendues.")
    pt = pd.DataFrame({
        "jour": pd.to_datetime(pointages[c_date], errors="coerce", dayfirst=True).dt.normalize(),
        "agent": _code(pointages[c_agent]),
        "conducteur": pointages[c_cond].astype(str).str.strip(),
    }).dropna(subset=["jour"]).drop_duplicates()

    if anonymiser:   # « Conducteur 001 »… : ordre aléatoire à chaque chargement, non réversible
        noms = pt["conducteur"].unique()
        ordre = np.random.default_rng().permutation(len(noms)) + 1
        pt["conducteur"] = pt["conducteur"].map({n: f"Conducteur {k:03d}" for n, k in zip(noms, ordre)})

    # --- Ponctualité : passages mesurés ; sa colonne « Conducteur » contient le code agent/service
    c = {k: _col(ponctualite, *v) for k, v in {
        "date": ("Date",), "service": ("Conducteur", "Agent"), "ligne": ("Ligne",), "sens": ("Sens",),
        "arret": ("Point d'arrêt", "Point d arret"), "course": ("Course",),
        "vehicule": ("Véhicule", "Vehicule"),
        "prev": ("Heure_départ_planifiée", "Heure_depart_planifiee"),
        "mes": ("Heure_départ_mesurée", "Heure_depart_mesuree")}.items()}
    manque = [k for k in ("date", "service", "prev", "mes") if not c[k]]
    if manque:
        raise ValueError("Fichier de ponctualité : colonnes introuvables pour " + ", ".join(manque))

    def val(k):
        return ponctualite[c[k]] if c[k] else pd.Series(np.nan, index=ponctualite.index)

    prev, mes = _minutes(val("prev")), _minutes(val("mes"))
    retard = (mes - prev).round(1)
    retard = np.where(retard < -720, retard + 1440, np.where(retard > 720, retard - 1440, retard))
    dt = pd.to_datetime(val("date"), errors="coerce", dayfirst=True).dt.normalize()
    heure = ((prev // 60) % 24).astype("Int64")

    po = pd.DataFrame({
        "jour": dt, "service": _code(val("service")), "ligne": _txt(val("ligne")),
        "sens": _txt(val("sens")), "arret": _txt(val("arret")), "course": _txt(val("course")),
        "vehicule": _txt(val("vehicule")),
        "heure_prevue": _hhmm(prev), "heure_mesuree": _hhmm(mes),
        "heure": heure, "retard": retard})
    po["tranche"] = _tranche(po["heure"].astype(float))
    po = po.dropna(subset=["jour"])

    # --- Rapprochement ponctualité <-> pointages (jour + agent)
    m = po.merge(pt, left_on=["jour", "service"], right_on=["jour", "agent"], how="left", indicator=True)
    rapproche = m["_merge"] == "both"
    pass_ = m[rapproche].drop(columns=["agent", "_merge"]).copy()
    multi = pt.groupby(["jour", "agent"]).conducteur.nunique()
    pass_["mois"] = pass_["jour"].dt.strftime("%Y-%m")
    iso = pass_["jour"].dt.isocalendar()
    pass_["semaine"] = iso["year"].astype(str) + "-S" + iso["week"].astype(str).str.zfill(2)
    pass_["jour_num"] = pass_["jour"].dt.dayofweek + 1
    pass_["jour_sem"] = pass_["jour_num"].map(JOURS)
    pass_["jour"] = pass_["jour"].dt.strftime("%Y-%m-%d")
    cols = ["jour", "mois", "semaine", "jour_num", "jour_sem", "heure", "tranche", "ligne", "sens", "arret",
            "course", "service", "vehicule", "conducteur", "heure_prevue", "heure_mesuree", "retard"]
    pass_ = pass_[cols].reset_index(drop=True)

    n = len(po)
    controle = pd.DataFrame([
        ("Lignes du fichier pointages", f"{len(pt):,}".replace(",", " ")),
        ("Lignes du fichier ponctualité", f"{n:,}".replace(",", " ")),
        ("Passages rapprochés avec un conducteur", f"{int(rapproche.sum()):,}".replace(",", " ")
         + (f" ({100 * rapproche.mean():.1f} %)" if n else "")),
        ("Passages sans pointage (exclus)", f"{int((~rapproche).sum()):,}".replace(",", " ")),
        ("Passages sans heure exploitable", f"{int(pass_['retard'].isna().sum()):,}".replace(",", " ")),
        ("Période analysée", f"{pass_['jour'].min()} → {pass_['jour'].max()}" if len(pass_) else "—"),
        ("Jours analysés", str(pass_["jour"].nunique())),
        ("Agent-jours avec plusieurs conducteurs (passages comptés pour chacun)", str(int((multi > 1).sum()))),
        ("Noms des conducteurs", "anonymisés" if anonymiser else "nominatifs"),
    ], columns=["Contrôle", "Valeur"])
    return pass_, controle


# ----------------------------------------------------------------------------
# 2. CATALOGUE DE REQUÊTES
# ----------------------------------------------------------------------------
@dataclass
class Param:
    nom: str
    label: str
    type: str = "int"          # int | float | text | choice | conducteur
    defaut: object = None
    choix: dict = None         # pour type "choice" : {libellé affiché: colonne SQL}
    aide: str = ""


@dataclass
class Requete:
    id: str
    theme: str
    titre: str
    question: str
    sql: str
    params: list = field(default_factory=list)
    graphique: dict = None     # {"type": "line|bar|heat", "x":..., "y":..., "val":...}
    pivot: tuple = None        # (index, colonnes, valeurs)
    lecture: str = ""

    @property
    def libelle(self):
        return f"{self.theme} · {self.titre}"


# Bloc d'indicateurs commun (ré-utilisé dans {taux})
_TAUX = """COUNT(*) AS passages,
       SUM(retard > :seuil) AS retards,
       ROUND(100.0 * SUM(retard > :seuil) / COUNT(*), 1) AS pct_retard,
       ROUND(100.0 * SUM(retard < :avance) / COUNT(*), 1) AS pct_avance,
       ROUND(100.0 * SUM(retard BETWEEN :avance AND :seuil) / COUNT(*), 1) AS pct_a_l_heure,
       ROUND(AVG(retard), 1) AS retard_moyen_min"""

_MIN = Param("min_passages", "Passages minimum", "int", 30, aide="Ignore les éléments avec trop peu de passages.")
_TOP = Param("top_n", "Nombre de lignes", "int", 20)
_NIV = {"Ligne": "ligne", "Service": "service", "Sens": "sens"}

_REQUETES = [
    Requete("kpi", "Vue d’ensemble", "Indicateurs globaux du réseau",
            "Quelle est la ponctualité du réseau sur la période ?",
            """SELECT COUNT(DISTINCT jour) AS jours, COUNT(DISTINCT conducteur) AS conducteurs, {taux}
FROM p WHERE retard IS NOT NULL""",
            lecture="À l’heure = entre -1 min (avance) et +5 min. Les seuils se règlent dans les filtres."),
    Requete("evolution", "Vue d’ensemble", "Évolution de la ponctualité",
            "La ponctualité s’améliore-t-elle ou se dégrade-t-elle ?",
            """SELECT {pas} AS periode, {taux}
FROM p WHERE retard IS NOT NULL GROUP BY {pas} ORDER BY {pas}""",
            [Param("pas", "Pas de temps", "choice", "Semaine", {"Semaine": "semaine", "Mois": "mois", "Jour": "jour"})],
            graphique={"type": "line", "x": "periode", "y": ["pct_retard", "pct_avance"]}),
    Requete("jour_heure", "Vue d’ensemble", "Carte jours × heures des retards",
            "À quels jours et à quelles heures les retards se concentrent-ils ?",
            """SELECT jour_num, jour_sem, heure, COUNT(*) AS passages, SUM(retard > :seuil) AS retards,
       ROUND(100.0 * SUM(retard > :seuil) / COUNT(*), 1) AS pct_retard
FROM p WHERE retard IS NOT NULL AND heure IS NOT NULL
GROUP BY jour_num, jour_sem, heure ORDER BY jour_num, heure""",
            graphique={"type": "heat", "x": "heure", "y": "jour_sem", "val": "pct_retard"},
            lecture="Heure = heure de départ planifiée."),

    Requete("classement", "Lignes et services", "Classement des lignes / services",
            "Quelles lignes ou quels services cumulent le plus de retards ?",
            """SELECT {niveau} AS element, {taux}, MAX(retard) AS pire_retard_min
FROM p WHERE retard IS NOT NULL GROUP BY {niveau}
HAVING COUNT(*) >= :min_passages ORDER BY pct_retard DESC LIMIT :top_n""",
            [Param("niveau", "Regrouper par", "choice", "Ligne", _NIV), _MIN, _TOP],
            graphique={"type": "bar", "x": "element", "y": "pct_retard"}),
    Requete("diag_structure", "Lignes et services", "Problème de ligne ou problème de conducteur ?",
            "Les retards touchent-ils tous les conducteurs (cause structurelle) ou seulement quelques-uns ?",
            """WITH base AS (
  SELECT {niveau} AS element, conducteur, COUNT(*) AS passages, SUM(retard > :seuil) AS retards
  FROM p WHERE retard IS NOT NULL GROUP BY {niveau}, conducteur),
agg AS (
  SELECT element, COUNT(*) AS nb_conducteurs, SUM(retards >= :retards_mini) AS nb_conducteurs_touches,
         SUM(passages) AS passages, SUM(retards) AS retards
  FROM base GROUP BY element)
SELECT element, nb_conducteurs, nb_conducteurs_touches,
       ROUND(100.0 * nb_conducteurs_touches / nb_conducteurs, 0) AS pct_conducteurs_touches,
       passages, retards, ROUND(100.0 * retards / passages, 1) AS pct_retard,
       CASE WHEN 100.0 * nb_conducteurs_touches / nb_conducteurs >= :pct_touches
                 AND 100.0 * retards / passages > :taux_service AND nb_conducteurs >= 2
            THEN '🔴 Problème structurel (tous les conducteurs touchés)'
            WHEN retards >= 3 AND nb_conducteurs_touches <= 2
            THEN '🟠 Problème individuel (peu de conducteurs)'
            ELSE '⚪ Pas de signal' END AS diagnostic
FROM agg WHERE passages >= :min_passages ORDER BY pct_retard DESC LIMIT :top_n""",
            [Param("niveau", "Regrouper par", "choice", "Ligne", {"Ligne": "ligne", "Service": "service"}),
             Param("pct_touches", "% de conducteurs touchés (seuil)", "int", 70),
             Param("taux_service", "Taux de retard minimum (%)", "int", 20),
             Param("retards_mini", "Retards mini pour être « touché »", "int", 1,
                   aide="Augmenter si trop d’éléments sont en rouge."), _MIN, _TOP],
            lecture="Rouge = au moins 70 % des conducteurs en retard ET taux > 20 % : cause structurelle (horaire, "
                    "temps de parcours). Orange = quelques conducteurs seulement."),
    Requete("arrets", "Lignes et services", "Arrêts les plus pénalisants",
            "Quels arrêts cumulent le plus de minutes de retard ?",
            """SELECT arret, ligne, sens, COUNT(*) AS passages, SUM(retard > :seuil) AS retards,
       ROUND(100.0 * SUM(retard > :seuil) / COUNT(*), 1) AS pct_retard,
       ROUND(AVG(CASE WHEN retard > :seuil THEN retard END), 1) AS retard_moyen_si_retard_min,
       ROUND(SUM(CASE WHEN retard > :seuil THEN retard ELSE 0 END), 0) AS minutes_de_retard_cumulees
FROM p WHERE retard IS NOT NULL AND arret IS NOT NULL GROUP BY arret, ligne, sens
HAVING COUNT(*) >= :min_passages ORDER BY minutes_de_retard_cumulees DESC LIMIT :top_n""",
            [_MIN, _TOP]),
    Requete("tension", "Lignes et services", "Tension par tranche horaire",
            "Sur quelles tranches horaires chaque service est-il le plus tendu ?",
            """SELECT {niveau} AS element, tranche, {taux}
FROM p WHERE retard IS NOT NULL AND tranche IS NOT NULL GROUP BY {niveau}, tranche
HAVING COUNT(*) >= :min_passages ORDER BY pct_retard DESC LIMIT :top_n""",
            [Param("niveau", "Regrouper par", "choice", "Service", {"Service": "service", "Ligne": "ligne"}),
             Param("min_passages", "Passages minimum", "int", 20), _TOP]),

    Requete("conducteurs_retard", "Conducteurs", "Conducteurs les plus en retard",
            "Quels conducteurs ont le plus fort taux de retard ?",
            """SELECT conducteur, COUNT(*) AS passages, SUM(retard > :seuil) AS retards,
       ROUND(100.0 * SUM(retard > :seuil) / COUNT(*), 1) AS pct_retard,
       SUM(retard > :seuil AND retard <= :seuil_grave) AS retards_moderes,
       SUM(retard > :seuil_grave) AS retards_graves,
       ROUND(AVG(CASE WHEN retard > :seuil THEN retard END), 1) AS retard_moyen_si_retard_min,
       MAX(retard) AS pire_retard_min
FROM p WHERE retard IS NOT NULL GROUP BY conducteur
HAVING COUNT(*) >= :min_passages ORDER BY pct_retard DESC LIMIT :top_n""",
            [Param("seuil_grave", "Seuil de retard grave (min)", "int", 15), _MIN, _TOP],
            graphique={"type": "bar", "x": "conducteur", "y": "pct_retard"}),
    Requete("conducteurs_reguliers", "Conducteurs", "Conducteurs les plus réguliers",
            "Qui sont les conducteurs les plus réguliers (à retenir comme référents) ?",
            """SELECT conducteur, COUNT(*) AS passages,
       ROUND(100.0 * SUM(retard BETWEEN 0 AND :reg_max) / COUNT(*), 1) AS score_regularite_pct,
       ROUND(100.0 * SUM(retard < :avance) / COUNT(*), 1) AS pct_avance,
       ROUND(100.0 * SUM(retard > :seuil) / COUNT(*), 1) AS pct_retard
FROM p WHERE retard IS NOT NULL GROUP BY conducteur
HAVING COUNT(*) >= :min_passages ORDER BY score_regularite_pct DESC LIMIT :top_n""",
            [Param("reg_max", "Retard maxi « pile à l’heure » (min)", "int", 3),
             Param("min_passages", "Passages minimum", "int", 100), Param("top_n", "Nombre de lignes", "int", 15)],
            lecture="Score = % de départs entre 0 et 3 min de retard (aucune avance)."),
    Requete("derive", "Conducteurs", "Conducteurs en dérive vs leurs pairs",
            "Quels conducteurs s’écartent significativement de la moyenne des pairs ?",
            """SELECT * FROM (
  WITH pc AS (
    SELECT conducteur, COUNT(*) AS passages, ROUND(100.0 * SUM(retard > :seuil) / COUNT(*), 1) AS taux
    FROM p WHERE retard IS NOT NULL GROUP BY conducteur HAVING COUNT(*) >= :min_passages),
  st AS (SELECT AVG(taux) AS moy, SQRT(MAX(AVG(taux * taux) - AVG(taux) * AVG(taux), 0)) AS sigma FROM pc)
  SELECT pc.conducteur, pc.passages, pc.taux AS pct_retard, ROUND(st.moy, 1) AS moyenne_pairs,
         ROUND(st.sigma, 2) AS sigma,
         ROUND((pc.taux - st.moy) / NULLIF(st.sigma, 0), 2) AS z_score,
         CASE WHEN pc.taux > st.moy + :nb_sigma * st.sigma THEN '🔴 Dérive avérée'
              WHEN pc.taux > st.moy + 2 * st.sigma THEN '🟠 À surveiller'
              ELSE '⚪ Dans la norme' END AS diagnostic
  FROM pc, st) {filtre}
ORDER BY z_score DESC LIMIT :top_n""",
            [Param("nb_sigma", "Seuil de dérive (nb de σ)", "int", 3),
             Param("filtre", "Afficher seulement les alertes", "choice", "Non",
                   {"Non": "", "Oui": "WHERE diagnostic NOT LIKE '⚪%'"}), _MIN, Param("top_n", "Nombre de lignes", "int", 50)],
            lecture="z-score = nombre d’écarts-types au-dessus de la moyenne des conducteurs."),
    Requete("degradation", "Conducteurs", "Dégradation récente par conducteur",
            "Quels conducteurs se sont dégradés ces derniers jours par rapport à leur historique ?",
            """WITH ref AS (SELECT MAX(jour) AS dernier FROM p),
h AS (
  SELECT conducteur,
    SUM(jour <= DATE(ref.dernier, '-' || :fenetre || ' days')) AS passages_avant,
    SUM(jour <= DATE(ref.dernier, '-' || :fenetre || ' days') AND retard > :seuil) AS retards_avant,
    SUM(jour >  DATE(ref.dernier, '-' || :fenetre || ' days')) AS passages_recents,
    SUM(jour >  DATE(ref.dernier, '-' || :fenetre || ' days') AND retard > :seuil) AS retards_recents
  FROM p CROSS JOIN ref WHERE retard IS NOT NULL GROUP BY conducteur)
SELECT conducteur, passages_avant,
       ROUND(100.0 * retards_avant / NULLIF(passages_avant, 0), 1) AS pct_retard_avant,
       passages_recents,
       ROUND(100.0 * retards_recents / NULLIF(passages_recents, 0), 1) AS pct_retard_recent,
       ROUND(100.0 * retards_recents / NULLIF(passages_recents, 0)
           - 100.0 * retards_avant / NULLIF(passages_avant, 0), 1) AS evolution_pts
FROM h WHERE passages_avant >= :min_passages AND passages_recents >= :min_recents
ORDER BY evolution_pts DESC LIMIT :top_n""",
            [Param("fenetre", "Fenêtre récente (jours)", "int", 14), _MIN,
             Param("min_recents", "Passages récents minimum", "int", 10), _TOP]),
    Requete("avances", "Conducteurs", "Conducteurs partant en avance",
            "Quels conducteurs partent le plus souvent en avance ?",
            """SELECT conducteur, COUNT(*) AS passages, SUM(retard < :avance) AS avances,
       ROUND(100.0 * SUM(retard < :avance) / COUNT(*), 1) AS pct_avance,
       ROUND(AVG(CASE WHEN retard < :avance THEN retard END), 1) AS avance_moyenne_min
FROM p WHERE retard IS NOT NULL GROUP BY conducteur
HAVING COUNT(*) >= :min_passages ORDER BY pct_avance DESC LIMIT :top_n""",
            [_MIN, _TOP]),
    Requete("retards_repetes", "Conducteurs", "Retards répétés ces derniers jours",
            "Qui cumule des retards sur la dernière semaine ?",
            """WITH ref AS (SELECT MAX(jour) AS dernier FROM p)
SELECT conducteur, COUNT(*) AS retards, COUNT(DISTINCT jour) AS jours_concernes,
       ROUND(AVG(retard), 1) AS retard_moyen_min, MAX(retard) AS pire_retard_min
FROM p CROSS JOIN ref
WHERE retard > :seuil AND jour > DATE(ref.dernier, '-' || :jours || ' days')
GROUP BY conducteur HAVING COUNT(*) >= :nb_min ORDER BY retards DESC LIMIT :top_n""",
            [Param("jours", "Derniers jours", "int", 7), Param("nb_min", "Retards minimum", "int", 3), _TOP]),
    Requete("suivi_mensuel", "Conducteurs", "Suivi mensuel par conducteur",
            "Comment évolue la ponctualité de chaque conducteur mois après mois ?",
            """SELECT conducteur, mois, COUNT(*) AS passages,
       ROUND(100.0 * SUM(retard BETWEEN :avance AND :seuil) / COUNT(*), 1) AS pct_a_l_heure
FROM p WHERE retard IS NOT NULL GROUP BY conducteur, mois
HAVING COUNT(*) >= :min_passages ORDER BY conducteur, mois""",
            [Param("min_passages", "Passages minimum / mois", "int", 10)],
            pivot=("conducteur", "mois", "pct_a_l_heure"),
            lecture="Une colonne par mois (% à l’heure) + évolution du dernier mois en points."),

    Requete("fiche_lignes", "Fiche conducteur", "Fiche : lignes parcourues",
            "Sur quelles lignes ce conducteur est-il en retard, comparé à ses collègues ?",
            """WITH c AS (SELECT * FROM p WHERE retard IS NOT NULL AND LOWER(conducteur) = LOWER(:conducteur)),
r AS (SELECT ligne, ROUND(100.0 * SUM(retard > :seuil) / COUNT(*), 1) AS pct_retard_reseau
      FROM p WHERE retard IS NOT NULL GROUP BY ligne)
SELECT c.ligne, COUNT(*) AS passages, SUM(c.retard > :seuil) AS retards,
       ROUND(100.0 * SUM(c.retard > :seuil) / COUNT(*), 1) AS pct_retard, r.pct_retard_reseau,
       ROUND(100.0 * SUM(c.retard > :seuil) / COUNT(*) - r.pct_retard_reseau, 1) AS ecart_vs_collegues_pts,
       ROUND(AVG(c.retard), 1) AS retard_moyen_min
FROM c JOIN r ON r.ligne = c.ligne GROUP BY c.ligne ORDER BY ecart_vs_collegues_pts DESC""",
            [Param("conducteur", "Conducteur", "conducteur", "")]),
    Requete("fiche_tranches", "Fiche conducteur", "Fiche : tranches horaires",
            "À quelles heures ce conducteur est-il en retard, comparé à ses collègues ?",
            """WITH c AS (SELECT * FROM p WHERE retard IS NOT NULL AND LOWER(conducteur) = LOWER(:conducteur)),
r AS (SELECT tranche, ROUND(100.0 * SUM(retard > :seuil) / COUNT(*), 1) AS pct_retard_reseau
      FROM p WHERE retard IS NOT NULL GROUP BY tranche)
SELECT c.tranche, COUNT(*) AS passages, SUM(c.retard > :seuil) AS retards,
       ROUND(100.0 * SUM(c.retard > :seuil) / COUNT(*), 1) AS pct_retard, r.pct_retard_reseau,
       ROUND(100.0 * SUM(c.retard > :seuil) / COUNT(*) - r.pct_retard_reseau, 1) AS ecart_vs_collegues_pts
FROM c JOIN r ON r.tranche = c.tranche GROUP BY c.tranche ORDER BY c.tranche""",
            [Param("conducteur", "Conducteur", "conducteur", "")],
            graphique={"type": "bar", "x": "tranche", "y": "pct_retard"}),
    Requete("fiche_retards", "Fiche conducteur", "Fiche : détail des retards",
            "Quels sont précisément les retards de ce conducteur ?",
            """SELECT jour, heure_prevue, ligne, sens, arret, service, ROUND(retard, 1) AS retard_min
FROM p WHERE retard > :seuil AND LOWER(conducteur) = LOWER(:conducteur)
ORDER BY jour, heure_prevue""",
            [Param("conducteur", "Conducteur", "conducteur", "")]),

    Requete("pires_retards", "Événements", "Retards exceptionnels",
            "Quels sont les plus gros retards constatés ?",
            """SELECT jour, heure_prevue, conducteur, service, ligne, sens, arret, ROUND(retard, 1) AS retard_min
FROM p WHERE retard > :seuil_grave ORDER BY retard DESC LIMIT :top_n""",
            [Param("seuil_grave", "Retard supérieur à (min)", "int", 15), Param("top_n", "Nombre de lignes", "int", 100)]),
]
CATALOGUE = {q.id: q for q in _REQUETES}
THEMES = list(dict.fromkeys(q.theme for q in _REQUETES))


def catalogue_df():
    return pd.DataFrame([(q.id, q.theme, q.titre, q.question) for q in _REQUETES],
                        columns=["id", "thème", "requête", "question posée"])


def _cast(p, v):
    if v is None or (isinstance(v, str) and v.strip() == ""):
        v = p.defaut
    if p.type == "int":
        return int(float(v))
    if p.type == "float":
        return float(v)
    return "" if v is None else str(v).strip()


# ----------------------------------------------------------------------------
# 3. BASE DE DONNÉES + EXÉCUTION
# ----------------------------------------------------------------------------
class Base:
    def __init__(self, passages, controle):
        self.passages, self.controle = passages, controle
        self._lock = threading.Lock()
        self.conn = sqlite3.connect(":memory:", check_same_thread=False)
        self.conn.create_function("SQRT", 1, lambda x: math.sqrt(x) if x is not None and x >= 0 else None)
        passages.to_sql("passages", self.conn, index=False, if_exists="replace")
        self.conn.execute("CREATE INDEX idx_jour ON passages(jour)")

    @classmethod
    def depuis_fichiers(cls, fichiers, anonymiser=False):
        """fichiers : dict {nom: bytes | chemin | fichier ouvert} ou liste de chemins (2 fichiers Excel)."""
        if not isinstance(fichiers, dict):
            fichiers = {str(f): f for f in fichiers}
        pts, pons = [], []
        for nom, src in fichiers.items():
            df = _lire(nom, src)
            t = _type_fichier(df)
            if t == "pointages":
                pts.append(df)
            elif t == "ponctualite":
                pons.append(df)
            else:
                raise ValueError(f"Fichier non reconnu : {nom} (ni pointages, ni ponctualité).")
        if not pts:
            raise ValueError("Fichier de pointages manquant (colonnes DATE_POINTAGE, Conducteur, Agent).")
        if not pons:
            raise ValueError("Fichier de ponctualité manquant (colonnes Date, Heure_départ_planifiée, …).")
        return cls(*preparer(pd.concat(pts, ignore_index=True), pd.concat(pons, ignore_index=True), anonymiser))

    # --- listes pour les filtres
    def conducteurs(self):
        return sorted(self.passages["conducteur"].dropna().unique())

    def lignes(self):
        return sorted(self.passages["ligne"].dropna().unique(), key=lambda x: (len(str(x)), str(x)))

    def periode(self):
        return self.passages["jour"].min(), self.passages["jour"].max()

    # --- vue filtrée `p`
    def _vue(self, d1, d2, lignes):
        cond = []
        if d1:
            cond.append(f"jour >= '{pd.Timestamp(d1).strftime('%Y-%m-%d')}'")
        if d2:
            cond.append(f"jour <= '{pd.Timestamp(d2).strftime('%Y-%m-%d')}'")
        if lignes:
            liste = ",".join("'" + str(l).strip().replace("'", "''") + "'" for l in lignes)
            cond.append(f"ligne IN ({liste})")
        where = (" WHERE " + " AND ".join(cond)) if cond else ""
        self.conn.execute("DROP VIEW IF EXISTS p")
        self.conn.execute(f"CREATE TEMP VIEW p AS SELECT * FROM passages{where}")

    def executer(self, qid, d1=None, d2=None, lignes=None, seuil=SEUIL_RETARD, avance=SEUIL_AVANCE, **params):
        """Exécute une requête du catalogue. Les paramètres vides prennent la valeur par défaut."""
        q = CATALOGUE[qid]
        if isinstance(lignes, str):
            lignes = [l for l in re.split(r"[;, ]+", lignes) if l]
        binds, frags = {"seuil": float(seuil), "avance": float(avance)}, {"taux": _TAUX}
        for p in q.params:
            v = params.get(p.nom)
            if p.type == "choice":
                lib = v if v in p.choix else p.defaut
                frags[p.nom] = p.choix[lib]
            else:
                binds[p.nom] = _cast(p, v)
        sql = q.sql.format(**frags)
        with self._lock:
            self._vue(d1, d2, lignes)
            df = pd.read_sql(sql, self.conn, params=binds)
        if q.pivot and len(df):
            i, c, v = q.pivot
            df = df.pivot(index=i, columns=c, values=v).reset_index()
            df.columns.name = None
            mois = [c for c in df.columns if c != i]
            if len(mois) >= 2:
                df["evolution_dernier_mois_pts"] = (df[mois[-1]] - df[mois[-2]]).round(1)
        return df

    def sql(self, requete, d1=None, d2=None, lignes=None):
        """Mode avancé : requête SQL libre sur la vue `p` (ou la table `passages`)."""
        with self._lock:
            self._vue(d1, d2, lignes)
            return pd.read_sql(requete, self.conn)

    def synthese(self, **filtres):
        """Pack « Direction » : dict {feuille: DataFrame}, prêt pour exporter_excel()."""
        e = lambda qid, **kw: self.executer(qid, **filtres, **kw)
        derive = e("derive", filtre="Oui")
        return {
            "Contrôle données": self.controle,
            "Indicateurs": e("kpi"),
            "Évolution hebdo": e("evolution", pas="Semaine"),
            "Lignes": e("classement", niveau="Ligne"),
            "Diagnostic lignes": e("diag_structure", niveau="Ligne"),
            "Arrêts pénalisants": e("arrets"),
            "Conducteurs en dérive": derive,
            "Conducteurs réguliers": e("conducteurs_reguliers"),
        }


# ----------------------------------------------------------------------------
# 4. SORTIES : GRAPHIQUE ET EXCEL
# ----------------------------------------------------------------------------
def graphique(df, spec):
    """Graphique Altair (Streamlit et Colab). Retourne None si pas de graphique pertinent."""
    if not spec or df is None or df.empty:
        return None
    import altair as alt
    ENCRE, OK, RETARD, AVANCE = "#12263F", "#1C9B6C", "#E0482B", "#2E6BD6"
    noms = {"pct_retard": "En retard", "pct_avance": "En avance", "pct_a_l_heure": "À l’heure"}
    coul = {"En retard": RETARD, "En avance": AVANCE, "À l’heure": OK}
    t = spec["type"]
    if t == "line":
        d = df.melt(id_vars=spec["x"], value_vars=spec["y"], var_name="indicateur", value_name="valeur")
        d["indicateur"] = d["indicateur"].map(noms).fillna(d["indicateur"])
        ch = (alt.Chart(d).mark_line(point=alt.OverlayMarkDef(filled=True, size=60), strokeWidth=3)
              .encode(x=alt.X(f"{spec['x']}:N", sort=None, title=None, axis=alt.Axis(labelAngle=-45)),
                      y=alt.Y("valeur:Q", title="% des passages"),
                      color=alt.Color("indicateur:N", title=None,
                                      scale=alt.Scale(domain=list(coul), range=list(coul.values())),
                                      legend=alt.Legend(orient="top")),
                      tooltip=[spec["x"], "indicateur", "valeur"]).properties(height=300))
    elif t == "bar":
        d = df.head(25)
        couleur = RETARD if "retard" in spec["y"] else ENCRE
        ch = (alt.Chart(d).mark_bar(color=couleur, cornerRadiusEnd=4)
              .encode(y=alt.Y(f"{spec['x']}:N", sort="-x", title=None), x=alt.X(f"{spec['y']}:Q", title=noms.get(spec["y"], spec["y"]) + " (%)"),
                      tooltip=list(d.columns[:6])).properties(height=max(200, 24 * len(d))))
    elif t == "heat":
        ch = (alt.Chart(df).mark_rect(stroke="white", strokeWidth=2, cornerRadius=3)
              .encode(x=alt.X(f"{spec['x']}:O", title="Heure de départ planifiée"),
                      y=alt.Y(f"{spec['y']}:N", sort=ORDRE_JOURS, title=None),
                      color=alt.Color(f"{spec['val']}:Q", title="% en retard",
                                      scale=alt.Scale(range=["#EEF3F8", "#F6C08A", RETARD, "#8F1D1D"])),
                      tooltip=list(df.columns)).properties(height=300))
    else:
        return None
    return (ch.configure_view(strokeWidth=0)
              .configure_axis(labelFont="IBM Plex Sans, system-ui, sans-serif", titleFont="IBM Plex Sans, system-ui, sans-serif", labelColor=ENCRE, titleColor="#5B6B7F",
                              gridColor="#E6ECF3", domainColor="#C9D6E6")
              .configure_legend(labelFont="IBM Plex Sans, system-ui, sans-serif", labelColor=ENCRE))


def exporter_excel(feuilles):
    """dict {nom: DataFrame} -> bytes .xlsx (en-têtes bleus, filtres, volets figés, largeurs auto)."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for nom, df in feuilles.items():
            nom = re.sub(r"[\[\]\*\?/\\:]", "", nom)[:31]
            df.to_excel(xw, sheet_name=nom, index=False)
            ws = xw.sheets[nom]
            for cell in ws[1]:
                cell.fill = PatternFill("solid", start_color="1F4E79")
                cell.font = Font(bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            for i, col in enumerate(df.columns, start=1):
                largeur = max([len(str(col))] + [len(str(v)) for v in df[col].head(200)]) + 3
                ws.column_dimensions[get_column_letter(i)].width = min(max(largeur, 12), 50)
            ws.freeze_panes = "A2"
            if len(df.columns):
                ws.auto_filter.ref = ws.dimensions
    return buf.getvalue()


def afficher(df, q):
    """Affichage Notebook : titre, question, graphique et tableau."""
    from IPython.display import Markdown, display
    display(Markdown(f"### {q.titre}\n**{q.question}**" + (f"\n\n*{q.lecture}*" if q.lecture else "")))
    if df is None or df.empty:
        display(Markdown("⚠️ **Aucun résultat** avec ces réglages (conducteur renseigné ? passages minimum trop élevé ?)."))
        return
    g = graphique(df, q.graphique)
    if g is not None:
        display(g)
    display(df.style.hide(axis="index") if len(df) <= 200 else df.head(200))


# ----------------------------------------------------------------------------
# 5. DONNÉES DE DÉMONSTRATION (100 % fictives, pour tester ou présenter l'application)
# ----------------------------------------------------------------------------
def donnees_demo(jours=35, services=14, graine=7):
    """Retourne {nom: bytes .xlsx} : un faux fichier pointages + un faux fichier ponctualité."""
    import datetime as dt
    rng = np.random.default_rng(graine)
    jrs = pd.date_range("2026-09-07", periods=jours)
    agents = [str(100 + i) for i in range(1, services + 1)]
    conds = [f"NOM{i:02d} Prénom{i:02d}" for i in range(1, services + 1)]
    pt, po = [], []
    for j in jrs:
        for k, a in enumerate(agents):
            pt.append((j, 20000 + k, conds[(k + j.day) % services], a))
            ligne = 1 + k % 3
            for c in range(8):
                h, m = 5 + c * 2 + k % 2, (c * 7) % 60
                r = rng.normal(1.5, 3.5) + (5 if ligne == 1 else 0) + (3 if h in (7, 8, 17) else 0)
                tot = h * 60 + m + r
                prev = dt.time(h, m)
                mes = dt.time(int(tot // 60) % 24, int(tot % 60))
                po.append((j, ligne, "Aller" if c % 2 == 0 else "Retour", f"Arrêt {c + 1}", a, prev, mes))
    b1, b2 = io.BytesIO(), io.BytesIO()
    pd.DataFrame(pt, columns=["DATE_POINTAGE", "MATRICULE", "Conducteur", "Agent"]).to_excel(b1, index=False)
    pd.DataFrame(po, columns=["Date", "Ligne", "Sens", "Point d'arrêt", "Conducteur",
                              "Heure_départ_planifiée", "Heure_départ_mesurée"]).to_excel(b2, index=False)
    return {"POINTAGES_DEMO.xlsx": b1.getvalue(), "PONCTUALITE_DEMO.xlsx": b2.getvalue()}
