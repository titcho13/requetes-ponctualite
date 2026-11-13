# -*- coding: utf-8 -*-
"""
Application « Requêtes Ponctualité »
Lancement : python -m streamlit run app_requetes.py

Facultatif : pré-charger les fichiers
  - dossier « donnees » à côté de l'application (tous les .xlsx/.csv sont lus), ou
  - variable PONCTU_FICHIERS=chemin1;chemin2
Facultatif : afficher le SQL sous chaque résultat (mode expert) -> AFFICHER_SQL=1
"""
import hashlib
import os
from datetime import date

import pandas as pd
import streamlit as st

import catalogue_requetes as cr

st.set_page_config(page_title="Requêtes Ponctualité", page_icon="🚌", layout="wide")

# Palette : encre (texte, accents), papier, et trois couleurs de statut
ENCRE, OK, RETARD, AVANCE = "#12263F", "#1C9B6C", "#E0482B", "#2E6BD6"

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@500;600;700&family=IBM+Plex+Sans:wght@400;500;600&display=swap');

.stApp { font-family: 'IBM Plex Sans', system-ui, sans-serif; }
h1, h2, h3, .brand, .hero-num, .q-title, .tile-num, button[data-baseweb="tab"] {
  font-family: 'Barlow Condensed', 'Arial Narrow', sans-serif !important; letter-spacing: .01em; }
#MainMenu, footer { visibility: hidden; }
.block-container { padding-top: 4.2rem; max-width: 1240px; }
.stAppDeployButton { display: none; }

/* Barre latérale */
section[data-testid="stSidebar"] { border-right: 1px solid #D5DEE9; }
section[data-testid="stSidebar"] h2, section[data-testid="stSidebar"] h3 { font-size: 1.25rem; color: #12263F; }
.brand { font-size: 1.9rem; font-weight: 700; color: #12263F; line-height: 1; margin: .2rem 0 .1rem; }
.brand-sub { color: #5B6B7F; font-size: .85rem; margin-bottom: 1.2rem; }
.brand-line { height: 4px; border-radius: 4px; margin: .4rem 0 .9rem;
  background: linear-gradient(90deg, #2E6BD6 0 25%, #1C9B6C 25% 75%, #E0482B 75% 100%); }

/* En-tête de page */
.page-title { font-family: 'Barlow Condensed', sans-serif; font-size: 2.6rem; font-weight: 700;
  color: #12263F; line-height: 1.05; margin: 0; }
.page-sub { color: #5B6B7F; margin: .25rem 0 1.4rem; }

/* Onglets */
button[data-baseweb="tab"] { font-size: 1.2rem; font-weight: 600; padding: .5rem 1.1rem; }
div[data-baseweb="tab-highlight"] { background-color: #12263F !important; height: 3px; }

/* Carte principale : la ponctualité en une ligne de bus */
.hero { background: #fff; border: 1px solid #D5DEE9; border-radius: 14px; padding: 1.5rem 1.8rem 1.3rem; margin-bottom: 1rem; }
.hero-label { color: #5B6B7F; font-size: .95rem; }
.hero-num { font-size: 4.6rem; font-weight: 700; color: #12263F; line-height: 1; margin: .15rem 0 .2rem; }
.hero-num span { font-size: 2rem; font-weight: 600; color: #1C9B6C; margin-left: .15rem; }
.hero-sub { color: #5B6B7F; font-size: .92rem; margin-bottom: 1.2rem; }
.route { display: flex; gap: 4px; height: 16px; position: relative; }
.seg { border-radius: 8px; position: relative; min-width: 6px; }
.seg:not(:last-child)::after { content: ""; position: absolute; right: -9px; top: -3px; width: 10px; height: 10px;
  background: #fff; border: 3px solid #12263F; border-radius: 50%; z-index: 2; box-sizing: content-box; }
.legend { display: flex; gap: 1.6rem; flex-wrap: wrap; margin-top: .9rem; color: #12263F; font-size: .92rem; }
.legend b { font-weight: 600; }
.dot { display: inline-block; width: 10px; height: 10px; border-radius: 50%; margin-right: .4rem; }

.tile { background: #fff; border: 1px solid #D5DEE9; border-radius: 12px; padding: .9rem 1.1rem; height: 100%; }
.tile-num { font-size: 2.1rem; font-weight: 700; color: #12263F; line-height: 1.05; }
.tile-label { color: #5B6B7F; font-size: .88rem; }

/* Questions */
.q-title { font-size: 2rem; font-weight: 600; color: #12263F; line-height: 1.15; margin: 1rem 0 .6rem; }
.note { background: #E6ECF3; border-left: 4px solid #12263F; border-radius: 6px; padding: .6rem .9rem;
  color: #12263F; font-size: .92rem; margin: .2rem 0 .9rem; }
.count { color: #5B6B7F; font-size: .9rem; margin-bottom: .4rem; }

/* Accueil */
.welcome { background: #12263F; color: #fff; border-radius: 16px; padding: 2.6rem 2.6rem 2.2rem; margin-bottom: 1.2rem; }
.welcome h1 { color: #fff; font-size: 3rem; margin: 0 0 .5rem; line-height: 1.05; }
.welcome p { color: #C9D6E6; font-size: 1.05rem; max-width: 640px; margin: 0; }
.steps { display: flex; gap: 1rem; flex-wrap: wrap; margin: 1.2rem 0 1.4rem; }
.step { flex: 1 1 220px; background: #fff; border: 1px solid #D5DEE9; border-radius: 12px; padding: 1rem 1.2rem; }
.step-n { font-family: 'Barlow Condensed', sans-serif; font-size: 2rem; font-weight: 700; color: #2E6BD6; line-height: 1; }
.step b { display: block; color: #12263F; margin: .2rem 0; }
.step span { color: #5B6B7F; font-size: .92rem; }

/* Boutons */
.stButton > button, .stDownloadButton > button { border-radius: 8px; font-weight: 600; }
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] { background: #12263F; border-color: #12263F; color: #fff; }
.stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover { background: #1B3A63; border-color: #1B3A63; }
@media (max-width: 700px) { .hero-num { font-size: 3.4rem; } .page-title { font-size: 2rem; } .welcome h1 { font-size: 2.2rem; } }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)


def html(s):
    """HTML sur une seule ligne (évite que Markdown le prenne pour du code)."""
    st.markdown(" ".join(l.strip() for l in s.splitlines()), unsafe_allow_html=True)


def fr(n, d=0):
    return f"{n:,.{d}f}".replace(",", "\u202f").replace(".", ",").replace("-", "−")


LIB = {"pct_retard": "% en retard", "pct_avance": "% en avance", "pct_a_l_heure": "% à l’heure",
       "retard_moyen_min": "Retard moyen (min)", "pire_retard_min": "Pire retard (min)", "retard_min": "Retard (min)",
       "retard_moyen_si_retard_min": "Retard moyen si retard (min)", "avance_moyenne_min": "Avance moyenne (min)",
       "pct_conducteurs_touches": "% conducteurs touchés", "pct_retard_reseau": "% en retard (réseau)",
       "ecart_vs_collegues_pts": "Écart vs collègues (pts)", "evolution_pts": "Évolution (pts)",
       "evolution_dernier_mois_pts": "Évolution dernier mois (pts)", "score_regularite_pct": "Régularité (%)",
       "minutes_de_retard_cumulees": "Minutes de retard cumulées", "z_score": "Z-score", "sigma": "Écart-type", "element": "Élément", "arret": "Arrêt"}


def joli(df):
    """Libellés lisibles pour l'affichage et les exports."""
    def nom(c):
        c = str(c)
        return LIB.get(c, c if c[:1].isdigit() else c.replace("_", " ").capitalize())
    return df.rename(columns=nom)


# ----------------------------------------------------------------------------
# Chargement (mis en cache : les 2 fichiers ne sont lus qu'une fois)
# ----------------------------------------------------------------------------
@st.cache_resource(show_spinner="Préparation des données…")
def charger(cle, _fichiers, anonymiser):
    return cr.Base.depuis_fichiers(_fichiers, anonymiser)


def fichiers_depuis_sidebar():
    st.sidebar.markdown('<div class="brand">Ponctualité</div><div class="brand-sub">Requêtes prêtes à l’emploi</div>'
                        '<div class="brand-line"></div>', unsafe_allow_html=True)
    st.sidebar.subheader("Vos fichiers")
    ups = st.sidebar.file_uploader("Pointages et passages mesurés (.xlsx)", type=["xlsx", "xls", "csv"],
                                   accept_multiple_files=True,
                                   help="Déposez les 2 fichiers ensemble, dans n’importe quel ordre : ils sont reconnus à leurs colonnes.")
    if ups:
        return {u.name: u.getvalue() for u in ups}
    pre = os.environ.get("PONCTU_FICHIERS")
    if pre:
        chemins = [c for c in pre.split(os.pathsep) if c.strip()]
        return {os.path.basename(c): open(c, "rb").read() for c in chemins}
    dossier = os.path.join(os.path.dirname(os.path.abspath(__file__)), "donnees")
    if os.path.isdir(dossier):
        chemins = [os.path.join(dossier, f) for f in sorted(os.listdir(dossier))
                   if f.lower().endswith((".xlsx", ".xls", ".csv")) and not f.startswith("~$")]
        if chemins:
            st.sidebar.success("Base pré-chargée : " + ", ".join(os.path.basename(c) for c in chemins))
            return {os.path.basename(c): open(c, "rb").read() for c in chemins}
    if st.session_state.get("demo"):
        st.sidebar.warning("Données de démonstration (fictives).")
        return cr.donnees_demo()
    return None


fichiers = fichiers_depuis_sidebar()

# ----------------------------------------------------------------------------
# Écran d'accueil
# ----------------------------------------------------------------------------
if not fichiers:
    html("""<div class="welcome"><h1>Vos retards, expliqués en quelques clics</h1>
    <p>Déposez deux fichiers Excel. L’application les rapproche et répond à 18 questions prêtes à l’emploi,
    sans une ligne de SQL.</p></div>
    <div class="steps">
    <div class="step"><div class="step-n">1</div><b>Déposez vos fichiers</b><span>Pointages et passages mesurés, dans la colonne de gauche.</span></div>
    <div class="step"><div class="step-n">2</div><b>Choisissez une question</b><span>Lignes en retard, conducteurs en dérive, heures de tension…</span></div>
    <div class="step"><div class="step-n">3</div><b>Lisez et exportez</b><span>Tableau, graphique et export Excel en un clic.</span></div>
    </div>""")
    if st.button("Essayer avec des données fictives", type="primary"):
        st.session_state["demo"] = True
        st.rerun()
    st.stop()

anonymiser = st.sidebar.checkbox("Anonymiser les conducteurs", value=bool(st.session_state.get("demo")),
                                 help="Remplace les noms par « Conducteur 001 », … (ordre aléatoire).")
try:
    cle = hashlib.md5(b"".join(sorted(hashlib.md5(v).digest() for v in fichiers.values()))).hexdigest()
    base = charger(cle + str(anonymiser), fichiers, anonymiser)
except ValueError as e:
    st.error(str(e))
    st.stop()

# ----------------------------------------------------------------------------
# Filtres communs
# ----------------------------------------------------------------------------
d_min, d_max = [pd.Timestamp(x).date() for x in base.periode()]
st.sidebar.subheader("Filtres")
periode = st.sidebar.date_input("Période", value=(d_min, d_max), min_value=d_min, max_value=d_max)
lignes = st.sidebar.multiselect("Lignes", base.lignes(), placeholder="Toutes les lignes")
with st.sidebar.expander("Seuils de ponctualité"):
    seuil = st.number_input("En retard au-delà de (min)", value=5.0, step=1.0)
    avance = st.number_input("En avance en deçà de (min)", value=-1.0, step=1.0)
d1, d2 = (periode if isinstance(periode, (tuple, list)) and len(periode) == 2 else (d_min, d_max))
F = dict(d1=d1, d2=d2, lignes=lignes or None, seuil=seuil, avance=avance)

st.markdown('<div class="page-title">Ponctualité du réseau</div>', unsafe_allow_html=True)
st.markdown(f'<div class="page-sub">Du {d1:%d/%m/%Y} au {d2:%d/%m/%Y}'
            + (f", lignes {', '.join(lignes)}" if lignes else ", toutes les lignes") + "</div>", unsafe_allow_html=True)

tab_syn, tab_req, tab_ctl = st.tabs(["Synthèse", "Poser une question", "Contrôle des données"])

# ----------------------------------------------------------------------------
# Synthèse
# ----------------------------------------------------------------------------
with tab_syn:
    k = base.executer("kpi", **F)
    if k.empty or int(k.loc[0, "passages"]) == 0:
        st.warning("Aucun passage sur cette sélection : élargissez la période ou les lignes.")
    else:
        r = k.loc[0]
        ok, late, early = float(r.pct_a_l_heure), float(r.pct_retard), float(r.pct_avance)
        html(f"""<div class="hero">
        <div class="hero-label">Départs à l’heure</div>
        <div class="hero-num">{fr(ok, 1).split(',')[0]}<span>,{fr(ok, 1).split(',')[1]} %</span></div>
        <div class="hero-sub">entre {fr(avance)} et +{fr(seuil)} min de l’horaire, sur {fr(int(r.passages))} passages et {int(r.jours)} jours</div>
        <div class="route">
          <div class="seg" style="flex:{max(early, .01)} 1 0; background:{AVANCE}"></div>
          <div class="seg" style="flex:{max(ok, .01)} 1 0; background:{OK}"></div>
          <div class="seg" style="flex:{max(late, .01)} 1 0; background:{RETARD}"></div>
        </div>
        <div class="legend">
          <div><span class="dot" style="background:{AVANCE}"></span>En avance <b>{fr(early, 1)} %</b></div>
          <div><span class="dot" style="background:{OK}"></span>À l’heure <b>{fr(ok, 1)} %</b></div>
          <div><span class="dot" style="background:{RETARD}"></span>En retard <b>{fr(late, 1)} %</b></div>
        </div></div>""")
        c = st.columns(3)
        for col, num, lab in zip(c, [fr(int(r.passages)), fr(float(r.retard_moyen_min), 1) + " min", fr(int(r.conducteurs))],
                                 ["passages analysés", "retard moyen (toutes situations)", "conducteurs concernés"]):
            col.markdown(f'<div class="tile"><div class="tile-num">{num}</div><div class="tile-label">{lab}</div></div>',
                         unsafe_allow_html=True)

        st.subheader("Évolution par semaine")
        ev = base.executer("evolution", pas="Semaine", **F)
        g = cr.graphique(ev, cr.CATALOGUE["evolution"].graphique)
        if g is not None:
            st.altair_chart(g)
        a, b = st.columns(2)
        with a:
            st.subheader("Lignes les plus en retard")
            st.dataframe(joli(base.executer("classement", niveau="Ligne", top_n=5, **F)), hide_index=True)
        with b:
            st.subheader("Conducteurs à surveiller")
            dv = base.executer("derive", filtre="Oui", **F)
            if len(dv):
                st.dataframe(joli(dv), hide_index=True)
            else:
                st.success("Aucun conducteur en dérive sur cette période.")
        st.download_button("Télécharger la synthèse (Excel)", cr.exporter_excel({k_: joli(v_) for k_, v_ in base.synthese(**F).items()}), type="primary",
                           file_name=f"synthese_ponctualite_{date.today():%Y%m%d}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ----------------------------------------------------------------------------
# Catalogue de questions
# ----------------------------------------------------------------------------
with tab_req:
    theme = st.radio("Thème", cr.THEMES, horizontal=True)
    qs = [q for q in cr.CATALOGUE.values() if q.theme == theme]
    q = st.selectbox("Question", qs, format_func=lambda x: x.titre)
    st.markdown(f'<div class="q-title">{q.question}</div>', unsafe_allow_html=True)

    valeurs = {}
    if q.params:
        with st.container(border=True):
            cols = st.columns(min(len(q.params), 4))
            for i, p in enumerate(q.params):
                with cols[i % len(cols)]:
                    k_ = f"{q.id}_{p.nom}"
                    if p.type == "choice":
                        valeurs[p.nom] = st.selectbox(p.label, list(p.choix), index=list(p.choix).index(p.defaut), key=k_, help=p.aide or None)
                    elif p.type == "conducteur":
                        valeurs[p.nom] = st.selectbox(p.label, base.conducteurs(), key=k_)
                    elif p.type == "float":
                        valeurs[p.nom] = st.number_input(p.label, value=float(p.defaut), key=k_, help=p.aide or None)
                    else:
                        valeurs[p.nom] = st.number_input(p.label, value=int(p.defaut), step=1, key=k_, help=p.aide or None)

    with st.spinner("Calcul en cours…"):
        res = base.executer(q.id, **F, **valeurs)
    if res.empty:
        st.warning("Aucun résultat avec ces réglages. Essayez de baisser le nombre de passages minimum.")
    else:
        if q.lecture:
            st.markdown(f'<div class="note">{q.lecture}</div>', unsafe_allow_html=True)
        g = cr.graphique(res, q.graphique)
        if g is not None:
            st.altair_chart(g)
        st.markdown(f'<div class="count">{len(res)} ligne(s)</div>', unsafe_allow_html=True)
        st.dataframe(joli(res), hide_index=True)
        d1_, d2_, _ = st.columns([1, 1, 4])
        d1_.download_button("Exporter en Excel", cr.exporter_excel({q.titre: joli(res)}), file_name=f"{q.id}.xlsx", type="primary",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        d2_.download_button("Exporter en CSV", joli(res).to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
                            file_name=f"{q.id}.csv", mime="text/csv")
    if os.environ.get("AFFICHER_SQL") == "1":
        with st.expander("SQL (mode expert)"):
            st.code(q.sql, language="sql")

# ----------------------------------------------------------------------------
# Contrôle qualité
# ----------------------------------------------------------------------------
with tab_ctl:
    st.markdown("Le rapprochement associe chaque passage au conducteur pointé le même jour sur le même service. "
                "Un taux faible signale un fichier incomplet ou des codes agent différents entre les 2 fichiers.")
    st.dataframe(base.controle, hide_index=True)
    st.markdown("**Définitions.** Retard = départ mesuré moins départ planifié, en minutes. "
                "À l’heure = entre le seuil d’avance et le seuil de retard. Les heures et tranches horaires "
                "correspondent à l’heure de départ planifiée.")
