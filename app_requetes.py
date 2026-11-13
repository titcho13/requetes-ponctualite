# -*- coding: utf-8 -*-
"""
Application « Requêtes Ponctualité »
Lancement : streamlit run app_requetes.py   (ou double-clic sur lancer_app.bat)

Facultatif : pré-charger les 2 fichiers sans les glisser à chaque fois
  set PONCTU_FICHIERS=C:\\chemin\\POINTAGESSO.xlsx;C:\\chemin\\PonctualiteSO.xlsx
Facultatif : afficher le SQL sous chaque résultat (mode expert)  ->  set AFFICHER_SQL=1
"""
import hashlib
import os
from datetime import date

import pandas as pd
import streamlit as st

import catalogue_requetes as cr

st.set_page_config(page_title="Requêtes Ponctualité", page_icon="🚌", layout="wide")


# ----------------------------------------------------------------------------
# Chargement (mis en cache : les 2 fichiers ne sont lus qu'une fois)
# ----------------------------------------------------------------------------
@st.cache_resource(show_spinner="Préparation des données…")
def charger(cle, _fichiers, anonymiser):
    return cr.Base.depuis_fichiers(_fichiers, anonymiser)


def fichiers_depuis_sidebar():
    st.sidebar.header("1 · Les 2 fichiers")
    ups = st.sidebar.file_uploader("Déposer POINTAGES + PONCTUALITÉ (.xlsx)", type=["xlsx", "xls", "csv"],
                                   accept_multiple_files=True,
                                   help="L’ordre n’a pas d’importance : chaque fichier est reconnu à ses colonnes.")
    if ups:
        return {u.name: u.getvalue() for u in ups}
    pre = os.environ.get("PONCTU_FICHIERS")
    if pre:
        chemins = [c for c in pre.split(os.pathsep) if c.strip()]
        return {os.path.basename(c): open(c, "rb").read() for c in chemins}
    # Base pré-chargée : tous les fichiers du dossier "donnees" placé à côté de l'application
    dossier = os.path.join(os.path.dirname(os.path.abspath(__file__)), "donnees")
    if os.path.isdir(dossier):
        chemins = [os.path.join(dossier, f) for f in sorted(os.listdir(dossier))
                   if f.lower().endswith((".xlsx", ".xls", ".csv")) and not f.startswith("~$")]
        if chemins:
            st.sidebar.success("Base pré-chargée : " + ", ".join(os.path.basename(c) for c in chemins))
            return {os.path.basename(c): open(c, "rb").read() for c in chemins}
    if st.session_state.get("demo"):
        st.sidebar.warning("Mode démonstration : données fictives.")
        return cr.donnees_demo()
    return None


fichiers = fichiers_depuis_sidebar()
st.title("🚌 Requêtes Ponctualité")
st.caption("Analyse de ponctualité d’un réseau de transport · pointages × passages mesurés")

if not fichiers:
    st.info("👈 Déposez les 2 fichiers Excel (pointages et ponctualité) dans la colonne de gauche pour commencer.")
    if st.button("▶ Essayer avec des données de démonstration (fictives)"):
        st.session_state["demo"] = True
        st.rerun()
    st.stop()

anonymiser = st.sidebar.checkbox("Anonymiser les noms des conducteurs", value=bool(st.session_state.get("demo")),
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
st.sidebar.header("2 · Filtres (facultatif)")
periode = st.sidebar.date_input("Période", value=(d_min, d_max), min_value=d_min, max_value=d_max)
lignes = st.sidebar.multiselect("Lignes", base.lignes(), placeholder="Toutes les lignes")
with st.sidebar.expander("Seuils"):
    seuil = st.number_input("Retard si supérieur à (min)", value=5.0, step=1.0)
    avance = st.number_input("Avance si inférieure à (min)", value=-1.0, step=1.0)
d1, d2 = (periode if isinstance(periode, (tuple, list)) and len(periode) == 2 else (d_min, d_max))
F = dict(d1=d1, d2=d2, lignes=lignes or None, seuil=seuil, avance=avance)

tab_syn, tab_req, tab_ctl = st.tabs(["📊 Synthèse", "🔎 Poser une question", "🧪 Contrôle des données"])

# ----------------------------------------------------------------------------
# Synthèse direction
# ----------------------------------------------------------------------------
with tab_syn:
    k = base.executer("kpi", **F)
    if k.empty or int(k.loc[0, "passages"]) == 0:
        st.warning("Aucun passage sur cette sélection.")
    else:
        r = k.loc[0]
        c = st.columns(5)
        c[0].metric("Passages analysés", f"{int(r.passages):,}".replace(",", " "))
        c[1].metric("À l’heure", f"{r.pct_a_l_heure} %")
        c[2].metric("En retard", f"{r.pct_retard} %")
        c[3].metric("En avance", f"{r.pct_avance} %")
        c[4].metric("Retard moyen", f"{r.retard_moyen_min} min")
        st.subheader("Évolution")
        ev = base.executer("evolution", pas="Semaine", **F)
        g = cr.graphique(ev, cr.CATALOGUE["evolution"].graphique)
        if g is not None:
            st.altair_chart(g)
        a, b = st.columns(2)
        with a:
            st.subheader("Lignes les plus en retard")
            st.dataframe(base.executer("classement", niveau="Ligne", top_n=5, **F), hide_index=True)
        with b:
            st.subheader("Conducteurs en dérive")
            dv = base.executer("derive", filtre="Oui", **F)
            st.dataframe(dv, hide_index=True) if len(dv) else st.success("Aucun conducteur en dérive avérée ou à surveiller.")
        st.download_button("⬇️ Télécharger la synthèse (Excel)", cr.exporter_excel(base.synthese(**F)),
                           file_name=f"synthese_ponctualite_{date.today():%Y%m%d}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ----------------------------------------------------------------------------
# Catalogue de questions
# ----------------------------------------------------------------------------
with tab_req:
    c1, c2 = st.columns([1, 2])
    theme = c1.selectbox("Thème", cr.THEMES)
    qs = [q for q in cr.CATALOGUE.values() if q.theme == theme]
    q = c2.selectbox("Question", qs, format_func=lambda x: x.titre)
    st.markdown(f"**{q.question}**")

    valeurs = {}
    if q.params:
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

    with st.spinner("Calcul…"):
        res = base.executer(q.id, **F, **valeurs)
    if res.empty:
        st.warning("Aucun résultat avec ces réglages (essayez de baisser le nombre de passages minimum).")
    else:
        st.caption(f"{len(res)} ligne(s)" + (f" · {q.lecture}" if q.lecture else ""))
        g = cr.graphique(res, q.graphique)
        if g is not None:
            st.altair_chart(g)
        st.dataframe(res, hide_index=True)
        d1_, d2_ = st.columns(2)
        d1_.download_button("⬇️ Excel", cr.exporter_excel({q.titre: res}), file_name=f"{q.id}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        d2_.download_button("⬇️ CSV", res.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
                            file_name=f"{q.id}.csv", mime="text/csv")
    if os.environ.get("AFFICHER_SQL") == "1":
        with st.expander("SQL (mode expert)"):
            st.code(q.sql, language="sql")

# ----------------------------------------------------------------------------
# Contrôle qualité
# ----------------------------------------------------------------------------
with tab_ctl:
    st.markdown("Vérification du **rapprochement** entre les 2 fichiers (jour + code agent). "
                "Un taux de rapprochement bas signale un fichier incomplet ou des codes agent différents.")
    st.dataframe(base.controle, hide_index=True)
    st.markdown("**Définitions** — retard = départ mesuré − départ planifié (min) · À l’heure : entre l’avance et le seuil · "
                "Tranche / heure = heure de départ planifiée.")
