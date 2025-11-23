import os
import io
import datetime as dt
import pandas as pd
import streamlit as st

# ---------------------------
# Configuration de la page
# ---------------------------
st.set_page_config(page_title="Gestion Carburant", page_icon="⛽", layout="wide")

DATA_CSV = "carburant_operations.csv"

# ---------------------------
# Helpers de persistance
# ---------------------------
def load_data():
    if os.path.exists(DATA_CSV):
        try:
            df = pd.read_csv(DATA_CSV, parse_dates=["Date"])
            # Normalisation des colonnes si besoin
            expected_cols = ["Date", "Engin", "PU", "Montant", "Quantite_L"]
            for col in expected_cols:
                if col not in df.columns:
                    df[col] = pd.Series(dtype="float64") if col != "Date" and col != "Engin" else ""
            # Types
            df["PU"] = pd.to_numeric(df["PU"], errors="coerce")
            df["Montant"] = pd.to_numeric(df["Montant"], errors="coerce")
            df["Quantite_L"] = pd.to_numeric(df["Quantite_L"], errors="coerce")
            return df[["Date", "Engin", "PU", "Montant", "Quantite_L"]].sort_values("Date")
        except Exception:
            pass
    # DataFrame vide par défaut
    return pd.DataFrame(columns=["Date", "Engin", "PU", "Montant", "Quantite_L"])

def save_data(df):
    df_out = df.copy()
    df_out.to_csv(DATA_CSV, index=False)

# ---------------------------
# Session state
# ---------------------------
if "df" not in st.session_state:
    st.session_state.df = load_data()

# ---------------------------
# Sidebar: paramètres
# ---------------------------
st.sidebar.header("Paramètres")
stock_initial = st.sidebar.number_input("Stock initial (L)", min_value=0, value=9000, step=100)
pu = st.sidebar.number_input("PU (Prix unitaire)", min_value=0.0, value=775.0, step=1.0, format="%.2f")
seuil = st.sidebar.number_input("Seuil sécurité (L)", min_value=0, value=1000, step=100)

with st.sidebar.expander("Options", expanded=False):
    if st.button("Réinitialiser l'historique"):
        st.session_state.df = pd.DataFrame(columns=["Date", "Engin", "PU", "Montant", "Quantite_L"])
        save_data(st.session_state.df)
        st.success("Historique réinitialisé.")

# ---------------------------
# Header
# ---------------------------
st.title("⛽ Gestion Carburant — Modèle interactif")
st.caption("Saisis le montant ou la quantité; le reste est calculé. Stock réel cumulatif, alerte sécurité et dashboard visuel.")

# ---------------------------
# Formulaire d'ajout d'opération
# ---------------------------
st.subheader("Ajouter une opération")
col_form1, col_form2, col_form3, col_form4 = st.columns([1.1, 1.1, 1.1, 1.1])

with col_form1:
    date_op = st.date_input("Date", value=dt.date.today())
with col_form2:
    engin = st.text_input("N° Engin", placeholder="Ex: CHASSIS-123")
with col_form3:
    montant = st.number_input("Montant", min_value=0.0, value=0.0, step=1000.0)
with col_form4:
    quantite = st.number_input("Quantité achetée (L)", min_value=0.0, value=0.0, step=10.0)

# Règles de calcul:
# - Si montant > 0 et quantite == 0 -> quantite = montant / pu
# - Si quantite > 0 et montant == 0 -> montant = pu * quantite
# - Si les deux sont fournis, on vérifie la cohérence et on privilégie la quantité (avec recalcul du montant)
calc_info = ""
if pu and pu > 0:
    if montant > 0 and quantite == 0:
        quantite_calc = montant / pu
        calc_info = f"Quantité calculée automatiquement: {quantite_calc:.2f} L (Montant / PU)"
    elif quantite > 0 and montant == 0:
        montant_calc = pu * quantite
        calc_info = f"Montant calculé automatiquement: {montant_calc:.0f} (PU × Quantité)"
    elif montant > 0 and quantite > 0:
        # Vérification de cohérence (tolérance)
        montant_expected = pu * quantite
        if abs(montant - montant_expected) > max(1.0, 0.01 * montant_expected):
            calc_info = f"Attention: incohérence détectée (PU × Quantité ≈ {montant_expected:.0f}). Le montant sera recalculé."
        else:
            calc_info = "Données cohérentes."
else:
    st.warning("PU doit être > 0 pour effectuer les calculs.")

if calc_info:
    st.info(calc_info)

btn_add = st.button("Valider l'achat ✅", type="primary")


if btn_add:
    if pu <= 0:
        st.error("PU doit être supérieur à 0.")
    else:
        # Calcul automatique
        if montant > 0 and quantite == 0:
            quantite = montant / pu
        elif quantite > 0 and montant == 0:
            montant = pu * quantite
        elif montant == 0 and quantite == 0:
            st.error("Saisir au moins Montant ou Quantité.")
        else:
            montant = pu * quantite  # cohérence

        # Ajout direct dans l’historique
        new_row = {
            "Date": pd.to_datetime(date_op),
            "Engin": engin.strip(),
            "PU": float(pu),
            "Montant": float(montant),
            "Quantite_L": float(quantite),
        }
        st.session_state.df = pd.concat([st.session_state.df, pd.DataFrame([new_row])], ignore_index=True)
        st.session_state.df = st.session_state.df.sort_values("Date").reset_index(drop=True)
        save_data(st.session_state.df)
        st.success("Achat validé et ajouté au tableau ✅")

# ---------------------------
# Calculs cumulés et affichage
# ---------------------------
df = st.session_state.df.copy()

if not df.empty:
    # Stock réel cumulatif
    df["Quantite_L"] = pd.to_numeric(df["Quantite_L"], errors="coerce").fillna(0.0)
    df["Montant"] = pd.to_numeric(df["Montant"], errors="coerce").fillna(0.0)
    df["PU"] = pd.to_numeric(df["PU"], errors="coerce").fillna(pu)

    df["Cumul_Quantite_L"] = df["Quantite_L"].cumsum()
    df["Stock_reel_L"] = stock_initial - df["Cumul_Quantite_L"]
    df["Alerte"] = df["Stock_reel_L"].apply(lambda x: "ALERTE" if x < seuil else "OK")

    st.subheader("Historique des opérations")
    show_cols = ["Date", "Engin", "PU", "Montant", "Quantite_L", "Stock_reel_L", "Alerte"]
    st.dataframe(df[show_cols], use_container_width=True)

    # KPIs
    col_kpi1, col_kpi2, col_kpi3 = st.columns(3)
    with col_kpi1:
        st.metric("Stock initial (L)", f"{stock_initial:,.0f}")
    with col_kpi2:
        total_qte = df["Quantite_L"].sum()
        st.metric("Total quantités achetées (L)", f"{total_qte:,.0f}")
    with col_kpi3:
        current_stock = df["Stock_reel_L"].iloc[-1]
        st.metric("Stock réel (L)", f"{current_stock:,.0f}", delta=None)

    # Charts
    st.subheader("Dashboard")
    c1, c2 = st.columns(2)

    with c1:
        st.write("Évolution du stock réel")
        chart_df = df[["Date", "Stock_reel_L"]].set_index("Date")
        st.line_chart(chart_df)

    with c2:
        st.write("Quantité achetée par Engin")
        bar_df = df.groupby("Engin", as_index=True)["Quantite_L"].sum().sort_values(ascending=False)
        st.bar_chart(bar_df)

    # Export
    st.subheader("Exports")
    col_exp1, col_exp2 = st.columns(2)

    with col_exp1:
        if st.button("Exporter en CSV"):
            save_data(df[["Date", "Engin", "PU", "Montant", "Quantite_L"]])
            st.success(f"Export CSV enregistré: {DATA_CSV}")

    with col_exp2:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            # Feuille opérations
            df.to_excel(writer, index=False, sheet_name="Operations")
            # Feuille résumé rapide
            resume = pd.DataFrame({
                "Stock_initial_L": [stock_initial],
                "Seuil_securite_L": [seuil],
                "PU": [pu],
                "Total_quantite_L": [df["Quantite_L"].sum()],
                "Stock_reel_L": [df["Stock_reel_L"].iloc[-1]]
            })
            resume.to_excel(writer, index=False, sheet_name="Resume")
        st.download_button(
            label="Télécharger Excel",
            data=buffer.getvalue(),
            file_name="carburant_dashboard.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Aucune opération pour l’instant. Ajoute une ligne ci‑dessus pour démarrer.")

# ---------------------------
# Notes d’usage
# ---------------------------
with st.expander("Notes"):
    st.markdown(
        "- Saisis soit Montant, soit Quantité; l’app calcule l’autre valeur automatiquement.\n"
        "- Le Stock réel est calculé de manière cumulative: Stock initial − somme des quantités.\n"
        "- Modifie PU, Stock initial et Seuil sécurité dans la barre latérale pour adapter le calcul.\n"
        "- Les données sont sauvegardées en local dans un CSV (utile hors connexion)."
    )
