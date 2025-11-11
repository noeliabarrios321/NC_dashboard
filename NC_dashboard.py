import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from datetime import date

# ==========================
# CONFIG
# ==========================
st.set_page_config(page_title="NC Dashboard", layout="wide")
TODAY = pd.Timestamp(date.today())

# ==========================
# HELPERS
# ==========================
def safe_to_datetime(series):
    return pd.to_datetime(series, errors="coerce")

def count_unique(series):
    return int(pd.Series(series).dropna().nunique())

# ==========================
# LOAD DATA
# ==========================
@st.cache_data
def load_data(path="NC_monitoring.xlsx"):
    xls = pd.ExcelFile(path)
    df = pd.read_excel(path, sheet_name=xls.sheet_names[0])

    # Normalizar columnas
    df = df.rename(columns={c: c.strip().lower() for c in df.columns})
    rename_map = {
        "nc number": "nc_number",
        "status": "status",
        "responsible site": "responsible_site",
        "initiation date": "created_date",
        "closed date": "closed_date"
    }
    df = df.rename(columns=rename_map)

    # Fechas
    for col in ["created_date", "closed_date"]:
        if col in df.columns:
            df[col] = safe_to_datetime(df[col])

    df["responsible_site"] = df["responsible_site"].astype(str).str.strip()
    df["status"] = df["status"].astype(str).str.strip()
    df = df[df["nc_number"].notna()].copy()

    # Consolidar por NC
    df_unique = df.groupby("nc_number").agg(
        created_date=("created_date", "min"),
        closed_date=("closed_date", "max"),
        status=("status", "last"),
        responsible_site=("responsible_site", "last")
    ).reset_index()

    return df_unique

df_unique = load_data()

# ==========================
# SECTION: Planned Closure Date Calculation
# ==========================
xls_full = pd.ExcelFile("NC_monitoring.xlsx")
df_full = pd.read_excel("NC_monitoring.xlsx", sheet_name=xls_full.sheet_names[0])

# Normalizar nombres
df_full.columns = df_full.columns.str.strip().str.lower()
df_full = df_full.rename(columns={
    "nc number": "nc_number",
    "initiation date": "created_date",
    "sign-off date": "sign_off_date",
    "step id": "step_id",
    "status": "status",
    "responsible site": "responsible_site",
    "title": "title",
    "nc owner": "nc_owner",
    "nc coordinator": "nc_coordinator",
    "nc related to": "nc_related_to",
    "closed date": "closed_date"
})

df_full["created_date"] = pd.to_datetime(df_full["created_date"], errors="coerce")
df_full["sign_off_date"] = pd.to_datetime(df_full["sign_off_date"], errors="coerce")

# === EXTENSIONES ===
# Se considera UNA extensión por día (no por hora) y nunca más de dos en total
df_ext = df_full[df_full["step_id"] == "tApproveDueDateExtension"].copy()
df_ext["sign_off_day"] = df_ext["sign_off_date"].dt.date

extensions = (
    df_ext.drop_duplicates(subset=["nc_number", "sign_off_day"])
          .groupby("nc_number")["sign_off_day"]
          .count()
          .rename("extension_count")
)

# Unir extensiones con el dataset principal
df_unique = df_unique.merge(extensions, on="nc_number", how="left")
df_unique["extension_count"] = df_unique["extension_count"].fillna(0).astype(int)

# Limitar máximo a 2 extensiones (según el proceso)
df_unique["extension_count"] = df_unique["extension_count"].clip(upper=2)

# === FECHA DE CIERRE PREVISTA ===
# Cálculo: fecha de creación + 90 días iniciales + 90 por cada extensión
# (el día de creación cuenta como día 1)
df_unique["planned_closure_date"] = df_unique["created_date"] + pd.to_timedelta(
    90 * (1 + df_unique["extension_count"]) - 1, unit="D"
)

# Usar esta fecha para cálculo de overdue
df_unique["calculated_completion_date"] = df_unique["planned_closure_date"]

# ==========================
# DASHBOARD
# ==========================
st.header("NC DASHBOARD")

# ==========================
# SECTION 1: Indicators
# ==========================
st.subheader("Indicators")

df_unique["is_closed"] = df_unique["status"].str.lower().eq("closed")
df_unique["is_inworks"] = ~df_unique["is_closed"]

total_global = df_unique["nc_number"].nunique()
inworks_global = df_unique.loc[df_unique["is_inworks"], "nc_number"].nunique()
closed_global = df_unique.loc[df_unique["is_closed"], "nc_number"].nunique()

df_1100 = df_unique[df_unique["responsible_site"] == "1100"]
total_1100 = df_1100["nc_number"].nunique()
inworks_1100 = df_1100.loc[df_1100["is_inworks"], "nc_number"].nunique()
closed_1100 = df_1100.loc[df_1100["is_closed"], "nc_number"].nunique()

col1, col2 = st.columns(2)
with col1:
    st.markdown("**General Numbers - Global**")
    st.table(pd.DataFrame({
        "Item": ["Total NC", "Total In works", "Total Closed"],
        "Number": [total_global, inworks_global, closed_global]
    }))
with col2:
    st.markdown("**General Numbers - Site 1100**")
    st.table(pd.DataFrame({
        "Item": ["Total NC", "Total In works", "Total Closed"],
        "Number": [total_1100, inworks_1100, closed_1100]
    }))

# ==========================
# SECTION 2: Overdue Summary
# ==========================
st.subheader("Overdue Summary (2020–2025)")

start_year, end_year = 2020, 2025
mask_window = df_unique["created_date"].dt.year.between(start_year, end_year)
df_win = df_unique[mask_window].copy()

# Clasificar solo INWORKS
def classify_inworks(frame):
    f = frame[frame["is_inworks"]].copy()
    due = f["calculated_completion_date"]
    f["due_status"] = "On time"
    f.loc[(due.notna()) & (due < TODAY), "due_status"] = "Overdue"
    f.loc[due.isna(), "due_status"] = "No due date"
    return f

# Aplicar clasificación
finw = classify_inworks(df_win)

# Tablas resumen
tbl_inw_global = pd.DataFrame({
    "Status": ["On time", "Overdue"],
    "NC": [
        finw.loc[finw["due_status"] == "On time", "nc_number"].nunique(),
        finw.loc[finw["due_status"] == "Overdue", "nc_number"].nunique()
    ]
})

tbl_inw_1100 = pd.DataFrame({
    "Status": ["On time", "Overdue"],
    "NC": [
        finw.loc[(finw["due_status"] == "On time") & (finw["responsible_site"] == "1100"), "nc_number"].nunique(),
        finw.loc[(finw["due_status"] == "Overdue") & (finw["responsible_site"] == "1100"), "nc_number"].nunique()
    ]
})

c1, c2 = st.columns(2)
with c1:
    st.markdown("**In works — Global**")
    st.table(tbl_inw_global)
with c2:
    st.markdown("**In works — Site 1100**")
    st.table(tbl_inw_1100)

# ==========================
# TABLE: Overdue Details
# ==========================
st.subheader("Overdue Details")

# Filtrar solo las NC Overdue
overdue_nc = finw.loc[finw["due_status"] == "Overdue", "nc_number"].unique()

df_overdue = (
    df_full[df_full["nc_number"].isin(overdue_nc)]
    .drop_duplicates(subset=["nc_number"])
    .merge(
        df_unique[["nc_number", "planned_closure_date", "extension_count"]],
        on="nc_number",
        how="left"
    )
)

df_overdue["extension_count"] = df_overdue["extension_count"].fillna(0).astype(int)

# Seleccionar y renombrar columnas finales
df_overdue = df_overdue[[
    "nc_number",
    "title",
    "nc_owner",
    "nc_coordinator",
    "nc_related_to",
    "responsible_site",
    "created_date",
    "planned_closure_date",
    "extension_count"
]].rename(columns={
    "nc_number": "NC Number",
    "title": "Title",
    "nc_owner": "NC Owner",
    "nc_coordinator": "NC Coordinator",
    "nc_related_to": "NC Related To",
    "responsible_site": "Responsible Site",
    "created_date": "Initiation Date",
    "planned_closure_date": "Closed Date (planned)",
    "extension_count": "Due date extensions"
})

# Formato de fechas tipo "15-Mar-25"
df_overdue["Initiation Date"] = pd.to_datetime(df_overdue["Initiation Date"], errors="coerce").dt.strftime("%d-%b-%y")
df_overdue["Closed Date (planned)"] = pd.to_datetime(df_overdue["Closed Date (planned)"], errors="coerce").dt.strftime("%d-%b-%y")

# Mostrar tabla final (sin índice)
st.dataframe(
    df_overdue.sort_values("Closed Date (planned)").reset_index(drop=True),
    use_container_width=True,
    hide_index=True
)
# ==========================
# TABLE: Next Overdue (Due in next 2 months)
# ==========================
st.subheader("Next Overdue (Due within 2 months)")

# Definir rango de fechas (desde hoy hasta 60 días adelante)
next_due_start = TODAY
next_due_end = TODAY + pd.Timedelta(days=60)

# Filtrar NC INWORKS que vencen dentro de los próximos 2 meses
upcoming_due = finw.loc[
    (finw["is_inworks"]) &
    (finw["calculated_completion_date"].notna()) &
    (finw["calculated_completion_date"] >= next_due_start) &
    (finw["calculated_completion_date"] <= next_due_end)
].copy()

# Ordenar por fecha de vencimiento (más próxima primero)
upcoming_due = upcoming_due.sort_values("calculated_completion_date")

# Unir información adicional desde df_full
df_upcoming = (
    df_full[df_full["nc_number"].isin(upcoming_due["nc_number"])]
    .drop_duplicates(subset=["nc_number"])
    .merge(
        df_unique[["nc_number", "planned_closure_date", "extension_count"]],
        on="nc_number",
        how="left"
    )
)

df_upcoming["extension_count"] = df_upcoming["extension_count"].fillna(0).astype(int)

# Seleccionar columnas finales
df_upcoming = df_upcoming[[
    "nc_number",
    "title",
    "nc_owner",
    "nc_coordinator",
    "nc_related_to",
    "responsible_site",
    "created_date",
    "planned_closure_date",
    "extension_count"
]].rename(columns={
    "nc_number": "NC Number",
    "title": "Title",
    "nc_owner": "NC Owner",
    "nc_coordinator": "NC Coordinator",
    "nc_related_to": "NC Related To",
    "responsible_site": "Responsible Site",
    "created_date": "Initiation Date",
    "planned_closure_date": "Closed Date (planned)",
    "extension_count": "Due date extensions"
})

# Formato de fechas
df_upcoming["Initiation Date"] = pd.to_datetime(df_upcoming["Initiation Date"], errors="coerce").dt.strftime("%d-%b-%y")
df_upcoming["Closed Date (planned)"] = pd.to_datetime(df_upcoming["Closed Date (planned)"], errors="coerce").dt.strftime("%d-%b-%y")

# Mostrar tabla final ordenada por fecha más próxima
st.dataframe(
    df_upcoming.sort_values("Closed Date (planned)").reset_index(drop=True),
    use_container_width=True,
    hide_index=True
)

# ==========================
# SECTION 3: Ages
# ==========================
#st.subheader("Age Distribution")

#def age_brackets(frame):
    #f = frame[frame["is_inworks"]].copy()
    #f = f[f["created_date"].notna()].copy()
    #f["age_days"] = (TODAY - f["created_date"]).dt.days
    #age_data = {
        #"Item": ["Older than 2 years", "Older than 1 year"],
        #"Number": [
            #int(f.loc[f["age_days"] > 730, "nc_number"].nunique()),
            #int(f.loc[f["age_days"] > 365, "nc_number"].nunique())
       # ]
    #}
   # return pd.DataFrame(age_data)

#col_age1, col_age2 = st.columns(2)
#with col_age1:
    #st.markdown("**Global**")
    #st.table(age_brackets(df_win))
#with col_age2:
    #st.markdown("**Site 1100**")
    #st.table(age_brackets(df_win[df_win["responsible_site"] == "1100"]))
# ==========================
# SECTION 4: Annual Trends
# ==========================
st.subheader("Annual Trends")

def annual_trends(frame, title=""):
    years = list(range(2020, TODAY.year + 1))

    created_per_year = frame.groupby(frame["created_date"].dt.year)["nc_number"].nunique()
    closed_per_year = frame.dropna(subset=["closed_date"]).groupby(frame["closed_date"].dt.year)["nc_number"].nunique()

    inworks_counts = {}
    for y in years:
        cutoff = pd.Timestamp(year=y, month=12, day=31)
        inworks_counts[y] = frame[
            (frame["created_date"] <= cutoff) &
            ((frame["closed_date"].isna()) | (frame["closed_date"] > cutoff))
        ]["nc_number"].nunique()

    data = pd.DataFrame({
        "Year": years,
        "Created": created_per_year.reindex(years, fill_value=0).values,
        "Closed": closed_per_year.reindex(years, fill_value=0).values,
        "In works": pd.Series(inworks_counts).reindex(years).values
    })

    fig = go.Figure()
    fig.add_bar(x=data["Year"], y=data["Created"], name="Created", text=data["Created"], textposition="outside")
    fig.add_bar(x=data["Year"], y=data["Closed"], name="Closed", text=data["Closed"], textposition="outside")
    fig.add_trace(go.Scatter(x=data["Year"], y=data["In works"], mode="lines+markers+text",
                             name="In works", text=data["In works"], textposition="top center"))
    fig.update_layout(title=title, barmode="group", legend=dict(orientation="h", y=-0.2))
    return fig

fig_global = annual_trends(df_win, title="Global")
fig_1100 = annual_trends(df_win[df_win["responsible_site"] == "1100"], title="Site 1100")

cg, c1100 = st.columns(2)
with cg:
    st.plotly_chart(fig_global, use_container_width=True)
with c1100:
    st.plotly_chart(fig_1100, use_container_width=True)
