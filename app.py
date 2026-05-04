import io
import math
import time
from datetime import datetime, timedelta, time as dtime
from typing import Dict, List, Optional, Tuple
from urllib.parse import quote_plus

import folium
import pandas as pd
import requests
import streamlit as st
from streamlit_folium import st_folium

st.set_page_config(page_title="Routage RDV depuis Excel", page_icon="🚗", layout="wide")

NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
OSRM_ROUTE_URL = "https://router.project-osrm.org/route/v1/driving"
USER_AGENT = "routage-rdv-excel/1.0"

DEFAULT_COLUMNS = {
    "adresse": "adresse_du_prospect",
    "cp": "code_postal_du_prospect",
    "ville": "ville_du_prospect",
    "date": "date_rendez_vous",
    "heure": "debut",
    "nom": "nom_du_prospect",
    "prenom": "prenom_du_prospect",
    "tel": "telephone_du_prospect",
    "email": "email_du_prospect",
    "commercial_nom": "nom_du_commercial",
    "commercial_prenom": "prenom_du_commercial",
    "statut": "statut_prospect",
}


def normalize_colname(name: str) -> str:
    return str(name).strip().lower().replace(" ", "_")


def find_column(df: pd.DataFrame, candidates: List[str], fallback: Optional[str] = None) -> Optional[str]:
    normalized = {normalize_colname(c): c for c in df.columns}
    for cand in candidates:
        key = normalize_colname(cand)
        if key in normalized:
            return normalized[key]
    return fallback


def clean_phone(value) -> str:
    if pd.isna(value):
        return ""
    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]
    if len(s) == 9 and not s.startswith("0"):
        s = "0" + s
    return s


def parse_time(value) -> Optional[dtime]:
    if pd.isna(value):
        return None
    if isinstance(value, dtime):
        return value
    if isinstance(value, datetime):
        return value.time()
    s = str(value).strip()
    for fmt in ["%H:%M:%S", "%H:%M", "%Hh%M", "%Hh"]:
        try:
            return datetime.strptime(s, fmt).time()
        except ValueError:
            pass
    return None


def parse_date(value) -> Optional[datetime]:
    if pd.isna(value):
        return None
    if isinstance(value, datetime):
        return value
    s = str(value).strip()
    for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"]:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    try:
        return pd.to_datetime(value, dayfirst=True).to_pydatetime()
    except Exception:
        return None


def build_full_address(row: pd.Series, col_addr: str, col_cp: Optional[str], col_city: Optional[str]) -> str:
    parts = []
    for col in [col_addr, col_cp, col_city]:
        if col and col in row and not pd.isna(row[col]):
            val = str(row[col]).strip()
            if val and val.lower() != "nan":
                if val.endswith(".0"):
                    val = val[:-2]
                parts.append(val)
    return ", ".join(parts) + ", France"


@st.cache_data(show_spinner=False)
def geocode_address(address: str) -> Optional[Tuple[float, float, str]]:
    params = {"q": address, "format": "json", "limit": 1, "countrycodes": "fr"}
    try:
        r = requests.get(NOMINATIM_URL, params=params, headers={"User-Agent": USER_AGENT}, timeout=20)
        r.raise_for_status()
        data = r.json()
        if not data:
            return None
        item = data[0]
        # Important: Nominatim = lat/lon. OSRM attend lon,lat.
        return float(item["lat"]), float(item["lon"]), item.get("display_name", address)
    except Exception:
        return None


@st.cache_data(show_spinner=False)
def osrm_route(a_lat: float, a_lon: float, b_lat: float, b_lon: float) -> Optional[Dict[str, float]]:
    coords = f"{a_lon},{a_lat};{b_lon},{b_lat}"
    params = {"overview": "false", "alternatives": "false", "steps": "false"}
    try:
        r = requests.get(f"{OSRM_ROUTE_URL}/{coords}", params=params, timeout=25)
        r.raise_for_status()
        data = r.json()
        if data.get("code") != "Ok" or not data.get("routes"):
            return None
        route = data["routes"][0]
        return {"distance_km": route["distance"] / 1000, "duration_min": route["duration"] / 60}
    except Exception:
        return None


def haversine_km(a_lat, a_lon, b_lat, b_lon) -> float:
    r = 6371
    phi1 = math.radians(a_lat)
    phi2 = math.radians(b_lat)
    dphi = math.radians(b_lat - a_lat)
    dlambda = math.radians(b_lon - a_lon)
    x = math.sin(dphi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda / 2) ** 2
    return 2 * r * math.atan2(math.sqrt(x), math.sqrt(1 - x))


def nearest_neighbor_order(points: pd.DataFrame, start_lat: float, start_lon: float) -> List[int]:
    remaining = set(points.index.tolist())
    order = []
    cur_lat, cur_lon = start_lat, start_lon
    while remaining:
        nxt = min(
            remaining,
            key=lambda idx: haversine_km(cur_lat, cur_lon, points.loc[idx, "latitude"], points.loc[idx, "longitude"]),
        )
        order.append(nxt)
        remaining.remove(nxt)
        cur_lat, cur_lon = points.loc[nxt, "latitude"], points.loc[nxt, "longitude"]
    return order


def compute_route(df: pd.DataFrame, start_address: str, start_time: dtime, mode: str, pause_minutes: int) -> Tuple[pd.DataFrame, Dict[str, float], folium.Map]:
    start_geo = geocode_address(start_address)
    if not start_geo:
        raise ValueError("Adresse de départ introuvable. Essaie avec une adresse plus complète.")
    start_lat, start_lon, start_label = start_geo

    valid = df[df["geocoding_ok"]].copy()
    invalid = df[~df["geocoding_ok"]].copy()

    if valid.empty:
        raise ValueError("Aucune adresse n'a pu être géocodée.")

    if mode == "Optimiser la distance":
        order_idx = nearest_neighbor_order(valid, start_lat, start_lon)
        valid = valid.loc[order_idx].copy()
    else:
        valid = valid.sort_values(["date_rdv", "heure_rdv"], na_position="last").copy()

    current_dt = datetime.combine(valid["date_rdv"].dropna().min().date() if valid["date_rdv"].notna().any() else datetime.today().date(), start_time)
    prev_lat, prev_lon = start_lat, start_lon
    rows = []
    total_km = 0.0
    total_min = 0.0

    for i, (_, row) in enumerate(valid.iterrows(), start=1):
        route = osrm_route(prev_lat, prev_lon, row["latitude"], row["longitude"])
        if route is None:
            km = haversine_km(prev_lat, prev_lon, row["latitude"], row["longitude"]) * 1.30
            minutes = km / 55 * 60
            source = "Estimation à vol d'oiseau x1,30"
        else:
            km = route["distance_km"]
            minutes = route["duration_min"]
            source = "OSRM"
        current_dt = current_dt + timedelta(minutes=minutes)
        rdv_dt = None
        if pd.notna(row.get("date_rdv")) and row.get("heure_rdv"):
            rdv_dt = datetime.combine(row["date_rdv"].date(), row["heure_rdv"])
        retard_min = None
        avance_min = None
        if rdv_dt:
            diff = (current_dt - rdv_dt).total_seconds() / 60
            retard_min = max(0, round(diff))
            avance_min = max(0, round(-diff))
        enriched = row.to_dict()
        enriched.update({
            "ordre_tournee": i,
            "depart_depuis": "Départ" if i == 1 else "RDV précédent",
            "distance_depuis_precedent_km": round(km, 1),
            "temps_depuis_precedent_min": round(minutes),
            "heure_arrivee_estimee": current_dt.strftime("%H:%M"),
            "retard_estime_min": retard_min,
            "avance_estimee_min": avance_min,
            "source_calcul": source,
            "lien_google_maps": "https://www.google.com/maps/search/?api=1&query=" + quote_plus(row["adresse_complete"]),
        })
        rows.append(enriched)
        total_km += km
        total_min += minutes
        current_dt = current_dt + timedelta(minutes=pause_minutes)
        prev_lat, prev_lon = row["latitude"], row["longitude"]

    result = pd.DataFrame(rows)
    if not invalid.empty:
        invalid = invalid.copy()
        invalid["ordre_tournee"] = "Non calculé"
        invalid["source_calcul"] = "Adresse non trouvée"
        result = pd.concat([result, invalid], ignore_index=True, sort=False)

    m = folium.Map(location=[start_lat, start_lon], zoom_start=8)
    folium.Marker([start_lat, start_lon], tooltip="Départ", popup=start_label, icon=folium.Icon(color="green")).add_to(m)
    line = [[start_lat, start_lon]]
    for _, row in result[result["geocoding_ok"]].sort_values("ordre_tournee").iterrows():
        folium.Marker(
            [row["latitude"], row["longitude"]],
            tooltip=f"#{row['ordre_tournee']} - {row.get('prenom_du_prospect', '')} {row.get('nom_du_prospect', '')}",
            popup=f"{row['adresse_complete']}<br>Arrivée estimée : {row.get('heure_arrivee_estimee', '')}",
        ).add_to(m)
        line.append([row["latitude"], row["longitude"]])
    folium.PolyLine(line, weight=4, opacity=0.8).add_to(m)

    summary = {"total_km": round(total_km, 1), "total_trajet_min": round(total_min), "nb_rdv": int(valid.shape[0]), "adresses_non_trouvees": int(invalid.shape[0])}
    return result, summary, m


def to_excel_bytes(df: pd.DataFrame, summary: Dict[str, float]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df = pd.DataFrame([
            ["Nombre de RDV calculés", summary["nb_rdv"]],
            ["Distance totale estimée km", summary["total_km"]],
            ["Temps total de trajet min", summary["total_trajet_min"]],
            ["Adresses non trouvées", summary["adresses_non_trouvees"]],
        ], columns=["Indicateur", "Valeur"])
        summary_df.to_excel(writer, sheet_name="Résumé", index=False)
        df.to_excel(writer, sheet_name="Tournée", index=False)
        for sheet in writer.book.worksheets:
            sheet.freeze_panes = "A2"
            for col in sheet.columns:
                max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col[:100])
                sheet.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 12), 45)
    return output.getvalue()


st.title("🚗 Appli de routage depuis un fichier Excel")
st.caption("Importe tes RDV, calcule une tournée, affiche distance/temps/heure estimée et exporte un Excel enrichi.")

with st.sidebar:
    st.header("Paramètres")
    start_address = st.text_input("Adresse de départ", value="Maisons-Alfort, France")
    start_time_input = st.time_input("Heure de départ", value=dtime(8, 0))
    pause_minutes = st.number_input("Temps sur place par RDV, en minutes", min_value=0, max_value=240, value=45, step=5)
    mode = st.radio("Mode de tournée", ["Respecter les heures de RDV", "Optimiser la distance"])
    st.info("Astuce : pour une vraie tournée commerciale, commence par respecter les heures de RDV. Pour une prospection sans horaire fixe, optimise la distance.")

uploaded = st.file_uploader("Charge ton fichier Excel", type=["xlsx", "xls"])

if uploaded:
    df = pd.read_excel(uploaded)
    df.columns = [str(c).strip() for c in df.columns]

    col_addr = find_column(df, [DEFAULT_COLUMNS["adresse"], "adresse", "address"])
    col_cp = find_column(df, [DEFAULT_COLUMNS["cp"], "code_postal", "cp", "postcode"])
    col_city = find_column(df, [DEFAULT_COLUMNS["ville"], "ville", "city"])
    col_date = find_column(df, [DEFAULT_COLUMNS["date"], "date", "date_rdv"])
    col_time = find_column(df, [DEFAULT_COLUMNS["heure"], "heure", "heure_rdv", "debut", "début"])

    if not col_addr:
        st.error("Je ne trouve pas de colonne adresse. Renomme une colonne en `adresse_du_prospect` ou sélectionne un fichier compatible.")
        st.stop()

    st.subheader("Aperçu du fichier")
    st.dataframe(df.head(20), use_container_width=True)

    if st.button("Calculer la tournée", type="primary"):
        work = df.copy()
        work["adresse_complete"] = work.apply(lambda r: build_full_address(r, col_addr, col_cp, col_city), axis=1)
        work["date_rdv"] = work[col_date].apply(parse_date) if col_date else pd.NaT
        work["heure_rdv"] = work[col_time].apply(parse_time) if col_time else None
        if "telephone_du_prospect" in work.columns:
            work["telephone_du_prospect"] = work["telephone_du_prospect"].apply(clean_phone)

        progress = st.progress(0, text="Géocodage des adresses...")
        lats, lons, labels, oks = [], [], [], []
        for i, address in enumerate(work["adresse_complete"].tolist()):
            geo = geocode_address(address)
            if geo:
                lat, lon, label = geo
                lats.append(lat); lons.append(lon); labels.append(label); oks.append(True)
            else:
                lats.append(None); lons.append(None); labels.append(""); oks.append(False)
            progress.progress((i + 1) / max(len(work), 1), text=f"Géocodage {i + 1}/{len(work)}")
            time.sleep(1.0)  # Respect du service gratuit Nominatim.
        work["latitude"] = lats
        work["longitude"] = lons
        work["adresse_trouvee"] = labels
        work["geocoding_ok"] = oks

        with st.spinner("Calcul des distances et temps de trajet..."):
            result, summary, fmap = compute_route(work, start_address, start_time_input, mode, int(pause_minutes))

        st.success("Tournée calculée")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("RDV calculés", summary["nb_rdv"])
        c2.metric("Distance totale", f"{summary['total_km']} km")
        c3.metric("Temps de trajet", f"{summary['total_trajet_min']} min")
        c4.metric("Adresses non trouvées", summary["adresses_non_trouvees"])

        preferred_cols = [
            "ordre_tournee", "date_rdv", "heure_rdv", "heure_arrivee_estimee", "retard_estime_min", "avance_estimee_min",
            "prenom_du_prospect", "nom_du_prospect", "telephone_du_prospect", "adresse_complete",
            "distance_depuis_precedent_km", "temps_depuis_precedent_min", "source_calcul", "lien_google_maps"
        ]
        cols = [c for c in preferred_cols if c in result.columns] + [c for c in result.columns if c not in preferred_cols]
        st.subheader("Résultat")
        st.dataframe(result[cols], use_container_width=True)

        st.subheader("Carte")
        st_folium(fmap, width=None, height=520)

        excel_bytes = to_excel_bytes(result[cols], summary)
        st.download_button(
            "Télécharger l'Excel de tournée",
            data=excel_bytes,
            file_name="tournee_rdv_calculee.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Charge ton fichier Excel pour commencer. Le fichier fourni dans ce dossier `exemple_rdv.xlsx` est déjà compatible.")
