import math
import time
from datetime import datetime, timedelta, time as dtime
from io import BytesIO
from urllib.parse import quote_plus

import folium
import pandas as pd
import requests
import streamlit as st
from geopy.geocoders import Nominatim
from streamlit_folium import st_folium

st.set_page_config(page_title="Routage Pro Terrain V7", page_icon="🚗", layout="wide")

st.title("🚗 Routage Pro Terrain V7")
st.caption("Import Excel → tournée optimisée → distances, temps, carte, Waze, Google Maps, appel et export")

# -----------------------------
# Helpers
# -----------------------------

def normalize_col(c):
    return str(c).strip().lower().replace("é", "e").replace("è", "e").replace("ê", "e").replace("à", "a").replace("ç", "c")


def find_column(df, candidates):
    norm_map = {normalize_col(c): c for c in df.columns}
    for cand in candidates:
        nc = normalize_col(cand)
        if nc in norm_map:
            return norm_map[nc]
    # fuzzy contains
    for col in df.columns:
        n = normalize_col(col)
        for cand in candidates:
            if normalize_col(cand) in n:
                return col
    return None


def build_full_address(row, address_col=None, cp_col=None, city_col=None):
    parts = []
    if address_col and pd.notna(row.get(address_col, "")):
        parts.append(str(row[address_col]).strip())
    if cp_col and pd.notna(row.get(cp_col, "")):
        cp = str(row[cp_col]).strip().replace(".0", "")
        parts.append(cp)
    if city_col and pd.notna(row.get(city_col, "")):
        parts.append(str(row[city_col]).strip())
    return ", ".join([p for p in parts if p and p.lower() != "nan"])


def haversine_km(lat1, lon1, lat2, lon2):
    R = 6371.0
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)
    a = math.sin(dphi/2)**2 + math.cos(phi1)*math.cos(phi2)*math.sin(dlambda/2)**2
    return 2 * R * math.atan2(math.sqrt(a), math.sqrt(1-a))


def estimate_drive_from_air_km(air_km):
    # approximation terrain France : route souvent 20-35% plus longue que ligne droite
    road_km = air_km * 1.30
    # moyenne prudente : 55 km/h urbain/périurbain
    minutes = max(3, int(round((road_km / 55) * 60)))
    return road_km, minutes


def osrm_route(points):
    """Return distances/durations between consecutive points and full geometry if OSRM works."""
    if len(points) < 2:
        return [], [], []
    coords = ";".join([f"{lon},{lat}" for lat, lon in points])
    url = f"https://router.project-osrm.org/route/v1/driving/{coords}?overview=full&geometries=geojson&steps=false"
    try:
        r = requests.get(url, timeout=20)
        if r.status_code != 200:
            return None, None, None
        data = r.json()
        if not data.get("routes"):
            return None, None, None
        route = data["routes"][0]
        legs = route.get("legs", [])
        distances_km = [leg.get("distance", 0) / 1000 for leg in legs]
        durations_min = [int(round(leg.get("duration", 0) / 60)) for leg in legs]
        geometry = route.get("geometry", {}).get("coordinates", [])
        # convert lon,lat to lat,lon for folium
        line = [(lat, lon) for lon, lat in geometry]
        return distances_km, durations_min, line
    except Exception:
        return None, None, None


def nearest_neighbor_order(points, start_idx=0):
    # points = list of dict with lat/lon. start_idx is virtual if separate start point not in prospects
    remaining = list(range(len(points)))
    order = []
    current_lat = st.session_state.start_lat
    current_lon = st.session_state.start_lon
    while remaining:
        best = min(remaining, key=lambda i: haversine_km(current_lat, current_lon, points[i]["lat"], points[i]["lon"]))
        order.append(best)
        current_lat, current_lon = points[best]["lat"], points[best]["lon"]
        remaining.remove(best)
    return order


def parse_time_value(v):
    if pd.isna(v):
        return None
    if isinstance(v, datetime):
        return v.time()
    if isinstance(v, dtime):
        return v
    s = str(v).strip()
    if not s:
        return None
    for fmt in ["%H:%M:%S", "%H:%M", "%Hh%M", "%Hh"]:
        try:
            return datetime.strptime(s, fmt).time()
        except Exception:
            pass
    # Excel sometimes full datetime as string
    try:
        return pd.to_datetime(s).time()
    except Exception:
        return None


def clean_phone(v):
    if pd.isna(v):
        return ""
    s = str(v).strip().replace(".0", "")
    return s


def maps_link(address):
    return "https://www.google.com/maps/search/?api=1&query=" + quote_plus(address)


def waze_link(lat, lon):
    return f"https://waze.com/ul?ll={lat},{lon}&navigate=yes"


def tel_link(phone):
    p = "".join(ch for ch in str(phone) if ch.isdigit() or ch == "+")
    return f"tel:{p}" if p else ""


def to_excel(df_detail, df_source):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_detail.to_excel(writer, index=False, sheet_name="TOURNEE_DETAIL")
        df_source.to_excel(writer, index=False, sheet_name="DONNEES_ORIGINE")
        workbook = writer.book
        for sheet_name in ["TOURNEE_DETAIL", "DONNEES_ORIGINE"]:
            ws = writer.sheets[sheet_name]
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, 0, max(0, len((df_detail if sheet_name == 'TOURNEE_DETAIL' else df_source).columns)-1))
            for idx, col in enumerate((df_detail if sheet_name == 'TOURNEE_DETAIL' else df_source).columns):
                width = min(max(len(str(col)) + 2, 12), 35)
                ws.set_column(idx, idx, width)
    output.seek(0)
    return output

# -----------------------------
# Sidebar settings
# -----------------------------
with st.sidebar:
    st.header("⚙️ Paramètres")
    start_address = st.text_input("Adresse de départ", value="Maisons-Alfort, France")
    start_time = st.time_input("Heure de départ", value=dtime(8, 30))
    visit_minutes = st.number_input("Durée moyenne sur place / RDV (min)", min_value=0, max_value=240, value=45, step=5)
    respect_times = st.checkbox("Respecter les horaires de RDV si présents", value=True)
    st.info("Waze ne prend pas les étapes multiples : l'app génère un bouton Waze pour chaque prochain client.")

uploaded = st.file_uploader("📤 Upload ton fichier Excel", type=["xlsx", "xls"])

if "geocode_cache" not in st.session_state:
    st.session_state.geocode_cache = {}

if uploaded:
    try:
        df = pd.read_excel(uploaded)
    except Exception as e:
        st.error(f"Impossible de lire le fichier Excel : {e}")
        st.stop()

    if df.empty:
        st.warning("Le fichier est vide.")
        st.stop()

    st.subheader("1️⃣ Colonnes détectées")
    col1, col2, col3, col4 = st.columns(4)

    default_address = find_column(df, ["adresse", "address", "rue", "adresse client", "adresse complete"])
    default_cp = find_column(df, ["code postal", "cp", "post_code", "postcode", "zip"])
    default_city = find_column(df, ["ville", "city", "commune"])
    default_name = find_column(df, ["prospect", "nom", "full name", "client", "contact", "prenom nom"])
    default_phone = find_column(df, ["telephone", "téléphone", "phone", "tel", "mobile"])
    default_time = find_column(df, ["heure", "horaire", "heure rdv", "rdv", "date heure", "date/heure", "appointment"])
    default_sales = find_column(df, ["commercial", "vendeur", "technicien"])

    cols = [None] + list(df.columns)
    with col1:
        address_col = st.selectbox("Colonne adresse", cols, index=cols.index(default_address) if default_address in cols else 0)
        cp_col = st.selectbox("Colonne CP", cols, index=cols.index(default_cp) if default_cp in cols else 0)
    with col2:
        city_col = st.selectbox("Colonne ville", cols, index=cols.index(default_city) if default_city in cols else 0)
        name_col = st.selectbox("Colonne nom prospect", cols, index=cols.index(default_name) if default_name in cols else 0)
    with col3:
        phone_col = st.selectbox("Colonne téléphone", cols, index=cols.index(default_phone) if default_phone in cols else 0)
        time_col = st.selectbox("Colonne heure RDV", cols, index=cols.index(default_time) if default_time in cols else 0)
    with col4:
        sales_col = st.selectbox("Colonne commercial", cols, index=cols.index(default_sales) if default_sales in cols else 0)

    st.dataframe(df.head(20), use_container_width=True)

    if st.button("🚀 Calculer la tournée", type="primary"):
        try:
            geolocator = Nominatim(user_agent="routage_pro_froid24_v7")

            with st.spinner("Géocodage de l'adresse de départ..."):
                start_key = "START::" + start_address
                if start_key in st.session_state.geocode_cache:
                    st.session_state.start_lat, st.session_state.start_lon = st.session_state.geocode_cache[start_key]
                else:
                    loc = geolocator.geocode(start_address, country_codes="fr", timeout=15)
                    if not loc:
                        st.error("Adresse de départ introuvable. Essaie d'ajouter ville + code postal.")
                        st.stop()
                    st.session_state.start_lat, st.session_state.start_lon = loc.latitude, loc.longitude
                    st.session_state.geocode_cache[start_key] = (loc.latitude, loc.longitude)
                    time.sleep(1)

            prospects = []
            progress = st.progress(0)
            total = len(df)
            bad_rows = []

            for idx, row in df.iterrows():
                progress.progress(min(1.0, (idx + 1) / max(1, total)), text=f"Géocodage {idx+1}/{total}")
                full_addr = build_full_address(row, address_col, cp_col, city_col)
                if not full_addr:
                    bad_rows.append((idx + 2, "Adresse vide"))
                    continue
                if full_addr in st.session_state.geocode_cache:
                    lat, lon = st.session_state.geocode_cache[full_addr]
                else:
                    try:
                        loc = geolocator.geocode(full_addr, country_codes="fr", timeout=15)
                    except Exception as e:
                        loc = None
                    if not loc:
                        bad_rows.append((idx + 2, full_addr))
                        continue
                    lat, lon = loc.latitude, loc.longitude
                    st.session_state.geocode_cache[full_addr] = (lat, lon)
                    time.sleep(1)  # Nominatim usage policy
                prospects.append({
                    "source_index": idx,
                    "adresse_complete": full_addr,
                    "lat": lat,
                    "lon": lon,
                    "nom": str(row[name_col]).strip() if name_col and pd.notna(row.get(name_col, "")) else f"Prospect {len(prospects)+1}",
                    "telephone": clean_phone(row[phone_col]) if phone_col else "",
                    "heure_rdv": parse_time_value(row[time_col]) if time_col else None,
                    "commercial": str(row[sales_col]).strip() if sales_col and pd.notna(row.get(sales_col, "")) else "",
                    "row": row,
                })
            progress.empty()

            if not prospects:
                st.error("Aucune adresse n'a pu être géocodée.")
                if bad_rows:
                    st.write("Adresses non trouvées :", bad_rows[:20])
                st.stop()

            # Order
            if respect_times and any(p["heure_rdv"] for p in prospects):
                with_time = [p for p in prospects if p["heure_rdv"]]
                no_time = [p for p in prospects if not p["heure_rdv"]]
                with_time.sort(key=lambda p: p["heure_rdv"])
                # Insert no-time prospects by nearest neighbor before/after timed list using simple approach: nearest order for no-time appended
                if no_time:
                    temp = no_time
                    order_idx = nearest_neighbor_order(temp)
                    ordered = with_time + [temp[i] for i in order_idx]
                else:
                    ordered = with_time
            else:
                order_idx = nearest_neighbor_order(prospects)
                ordered = [prospects[i] for i in order_idx]

            all_points = [(st.session_state.start_lat, st.session_state.start_lon)] + [(p["lat"], p["lon"]) for p in ordered]
            distances, durations, route_line = osrm_route(all_points)
            used_osrm = distances is not None and durations is not None
            if not used_osrm:
                distances, durations = [], []
                for i in range(len(all_points)-1):
                    air = haversine_km(all_points[i][0], all_points[i][1], all_points[i+1][0], all_points[i+1][1])
                    d, m = estimate_drive_from_air_km(air)
                    distances.append(d)
                    durations.append(m)
                route_line = all_points

            # Build detail
            detail_rows = []
            current_dt = datetime.combine(datetime.today(), start_time)
            cumulative_km = 0.0
            cumulative_drive_min = 0
            for i, p in enumerate(ordered, start=1):
                leg_km = distances[i-1] if i-1 < len(distances) else 0
                leg_min = durations[i-1] if i-1 < len(durations) else 0
                current_dt = current_dt + timedelta(minutes=leg_min)
                arrival = current_dt
                rdv_time = p["heure_rdv"]
                ecart = ""
                ecart_min = None
                if rdv_time:
                    rdv_dt = datetime.combine(datetime.today(), rdv_time)
                    ecart_min = int(round((arrival - rdv_dt).total_seconds() / 60))
                    if ecart_min > 0:
                        ecart = f"Retard {ecart_min} min"
                    elif ecart_min < 0:
                        ecart = f"Avance {abs(ecart_min)} min"
                    else:
                        ecart = "À l'heure"
                cumulative_km += leg_km
                cumulative_drive_min += leg_min
                detail_rows.append({
                    "Ordre": i,
                    "Prospect": p["nom"],
                    "Téléphone": p["telephone"],
                    "Commercial": p["commercial"],
                    "Adresse": p["adresse_complete"],
                    "RDV prévu": rdv_time.strftime("%H:%M") if rdv_time else "",
                    "Arrivée estimée": arrival.strftime("%H:%M"),
                    "Écart RDV": ecart,
                    "Distance depuis étape précédente (km)": round(leg_km, 1),
                    "Temps de route depuis étape précédente": f"{leg_min} min",
                    "Distance cumulée (km)": round(cumulative_km, 1),
                    "Temps route cumulé": f"{cumulative_drive_min} min",
                    "Latitude": p["lat"],
                    "Longitude": p["lon"],
                    "Lien Waze": waze_link(p["lat"], p["lon"]),
                    "Lien Google Maps": maps_link(p["adresse_complete"]),
                    "Statut terrain": "À faire",
                })
                current_dt = current_dt + timedelta(minutes=int(visit_minutes))

            detail_df = pd.DataFrame(detail_rows)
            st.session_state["detail_df"] = detail_df
            st.session_state["ordered"] = ordered
            st.session_state["route_line"] = route_line
            st.session_state["used_osrm"] = used_osrm
            st.session_state["bad_rows"] = bad_rows
            st.session_state["source_df"] = df
            st.success("✅ Tournée calculée")
        except Exception as e:
            st.exception(e)

# -----------------------------
# Results persistent display
# -----------------------------
if "detail_df" in st.session_state:
    detail_df = st.session_state["detail_df"]
    ordered = st.session_state["ordered"]
    route_line = st.session_state["route_line"]

    st.subheader("2️⃣ Résumé tournée")
    total_km = detail_df["Distance depuis étape précédente (km)"].sum()
    total_min = sum(int(str(x).replace(" min", "")) for x in detail_df["Temps de route depuis étape précédente"])
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("RDV géocodés", len(detail_df))
    c2.metric("Distance totale", f"{total_km:.1f} km")
    c3.metric("Temps route total", f"{total_min//60}h{total_min%60:02d}")
    c4.metric("Type calcul", "Route réelle OSRM" if st.session_state.get("used_osrm") else "Approximation secours")

    bad_rows = st.session_state.get("bad_rows", [])
    if bad_rows:
        with st.expander(f"⚠️ {len(bad_rows)} adresse(s) non trouvée(s)"):
            st.write(pd.DataFrame(bad_rows, columns=["Ligne Excel", "Adresse / erreur"]))

    st.subheader("3️⃣ Détail étape par étape")
    st.dataframe(detail_df, use_container_width=True, hide_index=True)

    st.subheader("4️⃣ Mode terrain")
    for _, row in detail_df.iterrows():
        with st.container(border=True):
            st.markdown(f"### {int(row['Ordre'])}. {row['Prospect']}")
            st.write(f"📍 {row['Adresse']}")
            st.write(f"🚗 Depuis étape précédente : **{row['Distance depuis étape précédente (km)']} km** — **{row['Temps de route depuis étape précédente']}**")
            if row['RDV prévu']:
                st.write(f"🕘 RDV prévu : **{row['RDV prévu']}** | arrivée estimée : **{row['Arrivée estimée']}** | {row['Écart RDV']}")
            else:
                st.write(f"🕘 Arrivée estimée : **{row['Arrivée estimée']}**")
            b1, b2, b3 = st.columns(3)
            b1.link_button("🚗 Ouvrir Waze", row["Lien Waze"], use_container_width=True)
            b2.link_button("🗺️ Google Maps", row["Lien Google Maps"], use_container_width=True)
            if row["Téléphone"]:
                b3.link_button("📞 Appeler", tel_link(row["Téléphone"]), use_container_width=True)
            else:
                b3.button("📞 Pas de téléphone", disabled=True, use_container_width=True)

    st.subheader("5️⃣ Carte avec noms prospects")
    start_lat, start_lon = st.session_state.start_lat, st.session_state.start_lon
    m = folium.Map(location=[start_lat, start_lon], zoom_start=9)
    folium.Marker(
        [start_lat, start_lon],
        popup="Départ",
        tooltip="Départ",
        icon=folium.Icon(color="green", icon="home")
    ).add_to(m)
    if route_line and len(route_line) >= 2:
        folium.PolyLine(route_line, weight=5, opacity=0.8, tooltip="Itinéraire").add_to(m)
    for _, row in detail_df.iterrows():
        label = f"{int(row['Ordre'])} - {row['Prospect']}"
        popup_html = f"""
        <b>{label}</b><br>
        {row['Adresse']}<br>
        Distance étape: {row['Distance depuis étape précédente (km)']} km<br>
        Temps étape: {row['Temps de route depuis étape précédente']}<br>
        Arrivée estimée: {row['Arrivée estimée']}
        """
        folium.Marker(
            [row["Latitude"], row["Longitude"]],
            popup=folium.Popup(popup_html, max_width=320),
            tooltip=label,
            icon=folium.DivIcon(html=f"""
                <div style="font-size:12px;font-weight:bold;color:white;background:#1f77b4;border-radius:14px;padding:4px 8px;border:2px solid white;box-shadow:0 1px 4px #333;white-space:nowrap;">
                {label}
                </div>
            """)
        ).add_to(m)
    st_folium(m, width=None, height=650)

    st.subheader("6️⃣ Export")
    excel = to_excel(detail_df, st.session_state.get("source_df", pd.DataFrame()))
    st.download_button(
        "📥 Télécharger l'Excel de tournée V7",
        data=excel,
        file_name="tournee_routage_pro_V7.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
else:
    st.info("Upload ton Excel puis clique sur Calculer la tournée.")
