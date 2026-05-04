import io
import math
import time
import urllib.parse
from datetime import datetime, time as dtime

import pandas as pd
import requests
import streamlit as st
import folium
import streamlit.components.v1 as components

st.set_page_config(page_title="Routage PRO Excel V5", page_icon="🚗", layout="wide")
OSM_HEADERS = {"User-Agent": "RoutageProFroid24/4.0 (contact: mauricefroid24@gmail.com)"}


def normalize_text(x):
    return str(x).strip().lower().replace("é", "e").replace("è", "e").replace("ê", "e").replace("à", "a").replace("ç", "c")


def find_col(df, candidates):
    for c in df.columns:
        n = normalize_text(c)
        for cand in candidates:
            if cand in n:
                return c
    return None


def parse_time(value):
    if value is None or pd.isna(value) or value == "":
        return None
    if isinstance(value, pd.Timestamp):
        return value.time()
    if isinstance(value, datetime):
        return value.time()
    if isinstance(value, dtime):
        return value
    s = str(value).strip().lower().replace("h", ":")
    for fmt in ["%H:%M:%S", "%H:%M", "%H"]:
        try:
            return datetime.strptime(s, fmt).time()
        except Exception:
            pass
    return None


def minutes_from_time(t):
    if t is None:
        return None
    return t.hour * 60 + t.minute


def format_minutes(minutes):
    if minutes is None or pd.isna(minutes):
        return ""
    minutes = int(round(minutes)) % (24 * 60)
    return f"{minutes // 60:02d}:{minutes % 60:02d}"


def haversine_km(lat1, lon1, lat2, lon2):
    R = 6371
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat / 2) ** 2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon / 2) ** 2
    return 2 * R * math.asin(math.sqrt(a))


@st.cache_data(show_spinner=False, ttl=24 * 3600)
def geocode_one(address):
    url = "https://nominatim.openstreetmap.org/search"
    params = {"q": address, "format": "json", "limit": 1, "countrycodes": "fr", "addressdetails": 1}
    r = requests.get(url, params=params, headers=OSM_HEADERS, timeout=15)
    r.raise_for_status()
    data = r.json()
    if not data:
        return None
    item = data[0]
    return float(item["lat"]), float(item["lon"]), item.get("display_name", "")


@st.cache_data(show_spinner=False, ttl=3600)
def osrm_matrix(coords_tuple):
    coords = list(coords_tuple)
    coord_str = ";".join([f"{lon},{lat}" for lat, lon in coords])
    url = f"https://router.project-osrm.org/table/v1/driving/{coord_str}"
    params = {"annotations": "duration,distance"}
    r = requests.get(url, params=params, timeout=20)
    r.raise_for_status()
    data = r.json()
    if "durations" not in data or "distances" not in data:
        raise ValueError("Réponse OSRM invalide")
    durations = [[0 if x is None else int(x / 60) for x in row] for row in data["durations"]]
    distances = [[0 if x is None else round(x / 1000, 1) for x in row] for row in data["distances"]]
    return durations, distances


def fallback_matrix(coords):
    n = len(coords)
    distances = [[0] * n for _ in range(n)]
    durations = [[0] * n for _ in range(n)]
    for i in range(n):
        for j in range(n):
            if i == j:
                continue
            km = haversine_km(coords[i][0], coords[i][1], coords[j][0], coords[j][1]) * 1.30
            distances[i][j] = round(km, 1)
            durations[i][j] = max(1, int(km / 55 * 60))
    return durations, distances


def nearest_order(durations, start=0):
    n = len(durations)
    remaining = set(range(1, n))
    order = [start]
    current = start
    while remaining:
        nxt = min(remaining, key=lambda j: durations[current][j])
        order.append(nxt)
        remaining.remove(nxt)
        current = nxt
    return order


def google_maps_link(addresses):
    clean = [a for a in addresses if a]
    if len(clean) < 2:
        return ""
    # Google limite les waypoints, donc on garde les 10 premières étapes pour éviter un lien cassé.
    clean = clean[:10]
    origin = urllib.parse.quote_plus(clean[0])
    destination = urllib.parse.quote_plus(clean[-1])
    waypoints = "|".join(urllib.parse.quote_plus(a) for a in clean[1:-1])
    url = f"https://www.google.com/maps/dir/?api=1&origin={origin}&destination={destination}&travelmode=driving"
    if waypoints:
        url += f"&waypoints={waypoints}"
    return url


st.title("🚗 Appli de routage depuis un fichier Excel — V5 stable")
st.caption("Version stable : les résultats restent affichés après calcul, même si la carte recharge la page.")

with st.sidebar:
    st.header("Paramètres")
    start_address = st.text_input("Adresse de départ", "Maisons-Alfort, France")
    start_time_input = st.time_input("Heure de départ", value=dtime(7, 0))
    visit_duration = st.number_input("Temps sur place par RDV, en minutes", min_value=0, max_value=240, value=90, step=5)
    mode = st.radio("Mode de tournée", ["Respecter les heures de RDV", "Optimiser la distance"], index=0)
    return_to_start = st.checkbox("Retour au point de départ", value=False)

uploaded = st.file_uploader("Charge ton fichier Excel", type=["xlsx", "xls"])

if uploaded:
    df = pd.read_excel(uploaded)
    st.subheader("Aperçu du fichier")
    st.dataframe(df.head(20), use_container_width=True)

    adresse_col = find_col(df, ["adresse", "address", "rue"])
    cp_col = find_col(df, ["code postal", "postcode", "post_code", "cp"])
    ville_col = find_col(df, ["ville", "city", "commune"])
    heure_col = find_col(df, ["heure", "horaire", "creneau", "date", "rdv"])
    prospect_col = find_col(df, ["prospect", "client", "nom", "full name", "prenom"])
    phone_col = find_col(df, ["phone", "tel", "telephone"])
    commercial_col = find_col(df, ["commercial", "vendeur", "closer"])

    with st.expander("Vérifier / modifier les colonnes détectées", expanded=False):
        cols = [None] + list(df.columns)
        c1, c2, c3 = st.columns(3)
        with c1:
            adresse_col = st.selectbox("Adresse", cols, index=cols.index(adresse_col) if adresse_col in cols else 0)
            cp_col = st.selectbox("Code postal", cols, index=cols.index(cp_col) if cp_col in cols else 0)
        with c2:
            ville_col = st.selectbox("Ville", cols, index=cols.index(ville_col) if ville_col in cols else 0)
            heure_col = st.selectbox("Heure/date RDV", cols, index=cols.index(heure_col) if heure_col in cols else 0)
        with c3:
            prospect_col = st.selectbox("Nom / prospect", cols, index=cols.index(prospect_col) if prospect_col in cols else 0)
            phone_col = st.selectbox("Téléphone", cols, index=cols.index(phone_col) if phone_col in cols else 0)

    if st.button("Calculer la tournée", type="primary"):
        try:
            if not adresse_col:
                st.error("Sélectionne au minimum la colonne Adresse.")
                st.stop()

            work = df.copy().reset_index(drop=True)

            def full_address(row):
                parts = []
                for col in [adresse_col, cp_col, ville_col]:
                    if col and col in row and pd.notna(row[col]):
                        p = str(row[col]).strip()
                        if p and p.lower() != "nan":
                            parts.append(p)
                return ", ".join(parts)

            work["Adresse complete"] = work.apply(full_address, axis=1)
            work = work[work["Adresse complete"].astype(str).str.len() > 3].copy().reset_index(drop=True)
            if work.empty:
                st.error("Aucune adresse exploitable trouvée dans le fichier.")
                st.stop()

            progress = st.progress(0)
            status = st.empty()
            points = []
            all_addresses = [start_address] + work["Adresse complete"].tolist()

            for i, addr in enumerate(all_addresses):
                status.info(f"Géocodage {i + 1}/{len(all_addresses)} : {addr}")
                geo = geocode_one(addr)
                if geo is None:
                    points.append({"address": addr, "lat": None, "lon": None})
                else:
                    lat, lon, display = geo
                    points.append({"address": addr, "lat": lat, "lon": lon, "display": display})
                progress.progress((i + 1) / len(all_addresses))
                time.sleep(0.05)

            geocoded = pd.DataFrame(points)
            bad = geocoded[geocoded["lat"].isna()]
            if len(bad) > 0:
                st.error("Certaines adresses n’ont pas été trouvées. Corrige-les puis relance.")
                st.dataframe(bad[["address"]], use_container_width=True)
                st.stop()

            coords = tuple(zip(geocoded["lat"], geocoded["lon"]))
            status.info("Calcul des temps et distances...")
            try:
                durations, distances = osrm_matrix(coords)
                calcul_source = "Routes OSRM"
            except Exception as e:
                durations, distances = fallback_matrix(coords)
                calcul_source = "Calcul de secours sans trafic"
                st.warning(f"Le service de routage OSRM n’a pas répondu. J’utilise un calcul de secours. Détail : {e}")

            if mode == "Respecter les heures de RDV" and heure_col:
                rdv_infos = []
                for i, row in work.iterrows():
                    t = parse_time(row[heure_col])
                    rdv_infos.append((i + 1, minutes_from_time(t) if t else 99999))
                order = [0] + [idx for idx, _ in sorted(rdv_infos, key=lambda x: x[1])]
            else:
                order = nearest_order(durations, 0)

            if return_to_start:
                order.append(0)

            current_min = minutes_from_time(start_time_input)
            total_drive = 0
            total_km = 0.0
            prev = None
            rows = []

            for step, idx in enumerate(order):
                if prev is None:
                    drive_min, dist_km = 0, 0.0
                else:
                    drive_min, dist_km = durations[prev][idx], distances[prev][idx]
                    current_min += drive_min
                    total_drive += drive_min
                    total_km += dist_km

                rdv_min = None
                wait = 0
                if idx == 0:
                    nom, tel, address, type_row = "Base", "", start_address, "Départ/Retour"
                    rdv_file = ""
                    duration_place = 0
                else:
                    row = work.iloc[idx - 1]
                    nom = str(row[prospect_col]) if prospect_col and pd.notna(row[prospect_col]) else ""
                    tel = str(row[phone_col]) if phone_col and pd.notna(row[phone_col]) else ""
                    address = row["Adresse complete"]
                    type_row = "RDV"
                    duration_place = int(visit_duration)
                    t = parse_time(row[heure_col]) if heure_col else None
                    rdv_min = minutes_from_time(t)
                    rdv_file = format_minutes(rdv_min) if rdv_min is not None else ""
                    if mode == "Respecter les heures de RDV" and rdv_min is not None and current_min < rdv_min:
                        wait = rdv_min - current_min
                        current_min = rdv_min

                retard = ""
                if rdv_min is not None:
                    diff = current_min - rdv_min
                    retard = "+{} min".format(diff) if diff > 0 else ("{} min".format(diff) if diff < 0 else "à l'heure")

                rows.append({
                    "Ordre": step,
                    "Type": type_row,
                    "Nom": nom,
                    "Téléphone": tel,
                    "Adresse complete": address,
                    "Trajet depuis étape précédente (min)": drive_min,
                    "Distance depuis étape précédente (km)": dist_km,
                    "Attente avant RDV (min)": wait,
                    "Heure arrivée estimée": format_minutes(current_min),
                    "Heure RDV fichier": rdv_file,
                    "Avance / retard": retard,
                    "Temps sur place prévu (min)": duration_place,
                    "Latitude": geocoded.iloc[idx]["lat"],
                    "Longitude": geocoded.iloc[idx]["lon"],
                })

                if idx != 0:
                    current_min += int(visit_duration)
                prev = idx

            result = pd.DataFrame(rows)
            gmaps = google_maps_link(result["Adresse complete"].tolist())
            summary = {
                "Nombre de RDV": len(work),
                "Distance totale estimée": f"{round(total_km, 1)} km",
                "Temps de conduite estimé": f"{int(total_drive//60)}h{int(total_drive%60):02d}",
                "Heure fin estimée": format_minutes(current_min),
                "Source calcul": calcul_source,
                "Lien Google Maps": gmaps,
            }
            st.session_state["last_result"] = result
            st.session_state["last_summary"] = summary
            st.session_state["last_calcul_source"] = calcul_source
            st.session_state["last_total_km"] = total_km
            st.session_state["last_total_drive"] = total_drive
            st.session_state["last_end_min"] = current_min
            st.session_state["last_gmaps"] = gmaps
            status.success("Tournée calculée. Les résultats sont affichés plus bas.")

        except Exception as e:
            st.error("Erreur pendant le calcul. Voici le détail technique pour correction :")
            st.exception(e)
    if "last_result" in st.session_state:
        result = st.session_state["last_result"]
        summary = st.session_state.get("last_summary", {})
        st.divider()
        st.success(f"Tournée calculée avec succès — source : {st.session_state.get('last_calcul_source', '')}")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("RDV", max(0, len(result) - 1))
        c2.metric("Distance", f"{round(st.session_state.get('last_total_km', 0), 1)} km")
        total_drive = st.session_state.get("last_total_drive", 0)
        c3.metric("Conduite", f"{int(total_drive//60)}h{int(total_drive%60):02d}")
        c4.metric("Fin estimée", format_minutes(st.session_state.get("last_end_min", 0)))

        st.subheader("Résultat de la tournée")
        display_cols = [
            "Ordre", "Type", "Nom", "Téléphone", "Adresse complete",
            "Trajet depuis étape précédente (min)", "Distance depuis étape précédente (km)",
            "Heure arrivée estimée", "Heure RDV fichier", "Avance / retard",
            "Attente avant RDV (min)", "Temps sur place prévu (min)"
        ]
        st.dataframe(result[[c for c in display_cols if c in result.columns]], use_container_width=True, height=420)

        gmaps = st.session_state.get("last_gmaps", "")
        if gmaps:
            st.link_button("Ouvrir la tournée dans Google Maps", gmaps, type="primary")

        st.subheader("Carte")
        try:
            center_lat = float(result["Latitude"].mean())
            center_lon = float(result["Longitude"].mean())
            m = folium.Map(location=[center_lat, center_lon], zoom_start=9)
            route_coords = []
            for _, r in result.iterrows():
                route_coords.append([r["Latitude"], r["Longitude"]])
                folium.Marker(
                    [r["Latitude"], r["Longitude"]],
                    tooltip=f"{r['Ordre']} - {r['Type']}",
                    popup=f"{r['Ordre']} - {r['Nom']}<br>{r['Adresse complete']}<br>Arrivée : {r['Heure arrivée estimée']}",
                ).add_to(m)
            folium.PolyLine(route_coords, weight=4, opacity=0.8).add_to(m)
            components.html(m._repr_html_(), height=560, scrolling=True)
        except Exception as e:
            st.warning(f"Carte non affichée, mais la tournée est bien calculée. Détail : {e}")

        st.subheader("Export")
        export = io.BytesIO()
        with pd.ExcelWriter(export, engine="openpyxl") as writer:
            result.to_excel(writer, sheet_name="Tournee optimisee", index=False)
            pd.DataFrame(summary.items(), columns=["Indicateur", "Valeur"]).to_excel(writer, sheet_name="Resume", index=False)
        export.seek(0)
        st.download_button(
            "Télécharger l’Excel optimisé",
            data=export,
            file_name="tournee_optimisee.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Dépose ton Excel pour commencer.")
