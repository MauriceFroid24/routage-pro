import io
import re
import math
from pathlib import Path
from datetime import datetime, date, time as dtime, timedelta
from urllib.parse import quote_plus

import pandas as pd
import streamlit as st
import requests
from PIL import Image as PILImage
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from geopy.distance import geodesic
import folium
from streamlit_folium import st_folium
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

st.set_page_config(page_title="Routage PRO V17", page_icon="🚗", layout="wide")

DEFAULT_START = "72 avenue des Tourelles, 94490 Ormesson-sur-Marne"
AVG_SPEED_KMH = 38
LAST_UPLOAD_PATH = Path("/tmp/routage_pro_dernier_fichier.xlsx")

COLS = {
    "numero_rdv": 0, "adresse": 1, "code_postal": 2, "date_rdv": 3, "heure_debut": 4,
    "email": 5, "fournisseur": 7, "commercial_nom": 8, "nom": 9, "telepros_nom": 11,
    "commercial_prenom": 12, "prenom": 13, "telepros_prenom": 14, "telephone": 16, "ville": 17,
}

st.title("🚗 Routage PRO V17 — terrain + IK comptable")
st.caption("Mode sombre lisible · carte claire · calcul automatique · module indemnités kilométriques PDF")

st.markdown("""
<style>
.stApp {
    background-color:#050505;
    color:#f5f5f5;
}
[data-testid="stSidebar"] {
    background-color:#0f1115;
}
.block-container {
    padding-top: 1.2rem;
}
.stAlert {
    border-radius:14px;
    background-color:#171717 !important;
    color:#ffffff !important;
    border:1px solid #333333 !important;
}
.stAlert p, .stAlert div {
    color:#ffffff !important;
}
[data-testid="stMetric"] {
    background-color:#161616;
    border-radius:14px;
    padding:12px;
    border:1px solid #333333;
}
[data-testid="stMetricValue"], [data-testid="stMetricLabel"] {
    color:#ffffff !important;
}
div[data-testid="stDataFrame"] {
    border-radius:14px;
    overflow:hidden;
    border:1px solid #333333;
}
.stButton button, .stDownloadButton button, .stLinkButton a {
    border-radius:14px !important;
    font-weight:800 !important;
    background-color:#222222 !important;
    color:#ffffff !important;
    border:1px solid #444444 !important;
}
.stLinkButton a {
    text-decoration:none !important;
}
hr {
    border-color:#333333 !important;
}

/* V15 — corrections PC : bandeau haut + expanders blancs sur blanc */
[data-testid="stHeader"] {
    background: #050505 !important;
    color: #ffffff !important;
    border-bottom: 1px solid #1f1f1f !important;
}
[data-testid="stToolbar"] {
    background: transparent !important;
}
[data-testid="stDecoration"] {
    background: #050505 !important;
}
[data-testid="stExpander"] {
    background-color: #121212 !important;
    border: 1px solid #3a3a3a !important;
    border-radius: 14px !important;
    overflow: hidden !important;
    margin-bottom: 10px !important;
}
[data-testid="stExpander"] details {
    background-color: #121212 !important;
    color: #ffffff !important;
}
[data-testid="stExpander"] summary {
    background-color: #1f2937 !important;
    color: #ffffff !important;
    font-weight: 900 !important;
    font-size: 1.02rem !important;
}
[data-testid="stExpander"] summary * {
    color: #ffffff !important;
}
[data-testid="stExpander"] div,
[data-testid="stExpander"] p,
[data-testid="stExpander"] span {
    color: #ffffff !important;
}
[data-testid="stExpander"] svg {
    fill: #ffffff !important;
    color: #ffffff !important;
}
/* Force les textes Streamlit à rester lisibles sur PC */
.stMarkdown, .stMarkdown p, .stMarkdown span, label, div[data-testid="stText"] {
    color: #f5f5f5 !important;
}

</style>
""", unsafe_allow_html=True)



def safe_get(row, idx):
    try:
        v = row.iloc[idx]
        if pd.isna(v):
            return ""
        return str(v).strip()
    except Exception:
        return ""


def parse_date(v):
    if pd.isna(v) or v == "":
        return None
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y"]:
        try:
            return datetime.strptime(s[:10], fmt).date()
        except Exception:
            pass
    try:
        return pd.to_datetime(s, dayfirst=True).date()
    except Exception:
        return None


def parse_time(v):
    if pd.isna(v) or v == "":
        return None
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.time().replace(second=0, microsecond=0)
    if isinstance(v, dtime):
        return v.replace(second=0, microsecond=0)
    if isinstance(v, (int, float)) and 0 <= v < 1:
        total_minutes = int(round(v * 24 * 60))
        return dtime(total_minutes // 60, total_minutes % 60)
    s = str(v).strip().replace("h", ":").replace("H", ":")
    if re.match(r"^\d{1,2}:\d{2}(:\d{2})?$", s):
        h, m = [int(x) for x in s.split(":")[:2]]
        return dtime(h, m)
    if re.match(r"^\d{1,2}$", s):
        return dtime(int(s), 0)
    try:
        return pd.to_datetime(s).time().replace(second=0, microsecond=0)
    except Exception:
        return None


def dt_from_row(d, t):
    if isinstance(d, date) and isinstance(t, dtime):
        return datetime.combine(d, t)
    return None


def format_phone(raw):
    digits = re.sub(r"\D", "", str(raw or ""))
    if digits.startswith("33") and len(digits) == 11:
        digits = "0" + digits[2:]
    if len(digits) == 9 and digits[0] in "123456789":
        digits = "0" + digits
    if len(digits) == 10:
        return " ".join([digits[0:2], digits[2:4], digits[4:6], digits[6:8], digits[8:10]]), digits
    return str(raw or ""), digits


def full_name(prenom, nom):
    parts = [str(x).strip() for x in [prenom, nom] if str(x).strip() and str(x).strip().lower() != "nan"]
    return " ".join(parts) or "Prospect"


def build_address(adresse, cp, ville):
    return ", ".join([str(x).strip() for x in [adresse, cp, ville] if str(x).strip()])


def waze_link(lat, lon, address):
    # Pour Waze, on privilégie toujours l'adresse complète plutôt que les coordonnées.
    # Les coordonnées issues du géocodage peuvent parfois pointer seulement la rue,
    # alors que l'adresse encodée conserve mieux le numéro de rue dans l'application.
    clean_address = str(address or "").strip()
    if clean_address:
        return f"https://www.waze.com/ul?q={quote_plus(clean_address)}&navigate=yes"
    if lat and lon:
        return f"https://www.waze.com/ul?ll={lat},{lon}&navigate=yes"
    return "https://www.waze.com/"


def maps_link(address):
    return f"https://www.google.com/maps/search/?api=1&query={quote_plus(address)}"


def streetview_link(lat, lon, address):
    # Lien volontairement simple et fiable sur iPhone/Windows : ouvre Google Maps sur l’adresse.
    # Les liens Street View directs sont instables et peuvent donner un écran noir selon Safari/Chrome.
    return f"https://www.google.com/maps/search/{quote_plus(address)}"


def directions_link(origin, destination):
    return f"https://www.google.com/maps/dir/?api=1&origin={quote_plus(origin)}&destination={quote_plus(destination)}&travelmode=driving"


def fmt_date(x):
    if isinstance(x, date):
        return x.strftime("%d/%m/%Y")
    return str(x) if x else ""


def fmt_time(x):
    if isinstance(x, dtime):
        return x.strftime("%H:%M")
    return str(x) if x else ""


def fmt_dt(x):
    if isinstance(x, datetime):
        return x.strftime("%H:%M")
    return ""


def fmt_duration(m):
    if m == "" or m is None:
        return ""
    try:
        m = int(round(float(m)))
    except Exception:
        return ""
    return f"{m//60}h{m%60:02d}" if m >= 60 else f"{m} min"


@st.cache_data(show_spinner=False)
def geocode_one(address):
    """Géocodage robuste pour la France : API adresse.data.gouv.fr puis Nominatim."""
    if not address:
        return {"lat": None, "lon": None, "source": "vide"}
    q = str(address).strip()
    # 1) API officielle française, très fiable avec adresse + CP + ville
    try:
        r = requests.get(
            "https://api-adresse.data.gouv.fr/search/",
            params={"q": q, "limit": 1, "autocomplete": 0},
            timeout=8,
        )
        data = r.json()
        feats = data.get("features", [])
        if feats:
            lon, lat = feats[0].get("geometry", {}).get("coordinates", [None, None])
            if lat and lon:
                return {"lat": float(lat), "lon": float(lon), "source": "adresse.data.gouv.fr"}
    except Exception:
        pass
    # 2) Fallback Nominatim
    try:
        geolocator = Nominatim(user_agent="routage_pro_v13_froid24")
        loc = geolocator.geocode(q + ", France", timeout=8)
        if loc:
            return {"lat": loc.latitude, "lon": loc.longitude, "source": "Nominatim"}
    except Exception:
        pass
    return {"lat": None, "lon": None, "source": "non trouvé"}


@st.cache_data(show_spinner=False)
def geocode_addresses(addresses):
    out = {}
    seen = []
    for a in addresses:
        if a and a not in seen:
            seen.append(a)
    for address in seen:
        out[address] = geocode_one(address)
    return out


@st.cache_data(show_spinner=False)
def osrm_route(origin_lat, origin_lon, dest_lat, dest_lon):
    if not all([origin_lat, origin_lon, dest_lat, dest_lon]):
        return None
    url = f"https://router.project-osrm.org/route/v1/driving/{origin_lon},{origin_lat};{dest_lon},{dest_lat}?overview=full&geometries=geojson"
    try:
        r = requests.get(url, timeout=8)
        data = r.json()
        if data.get("code") == "Ok" and data.get("routes"):
            route = data["routes"][0]
            coords = []
            try:
                coords = [[lat, lon] for lon, lat in route.get("geometry", {}).get("coordinates", [])]
            except Exception:
                coords = []
            return {"km": route["distance"] / 1000, "min": route["duration"] / 60, "source": "OSRM", "geometry": coords}
    except Exception:
        return None
    return None


@st.cache_data(show_spinner=False)
def google_distance_matrix(origin, destination, arrival_dt, api_key):
    if not api_key or not arrival_dt:
        return None
    try:
        departure = max(datetime.now(), arrival_dt - timedelta(hours=2))
        params = {
            "origins": origin,
            "destinations": destination,
            "mode": "driving",
            "departure_time": int(departure.timestamp()),
            "key": api_key,
        }
        r = requests.get("https://maps.googleapis.com/maps/api/distancematrix/json", params=params, timeout=10)
        data = r.json()
        el = data["rows"][0]["elements"][0]
        if el.get("status") == "OK":
            dur = el.get("duration_in_traffic", el.get("duration", {})).get("value", 0) / 60
            dist = el.get("distance", {}).get("value", 0) / 1000
            return {"km": dist, "min": dur, "source": "Google trafic"}
    except Exception:
        return None
    return None


def traffic_factor(arrival_dt):
    if not isinstance(arrival_dt, datetime):
        return 1.25
    h = arrival_dt.hour + arrival_dt.minute / 60
    if 7 <= h <= 10 or 16.5 <= h <= 20:
        return 1.55
    if 11 <= h <= 16.5:
        return 1.25
    return 1.12


def prepare_dataframe(file):
    df = pd.read_excel(file, header=0)
    rows = []
    for _, row in df.iterrows():
        adresse = safe_get(row, COLS["adresse"])
        cp = safe_get(row, COLS["code_postal"])
        ville = safe_get(row, COLS["ville"])
        if not adresse and not ville:
            continue
        d = parse_date(row.iloc[COLS["date_rdv"]] if len(row) > COLS["date_rdv"] else "")
        h = parse_time(row.iloc[COLS["heure_debut"]] if len(row) > COLS["heure_debut"] else "")
        phone_fmt, phone_digits = format_phone(safe_get(row, COLS["telephone"]))
        telepros_full = full_name(safe_get(row, COLS["telepros_prenom"]), safe_get(row, COLS["telepros_nom"]))
        rows.append({
            "numero_rdv": safe_get(row, COLS["numero_rdv"]),
            "nom_prospect": full_name(safe_get(row, COLS["prenom"]), safe_get(row, COLS["nom"])),
            "adresse": adresse,
            "code_postal": cp,
            "ville": ville,
            "adresse_complete": build_address(adresse, cp, ville),
            "date_rdv": d,
            "heure_rdv": h,
            "rdv_datetime": dt_from_row(d, h),
            "telephone": phone_fmt,
            "telephone_tel": phone_digits,
            "email": safe_get(row, COLS["email"]),
            "fournisseur": safe_get(row, COLS["fournisseur"]),
            "commercial": full_name(safe_get(row, COLS["commercial_prenom"]), safe_get(row, COLS["commercial_nom"])),
            "teleprospecteur": telepros_full,
        })
    result = pd.DataFrame(rows)
    if not result.empty:
        result["sort_date"] = result["date_rdv"].apply(lambda x: x or date.max)
        result["sort_time"] = result["heure_rdv"].apply(lambda x: x or dtime.max)
        result = result.sort_values(["sort_date", "sort_time", "numero_rdv"]).drop(columns=["sort_date", "sort_time"]).reset_index(drop=True)
        # Conservation du numéro RDV original + renumérotation terrain dans l’ordre chronologique
        if "numero_rdv_source" not in result.columns:
            result["numero_rdv_source"] = result["numero_rdv"]
        result["numero_rdv"] = range(1, len(result) + 1)
        result.insert(0, "ordre", range(1, len(result) + 1))
    return result


def renumber_route_df(df):
    """Force la numérotation terrain 1, 2, 3... dans l'ordre chronologique réel."""
    if df is None or df.empty:
        return df
    out = df.copy()
    # Garde le numéro d'origine du fichier si présent, mais n'utilise plus ce numéro pour l'affichage terrain.
    if "numero_rdv_source" not in out.columns and "numero_rdv" in out.columns:
        out["numero_rdv_source"] = out["numero_rdv"]
    # Tri le plus fiable possible : date/heure RDV puis ordre existant si disponible.
    sort_cols = []
    if "rdv_datetime" in out.columns:
        out["__sort_dt"] = pd.to_datetime(out["rdv_datetime"], errors="coerce")
        sort_cols.append("__sort_dt")
    if "date_rdv" in out.columns:
        out["__sort_date"] = pd.to_datetime(out["date_rdv"], errors="coerce")
        sort_cols.append("__sort_date")
    if "heure_rdv" in out.columns:
        out["__sort_time"] = out["heure_rdv"].apply(lambda x: fmt_time(x) if pd.notna(x) else "99:99")
        sort_cols.append("__sort_time")
    if "ordre" in out.columns:
        out["__sort_ordre"] = pd.to_numeric(out["ordre"], errors="coerce").fillna(999999)
        sort_cols.append("__sort_ordre")
    if sort_cols:
        out = out.sort_values(sort_cols, na_position="last").reset_index(drop=True)
    out["numero_rdv"] = list(range(1, len(out) + 1))
    out["ordre"] = list(range(1, len(out) + 1))
    out = out.drop(columns=[c for c in out.columns if c.startswith("__sort_")], errors="ignore")
    return out


def route_between(prev_addr, prev_geo, addr, coord, arrival_dt, api_key, use_google):
    if use_google and api_key:
        g = google_distance_matrix(prev_addr, addr, arrival_dt, api_key)
        if g:
            return g
    o = osrm_route(prev_geo.get("lat"), prev_geo.get("lon"), coord.get("lat"), coord.get("lon"))
    if o:
        return o
    if prev_geo.get("lat") and prev_geo.get("lon") and coord.get("lat") and coord.get("lon"):
        dist = geodesic((prev_geo["lat"], prev_geo["lon"]), (coord["lat"], coord["lon"])).km * 1.28
        mins = (dist / AVG_SPEED_KMH) * 60
        return {"km": dist, "min": mins, "source": "Estimation", "geometry": []}
    return {"km": None, "min": None, "source": "Non calculé", "geometry": []}


def enrich_route(df, start_address, safety_min, visit_min, use_google, api_key):
    addresses = [start_address] + df["adresse_complete"].tolist()
    geo = geocode_addresses(addresses)
    prev_addr = start_address
    prev_geo = geo.get(start_address, {})
    previous_rdv_end = None
    out = []
    cumulative_km = 0.0
    cumulative_min = 0.0

    for _, row in df.iterrows():
        addr = row["adresse_complete"]
        coord = geo.get(addr, {})
        arrival_dt = row.get("rdv_datetime")
        rb = route_between(prev_addr, prev_geo, addr, coord, arrival_dt, api_key, use_google)
        km = rb.get("km")
        raw_min = rb.get("min")
        if raw_min is not None:
            if rb.get("source") == "Google trafic":
                drive_min = int(math.ceil(raw_min))
                traffic_note = "trafic Google"
            else:
                drive_min = int(math.ceil(raw_min * traffic_factor(arrival_dt)))
                traffic_note = "trafic estimé"
        else:
            drive_min = None
            traffic_note = "non calculé"

        advised_departure = arrival_dt - timedelta(minutes=(drive_min or 0) + safety_min) if arrival_dt and drive_min is not None else None
        if previous_rdv_end and advised_departure:
            pause_min = int((advised_departure - previous_rdv_end).total_seconds() // 60)
        else:
            pause_min = None
        previous_rdv_end = arrival_dt + timedelta(minutes=visit_min) if arrival_dt else None
        cumulative_km += km or 0
        cumulative_min += drive_min or 0

        r = row.to_dict()
        r.update({
            "lat": coord.get("lat"), "lon": coord.get("lon"), "source_geocodage": coord.get("source", ""),
            "distance_depuis_precedent_km": round(km, 1) if km is not None else "",
            "temps_route_depuis_precedent_min": drive_min if drive_min is not None else "",
            "source_temps": rb.get("source", ""),
            "note_trafic": traffic_note,
            "depart_conseille": advised_departure,
            "marge_securite_min": safety_min,
            "pause_avant_rdv_min": pause_min if pause_min is not None else "",
            "distance_cumulee_km": round(cumulative_km, 1),
            "temps_route_cumule_min": int(cumulative_min),
            "waze": waze_link(coord.get("lat"), coord.get("lon"), addr),
            "google_maps": maps_link(addr),
            "street_view": streetview_link(coord.get("lat"), coord.get("lon"), addr),
            "itineraire_depuis_precedent": directions_link(prev_addr, addr),
            "route_geometry": rb.get("geometry", []),
        })
        out.append(r)
        prev_addr = addr
        prev_geo = coord

    route_df = renumber_route_df(pd.DataFrame(out))

    # Retour base après le dernier RDV
    return_row = None
    if not route_df.empty:
        last = route_df.iloc[-1]
        last_addr = last["adresse_complete"]
        last_geo = {"lat": last.get("lat"), "lon": last.get("lon")}
        last_end = last.get("rdv_datetime") + timedelta(minutes=visit_min) if isinstance(last.get("rdv_datetime"), datetime) else None
        rb = route_between(last_addr, last_geo, start_address, geo.get(start_address, {}), last_end, api_key, use_google)
        km = rb.get("km")
        raw_min = rb.get("min")
        ret_min = int(math.ceil(raw_min * (1 if rb.get("source") == "Google trafic" else traffic_factor(last_end)))) if raw_min is not None else ""
        return_row = {
            "ordre": "Retour", "numero_rdv": "BASE", "date_rdv": last.get("date_rdv", ""), "heure_rdv": "",
            "rdv_datetime": last_end, "nom_prospect": "Retour base", "telephone": "", "telephone_tel": "",
            "adresse_complete": start_address, "lat": geo.get(start_address, {}).get("lat"), "lon": geo.get(start_address, {}).get("lon"),
            "distance_depuis_precedent_km": round(km, 1) if km is not None else "",
            "temps_route_depuis_precedent_min": ret_min,
            "source_temps": rb.get("source", ""), "note_trafic": "retour inclus",
            "depart_conseille": last_end, "pause_avant_rdv_min": "", "marge_securite_min": 0,
            "distance_cumulee_km": round(cumulative_km + (km or 0), 1),
            "temps_route_cumule_min": int(cumulative_min + (ret_min if isinstance(ret_min, int) else 0)),
            "waze": waze_link(geo.get(start_address, {}).get("lat"), geo.get(start_address, {}).get("lon"), start_address),
            "google_maps": maps_link(start_address), "street_view": maps_link(start_address),
            "itineraire_depuis_precedent": directions_link(last_addr, start_address),
            "route_geometry": rb.get("geometry", []),
        }
    return route_df, return_row, geo.get(start_address, {})


def make_map(df, return_row, start_address, start_geo):
    map_df = df.copy()
    if return_row:
        map_df = pd.concat([map_df, pd.DataFrame([return_row])], ignore_index=True)
    valid = map_df.dropna(subset=["lat", "lon"])
    if not valid.empty:
        center = [valid["lat"].mean(), valid["lon"].mean()]
    elif start_geo.get("lat"):
        center = [start_geo["lat"], start_geo["lon"]]
    else:
        center = [48.79, 2.53]
    m = folium.Map(location=center, zoom_start=11, tiles="OpenStreetMap")
    points = []
    if start_geo.get("lat") and start_geo.get("lon"):
        folium.Marker([start_geo["lat"], start_geo["lon"]], tooltip="Départ / retour", popup=start_address, icon=folium.Icon(color="green", icon="home")).add_to(m)
        points.append([start_geo["lat"], start_geo["lon"]])
    for _, r in df.iterrows():
        if not r.get("lat") or not r.get("lon"):
            continue
        time_label = fmt_time(r.get("heure_rdv"))
        label = f"{r.get('numero_rdv','')} - {r.get('nom_prospect','')} - {time_label}"
        html = f"""
        <div style='font-size:20px;line-height:23px;font-weight:900;background:#ff8c00;color:#000;border:3px solid #fff;border-radius:10px;padding:7px 10px;white-space:nowrap;box-shadow:0 3px 12px rgba(0,0,0,.55);'>
        #{r.get('numero_rdv','')} · {r.get('nom_prospect','')}<br>🕒 {time_label}
        </div>"""
        folium.Marker([r["lat"], r["lon"]], tooltip=label, popup=folium.Popup(label, max_width=380), icon=folium.Icon(color="blue", icon="user")).add_to(m)
        folium.map.Marker([r["lat"], r["lon"]], icon=folium.DivIcon(html=html)).add_to(m)
        points.append([r["lat"], r["lon"]])
    if return_row and return_row.get("lat") and return_row.get("lon"):
        points.append([return_row["lat"], return_row["lon"]])
    # Tracé des routes réelles quand OSRM a fourni la géométrie, sinon ligne droite de secours
    route_drawn = False
    for _, r in df.iterrows():
        geom = r.get("route_geometry", [])
        if isinstance(geom, list) and len(geom) >= 2:
            folium.PolyLine(geom, weight=5, opacity=0.9, color="red").add_to(m)
            route_drawn = True
    if return_row:
        geom = return_row.get("route_geometry", [])
        if isinstance(geom, list) and len(geom) >= 2:
            folium.PolyLine(geom, weight=5, opacity=0.9, color="red", dash_array="8,6").add_to(m)
            route_drawn = True
    # Pas de ligne droite de secours : on évite l'effet "avion".
    # Si aucune géométrie routière n'est disponible, les marqueurs restent visibles sans faux tracé.
    return m


def streetview_static_image(lat, lon, api_key):
    if not api_key or not lat or not lon:
        return None
    try:
        params = {"size": "420x240", "location": f"{lat},{lon}", "fov": "90", "heading": "0", "pitch": "0", "key": api_key}
        r = requests.get("https://maps.googleapis.com/maps/api/streetview", params=params, timeout=8)
        if r.status_code == 200 and r.content:
            return io.BytesIO(r.content)
    except Exception:
        return None
    return None


def create_pdf(df, return_row, start_address, include_photos, google_key, visit_min=150):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=0.7*cm, leftMargin=0.7*cm, topMargin=0.7*cm, bottomMargin=0.7*cm)
    styles = getSampleStyleSheet()
    title = ParagraphStyle('TitleCustom', parent=styles['Title'], fontSize=18, leading=22, spaceAfter=8)
    h2 = ParagraphStyle('H2Custom', parent=styles['Heading2'], fontSize=13, leading=15, spaceBefore=8, spaceAfter=4)
    small = ParagraphStyle('Small', parent=styles['Normal'], fontSize=7.4, leading=9)
    normal = ParagraphStyle('NormalCustom', parent=styles['Normal'], fontSize=9, leading=11)
    story = []

    # Totaux robustes : évite les erreurs numpy quand le retour base est stocké en texte/objet.
    total_km = float(pd.to_numeric(df.get("distance_depuis_precedent_km", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
    total_km += to_float(return_row.get("distance_depuis_precedent_km", 0) if return_row else 0)
    total_min = int(pd.to_numeric(df.get("temps_route_depuis_precedent_min", pd.Series(dtype=float)), errors="coerce").fillna(0).sum())
    total_min += to_minutes(return_row.get("temps_route_depuis_precedent_min", 0) if return_row else 0)

    story.append(Paragraph("Tournée terrain — Routage PRO V14", title))
    story.append(Paragraph(f"Départ / retour : {start_address}", normal))
    story.append(Paragraph(f"RDV : {len(df)} · Distance totale retour inclus : {total_km:.1f} km · Temps route : {fmt_duration(total_min)}", normal))
    story.append(Spacer(1, 0.2*cm))

    story.append(Paragraph("Fil conducteur terrain", h2))
    tdata = [["Étape", "Heure", "Détail"]]
    for item in build_timeline(df, return_row, start_address, visit_min):
        tdata.append([Paragraph(item.get("Étape", ""), small), Paragraph(item.get("Heure conseillée", ""), small), Paragraph(item.get("Détail", ""), small)])
    tt = Table(tdata, colWidths=[2.2*cm, 2.0*cm, 14.0*cm], repeatRows=1)
    tt.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1f2937')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-1),0.25,colors.grey),('VALIGN',(0,0),(-1,-1),'TOP'),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#f3f4f6')])]))
    story.append(tt)
    story.append(Spacer(1, 0.25*cm))

    data = [["#", "RDV", "Client", "Adresse", "Trajet", "Départ conseillé", "Pause", "Liens"]]
    for _, r in df.iterrows():
        links = f"<a href='{r['waze']}'>Waze</a><br/><a href='{r['google_maps']}'>Maps</a><br/><a href='{r['street_view']}'>Maison</a>"
        if r.get('telephone_tel'):
            links += f"<br/><a href='tel:{r['telephone_tel']}'>Appeler</a>"
        pause = r.get("pause_avant_rdv_min", "")
        pause_txt = "" if pause == "" else (fmt_duration(pause) if to_minutes(pause) >= 0 else f"⚠ retard {fmt_duration(abs(to_minutes(pause)))}")
        data.append([
            str(r.get('numero_rdv','')),
            Paragraph(f"{fmt_date(r.get('date_rdv'))}<br/><b>{fmt_time(r.get('heure_rdv'))}</b>", small),
            Paragraph(f"<b>{r.get('nom_prospect','')}</b><br/>{r.get('telephone','')}<br/>Télépro : {r.get('teleprospecteur','')}", small),
            Paragraph(r.get('adresse_complete',''), small),
            Paragraph(f"{r.get('distance_depuis_precedent_km','')} km<br/>{fmt_duration(r.get('temps_route_depuis_precedent_min',''))}<br/>{r.get('note_trafic','')}", small),
            Paragraph(fmt_dt(r.get('depart_conseille')), small),
            Paragraph(pause_txt, small),
            Paragraph(links, small),
        ])
    if return_row:
        data.append(["BASE", "", Paragraph("<b>Retour base</b>", small), Paragraph(start_address, small), Paragraph(f"{return_row.get('distance_depuis_precedent_km','')} km<br/>{fmt_duration(return_row.get('temps_route_depuis_precedent_min',''))}", small), "", "", Paragraph(f"<a href='{return_row.get('waze','#')}'>Waze</a><br/><a href='{return_row.get('google_maps','#')}'>Maps</a>", small)])
    table = Table(data, colWidths=[1.0*cm, 1.7*cm, 2.8*cm, 5.0*cm, 2.0*cm, 2.0*cm, 1.4*cm, 2.0*cm], repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#111827')), ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('GRID', (0,0), (-1,-1), 0.25, colors.grey), ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('FONTSIZE', (0,0), (-1,-1), 7.2), ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f3f4f6')]),
    ]))
    story.append(table)

    story.append(PageBreak())
    story.append(Paragraph("Fiches prospects", title))
    for _, r in df.iterrows():
        story.append(Paragraph(f"#{r.get('numero_rdv','')} — {r.get('nom_prospect','')} — {fmt_time(r.get('heure_rdv'))}", h2))
        info = f"<b>Adresse :</b> {r.get('adresse_complete','')}<br/><b>Téléphone :</b> {r.get('telephone','')}<br/><b>Téléprospecteur :</b> {r.get('teleprospecteur','')}<br/><b>Départ conseillé :</b> {fmt_dt(r.get('depart_conseille'))}<br/><b>Trajet :</b> {r.get('distance_depuis_precedent_km','')} km · {fmt_duration(r.get('temps_route_depuis_precedent_min',''))}<br/><a href='{r.get('waze','#')}'>Ouvrir Waze</a> · <a href='{r.get('google_maps','#')}'>Google Maps</a> · <a href='{r.get('street_view','#')}'>Voir maison</a>"
        if r.get('telephone_tel'):
            info += f" · <a href='tel:{r.get('telephone_tel')}'>Appeler</a>"
        story.append(Paragraph(info, normal))
        if include_photos and google_key:
            img_bytes = streetview_static_image(r.get('lat'), r.get('lon'), google_key)
            if img_bytes:
                try:
                    story.append(Image(img_bytes, width=11*cm, height=6.3*cm))
                except Exception:
                    story.append(Paragraph("Image Street View indisponible — utiliser le lien Voir maison.", small))
            else:
                story.append(Paragraph("Image Street View indisponible — utiliser le lien Voir maison.", small))
        else:
            story.append(Paragraph("Photo maison : lien cliquable Voir maison disponible ci-dessus. Pour intégrer les photos directement, renseigner une clé Google Maps API.", small))
        story.append(Spacer(1, 0.25*cm))
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def to_recap_csv(df, return_row):
    export = renumber_route_df(df.copy())
    cols = ["ordre", "numero_rdv", "numero_rdv_source", "date_rdv", "heure_rdv", "depart_conseille", "pause_avant_rdv_min", "nom_prospect", "teleprospecteur", "telephone", "adresse_complete", "distance_depuis_precedent_km", "temps_route_depuis_precedent_min", "source_temps", "waze", "google_maps", "street_view", "lien_appel"]
    # Sécurité : si une ancienne tournée sauvegardée ne contient pas toutes les colonnes, on les recrée vides au lieu de planter.
    for c in cols:
        if c not in export.columns:
            export[c] = ""
    export["date_rdv"] = export["date_rdv"].apply(fmt_date)
    export["heure_rdv"] = export["heure_rdv"].apply(fmt_time)
    export["depart_conseille"] = export["depart_conseille"].apply(fmt_dt)
    if "telephone_tel" in export.columns:
        export["lien_appel"] = export["telephone_tel"].apply(lambda x: f"tel:{x}" if x else "")
    if return_row:
        row = {c: return_row.get(c, "") for c in cols}
        export = pd.concat([export[cols], pd.DataFrame([row])], ignore_index=True)
    return export[cols].to_csv(index=False, sep=";").encode("utf-8-sig")


def build_timeline(df, return_row, start_address, visit_min):
    lines = []
    if df.empty:
        return lines
    first = df.iloc[0]
    lines.append({
        "Étape": "Départ base",
        "Lieu": start_address,
        "Heure conseillée": fmt_dt(first.get("depart_conseille")),
        "Détail": f"Départ conseillé pour arriver chez {first.get('nom_prospect','')} à {fmt_time(first.get('heure_rdv'))} avec sécurité.",
    })
    rows = list(df.iterrows())
    for i, (_, r) in enumerate(rows):
        rdv_dt = r.get("rdv_datetime")
        fin_prevue = rdv_dt + timedelta(minutes=visit_min) if isinstance(rdv_dt, datetime) else None
        if i + 1 < len(rows):
            next_r = rows[i+1][1]
            depart_max = next_r.get("depart_conseille")
            pause_min = int((depart_max - fin_prevue).total_seconds() // 60) if isinstance(depart_max, datetime) and isinstance(fin_prevue, datetime) else None
            if pause_min is not None:
                pause_txt = f"Pause possible : {fmt_duration(pause_min)}" if pause_min >= 0 else f"⚠ retard probable : {fmt_duration(abs(pause_min))}"
            else:
                pause_txt = "Pause non calculée"
            detail = f"RDV prévu {fmt_time(r.get('heure_rdv'))} → fin estimée {fmt_dt(fin_prevue)}. Départ max vers le RDV suivant : {fmt_dt(depart_max)}. {pause_txt}."
        else:
            depart_retour = fin_prevue
            arrivee_retour = depart_retour + timedelta(minutes=int(return_row.get('temps_route_depuis_precedent_min', 0))) if return_row and isinstance(return_row.get('temps_route_depuis_precedent_min'), int) and isinstance(depart_retour, datetime) else None
            detail = f"RDV prévu {fmt_time(r.get('heure_rdv'))} → fin estimée {fmt_dt(fin_prevue)}. Retour base conseillé à {fmt_dt(depart_retour)}. Arrivée base estimée {fmt_dt(arrivee_retour)}."
        lines.append({
            "Étape": f"RDV {r.get('numero_rdv','')}",
            "Lieu": f"{r.get('nom_prospect','')} — {r.get('adresse_complete','')}",
            "Heure conseillée": fmt_time(r.get("heure_rdv")),
            "Détail": detail,
        })
    return lines



def to_float(value, default=0.0):
    try:
        if value is None or value == "":
            return default
        return float(value)
    except Exception:
        return default

def to_minutes(value, default=0):
    try:
        if value is None or value == "":
            return default
        return int(float(value))
    except Exception:
        return default

def save_last_uploaded(uploaded_file):
    try:
        data = uploaded_file.getvalue()
        LAST_UPLOAD_PATH.write_bytes(data)
        st.session_state["last_upload_bytes"] = data
        st.session_state["last_upload_name"] = uploaded_file.name
        return io.BytesIO(data)
    except Exception:
        return uploaded_file

def get_last_uploaded_file():
    if st.session_state.get("last_upload_bytes"):
        return io.BytesIO(st.session_state["last_upload_bytes"])
    if LAST_UPLOAD_PATH.exists():
        try:
            data = LAST_UPLOAD_PATH.read_bytes()
            st.session_state["last_upload_bytes"] = data
            st.session_state["last_upload_name"] = "dernier_fichier.xlsx"
            return io.BytesIO(data)
        except Exception:
            return None
    return None


# ==============================
# V17 — Module indemnités kilométriques
# ==============================
# Barème kilométrique voitures 2026 pour revenus 2025 — voitures thermiques, hybrides, hydrogène.
# Pour véhicule 100% électrique, une majoration de 20% est appliquée.
IK_BAREME_2026 = {
    3: ((5000, 0.529, 0), (20000, 0.316, 1065), (float("inf"), 0.370, 0)),
    4: ((5000, 0.606, 0), (20000, 0.340, 1330), (float("inf"), 0.407, 0)),
    5: ((5000, 0.636, 0), (20000, 0.357, 1395), (float("inf"), 0.427, 0)),
    6: ((5000, 0.665, 0), (20000, 0.374, 1457), (float("inf"), 0.447, 0)),
    7: ((5000, 0.697, 0), (20000, 0.394, 1515), (float("inf"), 0.470, 0)),
}


def ik_cv_key(cv):
    try:
        cv = int(cv)
    except Exception:
        cv = 7
    if cv <= 3:
        return 3
    if cv >= 7:
        return 7
    return cv


def calc_ik_amount(km, cv=7, electric=False):
    """Calcule l'indemnité kilométrique selon le barème 2026 voiture.
    Attention : le barème est annuel. Pour un export mensuel, l'app calcule sur les km de la période sélectionnée.
    Une régularisation annuelle peut être faite par le comptable si nécessaire.
    """
    d = max(0.0, float(km or 0))
    key = ik_cv_key(cv)
    brackets = IK_BAREME_2026[key]
    if d <= brackets[0][0]:
        amount = d * brackets[0][1] + brackets[0][2]
        formula = f"{d:.1f} km × {brackets[0][1]:.3f}"
    elif d <= brackets[1][0]:
        amount = d * brackets[1][1] + brackets[1][2]
        formula = f"({d:.1f} km × {brackets[1][1]:.3f}) + {brackets[1][2]:.0f}"
    else:
        amount = d * brackets[2][1] + brackets[2][2]
        formula = f"{d:.1f} km × {brackets[2][1]:.3f}"
    if electric:
        amount *= 1.20
        formula = f"({formula}) × 1,20 véhicule électrique"
    return round(amount, 2), formula


def euro(v):
    try:
        return f"{float(v):,.2f} €".replace(',', ' ').replace('.', ',')
    except Exception:
        return "0,00 €"


def build_ik_register(route_df, return_row=None, include_return=True, start_address=DEFAULT_START):
    rows = []
    if route_df is None or route_df.empty:
        return pd.DataFrame()
    for _, r in route_df.iterrows():
        d = r.get("date_rdv")
        date_txt = d.strftime("%d/%m/%Y") if isinstance(d, date) else str(d or "")
        rows.append({
            "Date": date_txt,
            "Objet": f"RDV client — {r.get('nom_prospect','')}",
            "Départ": "Base" if int(r.get('ordre', 1) or 1) == 1 else "RDV précédent",
            "Arrivée": r.get("adresse_complete", ""),
            "Client": r.get("nom_prospect", ""),
            "Téléprospecteur": r.get("teleprospecteur", ""),
            "Km": round(to_float(r.get("distance_depuis_precedent_km")), 1),
            "Temps": fmt_duration(r.get("temps_route_depuis_precedent_min", 0)),
            "Justificatif": f"RDV {r.get('numero_rdv','')} à {fmt_time(r.get('heure_rdv'))}",
        })
    if include_return and return_row:
        rows.append({
            "Date": "",
            "Objet": "Retour base",
            "Départ": "Dernier RDV",
            "Arrivée": start_address,
            "Client": "Retour",
            "Téléprospecteur": "",
            "Km": round(to_float(return_row.get("distance_depuis_precedent_km")), 1),
            "Temps": fmt_duration(return_row.get("temps_route_depuis_precedent_min", 0)),
            "Justificatif": "Retour base après tournée",
        })
    return pd.DataFrame(rows)


def create_ik_pdf(register_df, params):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=1.2*cm, leftMargin=1.2*cm, topMargin=1.0*cm, bottomMargin=1.0*cm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('IKTitle', parent=styles['Title'], fontSize=18, leading=22, textColor=colors.HexColor('#111827'))
    h_style = ParagraphStyle('IKH', parent=styles['Heading2'], fontSize=12, leading=15, textColor=colors.HexColor('#111827'))
    small = ParagraphStyle('IKSmall', parent=styles['Normal'], fontSize=8, leading=10)
    normal = ParagraphStyle('IKNormal', parent=styles['Normal'], fontSize=9, leading=12)
    story = []
    story.append(Paragraph("Note de frais — Indemnités kilométriques", title_style))
    story.append(Paragraph("Document généré depuis l’application Routage PRO V17", small))
    story.append(Spacer(1, 0.25*cm))
    info = [
        ["Bénéficiaire", params.get('beneficiaire','Mr Dahan'), "Société", params.get('societe','')],
        ["Période", params.get('periode',''), "Véhicule", params.get('vehicule','')],
        ["Immatriculation", params.get('immat',''), "Puissance fiscale", f"{params.get('cv','')} CV"],
        ["Base barème", params.get('bareme','Barème kilométrique 2026'), "Électrique", "Oui" if params.get('electric') else "Non"],
    ]
    t = Table(info, colWidths=[3.2*cm, 5.0*cm, 3.2*cm, 5.0*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,-1),colors.HexColor('#F3F4F6')),
        ('BOX',(0,0),(-1,-1),0.5,colors.HexColor('#CBD5E1')),
        ('INNERGRID',(0,0),(-1,-1),0.25,colors.HexColor('#CBD5E1')),
        ('FONTNAME',(0,0),(-1,-1),'Helvetica'),
        ('FONTSIZE',(0,0),(-1,-1),8),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.35*cm))
    total_km = float(register_df.get('Km', pd.Series(dtype=float)).sum()) if not register_df.empty else 0.0
    amount = params.get('amount', 0.0)
    summary = [
        ["Total kilomètres professionnels", f"{total_km:.1f} km"],
        ["Formule appliquée", params.get('formula','')],
        ["Montant à rembourser", euro(amount)],
    ]
    stbl = Table(summary, colWidths=[7.5*cm, 8.9*cm])
    stbl.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,1),colors.HexColor('#E0F2FE')),
        ('BACKGROUND',(0,2),(-1,2),colors.HexColor('#DCFCE7')),
        ('BOX',(0,0),(-1,-1),0.75,colors.HexColor('#0F172A')),
        ('INNERGRID',(0,0),(-1,-1),0.25,colors.HexColor('#94A3B8')),
        ('FONTNAME',(0,0),(-1,-1),'Helvetica-Bold'),
        ('FONTSIZE',(0,0),(-1,-1),9),
    ]))
    story.append(stbl)
    story.append(Spacer(1, 0.35*cm))
    story.append(Paragraph("Détail des déplacements", h_style))
    table_data = [["Date", "Objet", "Arrivée", "Km", "Temps", "Justificatif"]]
    for _, r in register_df.iterrows():
        table_data.append([
            str(r.get('Date','')),
            Paragraph(str(r.get('Objet','')), small),
            Paragraph(str(r.get('Arrivée','')), small),
            f"{to_float(r.get('Km')):.1f}",
            str(r.get('Temps','')),
            Paragraph(str(r.get('Justificatif','')), small),
        ])
    tbl = Table(table_data, colWidths=[1.7*cm, 3.5*cm, 6.2*cm, 1.3*cm, 1.7*cm, 2.8*cm], repeatRows=1)
    tbl.setStyle(TableStyle([
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#111827')),
        ('TEXTCOLOR',(0,0),(-1,0),colors.white),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('FONTSIZE',(0,0),(-1,-1),7),
        ('GRID',(0,0),(-1,-1),0.25,colors.HexColor('#CBD5E1')),
        ('VALIGN',(0,0),(-1,-1),'TOP'),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white, colors.HexColor('#F8FAFC')]),
    ]))
    story.append(tbl)
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph("Déclaration", h_style))
    story.append(Paragraph("Je certifie que les déplacements ci-dessus ont été effectués dans le cadre professionnel et que les kilomètres déclarés correspondent aux trajets calculés par l’application à partir des adresses de rendez-vous.", normal))
    story.append(Spacer(1, 0.5*cm))
    sig = Table([["Date :", "", "Signature bénéficiaire :", ""]], colWidths=[2*cm,4*cm,4*cm,5*cm])
    sig.setStyle(TableStyle([('LINEBELOW',(1,0),(1,0),0.5,colors.black),('LINEBELOW',(3,0),(3,0),0.5,colors.black),('FONTSIZE',(0,0),(-1,-1),9)]))
    story.append(sig)
    story.append(Spacer(1, 0.3*cm))
    story.append(Paragraph("Note : le barème kilométrique couvre notamment la dépréciation, l’entretien, les pneumatiques, le carburant et l’assurance. Les péages et stationnements peuvent être ajoutés séparément sur justificatifs.", small))
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def df_to_csv_bytes(df):
    return df.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')

with st.sidebar:
    st.header("Réglages")
    start_address = st.text_input("Adresse de départ / retour", value=DEFAULT_START)
    safety_min = st.number_input("Marge sécurité avant RDV", min_value=0, max_value=60, value=15, step=5)
    visit_min = st.number_input("Durée moyenne d'un RDV", min_value=15, max_value=240, value=150, step=15)
    use_google = st.checkbox("Utiliser Google trafic / Street View si j'ai une clé API", value=False)
    google_key = st.text_input("Clé Google Maps API (optionnel)", type="password") if use_google else ""
    uploaded = st.file_uploader("Importer ton fichier Excel", type=["xlsx", "xls"])
    saved = st.file_uploader("Ou charger un récap CSV sauvegardé", type=["csv"], key="saved_csv")
    auto_reload = st.checkbox("Recharger automatiquement le dernier Excel de la journée", value=True)
    st.info("V13 : mode sombre + calcul automatique + géocodage France renforcé + tracé route réelle OSRM + dernier Excel auto.")

source_file = None
source_label = ""
if uploaded:
    source_file = save_last_uploaded(uploaded)
    source_label = uploaded.name
elif auto_reload:
    source_file = get_last_uploaded_file()
    source_label = st.session_state.get("last_upload_name", "dernier fichier")

if source_file:
    try:
        df = prepare_dataframe(source_file)
        if df.empty:
            st.error("Aucune adresse trouvée dans le fichier.")
            st.stop()
        st.success(f"{len(df)} RDV chargés depuis {source_label}. Calcul automatique lancé, aucun bouton à cliquer.")
        with st.spinner("Géocodage, trajets, pauses, départs conseillés..."):
            route_df, return_row, start_geo = enrich_route(df, start_address, int(safety_min), int(visit_min), use_google, google_key)
        st.session_state["route_df"] = route_df
        st.session_state["return_row"] = return_row
        st.session_state["start_address"] = start_address
        st.session_state["start_geo"] = start_geo
        st.session_state["google_key"] = google_key
        st.session_state["use_google"] = use_google
    except Exception as e:
        st.exception(e)
        st.stop()
elif saved:
    try:
        route_df = renumber_route_df(pd.read_csv(saved, sep=";"))
        st.session_state["route_df"] = route_df
        st.session_state["return_row"] = None
        st.session_state["start_address"] = start_address
        st.session_state["start_geo"] = {}
        st.session_state["google_key"] = ""
        st.session_state["use_google"] = False
        st.success("Récap chargé. Les liens restent utilisables.")
    except Exception as e:
        st.exception(e)

if "route_df" not in st.session_state:
    st.warning("Importe ton Excel dans la barre de gauche pour générer ta tournée.")
    st.markdown("""
### Format attendu
A numéro RDV · B adresse · C code postal · D date RDV · E heure RDV · J/N nom/prénom · Q téléphone · R ville.
""")
    st.stop()

route_df = renumber_route_df(st.session_state["route_df"])
st.session_state["route_df"] = route_df
return_row = st.session_state.get("return_row")
start_address = st.session_state.get("start_address", DEFAULT_START)
start_geo = st.session_state.get("start_geo", {})
google_key = st.session_state.get("google_key", "")
use_google = st.session_state.get("use_google", False)

distance_series = pd.to_numeric(route_df.get("distance_depuis_precedent_km"), errors="coerce").fillna(0)
time_series = pd.to_numeric(route_df.get("temps_route_depuis_precedent_min"), errors="coerce").fillna(0)
total_km = float(distance_series.sum()) + (to_float(return_row.get("distance_depuis_precedent_km", 0)) if return_row else 0.0)
total_min = int(time_series.sum()) + (to_minutes(return_row.get("temps_route_depuis_precedent_min", 0)) if return_row else 0)

col1, col2, col3, col4 = st.columns(4)
col1.metric("RDV", len(route_df))
col2.metric("Distance retour inclus", f"{total_km:.1f} km")
col3.metric("Temps route", fmt_duration(total_min))
if not route_df.empty:
    first_dep = route_df.iloc[0].get("depart_conseille")
    col4.metric("Premier départ conseillé", fmt_dt(first_dep))
    if fmt_dt(first_dep):
        st.success(f"Départ conseillé de la base : {fmt_dt(first_dep)}")
    else:
        st.warning("Départ conseillé non calculé : vérifie que chaque RDV a bien une date et une heure dans les colonnes D et E.")

st.subheader("🧭 Fil conducteur terrain")
timeline_df = pd.DataFrame(build_timeline(route_df, return_row, start_address, int(visit_min)))
st.dataframe(timeline_df, use_container_width=True, hide_index=True)

st.subheader("📊 Détail des trajets étape par étape")
show_cols = ["numero_rdv", "heure_rdv", "depart_conseille", "pause_avant_rdv_min", "nom_prospect", "teleprospecteur", "telephone", "adresse_complete", "distance_depuis_precedent_km", "temps_route_depuis_precedent_min", "note_trafic"]
display_df = route_df[show_cols].copy()
display_df["heure_rdv"] = display_df["heure_rdv"].apply(fmt_time)
display_df["depart_conseille"] = display_df["depart_conseille"].apply(fmt_dt)
display_df["pause_avant_rdv_min"] = display_df["pause_avant_rdv_min"].apply(lambda x: "" if x == "" else fmt_duration(x))
display_df["temps_route_depuis_precedent_min"] = display_df["temps_route_depuis_precedent_min"].apply(fmt_duration)
display_df = display_df.rename(columns={
    "numero_rdv": "N° RDV", "heure_rdv": "Heure RDV", "depart_conseille": "Départ conseillé",
    "pause_avant_rdv_min": "Pause avant RDV", "nom_prospect": "Client", "teleprospecteur": "Téléprospecteur", "telephone": "Téléphone",
    "adresse_complete": "Adresse", "distance_depuis_precedent_km": "Km depuis précédent",
    "temps_route_depuis_precedent_min": "Temps depuis précédent", "note_trafic": "Calcul"
})
st.dataframe(display_df, use_container_width=True, hide_index=True)
if return_row:
    st.info(f"Retour base inclus : {return_row.get('distance_depuis_precedent_km','')} km · {fmt_duration(return_row.get('temps_route_depuis_precedent_min',''))}")

st.subheader("📋 Mode terrain")
for _, r in route_df.iterrows():
    pause = r.get('pause_avant_rdv_min', '')
    pause_txt = "" if pause == "" else (f" · Pause dispo : {fmt_duration(pause)}" if to_minutes(pause) >= 0 else f" · ⚠ Retard probable : {fmt_duration(abs(to_minutes(pause)))}")
    title = f"RDV {r.get('numero_rdv','')} · {fmt_time(r.get('heure_rdv'))} · {r.get('nom_prospect','')}{pause_txt}"
    with st.expander(title, expanded=(str(r.get('ordre','')) == '1')):
        c1, c2 = st.columns([2, 1])
        with c1:
            st.markdown(f"**Adresse :** {r.get('adresse_complete','')}")
            st.markdown(f"**Téléphone :** {r.get('telephone','')}")
            st.markdown(f"**Départ conseillé :** {fmt_dt(r.get('depart_conseille'))} avec {r.get('marge_securite_min', safety_min)} min de sécurité")
            st.markdown(f"**Trajet depuis précédent :** {r.get('distance_depuis_precedent_km','')} km · {fmt_duration(r.get('temps_route_depuis_precedent_min',''))} · {r.get('note_trafic','')}")
        with c2:
            st.link_button("🚗 Waze", r.get('waze', '#'), use_container_width=True)
            st.link_button("🗺️ Google Maps", r.get('google_maps', '#'), use_container_width=True)
            st.link_button("🏠 Voir maison", r.get('street_view', '#'), use_container_width=True)
            if r.get('telephone_tel'):
                st.link_button("📞 Appeler", f"tel:{r.get('telephone_tel')}", use_container_width=True)

st.subheader("🗺️ Carte générale")
nb_routes = sum(1 for _, rr in route_df.iterrows() if isinstance(rr.get("route_geometry", []), list) and len(rr.get("route_geometry", [])) >= 2)
if return_row and isinstance(return_row.get("route_geometry", []), list) and len(return_row.get("route_geometry", [])) >= 2:
    nb_routes += 1
if nb_routes == 0:
    st.warning("Aucun tracé routier disponible pour l’instant. Vérifie la connexion ou les adresses. Les calculs peuvent quand même apparaître si les coordonnées sont trouvées.")
else:
    st.success(f"{nb_routes} trajet(s) routier(s) tracé(s) sur la carte.")
try:
    st_folium(make_map(route_df, return_row, start_address, start_geo), height=650, use_container_width=True)
except Exception as e:
    st.warning(f"Carte non disponible : {e}")

st.subheader("📤 Exports terrain")
include_photos = st.checkbox("Essayer d'intégrer les photos Street View dans le PDF", value=bool(google_key), help="Nécessite une clé Google Maps API. Sinon le PDF contient le lien Voir maison cliquable.")
pdf_bytes = create_pdf(route_df, return_row, start_address, include_photos, google_key, int(visit_min))
csv_bytes = to_recap_csv(route_df, return_row)

c1, c2 = st.columns(2)
with c1:
    st.download_button("📄 Télécharger PDF enrichi cliquable", data=pdf_bytes, file_name="tournee_terrain_v17.pdf", mime="application/pdf", use_container_width=True)
with c2:
    st.download_button("💾 Sauvegarde CSV réutilisable", data=csv_bytes, file_name="tournee_sauvegarde_v17.csv", mime="text/csv", use_container_width=True)


st.subheader("💰 Indemnités kilométriques — PDF professionnel")
st.caption("Module V17 séparé : il ne modifie pas la tournée. Tu peux générer une note de frais mensuelle/propre à partir des trajets calculés.")
with st.expander("Paramétrer la note de frais IK", expanded=False):
    a,b,c = st.columns(3)
    with a:
        beneficiaire = st.text_input("Bénéficiaire", value="Mr Dahan")
        societe = st.text_input("Société à facturer / rembourser", value="")
        periode = st.text_input("Période", value=datetime.now().strftime("%B %Y"))
    with b:
        vehicule = st.text_input("Véhicule", value="")
        immat = st.text_input("Immatriculation", value="")
        cv = st.selectbox("Puissance fiscale", options=[3,4,5,6,7], index=4, help="7 = 7 CV et plus")
    with c:
        electric = st.checkbox("Véhicule 100% électrique (+20%)", value=False)
        include_return_ik = st.checkbox("Inclure le retour à la base", value=True)
        st.info("Barème voiture 2026, revenus 2025. À valider avec ton comptable selon ton montage exact.")

    ik_register = build_ik_register(route_df, return_row, include_return_ik, start_address)
    total_ik_km = float(ik_register["Km"].sum()) if not ik_register.empty else 0.0
    ik_amount, ik_formula = calc_ik_amount(total_ik_km, cv=cv, electric=electric)
    k1,k2,k3 = st.columns(3)
    k1.metric("Km IK", f"{total_ik_km:.1f} km")
    k2.metric("Montant IK", euro(ik_amount))
    k3.metric("Barème", f"{cv} CV" + (" électrique" if electric else ""))
    st.dataframe(ik_register, use_container_width=True, hide_index=True)

    ik_params = {
        "beneficiaire": beneficiaire,
        "societe": societe,
        "periode": periode,
        "vehicule": vehicule,
        "immat": immat,
        "cv": cv,
        "electric": electric,
        "bareme": "Barème kilométrique 2026 — voitures",
        "amount": ik_amount,
        "formula": ik_formula,
    }
    ik_pdf = create_ik_pdf(ik_register, ik_params)
    d1,d2 = st.columns(2)
    with d1:
        st.download_button("📄 Télécharger la note IK PDF", data=ik_pdf, file_name="note_indemnites_kilometriques.pdf", mime="application/pdf", use_container_width=True)
    with d2:
        st.download_button("📊 Télécharger le registre IK CSV", data=df_to_csv_bytes(ik_register), file_name="registre_indemnites_kilometriques.csv", mime="text/csv", use_container_width=True)


st.caption("V17 : V16 stable + module IK PDF professionnel. Sans clé Google, le trafic est une estimation prudente. Les tracés routiers utilisent OSRM gratuit quand les coordonnées sont trouvées.")
