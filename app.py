import io
import re
import time
import math
import json
from datetime import datetime, date, time as dtime, timedelta
from urllib.parse import quote_plus

import pandas as pd
import streamlit as st
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from geopy.distance import geodesic
import folium
from streamlit_folium import st_folium
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase.pdfmetrics import stringWidth

st.set_page_config(page_title="Routage PRO V8", page_icon="🚗", layout="wide")

DEFAULT_START = "72 avenue des Tourelles, 94490 Ormesson-sur-Marne"
AVG_SPEED_KMH = 42  # estimation route IDF hors trafic réel

COLS = {
    "numero_rdv": 0,
    "adresse": 1,
    "code_postal": 2,
    "date_rdv": 3,
    "heure_debut": 4,
    "email": 5,
    "fournisseur": 7,
    "commercial_nom": 8,
    "nom": 9,
    "telepros_nom": 11,
    "commercial_prenom": 12,
    "prenom": 13,
    "telephone": 16,
    "ville": 17,
}

st.title("🚗 Routage PRO V8 — terrain")
st.caption("Ordre imposé par l'heure de RDV · Waze · Google Maps · Street View · PDF cliquable")


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
    # Excel decimal time
    if isinstance(v, (int, float)) and 0 <= v < 1:
        total_minutes = int(round(v * 24 * 60))
        return dtime(total_minutes // 60, total_minutes % 60)
    s = str(v).strip().replace("h", ":").replace("H", ":")
    if re.match(r"^\d{1,2}:\d{2}(:\d{2})?$", s):
        parts = [int(x) for x in s.split(":")[:2]]
        return dtime(parts[0], parts[1])
    if re.match(r"^\d{1,2}$", s):
        return dtime(int(s), 0)
    try:
        return pd.to_datetime(s).time().replace(second=0, microsecond=0)
    except Exception:
        return None


def format_phone(raw):
    digits = re.sub(r"\D", "", str(raw or ""))
    if digits.startswith("33") and len(digits) == 11:
        digits = "0" + digits[2:]
    if len(digits) == 9 and digits[0] in "123456789":
        digits = "0" + digits
    if len(digits) == 10:
        formatted = " ".join([digits[0:2], digits[2:4], digits[4:6], digits[6:8], digits[8:10]])
        return formatted, digits
    return str(raw or ""), digits


def full_name(prenom, nom):
    return " ".join([x for x in [str(prenom).strip(), str(nom).strip()] if x and x.lower() != "nan"]).strip() or "Prospect"


def build_address(adresse, cp, ville):
    return ", ".join([x for x in [adresse, cp, ville] if str(x).strip()])


def waze_link(lat, lon, address):
    if lat and lon:
        return f"https://waze.com/ul?ll={lat},{lon}&navigate=yes"
    return f"https://waze.com/ul?q={quote_plus(address)}&navigate=yes"


def maps_link(address):
    return f"https://www.google.com/maps/search/?api=1&query={quote_plus(address)}"


def streetview_link(address):
    return f"https://www.google.com/maps/@?api=1&map_action=pano&viewpoint={quote_plus(address)}"


def directions_link(origin, destination):
    return f"https://www.google.com/maps/dir/?api=1&origin={quote_plus(origin)}&destination={quote_plus(destination)}&travelmode=driving"

@st.cache_data(show_spinner=False)
def geocode_addresses(addresses):
    geolocator = Nominatim(user_agent="routage_pro_v8_froid24")
    geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1, max_retries=1, swallow_exceptions=True)
    out = {}
    for address in addresses:
        loc = geocode(address + ", France")
        if loc:
            out[address] = {"lat": loc.latitude, "lon": loc.longitude}
        else:
            out[address] = {"lat": None, "lon": None}
    return out


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
        item = {
            "numero_rdv": safe_get(row, COLS["numero_rdv"]),
            "nom_prospect": full_name(safe_get(row, COLS["prenom"]), safe_get(row, COLS["nom"])),
            "adresse": adresse,
            "code_postal": cp,
            "ville": ville,
            "adresse_complete": build_address(adresse, cp, ville),
            "date_rdv": d,
            "heure_rdv": h,
            "telephone": phone_fmt,
            "telephone_tel": phone_digits,
            "email": safe_get(row, COLS["email"]),
            "fournisseur": safe_get(row, COLS["fournisseur"]),
            "commercial": full_name(safe_get(row, COLS["commercial_prenom"]), safe_get(row, COLS["commercial_nom"])),
            "teleprospecteur": full_name(safe_get(row, COLS["telepros_nom"]), safe_get(row, COLS["telepros_nom"])),
        }
        rows.append(item)
    result = pd.DataFrame(rows)
    if not result.empty:
        result["sort_date"] = result["date_rdv"].apply(lambda x: x or date.max)
        result["sort_time"] = result["heure_rdv"].apply(lambda x: x or dtime.max)
        result = result.sort_values(["sort_date", "sort_time", "numero_rdv"]).drop(columns=["sort_date", "sort_time"]).reset_index(drop=True)
        result.insert(0, "ordre", range(1, len(result) + 1))
    return result


def enrich_route(df, start_address):
    addresses = [start_address] + df["adresse_complete"].tolist()
    geo = geocode_addresses(addresses)
    prev_addr = start_address
    prev_coord = geo.get(start_address, {})
    out = []
    cumulative_km = 0
    cumulative_min = 0
    for _, row in df.iterrows():
        addr = row["adresse_complete"]
        coord = geo.get(addr, {})
        lat, lon = coord.get("lat"), coord.get("lon")
        dist = None
        mins = None
        if prev_coord.get("lat") and prev_coord.get("lon") and lat and lon:
            dist = geodesic((prev_coord["lat"], prev_coord["lon"]), (lat, lon)).km * 1.25  # facteur route estimatif
            mins = int(round((dist / AVG_SPEED_KMH) * 60))
        cumulative_km += dist or 0
        cumulative_min += mins or 0
        r = row.to_dict()
        r.update({
            "lat": lat,
            "lon": lon,
            "distance_depuis_precedent_km": round(dist, 1) if dist is not None else "",
            "temps_route_depuis_precedent_min": mins if mins is not None else "",
            "distance_cumulee_km": round(cumulative_km, 1),
            "temps_route_cumule_min": cumulative_min,
            "waze": waze_link(lat, lon, addr),
            "google_maps": maps_link(addr),
            "street_view": streetview_link(addr),
            "itineraire_depuis_precedent": directions_link(prev_addr, addr),
        })
        out.append(r)
        prev_addr = addr
        prev_coord = coord
    return pd.DataFrame(out), geo.get(start_address, {})


def make_map(df, start_address, start_geo):
    valid = df.dropna(subset=["lat", "lon"])
    if not valid.empty:
        center = [valid["lat"].mean(), valid["lon"].mean()]
    elif start_geo.get("lat"):
        center = [start_geo["lat"], start_geo["lon"]]
    else:
        center = [48.79, 2.53]
    m = folium.Map(location=center, zoom_start=11)
    if start_geo.get("lat") and start_geo.get("lon"):
        folium.Marker([start_geo["lat"], start_geo["lon"]], tooltip="Départ", popup=start_address, icon=folium.Icon(color="green", icon="home")).add_to(m)
    points = []
    for _, r in df.iterrows():
        if not r.get("lat") or not r.get("lon"):
            continue
        label = f"{r['numero_rdv']} - {r['nom_prospect']} - {r['heure_rdv'].strftime('%H:%M') if isinstance(r['heure_rdv'], dtime) else ''} - {r['telephone']}"
        html = f"""
        <div style='font-size:12px;font-weight:bold;background:white;border:1px solid #333;border-radius:5px;padding:3px;white-space:nowrap;'>
        {r['numero_rdv']} - {r['nom_prospect']}<br>🕒 {r['heure_rdv'].strftime('%H:%M') if isinstance(r['heure_rdv'], dtime) else ''} &nbsp; 📞 {r['telephone']}
        </div>"""
        folium.Marker([r["lat"], r["lon"]], tooltip=label, popup=folium.Popup(label, max_width=350), icon=folium.Icon(color="blue", icon="user")).add_to(m)
        folium.map.Marker([r["lat"], r["lon"]], icon=folium.DivIcon(html=html)).add_to(m)
        points.append([r["lat"], r["lon"]])
    if start_geo.get("lat") and start_geo.get("lon"):
        points = [[start_geo["lat"], start_geo["lon"]]] + points
    if len(points) >= 2:
        folium.PolyLine(points, weight=3, opacity=0.7).add_to(m)
    return m


def fmt_date(x):
    return x.strftime("%d/%m/%Y") if isinstance(x, date) else ""

def fmt_time(x):
    return x.strftime("%H:%M") if isinstance(x, dtime) else ""

def fmt_duration(m):
    if m == "" or pd.isna(m): return ""
    m = int(m)
    return f"{m//60}h{m%60:02d}" if m >= 60 else f"{m} min"


def create_pdf(df, start_address):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=0.8*cm, leftMargin=0.8*cm, topMargin=0.8*cm, bottomMargin=0.8*cm)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleCustom', parent=styles['Title'], fontSize=18, leading=22, spaceAfter=8)
    small = ParagraphStyle('Small', parent=styles['Normal'], fontSize=7.5, leading=9)
    normal = ParagraphStyle('NormalCustom', parent=styles['Normal'], fontSize=9, leading=11)
    story = []
    story.append(Paragraph("Tournée terrain — Routage PRO V8", title_style))
    story.append(Paragraph(f"Départ : {start_address}", normal))
    story.append(Paragraph("Ordre imposé par l'heure de RDV. Liens cliquables sur PC/iPhone.", small))
    story.append(Spacer(1, 0.25*cm))
    data = [["#", "RDV", "Client", "Adresse", "Trajet", "Actions"]]
    for _, r in df.iterrows():
        client = f"{r['nom_prospect']}<br/>{r['telephone']}"
        rdv = f"{fmt_date(r['date_rdv'])}<br/>{fmt_time(r['heure_rdv'])}"
        addr = r['adresse_complete']
        trajet = f"{r['distance_depuis_precedent_km']} km<br/>{fmt_duration(r['temps_route_depuis_precedent_min'])}"
        actions = f"<a href='{r['waze']}'>Waze</a><br/><a href='{r['google_maps']}'>Maps</a><br/><a href='{r['street_view']}'>Maison</a>"
        if r.get('telephone_tel'):
            actions += f"<br/><a href='tel:{r['telephone_tel']}'>Appeler</a>"
        data.append([str(r['ordre']), rdv, Paragraph(client, small), Paragraph(addr, small), Paragraph(trajet, small), Paragraph(actions, small)])
    table = Table(data, colWidths=[0.7*cm, 1.8*cm, 3.1*cm, 6.2*cm, 1.7*cm, 2.5*cm], repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1f2937')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('GRID', (0,0), (-1,-1), 0.25, colors.grey),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('FONTSIZE', (0,0), (-1,-1), 7.5),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#f3f4f6')]),
    ]))
    story.append(table)
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def to_recap_csv(df):
    export = df.copy()
    export["date_rdv"] = export["date_rdv"].apply(fmt_date)
    export["heure_rdv"] = export["heure_rdv"].apply(fmt_time)
    export["lien_appel"] = export["telephone_tel"].apply(lambda x: f"tel:{x}" if x else "")
    cols = ["ordre", "numero_rdv", "date_rdv", "heure_rdv", "nom_prospect", "telephone", "adresse_complete", "distance_depuis_precedent_km", "temps_route_depuis_precedent_min", "waze", "google_maps", "street_view", "lien_appel"]
    return export[cols].to_csv(index=False, sep=";").encode("utf-8-sig")

with st.sidebar:
    st.header("Réglages")
    start_address = st.text_input("Adresse de départ", value=DEFAULT_START)
    uploaded = st.file_uploader("Importer ton fichier Excel", type=["xlsx", "xls"])
    saved = st.file_uploader("Ou charger un récap CSV sauvegardé", type=["csv"], key="saved_csv")
    st.info("V8 : plus besoin d'assigner les colonnes. Le format est fixé selon ton fichier.")

if uploaded:
    try:
        df = prepare_dataframe(uploaded)
        if df.empty:
            st.error("Aucune adresse trouvée dans le fichier.")
            st.stop()
        st.success(f"{len(df)} RDV chargés. Ordre trié par date et heure de RDV.")
        with st.spinner("Géocodage des adresses et préparation des liens terrain..."):
            route_df, start_geo = enrich_route(df, start_address)
        st.session_state["route_df"] = route_df
        st.session_state["start_address"] = start_address
        st.session_state["start_geo"] = start_geo
    except Exception as e:
        st.exception(e)
        st.stop()
elif saved:
    try:
        route_df = pd.read_csv(saved, sep=";")
        st.session_state["route_df"] = route_df
        st.session_state["start_address"] = start_address
        st.session_state["start_geo"] = {}
        st.success("Récap chargé.")
    except Exception as e:
        st.exception(e)

if "route_df" not in st.session_state:
    st.warning("Importe ton Excel dans la barre de gauche pour générer ta tournée.")
    st.markdown("""
### Format attendu
A numéro RDV · B adresse · C code postal · D date RDV · E heure RDV · J/N nom/prénom · Q téléphone · R ville.
""")
    st.stop()

route_df = st.session_state["route_df"]
start_address = st.session_state.get("start_address", DEFAULT_START)
start_geo = st.session_state.get("start_geo", {})

col1, col2, col3 = st.columns(3)
col1.metric("RDV", len(route_df))
if "distance_depuis_precedent_km" in route_df:
    col2.metric("Distance estimée", f"{pd.to_numeric(route_df['distance_depuis_precedent_km'], errors='coerce').fillna(0).sum():.1f} km")
    col3.metric("Temps route estimé", fmt_duration(pd.to_numeric(route_df['temps_route_depuis_precedent_min'], errors='coerce').fillna(0).sum()))

st.subheader("📋 Tournée terrain")
for _, r in route_df.iterrows():
    title = f"{r.get('ordre','')} — RDV {r.get('numero_rdv','')} · {fmt_time(r.get('heure_rdv')) if not isinstance(r.get('heure_rdv'), str) else r.get('heure_rdv','')} · {r.get('nom_prospect','')}"
    with st.expander(title, expanded=(int(r.get('ordre', 99)) == 1 if str(r.get('ordre','')).isdigit() else False)):
        c1, c2 = st.columns([2, 1])
        with c1:
            st.markdown(f"**Adresse :** {r.get('adresse_complete','')}")
            st.markdown(f"**Téléphone :** {r.get('telephone','')}")
            st.markdown(f"**Trajet depuis précédent :** {r.get('distance_depuis_precedent_km','')} km · {fmt_duration(r.get('temps_route_depuis_precedent_min',''))}")
        with c2:
            st.link_button("🚗 Waze", r.get('waze', '#'))
            st.link_button("🗺️ Google Maps", r.get('google_maps', '#'))
            st.link_button("🏠 Voir maison", r.get('street_view', '#'))
            if r.get('telephone_tel'):
                st.link_button("📞 Appeler", f"tel:{r.get('telephone_tel')}")

st.subheader("🗺️ Carte générale")
try:
    st_folium(make_map(route_df, start_address, start_geo), height=620, use_container_width=True)
except Exception as e:
    st.warning(f"Carte non disponible : {e}")

st.subheader("📤 Exports terrain")
pdf_bytes = create_pdf(route_df, start_address)
csv_bytes = to_recap_csv(route_df)

c1, c2 = st.columns(2)
with c1:
    st.download_button("📄 Télécharger PDF cliquable", data=pdf_bytes, file_name="tournee_terrain_v8.pdf", mime="application/pdf")
with c2:
    st.download_button("💾 Sauvegarde CSV réutilisable", data=csv_bytes, file_name="tournee_sauvegarde_v8.csv", mime="text/csv")

st.caption("Astuce iPhone : le PDF conserve les liens Waze / Maps / Appeler. Si l'upload iPhone bloque, utilise Chrome ou prépare la tournée sur Surface puis ouvre le PDF sur iPhone.")
