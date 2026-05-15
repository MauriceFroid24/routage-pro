import io
import re
import math
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

st.set_page_config(page_title="Routage PRO V10", page_icon="🚗", layout="wide")

DEFAULT_START = "72 avenue des Tourelles, 94490 Ormesson-sur-Marne"
AVG_SPEED_KMH = 38

COLS = {
    "numero_rdv": 0, "adresse": 1, "code_postal": 2, "date_rdv": 3, "heure_debut": 4,
    "email": 5, "fournisseur": 7, "commercial_nom": 8, "nom": 9, "telepros_nom": 11,
    "commercial_prenom": 12, "prenom": 13, "telephone": 16, "ville": 17,
}

st.title("🚗 Routage PRO V10 — terrain iPhone / Surface")
st.caption("Ordre par heure de RDV · retour base · pauses · départ conseillé · PDF enrichi · Waze / Maps / appel")


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
    if lat and lon:
        return f"https://waze.com/ul?ll={lat},{lon}&navigate=yes"
    return f"https://waze.com/ul?q={quote_plus(address)}&navigate=yes"


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
def geocode_addresses(addresses):
    geolocator = Nominatim(user_agent="routage_pro_v10_froid24")
    geocode = RateLimiter(geolocator.geocode, min_delay_seconds=1, max_retries=1, swallow_exceptions=True)
    out = {}
    for address in addresses:
        loc = geocode(address + ", France")
        out[address] = {"lat": loc.latitude, "lon": loc.longitude} if loc else {"lat": None, "lon": None}
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
            "teleprospecteur": safe_get(row, COLS["telepros_nom"]),
        })
    result = pd.DataFrame(rows)
    if not result.empty:
        result["sort_date"] = result["date_rdv"].apply(lambda x: x or date.max)
        result["sort_time"] = result["heure_rdv"].apply(lambda x: x or dtime.max)
        result = result.sort_values(["sort_date", "sort_time", "numero_rdv"]).drop(columns=["sort_date", "sort_time"]).reset_index(drop=True)
        result.insert(0, "ordre", range(1, len(result) + 1))
    return result


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
            "lat": coord.get("lat"), "lon": coord.get("lon"),
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

    route_df = pd.DataFrame(out)

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
        <div style='font-size:20px;line-height:23px;font-weight:900;background:#ffea00;color:#111;border:3px solid #000;border-radius:10px;padding:7px 10px;white-space:nowrap;box-shadow:0 3px 12px rgba(0,0,0,.55);'>
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
    if not route_drawn and len(points) >= 2:
        folium.PolyLine(points, weight=4, opacity=0.8, color="red").add_to(m)
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

    total_km = pd.to_numeric(df["distance_depuis_precedent_km"], errors="coerce").fillna(0).sum() + (return_row.get("distance_depuis_precedent_km", 0) if return_row else 0)
    total_min = pd.to_numeric(df["temps_route_depuis_precedent_min"], errors="coerce").fillna(0).sum() + (return_row.get("temps_route_depuis_precedent_min", 0) if return_row and isinstance(return_row.get("temps_route_depuis_precedent_min"), int) else 0)

    story.append(Paragraph("Tournée terrain — Routage PRO V10", title))
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
        pause_txt = "" if pause == "" else (f"{pause} min" if int(pause) >= 0 else f"⚠ retard {abs(int(pause))} min")
        data.append([
            str(r.get('numero_rdv','')),
            Paragraph(f"{fmt_date(r.get('date_rdv'))}<br/><b>{fmt_time(r.get('heure_rdv'))}</b>", small),
            Paragraph(f"<b>{r.get('nom_prospect','')}</b><br/>{r.get('telephone','')}", small),
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
        info = f"<b>Adresse :</b> {r.get('adresse_complete','')}<br/><b>Téléphone :</b> {r.get('telephone','')}<br/><b>Départ conseillé :</b> {fmt_dt(r.get('depart_conseille'))}<br/><b>Trajet :</b> {r.get('distance_depuis_precedent_km','')} km · {fmt_duration(r.get('temps_route_depuis_precedent_min',''))}<br/><a href='{r.get('waze','#')}'>Ouvrir Waze</a> · <a href='{r.get('google_maps','#')}'>Google Maps</a> · <a href='{r.get('street_view','#')}'>Voir maison</a>"
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
    export = df.copy()
    export["date_rdv"] = export["date_rdv"].apply(fmt_date)
    export["heure_rdv"] = export["heure_rdv"].apply(fmt_time)
    export["depart_conseille"] = export["depart_conseille"].apply(fmt_dt)
    export["lien_appel"] = export["telephone_tel"].apply(lambda x: f"tel:{x}" if x else "")
    cols = ["ordre", "numero_rdv", "date_rdv", "heure_rdv", "depart_conseille", "pause_avant_rdv_min", "nom_prospect", "telephone", "adresse_complete", "distance_depuis_precedent_km", "temps_route_depuis_precedent_min", "source_temps", "waze", "google_maps", "street_view", "lien_appel"]
    if return_row:
        export = pd.concat([export[cols], pd.DataFrame([{c: return_row.get(c, "") for c in cols}])], ignore_index=True)
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


with st.sidebar:
    st.header("Réglages")
    start_address = st.text_input("Adresse de départ / retour", value=DEFAULT_START)
    safety_min = st.number_input("Marge sécurité avant RDV", min_value=0, max_value=60, value=15, step=5)
    visit_min = st.number_input("Durée moyenne d'un RDV", min_value=15, max_value=240, value=150, step=15)
    use_google = st.checkbox("Utiliser Google trafic / Street View si j'ai une clé API", value=False)
    google_key = st.text_input("Clé Google Maps API (optionnel)", type="password") if use_google else ""
    uploaded = st.file_uploader("Importer ton fichier Excel", type=["xlsx", "xls"])
    saved = st.file_uploader("Ou charger un récap CSV sauvegardé", type=["csv"], key="saved_csv")
    st.info("V10 : colonnes fixes, ordre par heure RDV, retour base, pauses et PDF enrichi.")

if uploaded:
    try:
        df = prepare_dataframe(uploaded)
        if df.empty:
            st.error("Aucune adresse trouvée dans le fichier.")
            st.stop()
        st.success(f"{len(df)} RDV chargés. Ordre imposé par date + heure de RDV.")
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
        route_df = pd.read_csv(saved, sep=";")
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

route_df = st.session_state["route_df"]
return_row = st.session_state.get("return_row")
start_address = st.session_state.get("start_address", DEFAULT_START)
start_geo = st.session_state.get("start_geo", {})
google_key = st.session_state.get("google_key", "")
use_google = st.session_state.get("use_google", False)

total_km = pd.to_numeric(route_df.get("distance_depuis_precedent_km"), errors="coerce").fillna(0).sum() + (return_row.get("distance_depuis_precedent_km", 0) if return_row else 0)
total_min = pd.to_numeric(route_df.get("temps_route_depuis_precedent_min"), errors="coerce").fillna(0).sum() + (return_row.get("temps_route_depuis_precedent_min", 0) if return_row and isinstance(return_row.get("temps_route_depuis_precedent_min"), int) else 0)

col1, col2, col3, col4 = st.columns(4)
col1.metric("RDV", len(route_df))
col2.metric("Distance retour inclus", f"{total_km:.1f} km")
col3.metric("Temps route", fmt_duration(total_min))
if not route_df.empty:
    first_dep = route_df.iloc[0].get("depart_conseille")
    col4.metric("Premier départ conseillé", fmt_dt(first_dep))

st.subheader("🧭 Fil conducteur terrain")
timeline_df = pd.DataFrame(build_timeline(route_df, return_row, start_address, int(visit_min)))
st.dataframe(timeline_df, use_container_width=True, hide_index=True)

st.subheader("📊 Détail des trajets étape par étape")
show_cols = ["numero_rdv", "heure_rdv", "depart_conseille", "pause_avant_rdv_min", "nom_prospect", "telephone", "adresse_complete", "distance_depuis_precedent_km", "temps_route_depuis_precedent_min", "note_trafic"]
display_df = route_df[show_cols].copy()
display_df["heure_rdv"] = display_df["heure_rdv"].apply(fmt_time)
display_df["depart_conseille"] = display_df["depart_conseille"].apply(fmt_dt)
display_df["pause_avant_rdv_min"] = display_df["pause_avant_rdv_min"].apply(lambda x: "" if x == "" else fmt_duration(x))
display_df["temps_route_depuis_precedent_min"] = display_df["temps_route_depuis_precedent_min"].apply(fmt_duration)
display_df = display_df.rename(columns={
    "numero_rdv": "N° RDV", "heure_rdv": "Heure RDV", "depart_conseille": "Départ conseillé",
    "pause_avant_rdv_min": "Pause avant RDV", "nom_prospect": "Client", "telephone": "Téléphone",
    "adresse_complete": "Adresse", "distance_depuis_precedent_km": "Km depuis précédent",
    "temps_route_depuis_precedent_min": "Temps depuis précédent", "note_trafic": "Calcul"
})
st.dataframe(display_df, use_container_width=True, hide_index=True)
if return_row:
    st.info(f"Retour base inclus : {return_row.get('distance_depuis_precedent_km','')} km · {fmt_duration(return_row.get('temps_route_depuis_precedent_min',''))}")

st.subheader("📋 Mode terrain")
for _, r in route_df.iterrows():
    pause = r.get('pause_avant_rdv_min', '')
    pause_txt = "" if pause == "" else (f" · Pause dispo : {fmt_duration(pause)}" if int(pause) >= 0 else f" · ⚠ Retard probable : {fmt_duration(abs(int(pause)))}")
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
    st.download_button("📄 Télécharger PDF enrichi cliquable", data=pdf_bytes, file_name="tournee_terrain_v10.pdf", mime="application/pdf", use_container_width=True)
with c2:
    st.download_button("💾 Sauvegarde CSV réutilisable", data=csv_bytes, file_name="tournee_sauvegarde_v10.csv", mime="text/csv", use_container_width=True)

st.caption("Sans clé Google, le trafic est une estimation prudente. Avec une clé Google Maps API, l'app peut utiliser les durées trafic Google et intégrer des images Street View dans le PDF.")
