import io
import zipfile
import xml.etree.ElementTree as ET
import re
import math
import json
import unicodedata
from pathlib import Path
from datetime import datetime, date, time as dtime, timedelta, timezone
from zoneinfo import ZoneInfo
from urllib.parse import quote_plus

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import requests
from PIL import Image as PILImage
from geopy.geocoders import Nominatim
from geopy.extra.rate_limiter import RateLimiter
from geopy.distance import geodesic
import folium
from streamlit_folium import st_folium

try:
    from streamlit_geolocation import streamlit_geolocation
    GEOLOCATION_COMPONENT_AVAILABLE = True
except Exception:
    streamlit_geolocation = None
    GEOLOCATION_COMPONENT_AVAILABLE = False
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

st.set_page_config(page_title="Routage PRO V28.5 — 28/07/2026", page_icon="🚗", layout="wide")

DEFAULT_START = "72 avenue des Tourelles, 94490 Ormesson-sur-Marne"
AVG_SPEED_KMH = 38

def _get_data_dir():
    # Espace de sauvegarde serveur : tient aux simples rafraîchissements de page.
    # Sur Streamlit Cloud, il peut être réinitialisé lors d'un redémarrage ou redéploiement.
    for candidate in [Path.home() / ".routage_pro", Path("/tmp/routage_pro")]:
        try:
            candidate.mkdir(parents=True, exist_ok=True)
            test = candidate / ".write_test"
            test.write_text("ok", encoding="utf-8")
            test.unlink(missing_ok=True)
            return candidate
        except Exception:
            pass
    return Path("/tmp")

DATA_DIR = _get_data_dir()
LAST_UPLOAD_PATH = DATA_DIR / "dernier_fichier.xlsx"
IK_HISTORY_DIR = DATA_DIR / "ik_history"
APP_STATE_PATH = DATA_DIR / "settings_v20.json"
APP_STATE_FALLBACK_PATHS = [Path("/tmp/routage_pro_v19_settings.json"), DATA_DIR / "settings_v19.json"]
CRM_HISTORY_PATH = DATA_DIR / "crm_v20.csv"
CRM_HISTORY_FALLBACK_PATHS = [Path("/tmp/routage_pro_v19_crm.csv"), DATA_DIR / "crm_v19.csv"]

COLS = {
    "numero_rdv": 0, "adresse": 1, "code_postal": 2, "date_rdv": 3, "heure_debut": 4,
    "email": 5, "fournisseur": 7, "commercial_nom": 8, "nom": 9, "telepros_nom": 11,
    "commercial_prenom": 12, "prenom": 13, "telepros_prenom": 14, "telephone": 16, "ville": 17,
}

# Colonnes ajoutées par le robot CRM local V25/V26.
CRM_DETAIL_COLUMNS = ["Remarque", "remarque", "details_crm_ia", "details_crm", "Détails CRM", "Analyse IA"]

def _norm_col_name(value):
    """Normalise un nom de colonne Excel pour retrouver Remarque / details_crm_ia même si le robot change un peu le libellé."""
    try:
        s = str(value or "").strip().lower()
        s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
        s = re.sub(r"[^a-z0-9]+", "", s)
        return s
    except Exception:
        return str(value or "").strip().lower()


def safe_get_named(row, names, default=""):
    """Récupère une valeur par nom de colonne, même si la casse/accents/espaces changent légèrement."""
    try:
        normalized_cols = {_norm_col_name(c): c for c in row.index}
        # 1) match exact normalisé
        for name in names:
            key = _norm_col_name(name)
            if key in normalized_cols:
                val = row.get(normalized_cols[key], default)
                if pd.isna(val):
                    return default
                return str(val).strip()
        # 2) match souple : utile si la colonne s'appelle par exemple "Remarque client"
        for name in names:
            key = _norm_col_name(name)
            for norm_col, real_col in normalized_cols.items():
                if key and (key in norm_col or norm_col in key):
                    val = row.get(real_col, default)
                    if pd.isna(val):
                        return default
                    return str(val).strip()
    except Exception:
        pass
    return default

def latest_local_crm_export():
    """Détecte le dernier Excel enrichi généré par le robot local dans exports_crm/."""
    candidates = []
    for base in [Path.cwd(), Path(__file__).resolve().parent]:
        d = base / "exports_crm"
        if d.exists():
            candidates += list(d.glob("rdv_enrichi_IA_*.xlsx"))
            candidates += list(d.glob("*.xlsx"))
    candidates = [c for c in candidates if c.exists()]
    if not candidates:
        return None
    return max(candidates, key=lambda x: x.stat().st_mtime)

st.title("🚗 Routage PRO · GDH — V28.5 — 28/07/2026")
st.caption("Copilote terrain · trafic Google · Waze · Voir maison · CRM · rappels · IK")

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
    padding-top: 4.8rem;
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
    height: 3.4rem !important;
}
/* Evite que le bandeau Streamlit coupe le titre sur PC */
header[data-testid="stHeader"] + div {
    padding-top: 0.5rem !important;
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

/* V20.2 — garde le scroll iPhone fluide quand la carte est verrouillée */
.map-scroll-safe iframe {
    pointer-events: none !important;
}

/* Force les textes Streamlit à rester lisibles sur PC */
.stMarkdown, .stMarkdown p, .stMarkdown span, label, div[data-testid="stText"] {
    color: #f5f5f5 !important;
}


/* V21.2 — rappels intelligents */
.reminder-card {
    border-radius: 16px;
    padding: 14px 16px;
    margin: 10px 0 8px 0;
    border: 2px solid #333;
    background: #111827;
    color: #fff;
}
.reminder-card strong { font-size: 1.08rem; }
.reminder-today {
    background: linear-gradient(90deg, #7f1d1d, #f97316);
    border-color: #fde047;
    box-shadow: 0 0 18px rgba(249, 115, 22, 0.85);
    animation: rappelPulse 1s infinite alternate;
}
.reminder-late {
    background: linear-gradient(90deg, #450a0a, #991b1b);
    border-color: #ef4444;
}
.reminder-future {
    background: #111827;
    border-color: #2563eb;
}
.reminder-treated {
    background: #0f172a;
    border-color: #475569;
    opacity: 0.75;
}
@keyframes rappelPulse {
    from { filter: brightness(1); transform: scale(1); }
    to { filter: brightness(1.28); transform: scale(1.01); }
}

/* V28.5 — lisibilité des champs IA désactivés sur iPhone */
textarea:disabled {
    -webkit-text-fill-color: #111827 !important;
    color: #111827 !important;
    background: #f8fafc !important;
    opacity: 1 !important;
    border: 1px solid #cbd5e1 !important;
}


/* V27 — Interface terrain premium iPhone */
@media (max-width: 768px) {
    .block-container { padding-top: 3.9rem !important; padding-left: .65rem !important; padding-right: .65rem !important; }
    h1 { font-size: 1.45rem !important; line-height: 1.2 !important; }
    h2 { font-size: 1.2rem !important; }
    h3 { font-size: 1.05rem !important; }
}
.day-summary { background:linear-gradient(135deg,#111827,#1f2937); border:1px solid #374151; border-radius:18px; padding:14px 16px; margin:8px 0 12px 0; box-shadow:0 8px 26px rgba(0,0,0,.24); }
.day-summary-title { font-size:1.05rem; font-weight:900; color:#fff; margin-bottom:6px; }
.day-summary-line { font-size:.98rem; font-weight:700; color:#e5e7eb; line-height:1.55; }
.next-card { background:linear-gradient(135deg,#0f172a,#172554); border:2px solid #2563eb; border-radius:20px; padding:16px; margin:12px 0; box-shadow:0 10px 30px rgba(37,99,235,.20); }
.next-card .eyebrow { color:#93c5fd; font-weight:900; letter-spacing:.06em; font-size:.76rem; }
.next-card .time { color:#fff; font-size:2.05rem; font-weight:950; line-height:1; margin-top:5px; }
.next-card .client { color:#fff; font-size:1.28rem; font-weight:900; margin-top:6px; }
.next-card .address { color:#dbeafe; font-size:.94rem; margin-top:3px; }
.next-card .route { color:#fff; font-size:1.02rem; font-weight:800; margin-top:12px; }
.next-card .depart { display:inline-block; margin-top:10px; padding:7px 10px; border-radius:10px; background:#052e16; border:1px solid #22c55e; color:#dcfce7; font-weight:900; }
.route-card { background:#111827; border:1px solid #374151; border-radius:16px; padding:12px 14px; margin:8px 0; }
.route-card strong { color:#fff; }
.route-card .muted { color:#cbd5e1; }
.route-card .go { color:#86efac; font-weight:850; }
.map-legend { font-size:.82rem; color:#cbd5e1; margin:-2px 0 8px 0; }


/* V28 — Cockpit premium inspiré des apps de conduite modernes */
.stApp {
    background:
      radial-gradient(circle at 20% -10%, rgba(0,194,255,.12), transparent 28%),
      radial-gradient(circle at 95% 8%, rgba(44,95,255,.12), transparent 26%),
      #050810 !important;
}
[data-testid="stSidebar"] {
    background: linear-gradient(180deg,#07101c 0%,#090d16 100%) !important;
    border-right:1px solid #172235 !important;
}
.block-container { max-width: 1500px; }
h1,h2,h3 { letter-spacing:-.02em; }
div[data-testid="stExpander"]{
    background:rgba(10,17,29,.72)!important;
    border:1px solid #1e2b40!important;
    border-radius:18px!important;
}
.stButton button,.stDownloadButton button,.stLinkButton a{
    min-height:46px!important;
    background:linear-gradient(180deg,#132238,#0c1727)!important;
    border:1px solid #29405f!important;
    box-shadow:0 6px 18px rgba(0,0,0,.22)!important;
}
.stButton button:hover,.stLinkButton a:hover{
    border-color:#00c2ff!important;
    box-shadow:0 0 0 1px rgba(0,194,255,.25),0 8px 24px rgba(0,194,255,.14)!important;
}
.cockpit-head{
    display:flex;justify-content:space-between;gap:12px;align-items:center;
    background:linear-gradient(135deg,rgba(8,17,31,.94),rgba(8,22,40,.82));
    border:1px solid #1d3552;border-radius:22px;padding:14px 16px;margin:8px 0 12px;
    box-shadow:0 18px 45px rgba(0,0,0,.28);
}
.cockpit-date{font-size:1.12rem;font-weight:950;color:#fff}
.cockpit-stats{font-size:.9rem;font-weight:750;color:#a9bdd8;margin-top:4px}
.live-pills{display:flex;gap:6px;flex-wrap:wrap;justify-content:flex-end}
.live-pill{font-size:.72rem;font-weight:900;padding:7px 9px;border-radius:999px;background:#0b1b2d;border:1px solid #21425f;color:#b9edff}
.live-pill.ok{background:#072319;border-color:#116b4a;color:#70f3bc}
.live-pill.alert{background:#2c1016;border-color:#7e2333;color:#ff9aaa}
.next-card{
    background:linear-gradient(145deg,#07101d 0%,#0a1d36 60%,#08294b 100%)!important;
    border:1px solid #1f9ee8!important;
    box-shadow:0 18px 45px rgba(0,140,255,.17)!important;
}
.next-card .depart{
    background:#07281e!important;border:1px solid #1ac683!important;color:#9bffd5!important;
}
.route-card{background:rgba(10,16,27,.72)!important;border:1px solid #1d2b41!important}
.day-summary{display:none!important}
@media(max-width:768px){
    .cockpit-head{display:block}
    .live-pills{justify-content:flex-start;margin-top:10px}
}

.terrain-anchor{scroll-margin-top:82px;}
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


def extract_departement(text):
    m = re.search(r"\b(\d{2})\d{3}\b", str(text or ""))
    if not m:
        return ""
    dep = m.group(1)
    if dep == "97":
        m2 = re.search(r"\b(97\d)\d{2}\b", str(text or ""))
        return m2.group(1) if m2 else dep
    return dep


def whatsapp_report_link(client="", departement="", statut="", commentaire="", rappel_date="", rappel_heure="", adresse="", telephone=""):
    lines = [
        f"RDV {str(client or '').strip()}" + (f" ({departement})" if departement else ""),
        "",
        f"Statut : {statut or ''}",
        f"Téléphone : {telephone or ''}",
        f"Adresse : {adresse or ''}",
        "",
        "Note :",
        str(commentaire or '').strip(),
    ]
    if rappel_date or rappel_heure:
        lines += ["", f"Rappel : {rappel_date or ''} {rappel_heure or ''}".strip()]
    lines += ["", "- Mr Dahan"]
    msg = "\n".join(lines)
    return "https://wa.me/?text=" + quote_plus(msg)



def analyse_crm_details(details_text):
    """Analyse terrain renforcée des infos copiées depuis le CRM.
    V23 : plus détaillée, orientée closing, avec plan d'entretien et phrases utilisables.
    Fonctionne sans API externe pour rester fiable sur Streamlit Cloud.
    """
    raw = str(details_text or "").strip()
    t = raw.lower()
    if not raw:
        return ""

    def has(*words):
        return any(str(w).lower() in t for w in words)

    def find_amounts():
        amounts = re.findall(r"(?:\b\d{2,4}\s*(?:€|eur|euro|euros)|\b\d{2,4}\s*/\s*mois)", raw, flags=re.I)
        return list(dict.fromkeys([a.strip() for a in amounts]))[:6]

    def find_surface():
        m = re.search(r"(\d{2,3})\s*m\s*(?:2|²)", raw, flags=re.I)
        return m.group(0) if m else ""

    def find_age():
        m = re.search(r"\b(\d{2})\s*ans\b", raw, flags=re.I)
        return m.group(1) + " ans" if m else ""

    signals = []
    profil = []
    douleurs = []
    axes = []
    objections = []
    questions = []
    phrases = []
    erreurs = []
    pieces = []
    closing_steps = []

    surface = find_surface()
    age = find_age()
    amounts = find_amounts()

    # Profil / décision
    if has("retrait", "retraite", "retraité", "retraitée"):
        profil.append("Client retraité : privilégier un rythme calme, rassurant, très concret, sans jargon technique.")
        axes.append("Mettre en avant le confort au quotidien, la stabilité de température, moins de manutention et un accompagnement administratif complet.")
        objections.append("Crainte de se faire embarquer dans des démarches compliquées ou un financement trop long.")
    if has("mari", "marié", "mariée", "épouse", "epouse", "conjoint", "conjointe", "madame", "monsieur"):
        profil.append("Décision probablement à deux : valider rapidement qui doit être convaincu et éviter de closer si un décisionnaire manque.")
        questions.append("Est-ce que vous décidez ensemble ? Si on trouve une solution cohérente, qui doit absolument valider avant de lancer ?")
    if has("vacance", "vacances", "retour", "pas dispo"):
        profil.append("Le client a déjà eu une contrainte de disponibilité : être efficace, montrer que le RDV ne sera pas une perte de temps.")

    # Chauffage et douleurs
    if has("bois", "cheminée", "cheminee", "poêle", "poele", "insert", "granulé", "granules"):
        douleurs.append("Chauffage bois : manutention, saleté, stockage, température irrégulière, contraintes d'âge/santé.")
        axes.append("Angle fort : garder le bois en appoint/plaisir, mais avoir un chauffage automatique confortable au quotidien.")
        phrases.append("L'idée n'est pas forcément de supprimer le bois si vous l'aimez, mais de ne plus dépendre de lui tous les jours.")
    if has("gaz"):
        douleurs.append("Gaz : facture mensuelle, dépendance aux hausses, entretien chaudière, rendement qui baisse avec l'âge.")
        axes.append("Comparer la mensualité/économie à la dépense actuelle et parler de modernisation du système.")
        phrases.append("Aujourd'hui vous payez déjà votre chauffage ; l'objectif est de transformer une dépense subie en solution durable et mieux aidée.")
    if has("fioul", "fuel"):
        douleurs.append("Fioul : énergie coûteuse, livraison, odeur, entretien, image vieillissante.")
        axes.append("Sortie du fioul = argument très fort : confort, aides, valeur maison, sérénité.")
    if has("électrique", "electrique", "convecteur", "grille pain", "radiateur électrique"):
        douleurs.append("Électrique direct : facture élevée et confort souvent médiocre.")
        axes.append("Insister sur le rendement et la baisse de consommation si le dimensionnement est cohérent.")
    if amounts:
        signals.append("Montants repérés : " + ", ".join(amounts) + ". S'en servir pour comparer au reste à charge ou à la mensualité.")
    if surface:
        signals.append(f"Surface repérée : {surface}. À utiliser pour crédibiliser le dimensionnement et les économies.")
    if age:
        signals.append(f"Âge repéré : {age}. Adapter le discours : simplicité, sécurité, pas de complexité administrative.")

    # Aides / revenus / maison
    if has("rfr", "revenu fiscal", "bleu", "jaune", "modeste", "très modeste", "tres modeste", "maprimerenov", "prime"):
        axes.append("Aides : présenter comme une opportunité à sécuriser, sans promettre de montant définitif avant vérification.")
        questions.append("Vous avez bien votre dernier avis d'imposition ? C'est ce qui permet de verrouiller les aides et d'éviter les mauvaises surprises.")
        pieces.append("Avis d'imposition / RFR")
    if has("propriétaire", "proprietaire", "maison", "résidence principale", "residence principale"):
        signals.append("Maison / résidence principale : bon terrain pour parler confort, valorisation du bien et solution long terme.")
    if has("isolation", "combles", "ite", "iti", "fenêtre", "fenetre", "simple vitrage", "sous-sol", "sous sol"):
        douleurs.append("Isolation à vérifier : peut conditionner le confort, les économies et parfois l'éligibilité.")
        questions.append("Qu'est-ce qui a déjà été isolé, en quelle année, et par qui ?")

    # Méfiance / objection
    if has("méfiant", "mefiant", "arnaque", "harcel", "déjà appelé", "deja appelé", "doute", "pas intéressé", "pas interesse"):
        objections.append("Méfiance forte : ne pas commencer par vendre. Commencer par expliquer qui tu es, pourquoi tu es là, et comment le client garde le contrôle.")
        phrases.append("Je comprends votre méfiance, vous avez sûrement été beaucoup sollicité. Mon but aujourd'hui c'est d'abord de vérifier si le projet est cohérent, pas de vous forcer la main.")
        erreurs.append("Éviter les phrases type 'offre exceptionnelle' ou 'il faut signer maintenant' trop tôt.")
    if has("devis", "concurrent", "réfléchir", "reflechir", "voir", "comparer"):
        objections.append("Client en comparaison/réflexion : il faudra vendre la sécurité, l'accompagnement, les délais, les garanties, pas seulement le prix.")
        phrases.append("Vous avez raison de comparer. Mon rôle c'est de vous aider à comparer ce qui est vraiment comparable : matériel, pose, garanties, aides et suivi administratif.")
    if has("cher", "prix", "budget", "mensual", "crédit", "credit", "financement"):
        objections.append("Sensibilité prix/financement : parler d'abord valeur et économies avant de présenter mensualité ou reste à charge.")
        questions.append("Aujourd'hui, entre chauffage, entretien et inconfort, combien ce système vous coûte réellement par mois ?")
    if has("urgent", "rapidement", "vite", "panne", "froid"):
        axes.append("Urgence : mettre en avant la réactivité et la capacité à sécuriser une solution rapidement.")

    # Defaults
    if not signals:
        signals.append("Informations CRM à exploiter : situation familiale, chauffage actuel, facture, surface, motivation et freins.")
    if not profil:
        profil.append("Profil à découvrir sur place : décisionnaire, niveau de méfiance, urgence réelle, sensibilité au prix.")
    if not douleurs:
        douleurs.append("Douleurs à faire verbaliser : facture, inconfort, pannes, manutention, peur de l'avenir, valeur de la maison.")
    if not axes:
        axes.append("Axe principal : partir du problème du client, puis montrer que la solution répond précisément à ce problème.")
    if not objections:
        objections.extend(["Prix / reste à charge", "Besoin de réfléchir", "Méfiance envers les aides", "Peur des travaux"])
    if not questions:
        questions.extend([
            "Qu'est-ce qui vous a fait accepter le RDV ?",
            "Qu'est-ce qui vous dérange le plus dans votre chauffage actuel ?",
            "Si le projet est éligible et cohérent, qu'est-ce qui pourrait vous empêcher d'avancer ?",
        ])
    if not phrases:
        phrases.extend([
            "Je vais d'abord vérifier si le projet est cohérent chez vous, et seulement après on parlera solution.",
            "Le but n'est pas de vous vendre quelque chose d'inutile, mais de voir si les aides rendent le projet intéressant.",
        ])
    if not erreurs:
        erreurs.extend([
            "Ne pas annoncer un montant d'aide définitif sans vérification.",
            "Ne pas aller trop vite au prix avant d'avoir fait exprimer le besoin.",
            "Ne pas parler uniquement technique : le client achète surtout du confort, de la sécurité et de la simplicité.",
        ])
    if not pieces:
        pieces.extend(["Avis d'imposition", "Factures énergie", "Photos/infos chauffage actuel", "Surface et isolation"])

    closing_steps = [
        "1. Rassurer : expliquer ton rôle, le déroulé du RDV et que rien n'est validé sans vérification.",
        "2. Découverte : faire parler le client 5-10 minutes sur chauffage, factures, confort, travaux déjà faits.",
        "3. Reformulation : résumer son problème avec ses mots pour créer l'accord.",
        "4. Diagnostic : vérifier maison, chauffage, isolation, place matériel, contraintes techniques.",
        "5. Valeur : relier la solution aux problèmes exprimés, pas à une fiche produit.",
        "6. Aides/financement : présenter prudemment le scénario, puis le reste à charge/mensualité.",
        "7. Closing doux : demander ce qui bloque réellement et traiter une objection à la fois.",
    ]

    # Priorité commerciale
    if has("veut réfléchir", "veut reflechir", "réfléchir", "reflechir"):
        priorite = "🟠 Priorité : client à travailler en réassurance. Objectif RDV : comprendre le vrai frein et programmer une suite claire."
    elif has("urgent", "panne", "vite", "rapidement"):
        priorite = "🔴 Priorité : forte urgence possible. Objectif RDV : sécuriser rapidement la faisabilité et le délai."
    elif has("rfr", "modeste", "très modeste", "tres modeste", "bleu", "jaune"):
        priorite = "🟢 Priorité : potentiel aides intéressant. Objectif RDV : vérifier l'éligibilité et cadrer le reste à charge."
    else:
        priorite = "🟡 Priorité : à qualifier sur place. Objectif RDV : identifier douleur + décisionnaire + budget."

    def block(title, items, limit=10):
        lines = [title]
        for it in items[:limit]:
            lines.append(f"- {it}")
        return "\n".join(lines)

    return "\n\n".join([
        "🔥 SYNTHÈSE CLOSING TERRAIN",
        priorite,
        block("🎯 Signaux importants repérés", signals, 8),
        block("👤 Lecture du profil client", profil, 8),
        block("💥 Douleurs à faire ressortir", douleurs, 8),
        block("🧠 Angle de vente recommandé", axes, 10),
        block("⚠️ Objections probables + réponse à préparer", objections, 10),
        block("❓ Questions puissantes à poser", questions, 10),
        block("🗣️ Phrases utiles à dire sur place", phrases, 10),
        block("🧾 Pièces / points à vérifier", pieces, 8),
        block("🪜 Plan de closing étape par étape", closing_steps, 10),
        block("🚫 Erreurs à éviter", erreurs, 8),
        "📌 Objectif final du RDV\n- Obtenir soit une signature, soit une suite cadrée : document manquant, rappel daté, décisionnaire à revoir, ou objection précise à traiter.\n- Ne jamais repartir avec un simple 'je vais réfléchir' sans date de rappel et raison exacte."
    ])

def whatsapp_ai_prep_link(client="", departement="", adresse="", details="", analyse=""):
    msg = f"""Préparation RDV {client}{' (' + departement + ')' if departement else ''}
Adresse : {adresse}

Infos CRM :
{str(details or '').strip()}

Conseils terrain :
{str(analyse or '').strip()}

- Mr Dahan"""
    return "https://wa.me/?text=" + quote_plus(msg)


def streetview_link(lat, lon, address):
    # Lien volontairement simple et fiable sur iPhone/Windows : ouvre Google Maps sur l’adresse.
    # Les liens Street View directs sont instables et peuvent donner un écran noir selon Safari/Chrome.
    return f"https://www.google.com/maps/search/{quote_plus(address)}"


def directions_link(origin, destination):
    return f"https://www.google.com/maps/dir/?api=1&origin={quote_plus(origin)}&destination={quote_plus(destination)}&travelmode=driving"


def google_drive_to_link(destination):
    """Ouvre Google Maps en mode itinéraire vers le prospect, depuis la position courante si disponible."""
    return f"https://www.google.com/maps/dir/?api=1&destination={quote_plus(str(destination or ''))}&travelmode=driving"



FR_WEEKDAYS = ["Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi","Dimanche"]
FR_MONTHS = ["janvier","février","mars","avril","mai","juin","juillet","août","septembre","octobre","novembre","décembre"]

def french_long_date(d):
    if not isinstance(d, date):
        return ""
    return f"{FR_WEEKDAYS[d.weekday()]} {d.day} {FR_MONTHS[d.month-1]}"


def route_delay_status(row, row_index=0):
    """Retard probable en minutes vers ce RDV."""
    try:
        pause = row.get("pause_avant_rdv_min", "")
        if row_index > 0 and pause not in ("", None):
            return max(0, -int(to_minutes(pause)))

        advised = row.get("depart_conseille")
        rdv_dt = row.get("rdv_datetime")
        if isinstance(advised, datetime) and isinstance(rdv_dt, datetime):
            now = datetime.now(ZoneInfo("Europe/Paris")).replace(tzinfo=None)
            if now <= rdv_dt:
                return max(0, int((now - advised).total_seconds() // 60))
    except Exception:
        pass
    return 0


def find_next_rdv(df):
    if df is None or df.empty:
        return None
    now_paris = datetime.now(ZoneInfo("Europe/Paris")).replace(tzinfo=None)
    for _, rr in df.iterrows():
        dt = rr.get("rdv_datetime")
        if isinstance(dt, datetime) and dt >= now_paris:
            return rr
    return df.iloc[-1]

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


def decode_google_polyline(encoded):
    if not encoded:
        return []
    coords=[]; index=0; lat=0; lng=0
    try:
        while index < len(encoded):
            for is_lat in (True, False):
                result=0; shift=0
                while True:
                    b=ord(encoded[index])-63; index+=1
                    result |= (b & 0x1f) << shift; shift += 5
                    if b < 0x20: break
                delta = ~(result >> 1) if (result & 1) else (result >> 1)
                if is_lat: lat += delta
                else: lng += delta
            coords.append([lat/1e5,lng/1e5])
    except Exception:
        return []
    return coords

def money_to_float(money):
    if not isinstance(money, dict): return 0.0
    try:
        return float(money.get('units',0) or 0)+float(money.get('nanos',0) or 0)/1_000_000_000
    except Exception:
        return 0.0

ROUTE_PREF_LABELS={
    'recommended':'⚡ Recommandé',
    'no_tolls':'💶 Sans péage',
    'shortest':'📏 Plus court',
    'no_highways':'🛣️ Sans autoroute',
}

@st.cache_data(show_spinner=False, ttl=300)
def google_routes_traffic(origin, destination, departure_dt, api_key, route_pref="recommended", include_tolls=True):
    if not api_key or not isinstance(departure_dt, datetime): return None
    try:
        paris=ZoneInfo("Europe/Paris"); now_paris=datetime.now(paris)
        dep=departure_dt.replace(tzinfo=paris) if departure_dt.tzinfo is None else departure_dt.astimezone(paris)
        if dep < now_paris - timedelta(minutes=1): return None
        payload={
            "origin":{"address":str(origin)},"destination":{"address":str(destination)},
            "travelMode":"DRIVE","routingPreference":"TRAFFIC_AWARE_OPTIMAL","trafficModel":"BEST_GUESS",
            "departureTime":dep.astimezone(timezone.utc).isoformat().replace("+00:00","Z"),
            "languageCode":"fr-FR","regionCode":"FR","units":"METRIC",
            "routeModifiers":{"avoidTolls":route_pref=="no_tolls","avoidHighways":route_pref=="no_highways","vehicleInfo":{"emissionType":"DIESEL"}},
        }
        if include_tolls: payload["extraComputations"]=["TOLLS"]
        if route_pref=="shortest":
            payload["requestedReferenceRoutes"]=["SHORTER_DISTANCE"]; payload["routingPreference"]="TRAFFIC_AWARE"
        fields=["routes.duration","routes.staticDuration","routes.distanceMeters","routes.routeLabels","routes.routeToken","routes.polyline.encodedPolyline"]
        if include_tolls:
            fields.extend([
                "routes.travelAdvisory.tollInfo",
                "routes.legs.travelAdvisory.tollInfo",
            ])
        headers={"Content-Type":"application/json","X-Goog-Api-Key":api_key,"X-Goog-FieldMask":",".join(fields)}
        r=requests.post("https://routes.googleapis.com/directions/v2:computeRoutes",headers=headers,json=payload,timeout=15); r.raise_for_status()
        routes=r.json().get("routes",[])
        if not routes: return None
        route=routes[0]
        if route_pref=="shortest":
            shorter_candidates = [
                cand for cand in routes
                if "SHORTER_DISTANCE" in cand.get("routeLabels", [])
            ]
            if not shorter_candidates:
                return None
            # Si Google renvoie plusieurs variantes SHORTER_DISTANCE,
            # on garde réellement la moins kilométrée.
            route = min(
                shorter_candidates,
                key=lambda cand: float(cand.get("distanceMeters", 10**15) or 10**15)
            )
        def sec(v):
            try:return float(str(v).rstrip("s"))
            except:return 0.0
        traffic_min=sec(route.get("duration","0s"))/60; static_min=sec(route.get("staticDuration","0s"))/60; km=float(route.get("distanceMeters",0) or 0)/1000
        route_toll_info = route.get("travelAdvisory", {}).get("tollInfo", {})
        route_est = route_toll_info.get("estimatedPrice", []) if isinstance(route_toll_info, dict) else []

        leg_toll_infos = []
        for leg in route.get("legs", []) or []:
            info = leg.get("travelAdvisory", {}).get("tollInfo", {})
            if isinstance(info, dict) and info:
                leg_toll_infos.append(info)

        toll_detected = bool(route_toll_info) or bool(leg_toll_infos)
        toll_amount = None
        toll_known = False

        # Priorité au total d'itinéraire renvoyé par Google.
        if route_est:
            toll_amount = sum(money_to_float(m) for m in route_est if isinstance(m, dict))
            toll_known = True
        elif leg_toll_infos:
            # Secours : addition des prix connus par étape.
            leg_prices = []
            all_legs_priced = True
            for info in leg_toll_infos:
                prices = info.get("estimatedPrice", []) if isinstance(info, dict) else []
                if not prices:
                    all_legs_priced = False
                    continue
                leg_prices.append(sum(money_to_float(m) for m in prices if isinstance(m, dict)))
            if leg_prices and all_legs_priced:
                toll_amount = sum(leg_prices)
                toll_known = True

        if traffic_min<=0 or km<=0:return None
        return {
            "km":km,
            "min":traffic_min,
            "static_min":static_min,
            "traffic_delay_min":max(0.0,traffic_min-static_min),
            "source":"Google Routes trafic",
            "geometry":decode_google_polyline(route.get("polyline",{}).get("encodedPolyline","")),
            "toll_amount":round(toll_amount,2) if toll_amount is not None else None,
            "toll_known":toll_known,
            "toll_detected":toll_detected,
            "route_pref":route_pref,
        }
    except Exception:
        return None


def google_routes_diagnostic(api_key, origin="72 avenue des Tourelles, 94490 Ormesson-sur-Marne", destination="10 Rue de la Planche, 89210 Esnon"):
    """Test direct de Google Routes API et retourne un diagnostic lisible, sans afficher la clé."""
    if not api_key:
        return {"ok": False, "status": None, "message": "Aucune clé trouvée dans Streamlit Secrets."}

    try:
        paris = ZoneInfo("Europe/Paris")
        dep = datetime.now(paris) + timedelta(minutes=10)
        payload = {
            "origin": {"address": str(origin)},
            "destination": {"address": str(destination)},
            "travelMode": "DRIVE",
            "routingPreference": "TRAFFIC_AWARE_OPTIMAL",
            "trafficModel": "BEST_GUESS",
            "departureTime": dep.astimezone(timezone.utc).isoformat().replace("+00:00", "Z"),
            "languageCode": "fr-FR",
            "regionCode": "FR",
            "units": "METRIC",
        }
        headers = {
            "Content-Type": "application/json",
            "X-Goog-Api-Key": api_key,
            "X-Goog-FieldMask": "routes.duration,routes.staticDuration,routes.distanceMeters",
        }
        r = requests.post(
            "https://routes.googleapis.com/directions/v2:computeRoutes",
            headers=headers,
            json=payload,
            timeout=15,
        )
        try:
            body = r.json()
        except Exception:
            body = {"raw": (r.text or "")[:1000]}

        if r.ok:
            routes = body.get("routes", []) if isinstance(body, dict) else []
            if routes:
                return {"ok": True, "status": r.status_code, "message": "Google Routes répond correctement."}
            return {"ok": False, "status": r.status_code, "message": "Google a répondu, mais aucun itinéraire n'a été renvoyé."}

        msg = ""
        if isinstance(body, dict):
            err = body.get("error", {})
            if isinstance(err, dict):
                msg = str(err.get("message", "") or "")
                status_txt = str(err.get("status", "") or "")
                if status_txt and status_txt not in msg:
                    msg = f"{status_txt} — {msg}".strip(" —")
        if not msg:
            msg = str(body)[:1000]
        return {"ok": False, "status": r.status_code, "message": msg}
    except Exception as e:
        return {"ok": False, "status": None, "message": f"{type(e).__name__}: {e}"}


def traffic_factor(arrival_dt):
    if not isinstance(arrival_dt, datetime):
        return 1.25
    h = arrival_dt.hour + arrival_dt.minute / 60
    if 7 <= h <= 10 or 16.5 <= h <= 20:
        return 1.55
    if 11 <= h <= 16.5:
        return 1.25
    return 1.12


def extra_step_to_row(step):
    d=step.get("date"); h=step.get("heure"); phone_fmt,phone_digits=format_phone(step.get("telephone","")); label=step.get("nom","").strip() or step.get("type","Étape")
    return {"ordre":0,"numero_rdv":"","numero_rdv_source":"","nom_prospect":label,"adresse":step.get("adresse",""),"code_postal":"","ville":"","adresse_complete":step.get("adresse",""),
            "date_rdv":d,"heure_rdv":h,"rdv_datetime":dt_from_row(d,h),"telephone":phone_fmt,"telephone_tel":phone_digits,"email":"","fournisseur":"","commercial":"","teleprospecteur":"",
            "remarque_crm":step.get("note",""),"details_crm":step.get("note",""),"analyse_ia_importee":"","type_etape":step.get("type","Autre"),"duree_etape_min":int(step.get("duree",30) or 30),"is_extra_step":True}

def apply_extra_steps(df):
    base=df.copy(); base["is_extra_step"]=False; base["type_etape"]="RDV client"; base["duree_etape_min"]=0
    steps=st.session_state.get("extra_steps",[])
    if not steps:return base
    extras=pd.DataFrame([extra_step_to_row(s) for s in steps]); merged=pd.concat([base,extras],ignore_index=True,sort=False)
    merged["__sort_dt"]=pd.to_datetime(merged["rdv_datetime"],errors="coerce"); merged=merged.sort_values("__sort_dt",na_position="last").drop(columns=["__sort_dt"]).reset_index(drop=True)
    merged["numero_rdv"]=range(1,len(merged)+1); merged["ordre"]=range(1,len(merged)+1); return merged

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
            # Données enrichies par le robot CRM local : elles alimentent directement la préparation IA.
            # Priorité : Remarque du CRM, puis details_crm_ia. Ces valeurs doivent pré-remplir le bloc "Préparation IA".
            "remarque_crm": safe_get_named(row, ["Remarque", "remarque", "Remarque client", "remarque_client"]),
            "details_crm": (
                safe_get_named(row, ["Remarque", "remarque", "Remarque client", "remarque_client"])
                or safe_get_named(row, ["details_crm_ia", "details_crm", "Détails CRM", "Details CRM", "crm_row_text"])
            ),
            "analyse_ia_importee": safe_get_named(row, ["Analyse IA", "analyse_ia", "analyse"]),
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


def route_between(prev_addr, prev_geo, addr, coord, departure_dt, api_key, use_google, route_pref="recommended", include_tolls=True):
    """Calcule un trajet.

    Le temps Google est prioritaire quand Routes API est disponible.
    OSRM reste utilisé pour tracer la route sur la carte et comme secours.
    """
    osrm = osrm_route(
        prev_geo.get("lat"), prev_geo.get("lon"),
        coord.get("lat"), coord.get("lon")
    )

    if use_google and api_key and isinstance(departure_dt, datetime):
        g = google_routes_traffic(prev_addr, addr, departure_dt, api_key, route_pref=route_pref, include_tolls=include_tolls)
        if g:
            if not g.get("geometry") and osrm:
                g["geometry"] = osrm.get("geometry", [])
            return g

    if osrm:
        return osrm

    if prev_geo.get("lat") and prev_geo.get("lon") and coord.get("lat") and coord.get("lon"):
        dist = geodesic(
            (prev_geo["lat"], prev_geo["lon"]),
            (coord["lat"], coord["lon"])
        ).km * 1.28
        mins = (dist / AVG_SPEED_KMH) * 60
        return {"km": dist, "min": mins, "source": "Estimation", "geometry": []}

    return {"km": None, "min": None, "source": "Non calculé", "geometry": []}


def enrich_route(df, start_address, safety_min, visit_min, use_google, api_key, route_pref="recommended", include_tolls=True):
    addresses = [start_address] + df["adresse_complete"].tolist()
    geo = geocode_addresses(addresses)

    prev_addr = start_address
    prev_geo = geo.get(start_address, {})
    previous_rdv_end = None
    out = []
    cumulative_km = 0.0
    cumulative_min = 0.0

    paris = ZoneInfo("Europe/Paris")

    for i, (_, row) in enumerate(df.iterrows()):
        addr = row["adresse_complete"]
        coord = geo.get(addr, {})
        arrival_dt = row.get("rdv_datetime")

        # Heure de départ utilisée pour demander le trafic :
        # - 1er RDV : première estimation 2 h avant, puis recalcul itératif ;
        # - RDV suivants : fin du RDV précédent, puis recalcul au départ conseillé.
        if isinstance(arrival_dt, datetime):
            initial_departure = (
                previous_rdv_end
                if isinstance(previous_rdv_end, datetime)
                else arrival_dt - timedelta(hours=2)
            )
        else:
            initial_departure = None

        rb = route_between(
            prev_addr, prev_geo, addr, coord,
            initial_departure, api_key, use_google, route_pref, include_tolls
        )

        # Pour Google, une seconde passe utilise l'heure de départ conseillée
        # afin que le trafic corresponde réellement à cette heure.
        if (
            use_google and api_key and isinstance(arrival_dt, datetime)
            and rb.get("source") == "Google Routes trafic"
            and rb.get("min") is not None
        ):
            first_min = int(math.ceil(rb["min"]))
            suggested = arrival_dt - timedelta(minutes=first_min + safety_min)

            rb2 = route_between(
                prev_addr, prev_geo, addr, coord,
                suggested, api_key, use_google, route_pref, include_tolls
            )
            if rb2.get("source") == "Google Routes trafic":
                rb = rb2

        km = rb.get("km")
        raw_min = rb.get("min")

        if raw_min is not None:
            if rb.get("source") == "Google Routes trafic":
                drive_min = int(math.ceil(raw_min))
                delay = int(round(float(rb.get("traffic_delay_min", 0) or 0)))
                traffic_note = (
                    f"Google trafic (+{delay} min)"
                    if delay > 0 else
                    "Google trafic"
                )
            elif rb.get("source") == "OSRM":
                # Important : OSRM n'est plus multiplié artificiellement.
                drive_min = int(math.ceil(raw_min))
                traffic_note = "sans trafic réel"
            else:
                drive_min = int(math.ceil(raw_min))
                traffic_note = "estimation sans trafic"
        else:
            drive_min = None
            traffic_note = "non calculé"

        advised_departure = (
            arrival_dt - timedelta(minutes=(drive_min or 0) + safety_min)
            if isinstance(arrival_dt, datetime) and drive_min is not None
            else None
        )

        if previous_rdv_end and advised_departure:
            pause_min = int((advised_departure - previous_rdv_end).total_seconds() // 60)
        else:
            pause_min = None

        step_duration = int(row.get("duree_etape_min",0) or 0) if bool(row.get("is_extra_step",False)) else int(visit_min)
        previous_rdv_end = arrival_dt + timedelta(minutes=step_duration) if isinstance(arrival_dt,datetime) else None

        cumulative_km += km or 0
        cumulative_min += drive_min or 0

        r = row.to_dict()
        r.update({
            "lat": coord.get("lat"),
            "lon": coord.get("lon"),
            "source_geocodage": coord.get("source", ""),
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
            "peage_estime": float(rb.get("toll_amount", 0) or 0),
            "peage_connu": bool(rb.get("toll_known", False)),
            "peage_detecte": bool(rb.get("toll_detected", False)),
            "route_pref": rb.get("route_pref", route_pref),
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
        last_duration=int(last.get("duree_etape_min",0) or 0) if bool(last.get("is_extra_step",False)) else int(visit_min)
        last_end=last.get("rdv_datetime")+timedelta(minutes=last_duration) if isinstance(last.get("rdv_datetime"),datetime) else None

        rb = route_between(
            last_addr, last_geo, start_address,
            geo.get(start_address, {}),
            last_end, api_key, use_google, route_pref, include_tolls
        )

        km = rb.get("km")
        raw_min = rb.get("min")

        if raw_min is not None:
            ret_min = int(math.ceil(raw_min))
            if rb.get("source") == "Google Routes trafic":
                delay = int(round(float(rb.get("traffic_delay_min", 0) or 0)))
                ret_note = f"Google trafic (+{delay} min)" if delay > 0 else "Google trafic"
            elif rb.get("source") == "OSRM":
                ret_note = "sans trafic réel"
            else:
                ret_note = "estimation sans trafic"
        else:
            ret_min = ""
            ret_note = "non calculé"

        return_row = {
            "ordre": "Retour",
            "numero_rdv": "BASE",
            "date_rdv": last.get("date_rdv", ""),
            "heure_rdv": "",
            "rdv_datetime": last_end,
            "nom_prospect": "Retour base",
            "telephone": "",
            "telephone_tel": "",
            "adresse_complete": start_address,
            "lat": geo.get(start_address, {}).get("lat"),
            "lon": geo.get(start_address, {}).get("lon"),
            "distance_depuis_precedent_km": round(km, 1) if km is not None else "",
            "temps_route_depuis_precedent_min": ret_min,
            "source_temps": rb.get("source", ""),
            "note_trafic": ret_note,
            "depart_conseille": last_end,
            "pause_avant_rdv_min": "",
            "marge_securite_min": 0,
            "distance_cumulee_km": round(cumulative_km + (km or 0), 1),
            "temps_route_cumule_min": int(cumulative_min + (ret_min if isinstance(ret_min, int) else 0)),
            "waze": waze_link(
                geo.get(start_address, {}).get("lat"),
                geo.get(start_address, {}).get("lon"),
                start_address
            ),
            "google_maps": maps_link(start_address),
            "street_view": maps_link(start_address),
            "itineraire_depuis_precedent": directions_link(last_addr, start_address),
            "route_geometry": rb.get("geometry", []),
            "peage_estime": float(rb.get("toll_amount", 0) or 0),
            "peage_connu": bool(rb.get("toll_known", False)),
            "peage_detecte": bool(rb.get("toll_detected", False)),
            "route_pref": rb.get("route_pref", route_pref),
        }

    return route_df, return_row, geo.get(start_address, {})



RADARS_DATASET_SLUG = "liste-des-radars-fixes-en-france"

@st.cache_data(show_spinner=False, ttl=86400)
def load_fixed_radars():
    """Charge dynamiquement le CSV le plus récent du jeu officiel data.gouv.fr."""
    try:
        meta_url = f"https://www.data.gouv.fr/api/1/datasets/{RADARS_DATASET_SLUG}/"
        meta = requests.get(meta_url, timeout=15)
        meta.raise_for_status()
        dataset = meta.json()

        resources = dataset.get("resources", [])
        csv_resources = [
            r for r in resources
            if str(r.get("format", "")).lower() == "csv" and r.get("url")
        ]
        if not csv_resources:
            return pd.DataFrame(columns=["lat","lon","type","vma","id"])

        csv_resources.sort(
            key=lambda r: str(r.get("last_modified") or r.get("latest") or ""),
            reverse=True
        )
        resource_url = csv_resources[0]["url"]

        resp = requests.get(resource_url, timeout=20)
        resp.raise_for_status()

        raw = None
        for enc in ["utf-8-sig", "utf-8", "latin-1"]:
            try:
                raw = pd.read_csv(io.BytesIO(resp.content), sep=None, engine="python", encoding=enc)
                if raw is not None and not raw.empty:
                    break
            except Exception:
                raw = None
        if raw is None or raw.empty:
            return pd.DataFrame(columns=["lat","lon","type","vma","id"])

        norm = {_norm_col_name(c): c for c in raw.columns}

        def find_col(candidates):
            for cand in candidates:
                nc = _norm_col_name(cand)
                if nc in norm:
                    return norm[nc]
            for nc, real in norm.items():
                for cand in candidates:
                    cc = _norm_col_name(cand)
                    if cc and (cc in nc or nc in cc):
                        return real
            return None

        lat_col = find_col(["latitude","lat"])
        lon_col = find_col(["longitude","lon","lng"])
        type_col = find_col(["type","type radar","type_de_radar","typeequipement"])
        vma_col = find_col(["vma","vitesse","vitesse maximale autorisee","vitessemaximaleautorisee"])
        id_col = find_col(["id","identifiant","id radar","idequipement"])

        if lat_col is None or lon_col is None:
            coord_col = find_col(["coordonnees","coordonnées","geopoint","geo_point_2d","position"])
            if coord_col:
                extracted = raw[coord_col].astype(str).str.extract(
                    r"(-?\d+(?:[.,]\d+)?)\s*[,; ]+\s*(-?\d+(?:[.,]\d+)?)"
                )
                raw["__lat"] = pd.to_numeric(extracted[0].str.replace(",", ".", regex=False), errors="coerce")
                raw["__lon"] = pd.to_numeric(extracted[1].str.replace(",", ".", regex=False), errors="coerce")
                lat_col, lon_col = "__lat", "__lon"

        if lat_col is None or lon_col is None:
            return pd.DataFrame(columns=["lat","lon","type","vma","id"])

        out = pd.DataFrame({
            "lat": pd.to_numeric(raw[lat_col].astype(str).str.replace(",", ".", regex=False), errors="coerce"),
            "lon": pd.to_numeric(raw[lon_col].astype(str).str.replace(",", ".", regex=False), errors="coerce"),
            "type": raw[type_col].astype(str) if type_col else "",
            "vma": raw[vma_col].astype(str) if vma_col else "",
            "id": raw[id_col].astype(str) if id_col else "",
        }).dropna(subset=["lat","lon"])

        out = out[
            out["lat"].between(41.0, 51.8)
            & out["lon"].between(-5.5, 10.2)
        ].copy()
        return out.reset_index(drop=True)
    except Exception:
        return pd.DataFrame(columns=["lat","lon","type","vma","id"])


def radars_near_route(df, return_row=None, current_position=None, margin_deg=0.06):
    """Retourne uniquement les radars proches du tracé routier réel de la tournée."""
    route_pts = []

    def add_geometry(geom):
        if not isinstance(geom, list) or not geom:
            return
        # Jusqu'à ~500 points pour conserver une bonne précision sans alourdir l'iPhone.
        step = max(1, len(geom) // 500)
        for p in geom[::step]:
            if isinstance(p, (list, tuple)) and len(p) >= 2:
                try:
                    route_pts.append((float(p[0]), float(p[1])))
                except Exception:
                    pass

    for _, rr in df.iterrows():
        add_geometry(rr.get("route_geometry", []))
    if return_row:
        add_geometry(return_row.get("route_geometry", []))

    if not route_pts:
        return pd.DataFrame(columns=["lat","lon","type","vma","id"])

    all_radars = load_fixed_radars()
    if all_radars.empty:
        return all_radars

    lats = [p[0] for p in route_pts]
    lons = [p[1] for p in route_pts]
    subset = all_radars[
        all_radars["lat"].between(min(lats)-margin_deg, max(lats)+margin_deg)
        & all_radars["lon"].between(min(lons)-margin_deg, max(lons)+margin_deg)
    ].copy()

    if subset.empty:
        return subset

    # Distance approchée en km vers le point échantillonné le plus proche.
    # 1° latitude ≈111 km ; longitude corrigée autour de la France.
    import math as _math
    mean_lat = sum(lats) / len(lats)
    lon_km = 111.0 * _math.cos(_math.radians(mean_lat))
    max_dist_km = 1.2

    keep = []
    for idx, rd in subset.iterrows():
        rlat, rlon = float(rd["lat"]), float(rd["lon"])
        best2 = None
        for plat, plon in route_pts:
            dy = (rlat - plat) * 111.0
            dx = (rlon - plon) * lon_km
            d2 = dx*dx + dy*dy
            if best2 is None or d2 < best2:
                best2 = d2
                if best2 <= 0.04:  # ~200 m
                    break
        if best2 is not None and best2 <= max_dist_km * max_dist_km:
            keep.append(idx)

    return subset.loc[keep].head(250).reset_index(drop=True)



FUEL_REALTIME_URL = "https://donnees.roulez-eco.fr/opendata/instantane"

@st.cache_data(show_spinner=False, ttl=600)
def load_fuel_stations():
    """Flux gouvernemental instantané des stations-service et prix carburants."""
    cols = ["lat","lon","adresse","ville","cp","brand","gazole","sp95","e10","sp98","e85","gplc","updated"]
    try:
        r = requests.get(FUEL_REALTIME_URL, timeout=25)
        r.raise_for_status()
        with zipfile.ZipFile(io.BytesIO(r.content)) as z:
            xml_names = [n for n in z.namelist() if n.lower().endswith(".xml")]
            if not xml_names:
                return pd.DataFrame(columns=cols)
            xml_bytes = z.read(xml_names[0])

        root = ET.fromstring(xml_bytes)
        rows = []
        for pdv in root.findall(".//pdv"):
            try:
                lat = float(pdv.attrib.get("latitude")) / 100000.0
                lon = float(pdv.attrib.get("longitude")) / 100000.0
            except Exception:
                continue

            adresse = (pdv.findtext("adresse") or "").strip()
            ville = (pdv.findtext("ville") or "").strip()
            cp = str(pdv.attrib.get("cp", "") or "")
            brand = ""
            for possible in ["marque", "enseigne", "nom"]:
                node = pdv.find(possible)
                if node is not None and (node.text or "").strip():
                    brand = (node.text or "").strip()
                    break

            prices = {}
            updated = ""
            for prix in pdv.findall("prix"):
                fuel_name = str(prix.attrib.get("nom", "") or "").strip().lower()
                try:
                    value_f = float(str(prix.attrib.get("valeur")).replace(",", "."))
                except Exception:
                    continue
                updated = max(updated, str(prix.attrib.get("maj", "") or ""))
                if "gazole" in fuel_name or "diesel" in fuel_name:
                    prices["gazole"] = value_f
                elif fuel_name == "sp95":
                    prices["sp95"] = value_f
                elif "e10" in fuel_name:
                    prices["e10"] = value_f
                elif "sp98" in fuel_name:
                    prices["sp98"] = value_f
                elif "e85" in fuel_name:
                    prices["e85"] = value_f
                elif "gpl" in fuel_name:
                    prices["gplc"] = value_f

            rows.append({
                "lat": lat, "lon": lon, "adresse": adresse, "ville": ville, "cp": cp,
                "brand": brand, "gazole": prices.get("gazole"), "sp95": prices.get("sp95"),
                "e10": prices.get("e10"), "sp98": prices.get("sp98"),
                "e85": prices.get("e85"), "gplc": prices.get("gplc"), "updated": updated,
            })

        return pd.DataFrame(rows, columns=cols)
    except Exception:
        return pd.DataFrame(columns=cols)


def stations_near_route(df, return_row=None, current_position=None, margin_deg=0.11):
    pts = []
    for _, rr in df.iterrows():
        try:
            pts.append((float(rr.get("lat")), float(rr.get("lon"))))
        except Exception:
            pass
        geom = rr.get("route_geometry", [])
        if isinstance(geom, list) and geom:
            step = max(1, len(geom)//25)
            for p in geom[::step]:
                try:
                    pts.append((float(p[0]), float(p[1])))
                except Exception:
                    pass
    if return_row:
        geom = return_row.get("route_geometry", [])
        if isinstance(geom, list) and geom:
            step = max(1, len(geom)//25)
            for p in geom[::step]:
                try:
                    pts.append((float(p[0]), float(p[1])))
                except Exception:
                    pass
    if isinstance(current_position, dict):
        try:
            pts.append((float(current_position["latitude"]), float(current_position["longitude"])))
        except Exception:
            pass
    if not pts:
        return pd.DataFrame()

    all_stations = load_fuel_stations()
    if all_stations.empty:
        return all_stations

    lats = [p[0] for p in pts]
    lons = [p[1] for p in pts]
    return all_stations[
        all_stations["lat"].between(min(lats)-margin_deg, max(lats)+margin_deg)
        & all_stations["lon"].between(min(lons)-margin_deg, max(lons)+margin_deg)
    ].head(300).reset_index(drop=True)


def _html_escape(s):
    return (str(s or "")
            .replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
            .replace('"',"&quot;").replace("'","&#39;"))


def build_maplibre_html(df, return_row, start_address, start_geo, current_position=None,
                        show_radars=True, show_fuel=False, fuel_type="gazole", style_name="Liberty",
                        app_url=""):
    """Carte vectorielle moderne OpenFreeMap/MapLibre."""
    import json

    style_key = {
        "Liberty": "liberty",
        "Bright": "bright",
        "Positron": "positron",
        "Dark": "dark",
        "Fiord": "fiord",
    }.get(style_name, "liberty")
    style_url = f"https://tiles.openfreemap.org/styles/{style_key}"

    route_features = []
    route_labels = []
    stop_features = []
    bounds_pts = []

    # Base
    if start_geo.get("lat") and start_geo.get("lon"):
        base_lon = float(start_geo["lon"]); base_lat = float(start_geo["lat"])
        bounds_pts.append([base_lon, base_lat])
        stop_features.append({
            "type":"Feature",
            "properties":{"kind":"base","title":"Base","subtitle":start_address,"order":"⌂"},
            "geometry":{"type":"Point","coordinates":[base_lon,base_lat]},
        })

    for _, rr in df.iterrows():
        try:
            lat = float(rr.get("lat")); lon = float(rr.get("lon"))
        except Exception:
            continue
        bounds_pts.append([lon,lat])
        stop_features.append({
            "type":"Feature",
            "properties":{
                "kind":"stop",
                "order":str(rr.get("numero_rdv","")),
                "time":fmt_time(rr.get("heure_rdv")),
                "client":str(rr.get("nom_prospect","")),
                "address":str(rr.get("adresse_complete","")),
                "waze":str(rr.get("waze","#")),
                "groute":google_drive_to_link(rr.get("adresse_complete","")),
                "house":str(rr.get("street_view","#")),
                "terrain":f"{str(app_url).rstrip('/')}?terrain={quote_plus(str(rr.get('numero_rdv','')))}#terrain-{quote_plus(str(rr.get('numero_rdv','')))}" if app_url else "",
                "phone":str(rr.get("telephone_tel","")),
                "distance":str(rr.get("distance_depuis_precedent_km","")),
                "duration":fmt_duration(rr.get("temps_route_depuis_precedent_min","")),
                "depart":fmt_dt(rr.get("depart_conseille")),
                "ik":euro(rr.get("ik_montant_trajet",0)),
                "toll":(
                    euro(rr.get("peage_estime",0))
                    if bool(rr.get("peage_connu",False)) and float(rr.get("peage_estime",0) or 0) > 0
                    else ("Tarif indisponible" if bool(rr.get("peage_detecte",False)) else "")
                ),
            },
            "geometry":{"type":"Point","coordinates":[lon,lat]},
        })

        geom = rr.get("route_geometry", [])
        if isinstance(geom, list) and len(geom)>=2:
            coords = [[float(p[1]),float(p[0])] for p in geom if isinstance(p,(list,tuple)) and len(p)>=2]
            if len(coords)>=2:
                route_features.append({
                    "type":"Feature",
                    "properties":{"return":False},
                    "geometry":{"type":"LineString","coordinates":coords},
                })
                mid = coords[len(coords)//2]
                ik_txt = euro(rr.get("ik_montant_trajet", 0))
                
                delay_min = int(route_delay_status(rr, int(rr.name) if isinstance(rr.name, int) else 0))
                route_labels.append({
                    "coords":mid,
                    "line1":label,
                    "line2":f"Départ {fmt_dt(rr.get('depart_conseille'))}",
                    "delay_min":delay_min,
                    "alert":bool(delay_min > 0),
                })

    if return_row:
        geom = return_row.get("route_geometry", [])
        if isinstance(geom, list) and len(geom)>=2:
            coords = [[float(p[1]),float(p[0])] for p in geom if isinstance(p,(list,tuple)) and len(p)>=2]
            if len(coords)>=2:
                route_features.append({
                    "type":"Feature","properties":{"return":True},
                    "geometry":{"type":"LineString","coordinates":coords},
                })
                mid=coords[len(coords)//2]
                ik_ret = euro(return_row.get("ik_montant_trajet", 0))
                
                route_labels.append({"coords":mid,"line1":label,"line2":"Retour base","delay_min":0,"alert":False})

    gps = None
    if isinstance(current_position, dict):
        try:
            gps={"lat":float(current_position["latitude"]),"lon":float(current_position["longitude"])}
            bounds_pts.append([gps["lon"],gps["lat"]])
        except Exception:
            gps=None

    radar_features=[]
    if show_radars:
        rads=radars_near_route(df, return_row, current_position)
        for _, rd in rads.iterrows():
            radar_features.append({
                "type":"Feature",
                "properties":{
                    "type":str(rd.get("type","")),
                    "vma":str(rd.get("vma","")),
                    "id":str(rd.get("id","")),
                },
                "geometry":{"type":"Point","coordinates":[float(rd["lon"]),float(rd["lat"])]},
            })

    fuel_features=[]
    if show_fuel:
        stations = stations_near_route(df, return_row, current_position)
        for _, stn in stations.iterrows():
            try:
                price = None if pd.isna(stn.get(fuel_type)) else float(stn.get(fuel_type))
            except Exception:
                price = None
            def fv(key):
                try:
                    return None if pd.isna(stn.get(key)) else float(stn.get(key))
                except Exception:
                    return None
            fuel_features.append({
                "type":"Feature",
                "properties":{
                    "brand":str(stn.get("brand","") or ""),
                    "adresse":str(stn.get("adresse","") or ""),
                    "ville":str(stn.get("ville","") or ""),
                    "cp":str(stn.get("cp","") or ""),
                    "price":price,
                    "updated":str(stn.get("updated","") or ""),
                    "gazole":fv("gazole"),"sp95":fv("sp95"),"e10":fv("e10"),
                    "sp98":fv("sp98"),"e85":fv("e85"),"gplc":fv("gplc"),
                },
                "geometry":{"type":"Point","coordinates":[float(stn["lon"]),float(stn["lat"])]},
            })

    data = {
        "routes":{"type":"FeatureCollection","features":route_features},
        "stops":{"type":"FeatureCollection","features":stop_features},
        "radars":{"type":"FeatureCollection","features":radar_features},
        "fuel":{"type":"FeatureCollection","features":fuel_features},
        "labels":route_labels,
        "gps":gps,
        "bounds":bounds_pts,
        "style":style_url,
    }
    payload=json.dumps(data, ensure_ascii=False).replace("</","<\\/")

    return f"""<!doctype html>
<html>
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no"/>
<link href="https://unpkg.com/maplibre-gl@5/dist/maplibre-gl.css" rel="stylesheet"/>
<script src="https://unpkg.com/maplibre-gl@5/dist/maplibre-gl.js"></script>
<style>
html,body,#map{{margin:0;width:100%;height:100%;background:#070b14;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif}}
#map{{border-radius:22px;overflow:hidden}}
.maplibregl-ctrl-group{{border-radius:14px!important;overflow:hidden}}
.maplibregl-popup-content{{background:#0b1220!important;color:#fff!important;border:1px solid #263248;border-radius:18px!important;padding:14px!important;box-shadow:0 16px 40px rgba(0,0,0,.45)!important}}
.maplibregl-popup-tip{{border-top-color:#0b1220!important;border-bottom-color:#0b1220!important}}
.stop-marker{{min-width:94px;background:#05070b;color:#fff;border:1px solid #222b39;border-radius:14px;padding:7px 9px;box-shadow:0 6px 20px rgba(0,0,0,.42);font-weight:900;text-align:center;line-height:1.1}}
.stop-marker .rdv-head{{display:flex;align-items:center;justify-content:center;gap:6px}} .stop-marker .rdv-num{{width:23px;height:23px;border-radius:50%;display:inline-flex;align-items:center;justify-content:center;background:#00bfff;color:#03101a;font-size:12px;font-weight:950}} .stop-marker .t{{font-size:13px;color:#8beaff}} .stop-marker .n{{font-size:12px;margin-top:5px;max-width:155px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
.base-marker{{width:38px;height:38px;border-radius:50%;display:flex;align-items:center;justify-content:center;background:#0b1220;border:3px solid #62f6b7;color:#fff;font-size:20px;box-shadow:0 4px 18px rgba(98,246,183,.35)}}
.route-pill{{background:rgba(92,99,109,.97);color:#fff;border:1px solid #a7adb6;border-radius:12px;padding:7px 10px;box-shadow:0 5px 18px rgba(0,0,0,.24);font-weight:850;font-size:12px;line-height:1.3;white-space:nowrap}}
.route-pill .sub{{color:#80e8ff;font-weight:900}}
.route-pill.delay-alert{{
  background:rgba(210,34,51,.97)!important;
  border:2px solid #fff!important;
  box-shadow:0 0 0 0 rgba(255,49,70,.65),0 6px 22px rgba(255,0,34,.38)!important;
  animation:delayPulse 1.05s infinite alternate;
}}
.route-pill .delay-line{{
  color:#fff;
  font-weight:950;
  font-size:12px;
  margin-top:3px;
}}
@keyframes delayPulse{{
  0%{{transform:scale(1);box-shadow:0 0 0 0 rgba(255,49,70,.55),0 6px 22px rgba(255,0,34,.30)}}
  100%{{transform:scale(1.035);box-shadow:0 0 0 8px rgba(255,49,70,0),0 8px 28px rgba(255,0,34,.52)}}
}}

.radar-marker{{width:34px;height:34px;border-radius:50%;background:#ff293d;border:3px solid #fff;color:#fff;display:flex;align-items:center;justify-content:center;font-size:16px;font-weight:950;box-shadow:0 4px 15px rgba(255,41,61,.45)}}
.fuel-marker{{min-width:40px;height:38px;border-radius:13px;background:#0a8f63;border:2px solid #fff;color:#fff;display:flex;align-items:center;justify-content:center;font-size:17px;font-weight:950;box-shadow:0 4px 15px rgba(10,143,99,.42);padding:0 5px}}
.fuel-price{{font-size:10px;margin-left:3px;font-weight:950}}

.gps-dot{{width:22px;height:22px;border-radius:50%;background:#178bff;border:4px solid #fff;box-shadow:0 0 0 10px rgba(23,139,255,.22),0 0 25px rgba(23,139,255,.55)}}
.popup-title{{font-size:18px;font-weight:950;margin-bottom:4px}} .popup-client{{font-size:15px;font-weight:850;color:#7ee7ff;margin-bottom:6px}}
.popup-addr{{font-size:12px;color:#cbd5e1;margin-bottom:10px}}
.popup-route{{font-size:12px;color:#dbeafe;margin:8px 0}}
.pbtn{{display:inline-block;text-decoration:none;color:#fff!important;background:#142033;border:1px solid #31425f;border-radius:10px;padding:8px 10px;margin:3px 2px;font-weight:850;font-size:12px}}
.pbtn.primary{{background:#087de8;border-color:#21bfff}} .pbtn.house{{background:#5b2fc6;border-color:#8b5cf6}}
</style>
</head>
<body>
<div id="map"></div>
<script>
const DATA={payload};
const map=new maplibregl.Map({{
  container:"map", style:DATA.style, center:[2.35,48.85], zoom:7,
  attributionControl:true
}});
map.addControl(new maplibregl.NavigationControl({{showCompass:true}}),"top-right");

function esc(v){{return String(v??"").replace(/[&<>"']/g,m=>({{"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}}[m]));}}
map.on("load",()=>{{
  map.addSource("routes",{{type:"geojson",data:DATA.routes}});
  map.addLayer({{id:"route-halo",type:"line",source:"routes",paint:{{"line-color":"#06101d","line-width":10,"line-opacity":.78}},layout:{{"line-cap":"round","line-join":"round"}}}});
  map.addLayer({{id:"routes",type:"line",source:"routes",paint:{{"line-color":["case",["get","return"],"#7d8da7","#00bfff"],"line-width":["case",["get","return"],5,6],"line-opacity":.98,"line-dasharray":["case",["get","return"],["literal",[2,1]],["literal",[1,0]]]}},layout:{{"line-cap":"round","line-join":"round"}}}});

  // Labels trajet
  DATA.labels.forEach(l=>{{
    const el=document.createElement("div");
    el.className="route-pill" + (l.alert ? " delay-alert" : "");
    const alertLine=l.alert?`<div class="delay-line">⚠ RETARD PROBABLE · +${{esc(l.delay_min)}} min</div>`:"";
    el.innerHTML=esc(l.line1)+`<br><span class="sub">${{esc(l.line2)}}</span>`+alertLine;
    new maplibregl.Marker({{element:el,anchor:"center"}}).setLngLat(l.coords).addTo(map);
  }});

  // Stops
  DATA.stops.features.forEach(f=>{{
    const p=f.properties, c=f.geometry.coordinates;
    const el=document.createElement("div");
    if(p.kind==="base"){{ el.className="base-marker"; el.innerHTML="⌂"; }}
    else{{ el.className="stop-marker"; el.innerHTML=`<div class="rdv-head"><span class="rdv-num">${{esc(p.order)}}</span><span class="t">${{esc(p.time)}}</span></div><div class="n">${{esc(p.client)}}</div>`; }}
    const marker=new maplibregl.Marker({{element:el,anchor:"bottom"}}).setLngLat(c).addTo(map);
    if(p.kind==="stop"){{
      const phone=p.phone?`<a class="pbtn" href="tel:${{esc(p.phone)}}">📞 Appeler</a>`:"";
      const toll=p.toll?` · 🛣️ ${{esc(p.toll)}}`:"";
      const html=`<div class="popup-title"><span style="display:inline-flex;width:27px;height:27px;border-radius:50%;align-items:center;justify-content:center;background:#00bfff;color:#03101a;font-size:13px;margin-right:7px">${{esc(p.order)}}</span>${{esc(p.time)}}</div>
      <div class="popup-client">${{esc(p.client)}}</div><div class="popup-addr">${{esc(p.address)}}</div>
      <div class="popup-route">🚗 ${{esc(p.duration)}} · ${{esc(p.distance)}} km · IK ${{esc(p.ik)}}${{toll}}<br>⏰ Départ ${{esc(p.depart)}}</div>
      <a class="pbtn primary" target="_blank" href="${{esc(p.waze)}}">🚗 Waze</a>
      <a class="pbtn" target="_blank" href="${{esc(p.groute)}}">🗺️ Trajet Google</a>
      <a class="pbtn house" target="_blank" href="${{esc(p.house)}}">🏠 Voir maison</a>
      ${{p.terrain ? `<a class="pbtn" target="_top" href="${{esc(p.terrain)}}">🧠 Mode terrain</a>` : ""}}${{phone}}`;
      marker.setPopup(new maplibregl.Popup({{offset:28,maxWidth:"340px"}}).setHTML(html));
    }} else {{
      marker.setPopup(new maplibregl.Popup({{offset:22}}).setHTML(`<b>⌂ Base</b><br>${{esc(p.subtitle)}}`));
    }}
  }});

  // GPS
  if(DATA.gps){{
    const el=document.createElement("div"); el.className="gps-dot";
    new maplibregl.Marker({{element:el}}).setLngLat([DATA.gps.lon,DATA.gps.lat]).setPopup(new maplibregl.Popup().setHTML("<b>📍 Ma position</b>")).addTo(map);
  }}

  // Radars fixes publics
  DATA.radars.features.forEach(f=>{{
    const p=f.properties,c=f.geometry.coordinates,el=document.createElement("div");
    el.className="radar-marker"; el.innerHTML="📷";
    let details=`<div class="popup-title">📷 Radar fixe</div>`;
    if(p.type) details+=`<div class="popup-client">${{esc(p.type)}}</div>`;
    if(p.vma && p.vma!=="nan") details+=`<div class="popup-route">Limitation : <b>${{esc(p.vma)}} km/h</b></div>`;
    new maplibregl.Marker({{element:el,anchor:"center"}}).setLngLat(c).setPopup(new maplibregl.Popup({{offset:20}}).setHTML(details)).addTo(map);
  }});

  DATA.fuel.features.forEach(f=>{{
    const p=f.properties,c=f.geometry.coordinates,el=document.createElement("div");
    el.className="fuel-marker";
    const price=(p.price!==null&&p.price!==undefined)?Number(p.price).toFixed(3):"";
    el.innerHTML=price?`⛽<span class="fuel-price">${{price}}€</span>`:"⛽";
    let title=p.brand?esc(p.brand):"Station-service";
    let details=`<div class="popup-title">⛽ ${{title}}</div>`;
    const addr=[p.adresse,p.cp,p.ville].filter(Boolean).join(" ");
    if(addr) details+=`<div class="popup-addr">${{esc(addr)}}</div>`;
    const fuels=[["Gazole",p.gazole],["SP95",p.sp95],["E10",p.e10],["SP98",p.sp98],["E85",p.e85],["GPLc",p.gplc]].filter(x=>x[1]!==null&&x[1]!==undefined);
    if(fuels.length) details+=`<div class="popup-route">${{fuels.map(x=>`${{x[0]}} : <b>${{Number(x[1]).toFixed(3)}} €/L</b>`).join("<br>")}}</div>`;
    if(p.updated) details+=`<div style="font-size:10px;color:#94a3b8">Mise à jour : ${{esc(p.updated)}}</div>`;
    new maplibregl.Marker({{element:el,anchor:"center"}}).setLngLat(c).setPopup(new maplibregl.Popup({{offset:20,maxWidth:"300px"}}).setHTML(details)).addTo(map);
  }});

  if(DATA.bounds && DATA.bounds.length){{
    const b=new maplibregl.LngLatBounds();
    DATA.bounds.forEach(p=>b.extend(p));
    map.fitBounds(b,{{padding:70,maxZoom:12,duration:0}});
  }}
}});
</script>
</body></html>"""


def make_map(df, return_row, start_address, start_geo, interactive=True, current_position=None):
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

    m = folium.Map(location=center, zoom_start=11, tiles="OpenStreetMap",
                   dragging=interactive, scrollWheelZoom=interactive,
                   touchZoom=interactive, doubleClickZoom=interactive, zoom_control=True)

    points = []
    if start_geo.get("lat") and start_geo.get("lon"):
        folium.Marker([start_geo["lat"], start_geo["lon"]],
                      tooltip="Départ / retour",
                      popup=folium.Popup(f"<b>🏠 Base</b><br>{start_address}", max_width=320),
                      icon=folium.Icon(color="green", icon="home")).add_to(m)
        points.append([start_geo["lat"], start_geo["lon"]])

    # Position GPS actuelle de l'iPhone.
    if isinstance(current_position, dict):
        gps_lat = current_position.get("latitude")
        gps_lon = current_position.get("longitude")
        if gps_lat is not None and gps_lon is not None:
            try:
                gps_lat = float(gps_lat)
                gps_lon = float(gps_lon)
                folium.CircleMarker(
                    [gps_lat, gps_lon],
                    radius=10,
                    tooltip="📍 Vous êtes ici",
                    popup=folium.Popup("<b>📍 Ma position actuelle</b>", max_width=220),
                    color="#ffffff",
                    weight=3,
                    fill=True,
                    fill_color="#2563eb",
                    fill_opacity=1.0,
                ).add_to(m)
                folium.CircleMarker(
                    [gps_lat, gps_lon],
                    radius=18,
                    color="#2563eb",
                    weight=2,
                    fill=False,
                    opacity=0.35,
                ).add_to(m)
                points.append([gps_lat, gps_lon])
            except Exception:
                pass

    for _, r in df.iterrows():
        if not r.get("lat") or not r.get("lon"):
            continue

        time_label = fmt_time(r.get("heure_rdv"))
        client = str(r.get("nom_prospect",""))
        address = str(r.get("adresse_complete",""))
        tel = str(r.get("telephone_tel",""))
        waze = str(r.get("waze","#"))
        house = str(r.get("street_view","#"))
        phone_html = f"<a href='tel:{tel}' style='display:inline-block;padding:7px 9px;margin:4px 2px;background:#111827;color:#fff;border-radius:8px;text-decoration:none;'>📞 Appeler</a>" if tel else ""
        popup_html = f"""
        <div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;min-width:230px;">
          <div style="font-size:18px;font-weight:900;">#{r.get('numero_rdv','')} · {time_label}</div>
          <div style="font-size:16px;font-weight:800;margin-top:3px;">{client}</div>
          <div style="font-size:13px;margin:6px 0 9px 0;">{address}</div>
          <a href="{waze}" target="_blank" style="display:inline-block;padding:7px 9px;margin:4px 2px;background:#2563eb;color:#fff;border-radius:8px;text-decoration:none;">🚗 Waze</a>
          <a href="{house}" target="_blank" style="display:inline-block;padding:7px 9px;margin:4px 2px;background:#7c3aed;color:#fff;border-radius:8px;text-decoration:none;">🏠 Voir maison</a>
          {phone_html}
          <div style="font-size:12px;margin-top:7px;">
            🚗 {r.get('distance_depuis_precedent_km','')} km · {fmt_duration(r.get('temps_route_depuis_precedent_min',''))}<br>
            ⏰ Départ conseillé : <b>{fmt_dt(r.get('depart_conseille'))}</b>
          </div>
        </div>
        """

        folium.Marker([r["lat"], r["lon"]],
                      tooltip=f"#{r.get('numero_rdv','')} {client} · {time_label}",
                      popup=folium.Popup(popup_html, max_width=340),
                      icon=folium.Icon(color="blue", icon="user")).add_to(m)

        marker_html = f"""<div style='font-size:16px;line-height:19px;font-weight:900;background:#ffb347;color:#111827;border:2px solid #111827;border-radius:9px;padding:5px 8px;white-space:nowrap;box-shadow:0 3px 10px rgba(0,0,0,.45);'>#{r.get('numero_rdv','')} · {time_label}<br>{client}</div>"""
        folium.map.Marker([r["lat"], r["lon"]], icon=folium.DivIcon(html=marker_html)).add_to(m)
        points.append([r["lat"], r["lon"]])

        geom = r.get("route_geometry", [])
        if isinstance(geom, list) and len(geom) >= 2:
            folium.PolyLine(geom, weight=5, opacity=0.9, color="red").add_to(m)
            mid = geom[len(geom)//2]
            toll_txt = f" · 🛣️ {euro(r.get('peage_estime',0))}" if float(r.get('peage_estime',0) or 0)>0 else ""
            route_label = f"""<div style='font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;font-size:12px;line-height:15px;font-weight:900;background:#fff;color:#111827;border:2px solid #111827;border-radius:9px;padding:6px 8px;white-space:nowrap;box-shadow:0 2px 8px rgba(0,0,0,.35);'>🚗 {fmt_duration(r.get('temps_route_depuis_precedent_min',''))} · {r.get('distance_depuis_precedent_km','')} km{toll_txt}<br>⏰ Départ {fmt_dt(r.get('depart_conseille'))}</div>"""
            folium.map.Marker(mid, icon=folium.DivIcon(html=route_label)).add_to(m)

    if return_row and return_row.get("lat") and return_row.get("lon"):
        points.append([return_row["lat"], return_row["lon"]])
        geom = return_row.get("route_geometry", [])
        if isinstance(geom, list) and len(geom) >= 2:
            folium.PolyLine(geom, weight=5, opacity=0.85, color="red", dash_array="8,6").add_to(m)
            mid = geom[len(geom)//2]
            return_toll_txt = f" · 🛣️ {euro(return_row.get('peage_estime',0))}" if float(return_row.get('peage_estime',0) or 0) > 0 else ""
            route_label = f"""<div style='font-size:12px;line-height:15px;font-weight:900;background:#ffffff;color:#111827;border:2px solid #111827;border-radius:9px;padding:6px 8px;white-space:nowrap;box-shadow:0 2px 8px rgba(0,0,0,.35);'>🏠 Retour · {fmt_duration(return_row.get('temps_route_depuis_precedent_min',''))} · {return_row.get('distance_depuis_precedent_km','')} km{return_toll_txt}</div>"""
            folium.map.Marker(mid, icon=folium.DivIcon(html=route_label)).add_to(m)

    try:
        if len(points) >= 2:
            m.fit_bounds(points, padding=(35,35))
        elif len(points) == 1:
            m.location = points[0]
            m.zoom_start = 13
    except Exception:
        pass
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
            Paragraph(f"{r.get('distance_depuis_precedent_km','')} km<br/>{fmt_duration(r.get('temps_route_depuis_precedent_min',''))}<br/>IK estimée : {euro(r.get('ik_montant_trajet', 0))}<br/>{r.get('note_trafic','')}", small),
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
        info = f"<b>Adresse :</b> {r.get('adresse_complete','')}<br/><b>Téléphone :</b> {r.get('telephone','')}<br/><b>Téléprospecteur :</b> {r.get('teleprospecteur','')}<br/><b>Départ conseillé :</b> {fmt_dt(r.get('depart_conseille'))}<br/><b>Trajet :</b> {r.get('distance_depuis_precedent_km','')} km · {fmt_duration(r.get('temps_route_depuis_precedent_min',''))}<br/><b>IK estimée trajet :</b> {euro(r.get('ik_montant_trajet', 0))}<br/><a href='{r.get('waze','#')}'>Ouvrir Waze</a> · <a href='{r.get('google_maps','#')}'>Google Maps</a> · <a href='{r.get('street_view','#')}'>Voir maison</a>"
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


def load_app_settings():
    for path in [APP_STATE_PATH] + APP_STATE_FALLBACK_PATHS:
        try:
            if path.exists():
                return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def save_app_settings(data):
    try:
        current = load_app_settings()
        current.update(data or {})
        APP_STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
        APP_STATE_PATH.write_text(json.dumps(current, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def crm_key(row):
    d = row.get("date_rdv", "")
    d_txt = d.isoformat() if isinstance(d, date) else str(d or "")
    return "|".join([d_txt, str(row.get("heure_rdv", "")), str(row.get("nom_prospect", "")), str(row.get("telephone_tel", "")), str(row.get("adresse_complete", ""))])


def crm_columns():
    return ["key", "date_rdv", "heure_rdv", "client", "telephone", "telephone_tel", "adresse", "teleprospecteur", "fournisseur", "commercial", "statut", "commentaire", "details_crm", "analyse_ia", "date_rappel", "heure_rappel", "rappel_traite", "traite_at", "created_at", "updated_at"]


def empty_crm_df():
    return pd.DataFrame(columns=crm_columns())


def load_crm_history():
    for path in [CRM_HISTORY_PATH] + CRM_HISTORY_FALLBACK_PATHS:
        try:
            if path.exists():
                df = pd.read_csv(path, sep=";", dtype=str).fillna("")
                for c in crm_columns():
                    if c not in df.columns:
                        df[c] = ""
                return df[crm_columns()]
        except Exception:
            pass
    return empty_crm_df()


def save_crm_record(key, row, statut, commentaire, rappel_date=None, rappel_time=None, details_crm="", analyse_ia=""):
    hist = load_crm_history()
    now = datetime.now().strftime("%d/%m/%Y %H:%M")
    d = row.get("date_rdv", "")
    d_txt = d.strftime("%d/%m/%Y") if isinstance(d, date) else str(d or "")
    t_txt = fmt_time(row.get("heure_rdv"))
    record = {
        "key": key,
        "date_rdv": d_txt,
        "heure_rdv": t_txt,
        "client": row.get("nom_prospect", ""),
        "telephone": row.get("telephone", ""),
        "telephone_tel": row.get("telephone_tel", ""),
        "adresse": row.get("adresse_complete", ""),
        "teleprospecteur": row.get("teleprospecteur", ""),
        "fournisseur": row.get("fournisseur", ""),
        "commercial": row.get("commercial", ""),
        "statut": statut or "",
        "commentaire": commentaire or "",
        "details_crm": details_crm or "",
        "analyse_ia": analyse_ia or "",
        "date_rappel": rappel_date.strftime("%d/%m/%Y") if isinstance(rappel_date, date) else str(rappel_date or ""),
        "heure_rappel": fmt_time(rappel_time) if isinstance(rappel_time, dtime) else str(rappel_time or ""),
        "created_at": now,
        "updated_at": now,
    }
    # Par défaut, un rappel nouvellement saisi est non traité.
    # Si le RDV existait déjà et avait été marqué traité, on conserve l'état traité tant que la date/heure de rappel ne change pas.
    record["rappel_traite"] = "non"
    record["traite_at"] = ""
    if hist.empty or key not in hist.get("key", pd.Series(dtype=str)).astype(str).tolist():
        hist = pd.concat([hist, pd.DataFrame([record])], ignore_index=True)
    else:
        idx = hist.index[hist["key"].astype(str) == key]
        created = hist.loc[idx[0], "created_at"] if len(idx) else now
        old_date = str(hist.loc[idx[0], "date_rappel"]) if len(idx) and "date_rappel" in hist.columns else ""
        old_time = str(hist.loc[idx[0], "heure_rappel"]) if len(idx) and "heure_rappel" in hist.columns else ""
        old_done = str(hist.loc[idx[0], "rappel_traite"]).lower() if len(idx) and "rappel_traite" in hist.columns else "non"
        old_done_at = str(hist.loc[idx[0], "traite_at"]) if len(idx) and "traite_at" in hist.columns else ""
        old_details = str(hist.loc[idx[0], "details_crm"]) if len(idx) and "details_crm" in hist.columns else ""
        old_analyse = str(hist.loc[idx[0], "analyse_ia"]) if len(idx) and "analyse_ia" in hist.columns else ""
        if not record.get("details_crm") and old_details:
            record["details_crm"] = old_details
        if not record.get("analyse_ia") and old_analyse:
            record["analyse_ia"] = old_analyse
        record["created_at"] = created
        if old_date == record.get("date_rappel", "") and old_time == record.get("heure_rappel", ""):
            record["rappel_traite"] = "oui" if old_done in ["oui", "true", "1", "yes"] else "non"
            record["traite_at"] = old_done_at if record["rappel_traite"] == "oui" else ""
        for k, v in record.items():
            hist.loc[idx, k] = v
    try:
        for c in crm_columns():
            if c not in hist.columns:
                hist[c] = ""
        CRM_HISTORY_PATH.parent.mkdir(parents=True, exist_ok=True)
        hist[crm_columns()].to_csv(CRM_HISTORY_PATH, index=False, sep=";", encoding="utf-8-sig")
    except Exception:
        pass


def mark_reminder_treated(key, treated=True):
    # Marque un rappel comme traité / non traité dans l'historique CRM.
    hist = load_crm_history()
    if hist.empty:
        return
    if "rappel_traite" not in hist.columns:
        hist["rappel_traite"] = "non"
    if "traite_at" not in hist.columns:
        hist["traite_at"] = ""
    idx = hist.index[hist["key"].astype(str) == str(key)]
    if len(idx):
        hist.loc[idx, "rappel_traite"] = "oui" if treated else "non"
        hist.loc[idx, "traite_at"] = datetime.now().strftime("%d/%m/%Y %H:%M") if treated else ""
        try:
            for c in crm_columns():
                if c not in hist.columns:
                    hist[c] = ""
            CRM_HISTORY_PATH.parent.mkdir(parents=True, exist_ok=True)
            hist[crm_columns()].to_csv(CRM_HISTORY_PATH, index=False, sep=";", encoding="utf-8-sig")
        except Exception:
            pass


def is_reminder_treated(row):
    return str(row.get("rappel_traite", "")).strip().lower() in ["oui", "true", "1", "yes", "traité", "traite"]


def reminder_section_label(dt, now_dt):
    if dt is None:
        return "future"
    if dt.date() < now_dt.date():
        return "late"
    if dt.date() == now_dt.date():
        return "today"
    return "future"


def reminder_datetime(row):
    try:
        d = pd.to_datetime(row.get("date_rappel", ""), dayfirst=True, errors="coerce")
        if pd.isna(d):
            return None
        t = parse_time(row.get("heure_rappel", "")) or dtime(9,0)
        return datetime.combine(d.date(), t)
    except Exception:
        return None

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
    # Priorité au dernier export CRM enrichi local, si le robot vient de tourner sur la Surface.
    latest_crm = latest_local_crm_export()
    if latest_crm is not None:
        try:
            data = latest_crm.read_bytes()
            LAST_UPLOAD_PATH.write_bytes(data)
            st.session_state["last_upload_bytes"] = data
            st.session_state["last_upload_name"] = latest_crm.name
            return io.BytesIO(data)
        except Exception:
            pass
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


def calc_ik_amount(km, cv=7, electric=False, bareme=None):
    """Calcule l'indemnité kilométrique selon le barème 2026 voiture.
    Attention : le barème est annuel. Pour un export mensuel, l'app calcule sur les km de la période sélectionnée.
    Une régularisation annuelle peut être faite par le comptable si nécessaire.
    """
    d = max(0.0, float(km or 0))
    key = ik_cv_key(cv)
    brackets = (bareme or IK_BAREME_2026)[key]
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


def add_ik_amounts_to_register(register_df, total_amount=0.0):
    """Répartit le montant total IK sur chaque trajet au prorata des kilomètres.
    C'est une ventilation pratique : le barème officiel reste calculé sur le total de la période.
    """
    out = register_df.copy()
    if out.empty:
        out["Montant IK"] = []
        return out
    km_series = pd.to_numeric(out.get("Km", pd.Series(dtype=float)), errors="coerce").fillna(0)
    total_km = float(km_series.sum())
    if total_km <= 0:
        out["Montant IK"] = 0.0
    else:
        out["Montant IK"] = (km_series / total_km * float(total_amount or 0)).round(2)
    return out


def calc_ik_amount_flexible(km, cv=7, electric=False, bareme=None, mode="auto", manual_rate=0.47):
    d = max(0.0, float(km or 0))
    if mode == "manuel":
        rate = max(0.0, float(manual_rate or 0))
        return round(d * rate, 2), f"Forfait interne : {rate:.3f} €/km"
    return calc_ik_amount(d, cv=cv, electric=electric, bareme=bareme)


def add_ik_amounts_to_route(route_df, return_row, cv=7, electric=False, bareme=None, include_return=True, mode="auto", manual_rate=0.47):
    """Ajoute une estimation IK par trajet dans la tournée affichée.
    En mode manuel, chaque trajet utilise directement le forfait interne choisi.
    En mode auto, le montant total barème est ventilé au prorata des kilomètres.
    """
    out = route_df.copy()
    reg = build_ik_register(out, return_row, include_return=include_return)
    total_km = float(pd.to_numeric(reg.get("Km", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) if not reg.empty else 0.0
    total_amount, formula = calc_ik_amount_flexible(total_km, cv=cv, electric=electric, bareme=bareme, mode=mode, manual_rate=manual_rate)
    km_route = pd.to_numeric(out.get("distance_depuis_precedent_km", pd.Series(dtype=float)), errors="coerce").fillna(0)
    if mode == "manuel":
        out["ik_montant_trajet"] = (km_route * float(manual_rate or 0)).round(2)
    else:
        out["ik_montant_trajet"] = (km_route / total_km * total_amount).round(2) if total_km > 0 else 0.0
    ret_amount = 0.0
    if include_return and return_row:
        ret_km = to_float(return_row.get("distance_depuis_precedent_km"))
        ret_amount = round(ret_km * float(manual_rate or 0), 2) if mode == "manuel" else (round(ret_km / total_km * total_amount, 2) if total_km > 0 else 0.0)
    return out, ret_amount, total_amount, formula


def custom_bareme_from_inputs(prefix="ik"):
    """Écran de paramétrage complet du barème IK.
    Les valeurs par défaut correspondent au barème intégré. L'utilisateur peut les modifier si le barème évolue.
    """
    st.markdown("**Barème paramétrable** — modifie uniquement si ton comptable ou l'administration publie un nouveau barème.")
    rows = []
    for cv_key in [3, 4, 5, 6, 7]:
        b = IK_BAREME_2026[cv_key]
        c1, c2, c3, c4, c5 = st.columns([0.8, 1.4, 1.4, 1.4, 1.4])
        with c1:
            st.markdown(f"**{cv_key} CV{'+' if cv_key == 7 else ''}**")
        with c2:
            r1 = st.number_input(f"≤ 5 000 km taux {cv_key}CV", value=float(b[0][1]), format="%.3f", step=0.001, key=f"{prefix}_{cv_key}_r1")
        with c3:
            r2 = st.number_input(f"5 001–20 000 km taux {cv_key}CV", value=float(b[1][1]), format="%.3f", step=0.001, key=f"{prefix}_{cv_key}_r2")
        with c4:
            f2 = st.number_input(f"forfait {cv_key}CV", value=float(b[1][2]), step=1.0, key=f"{prefix}_{cv_key}_f2")
        with c5:
            r3 = st.number_input(f"> 20 000 km taux {cv_key}CV", value=float(b[2][1]), format="%.3f", step=0.001, key=f"{prefix}_{cv_key}_r3")
        rows.append((cv_key, ((5000, r1, 0), (20000, r2, f2), (float("inf"), r3, 0))))
    return dict(rows)


def history_file_name(register_df):
    if register_df is None or register_df.empty:
        return None
    dates = [str(x) for x in register_df.get("Date", []) if str(x).strip()]
    date_key = dates[0].replace("/", "-") if dates else datetime.now().strftime("%d-%m-%Y")
    km_key = int(round(float(pd.to_numeric(register_df.get("Km", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) * 10))
    return f"IK_{date_key}_{len(register_df)}_{km_key}.csv"


def save_ik_history(register_df):
    try:
        IK_HISTORY_DIR.mkdir(parents=True, exist_ok=True)
        name = history_file_name(register_df)
        if not name:
            return None
        path = IK_HISTORY_DIR / name
        register_df.to_csv(path, index=False, sep=";", encoding="utf-8-sig")
        return path
    except Exception:
        return None


def load_ik_history():
    frames = []
    try:
        if not IK_HISTORY_DIR.exists():
            return pd.DataFrame()
        for path in sorted(IK_HISTORY_DIR.glob("IK_*.csv")):
            try:
                df = pd.read_csv(path, sep=";")
                df["Source"] = path.name
                frames.append(df)
            except Exception:
                pass
    except Exception:
        pass
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def filter_register_by_period(register_df, month, year):
    if register_df is None or register_df.empty or "Date" not in register_df.columns:
        return pd.DataFrame()
    out = register_df.copy()
    parsed = pd.to_datetime(out["Date"], dayfirst=True, errors="coerce")
    return out[(parsed.dt.month == int(month)) & (parsed.dt.year == int(year))].copy()


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
    story.append(Paragraph("Document généré depuis l’application Routage PRO V19", small))
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
    table_data = [["Date", "Objet", "Arrivée", "Km", "Montant", "Temps", "Justificatif"]]
    for _, r in register_df.iterrows():
        table_data.append([
            str(r.get('Date','')),
            Paragraph(str(r.get('Objet','')), small),
            Paragraph(str(r.get('Arrivée','')), small),
            f"{to_float(r.get('Km')):.1f}",
            euro(r.get('Montant IK', 0)),
            str(r.get('Temps','')),
            Paragraph(str(r.get('Justificatif','')), small),
        ])
    tbl = Table(table_data, colWidths=[1.5*cm, 3.2*cm, 5.3*cm, 1.2*cm, 1.5*cm, 1.5*cm, 2.2*cm], repeatRows=1)
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

settings = load_app_settings()
with st.sidebar:
    st.header("Réglages")
    start_address = st.text_input("Adresse de départ / retour", value=settings.get("start_address", DEFAULT_START))
    safety_min = st.number_input("Marge sécurité avant RDV", min_value=0, max_value=60, value=int(settings.get("safety_min", 15)), step=5)
    visit_min = st.number_input("Durée moyenne d'un RDV", min_value=15, max_value=240, value=int(settings.get("visit_min", 150)), step=15)
    # Clé Google stockée côté serveur dans Streamlit Secrets, jamais affichée dans l'app.
    try:
        google_key = str(st.secrets.get("GOOGLE_MAPS_API_KEY", "")).strip()
    except Exception:
        google_key = ""

    google_routes_ready = bool(google_key)
    use_google = st.checkbox(
        "Utiliser Google Routes avec trafic",
        value=google_routes_ready,
        disabled=not google_routes_ready,
        help="La clé est lue automatiquement depuis Streamlit Secrets."
    )
    if google_routes_ready:
        st.caption("🟢 Trafic Google actif")
        with st.expander("🛠️ Diagnostic", expanded=False):
            if st.button("🧪 Tester Google Routes", use_container_width=True):
                with st.spinner("Test de Google Routes en cours…"):
                    diagnostic = google_routes_diagnostic(google_key)
                if diagnostic.get("ok"):
                    st.success(f"✅ Test Google Routes : {diagnostic.get('message')}")
                else:
                    status = diagnostic.get("status")
                    prefix = f"Erreur HTTP {status}" if status else "Erreur"
                    st.error(f"🔴 {prefix} : {diagnostic.get('message')}")
                    st.caption("La clé API n'est jamais affichée par ce diagnostic.")
    else:
        st.warning("🟠 Google Routes API non configurée : temps sans trafic réel")
    st.divider()
    st.subheader("🧭 Itinéraire")
    route_pref_label=st.selectbox("Préférence",["⚡ Recommandé","💶 Sans péage","📏 Plus court","🛣️ Sans autoroute"], index=0)
    route_pref={"⚡ Recommandé":"recommended","💶 Sans péage":"no_tolls","📏 Plus court":"shortest","🛣️ Sans autoroute":"no_highways"}[route_pref_label]
    include_tolls=st.checkbox("Afficher le coût estimé des péages",value=True)
    with st.expander("➕ Ajouter une étape",expanded=False):
        st.caption("Exemple : ouverture de chantier avant le premier RDV.")
        extra_type=st.selectbox("Type",["🏗️ Ouverture de chantier","🏢 Bureau / dépôt","☕ Pause","📍 Autre"],key="extra_type")
        extra_nom=st.text_input("Nom / client",key="extra_nom")
        extra_adresse=st.text_input("Adresse complète",key="extra_adresse")
        cdate,ctime=st.columns(2)
        with cdate: extra_date=st.date_input("Date",value=date.today()+timedelta(days=1),key="extra_date")
        with ctime: extra_heure=st.time_input("Heure",value=dtime(8,30),key="extra_heure")
        extra_duree=st.number_input("Durée sur place (min)",min_value=0,max_value=240,value=30,step=5,key="extra_duree")
        extra_tel=st.text_input("Téléphone (optionnel)",key="extra_tel")
        extra_note=st.text_area("Note (optionnel)",key="extra_note",height=60)
        if st.button("➕ Ajouter à la tournée",use_container_width=True):
            if not str(extra_adresse).strip(): st.error("Renseigne l'adresse de l'étape.")
            else:
                st.session_state.setdefault("extra_steps",[]); st.session_state["extra_steps"].append({"type":extra_type,"nom":extra_nom or extra_type,"adresse":extra_adresse.strip(),"date":extra_date,"heure":extra_heure,"duree":int(extra_duree),"telephone":extra_tel,"note":extra_note}); st.rerun()
        for j,s in enumerate(st.session_state.get("extra_steps",[])):
            cc1,cc2=st.columns([4,1]); cc1.caption(f"{fmt_time(s.get('heure'))} · {s.get('nom')} · {s.get('adresse')}")
            if cc2.button("✕",key=f"del_extra_{j}"): st.session_state["extra_steps"].pop(j); st.rerun()

    uploaded = st.file_uploader("Importer ton fichier Excel", type=["xlsx", "xls"])
    saved = st.file_uploader("Ou charger un récap CSV sauvegardé", type=["csv"], key="saved_csv")
    auto_reload = st.checkbox("Recharger automatiquement le dernier Excel de la journée", value=True)
    st.divider()
    st.subheader("💰 Réglages IK rapides")
    sidebar_cv = st.selectbox("Puissance fiscale IK", options=[3,4,5,6,7], index=[3,4,5,6,7].index(int(settings.get("sidebar_cv", 7))), help="7 = 7 CV et plus", key="sidebar_cv")
    sidebar_electric = st.checkbox("Véhicule 100% électrique IK (+20%)", value=bool(settings.get("sidebar_electric", False)), key="sidebar_electric")
    sidebar_include_return = st.checkbox("Inclure retour base dans IK", value=bool(settings.get("sidebar_return_ik", True)), key="sidebar_return_ik")
    ik_mode_label = st.radio("Mode de calcul IK", ["Barème officiel", "Forfait interne €/km"], index=1 if settings.get("ik_mode", "manuel") == "manuel" else 0)
    sidebar_ik_mode = "manuel" if ik_mode_label.startswith("Forfait") else "auto"
    sidebar_manual_rate = st.number_input("Forfait interne €/km", min_value=0.0, max_value=2.0, value=float(settings.get("manual_rate", 0.47)), step=0.01, format="%.3f", disabled=(sidebar_ik_mode != "manuel"))
    save_app_settings({
        "start_address": start_address, "safety_min": int(safety_min), "visit_min": int(visit_min),
        "use_google": bool(use_google), "sidebar_cv": int(sidebar_cv),
        "sidebar_electric": bool(sidebar_electric), "sidebar_return_ik": bool(sidebar_include_return),
        "ik_mode": sidebar_ik_mode, "manual_rate": float(sidebar_manual_rate),
    })
    st.info("V28.5 — 28/07/2026 : Google Routes API avec trafic réel/prédictif + import CRM enrichi + rappels + IA.")

source_file = None
source_label = ""
local_crm_file = latest_local_crm_export()

if uploaded:
    source_file = save_last_uploaded(uploaded)
    source_label = uploaded.name
elif auto_reload and local_crm_file:
    # Priorité au fichier enrichi généré par le robot CRM local.
    # C'est ce fichier qui contient les colonnes Remarque / details_crm_ia.
    source_file = local_crm_file
    source_label = f"export CRM enrichi : {local_crm_file.name}"
elif auto_reload:
    source_file = get_last_uploaded_file()
    source_label = st.session_state.get("last_upload_name", "dernier fichier")

if local_crm_file:
    st.sidebar.success(f"Dernier export CRM détecté : {local_crm_file.name}")
else:
    st.sidebar.caption("Aucun export CRM enrichi détecté dans exports_crm pour le moment.")

if source_file:
    try:
        df = prepare_dataframe(source_file)
        df = apply_extra_steps(df)
        if df.empty:
            st.error("Aucune adresse trouvée dans le fichier.")
            st.stop()
        st.success(f"{len(df)} RDV chargés depuis {source_label}. Calcul automatique lancé, aucun bouton à cliquer.")
        with st.spinner("Géocodage, trajets, pauses, départs conseillés..."):
            route_df, return_row, start_geo = enrich_route(df, start_address, int(safety_min), int(visit_min), use_google, google_key, route_pref, include_tolls)
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
# IK estimée par trajet, basée sur les réglages rapides de la barre latérale.
route_df, return_ik_amount, current_ik_total, current_ik_formula = add_ik_amounts_to_route(
    route_df, return_row,
    cv=st.session_state.get("sidebar_cv", 7),
    electric=st.session_state.get("sidebar_electric", False),
    include_return=st.session_state.get("sidebar_return_ik", True),
    mode=sidebar_ik_mode,
    manual_rate=sidebar_manual_rate,
)
if return_row is not None:
    return_row["ik_montant_trajet"] = float(return_ik_amount or 0)


distance_series = pd.to_numeric(route_df.get("distance_depuis_precedent_km"), errors="coerce").fillna(0)
time_series = pd.to_numeric(route_df.get("temps_route_depuis_precedent_min"), errors="coerce").fillna(0)
total_km = float(distance_series.sum()) + (to_float(return_row.get("distance_depuis_precedent_km", 0)) if return_row else 0.0)
total_min = int(time_series.sum()) + (to_minutes(return_row.get("temps_route_depuis_precedent_min", 0)) if return_row else 0)
total_tolls=float(pd.to_numeric(route_df.get("peage_estime",pd.Series(dtype=float)),errors="coerce").fillna(0).sum())+(to_float(return_row.get("peage_estime",0)) if return_row else 0.0)



# ==========================================================
# V28.5 — Mode terrain direct depuis la carte
# ==========================================================
terrain_direct_no = str(st.query_params.get("terrain", "") or "")
if terrain_direct_no:
    matches = route_df[route_df["numero_rdv"].astype(str) == terrain_direct_no]
    if not matches.empty:
        tr = matches.iloc[0]
        st.markdown(
            f"""
            <div style="background:linear-gradient(145deg,#07101d,#0b2139);
                        border:2px solid #00bfff;border-radius:20px;padding:16px;
                        margin:8px 0 16px;box-shadow:0 14px 38px rgba(0,191,255,.18)">
              <div style="font-size:.78rem;color:#80e8ff;font-weight:950;letter-spacing:.06em">MODE TERRAIN</div>
              <div style="font-size:1.45rem;color:#fff;font-weight:950;margin-top:4px">
                {fmt_time(tr.get('heure_rdv'))} · {tr.get('nom_prospect','')}
              </div>
              <div style="color:#cbd5e1;margin-top:5px">{tr.get('adresse_complete','')}</div>
            </div>
            """,
            unsafe_allow_html=True
        )
        tc1,tc2,tc3=st.columns(3)
        tc1.link_button("🚗 Waze",tr.get("waze","#"),use_container_width=True)
        tc2.link_button("🗺️ Google Maps",tr.get("google_maps","#"),use_container_width=True)
        tc3.link_button("🏠 Voir maison",tr.get("street_view","#"),use_container_width=True)

        st.markdown("#### 🧠 Préparation commerciale")
        remark=str(tr.get("remarque_crm","") or "")
        if remark:
            st.info(remark)
        ia_key=f"prepa_{tr.get('numero_rdv','')}"
        ia_default=str(tr.get("analyse_ia_importee","") or build_preparation_advice(remark))
        ia_text=st.text_area("Préparation IA",value=st.session_state.get(ia_key,ia_default),height=180,key=f"direct_{ia_key}")
        st.session_state[ia_key]=ia_text

        st.markdown("#### 📝 Rapport de rendez-vous")
        d1,d2=st.columns(2)
        with d1:
            st.selectbox("Statut",["À faire","Signé","Négatif","Absent","À rappeler"],key=f"direct_status_{tr.get('numero_rdv','')}")
        with d2:
            st.date_input("Date de rappel",value=date.today(),key=f"direct_reminder_{tr.get('numero_rdv','')}")
        st.text_area("Commentaire / compte-rendu",key=f"direct_comment_{tr.get('numero_rdv','')}",height=130)
        st.divider()

# ===== V27 : dashboard terrain premium =====
tour_date = None
try:
    dates = [d for d in route_df.get("date_rdv", []) if isinstance(d, date)]
    if dates:
        tour_date = dates[0]
except Exception:
    pass

first_dep = route_df.iloc[0].get("depart_conseille") if not route_df.empty else None
retour_estime = None
if return_row and isinstance(return_row.get("rdv_datetime"), datetime):
    ret_m = to_minutes(return_row.get("temps_route_depuis_precedent_min", 0))
    retour_estime = return_row.get("rdv_datetime") + timedelta(minutes=ret_m)

summary_title = french_long_date(tour_date) if tour_date else "Ma tournée"
summary_bits=[f"{len(route_df)} étapes",f"{total_km:.0f} km",f"{fmt_duration(total_min)} de route"]
if total_tolls>0: summary_bits.append(f"{euro(total_tolls)} péages")
summary_bits.append(f"{euro(current_ik_total)} IK")
if fmt_dt(first_dep): summary_bits.append(f"départ {fmt_dt(first_dep)}")
if fmt_dt(retour_estime): summary_bits.append(f"retour ~{fmt_dt(retour_estime)}")

st.markdown(
    f"""<div class="cockpit-head">
      <div>
        <div class="cockpit-date">{summary_title}</div>
        <div class="cockpit-stats">{" · ".join(summary_bits)}</div>
      </div>
      <div class="live-pills">
        <span class="live-pill">GDH</span>
        <span class="live-pill ok">● TRAFIC GOOGLE</span>
        <span class="live-pill">GPS</span>
        <span class="live-pill alert">📷 RADARS</span>
      </div>
    </div>""",
    unsafe_allow_html=True
)

st.subheader("🗺️ Ma tournée")

# GPS iPhone : l'utilisateur garde le contrôle et autorise explicitement la localisation.
current_position = st.session_state.get("current_position")
with st.expander("📍 Ma position sur la carte", expanded=False):
    if GEOLOCATION_COMPONENT_AVAILABLE:
        st.caption("Sur iPhone, appuie sur le bouton ci-dessous puis autorise Safari à utiliser ta position.")
        gps_result = streamlit_geolocation()
        if isinstance(gps_result, dict) and gps_result.get("latitude") is not None and gps_result.get("longitude") is not None:
            st.session_state["current_position"] = {
                "latitude": gps_result.get("latitude"),
                "longitude": gps_result.get("longitude"),
                "accuracy": gps_result.get("accuracy"),
            }
            current_position = st.session_state["current_position"]
            accuracy = current_position.get("accuracy")
            if accuracy:
                st.success(f"📍 Position obtenue · précision ≈ {float(accuracy):.0f} m")
            else:
                st.success("📍 Position obtenue")
    else:
        st.warning("Le module GPS n'est pas installé. Ajoute `streamlit-geolocation` dans requirements.txt.")

st.markdown('<div class="map-legend">Carte vectorielle routière · clients · temps · départs · péages · position GPS · radars fixes publics.</div>', unsafe_allow_html=True)

mc1, mc2, mc3 = st.columns(3)
with mc1:
    map_style = st.selectbox(
        "Style de carte",
        ["Liberty","Bright","Positron","Dark","Fiord"],
        index=0,
        key="map_style_v28",
        help="Liberty est le style par défaut, lisible et routier."
    )
with mc2:
    show_radars = st.toggle("📷 Radars fixes", value=True, key="show_radars_v28")
with mc3:
    show_fuel = st.toggle("⛽ Stations-service", value=False, key="show_fuel_v28")

fuel_type = "gazole"
if show_fuel:
    fuel_label = st.segmented_control(
        "Prix affiché sur la carte",
        options=["Gazole","SP95","E10","SP98","E85","GPLc"],
        default="Gazole",
        key="fuel_type_v28"
    )
    fuel_type = {"Gazole":"gazole","SP95":"sp95","E10":"e10","SP98":"sp98","E85":"e85","GPLc":"gplc"}.get(fuel_label or "Gazole","gazole")

layer_status=[]
if show_radars:
    layer_status.append(f"📷 {len(radars_near_route(route_df, return_row, current_position))} radar(s)")
if show_fuel:
    layer_status.append(f"⛽ {len(stations_near_route(route_df, return_row, current_position))} station(s)")
if layer_status:
    st.caption(" · ".join(layer_status))

try:
    map_html = build_maplibre_html(
        route_df, return_row, start_address, start_geo,
        current_position=current_position,
        show_radars=show_radars,
        show_fuel=show_fuel,
        fuel_type=fuel_type,
        style_name=map_style,
        app_url=(st.context.url if hasattr(st, "context") else ""),
    )
    components.html(map_html, height=650, scrolling=False)
except Exception as e:
    st.warning(f"Carte moderne indisponible : {e}")

next_rdv = find_next_rdv(route_df)
if next_rdv is not None:
    with st.expander("🧭 Comparer les itinéraires du prochain trajet", expanded=False):
        st.caption("Même départ, même destination et même heure pour chaque option.")
        if google_key and isinstance(next_rdv.get("rdv_datetime"), datetime):
            idxs = route_df.index[
                route_df["numero_rdv"].astype(str) == str(next_rdv.get("numero_rdv"))
            ].tolist()
            ni = idxs[0] if idxs else 0
            cmp_origin = start_address if ni <= 0 else route_df.iloc[ni-1].get("adresse_complete", start_address)
            cmp_destination = next_rdv.get("adresse_complete", "")
            dep_guess = next_rdv.get("depart_conseille") or (next_rdv.get("rdv_datetime") - timedelta(hours=2))

            results_by_pref = {}
            for pk in ["recommended", "no_tolls", "no_highways", "shortest"]:
                rr_cmp = google_routes_traffic(
                    cmp_origin, cmp_destination, dep_guess, google_key,
                    route_pref=pk, include_tolls=True,
                )
                if rr_cmp:
                    results_by_pref[pk] = rr_cmp

            # La ligne "Plus court" reste silencieusement cohérente :
            # on affiche la variante la moins kilométrée obtenue.
            available = list(results_by_pref.items())
            if available:
                _, min_route = min(
                    available,
                    key=lambda item: float(item[1].get("km", 10**15) or 10**15)
                )
                if "shortest" not in results_by_pref or float(min_route.get("km", 10**15)) < float(results_by_pref["shortest"].get("km", 10**15)) - 0.1:
                    results_by_pref["shortest"] = dict(min_route)

            rows = []
            for pk in ["recommended", "no_tolls", "shortest", "no_highways"]:
                rr_cmp = results_by_pref.get(pk)
                if not rr_cmp:
                    continue
                else:
                    # Google Routes ne fournit pas toujours les tarifs de péage
                    # selon les pays/itinéraires. On évite d'afficher 0 € comme un prix certain.
                    toll_txt = "Non fourni par Google"

                rows.append({
                    "Itinéraire": ROUTE_PREF_LABELS[pk],
                    "Temps": fmt_duration(rr_cmp.get("min")),
                    "Distance": f"{float(rr_cmp.get('km',0)):.0f} km",
                })

            if rows:
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
            else:
                st.info("Comparaison indisponible pour ce trajet.")
        else:
            st.info("Comparaison disponible pour les trajets futurs avec Google Routes.")

with st.expander("📊 Détails de tournée (optionnel)", expanded=False):
    st.caption("Toutes les données restent disponibles, mais sont masquées par défaut pour garder l'écran terrain simple.")
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
    show_cols = ["numero_rdv", "heure_rdv", "depart_conseille", "pause_avant_rdv_min", "nom_prospect", "teleprospecteur", "telephone", "adresse_complete", "distance_depuis_precedent_km", "ik_montant_trajet", "temps_route_depuis_precedent_min", "note_trafic"]
    display_df = route_df[show_cols].copy()
    display_df["heure_rdv"] = display_df["heure_rdv"].apply(fmt_time)
    display_df["depart_conseille"] = display_df["depart_conseille"].apply(fmt_dt)
    display_df["pause_avant_rdv_min"] = display_df["pause_avant_rdv_min"].apply(lambda x: "" if x == "" else fmt_duration(x))
    display_df["temps_route_depuis_precedent_min"] = display_df["temps_route_depuis_precedent_min"].apply(fmt_duration)
    display_df["ik_montant_trajet"] = display_df["ik_montant_trajet"].apply(euro)
    display_df = display_df.rename(columns={
        "numero_rdv": "N° RDV", "heure_rdv": "Heure RDV", "depart_conseille": "Départ conseillé",
        "pause_avant_rdv_min": "Pause avant RDV", "nom_prospect": "Client", "teleprospecteur": "Téléprospecteur", "telephone": "Téléphone",
        "adresse_complete": "Adresse", "distance_depuis_precedent_km": "Km depuis précédent",
        "ik_montant_trajet": "IK trajet", "temps_route_depuis_precedent_min": "Temps depuis précédent", "note_trafic": "Calcul"
    })
    st.dataframe(display_df, use_container_width=True, hide_index=True)
    if return_row:
        st.info(f"Retour base inclus : {return_row.get('distance_depuis_precedent_km','')} km · {fmt_duration(return_row.get('temps_route_depuis_precedent_min',''))}")


# Rappels visibles en haut du mode terrain
crm_df = load_crm_history()
if not crm_df.empty:
    crm_df["_rappel_dt"] = crm_df.apply(reminder_datetime, axis=1)
    now_dt = datetime.now()
    reminders = crm_df[(crm_df.get("statut", "") == "À rappeler") & crm_df["_rappel_dt"].notna()].copy()
    reminders = reminders.sort_values("_rappel_dt")

    untreated = reminders[~reminders.apply(is_reminder_treated, axis=1)].copy() if not reminders.empty else reminders
    treated = reminders[reminders.apply(is_reminder_treated, axis=1)].copy() if not reminders.empty else reminders

    today_reminders = untreated[untreated["_rappel_dt"].apply(lambda x: reminder_section_label(x, now_dt) == "today")]
    late_reminders = untreated[untreated["_rappel_dt"].apply(lambda x: reminder_section_label(x, now_dt) == "late")]
    future_reminders = untreated[untreated["_rappel_dt"].apply(lambda x: reminder_section_label(x, now_dt) == "future")]

    def render_reminder_card(rr, label, css_class):
        dep = extract_departement(rr.get("adresse", ""))
        note = str(rr.get("commentaire", "") or "").strip()
        client = rr.get("client", "")
        dt_txt = f"{rr.get('date_rappel','')} {rr.get('heure_rappel','')}".strip()
        note_html = ("<span>📝 " + note + "</span>") if note else "<span>📝 Aucune note renseignée</span>"
        dep_txt = (" (" + dep + ")") if dep else ""
        st.markdown(f"""
<div class="reminder-card {css_class}">
  <strong>{label} — {client}{dep_txt}</strong><br>
  <span>🕒 {dt_txt}</span><br>
  {note_html}
</div>
""", unsafe_allow_html=True)
        cols_rem = st.columns([1, 1, 1, 1])
        tel_digits = re.sub(r"\D", "", str(rr.get("telephone", "")))
        if tel_digits:
            cols_rem[0].link_button("📞 Appeler", f"tel:{tel_digits}", use_container_width=True)
        cols_rem[1].link_button("🗺️ Adresse", maps_link(rr.get("adresse", "")), use_container_width=True)
        cols_rem[2].link_button("📤 WhatsApp", whatsapp_report_link(
            client=client, departement=dep, statut=rr.get("statut", ""),
            commentaire=note, rappel_date=rr.get("date_rappel", ""), rappel_heure=rr.get("heure_rappel", ""),
            adresse=rr.get("adresse", ""), telephone=rr.get("telephone", "")
        ), use_container_width=True)
        key_done = f"done_rem_{abs(hash(str(rr.get('key','')) + dt_txt))}"
        if cols_rem[3].button("✅ Traité", key=key_done, use_container_width=True):
            mark_reminder_treated(rr.get("key", ""), True)
            st.success("Rappel marqué comme traité.")
            st.rerun()

    if not today_reminders.empty or not late_reminders.empty or not future_reminders.empty:
        st.subheader("🔔 Rappels à traiter")
        if not today_reminders.empty:
            st.markdown("### 🚨 À faire aujourd’hui")
            for _, rr in today_reminders.iterrows():
                render_reminder_card(rr, "🚨 AUJOURD’HUI", "reminder-today")
        if not late_reminders.empty:
            st.markdown("### 🔴 En retard")
            for _, rr in late_reminders.iterrows():
                render_reminder_card(rr, "🔴 EN RETARD", "reminder-late")
        if not future_reminders.empty:
            st.markdown("### 🟡 Futurs rappels")
            for _, rr in future_reminders.head(20).iterrows():
                render_reminder_card(rr, "🟡 À VENIR", "reminder-future")

    with st.expander("✅ Rappels traités / archivés", expanded=False):
        if treated.empty:
            st.caption("Aucun rappel traité pour le moment.")
        else:
            for _, rr in treated.sort_values("_rappel_dt", ascending=False).head(50).iterrows():
                dep = extract_departement(rr.get("adresse", ""))
                note = str(rr.get("commentaire", "") or "").strip()
                dep_txt = (" (" + dep + ")") if dep else ""
                note_html = ("<span>📝 " + note + "</span>") if note else ""
                st.markdown(f"""
<div class="reminder-card reminder-treated">
  <strong>✅ TRAITÉ — {rr.get('client','')}{dep_txt}</strong><br>
  <span>🕒 {rr.get('date_rappel','')} {rr.get('heure_rappel','')} · traité le {rr.get('traite_at','')}</span><br>
  {note_html}
</div>
""", unsafe_allow_html=True)
                if st.button("↩️ Remettre à traiter", key=f"undone_rem_{abs(hash(str(rr.get('key',''))))}"):
                    mark_reminder_treated(rr.get("key", ""), False)
                    st.rerun()


# Recherche globale CRM + tournée courante
st.subheader("🔎 Recherche globale")
search_q = st.text_input("Rechercher un client, téléphone, adresse, téléprospecteur, commentaire ou statut", value="", placeholder="Ex : Dupont, 06..., À rappeler, signé...")
if search_q.strip():
    q = search_q.strip().lower()
    results = []
    # RDV de la tournée courante
    for _, rr in route_df.iterrows():
        hay = " ".join([str(rr.get(c, "")) for c in ["nom_prospect", "telephone", "adresse_complete", "teleprospecteur", "fournisseur", "commercial", "details_crm", "remarque_crm"]]).lower()
        if q in hay:
            results.append({
                "Source": "Tournée actuelle",
                "Date": rr.get("date_rdv", ""),
                "Heure": fmt_time(rr.get("heure_rdv", "")),
                "Client": rr.get("nom_prospect", ""),
                "Téléphone": rr.get("telephone", ""),
                "Adresse": rr.get("adresse_complete", ""),
                "Téléprospecteur": rr.get("teleprospecteur", ""),
                "Statut": "",
                "Commentaire": "",
            })
    # Historique CRM
    crm_search_df = load_crm_history()
    if not crm_search_df.empty:
        for _, rr in crm_search_df.iterrows():
            hay = " ".join([str(rr.get(c, "")) for c in crm_columns()]).lower()
            if q in hay:
                results.append({
                    "Source": "Historique CRM",
                    "Date": rr.get("date_rdv", ""),
                    "Heure": rr.get("heure_rdv", ""),
                    "Client": rr.get("client", ""),
                    "Téléphone": rr.get("telephone", ""),
                    "Adresse": rr.get("adresse", ""),
                    "Téléprospecteur": rr.get("teleprospecteur", ""),
                    "Statut": rr.get("statut", ""),
                    "Commentaire": rr.get("commentaire", ""),
                    "Détails CRM": rr.get("details_crm", ""),
                    "Analyse IA": rr.get("analyse_ia", ""),
                })
    if results:
        res_df = pd.DataFrame(results).drop_duplicates()
        st.dataframe(res_df, use_container_width=True, hide_index=True)
        for i, rr in res_df.head(10).iterrows():
            with st.container(border=True):
                st.markdown(f"**{rr.get('Client','')}** · {rr.get('Date','')} {rr.get('Heure','')} · {rr.get('Statut','')}")
                st.caption(f"{rr.get('Adresse','')} · Téléprospecteur : {rr.get('Téléprospecteur','')}")
                if rr.get("Commentaire"):
                    st.markdown(rr.get("Commentaire"))
                tel_digits = re.sub(r"\D", "", str(rr.get("Téléphone", "")))
                c_a, c_b, c_c = st.columns(3)
                if tel_digits:
                    c_a.link_button("📞 Appeler", f"tel:{tel_digits}", use_container_width=True)
                c_b.link_button("🗺️ Adresse", maps_link(rr.get("Adresse", "")), use_container_width=True)
                dep = extract_departement(rr.get("Adresse", ""))
                c_c.link_button("📤 WhatsApp", whatsapp_report_link(
                    client=rr.get("Client", ""), departement=dep, statut=rr.get("Statut", ""),
                    commentaire=rr.get("Commentaire", ""), adresse=rr.get("Adresse", ""), telephone=rr.get("Téléphone", "")
                ), use_container_width=True)
    else:
        st.info("Aucun résultat trouvé.")

st.subheader("📋 Mode terrain")
selected_terrain = str(st.query_params.get("terrain", "") or "")
for _, r in route_df.iterrows():
    terrain_no = str(r.get("numero_rdv", ""))
    st.markdown(f'<div id="terrain-{terrain_no}" class="terrain-anchor"></div>', unsafe_allow_html=True)
    pause = r.get('pause_avant_rdv_min', '')
    pause_txt = "" if pause == "" else (f" · Pause dispo : {fmt_duration(pause)}" if to_minutes(pause) >= 0 else f" · ⚠ Retard probable : {fmt_duration(abs(to_minutes(pause)))}")
    title = f"RDV {r.get('numero_rdv','')} · {fmt_time(r.get('heure_rdv'))} · {r.get('nom_prospect','')}{pause_txt}"
    is_selected_terrain = bool(selected_terrain) and selected_terrain == terrain_no
    with st.expander(title, expanded=(is_selected_terrain or (not selected_terrain and str(r.get('ordre','')) == '1'))):
        c1, c2 = st.columns([2, 1])
        with c1:
            st.markdown(f"**Adresse :** {r.get('adresse_complete','')}")
            st.markdown(f"**Téléphone :** {r.get('telephone','')}")
            st.caption("Trajet, départ conseillé et péage sont déjà visibles sur la carte.")
        with c2:
            st.link_button("🚗 Waze", r.get('waze', '#'), use_container_width=True)
            st.link_button("🗺️ Google Maps", r.get('google_maps', '#'), use_container_width=True)
            st.link_button("🏠 Voir maison", r.get('street_view', '#'), use_container_width=True)
            if r.get('telephone_tel'):
                st.link_button("📞 Appeler", f"tel:{r.get('telephone_tel')}", use_container_width=True)
            # Bouton WhatsApp : préremplit un compte rendu à envoyer manuellement au groupe.
            st.link_button("📤 WhatsApp", whatsapp_report_link(
                client=r.get('nom_prospect',''),
                departement=extract_departement(r.get('adresse_complete','')),
                statut='',
                commentaire='',
                adresse=r.get('adresse_complete',''),
                telephone=r.get('telephone','')
            ), use_container_width=True)
        st.divider()
        key = crm_key(r)
        crm_hist = load_crm_history()
        previous = crm_hist[crm_hist["key"].astype(str) == key].tail(1) if not crm_hist.empty else pd.DataFrame()
        prev_statut = previous.iloc[0].get("statut", "") if not previous.empty else ""
        prev_comment = previous.iloc[0].get("commentaire", "") if not previous.empty else ""
        imported_details_crm = str(r.get("details_crm", "") or r.get("remarque_crm", "") or "").strip()
        imported_analyse_ia = str(r.get("analyse_ia_importee", "") or "").strip()
        saved_details_crm = previous.iloc[0].get("details_crm", "") if not previous.empty else ""
        saved_analyse_ia = previous.iloc[0].get("analyse_ia", "") if not previous.empty else ""
        # V26.3 : quand le robot CRM a enrichi l'Excel, on préremplit toujours la Prépa IA avec la Remarque du fichier.
        # Si aucun export enrichi n'est présent, on conserve l'ancien contenu sauvegardé.
        prev_details_crm = imported_details_crm or str(saved_details_crm or "").strip()
        prev_analyse_ia = imported_analyse_ia or str(saved_analyse_ia or "").strip()
        prev_rappel_date = parse_date(previous.iloc[0].get("date_rappel", "")) if not previous.empty else None
        prev_rappel_time = parse_time(previous.iloc[0].get("heure_rappel", "")) if not previous.empty else None
        statuses = ["", "Signé", "Veut réfléchir", "Absent", "Négatif", "À rappeler", "VT à planifier", "À revoir"]
        sidx = statuses.index(prev_statut) if prev_statut in statuses else 0
        cr1, cr2 = st.columns([1,2])
        with cr1:
            statut = st.selectbox("Statut fin RDV", statuses, index=sidx, key=f"statut_{abs(hash(key))}")
        with cr2:
            commentaire = st.text_area("Commentaire fin RDV", value=prev_comment, key=f"comment_{abs(hash(key))}", height=80)
        st.markdown("#### 🧠 Préparation IA du RDV")
        details_crm = st.text_area(
            "Coller ici les infos détaillées du CRM (chauffage, RFR, situation, remarques télépro, etc.)",
            value=prev_details_crm,
            key=f"detailscrm_{abs(hash(key))}",
            height=120,
            placeholder="Ex : Mariés retraités, chauffage bois, RFR..., maison ancienne, hésite sur le financement..."
        )
        analyse_ia = analyse_crm_details(details_crm)
        if analyse_ia:
            st.markdown("**Conseils terrain générés :**")
            st.text_area("Analyse / stratégie avant RDV", value=analyse_ia, key=f"analyseia_{abs(hash(key))}", height=220, disabled=True)
            st.link_button("📤 Envoyer préparation WhatsApp", whatsapp_ai_prep_link(
                client=r.get('nom_prospect',''),
                departement=extract_departement(r.get('adresse_complete','')),
                adresse=r.get('adresse_complete',''),
                details=details_crm,
                analyse=analyse_ia
            ), use_container_width=True)
        else:
            analyse_ia = prev_analyse_ia

        rappel_date = None
        rappel_time = None
        if statut == "À rappeler":
            rr1, rr2 = st.columns(2)
            with rr1:
                rappel_date = st.date_input("Date de rappel", value=prev_rappel_date or date.today(), key=f"rdate_{abs(hash(key))}")
            with rr2:
                rappel_time = st.time_input("Heure de rappel", value=prev_rappel_time or dtime(9,0), key=f"rtime_{abs(hash(key))}")
        st.link_button("📤 Envoyer ce compte rendu WhatsApp", whatsapp_report_link(
            client=r.get('nom_prospect',''),
            departement=extract_departement(r.get('adresse_complete','')),
            statut=statut,
            commentaire=commentaire,
            rappel_date=rappel_date.strftime("%d/%m/%Y") if isinstance(rappel_date, date) else "",
            rappel_heure=fmt_time(rappel_time) if isinstance(rappel_time, dtime) else "",
            adresse=r.get('adresse_complete',''),
            telephone=r.get('telephone','')
        ), use_container_width=True)
        # V21 : sauvegarde automatique dès qu'un statut, commentaire ou rappel est saisi.
        if statut or str(commentaire).strip() or rappel_date or rappel_time:
            save_crm_record(key, r, statut, commentaire, rappel_date, rappel_time, details_crm, analyse_ia)
            st.caption("💾 Sauvegarde automatique active")
        if st.button("💾 Enregistrer compte rendu", key=f"savecrm_{abs(hash(key))}"):
            save_crm_record(key, r, statut, commentaire, rappel_date, rappel_time, details_crm, analyse_ia)
            st.success("Compte rendu enregistré.")


with st.expander("📤 Documents & exports", expanded=False):
    include_photos = st.checkbox("Essayer d'intégrer les photos Street View dans le PDF", value=bool(google_key), help="Nécessite une clé Google Maps API. Sinon le PDF contient le lien Voir maison cliquable.")
    pdf_bytes = create_pdf(route_df, return_row, start_address, include_photos, google_key, int(visit_min))
    csv_bytes = to_recap_csv(route_df, return_row)

    c1, c2 = st.columns(2)
    with c1:
        st.download_button("📄 Télécharger PDF enrichi cliquable", data=pdf_bytes, file_name="tournee_terrain_v18.pdf", mime="application/pdf", use_container_width=True)
    with c2:
        st.download_button("💾 Sauvegarde CSV réutilisable", data=csv_bytes, file_name="tournee_sauvegarde_v18.csv", mime="text/csv", use_container_width=True)




with st.expander("💰 Indemnités kilométriques", expanded=False):
    st.caption("Objectif : une note mensuelle professionnelle, avec historique des journées et barème modifiable si les règles changent.")

    # Registre du jour courant, sauvegardé automatiquement en historique local Streamlit.
    current_register_base = build_ik_register(route_df, return_row, st.session_state.get("sidebar_return_ik", True), start_address)
    if not current_register_base.empty:
        saved_path = save_ik_history(current_register_base)
        if saved_path:
            st.success(f"Journée enregistrée dans l'historique IK : {saved_path.name}")

    st.markdown("#### ⚙️ Paramètres de la note IK")
    with st.container(border=True):
        a,b,c = st.columns(3)
        with a:
            beneficiaire = st.text_input("Bénéficiaire", value=settings.get("beneficiaire", "Mr Dahan"))
            societe = st.text_input("Société à facturer / rembourser", value=settings.get("societe", ""))
            periode = st.text_input("Libellé période", value=datetime.now().strftime("%B %Y"))
        with b:
            vehicule = st.text_input("Véhicule", value=settings.get("vehicule", ""))
            immat = st.text_input("Immatriculation", value=settings.get("immat", ""))
            cv = st.selectbox("Puissance fiscale", options=[3,4,5,6,7], index=[3,4,5,6,7].index(st.session_state.get("sidebar_cv", 7)), help="7 = 7 CV et plus")
        with c:
            electric = st.checkbox("Véhicule 100% électrique (+20%)", value=st.session_state.get("sidebar_electric", False))
            include_return_ik = st.checkbox("Inclure le retour à la base", value=st.session_state.get("sidebar_return_ik", True))
            st.info("Mode appliqué : " + (f"forfait interne {sidebar_manual_rate:.3f} €/km" if sidebar_ik_mode == "manuel" else "barème officiel paramétrable"))

        custom_bareme_enabled = st.checkbox("Afficher / modifier le barème IK", value=False)
        custom_bareme = custom_bareme_from_inputs("ik_custom") if custom_bareme_enabled else IK_BAREME_2026
        save_app_settings({
            "beneficiaire": beneficiaire, "societe": societe, "vehicule": vehicule, "immat": immat,
            "sidebar_cv": int(cv), "sidebar_electric": bool(electric), "sidebar_return_ik": bool(include_return_ik),
        })

    st.markdown("#### 📚 Historique / facturation IK")
    with st.container(border=True):
        h1, h2, h3 = st.columns([1,1,2])
        today = date.today()
        with h1:
            period_start = st.date_input("Période du", value=date(today.year, today.month, 1))
        with h2:
            period_end = st.date_input("Au", value=today)
        with h3:
            uploaded_history = st.file_uploader("Ajouter un registre IK CSV ancien si besoin", type=["csv"], key="ik_history_upload")

        history_df = load_ik_history()
        frames = []
        if not history_df.empty:
            frames.append(history_df)
        # Toujours inclure la journée courante pour éviter d'attendre une sauvegarde serveur.
        if not current_register_base.empty:
            today_df = current_register_base.copy()
            today_df["Source"] = "journée actuelle"
            frames.append(today_df)
        if uploaded_history is not None:
            try:
                extra = pd.read_csv(uploaded_history, sep=";")
                extra["Source"] = uploaded_history.name
                frames.append(extra)
            except Exception as e:
                st.warning(f"Registre ajouté illisible : {e}")

        all_register = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    
        if not all_register.empty and "Date" in all_register.columns:
            tmp = all_register.copy()
            parsed = pd.to_datetime(tmp["Date"], dayfirst=True, errors="coerce")
            monthly_register = tmp[(parsed.dt.date >= period_start) & (parsed.dt.date <= period_end)].copy()
        else:
            monthly_register = pd.DataFrame()
        if not monthly_register.empty:
            # Déduplique les lignes principales pour éviter les doublons de la journée actuelle + historique.
            dedup_cols = [c for c in ["Date", "Objet", "Arrivée", "Km", "Justificatif"] if c in monthly_register.columns]
            monthly_register = monthly_register.drop_duplicates(subset=dedup_cols, keep="last") if dedup_cols else monthly_register
        st.caption("L'historique est conservé tant que l'application Streamlit garde son espace temporaire. Pour un archivage durable, télécharge le CSV mensuel.")

    ik_register = monthly_register if 'monthly_register' in locals() and not monthly_register.empty else current_register_base
    ik_register = ik_register.copy()
    total_ik_km = float(pd.to_numeric(ik_register.get("Km", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()) if not ik_register.empty else 0.0
    ik_amount, ik_formula = calc_ik_amount_flexible(total_ik_km, cv=cv, electric=electric, bareme=custom_bareme, mode=sidebar_ik_mode, manual_rate=sidebar_manual_rate)
    ik_register = add_ik_amounts_to_register(ik_register, ik_amount)

    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Km IK période", f"{total_ik_km:.1f} km")
    k2.metric("Montant IK", euro(ik_amount))
    k3.metric("Barème", f"{cv} CV" + (" électrique" if electric else ""))
    k4.metric("Lignes", len(ik_register))

    st.dataframe(ik_register, use_container_width=True, hide_index=True)

    ik_params = {
        "beneficiaire": beneficiaire,
        "societe": societe,
        "periode": periode,
        "vehicule": vehicule,
        "immat": immat,
        "cv": cv,
        "electric": electric,
        "bareme": "Barème kilométrique paramétrable — voitures",
        "amount": ik_amount,
        "formula": ik_formula,
    }
    ik_pdf = create_ik_pdf(ik_register, ik_params)
    d1,d2 = st.columns(2)
    with d1:
        st.download_button("📄 Télécharger la note IK mensuelle PDF", data=ik_pdf, file_name="note_indemnites_kilometriques_mensuelle.pdf", mime="application/pdf", use_container_width=True)
    with d2:
        st.download_button("📊 Télécharger le registre IK mensuel CSV", data=df_to_csv_bytes(ik_register), file_name="registre_indemnites_kilometriques_mensuel.csv", mime="text/csv", use_container_width=True)



st.caption("Routage PRO · GDH — V28.5 — 28/07/2026 · alerte retard sur carte · Mode terrain direct · GPS · radars · stations-service · trafic Google")
