import io
import re
import math
import json
import unicodedata
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

st.set_page_config(page_title="Routage PRO V26.3", page_icon="🚗", layout="wide")

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

st.title("🚗 Routage PRO V26.3 — terrain + IK + CRM + rappels + IA")
st.caption("Mode sombre · carte claire · IK · CRM persistant · rappels intelligents · WhatsApp · import CRM enrichi")

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


def make_map(df, return_row, start_address, start_geo, interactive=True):
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
    m = folium.Map(
        location=center,
        zoom_start=11,
        tiles="OpenStreetMap",
        dragging=interactive,
        scrollWheelZoom=interactive,
        touchZoom=interactive,
        doubleClickZoom=interactive,
        zoom_control=True,
    )
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

    # Zoom automatique : afficher toutes les destinations dès l'ouverture de la carte
    # (base + tous les RDV + retour base) plutôt qu'une simple portion de route.
    try:
        if len(points) >= 2:
            m.fit_bounds(points, padding=(35, 35))
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
    use_google = st.checkbox("Utiliser Google trafic / Street View si j'ai une clé API", value=bool(settings.get("use_google", False)))
    google_key = st.text_input("Clé Google Maps API (optionnel)", value=settings.get("google_key", ""), type="password") if use_google else ""
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
        "use_google": bool(use_google), "google_key": google_key, "sidebar_cv": int(sidebar_cv),
        "sidebar_electric": bool(sidebar_electric), "sidebar_return_ik": bool(sidebar_include_return),
        "ik_mode": sidebar_ik_mode, "manual_rate": float(sidebar_manual_rate),
    })
    st.info("V26.3 : charge en priorité le dernier export CRM enrichi généré par le robot local, avec remarques + IA.")

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
# IK estimée par trajet, basée sur les réglages rapides de la barre latérale.
route_df, return_ik_amount, current_ik_total, current_ik_formula = add_ik_amounts_to_route(
    route_df, return_row,
    cv=st.session_state.get("sidebar_cv", 7),
    electric=st.session_state.get("sidebar_electric", False),
    include_return=st.session_state.get("sidebar_return_ik", True),
    mode=sidebar_ik_mode,
    manual_rate=sidebar_manual_rate,
)

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
            st.markdown(f"**Indemnité estimée pour ce trajet :** {euro(r.get('ik_montant_trajet', 0))}")
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

st.subheader("🗺️ Carte générale")
nb_routes = sum(1 for _, rr in route_df.iterrows() if isinstance(rr.get("route_geometry", []), list) and len(rr.get("route_geometry", [])) >= 2)
if return_row and isinstance(return_row.get("route_geometry", []), list) and len(return_row.get("route_geometry", [])) >= 2:
    nb_routes += 1
if nb_routes == 0:
    st.warning("Aucun tracé routier disponible pour l’instant. Vérifie la connexion ou les adresses. Les calculs peuvent quand même apparaître si les coordonnées sont trouvées.")
else:
    st.success(f"{nb_routes} trajet(s) routier(s) tracé(s) sur la carte.")
map_interactive = st.toggle(
    "Activer déplacement / zoom sur la carte",
    value=False,
    help="Désactivé par défaut pour que le scroll iPhone fasse défiler l’application au lieu de bouger la carte."
)
if not map_interactive:
    st.caption("📱 Carte verrouillée : le scroll iPhone fait défiler l’application. Active le bouton ci-dessus si tu veux déplacer/zoomer la carte.")
    st.markdown("""<style>iframe[title="streamlit_folium.st_folium"]{pointer-events:none!important;}</style>""", unsafe_allow_html=True)
else:
    st.caption("🗺️ Carte interactive activée : tu peux zoomer/déplacer la carte.")
try:
    st_folium(make_map(route_df, return_row, start_address, start_geo, interactive=map_interactive), height=650, use_container_width=True)
except Exception as e:
    st.warning(f"Carte non disponible : {e}")

st.subheader("📤 Exports terrain")
include_photos = st.checkbox("Essayer d'intégrer les photos Street View dans le PDF", value=bool(google_key), help="Nécessite une clé Google Maps API. Sinon le PDF contient le lien Voir maison cliquable.")
pdf_bytes = create_pdf(route_df, return_row, start_address, include_photos, google_key, int(visit_min))
csv_bytes = to_recap_csv(route_df, return_row)

c1, c2 = st.columns(2)
with c1:
    st.download_button("📄 Télécharger PDF enrichi cliquable", data=pdf_bytes, file_name="tournee_terrain_v18.pdf", mime="application/pdf", use_container_width=True)
with c2:
    st.download_button("💾 Sauvegarde CSV réutilisable", data=csv_bytes, file_name="tournee_sauvegarde_v18.csv", mime="text/csv", use_container_width=True)



st.subheader("💰 Indemnités kilométriques — V18 mensuelle")
st.caption("Objectif : une note mensuelle professionnelle, avec historique des journées et barème modifiable si les règles changent.")

# Registre du jour courant, sauvegardé automatiquement en historique local Streamlit.
current_register_base = build_ik_register(route_df, return_row, st.session_state.get("sidebar_return_ik", True), start_address)
if not current_register_base.empty:
    saved_path = save_ik_history(current_register_base)
    if saved_path:
        st.success(f"Journée enregistrée dans l'historique IK : {saved_path.name}")

with st.expander("⚙️ Paramètres de la note IK", expanded=False):
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

with st.expander("📚 Historique / facturation IK", expanded=True):
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


st.caption("V22 : V21.2 stable + préparation RDV depuis détail CRM. Sans clé Google, le trafic est une estimation prudente. Les tracés routiers utilisent OSRM gratuit quand les coordonnées sont trouvées.")
