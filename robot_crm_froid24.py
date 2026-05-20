"""
Robot expérimental FROID24 / Alltoo CRM
Objectif : télécharger automatiquement l'Excel des RDV d'une date donnée.

Important :
- Lance ce script sur ta Surface, pas dans Streamlit Cloud.
- Tes identifiants sont saisis à l'écran et ne sont pas écrits dans le code.
- Le CRM peut changer : si un bouton/texte change, le robot peut nécessiter un ajustement.
"""
from __future__ import annotations

import asyncio
from pathlib import Path
from getpass import getpass
from datetime import datetime

try:
    from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError
except Exception:
    print("Playwright n'est pas installé. Lance : pip install playwright puis playwright install chromium")
    raise

BASE_URL = "https://froid24.callcenter.alltoo.fr"
LOGIN_URL = BASE_URL + "/accounts/login/?next=/"
RDV_URL = BASE_URL + "/rendez_vous/mes_rendez_vous/"
OUT_DIR = Path("exports_crm")
OUT_DIR.mkdir(exist_ok=True)

async def fill_first_available(page, selectors, value):
    for sel in selectors:
        try:
            loc = page.locator(sel).first
            if await loc.count():
                await loc.fill(value)
                return True
        except Exception:
            pass
    return False

async def click_text(page, texts, timeout=5000):
    for text in texts:
        try:
            await page.get_by_text(text, exact=False).first.click(timeout=timeout)
            return True
        except Exception:
            pass
    return False

async def main():
    print("=== Robot CRM FROID24 expérimental ===")
    username = input("Identifiant CRM : ").strip()
    password = getpass("Mot de passe CRM : ").strip()
    date_str = input("Date des RDV à exporter (JJ/MM/AAAA) : ").strip()
    if not date_str:
        date_str = datetime.now().strftime("%d/%m/%Y")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()
        print("Connexion...")
        await page.goto(LOGIN_URL, wait_until="domcontentloaded")

        ok_user = await fill_first_available(page, [
            'input[name="username"]', 'input[name="login"]', 'input[name="email"]',
            'input[type="text"]', 'input[placeholder*="ident" i]', 'input[placeholder*="login" i]'
        ], username)
        ok_pass = await fill_first_available(page, [
            'input[name="password"]', 'input[type="password"]'
        ], password)
        if not ok_user or not ok_pass:
            print("Impossible de trouver les champs login/mot de passe. Le CRM a peut-être changé.")
            await browser.close()
            return

        if not await click_text(page, ["Connexion", "Se connecter", "Login"], timeout=3000):
            await page.keyboard.press("Enter")
        await page.wait_for_load_state("networkidle")

        print("Ouverture de la page Rendez-vous...")
        await page.goto(RDV_URL, wait_until="networkidle")

        # Date début
        try:
            inputs = page.locator('input')
            n = await inputs.count()
            # Dans les captures, le premier champ date correspond à "Du".
            if n >= 1:
                await inputs.nth(0).fill(date_str)
        except Exception:
            pass

        # Actualiser
        await click_text(page, ["Actualisé", "Actualiser", "Rechercher"], timeout=3000)
        await page.wait_for_load_state("networkidle")

        print("Sélection des RDV visibles...")
        try:
            # Case globale ou cases ligne par ligne.
            checkboxes = page.locator('input[type="checkbox"]')
            count = await checkboxes.count()
            for i in range(count):
                try:
                    cb = checkboxes.nth(i)
                    if not await cb.is_checked():
                        await cb.check(force=True)
                except Exception:
                    pass
        except Exception:
            pass

        print("Export Excel...")
        try:
            async with page.expect_download(timeout=30000) as download_info:
                clicked = await click_text(page, ["Exporter les rendez-vous", "Exporter les rendez", "Exporter"], timeout=5000)
                if not clicked:
                    print("Clique manuellement sur Actions > Exporter les rendez-vous dans la fenêtre ouverte.")
            download = await download_info.value
            out = OUT_DIR / f"export_rdv_{date_str.replace('/', '-')}.xlsx"
            await download.save_as(out)
            print(f"✅ Excel téléchargé : {out.resolve()}")
        except PlaywrightTimeoutError:
            print("⚠️ Téléchargement non détecté automatiquement. Si le fichier s'est téléchargé dans Chrome, récupère-le dans Téléchargements.")
        except Exception as e:
            print(f"Erreur export : {e}")

        print("\nÉtape détails clients : expérimentale.")
        print("Pour l'instant, si le robot n'arrive pas à récupérer les remarques, ouvre chaque œil Voir/Editer et copie les remarques dans l'app V24.")
        input("Appuie sur Entrée pour fermer le navigateur...")
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
