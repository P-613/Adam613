#!/usr/bin/env python3
"""
Kleinanzeigen.de – Automatische KI-Antworten auf Wohnungsanfragen.

Benötigte Umgebungsvariablen:
  KLEINANZEIGEN_EMAIL    – Ihre E-Mail-Adresse bei Kleinanzeigen
  KLEINANZEIGEN_PASSWORT – Ihr Passwort
  ANTHROPIC_API_KEY      – Ihr Anthropic API-Schlüssel

Verwendung:
  python ebay_kleinanzeigen_auto_reply.py             # Vorschau (kein Absenden)
  python ebay_kleinanzeigen_auto_reply.py --senden    # Antworten absenden
  python ebay_kleinanzeigen_auto_reply.py --max 5     # Maximal 5 Nachrichten
  python ebay_kleinanzeigen_auto_reply.py --sichtbar  # Browser sichtbar (für CAPTCHA)
"""

import os
import sys
import time
import argparse

import anthropic
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout


# ── Konfiguration ────────────────────────────────────────────────────────────
EMAIL = os.environ.get("KLEINANZEIGEN_EMAIL", "")
PASSWORT = os.environ.get("KLEINANZEIGEN_PASSWORT", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")

BASE_URL = "https://www.kleinanzeigen.de"

SYSTEM_PROMPT = """\
Du bist ein freundlicher und professioneller Vermieter, der auf Wohnungsanfragen \
von Mietinteressenten antwortet. Schreibe ausschließlich auf Deutsch.

Deine Antwort soll:
- Herzlich und einladend, aber seriös klingen
- 3–5 kurze Sätze umfassen
- Dich für das Interesse bedanken
- Folgende Informationen erfragen:
  * Anzahl der einziehenden Personen (Erwachsene, Kinder, Haustiere)
  * Gewünschter Einzugstermin
  * Berufliche / finanzielle Situation (knapp)
- Einen Besichtigungstermin anbieten

Schreibe nur den Antworttext – keine Betreffzeile, kein „Hallo" als eigene Zeile."""

# Selektoren für die Kleinanzeigen-Nachrichten-Oberfläche (Stand 2025/26)
SEL_COOKIE_ACCEPT = (
    "#gdpr-banner-accept, [data-testid='gdpr-banner-accept'], "
    "button[id*='accept'], button[class*='accept-all']"
)
SEL_LOGIN_EMAIL = "#login-email, input[name='loginEmail'], input[type='email']"
SEL_LOGIN_PW = "#login-password, input[name='loginPassword'], input[type='password']"
SEL_LOGIN_BTN = "#login-submit, button[type='submit']"
SEL_KONV_LISTE = (
    ".conversations-list li, .msg-thread-item, "
    "[data-testid='conversation-item'], .conversation-item"
)
SEL_KONV_LINK = "a[href*='/m-nachrichten'], a[href*='/nachricht']"
SEL_NACHRICHT_EINGANG = (
    ".message--inbound, .msg-received, [data-testid='message-inbound'], "
    ".message-item.received, .messages-list .incoming"
)
SEL_NACHRICHT_TEXT = (
    ".message-text, .msg-body, [data-testid='message-text'], "
    ".text-body, p.message"
)
SEL_ANTWORT_FELD = (
    "textarea[name='message'], textarea#message, #message-input, "
    ".compose-text textarea, textarea[placeholder*='Nachricht']"
)
SEL_SENDEN_BTN = (
    "button[type='submit'][form*='message'], button.send-button, "
    "#send-message-btn, button[aria-label*='Senden'], button[aria-label*='senden']"
)


# ── Hilfsfunktionen ──────────────────────────────────────────────────────────

def check_env() -> None:
    fehlend = [
        var for var, val in [
            ("KLEINANZEIGEN_EMAIL", EMAIL),
            ("KLEINANZEIGEN_PASSWORT", PASSWORT),
            ("ANTHROPIC_API_KEY", ANTHROPIC_API_KEY),
        ] if not val
    ]
    if fehlend:
        print(
            f"Fehler: Folgende Umgebungsvariablen fehlen:\n"
            + "\n".join(f"  {v}" for v in fehlend)
            + "\n\nSiehe .env.example für die Einrichtung."
        )
        sys.exit(1)


def klick_cookie_banner(page) -> None:
    try:
        page.click(SEL_COOKIE_ACCEPT, timeout=4000)
        time.sleep(0.5)
    except PlaywrightTimeout:
        pass


def generiere_antwort(anfrage: str, inserat_titel: str, absender: str) -> str:
    """Ruft die Claude-API auf und gibt eine passende Vermieter-Antwort zurück."""
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    nutzer_prompt = (
        f"Inserat: {inserat_titel}\n"
        f"Anfrage von {absender}:\n\n{anfrage}\n\n"
        "Bitte schreibe eine passende Antwort als Vermieter."
    )
    antwort = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=500,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": nutzer_prompt}],
    )
    return antwort.content[0].text.strip()


# ── Browser-Aktionen ─────────────────────────────────────────────────────────

def login(page, email: str, passwort: str) -> bool:
    """Loggt sich bei Kleinanzeigen ein. Gibt True zurück bei Erfolg."""
    try:
        page.goto(f"{BASE_URL}/m-einloggen.html", wait_until="domcontentloaded", timeout=20000)
        klick_cookie_banner(page)

        page.fill(SEL_LOGIN_EMAIL, email, timeout=8000)
        time.sleep(0.4)
        page.fill(SEL_LOGIN_PW, passwort, timeout=5000)
        time.sleep(0.4)
        page.click(SEL_LOGIN_BTN, timeout=5000)
        page.wait_for_url(lambda url: "/m-einloggen.html" not in url, timeout=15000)

        print(f"  ✓ Eingeloggt als {email}")
        return True

    except PlaywrightTimeout:
        # Prüfen ob wir bereits eingeloggt sind (Redirect hat funktioniert)
        if "/m-einloggen.html" not in page.url:
            print(f"  ✓ Eingeloggt als {email}")
            return True
        print(
            "  Fehler: Login fehlgeschlagen oder Timeout.\n"
            "  Tipp: Mit --sichtbar können Sie CAPTCHA manuell lösen."
        )
        return False


def hole_konversationen(page) -> list[dict]:
    """Navigiert zur Nachrichtenübersicht und gibt eine Liste offener Anfragen zurück."""
    page.goto(f"{BASE_URL}/m-nachrichten.html", wait_until="domcontentloaded", timeout=20000)
    klick_cookie_banner(page)
    time.sleep(2)

    konversationen = []
    elemente = page.query_selector_all(SEL_KONV_LISTE)

    if not elemente:
        # Fallback: alle Links, die auf Konversationen zeigen
        elemente = page.query_selector_all("a[href*='nachrichten'], a[href*='nachricht']")
        for el in elemente:
            href = el.get_attribute("href") or ""
            if not href:
                continue
            url = href if href.startswith("http") else BASE_URL + href
            konversationen.append({"url": url, "sender": "?", "titel": "?"})
        return konversationen

    for el in elemente:
        link = el.query_selector("a") or el
        href = link.get_attribute("href") or ""
        if not href:
            continue
        url = href if href.startswith("http") else BASE_URL + href

        sender_el = el.query_selector(".username, .sender, .conversation-partner, strong")
        sender = sender_el.inner_text().strip() if sender_el else "Unbekannt"

        titel_el = el.query_selector(".subject, .ad-title, .conversation-title, span")
        titel = titel_el.inner_text().strip() if titel_el else "Unbekannte Anzeige"

        konversationen.append({"url": url, "sender": sender, "titel": titel})

    return konversationen


def bearbeite_konversation(
    page, konv: dict, senden: bool = False
) -> dict | None:
    """
    Öffnet eine Konversation. Ist die letzte Nachricht eine unbeantwortete Anfrage,
    wird eine KI-Antwort generiert – und bei senden=True abgeschickt.
    Gibt None zurück, wenn keine Aktion nötig ist.
    """
    try:
        page.goto(konv["url"], wait_until="domcontentloaded", timeout=20000)
    except PlaywrightTimeout:
        print(f"  Übersprungen (Timeout): {konv['url']}")
        return None

    time.sleep(1.5)

    # Alle eingehenden Nachrichten in dieser Konversation
    eingehend = page.query_selector_all(SEL_NACHRICHT_EINGANG)

    if not eingehend:
        # Alternative: letzte Nachricht allgemein holen und prüfen
        alle = page.query_selector_all(".message, .msg-item, [data-testid='message']")
        if not alle:
            return None
        letzte = alle[-1]
        # Wenn die letzte Nachricht ein "sent/outgoing"-Merkmal trägt, ist sie beantwortet
        klassen = letzte.get_attribute("class") or ""
        if any(k in klassen for k in ("sent", "outgoing", "own", "sender")):
            return None
        eingehend = [letzte]

    letzte_anfrage = eingehend[-1]

    # Nachrichtentext extrahieren
    text_el = letzte_anfrage.query_selector(SEL_NACHRICHT_TEXT)
    anfrage_text = (
        text_el.inner_text().strip() if text_el
        else letzte_anfrage.inner_text().strip()
    )
    if not anfrage_text or len(anfrage_text) < 5:
        return None

    # Inserat-Titel und Absender aus der Seite verfeinern (falls vorhanden)
    inserat_el = page.query_selector("h1, .ad-title, [data-testid='ad-title']")
    inserat_titel = inserat_el.inner_text().strip() if inserat_el else konv["titel"]

    sender_el = page.query_selector(".conversation-partner, .username, .sender-name")
    sender = sender_el.inner_text().strip() if sender_el else konv["sender"]

    # KI-Antwort generieren
    print(f"  → Generiere Antwort für Anfrage von {sender}...")
    antwort = generiere_antwort(anfrage_text, inserat_titel, sender)

    ergebnis: dict = {
        "sender": sender,
        "titel": inserat_titel,
        "anfrage": anfrage_text[:300] + ("…" if len(anfrage_text) > 300 else ""),
        "antwort": antwort,
        "gesendet": False,
        "fehler": None,
    }

    if senden:
        try:
            feld = page.query_selector(SEL_ANTWORT_FELD)
            if not feld:
                ergebnis["fehler"] = "Eingabefeld nicht gefunden"
            else:
                feld.click()
                feld.fill(antwort)
                time.sleep(0.6)
                btn = page.query_selector(SEL_SENDEN_BTN)
                if btn:
                    btn.click()
                    time.sleep(1.2)
                    ergebnis["gesendet"] = True
                else:
                    # Keyboard-Fallback: Strg+Enter
                    feld.press("Control+Return")
                    time.sleep(1.2)
                    ergebnis["gesendet"] = True
        except Exception as exc:
            ergebnis["fehler"] = str(exc)

    return ergebnis


# ── Ausgabe ──────────────────────────────────────────────────────────────────

def drucke_ergebnis(i: int, e: dict) -> None:
    trennlinie = "─" * 65
    print(f"\n{trennlinie}")
    print(f"  [{i}] Anfrage von: {e['sender']}")
    print(f"      Inserat:     {e['titel']}")
    print(f"\n  Anfrage (Auszug):\n    {e['anfrage']}")
    print(f"\n  Generierte Antwort:")
    for zeile in e["antwort"].split("\n"):
        print(f"    {zeile}")
    if e["fehler"]:
        print(f"\n  ⚠ Fehler beim Senden: {e['fehler']}")
    elif e["gesendet"]:
        print("\n  ✓ Antwort gesendet")
    else:
        print("\n  ○ Nicht gesendet (Vorschaumodus)")


# ── Hauptprogramm ─────────────────────────────────────────────────────────────

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Kleinanzeigen.de – Automatische KI-Antworten auf Wohnungsanfragen",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--senden", action="store_true",
        help="Antworten tatsächlich absenden (Standard: nur Vorschau)",
    )
    parser.add_argument(
        "--max", type=int, default=20, metavar="N",
        help="Maximale Anzahl zu bearbeitender Konversationen (Standard: 20)",
    )
    parser.add_argument(
        "--sichtbar", action="store_true",
        help="Browser sichtbar anzeigen – nützlich bei CAPTCHA-Abfragen",
    )
    args = parser.parse_args()

    check_env()

    modus = "SENDEN" if args.senden else "VORSCHAU (kein Absenden)"
    print("=" * 65)
    print("  Kleinanzeigen.de – Automatische Antworten auf Wohnungsanfragen")
    print(f"  Modus: {modus}  |  Max. Anfragen: {args.max}")
    print("=" * 65)

    ergebnisse: list[dict] = []

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=not args.sichtbar)
        context = browser.new_context(
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            locale="de-DE",
            viewport={"width": 1280, "height": 900},
        )
        page = context.new_page()

        print("\n1. Einloggen...")
        if not login(page, EMAIL, PASSWORT):
            browser.close()
            sys.exit(1)

        print("\n2. Nachrichten laden...")
        konversationen = hole_konversationen(page)

        if not konversationen:
            print("  Keine Konversationen gefunden.")
            browser.close()
            return

        print(f"  {len(konversationen)} Konversation(en) gefunden.\n")
        print("3. Anfragen bearbeiten...")

        for konv in konversationen[: args.max]:
            ergebnis = bearbeite_konversation(page, konv, senden=args.senden)
            if ergebnis:
                ergebnisse.append(ergebnis)
                time.sleep(1.5)

        browser.close()

    # Zusammenfassung
    print("\n" + "=" * 65)
    if not ergebnisse:
        print("  Keine unbeantworteten Anfragen gefunden.")
    else:
        print(f"  {len(ergebnisse)} Antwort(en) generiert:")
        for i, e in enumerate(ergebnisse, 1):
            drucke_ergebnis(i, e)

        if not args.senden:
            print(
                f"\n  Tipp: Mit --senden werden die {len(ergebnisse)} "
                "Antwort(en) automatisch abgeschickt."
            )
    print("=" * 65)


if __name__ == "__main__":
    main()
