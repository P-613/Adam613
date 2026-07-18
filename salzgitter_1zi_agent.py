#!/usr/bin/env python3
"""
Salzgitter 1-Zimmer-Wohnungen – Kalkulations-Agent (Jobcenter-Vermietung)

Der Agent:
1. Scrapt Kleinanzeigen.de nach 1-Zi-Wohnungen zum Kauf in Salzgitter
2. Berechnet vollständige Kaufkalkulation je Objekt:
   - Kaufnebenkosten (Niedersachsen: GrESt 5% + Notar 2% + ggf. Makler 3,57%)
   - Finanzierung & DSCR bei Vollfinanzierung inkl. Nebenkosten
   - Monatlicher Cashflow
   - Break-Even-Analyse (Monate bis die NK amortisiert sind)
3. Filtert: nur Objekte mit DSCR >= 1,3
4. Exportiert als formatierte Excel-Datei mit Ampelfarben

Verwendung:
  python salzgitter_1zi_agent.py                         # Vorschau im Terminal
  python salzgitter_1zi_agent.py --sichtbar              # Browser sichtbar
  python salzgitter_1zi_agent.py --ausgabe meine.xlsx    # Dateiname wählen
  python salzgitter_1zi_agent.py --max 30 --seiten 3    # Mehr Objekte
"""

import re
import sys
import time
import argparse
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout


# ══════════════════════════════════════════════════════════════════════════════
# PARAMETER  (alle zentral anpassbar)
# ══════════════════════════════════════════════════════════════════════════════

# ── Mieteinnahmen: Jobcenter Salzgitter (Bürgergeld / neue Grundsicherung) ──
BRUTTOWARMMIETE     = 441.0   # €/Mon – Gesamtzahlung Jobcenter (inkl. Heizung)
HEIZKOSTEN_MONAT    = 70.0    # €/Mon – Heizkosten (Landlord zahlt, Schätzwert)
NK_NICHT_UMLAGEFAEHIG = 25.0  # €/Mon – nicht umlagefähige Betriebskosten
# → effektive Kaltmiete, die für DSCR/Cashflow zählt:
KALTMIETE_MONAT = BRUTTOWARMMIETE - HEIZKOSTEN_MONAT - NK_NICHT_UMLAGEFAEHIG
# = 346 €/Monat

# ── Finanzierung ─────────────────────────────────────────────────────────────
ZINSSATZ      = 0.038   # 3,8 % Sollzins (Marktniveau 2026)
TILGUNG       = 0.010   # 1,0 % anfängliche Tilgung
ANNUITAET_RATE = ZINSSATZ + TILGUNG   # = 4,8 %

# ── Kaufnebenkosten Niedersachsen ────────────────────────────────────────────
GRUNDERWERBSTEUER = 0.050   # 5,0 %
NOTAR_GRUNDBUCH   = 0.020   # 2,0 %
MAKLER_PROVISION  = 0.0357  # 3,57 % (nur wenn Makler)

# ── Laufende Bewirtschaftungskosten ──────────────────────────────────────────
INSTANDHALTUNG_QM_JAHR = 10.0   # €/m²/Jahr
MIETAUSFALL_QUOTE      = 0.02   # 2 %
HAUSVERWALTUNG_MONAT   = 25.0   # €/Monat

# ── Filter ───────────────────────────────────────────────────────────────────
MIN_DSCR = 1.3

# ── Suche ────────────────────────────────────────────────────────────────────
# Kleinanzeigen-Kategorie 196 = Wohnungen kaufen, Salzgitter
SUCHE_URLS = [
    "https://www.kleinanzeigen.de/s-wohnung-kaufen/salzgitter/c196l8473",
    "https://www.kleinanzeigen.de/s-wohnung-kaufen/salzgitter/c196",
]
# Paginierung: ?pageNum=2 …
PAUSE_SEKUNDEN = 1.5   # Pause zwischen Seitenaufrufen


# ══════════════════════════════════════════════════════════════════════════════
# BERECHNUNG
# ══════════════════════════════════════════════════════════════════════════════

def berechne(kaufpreis: float, flaeche_qm: float, hat_makler: bool = False) -> dict:
    """Vollständige Kaufkalkulation für eine Wohnung."""

    # Kaufnebenkosten
    nk_quote   = GRUNDERWERBSTEUER + NOTAR_GRUNDBUCH + (MAKLER_PROVISION if hat_makler else 0)
    nebenkosten = kaufpreis * nk_quote
    gesamt      = kaufpreis + nebenkosten   # Gesamtdarlehen (Vollfinanzierung)

    # Finanzierung
    jahres_annuitaet  = gesamt * ANNUITAET_RATE
    monats_annuitaet  = jahres_annuitaet / 12
    monats_zinsen     = gesamt * ZINSSATZ / 12
    monats_tilgung    = gesamt * TILGUNG   / 12

    # NOI
    jahresmiete         = KALTMIETE_MONAT * 12
    mietausfall_jahr    = jahresmiete * MIETAUSFALL_QUOTE
    instandhaltung_jahr = flaeche_qm * INSTANDHALTUNG_QM_JAHR
    verwaltung_jahr     = HAUSVERWALTUNG_MONAT * 12
    noi = jahresmiete - mietausfall_jahr - instandhaltung_jahr - verwaltung_jahr

    # DSCR
    dscr = noi / jahres_annuitaet if jahres_annuitaet > 0 else 0.0

    # Monatlicher Cashflow
    cashflow_monat = (
        KALTMIETE_MONAT
        - monats_annuitaet
        - HAUSVERWALTUNG_MONAT
        - instandhaltung_jahr / 12
        - mietausfall_jahr    / 12
    )

    # Break-Even: Monate bis die Kaufnebenkosten (Eigenkapitaleinsatz) amortisiert sind
    if cashflow_monat > 0:
        be_monate = nebenkosten / cashflow_monat
        be_jahre  = be_monate / 12
    else:
        be_monate = float("inf")
        be_jahre  = float("inf")

    return {
        # Preise
        "kaufpreis":           kaufpreis,
        "flaeche_qm":          flaeche_qm,
        "kaufpreis_pro_qm":    kaufpreis / flaeche_qm if flaeche_qm else 0,
        # Nebenkosten
        "grunderwerbsteuer":   kaufpreis * GRUNDERWERBSTEUER,
        "notar_grundbuch":     kaufpreis * NOTAR_GRUNDBUCH,
        "makler":              kaufpreis * MAKLER_PROVISION if hat_makler else 0.0,
        "nebenkosten_gesamt":  nebenkosten,
        "nebenkosten_quote":   nk_quote * 100,
        "gesamtinvestition":   gesamt,
        # Finanzierung
        "jahres_annuitaet":    jahres_annuitaet,
        "monats_annuitaet":    monats_annuitaet,
        "monats_zinsen":       monats_zinsen,
        "monats_tilgung":      monats_tilgung,
        # Ertrag
        "kaltmiete_monat":     KALTMIETE_MONAT,
        "bruttowarmmiete":     BRUTTOWARMMIETE,
        "miete_pro_qm":        KALTMIETE_MONAT / flaeche_qm if flaeche_qm else 0,
        "noi":                 noi,
        # Kennzahlen
        "dscr":                dscr,
        "cashflow_monat":      cashflow_monat,
        "cashflow_jahr":       cashflow_monat * 12,
        "be_monate":           be_monate,
        "be_jahre":            be_jahre,
        "bruttorendite":       (KALTMIETE_MONAT * 12 / kaufpreis * 100) if kaufpreis else 0,
        "nettorendite":        (noi / kaufpreis * 100) if kaufpreis else 0,
        "hat_makler":          hat_makler,
    }


# ══════════════════════════════════════════════════════════════════════════════
# SCRAPER
# ══════════════════════════════════════════════════════════════════════════════

def _preis_aus_text(text: str) -> float | None:
    """Extrahiert ersten Zahlenwert aus Preistext ('38.500 €' → 38500.0)."""
    bereinigt = text.replace(".", "").replace(",", ".").replace(" ", "")
    treffer = re.search(r"(\d+(?:\.\d+)?)", bereinigt)
    return float(treffer.group(1)) if treffer else None


def _flaeche_aus_text(text: str) -> float | None:
    """Extrahiert Quadratmeter aus Text ('42 m²' → 42.0)."""
    treffer = re.search(r"(\d+(?:[.,]\d+)?)\s*m[²2]", text, re.IGNORECASE)
    if treffer:
        return float(treffer.group(1).replace(",", "."))
    return None


def _zimmer_aus_text(text: str) -> float | None:
    """Extrahiert Zimmeranzahl ('1,5 Zimmer' → 1.5)."""
    treffer = re.search(r"(\d+(?:[.,]\d+)?)\s*Zimmer", text, re.IGNORECASE)
    if treffer:
        return float(treffer.group(1).replace(",", "."))
    return None


def _cookie_wegklicken(page) -> None:
    for sel in [
        "#gdpr-banner-accept",
        "[data-testid='gdpr-banner-accept']",
        "button[id*='accept-all']",
        "button[class*='accept-all']",
    ]:
        try:
            page.click(sel, timeout=2500)
            time.sleep(0.4)
            return
        except PlaywrightTimeout:
            pass


def _extrahiere_karten(page) -> list[dict]:
    """Liest alle Inserate-Karten von der aktuellen Suchergebnisseite."""
    karten: list[dict] = []

    # Karten-Container
    items = page.query_selector_all(
        "article.aditem, li.ad-listitem article, [data-testid='ad-item']"
    )
    if not items:
        items = page.query_selector_all("li[data-adid]")

    for item in items:
        try:
            # Link + URL
            link_el = item.query_selector("a[href]")
            if not link_el:
                continue
            href = link_el.get_attribute("href") or ""
            url = href if href.startswith("http") else "https://www.kleinanzeigen.de" + href

            # Titel
            titel_el = item.query_selector(
                "h2.text-module-begin, .ellipsis, [class*='title'], h2"
            )
            titel = titel_el.inner_text().strip() if titel_el else "–"

            # Preis
            preis_el = item.query_selector(
                ".aditem-main--price, [class*='price'], strong"
            )
            preis_text = preis_el.inner_text().strip() if preis_el else ""
            preis = _preis_aus_text(preis_text)

            # Details (Fläche, Zimmer)
            details_text = item.inner_text()
            flaeche = _flaeche_aus_text(details_text)
            zimmer  = _zimmer_aus_text(details_text)

            # Ort
            ort_el = item.query_selector(
                ".aditem-main--tagbox, [class*='location'], .locality"
            )
            ort = ort_el.inner_text().strip() if ort_el else "Salzgitter"

            karten.append({
                "url":    url,
                "titel":  titel,
                "preis":  preis,
                "flaeche": flaeche,
                "zimmer": zimmer,
                "ort":    ort,
            })
        except Exception:
            continue

    return karten


def _detail_seite_anreichern(page, karte: dict) -> dict:
    """Öffnet die Detailseite und ergänzt fehlende Angaben (Fläche, Preis, Makler)."""
    try:
        page.goto(karte["url"], wait_until="domcontentloaded", timeout=20000)
        time.sleep(1.0)
    except PlaywrightTimeout:
        return karte

    seiten_text = page.inner_text("body")

    if not karte["preis"]:
        preis_el = page.query_selector(
            "h2.boxedarticle--price, [class*='price-header'], [class*='asking-price'], "
            "[data-testid='price']"
        )
        if preis_el:
            karte["preis"] = _preis_aus_text(preis_el.inner_text())
        else:
            karte["preis"] = _preis_aus_text(seiten_text)

    if not karte["flaeche"]:
        karte["flaeche"] = _flaeche_aus_text(seiten_text)

    if not karte["zimmer"]:
        karte["zimmer"] = _zimmer_aus_text(seiten_text)

    # Makler-Erkennung (Provision / Courtage im Text)
    karte["hat_makler"] = bool(
        re.search(r"(provision|courtage|makler|maklerprovision)", seiten_text, re.IGNORECASE)
    )

    return karte


def scrape(max_objekte: int, max_seiten: int, sichtbar: bool) -> list[dict]:
    """Scrapt Kleinanzeigen und gibt rohe Inserate-Dicts zurück."""
    gefunden: list[dict] = []

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=not sichtbar)
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

        # Einstiegs-URL finden
        start_url = None
        for url in SUCHE_URLS:
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=20000)
                _cookie_wegklicken(page)
                time.sleep(1.5)
                karten = _extrahiere_karten(page)
                if karten:
                    start_url = url
                    print(f"  ✓ Suchergebnisse gefunden: {url}")
                    break
                else:
                    print(f"  – Keine Treffer bei: {url}")
            except PlaywrightTimeout:
                print(f"  – Timeout: {url}")

        if not start_url:
            print("  Fehler: Keine Suchergebnisse gefunden. Bitte --sichtbar probieren.")
            browser.close()
            return []

        # Seiten durchlaufen
        for seite in range(1, max_seiten + 1):
            if seite > 1:
                seiten_url = f"{start_url}?pageNum={seite}"
                try:
                    page.goto(seiten_url, wait_until="domcontentloaded", timeout=20000)
                    time.sleep(PAUSE_SEKUNDEN)
                except PlaywrightTimeout:
                    print(f"  Seite {seite}: Timeout – Abbruch.")
                    break

            karten = _extrahiere_karten(page)
            if not karten:
                print(f"  Seite {seite}: Keine weiteren Inserate.")
                break

            print(f"  Seite {seite}: {len(karten)} Inserate gefunden.")

            for karte in karten:
                if len(gefunden) >= max_objekte:
                    break

                # Nur 1-Zimmer-Wohnungen (oder unbekannte Zimmeranzahl → prüfen wir)
                if karte["zimmer"] is not None and karte["zimmer"] > 1.5:
                    continue

                # Detailseite aufrufen wenn Daten fehlen
                if not karte["preis"] or not karte["flaeche"]:
                    karte = _detail_seite_anreichern(page, karte)
                    time.sleep(PAUSE_SEKUNDEN)
                else:
                    karte.setdefault("hat_makler", False)

                # Ungültige Objekte überspringen
                if not karte["preis"] or not karte["flaeche"]:
                    continue
                if karte["preis"] < 5_000 or karte["preis"] > 500_000:
                    continue
                if karte["flaeche"] < 10 or karte["flaeche"] > 80:
                    continue

                gefunden.append(karte)

            if len(gefunden) >= max_objekte:
                break

            time.sleep(PAUSE_SEKUNDEN)

        browser.close()

    return gefunden


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL-EXPORT
# ══════════════════════════════════════════════════════════════════════════════

def erstelle_excel(objekte: list[dict], dateiname: str) -> str:
    """Erstellt eine formatierte Excel-Datei mit allen kalkulierten Objekten."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Kalkulation 1-Zi Salzgitter"

    # ── Styles ───────────────────────────────────────────────────────────────
    def fett(farbe="000000", size=11):
        return Font(name="Calibri", bold=True, color=farbe, size=size)

    def fuell(hex_farbe):
        return PatternFill(start_color=hex_farbe, end_color=hex_farbe, fill_type="solid")

    BLAU_DUNKEL  = "1F4E79"
    GRUEN_HELL   = "C6EFCE"
    GRUEN_DUNKEL = "00B050"
    GELB_HELL    = "FFEB9C"
    ROT_HELL     = "FFC7CE"
    GRAU_HELL    = "F2F2F2"

    rand = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    mitte = Alignment(horizontal="center", vertical="center")
    links = Alignment(horizontal="left",   vertical="center")

    # ── Titelzeile ────────────────────────────────────────────────────────────
    ws.merge_cells("A1:T1")
    ws["A1"].value = (
        f"Salzgitter – 1-Zimmer-Wohnungen Kaufkalkulation  |  "
        f"Jobcenter-Miete: {BRUTTOWARMMIETE:.0f} € Bruttowarm / "
        f"{KALTMIETE_MONAT:.0f} € Kalt  |  "
        f"DSCR-Filter ≥ {MIN_DSCR}  |  Erstellt: {datetime.now().strftime('%d.%m.%Y')}"
    )
    ws["A1"].font = Font(name="Calibri", bold=True, size=12, color=BLAU_DUNKEL)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # ── Parameter-Zeile ───────────────────────────────────────────────────────
    ws.merge_cells("A2:T2")
    ws["A2"].value = (
        f"Vollfinanzierung inkl. NK  |  Zins {ZINSSATZ*100:.1f}%  "
        f"Tilgung {TILGUNG*100:.1f}%  |  "
        f"GrESt {GRUNDERWERBSTEUER*100:.1f}%  Notar {NOTAR_GRUNDBUCH*100:.1f}%  "
        f"Makler {MAKLER_PROVISION*100:.2f}%  |  "
        f"Instandhaltung {INSTANDHALTUNG_QM_JAHR:.0f} €/m²/J  "
        f"Mietausfall {MIETAUSFALL_QUOTE*100:.0f}%  Verwaltung {HAUSVERWALTUNG_MONAT:.0f} €/Mon"
    )
    ws["A2"].font = Font(name="Calibri", italic=True, size=9, color="666666")
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 16

    # ── Spalten ───────────────────────────────────────────────────────────────
    headers = [
        # (Bezeichnung,      Breite)
        ("Nr.",              5),
        ("Titel",           32),
        ("Ort",             18),
        ("m²",               7),
        ("Kaufpreis",       13),
        ("€/m² Kauf",       10),
        ("GrESt",           10),
        ("Notar/GB",        10),
        ("Makler",           9),
        ("NK gesamt",       11),
        ("Gesamtinvest.",   13),
        ("Rate/Mon.",       11),
        ("davon Zins",      10),
        ("davon Tilgung",   12),
        ("Kaltmiete/Mon.",  13),
        ("NOI/Jahr",        11),
        ("DSCR",             8),
        ("Cashflow/Mon.",   13),
        ("Cashflow/Jahr",   13),
        ("Break-Even (J.)", 14),
    ]

    for col, (name, breite) in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col, value=name)
        cell.font      = fett("FFFFFF")
        cell.fill      = fuell(BLAU_DUNKEL)
        cell.alignment = mitte
        cell.border    = rand
        ws.column_dimensions[get_column_letter(col)].width = breite

    ws.row_dimensions[4].height = 32
    ws.freeze_panes = "A5"

    # ── Datenzeilen ───────────────────────────────────────────────────────────
    for zeile_nr, obj in enumerate(objekte, 1):
        row = zeile_nr + 4
        k   = obj["kalk"]
        alternierend = (zeile_nr % 2 == 0)

        def zelle(col, wert, fmt=None, align=mitte):
            c = ws.cell(row=row, column=col, value=wert)
            c.border    = rand
            c.alignment = align
            if alternierend:
                c.fill = fuell(GRAU_HELL)
            if fmt:
                c.number_format = fmt
            return c

        zelle(1,  zeile_nr)
        zelle(2,  obj["titel"],  align=links)
        zelle(3,  obj["ort"],    align=links)
        zelle(4,  k["flaeche_qm"])
        zelle(5,  k["kaufpreis"],           '#,##0 "€"')
        zelle(6,  k["kaufpreis_pro_qm"],    '#,##0 "€"')
        zelle(7,  k["grunderwerbsteuer"],   '#,##0 "€"')
        zelle(8,  k["notar_grundbuch"],     '#,##0 "€"')
        zelle(9,  k["makler"] or None,      '#,##0 "€"')
        zelle(10, k["nebenkosten_gesamt"],  '#,##0 "€"')
        zelle(11, k["gesamtinvestition"],   '#,##0 "€"')
        zelle(12, k["monats_annuitaet"],    '#,##0 "€"')
        zelle(13, k["monats_zinsen"],       '#,##0 "€"')
        zelle(14, k["monats_tilgung"],      '#,##0 "€"')
        zelle(15, k["kaltmiete_monat"],     '#,##0 "€"')
        zelle(16, k["noi"],                 '#,##0 "€"')

        # DSCR-Zelle mit Ampelfarbe
        dscr_cell = zelle(17, round(k["dscr"], 2), '0.00')
        if k["dscr"] >= 1.5:
            dscr_cell.fill = fuell(GRUEN_DUNKEL)
            dscr_cell.font = fett("FFFFFF")
        elif k["dscr"] >= MIN_DSCR:
            dscr_cell.fill = fuell(GRUEN_HELL)
            dscr_cell.font = fett()
        else:
            dscr_cell.fill = fuell(ROT_HELL)

        # Cashflow mit Ampelfarbe
        cf_cell = zelle(18, k["cashflow_monat"], '#,##0 "€"')
        if k["cashflow_monat"] >= 50:
            cf_cell.fill = fuell(GRUEN_HELL)
        elif k["cashflow_monat"] < 0:
            cf_cell.fill = fuell(ROT_HELL)
        else:
            cf_cell.fill = fuell(GELB_HELL)

        zelle(19, k["cashflow_jahr"], '#,##0 "€"')

        # Break-Even
        if k["be_jahre"] == float("inf"):
            be_val = "negativ"
            be_cell = zelle(20, be_val)
            be_cell.fill = fuell(ROT_HELL)
        else:
            be_cell = zelle(20, round(k["be_jahre"], 1), '0.0 "J."')
            if k["be_jahre"] <= 10:
                be_cell.fill = fuell(GRUEN_HELL)
            elif k["be_jahre"] <= 20:
                be_cell.fill = fuell(GELB_HELL)
            else:
                be_cell.fill = fuell(ROT_HELL)

    # Autofilter
    letzteZeile = 4 + len(objekte)
    ws.auto_filter.ref = f"A4:T{letzteZeile}"

    # ── Legende (zweites Tabellenblatt) ──────────────────────────────────────
    legend = wb.create_sheet("Legende & Parameter")
    legend_daten = [
        ("Parameter",                "Wert",                            "Erläuterung"),
        ("Bruttowarmmiete Jobcenter", f"{BRUTTOWARMMIETE:.0f} €/Mon.",  "Jobcenter Salzgitter – Bürgergeld/Grundsicherung 1 Person"),
        ("− Heizkosten (Schätzung)", f"{HEIZKOSTEN_MONAT:.0f} €/Mon.", "Landlord zahlt Heizung; Schätzwert – je nach Energieträger anpassen"),
        ("− Nicht umlagef. NK",      f"{NK_NICHT_UMLAGEFAEHIG:.0f} €/Mon.", "Verwaltungs-/Wartungsanteile die nicht auf Mieter umgelegt werden"),
        ("= Kaltmiete (DSCR-Basis)", f"{KALTMIETE_MONAT:.0f} €/Mon.", "Diese Kaltmiete geht in DSCR und Cashflow-Berechnung ein"),
        ("",                         "",                                ""),
        ("Zinssatz",                 f"{ZINSSATZ*100:.1f} %",          "Sollzins Annuitätendarlehen"),
        ("Tilgung (anfänglich)",     f"{TILGUNG*100:.1f} %",           "Anfängliche Tilgungsrate"),
        ("",                         "",                                ""),
        ("Grunderwerbsteuer",        f"{GRUNDERWERBSTEUER*100:.1f} %", "Niedersachsen (Stand 2026)"),
        ("Notar + Grundbuch",        f"{NOTAR_GRUNDBUCH*100:.1f} %",   "Schätzwert – konkrete Angebote einholen"),
        ("Maklerprovision",          f"{MAKLER_PROVISION*100:.2f} %",  "Nur wenn Makler beteiligt (wird automatisch erkannt)"),
        ("",                         "",                                ""),
        ("Instandhaltung",           f"{INSTANDHALTUNG_QM_JAHR:.0f} €/m²/Jahr", "Rücklage für Reparaturen und Modernisierungen"),
        ("Mietausfallrisiko",        f"{MIETAUSFALL_QUOTE*100:.0f} %", "Puffer für Leerstand/Mietausfall"),
        ("Hausverwaltung",           f"{HAUSVERWALTUNG_MONAT:.0f} €/Mon.", "Verwaltungskosten"),
        ("",                         "",                                ""),
        ("DSCR-Mindestgrenze",       f"≥ {MIN_DSCR}",                  "Nur Objekte ≥ 1,3 werden ausgegeben"),
        ("",                         "",                                ""),
        ("DSCR-Ampel",               "",                                ""),
        ("≥ 1,5",                    "Dunkelgrün",                     "Sehr gut – hoher Sicherheitspuffer"),
        (f"≥ {MIN_DSCR}",            "Hellgrün",                       "Gut – Mindestanforderung erfüllt"),
        ("< 1,3",                    "Rot",                            "Nicht empfohlen"),
        ("",                         "",                                ""),
        ("Cashflow-Ampel",           "",                                ""),
        ("≥ +50 €/Mon.",             "Hellgrün",                       "Positiver Cashflow"),
        ("0–50 €/Mon.",              "Gelb",                           "Knapper Cashflow"),
        ("< 0 €/Mon.",               "Rot",                            "Negativer Cashflow"),
    ]
    for r, (a, b, c) in enumerate(legend_daten, 1):
        legend.cell(row=r, column=1, value=a).font = Font(bold=(r == 1))
        legend.cell(row=r, column=2, value=b)
        legend.cell(row=r, column=3, value=c)
    legend.column_dimensions["A"].width = 28
    legend.column_dimensions["B"].width = 20
    legend.column_dimensions["C"].width = 60

    wb.save(dateiname)
    return dateiname


# ══════════════════════════════════════════════════════════════════════════════
# HAUPTPROGRAMM
# ══════════════════════════════════════════════════════════════════════════════

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Salzgitter 1-Zimmer-Wohnungen – Kalkulations-Agent (Jobcenter)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--ausgabe",  default="salzgitter_1zi_kalkulation.xlsx",
                        help="Dateiname der Excel-Ausgabe (Standard: salzgitter_1zi_kalkulation.xlsx)")
    parser.add_argument("--max",      type=int, default=50,
                        help="Max. zu verarbeitende Inserate (Standard: 50)")
    parser.add_argument("--seiten",   type=int, default=5,
                        help="Anzahl Suchergebnis-Seiten (Standard: 5)")
    parser.add_argument("--sichtbar", action="store_true",
                        help="Browser sichtbar anzeigen (hilfreich bei CAPTCHA)")
    parser.add_argument("--alle",     action="store_true",
                        help="Alle Objekte ausgeben, nicht nur DSCR ≥ 1,3")
    args = parser.parse_args()

    print("=" * 70)
    print("  Salzgitter – 1-Zimmer-Wohnungen Kalkulations-Agent")
    print(f"  Jobcenter-Miete: {BRUTTOWARMMIETE:.0f} € bruttowarm / {KALTMIETE_MONAT:.0f} € kalt")
    print(f"  Mindest-DSCR: {MIN_DSCR}  |  Vollfinanzierung inkl. Nebenkosten")
    print("=" * 70)

    # ── Scraping ──────────────────────────────────────────────────────────────
    print(f"\n[1/3] Kleinanzeigen.de wird durchsucht (max. {args.seiten} Seiten)...")
    rohdaten = scrape(args.max, args.seiten, args.sichtbar)

    if not rohdaten:
        print("\nKeine Inserate gefunden. Tipps:")
        print("  • Mit --sichtbar starten und ggf. CAPTCHA manuell lösen")
        print("  • Internetverbindung prüfen")
        sys.exit(1)

    print(f"\n  {len(rohdaten)} Inserate mit Preis + Fläche gefunden.\n")

    # ── Kalkulation ───────────────────────────────────────────────────────────
    print("[2/3] Kalkulation läuft...")
    alle_objekte: list[dict] = []
    for ins in rohdaten:
        kalk = berechne(ins["preis"], ins["flaeche"], ins.get("hat_makler", False))
        alle_objekte.append({**ins, "kalk": kalk})

    # Sortieren nach DSCR absteigend
    alle_objekte.sort(key=lambda x: x["kalk"]["dscr"], reverse=True)

    # Filtern
    export_objekte = alle_objekte if args.alle else [
        o for o in alle_objekte if o["kalk"]["dscr"] >= MIN_DSCR
    ]

    # ── Terminalmeldung ───────────────────────────────────────────────────────
    gesamt = len(alle_objekte)
    bestanden = len(export_objekte)
    print(f"\n  {bestanden} von {gesamt} Objekten erfüllen DSCR ≥ {MIN_DSCR}:\n")

    for i, obj in enumerate(export_objekte, 1):
        k = obj["kalk"]
        be_str = (
            f"{k['be_jahre']:.1f} J." if k["be_jahre"] != float("inf") else "negativ"
        )
        print(
            f"  {i:>2}. {obj['titel'][:50]:<50}\n"
            f"      Kaufpreis: {k['kaufpreis']:>8,.0f} €  |  "
            f"m²: {k['flaeche_qm']:.0f}  |  "
            f"Rate: {k['monats_annuitaet']:>6,.0f} €/Mon  |  "
            f"DSCR: {k['dscr']:.2f}  |  "
            f"CF: {k['cashflow_monat']:+.0f} €/Mon  |  "
            f"Break-Even: {be_str}"
        )

    # ── Excel-Export ──────────────────────────────────────────────────────────
    print(f"\n[3/3] Excel-Export...")
    datei = erstelle_excel(export_objekte, args.ausgabe)
    print(f"\n  ✓ Datei erstellt: {datei}")
    print(f"  {bestanden} Objekte | Mindest-DSCR {MIN_DSCR} | "
          f"Kaltmiete {KALTMIETE_MONAT:.0f} €/Mon | Bruttowarm {BRUTTOWARMMIETE:.0f} €/Mon")
    print("=" * 70)


if __name__ == "__main__":
    main()
