#!/usr/bin/env python3
"""
ImmobilienScout24 – Wohnungssuche über die offizielle REST API mit DSCR-Analyse.

Dieses Skript:
1. Verbindet sich mit der ImmobilienScout24 REST API (OAuth 1.0)
2. Sucht nach Kaufimmobilien basierend auf konfigurierbaren Parametern
3. Berechnet den DSCR (Debt Service Coverage Ratio) bei Vollfinanzierung
4. Filtert nach DSCR >= Schwellenwert
5. Erstellt eine Excel-Liste mit den besten Angeboten

Voraussetzung: API-Zugang bei https://api.immobilienscout24.de/
"""

import configparser
import json
import sys
import time
from pathlib import Path
from datetime import datetime

import requests
from requests_oauthlib import OAuth1

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ── Konfiguration laden ──────────────────────────────────────────────────

def lade_config(pfad: str = "config.ini") -> configparser.ConfigParser:
    """Lädt die Konfigurationsdatei."""
    config = configparser.ConfigParser()
    config_path = Path(pfad)

    if not config_path.exists():
        print(f"FEHLER: Konfigurationsdatei '{pfad}' nicht gefunden!")
        print("Bitte kopiere 'config.example.ini' zu 'config.ini' und trage deine API-Credentials ein.")
        print()
        print("  cp config.example.ini config.ini")
        print("  # Dann config.ini bearbeiten und Credentials eintragen")
        sys.exit(1)

    config.read(pfad, encoding="utf-8")
    return config


# ── OAuth 1.0 Authentifizierung ──────────────────────────────────────────

def erstelle_oauth(config: configparser.ConfigParser) -> OAuth1:
    """Erstellt OAuth1-Authentifizierung für die IS24 API."""
    api = config["immoscout24_api"]
    return OAuth1(
        client_key=api["consumer_key"],
        client_secret=api["consumer_secret"],
        resource_owner_key=api["access_token"],
        resource_owner_secret=api["access_token_secret"],
    )


# ── API-Anfragen ─────────────────────────────────────────────────────────

def suche_immobilien(config: configparser.ConfigParser, oauth: OAuth1) -> list:
    """
    Sucht Immobilien über die IS24 REST API.

    Endpunkt: /restapi/api/search/v2.0/search/region
    Dokumentation: https://api.immobilienscout24.de/our-apis/search/
    """
    base_url = config.get("api_settings", "base_url",
                          fallback="https://rest.immobilienscout24.de")
    page_size = config.getint("api_settings", "page_size", fallback=20)

    such = config["suchparameter"]
    suchart = such.get("suchart", "buy")
    immobilientyp = such.get("immobilientyp", "apartment")

    # Immobilientyp-Mapping für die API
    realestatetype_map = {
        ("buy", "apartment"): "apartmentbuy",
        ("buy", "house"): "housebuy",
        ("rent", "apartment"): "apartmentrent",
        ("rent", "house"): "houserent",
    }
    realestatetype = realestatetype_map.get(
        (suchart, immobilientyp), "apartmentbuy"
    )

    # Suchparameter zusammenbauen
    params = {
        "realestatetype": realestatetype,
        "geocodes": such.get("geocode", ""),
        "price.min": such.get("preis_min", ""),
        "price.max": such.get("preis_max", ""),
        "numberofrooms.min": such.get("zimmer_min", ""),
        "numberofrooms.max": such.get("zimmer_max", ""),
        "livingspace.min": such.get("flaeche_min", ""),
        "livingspace.max": such.get("flaeche_max", ""),
        "pagesize": str(page_size),
        "pagenumber": "1",
    }

    # Leere Parameter entfernen
    params = {k: v for k, v in params.items() if v}

    max_ergebnisse = config.getint("suchparameter", "max_ergebnisse", fallback=50)

    alle_angebote = []
    seite = 1

    while len(alle_angebote) < max_ergebnisse:
        params["pagenumber"] = str(seite)

        url = f"{base_url}/restapi/api/search/v2.0/search/region"
        headers = {
            "Accept": "application/json",
            "Content-Type": "application/json",
        }

        print(f"  Lade Seite {seite}...")

        try:
            response = requests.get(
                url, params=params, auth=oauth, headers=headers, timeout=30
            )
            response.raise_for_status()
        except requests.exceptions.HTTPError as e:
            if response.status_code == 401:
                print("FEHLER: Authentifizierung fehlgeschlagen!")
                print("Bitte prüfe deine API-Credentials in config.ini")
                sys.exit(1)
            elif response.status_code == 403:
                print("FEHLER: Zugriff verweigert!")
                print("Dein API-Zugang hat möglicherweise keine Berechtigung für die Such-API.")
                sys.exit(1)
            elif response.status_code == 429:
                print("  Rate-Limit erreicht. Warte 60 Sekunden...")
                time.sleep(60)
                continue
            else:
                print(f"FEHLER: HTTP {response.status_code} – {e}")
                sys.exit(1)
        except requests.exceptions.RequestException as e:
            print(f"FEHLER: Verbindungsfehler – {e}")
            sys.exit(1)

        daten = response.json()

        # Ergebnisse parsen
        ergebnisse = _parse_suchergebnisse(daten, realestatetype)

        if not ergebnisse:
            break

        alle_angebote.extend(ergebnisse)

        # Prüfen ob es weitere Seiten gibt
        paging = daten.get("resultlistResultList", {}).get("paging", {})
        total_pages = paging.get("numberOfPages", 1)

        if seite >= total_pages:
            break

        seite += 1
        time.sleep(1)  # Rate-Limiting respektieren

    return alle_angebote[:max_ergebnisse]


def _parse_suchergebnisse(daten: dict, realestatetype: str) -> list:
    """Parst die JSON-Antwort der IS24 Such-API."""
    angebote = []

    result_list = daten.get("resultlistResultList", {})
    entries = result_list.get("resultlistEntries", [])

    if not entries:
        return angebote

    for entry_group in entries:
        results = entry_group.get("resultlistEntry", [])
        if not isinstance(results, list):
            results = [results]

        for result in results:
            expose = result.get("resultlist.realEstate", {})
            if not expose:
                continue

            # Adresse extrahieren
            adresse = expose.get("address", {})
            stadt = adresse.get("city", "Unbekannt")
            plz = adresse.get("postcode", "")
            strasse = adresse.get("street", "")
            hausnr = adresse.get("houseNumber", "")

            # Preisdaten extrahieren
            preis_obj = expose.get("price", {})
            kaufpreis = preis_obj.get("value", 0)

            # Flächendaten
            flaeche = expose.get("livingSpace", 0)

            # Zimmerzahl
            zimmer = expose.get("numberOfRooms", 0)

            # Baujahr
            baujahr = expose.get("constructionYear", 0)

            # Etage
            etage = expose.get("floor", 0)

            # Titel
            titel = expose.get("title", "Kein Titel")

            # Expose-ID und URL
            expose_id = result.get("@id", "")

            # Mieteinnahmen (falls vermietet - bei Kapitalanlagen)
            # IS24 gibt calculatedPrice oder Kaltmiete bei vermieteten Objekten
            kaltmiete = expose.get("calculatedPrice", {}).get("rentPrice", {}).get(
                "netColdRent", 0
            )

            # Alternativ: Geschätzte Miete aus Flächenmultiplikator
            # Falls keine Mietdaten vorliegen, schätzen wir basierend auf dem Standort
            if not kaltmiete and flaeche > 0:
                # Konservative Schätzung: durchschnittliche Miete pro m²
                # basierend auf dem Kaufpreisfaktor
                if kaufpreis > 0:
                    # Grobe Schätzung: Bruttorendite von ca. 5-8% in günstigen Lagen
                    geschaetzte_jahresmiete = kaufpreis * 0.06
                    kaltmiete = geschaetzte_jahresmiete / 12

            angebot = {
                "titel": titel,
                "stadt": stadt,
                "plz": plz,
                "strasse": f"{strasse} {hausnr}".strip(),
                "flaeche_qm": float(flaeche) if flaeche else 0,
                "kaufpreis": float(kaufpreis) if kaufpreis else 0,
                "kaltmiete_monat": float(kaltmiete) if kaltmiete else 0,
                "zimmer": float(zimmer) if zimmer else 0,
                "baujahr": int(baujahr) if baujahr else 0,
                "etage": int(etage) if etage else 0,
                "expose_id": str(expose_id),
                "url": f"https://www.immobilienscout24.de/expose/{expose_id}",
                "miete_geschaetzt": kaltmiete == 0,
            }

            # Nur Angebote mit gültigen Daten aufnehmen
            if angebot["kaufpreis"] > 0 and angebot["flaeche_qm"] > 0:
                angebote.append(angebot)

    return angebote


def hole_expose_details(expose_id: str, config: configparser.ConfigParser,
                        oauth: OAuth1) -> dict:
    """Holt detaillierte Informationen zu einem einzelnen Exposé."""
    base_url = config.get("api_settings", "base_url",
                          fallback="https://rest.immobilienscout24.de")

    url = f"{base_url}/restapi/api/search/v2.0/expose/{expose_id}"
    headers = {"Accept": "application/json"}

    try:
        response = requests.get(url, auth=oauth, headers=headers, timeout=30)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"  Warnung: Expose {expose_id} konnte nicht geladen werden: {e}")
        return {}


# ── DSCR-Berechnung ──────────────────────────────────────────────────────

def berechne_dscr(kaufpreis: float, kaltmiete_monat: float,
                  flaeche_qm: float, config: configparser.ConfigParser) -> dict:
    """Berechnet den DSCR und alle Kennzahlen bei Vollfinanzierung."""
    fin = config["finanzierung"]
    zinssatz = fin.getfloat("zinssatz", fallback=0.038)
    tilgung = fin.getfloat("tilgung", fallback=0.01)
    nebenkosten_quote = fin.getfloat("nebenkosten_quote", fallback=0.20)
    instandhaltung_qm = fin.getfloat("instandhaltung_qm", fallback=10.0)
    mietausfall_quote = fin.getfloat("mietausfall_quote", fallback=0.02)

    # Jahresmiete
    jahresmiete_brutto = kaltmiete_monat * 12

    # Abzüge
    nicht_umlagefaehig = jahresmiete_brutto * nebenkosten_quote
    instandhaltung = flaeche_qm * instandhaltung_qm
    mietausfall = jahresmiete_brutto * mietausfall_quote

    # Netto-Mieteinnahmen (NOI)
    noi = jahresmiete_brutto - nicht_umlagefaehig - instandhaltung - mietausfall

    # Kapitaldienst bei Vollfinanzierung
    annuitaet_rate = zinssatz + tilgung
    kapitaldienst_jahr = kaufpreis * annuitaet_rate

    # DSCR
    dscr = noi / kapitaldienst_jahr if kapitaldienst_jahr > 0 else 0

    # Weitere Kennzahlen
    bruttorendite = (jahresmiete_brutto / kaufpreis * 100) if kaufpreis > 0 else 0
    nettorendite = (noi / kaufpreis * 100) if kaufpreis > 0 else 0
    preis_pro_qm = kaufpreis / flaeche_qm if flaeche_qm > 0 else 0
    miete_pro_qm = kaltmiete_monat / flaeche_qm if flaeche_qm > 0 else 0

    return {
        "jahresmiete_brutto": jahresmiete_brutto,
        "noi": noi,
        "kapitaldienst_jahr": kapitaldienst_jahr,
        "dscr": dscr,
        "bruttorendite": bruttorendite,
        "nettorendite": nettorendite,
        "preis_pro_qm": preis_pro_qm,
        "miete_pro_qm": miete_pro_qm,
    }


# ── Excel-Export ──────────────────────────────────────────────────────────

def erstelle_excel(angebote_mit_dscr: list, dateiname: str,
                   config: configparser.ConfigParser):
    """Erstellt eine formatierte Excel-Datei mit den Angeboten."""
    fin = config["finanzierung"]
    zinssatz = fin.getfloat("zinssatz", fallback=0.038)
    tilgung = fin.getfloat("tilgung", fallback=0.01)
    min_dscr = fin.getfloat("min_dscr", fallback=1.2)
    nebenkosten_quote = fin.getfloat("nebenkosten_quote", fallback=0.20)
    instandhaltung_qm = fin.getfloat("instandhaltung_qm", fallback=10.0)
    mietausfall_quote = fin.getfloat("mietausfall_quote", fallback=0.02)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Wohnungsangebote DSCR"

    # Styles
    header_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496",
                              fill_type="solid")
    good_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE",
                            fill_type="solid")
    great_fill = PatternFill(start_color="00B050", end_color="00B050",
                             fill_type="solid")
    great_font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    warn_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC",
                            fill_type="solid")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    currency_fmt = '#,##0 €'
    pct_fmt = '0.00%'
    dscr_fmt = '0.00'

    # Titel-Zeile
    ws.merge_cells("A1:R1")
    title_cell = ws["A1"]
    title_cell.value = (
        f"ImmobilienScout24 API – Wohnungssuche mit DSCR-Analyse  |  "
        f"Erstellt: {datetime.now().strftime('%d.%m.%Y %H:%M')}  |  "
        f"Zins: {zinssatz*100:.1f}%  Tilgung: {tilgung*100:.1f}%  "
        f"Min-DSCR: {min_dscr}"
    )
    title_cell.font = Font(name="Calibri", bold=True, size=13, color="2F5496")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # Parameter-Zeile
    ws.merge_cells("A2:R2")
    param_cell = ws["A2"]
    param_cell.value = (
        f"Vollfinanzierung (100% LTV)  |  "
        f"NK-Quote: {nebenkosten_quote*100:.0f}%  |  "
        f"Instandhaltung: {instandhaltung_qm:.0f} €/m²/Jahr  |  "
        f"Mietausfallrisiko: {mietausfall_quote*100:.0f}%  |  "
        f"Datenquelle: ImmobilienScout24 REST API"
    )
    param_cell.font = Font(name="Calibri", italic=True, size=10, color="666666")
    param_cell.alignment = Alignment(horizontal="center")

    # Spaltenüberschriften
    headers = [
        ("Nr.", 5),
        ("Titel", 35),
        ("Stadt", 15),
        ("PLZ", 8),
        ("Zimmer", 8),
        ("Fläche (m²)", 12),
        ("Baujahr", 9),
        ("Kaufpreis", 14),
        ("€/m²", 10),
        ("Kaltmiete/Mon.", 14),
        ("€/m² Miete", 11),
        ("Jahresmiete", 14),
        ("NOI/Jahr", 14),
        ("Kapitaldienst/J.", 15),
        ("DSCR", 9),
        ("Bruttorendite", 13),
        ("Nettorendite", 13),
        ("IS24-Link", 40),
    ]

    for col_idx, (header, width) in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = thin_border
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[4].height = 30

    # Daten einfügen
    for row_idx, angebot in enumerate(angebote_mit_dscr, 5):
        a = angebot
        k = a["kennzahlen"]
        nr = row_idx - 4

        values = [
            nr,
            a["titel"],
            a["stadt"],
            a["plz"],
            a["zimmer"],
            a["flaeche_qm"],
            a["baujahr"],
            a["kaufpreis"],
            k["preis_pro_qm"],
            a["kaltmiete_monat"],
            k["miete_pro_qm"],
            k["jahresmiete_brutto"],
            k["noi"],
            k["kapitaldienst_jahr"],
            k["dscr"],
            k["bruttorendite"] / 100,
            k["nettorendite"] / 100,
            a["url"],
        ]

        for col_idx, val in enumerate(values, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

            # Formate
            if col_idx in (8, 12, 13, 14):
                cell.number_format = currency_fmt
            elif col_idx in (9, 11):
                cell.number_format = '#,##0.00 €'
            elif col_idx == 10:
                cell.number_format = currency_fmt
            elif col_idx == 15:
                cell.number_format = dscr_fmt
                if val >= 1.5:
                    cell.fill = great_fill
                    cell.font = great_font
                elif val >= min_dscr:
                    cell.fill = good_fill
            elif col_idx in (16, 17):
                cell.number_format = pct_fmt
            elif col_idx == 18:
                cell.alignment = Alignment(horizontal="left", vertical="center")

        # Warnung bei geschätzter Miete
        if a.get("miete_geschaetzt"):
            cell = ws.cell(row=row_idx, column=10)
            cell.fill = warn_fill
            cell.comment = openpyxl.comments.Comment(
                "Miete geschätzt (keine Angabe im Inserat)", "System"
            )

        # Zeile alternierend einfärben
        if nr % 2 == 0:
            for col_idx in range(1, len(values) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                if col_idx != 15 or k["dscr"] < min_dscr:
                    if not cell.fill or cell.fill.start_color.rgb == "00000000":
                        cell.fill = PatternFill(start_color="F2F2F2",
                                                end_color="F2F2F2",
                                                fill_type="solid")

    # Autofilter
    last_row = 4 + len(angebote_mit_dscr)
    ws.auto_filter.ref = f"A4:R{last_row}"

    # Einfrieren
    ws.freeze_panes = "A5"

    wb.save(dateiname)
    return dateiname


# ── Hauptprogramm ─────────────────────────────────────────────────────────

def main():
    print("=" * 70)
    print("  ImmobilienScout24 API – Wohnungssuche mit DSCR-Analyse")
    print("=" * 70)

    # Konfiguration laden
    config = lade_config()

    fin = config["finanzierung"]
    zinssatz = fin.getfloat("zinssatz", fallback=0.038)
    tilgung = fin.getfloat("tilgung", fallback=0.01)
    min_dscr = fin.getfloat("min_dscr", fallback=1.2)

    print(f"  Vollfinanzierung | Zins: {zinssatz*100:.1f}% | Tilgung: {tilgung*100:.1f}%")
    print(f"  Mindest-DSCR: {min_dscr}")
    print("-" * 70)

    # OAuth-Authentifizierung
    print("  Authentifizierung...")
    oauth = erstelle_oauth(config)

    # Immobilien suchen
    print("  Suche Immobilien auf ImmobilienScout24...")
    angebote = suche_immobilien(config, oauth)

    if not angebote:
        print("\n  Keine Angebote gefunden. Bitte prüfe deine Suchparameter.")
        sys.exit(0)

    print(f"  {len(angebote)} Angebote gefunden.")
    print("-" * 70)

    # DSCR berechnen
    print("  Berechne DSCR für alle Angebote...")
    ergebnisse = []
    for angebot in angebote:
        kennzahlen = berechne_dscr(
            angebot["kaufpreis"],
            angebot["kaltmiete_monat"],
            angebot["flaeche_qm"],
            config,
        )
        ergebnisse.append({**angebot, "kennzahlen": kennzahlen})

    # Nach DSCR filtern und sortieren
    gefiltert = [e for e in ergebnisse if e["kennzahlen"]["dscr"] >= min_dscr]
    gefiltert.sort(key=lambda x: x["kennzahlen"]["dscr"], reverse=True)

    # Top-Ergebnisse
    top = gefiltert[:20]

    print(f"\n  {len(gefiltert)} von {len(angebote)} Angeboten mit DSCR >= {min_dscr}")

    if not top:
        print("  Keine Angebote erfüllen das DSCR-Kriterium.")
        print("  Tipp: Senke den min_dscr-Wert oder erweitere den Suchbereich.")

        # Trotzdem alle exportieren (sortiert nach DSCR)
        ergebnisse.sort(key=lambda x: x["kennzahlen"]["dscr"], reverse=True)
        top = ergebnisse[:20]
        print(f"  Exportiere stattdessen die Top 20 nach DSCR (ungefiltert).")

    print()
    for i, a in enumerate(top[:10], 1):
        k = a["kennzahlen"]
        miete_hinweis = " *" if a.get("miete_geschaetzt") else ""
        print(f"  {i:2d}. {a['titel']}")
        print(f"      {a['stadt']} | Kaufpreis: {a['kaufpreis']:>10,.0f} €  |  "
              f"Kaltmiete: {a['kaltmiete_monat']:>5,.0f} €/Mon{miete_hinweis}  |  "
              f"DSCR: {k['dscr']:.2f}  |  "
              f"Brutto: {k['bruttorendite']:.1f}%")

    if any(a.get("miete_geschaetzt") for a in top[:10]):
        print("\n  * = Miete geschätzt (keine Mietangabe im Inserat)")

    # Excel erstellen
    dateiname = f"wohnungen_dscr_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    erstelle_excel(top, dateiname, config)
    print(f"\n  Excel-Datei erstellt: {dateiname}")
    print("=" * 70)


if __name__ == "__main__":
    main()
