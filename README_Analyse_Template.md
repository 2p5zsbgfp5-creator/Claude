# Analyse Template Handelsonderneming

Herbruikbaar, volledig formule-gedreven Excel-dashboard voor een
accountantskantoor. Inzetbaar voor vrijwel iedere handelsonderneming: de
gebruiker past alleen de **Instellingen**, de **Brondata** (grootboekexport) en
eventueel de **Mapping** aan — alle rapportages, KPI's, grafieken en
signaleringen rekenen automatisch door.

**Bestand:** `Analyse_Template_Handelsonderneming.xlsx`
**Generator:** `build_analyse_template.py` (reproduceerbaar; `python3 build_analyse_template.py`)

## Huisstijl
- Achtergrond wit · hoofdkleur donkergroen `#0C5F42` · accentgroen `#1CE175`
- Lettertype **Aptos**
- KPI-tegels, dynamische grafieken, doughnut/pie, databalken, voorwaardelijke
  opmaak (stoplichten), navigatieknoppen, professionele tabellen.

## Werkbladen (14 + verborgen rekenblad)
| # | Tab | Inhoud |
|---|-----|--------|
| 1 | **Start** | Welkomstscherm, bedrijfsgegevens, logo-placeholder, versie, accountant, navigatie |
| 2 | **Instellingen** | Algemene gegevens, periode-/vergelijkingsselectie, weergave, normeringen (invulvelden lichtgroen) |
| 3 | **Brondata** | Plakblad voor de grootboekexport (Exact Online e.d.); Saldo en Dashboardcategorie worden automatisch afgeleid |
| 4 | **Mapping** | Koppeling grootboekrekening → dashboardcategorie (vrij aanpasbaar) |
| 5 | **Dashboard** | 12 KPI-tegels (waarde, verschil vorig jaar, stoplicht) + 6 grafieken |
| 6 | **Resultatenrekening** | Dynamische W&V met vergelijking, verschil € / %, % van omzet |
| 7 | **Balans** | Activa/passiva, mutaties, balanscontrole, grafische verdeling |
| 8 | **Ratios** | 13 kengetallen met waarde, norm, stoplicht en toelichting |
| 9 | **Debiteuren** | Openstaande posten, ouderdomsanalyse, DSO, top-debiteuren |
| 10 | **Crediteuren** | Openstaande posten, DPO, top-crediteuren |
| 11 | **Voorraad** | Voorraadontwikkeling, rotatie, dagen, signaleringen |
| 12 | **Trendanalyse** | 24-maands ontwikkeling (omzet, kosten, resultaat, liquiditeit, eigen vermogen) + jaarvergelijking |
| 13 | **Signalering** | Automatische waarschuwingen t.o.v. de normeringen (stoplichten) |
| 14 | **Samenvatting** | Automatisch gegenereerde managementsamenvatting (sterke punten, risico's, adviespunten) |
| – | *Calc* | Verborgen rekenblad dat de grafieken voedt |

## Werkwijze
1. Vul de **Instellingen** in (bedrijfsnaam, boekjaar, periode, normen).
2. Plak de grootboekexport in **Brondata** (kolommen A t/m H en J t/m N).
3. Controleer/uitbreid de **Mapping**.
4. Klaar — het hele model rekent door. Wissel van periode of vergelijkingsjaar
   via de keuzelijsten in Instellingen.

## Techniek
- **Geen VBA, geen macro's.** Uitsluitend Excel-tabellen, gegevensvalidatie,
  voorwaardelijke opmaak, benoemde bereiken en dynamische formules.
- Het model rekent met **SUMIFS + INDEX/MATCH + IFERROR**. Bewuste keuze i.p.v.
  XLOOKUP/FILTER/SORT/UNIQUE: die spill-functies laten zich in de
  validatieomgeving niet betrouwbaar doorrekenen. SUMIFS/INDEX/MATCH werken
  identiek in Excel bij het plakken van nieuwe data en houden opgeslagen
  waarden correct. In Excel kan de mapping desgewenst naar XLOOKUP worden
  omgezet zonder de logica te wijzigen.
- Datamodel: periode 0 = beginbalans, periode 1–12 = maandmutaties. Balansposten
  cumuleren t/m de gekozen maand; W&V-posten tellen year-to-date t/m de gekozen
  maand. De balans sluit in **elke** periode (balanscontrole = 0).
- `fullCalcOnLoad` staat aan: Excel herberekent het volledige model bij openen.

## Voorbeelddata
Het bestand bevat een **fictieve, geanonimiseerde en sluitende** voorbeeld-
administratie (boekjaren 2024 en 2025) zodat alle schermen direct gevuld zijn.
Kerncijfers 2025: omzet € 4.200.000, brutomarge 30,0%, EBITDA € 331.000,
resultaat € 224.000, balanstotaal € 1.755.000, solvabiliteit 39,9%.

## Validatie
Alle 18.566 formule-nodes zijn onafhankelijk doorgerekend (Python `formulas`):
**0 formule-fouten**, en de kern- en balanscijfers komen exact overeen met de
verwachte waarden, ook bij wisselende periodeselectie.
