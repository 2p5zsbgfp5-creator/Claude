# Vermogensmonitor Premium — geïntegreerd standaardbestand

Eén schaalbaar Excel-standaardbestand voor de accountantspraktijk dat het actuele
vermogen van een klant (met of zonder partner) inzichtelijk maakt, corrigeert voor
**latente belastingen**, en toekomstgericht doorrekent via scenario's. Combineert
data uit de **aangifte inkomstenbelasting** en de **jaarrekening (BV/DGA)**.

**Hoofdbestand:** `Vermogensmonitor_Premium.xlsx` (v2 — premium groene uitstraling)
**Eerdere MVP:** `Vermogensmonitor_Concept.xlsx` (v1 — eenvoudiger, ter referentie)

Het Premium-bestand integreert het beste van drie aangeleverde bronbestanden:

| Bron | Overgenomen |
|------|-------------|
| Vermogensmonitor.xlsx | box 1/2/3-structuur, latente belastingen, KPI-set, **huisstijl-groen** |
| Private Wealth Dashboard.xlsx | data-ingang (IB + jaarrekening), scenario-events, adviesrapport-idee |
| Vermogensplanning_sjabloon.xlsx | meerjaren-prognose, scenario-delta, doelvermogen |

## Huisstijl

Donkergroen `#0C5F42` · accentgroen `#1CE175` · mint `#9AD9BE` · lichtgrijs `#F4F6F5`
(overgenomen uit het bronbestand). Dashboard geïnspireerd op de aangeleverde
referentie: icon-chip KPI-cards, doughnut, solvabiliteits-gauge, area-chart,
progress-bars, status-signalen en een donkergroene navigatierail.

## Architectuur — strikte laagscheiding (13 tabbladen)

| Laag | Tabbladen | Rol |
|------|-----------|-----|
| Navigatie | `START` | Klant, peildatum, partnerschakelaar, status |
| Input | `01_Klant`, `02_Input_IB`, `03_Input_JR`, `04_Vermogen`, `05_Historie`, `06_Latente` | Gestandaardiseerde brondata |
| Verwerking | `10_Berekeningen` | KPI's, ratio's, partner-consolidatie, latentie-correctie |
| Output | `20_Dashboard`, `21_Scenario` | Premium dashboard + financiële planning |
| Config/QA | `90_Config`, `91_Mapping`, `99_Controles` | Keuzelijsten, mapping, controles |

## Kernbeslissingen

- **Genormaliseerde feitentabel** (`04_Vermogen`): één regel per component met
  `Eigenaar` (Klant/Partner/Gezamenlijk), `Aandeel klant %`, `Box` en `Latente %`.
  Alle KPI's zijn `SUMIFS`/`SUMPRODUCT` hierover → schaalbaar zonder herbouw.
- **Latente belastingen** geïntegreerd via de kolom `Latente %` (bv. lijfrente 37%,
  AB-aandelen 31%) → gecorrigeerd nettovermogen ná belastinglatentie.
- **Eén partnerschakelaar** (START) stuurt het hele bestand; drie totalen
  (Klant / Partner / Gezamenlijk).
- **Ingebouwde reconciliaties** (`99_Controles`): box 3 monitor ↔ aangifte,
  balans-evenwicht jaarrekening, AB-waarde ↔ eigen vermogen BV, partnersplitsing.
- **Alleen breed-ondersteunde functies** (`SUMIFS`, `SUMPRODUCT`, `IF`, `IFERROR`) —
  geen `XLOOKUP`/spill-functies → maximale compatibiliteit.

## Dashboard (`20_Dashboard`)

- 8 KPI-cards met icon-chips: nettovermogen (+Δ t.o.v. vorig jaar), bezittingen,
  schulden, na latentie, privévermogen, onderneming (box 2), liquiditeit, LTV.
- Doughnut *vermogensverdeling*, solvabiliteits-*gauge*, *area-chart* ontwikkeling.
- Tabel belangrijkste posten met progress-bars, signalen-paneel.
- Print-klaar (liggend, 1 pagina).

## Scenario (`21_Scenario`)

- Meerjaren-projectie: basis / optimistisch / conservatief (aanpasbare aannames).
- Eenmalige scenario-events: verkoop onderneming, dividend, schenking, extra
  aflossing, overlijden (erfbelasting) — elk met effect op nettovermogen.

## Gebruik

1. Open in **Microsoft Excel** — formules herberekenen automatisch bij openen.
2. Vul op `START`: klantnaam, peildatum, fiscaal partner (Ja/Nee).
3. Vul de gele invoervelden op `01`–`06` (dropdowns via `90_Config`).
4. Controleer `99_Controles` (alles groen).
5. Bespreek `20_Dashboard` en `21_Scenario` met de klant.

## Reproduceren

```bash
python build_v2.py     # Premium (v2)
python build_vermogensmonitor.py   # Concept (v1)
```

> Alle bedragen zijn **fictieve voorbeelddata** (geanonimiseerd DGA-huishouden).
> Fiscale parameters op `90_Config` zijn voorbeeldwaarden — jaarlijks actualiseren.
> Automatische recalculatie via LibreOffice is in de bouwomgeving niet beschikbaar;
> formules zijn statisch gevalideerd (0 losse scheidingstekens, alle named ranges
> gedefinieerd, alle sheet-referenties geldig) en de kernbedragen + reconciliaties
> zijn handmatig geverifieerd. Excel berekent alles bij openen.
