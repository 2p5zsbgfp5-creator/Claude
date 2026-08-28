> **Let op — deze repo bevat twee losse projecten:**
> 1. **Vermogensmonitor** (Excel/Python) — hieronder beschreven.
> 2. **[Planning](daily-routine/)** — een mobiele web-app voor dagelijkse routines,
>    taken, een weekmenu en voortgang. Zie [`daily-routine/`](daily-routine/).

# Vermogensmonitor Premium — geïntegreerd standaardbestand

Eén schaalbaar Excel-standaardbestand voor de accountantspraktijk dat het actuele
vermogen van een klant (met of zonder partner) inzichtelijk maakt, corrigeert voor
**latente belastingen**, en toekomstgericht doorrekent via scenario's. Combineert
data uit de **aangifte inkomstenbelasting** en de **jaarrekening (BV/DGA)**.

**Hoofdbestand:** `Vermogensmonitor_Premium.xlsx` (v3 — balans-dashboard, import, planning)
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
referentie: icon-chip KPI-cards, doughnut, area-chart, status-signalen en een
donkergroene navigatierail.

## Architectuur — strikte laagscheiding (14 tabbladen)

| Laag | Tabbladen | Rol |
|------|-----------|-----|
| Navigatie | `START` | Klant, peildatum, partnerschakelaar, status |
| Input | `01_Klant`, `02_Input_IB`, `03_Input_JR`, `04_Vermogen`, `05_Historie`, `06_Latente`, `07_Import` | Gestandaardiseerde brondata + import |
| Verwerking | `10_Berekeningen` | KPI's, ratio's, partner-consolidatie, latentie-correctie |
| Output | `20_Dashboard`, `21_Planning` | Balans-dashboard + financiële planning |
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

## Dashboard (`20_Dashboard`) — balansopbouw

- 4 KPI-cards met icon-chips: nettovermogen (na latentie, +Δ vorig jaar),
  bezittingen, schulden + latentie, solvabiliteit.
- **Uitklapbare balans, gegroepeerd per box** (klik ▸/+ om open te klappen):
  - Links **BEZITTINGEN (ACTIVA)** per box (Box 1 / Box 2 / Box 3), met detailregels.
  - Rechtsboven het **NETTOVERMOGEN VAN DE KLANT** (na latentie).
  - Rechts daaronder **SCHULDEN & LATENTE BELASTINGEN (PASSIVA)** per box
    (Box 1 / Box 2 = latente / Box 3), met detailregels.
  - Balanscheck: Bezittingen = Nettovermogen + Schulden + Latenties.
- Doughnut *vermogensverdeling*, *area-chart* ontwikkeling, signalenpaneel.
- Print-klaar (liggend).

## Import (`07_Import`) — 80–90% automatisch

- Plak de export (omschrijving, bedrag, eigenaar); de **mappingtabel** vult
  automatisch categorie, box, type en latente% in.
- Genereert kant-en-klare regels in feitentabel-volgorde (kolommen R:AA) om met
  *plakken-speciaal (waarden)* naar `04_Vermogen` over te zetten.
- Toelichting voor **Power Query** één-klik file-import (fase 2) — af te stemmen
  op jullie software-export (RGS-saldibalans / IB-XML).

## Financiële planning (`21_Planning`) — 10–50 jaar

- Instelbare horizon (10–50 jaar), leeftijden, AOW-/pensioenleeftijd, doelvermogen.
- **Netto maandelijkse kasstroom**: inkomsten (salaris, pensioen/AOW, BV-dividend,
  huur, overig) en uitgaven (levensonderhoud, woonlasten, overig).
- **Meerjarenprognose** jaar-op-jaar: inkomen, uitgaven, saldo, rendement en
  vermogensontwikkeling; met inflatie-indexatie en pensioen/AOW-timing.
- 3 scenario's via Δ rendement/inflatie (Neutraal/Optimistisch/Pessimistisch).
- KPI's (vermogen nu, einde horizon, doel bereikt?) + grafieken
  vermogensontwikkeling t.o.v. doel en inkomsten-vs-uitgaven.

## Gebruik

1. Open in **Microsoft Excel** — formules herberekenen automatisch bij openen.
2. Vul op `START`: klantnaam, peildatum, fiscaal partner (Ja/Nee).
3. Vul de gele invoervelden op `01`–`07` (dropdowns via `90_Config`).
4. Controleer `99_Controles` (alles groen).
5. Bespreek `20_Dashboard` en `21_Planning` met de klant.

## Reproduceren

```bash
python build_v3.py     # Premium (v3 — balans-dashboard, import, planning)
python build_v2.py     # Premium (v2)
python build_vermogensmonitor.py   # Concept (v1)
```

> Alle bedragen zijn **fictieve voorbeelddata** (geanonimiseerd DGA-huishouden).
> Fiscale parameters op `90_Config` zijn voorbeeldwaarden — jaarlijks actualiseren.
> Automatische recalculatie via LibreOffice is in de bouwomgeving niet beschikbaar;
> formules zijn statisch gevalideerd (0 losse scheidingstekens, alle named ranges
> gedefinieerd, alle sheet-referenties geldig) en de kernbedragen + reconciliaties
> zijn handmatig geverifieerd. Excel berekent alles bij openen.
