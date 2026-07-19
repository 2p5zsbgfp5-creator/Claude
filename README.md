# Vermogensmonitor — geïntegreerd standaardbestand (concept / MVP)

Eén schaalbaar Excel-standaardbestand voor de accountantspraktijk dat het actuele
vermogen van een klant (met of zonder partner) inzichtelijk maakt én
toekomstgericht doorrekent via scenario's. Ontworpen om data uit de **aangifte
inkomstenbelasting** en de **jaarrekening** te combineren.

**Bestand:** `Vermogensmonitor_Concept.xlsx`

---

## Architectuur — strikte laagscheiding

| Laag | Tabbladen | Rol |
|------|-----------|-----|
| **Navigatie** | `START` | Klant, peildatum, partnerschakelaar, versie, status |
| **Input** (blauw) | `01_Klant`, `02_Input_IB`, `03_Input_JR`, `04_Input_Vermogen`, `05_Historie` | Gestandaardiseerde brondata-invoer |
| **Verwerking** (grijs) | `11_Berekeningen` | KPI's, ratio's, partner-consolidatie |
| **Output** (groen) | `20_Dashboard`, `21_Scenario` | Cliëntdashboard + financiële planning |
| **Config/QA** (oranje) | `90_Config`, `91_Mapping`, `99_Controles` | Keuzelijsten, mapping, kwaliteitscontroles |

**Datastroom:** `02–04 Input → 04 feitentabel → 11 Berekeningen → 20 Dashboard / 21 Scenario`,
bewaakt door `99_Controles`, gevoed door `90_Config` + `91_Mapping`.

## Kernontwerpbeslissingen

- **Genormaliseerde feitentabel** (`04_Input_Vermogen`): één regel per vermogenscomponent,
  met `Eigenaar` (Klant/Partner/Gezamenlijk) en `Aandeel klant %`. Alle KPI's zijn
  `SUMIFS`/`SUMPRODUCT` hierover → onbeperkt uitbreidbaar zonder herbouw.
- **Eén partnerschakelaar** (`Heeft_Partner` op `START`) stuurt het hele bestand;
  levert altijd drie totalen: Klant / Partner / Gezamenlijk.
- **Centrale configuratie**: categorieën, boxen, fiscale parameters en mapping op één plek
  → 1× per jaar bijwerken, geldt voor alle klanten.
- **Ingebouwde reconciliaties**: box 3 (monitor ↔ aangifte), balans-evenwicht,
  AB-waarde ↔ eigen vermogen BV, partnersplitsing, tekencontrole.
- **Alleen breed-ondersteunde functies** (`SUMIFS`, `SUMPRODUCT`, `IF`, `IFERROR`,
  `TEXT`, `ABS`) — geen `XLOOKUP`/spill-functies → maximale compatibiliteit.

## Belangrijke aannames

- Alle bedragen in het bestand zijn **fictieve voorbeelddata** (geanonimiseerd DGA-huishouden
  met partner) ter demonstratie. Vervang de blauwe/gele invoercellen door klantdata.
- Fiscale parameters op `90_Config` zijn **voorbeeldwaarden** — jaarlijks actualiseren.
- Het is een **MVP (fase 1, handmatige invoer)**. Power Query-import van IB/jaarrekening,
  volledige historie en uitgebreide scenario's zijn **fase 2**.

## Gebruik

1. Open in **Microsoft Excel**. Formules herberekenen automatisch bij openen
   (het bestand wordt zonder cache-waarden gegenereerd; Excel doet een volledige
   herberekening bij de eerste keer openen).
2. Vul op `START`: klantnaam, peildatum en **Heeft partner? (JA/NEE)**.
3. Vul de blauwe/gele cellen op `01`–`05` (dropdowns via `90_Config`).
4. Controleer `99_Controles` (OK = groen).
5. Bespreek `20_Dashboard` en `21_Scenario` met de klant.

## Genereren / reproduceren

Het bestand wordt programmatisch opgebouwd met `openpyxl`:

```bash
python build_vermogensmonitor.py
```

> Let op: automatische recalculatie via LibreOffice is in deze omgeving niet
> beschikbaar (de LibreOffice-import is defect). Formules zijn statisch
> gevalideerd (0 losse scheidingstekens, alle named ranges gedefinieerd, alle
> sheet-referenties geldig) en de kernbedragen zijn handmatig geverifieerd.
> Excel berekent alles bij openen.
