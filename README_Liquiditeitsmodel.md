# Liquiditeitsmodel Handelsbedrijf — 24 maanden

Professioneel, commercieel inzetbaar liquiditeits- en cashflowadviesmodel voor
handelsbedrijven, bedoeld voor de accountant / Opdrachtmanager. Het model rekent
maximaal **24 maanden** vooruit en helpt niet alleen *rekenen* maar vooral het
**verhaal achter de liquiditeit** te vertellen in het klantgesprek.

**Bestand:** `Liquiditeitsmodel_Handelsbedrijf_24m.xlsx`

## Opzet (geen black box)

Strikte laagscheiding met kleurcodering:

- **Geel / blauw** = invoer (handmatig)
- **Groen getal** = link naar ander tabblad
- **Zwart** = berekening
- **Donkergroen** = belangrijke output

Elk bedrag is via zichtbare formules herleidbaar tot een invoercel of aanname.
Geen VBA; alleen standaard, controleerbare Excel-functies (SUMIFS, INDEX/MATCH,
IF, IFERROR, EDATE, MOD, TEXT, …) en Excel-tabellen.

## 17 tabbladen

| Tab | Inhoud |
|-----|--------|
| 01 Dashboard | KPI-kaarten, 6 grafieken, dynamische signalen, modelstatus |
| 02 Instellingen | Parameters + **scenariomatrix** (Base/Best/Worst/Custom) |
| 03 Historie | Historische W&V + startposities balans |
| 04 Input prognose | Seizoenspatroon omzet, overige ontvangsten/uitgaven, privé/dividend |
| 05 Omzet & marge | Omzet × seizoen × scenario, brutomarge, inkoopwaarde |
| 06 Werkkapitaal | Debiteuren/voorraad/crediteuren via **DSO/DPO/DIO**, DIO+DSO−DPO = CCC |
| 07 Personeel | Loonsom, werkgeverslasten, vakantiegeld, FTE, what-if loonmutatie |
| 08 Overige kosten | Kostencategorieën: vast bedrag of % van omzet |
| 09 Investeringen | CAPEX-schema, kas-impact op betaalmoment, afschrijving apart |
| 10 Financiering | Lening, aflossing, rente, RC-limiet, nieuwe financiering |
| 11 Belastingen | BTW (kwartaal/maand), VPB; loonheffing als memo (geen dubbeltelling) |
| 12 Liquiditeitsprognose | Begin + inkomend − uitgaand = eind, buffer, financieringsbehoefte |
| 13 Scenarioanalyse | Base/Best/Worst naast elkaar + vergelijkingstabel |
| 14 Import | Exact Online-template + mappingtabel + Power Query-route |
| 15 Controle | Automatische controles → MODEL OK / ACTIE VEREIST |
| 16 Toelichting | Werking, methodiek, aannames, disclaimer |
| 17 Copilot Prompts | Kant-en-klare prompts voor Microsoft Copilot in Excel |

## Methodiek werkkapitaal

Voorraad, debiteuren en crediteuren worden gemodelleerd als **streefsaldi op
dagbasis**:

```
Voorraad(mnd)   = DIO / 30,4 × inkoopwaarde omzet
Debiteuren(mnd) = DSO / 30,4 × omzet incl. BTW
Crediteuren(mnd)= DPO / 30,4 × inkopen incl. BTW
Ontvangsten     = beginstand debiteuren + facturatie − eindstand debiteuren
```

Zo vertaalt een langere DSO of extra voorraad zich direct in een lagere
kaspositie en hogere financieringsbehoefte — het cash-effect wordt zichtbaar.

## Fictieve testdata & scenario-uitkomsten

Handelsbedrijf: jaaromzet € 5 mln · brutomarge 28% · voorraad € 750k ·
debiteuren € 600k · crediteuren € 500k · personeelskosten ± € 700k ·
banklening € 900k · openingssaldo € 300k · minimumbuffer € 100k.

| Scenario | Min. liquiditeit | Max. financieringsbehoefte | Eind (mnd 24) |
|----------|-----------------:|---------------------------:|--------------:|
| **Base** | € 76.530 | € 23.470 | € 233.936 |
| **Best** | € 132.136 | € 0 | € 1.092.123 |
| **Worst** | −€ 998.658 | € 1.098.658 | −€ 936.969 |

Base raakt de buffer licht (klein tekort), Worst laat een duidelijke
financieringsbehoefte zien — het model reageert logisch op omzet, marge, DSO,
DPO, DIO, personeelskosten en investeringen (zie `crosscheck.py`).

## Reproduceren

```
python3 build_liquiditeit.py     # bouwt het model (formules)
python3 inject_values.py         # rekent door en schrijft cache-waarden
```

`fullCalcOnLoad` staat aan: Excel herberekent bij openen. De cache-waarden
zorgen dat cijfers ook in andere viewers direct zichtbaar zijn.

## Verificatie

Het model is geverifieerd met **twee onafhankelijke rekenengines** (een
handmatige Python-replica in `crosscheck.py` én de pure-Python `formulas`-engine
op het xlsx-bestand). Beide geven identieke uitkomsten en **0 formulefouten**
over alle 2.814 formules. Statische controle bevestigt: alle kruisverwijzingen
en named ranges geldig, geen niet-ondersteunde functies.
