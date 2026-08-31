# Planning ⚡

Een simpele, mobielvriendelijke web-app voor je dagelijkse routines (met een korte,
max. ~10 min oefenroutine), een takenlijst, een weekmenu en je voortgang. Geen account
en geen backend. In de gepubliceerde app bewaart hij je gegevens via de
`artifact`-opslag (zie **Opslag** hieronder); open je het bestand lokaal, dan valt hij
terug op `localStorage` op je eigen apparaat.

## Wat kan het

- **Vandaag** — je dagelijkse routines op één scherm, een streak-teller (⚡ dagen op rij)
  en de taken die spelen.
- **Trainingsmodus** — een volledig-scherm intervaltimer die je door elke oefening loodst:
  werk-/rustblokken, groot aftellen, voortgangsbalk, geluids- en trilsignalen. Bij afronden
  gaat je streak omhoog. Oefeningen kunnen op **tijd** (aftellen) óf op **herhalingen** — bij
  reps toont het scherm je doel (bijv. 20×) en tik je zelf **Klaar** om door te gaan.
- **Routines** — twee soorten, allebei aan te maken/bewerken/verwijderen:
  - **Timer** (circuit met tijden), bijv. de meegeleverde *Military Ochtend (≈10 min)*.
  - **Checklist** (afvinken zonder timer), bijv. de meegeleverde *Avond voorbereiding*
    (broodjes smeren voor morgen, sportkleding klaarleggen, nieuws checken…).
  - Kies per routine op **welke dagen van de week** hij terugkomt (bijv. alleen ma/wo/vr, of
    elke dag). Op **Vandaag** verschijnen alleen de routines die voor die dag gepland staan —
    je kunt er meerdere hebben, bijv. één ochtend en één avond.
- **Taken** — één lijst met **Korte termijn** bovenaan en **Lange termijn** eronder. Per taak
  een icoon, een context/toelichting en een optionele vervaldatum; afvinken laat een taak
  naar onderen zakken. De ＋-knop opent een bewerk-scherm; tik op een taak om die aan te passen.
- **Menu** — een **weekmenu** voor het avondeten. Kies een **startdag** (bijv. za→vr of zo→vr)
  en vul per dag in wat je eet; met ◀/▶ blader je door weken (oude weken blijven bewaard).
  Bij een lekker gerecht tik je 🔖 om het in je **receptenboekje** (“Mijn gerechten”) te
  bewaren met **ingrediënten** en **bereidingswijze** — alleen voor de gerechten die je wilt
  bewaren. Opgeslagen gerechten stel je snel opnieuw in via de suggestielijst en kun je later
  bewerken.
- **Voortgang** — huidige & langste streak, totaal sessies, en een heatmap van gedane dagen.
  Hier houd je ook **Lichaam & gewicht** bij: lengte (cm), gewicht (kg), automatische **BMI**
  met categorie, en je **streefgewicht** — met een meet-logboek zodat je je gewicht over tijd
  volgt. Verder de instellingen (geluid/trillen) en een backup-export.

De meegeleverde *Military Ochtend* is een startpunt — pas oefeningen, tijden en volgorde
gerust aan, of maak je eigen niveaus op naarmate je fitter wordt.

## Openen

Het is één self-contained bestand.

- **Op de telefoon:** open de gepubliceerde link en kies *Toevoegen aan beginscherm* —
  dan opent hij als een app (fullscreen, offline).
- **Lokaal op je computer:** open `index.html` direct in je browser, of serveer de map:

  ```bash
  cd daily-routine
  python3 -m http.server 8000
  # open http://localhost:8000
  ```

> **Opslag (v2 — betrouwbaar).** Elke opslag krijgt een **tijdstempel** (`_savedAt`). Bij het
> openen verzamelt de app *alle* plekken waar state kan staan — het cloud-databestand
> (`data/state-v2.js`), de in de pagina ingebedde state, `localStorage` en een
> `sessionStorage`-stash van nog niet verzonden wijzigingen — en gebruikt simpelweg de
> **nieuwste**. Daardoor kan een verouderde bron nooit meer winnen (dat was de oorzaak van
> eerder gegevensverlies: het oude `data/state.json`/`state.js` had vaste voorrang).
>
> Verder: wijzigingen worden **direct** lokaal vastgelegd én veiliggesteld vóór elke publicatie
> (zodat een herlaad of `conflict` niets kost), bij het sluiten/achtergrond van de app wordt
> openstaand werk weggeschreven, en in **Voortgang** staat de opslagstatus met **"Laatst
> opgeslagen"**. Werkt de cloud niet, dan toont de app een duidelijke waarschuwing.
> **Backup exporteren** (echte bestandsopslag) en **Backup terugzetten** staan er als noodrem;
> de artifact-**versiegeschiedenis** is bovendien een automatisch archief.

## Techniek

Vanilla HTML/CSS/JavaScript, geen build-stap en geen dependencies. Geluid via de WebAudio-API,
trillen via de Vibration-API (waar ondersteund). Opslag: `artifact`-capability (cloud) in de
gepubliceerde app, met `localStorage` als terugval.
