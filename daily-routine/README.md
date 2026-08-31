# Planning ⚡

Een simpele, mobielvriendelijke web-app voor je dagelijkse routines (met een korte,
max. ~10 min oefenroutine), een takenlijst, een weekmenu en je voortgang. Geen account
en geen backend: je gegevens staan in de browseropslag van je eigen toestel en verlaten
dat toestel nooit. Zie **Opslag** hieronder.

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

- **Op de telefoon (zo hoort het):** open <https://2p5zsbgfp5-creator.github.io/planning/>
  en kies *Zet op beginscherm*. Daarna opent hij als een app: fullscreen, offline, en met
  opslag die de browser niet opruimt.
- **Lokaal op je computer:** serveer de map (niet dubbelklikken — via `file://` werkt de
  service worker niet en is de opslag onbetrouwbaar):

  ```bash
  cd daily-routine
  python3 -m http.server 8000
  # open http://localhost:8000
  ```

## Opslag (v3)

De app draait op een gewoon webadres, dus opslaan is **eerste-partij browseropslag op je eigen
toestel**: direct, synchroon, zonder netwerk. Er is niets om mee te botsen en niets dat kan
weigeren. Elke wijziging gaat meteen naar `localStorage` én naar IndexedDB.

Waarom dit anders is dan v1/v2: die versies bewaarden door de héle app opnieuw te *publiceren*
via de artifact-capability. Dat systeem beperkt de frequentie (`rate_limited`) en laat
botsingen (`conflict`) vervallen — allebei gedocumenteerd als normaal gedrag. Voor een app
waar je dagelijks in werkt betekende dat stil gegevensverlies.

Vangnetten:

- **Dagelijkse backups** in IndexedDB (laatste 14), die meelopen met je laatste wijziging en
  nooit door een armere versie worden overschreven.
- **Een controle bij het opstarten.** Ziet de app minder gegevens dan de laatste backup, dan
  gaat hij *niet* stil verder maar toont een herstelbanner met de datum van die backup.
- **Backup exporteren / terugzetten** in Voortgang (bestand kiezen of JSON plakken).
- In **Voortgang** staat `Opgeslagen: <datum tijd> · op dit toestel`. Heeft de browser
  blijvende opslag geweigerd, dan waarschuwt de app dat je de app op je beginscherm moet zetten.

Beperkingen, eerlijk: je gegevens staan op **dit ene toestel**. Browsergegevens wissen of je
telefoon kwijtraken betekent gegevens kwijt — vandaar de export. En iOS ruimt opslag van gewone
websites na ongeveer een week zonder gebruik op; dat gebeurt niet als de app op je beginscherm
staat. *Op het beginscherm zetten is dus geen luxe maar onderdeel van de opzet.*

## Techniek

Vanilla HTML/CSS/JavaScript, geen build-stap en geen dependencies. Geluid via de WebAudio-API,
trillen via de Vibration-API (waar ondersteund). Opslag via `localStorage` + IndexedDB, met
`navigator.storage.persist()`. `sw.js` is een service worker die alleen de app zelf cachet
(nooit je gegevens) zodat hij offline opent.
