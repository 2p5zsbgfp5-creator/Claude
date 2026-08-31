# Planning ⚡ — met automatische cloud-opslag (Supabase)

Dezelfde Planning-app (dagelijkse routines, trainingstimer, taken, weekmenu +
receptenboekje, en voortgang met streaks/heatmap/gewicht & BMI), maar nu slaat hij
**alles automatisch op in de cloud** via [Supabase](https://supabase.com). Elke
wijziging wordt meteen lokaal bewaard én — ge-debounced — naar je eigen Supabase-
database gestuurd. Zo raak je nooit meer data kwijt als je toestel wordt gewist of
kwijtraakt, en gebruik je dezelfde gegevens op **telefoon én laptop**.

> **Lokaal-first en veilig bij storingen.** De app werkt eerst uit de browseropslag
> op je toestel (snel, offline). De cloud is een extra, automatische synchronisatie-
> laag. Ben je offline, dan blijft alles lokaal werken en wordt je wijziging later
> vanzelf nagestuurd (je ziet "☁︎ offline — wordt later gesynct" onder in *Voortgang*).

## Werkt hij ook zonder Supabase?

Ja. Zolang je in `config.js` de placeholders laat staan, draait de app **puur lokaal**
op je toestel (zoals de oude versie). Vul je de twee waarden in, dan gaat de app
inloggen en synchroniseren. Je kunt dus eerst uitproberen en later de cloud aanzetten.

---

## Eenmalig instellen (± 10 minuten)

### 1. Maak een gratis Supabase-project

1. Ga naar <https://supabase.com> → **Start your project** → log in.
2. **New project** → geef het een naam en kies een regio (bv. *West EU*). Bewaar het
   database-wachtwoord ergens (heb je hier verder niet nodig).
3. Wacht tot het project klaar is.

### 2. Maak de tabel + beveiliging (Row Level Security)

Open in Supabase **SQL Editor** → **New query**, plak onderstaande SQL en klik **Run**:

```sql
-- Eén rij per gebruiker met de volledige app-staat als JSON.
create table if not exists public.app_state (
  user_id    uuid primary key references auth.users (id) on delete cascade,
  data       jsonb not null default '{}'::jsonb,
  updated_at timestamptz not null default now()
);

-- Beveiliging: iedere gebruiker mag ALLEEN de eigen rij zien/wijzigen.
alter table public.app_state enable row level security;

create policy "eigen rij lezen"    on public.app_state
  for select using  (auth.uid() = user_id);
create policy "eigen rij invoegen" on public.app_state
  for insert with check (auth.uid() = user_id);
create policy "eigen rij wijzigen" on public.app_state
  for update using  (auth.uid() = user_id) with check (auth.uid() = user_id);
```

De publieke `anon`-key mag hierna gerust in de app staan: zónder ingelogde gebruiker
geeft Row Level Security niets terug, en met inloggen zie je alleen je eigen rij.

### 3. Zet inloggen met e-mail + zelfgekozen code aan (géén mail)

De app logt in met je **e‑mailadres + een code die je zelf kiest** (een wachtwoord),
volledig in de app — er wordt geen mail verstuurd. Daarvoor moet e‑mailbevestiging uit:

1. **Authentication → Providers → Email** (of **Sign In / Providers → Email**): zorg dat
   **Email** aanstaat.
2. Zet **"Confirm email" UIT**. Dan maakt aanmelden meteen een actieve sessie zonder
   bevestigingsmail, zodat je puur met je e‑mail + code inlogt.

> Geen custom SMTP of e‑mailsjablonen nodig. Op het gratis niveau kun je die toch niet
> aanpassen — deze methode omzeilt dat volledig.
>
> Beveiliging: zonder e‑mailbevestiging kan in principe iedereen een account aanmaken,
> maar door Row Level Security ziet elk account alleen de **eigen** gegevens. Kies een
> code die je verder nergens gebruikt.

### 4. Vul je twee sleutels in `config.js`

In Supabase: **Project Settings → Data API** (of **API**). Neem over:

- **Project URL** → `SUPABASE_URL`
- **anon public key** → `SUPABASE_ANON_KEY`

Open `config.js` en vervang de placeholders:

```js
window.PLANNING_CONFIG = {
  SUPABASE_URL:      "https://abcdxyz.supabase.co",
  SUPABASE_ANON_KEY: "sb_publishable_...."   // de publishable/anon-key
};
```

> Zet hier **nooit** de `service_role`-key neer — die omzeilt de beveiliging.

### 5. Online zetten (GitHub Pages)

Zet de map `planning/` in je repository en zet **GitHub Pages** aan
(Settings → Pages → *Deploy from a branch*). De app staat dan op
`https://<gebruikersnaam>.github.io/<repo>/planning/`. Open dat adres op je telefoon
en kies **Zet op beginscherm** — dan opent hij fullscreen en offline.

---

## Zelf uitproberen / lokaal draaien

Niet dubbelklikken (`file://` werkt niet voor de service worker en het inloggen).
Serveer de map:

```bash
cd planning
python3 -m http.server 8000
# open http://localhost:8000
```

## Zo werkt het opslaan

- **Optimistische weergave**: je handeling is direct zichtbaar.
- **Direct lokaal opgeslagen** (`localStorage` + IndexedDB) — snel en offline.
- **Ge-debounced naar de cloud** (~1 sec na je laatste wijziging), zodat typen geen
  reeks netwerk-verzoeken oplevert.
- **Nieuwste wint**: bij inloggen worden cloud en lokaal samengevoegd op tijdstempel
  (`_savedAt`), zodat je verdergaat waar je gebleven was — ook op een nieuw toestel.
- **Offline-vangnet**: mislukt de cloud-schrijf, dan blijft de lokale kopie leidend en
  wordt de wijziging later automatisch nagestuurd (bij internet terug, bij het sluiten
  van de app, of periodiek). De bestaande dagelijkse lokale back-ups blijven bestaan.
- Onderin **Voortgang** zie je de status: *op dit toestel* + *☁︎ gesynct in de cloud*.

## Techniek

Vanilla HTML/CSS/JavaScript, geen build-stap. Cloud via `@supabase/supabase-js`
(CDN). Auth via Supabase e‑mail + wachtwoord (in-app, zonder mail). `sw.js` cachet alleen de app-schil (nooit je
gegevens). Opslag-laag haakt in op de bestaande `save()`/`load()`: de app-functionali-
teit is ongewijzigd, alleen waar het naartoe wordt opgeslagen is uitgebreid.
