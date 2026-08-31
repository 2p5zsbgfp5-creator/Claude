/* ============================================================
   Supabase-instellingen  —  VUL HIER JE EIGEN TWEE WAARDEN IN
   ============================================================
   Waar vind je deze? In je Supabase-project:
   Project Settings  ->  Data API  (of "API")
     • Project URL   ->  hieronder bij SUPABASE_URL
     • anon public key (Legacy API keys / "anon")  ->  SUPABASE_ANON_KEY

   Deze twee waarden mogen openbaar in de app staan: de anon-key kan zonder
   inloggen niets, want de database is beveiligd met Row Level Security (zie
   de README). Zet hier NOOIT de "service_role"-key neer.

   Zolang je deze twee op de placeholder laat staan, werkt de app gewoon
   LOKAAL op dit toestel (zoals vroeger). Zodra je ze invult, gaat de app
   inloggen en alles automatisch in de cloud opslaan en synchroniseren.
============================================================ */
window.PLANNING_CONFIG = {
  SUPABASE_URL:      "https://zcdvycsrkfzzfmugsdyb.supabase.co",
  SUPABASE_ANON_KEY: "sb_publishable_pgg1hnIb80j8Rw-63FsfmA_DjKEBPBS"
};
