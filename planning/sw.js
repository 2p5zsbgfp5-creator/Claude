/* Service worker voor Planning.
   Doel: de app opent vanaf je beginscherm ook zonder internet.
   Let op: hier wordt ALLEEN de app zelf gecachet — nooit je gegevens.
   Die staan in localStorage/IndexedDB en komen hier niet langs. */

const CACHE_VERSION = "planning-v5";
const SHELL = ["./", "./index.html", "./config.js", "./manifest.json",
               "./apple-touch-icon.png", "./icon-192.png", "./icon-512.png"];

self.addEventListener("install", e=>{
  e.waitUntil(
    caches.open(CACHE_VERSION)
      .then(c => c.addAll(SHELL))
      .catch(()=>{})            // één ontbrekend bestand mag de installatie niet slopen
  );
});

self.addEventListener("activate", e=>{
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE_VERSION).map(k => caches.delete(k))))
      .then(()=> self.clients.claim())
  );
});

self.addEventListener("fetch", e=>{
  const req = e.request;
  if(req.method !== "GET") return;
  if(new URL(req.url).origin !== self.location.origin) return;

  // Netwerk eerst, cache als terugval: zo krijg je updates zodra je online bent,
  // en opent de app alsnog wanneer je dat niet bent.
  e.respondWith(
    fetch(req)
      .then(res=>{
        if(res && res.ok){
          const kopie = res.clone();
          caches.open(CACHE_VERSION).then(c => c.put(req, kopie)).catch(()=>{});
        }
        return res;
      })
      .catch(()=> caches.match(req).then(hit => hit || caches.match("./index.html")))
  );
});
