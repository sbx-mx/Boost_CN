const CACHE="boost-cn-v10-excel-20260727";
const CORE=[
  "./",
  "./index.html",
  "./css/styles.css",
  "./js/app.js",
  "./manifest.json",
  "./assets/boost-cn.png",
  "./data/base-at.json",
  "./data/item-adt.json"
];
const MANIFEST="./data/manifest-data.json";

async function precache(){
  const cache=await caches.open(CACHE);
  const response=await fetch(MANIFEST,{cache:"no-store"});
  if(!response.ok)throw new Error(`No se pudo cargar ${MANIFEST}`);
  const manifest=await response.clone().json();
  const chunks=Array.isArray(manifest.chunks)?manifest.chunks.map(name=>`./data/${name}`):[];
  await cache.addAll([...CORE,...chunks]);
  await cache.put(MANIFEST,response);
}

self.addEventListener("install",event=>{
  event.waitUntil(precache().then(()=>self.skipWaiting()));
});

self.addEventListener("activate",event=>{
  event.waitUntil(
    caches.keys()
      .then(keys=>Promise.all(keys.filter(key=>key!==CACHE).map(key=>caches.delete(key))))
      .then(()=>self.clients.claim())
  );
});

self.addEventListener("fetch",event=>{
  if(event.request.method!=="GET")return;
  const url=new URL(event.request.url);
  const isData=url.pathname.includes("/data/")&&url.pathname.endsWith(".json");
  if(isData){
    event.respondWith(
      fetch(event.request,{cache:"no-store"}).then(response=>{
        const copy=response.clone();
        caches.open(CACHE).then(cache=>cache.put(event.request,copy));
        return response;
      }).catch(()=>caches.match(event.request))
    );
    return;
  }
  event.respondWith(
    caches.match(event.request).then(cached=>cached||fetch(event.request).then(response=>{
      const copy=response.clone();
      caches.open(CACHE).then(cache=>cache.put(event.request,copy));
      return response;
    }).catch(()=>caches.match("./index.html")))
  );
});
