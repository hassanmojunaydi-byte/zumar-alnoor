const C="zn-v22";const A=["./","index.html","manifest.webmanifest","icon-192.png","icon-512.png","apple-touch-icon.png"];
self.addEventListener("install",e=>{self.skipWaiting();e.waitUntil(caches.open(C).then(c=>c.addAll(A)).catch(()=>{}));});
self.addEventListener("activate",e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==C).map(x=>caches.delete(x)))).then(()=>self.clients.claim()));});
self.addEventListener("notificationclick",e=>{e.notification.close();const go=(e.notification.data&&e.notification.data.go)||"";e.waitUntil(clients.matchAll({type:"window",includeUncontrolled:true}).then(cs=>{for(const c of cs){if("focus"in c){c.navigate&&go&&c.navigate("./"+go).catch(()=>{});return c.focus();}}return clients.openWindow("./"+go);}));});
self.addEventListener("fetch",e=>{const r=e.request;if(r.method!=="GET"||new URL(r.url).origin!==location.origin)return;
 e.respondWith(fetch(r).then(res=>{const cp=res.clone();caches.open(C).then(c=>c.put(r,cp));return res;}).catch(()=>caches.match(r).then(m=>m||caches.match("index.html"))));});
