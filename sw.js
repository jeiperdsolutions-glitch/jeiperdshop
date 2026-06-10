const C='jeiperd-v1';
self.addEventListener('install',function(){self.skipWaiting();});
self.addEventListener('activate',function(e){e.waitUntil(caches.keys().then(function(ks){return Promise.all(ks.map(function(k){if(k!==C){return caches.delete(k);}}));}));self.clients.claim();});
self.addEventListener('fetch',function(e){
  var r=e.request; if(r.method!=='GET'){return;}
  var u=new URL(r.url);
  if(u.pathname.indexOf('/img/')!==-1){
    e.respondWith(caches.open(C).then(function(c){return c.match(r).then(function(hit){return hit||fetch(r).then(function(res){c.put(r,res.clone());return res;});});}));
  } else if(r.mode==='navigate'){
    e.respondWith(fetch(r).catch(function(){return caches.match(r);}));
  }
});
