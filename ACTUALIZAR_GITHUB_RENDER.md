{% extends 'base.html' %}{% block content %}
<div class="topline"><div><h1>Ubicación de vendedores</h1><p class="muted">Última posición enviada durante el seguimiento activo. Horarios expresados en Rivera, Uruguay.</p></div><span id="updated" class="pill">Actualizando…</span></div>
<section class="panel map-panel"><div id="map" class="live-map" aria-label="Mapa de ubicación de vendedores"></div><div id="mapMessage" class="map-message" hidden></div></section>
<section class="panel"><div class="section-head"><h2>Estado de dispositivos</h2><button type="button" class="secondary" id="recenterMap">Centrar mapa</button></div><div class="table-wrap small-table"><table><thead><tr><th>Vendedor</th><th>Estado</th><th>Última actualización</th><th>Hora Rivera</th><th>Precisión</th><th>Velocidad</th><th>Mapa</th></tr></thead><tbody id="locationRows"></tbody></table></div></section>
{% endblock %}{% block scripts %}
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css">
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<script>
let map, fittedOnce=false; const markers={}, accuracyCircles={};
const mapMessage=document.getElementById('mapMessage');
function initMap(){
  if(typeof L==='undefined'){mapMessage.hidden=false;mapMessage.textContent='No fue posible cargar el mapa. Verifica la conexión y recarga la página.';return;}
  map=L.map('map',{zoomControl:true,preferCanvas:true}).setView([-30.905,-55.55],13);
  L.tileLayer('https://tile.openstreetmap.org/{z}/{x}/{y}.png',{maxZoom:19,attribution:'&copy; OpenStreetMap',tileSize:256,updateWhenIdle:false,keepBuffer:3}).addTo(map);
  setTimeout(()=>map.invalidateSize(true),150);
  window.addEventListener('resize',()=>map.invalidateSize(false));
  document.addEventListener('visibilitychange',()=>{if(!document.hidden)setTimeout(()=>map.invalidateSize(true),100)});
}
function ago(seconds){seconds=Math.max(0,Number(seconds)||0);if(seconds<60)return 'hace '+Math.round(seconds)+' s';if(seconds<3600)return 'hace '+Math.floor(seconds/60)+' min';if(seconds<86400)return 'hace '+Math.floor(seconds/3600)+' h';return 'hace '+Math.floor(seconds/86400)+' días'}
function localTime(iso){try{return new Intl.DateTimeFormat('es-UY',{timeZone:'America/Montevideo',day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit',second:'2-digit'}).format(new Date(iso));}catch(e){return '—'}}
function fitAll(force=false){if(!map)return;const pts=Object.values(markers).map(m=>m.getLatLng());if(!pts.length)return;if(pts.length===1)map.setView(pts[0],15);else map.fitBounds(L.latLngBounds(pts),{padding:[45,45],maxZoom:15});fittedOnce=true;setTimeout(()=>map.invalidateSize(true),50)}
async function refresh(){
 try{
  const r=await fetch('/api/locations',{cache:'no-store'}); if(!r.ok)throw new Error('Error '+r.status); const data=await r.json();
  const rows=document.getElementById('locationRows'); rows.innerHTML=''; const active=new Set();
  for(const x of data){const tr=document.createElement('tr');
   if(x.latitude!==undefined){active.add(x.seller);const state=x.online?'<span class="status online">En línea</span>':'<span class="status offline">Sin actualización reciente</span>'; const precision=x.accuracy!=null?Math.round(x.accuracy)+' m':'—';
    tr.innerHTML=`<td><b>${x.seller}</b></td><td>${state}</td><td>${ago(x.age_seconds)}</td><td>${localTime(x.created_at)}</td><td>${precision}</td><td>${x.speed!=null?Math.round(x.speed*3.6)+' km/h':'—'}</td><td><a class="button secondary" target="_blank" rel="noopener" href="https://www.google.com/maps/search/?api=1&query=${x.latitude},${x.longitude}">Abrir</a></td>`;
    const pos=[x.latitude,x.longitude]; if(!markers[x.seller])markers[x.seller]=L.marker(pos,{title:x.seller}).addTo(map);else markers[x.seller].setLatLng(pos);
    markers[x.seller].bindPopup(`<b>${x.seller}</b><br>${x.online?'En línea':'Sin actualización reciente'}<br>${localTime(x.created_at)}<br>Precisión: ${precision}`);
    if(x.accuracy!=null){if(!accuracyCircles[x.seller])accuracyCircles[x.seller]=L.circle(pos,{radius:x.accuracy,weight:1,fillOpacity:.08}).addTo(map);else accuracyCircles[x.seller].setLatLng(pos).setRadius(x.accuracy);}
   }else tr.innerHTML=`<td><b>${x.seller}</b></td><td><span class="status offline">Sin datos</span></td><td>—</td><td>—</td><td>—</td><td>—</td><td>—</td>`;
   rows.appendChild(tr);
  }
  for(const seller of Object.keys(markers)){if(!active.has(seller)){map.removeLayer(markers[seller]);delete markers[seller];if(accuracyCircles[seller]){map.removeLayer(accuracyCircles[seller]);delete accuracyCircles[seller];}}}
  if(!fittedOnce)fitAll(); map.invalidateSize(false); updated.textContent='Actualizado '+new Intl.DateTimeFormat('es-UY',{timeZone:'America/Montevideo',hour:'2-digit',minute:'2-digit',second:'2-digit'}).format(new Date()); mapMessage.hidden=true;
 }catch(e){updated.textContent='Error de actualización';mapMessage.hidden=false;mapMessage.textContent='No se pudieron actualizar las ubicaciones: '+e.message;}
}
window.addEventListener('load',()=>{initMap();document.getElementById('recenterMap').onclick=()=>fitAll(true);setTimeout(()=>{if(map)map.invalidateSize(true);refresh();},250);setInterval(refresh,20000);});
</script>{% endblock %}
