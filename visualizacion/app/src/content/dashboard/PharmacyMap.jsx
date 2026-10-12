import React, { useEffect, useLayoutEffect, useRef, useState } from "react";
import * as L from "../assets/leaflet/leaflet.esm.js";
import "../assets/leaflet/leaflet.css";
import { CATEGORY_COLORS, eur, num, pop, pct, potentialColor, markLabelColor } from "./market-utils.js";
import { MarketIcon } from "./MarketIcons.jsx";

const COUNTRY_BOUNDS = [[27.5, -18.5], [44, 4.5]];
const REGIONS = { "España completa": COUNTRY_BOUNDS, "Península y Baleares": [[35.5, -10], [44, 4.5]], Canarias: [[27.5, -18.5], [29.6, -13.2]], "Ceuta y Melilla": [[35.1, -5.7], [36.1, -2.5]] };
const escape = value => String(value ?? "").replace(/[&<>"']/g, char => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[char]));
const fetchJSON = async url => { const response = await fetch(url); if (!response.ok) throw new Error(`No se pudo cargar ${url}`); return response.json(); };

export function PharmacyMap({ rows, selected, sbCode, ccaaCode, reference, colorMode, scaleMax, fitKey, onSelect, onSB, onCCAA }) {
  const host = useRef(null), mapRef = useRef(null), canvasRef = useRef(null), drawRef = useRef(null), marksRef = useRef([]), stateRef = useRef({});
  const hoverCardRef = useRef(null);
  const [hoverHeight, setHoverHeight] = useState(340);
  const provinceLayer = useRef(null), sbLayer = useRef(null), ccaaLayer = useRef(null);
  const [geography, setGeography] = useState(null), [error, setError] = useState(""), [hover, setHover] = useState(null), [zoom, setZoom] = useState(5), [coincident, setCoincident] = useState([]);
  const [territoryMode, setTerritoryMode] = useState("ccaa");
  const effectiveSB = sbCode && sbCode !== "all" ? sbCode : null;
  stateRef.current = { rows, selected, colorMode, scaleMax, onSelect, onSB, onCCAA };
  useEffect(() => { if (sbCode && sbCode !== "all") setTerritoryMode("sb"); }, [sbCode]);
  useLayoutEffect(() => { if (hoverCardRef.current) setHoverHeight(hoverCardRef.current.getBoundingClientRect().height); }, [hover?.mark.row?.id, hover?.mark.rows?.length, hover?.mark.potential, colorMode]);

  useEffect(() => {
    const controller = new AbortController();
    Promise.all([fetch("/data/ccaa.geojson", { signal: controller.signal }).then(r => { if (!r.ok) throw Error("CCAA no disponibles"); return r.json(); }), fetch("/data/provinces.geojson", { signal: controller.signal }).then(r => { if (!r.ok) throw Error("Cartografía no disponible"); return r.json(); }), fetch("/data/sanibricks.geojson", { signal: controller.signal }).then(r => { if (!r.ok) throw Error("Sanibricks no disponibles"); return r.json(); })])
      .then(([ccaa, provinces, sanibricks]) => setGeography({ ccaa, provinces, sanibricks })).catch(e => { if (e.name !== "AbortError") setError(e.message); });
    return () => controller.abort();
  }, []);

  useEffect(() => {
    const map = L.map(host.current, { center: [38, -7], zoom: 5, minZoom: 4, maxZoom: 19, preferCanvas: true, zoomControl: false, attributionControl: true, zoomAnimation: false, fadeAnimation: false });
    mapRef.current = map;
    L.control.zoom({ position: "topright" }).addTo(map);
    L.control.scale({ imperial: false, position: "bottomleft" }).addTo(map);
    map.attributionControl.setPrefix(false);
    map.attributionControl.addAttribution('Cartografía INE 2025 · <a href="https://leafletjs.com" target="_blank" rel="noreferrer">Leaflet</a>');
    const canvas = document.createElement("canvas"); canvas.className = "market-marker-canvas"; canvas.setAttribute("aria-hidden", "true"); map.getContainer().appendChild(canvas); canvasRef.current = canvas;
    let scheduled = 0;
    const draw = () => {
      const { rows: points, selected: active, colorMode: mode, scaleMax: nationalMaximum } = stateRef.current;
      const size = map.getSize(), ratio = window.devicePixelRatio || 1;
      canvas.width = size.x * ratio; canvas.height = size.y * ratio; canvas.style.width = `${size.x}px`; canvas.style.height = `${size.y}px`;
      const ctx = canvas.getContext("2d"); ctx.scale(ratio, ratio); ctx.clearRect(0, 0, size.x, size.y);
      const isGrouped = map.getZoom() < 10, maximum = nationalMaximum || 1;
      const buckets = new Map(), preciseBuckets = new Map(), marks = [];
      for (const row of points) {
        if (!Number.isFinite(row.lat) || !Number.isFinite(row.lon)) continue;
        const point = map.latLngToContainerPoint([row.lat, row.lon]);
        if (point.x < -30 || point.x > size.x + 30 || point.y < -30 || point.y > size.y + 30) continue;
        if (isGrouped) {
          const key = `${Math.floor(point.x / 38)}_${Math.floor(point.y / 38)}`;
          let bucket = buckets.get(key);
          if (!bucket) { bucket = { x: 0, y: 0, rows: [], potential: 0, counts: { A: 0, B: 0, C: 0, D: 0 }, pop65: 0, provisional: 0 }; buckets.set(key, bucket); }
          bucket.rows.push(row); bucket.x += point.x; bucket.y += point.y; bucket.potential += row.potential_eur; bucket.counts[row.category] += 1; bucket.pop65 += row.pop65; bucket.provisional += row.location_provisional ? 1 : 0;
        } else {
          const key = `${row.lat.toFixed(7)}_${row.lon.toFixed(7)}`;
          let bucket = preciseBuckets.get(key);
          if (!bucket) { bucket = { x: point.x, y: point.y, rows: [], potential: 0, counts: { A: 0, B: 0, C: 0, D: 0 }, pop65: 0, provisional: 0 }; preciseBuckets.set(key, bucket); }
          bucket.rows.push(row); bucket.potential += row.potential_eur; bucket.counts[row.category] += 1; bucket.pop65 += row.pop65; bucket.provisional += row.location_provisional ? 1 : 0;
        }
      }
      if (isGrouped) for (const bucket of buckets.values()) {
        bucket.x /= bucket.rows.length; bucket.y /= bucket.rows.length;
        if (bucket.rows.length === 1) marks.push({ x: bucket.x, y: bucket.y, r: 5.5, row: bucket.rows[0] });
        else marks.push({ ...bucket, r: Math.min(20, 11 + Math.log10(bucket.rows.length) * 3), grouped: true });
      }
      if (!isGrouped) for (const bucket of preciseBuckets.values()) {
        if (bucket.rows.length === 1) { const row = bucket.rows[0]; marks.push({ x: bucket.x, y: bucket.y, r: 4 + Math.min(5, Math.sqrt(row.potential_eur / maximum) * 5), row }); }
        else marks.push({ ...bucket, r: Math.min(18, 10 + Math.log10(bucket.rows.length) * 3), grouped: true, coincident: true });
      }
      marks.sort((a, b) => Number(Boolean(active && a.row?.id === active.id)) - Number(Boolean(active && b.row?.id === active.id)));
      for (const mark of marks) {
        const category = mark.grouped ? Object.entries(mark.counts).sort((a, b) => b[1] - a[1])[0][0] : mark.row.category;
        const color = mode === "euros" ? potentialColor(mark.grouped ? mark.potential / mark.rows.length : mark.row.potential_eur, maximum) : CATEGORY_COLORS[category] || CATEGORY_COLORS.D;
        ctx.beginPath(); ctx.arc(mark.x, mark.y, mark.r, 0, Math.PI * 2);
        ctx.fillStyle = mark.grouped || !mark.row.location_provisional ? color : "#ffffff"; ctx.fill();
        ctx.strokeStyle = mark.grouped ? "#ffffff" : color; ctx.lineWidth = mark.grouped ? 1.7 : mark.row.location_provisional ? 2 : 1; ctx.stroke();
        if (mark.grouped) { ctx.fillStyle = markLabelColor(color); ctx.font = "600 10px Segoe UI, sans-serif"; ctx.textAlign = "center"; ctx.textBaseline = "middle"; ctx.fillText(mark.rows.length > 999 ? `${Math.round(mark.rows.length / 100) / 10}k` : String(mark.rows.length), mark.x, mark.y); }
        if (active && (mark.row?.id === active.id || mark.rows?.some(row => row.id === active.id))) { ctx.beginPath(); ctx.arc(mark.x, mark.y, mark.r + 4, 0, Math.PI * 2); ctx.strokeStyle = "#00005A"; ctx.lineWidth = 2; ctx.stroke(); }
      }
      marksRef.current = marks;
    };
    const schedule = () => { if (!scheduled) scheduled = requestAnimationFrame(() => { scheduled = 0; draw(); }); };
    drawRef.current = schedule;
    const hit = point => { let found = null, distance = Infinity; for (const mark of marksRef.current) { const d = Math.hypot(mark.x - point.x, mark.y - point.y); if (d < mark.r + 4 && d < distance) { found = mark; distance = d; } } return found; };
    const move = event => { const mark = hit(event.containerPoint); map.getContainer().style.cursor = mark ? "pointer" : ""; setHover(mark ? { mark, x: event.containerPoint.x, y: event.containerPoint.y } : null); };
    const click = event => {
      const mark = hit(event.containerPoint); if (!mark) return;
      if (mark.coincident) setCoincident(mark.rows);
      else if (mark.grouped) map.fitBounds(L.latLngBounds(mark.rows.map(row => [row.lat, row.lon])), { padding: [70, 70], maxZoom: Math.max(11, map.getZoom() + 2), animate: false });
      else stateRef.current.onSelect(mark.row);
    };
    map.on("move resize zoom", schedule); map.on("zoomend", () => { setZoom(map.getZoom()); setHover(null); }); map.on("mousemove", move); map.on("click", click); map.on("mouseout", () => setHover(null));
    const observer = new ResizeObserver(() => { map.invalidateSize(); schedule(); }); observer.observe(host.current);
    map.fitBounds(COUNTRY_BOUNDS, { padding: [24, 24], animate: false });
    schedule();
    return () => { observer.disconnect(); cancelAnimationFrame(scheduled); drawRef.current = null; mapRef.current = null; map.remove(); };
  }, []);

  useEffect(() => { drawRef.current?.(); setHover(null); setCoincident([]); }, [rows, selected, colorMode]);
  useEffect(() => {
    const map = mapRef.current; if (!map || !geography) return;
    if (provinceLayer.current) map.removeLayer(provinceLayer.current);
    provinceLayer.current = L.geoJSON(geography.provinces, { style: { color: "#b8bfd3", weight: .8, fillColor: "#f8f9fc", fillOpacity: 1 }, interactive: false }).addTo(map);
    provinceLayer.current.bringToBack();
  }, [geography]);
  useEffect(() => {
    const map = mapRef.current; if (!map || !geography) return;
    for (const layer of [sbLayer.current, ccaaLayer.current]) if (layer) map.removeLayer(layer);
    sbLayer.current = null; ccaaLayer.current = null;
    const source = territoryMode === "sb" ? geography.sanibricks : geography.ccaa;
    const features = source.features.filter(f => territoryMode === "sb" ?
      (effectiveSB ? f.properties.sb_code === effectiveSB : !ccaaCode || ccaaCode === "all" || f.properties.ccaa === ccaaCode) :
      (!ccaaCode || ccaaCode === "all" || f.properties.ccaa === ccaaCode));
    const territoryElements = [];
    const layer = L.geoJSON({ type: "FeatureCollection", features }, {
      renderer: L.svg(),
      style: f => ({ color: f.properties.sb_priority_national ? "#E6007A" : "#586e9b", weight: effectiveSB ? 2 : territoryMode === "ccaa" ? 1.2 : .65, opacity: .8, fillColor: f.properties.sb_priority_national ? "#E6007A" : "#c8cfe4", fillOpacity: effectiveSB ? .12 : .035 }),
      onEachFeature: (feature, polygon) => {
        const p = feature.properties;
        const selectTerritory = () => territoryMode === "sb" ? stateRef.current.onSB(p.sb_code) : stateRef.current.onCCAA(p.ccaa);
        polygon.on("add", () => {
          const element = polygon.getElement();
          if (!element) return;
          element.setAttribute("role", "button");
          element.setAttribute("tabindex", territoryElements.length ? "-1" : "0");
          territoryElements.push(element);
          element.setAttribute("aria-label", territoryMode === "sb" ? `Seleccionar SB ${p.sb_code}` : `Seleccionar CCAA ${p.ccaa}`);
          element.addEventListener("keydown", event => {
            if (event.key === "Enter" || event.key === " ") { event.preventDefault(); selectTerritory(); }
            else if (["ArrowLeft", "ArrowRight", "ArrowUp", "ArrowDown", "Home", "End"].includes(event.key)) {
              event.preventDefault(); event.stopPropagation();
              const i = territoryElements.indexOf(element), n = territoryElements.length;
              const next = event.key === "Home" ? 0 : event.key === "End" ? n - 1 :
                (i + (["ArrowLeft", "ArrowUp"].includes(event.key) ? -1 : 1) + n) % n;
              territoryElements.forEach((node,j) => node.setAttribute("tabindex", j === next ? "0" : "-1"));
              territoryElements[next].focus();
            }
          });
        });
        polygon.bindTooltip(territoryMode === "sb" ? '<strong>' + escape(p.sb_code) + ' · ' + escape(p.sb_name) + '</strong><br/>Mercado RX agregado: ' + eur(p.original_eur) + '<br/>' + num(p.national_AB_n) + ' A/B nacionales · ' + eur(p.national_AB_eur) + (p.sb_priority_national ? '<br/>SB de visita prioritaria' : '') : '<strong>' + escape(p.ccaa) + '</strong><br/>' + num(p.n_stores) + ' establecimientos · ' + eur(p.market_eur), { sticky: true, className: "market-territory-tooltip" });
        polygon.on("click", event => {
          const point = event.containerPoint || map.mouseEventToContainerPoint(event.originalEvent);
          if (marksRef.current.some(mark => Math.hypot(mark.x-point.x,mark.y-point.y) <= mark.r+4)) return;
          if (event.originalEvent) L.DomEvent.stopPropagation(event.originalEvent);
          selectTerritory();
        });
      }
    }).addTo(map);
    if (territoryMode === "sb") sbLayer.current = layer; else ccaaLayer.current = layer;
  }, [geography, territoryMode, effectiveSB, ccaaCode]);
  useEffect(() => {
    const map = mapRef.current; if (!map || !rows.length) return;
    const valid = rows.filter(row => Number.isFinite(row.lat) && Number.isFinite(row.lon));
    if (valid.length) map.fitBounds(L.latLngBounds(valid.map(row => [row.lat, row.lon])), { padding: [35, 35], maxZoom: valid.length === 1 ? 15 : 13, animate: false });
  }, [fitKey]);
  useEffect(() => { if (selected && mapRef.current && Number.isFinite(selected.lat)) mapRef.current.setView([selected.lat, selected.lon], Math.max(13, mapRef.current.getZoom()), { animate: false }); }, [selected?.id]);

  return <div className="market-map-frame" data-reviewed-rows>
    <div className="market-map-toolbar">
      <div className="market-region-control"><MarketIcon name="crosshair" size={15}/><select aria-label="Vista territorial del mapa" defaultValue="España completa" onChange={event => mapRef.current?.fitBounds(REGIONS[event.target.value], { padding: [28, 28], animate: false })}>{Object.keys(REGIONS).map(region => <option key={region}>{region}</option>)}</select></div>
      <div className="market-layer-controls"><label>Límites <select aria-label="División territorial del mapa" value={territoryMode} onChange={e => setTerritoryMode(e.target.value)}><option value="ccaa">CCAA</option><option value="sb">Sanibricks</option></select></label></div>
    </div>
    <div className="market-leaflet-host" ref={host} role="region" aria-label="Mapa interactivo de España con farmacias. Usa la tabla para seleccionar con el teclado."/>
    <div className="market-map-context"><span className="market-live-dot"/>{zoom < 10 ? colorMode === "category" ? "Agrupaciones · color: categoría más frecuente" : "Agrupaciones · color: potencial medio por punto" : "Establecimientos · tamaño según potencial"}{effectiveSB ? <span className="market-map-context-sb">SB {effectiveSB}</span> : null}</div>

    {!rows.length && <div className="market-map-empty"><MarketIcon name="search" size={24}/><strong>Sin establecimientos en esta selección</strong><span>Amplía los filtros para explorar el territorio.</span></div>}
    {error && <div className="market-map-error" role="alert">{error}. Los puntos disponibles siguen visibles.</div>}
    {hover && <div className="market-map-hover" ref={hoverCardRef} style={{ left: Math.min(hover.x + 18, Math.max(8, (host.current?.clientWidth || 600) - 285)), top: Math.min(hover.y + 15, Math.max(8, (host.current?.clientHeight || 600) - hoverHeight - 12)) }}>
      {hover.mark.grouped ? <><span className="market-eyebrow">{hover.mark.coincident ? "UBICACIÓN COMPARTIDA" : "AGRUPACIÓN VISUAL"}</span><strong>{num(hover.mark.rows.length)} establecimientos</strong><div className="market-tooltip-value">{eur(hover.mark.potential)}</div><span>Potencial anual estimado · jul 2025–jun 2026</span><div className="market-tooltip-row"><span>Población atribuida 65+</span><b>{pop(hover.mark.pop65)}</b></div><div className="market-tooltip-row"><span>Ubicaciones provisionales</span><b>{num(hover.mark.provisional)}</b></div><div className="market-tooltip-categories">{Object.entries(hover.mark.counts).map(([category, count]) => <span key={category}><i style={{ background: CATEGORY_COLORS[category] }}/>{category} · {num(count)}</span>)}</div><small>{colorMode === "category" ? "Color: categoría más frecuente por número de puntos." : "Color: potencial medio de los puntos, escala nacional."}</small><small>{hover.mark.coincident ? "Haz clic para elegir un establecimiento." : "Haz clic para ampliar."}</small></> : <><span className="market-eyebrow">ESPAÑA {hover.mark.row.category_national} · CCAA {hover.mark.row.category_ccaa} · SB {hover.mark.row.sb_code}</span><strong>{hover.mark.row.name}</strong><span>{hover.mark.row.geographic_scope} · {hover.mark.row.rurality} · {hover.mark.row.visit_eligible?"Visitable":"Sin visita"}</span><div className="market-tooltip-value">{eur(hover.mark.row.potential_eur)}</div><span>Potencial anual estimado · jul 2025–jun 2026</span><div className="market-tooltip-pops">{[["65+", hover.mark.row.pop65], ["75+", hover.mark.row.pop75], ["85+", hover.mark.row.pop85]].map(([label, value]) => <div key={label}><b>{pop(value)}</b><span>{label}</span></div>)}</div><div className="market-tooltip-row"><span>Cuota sobre mercado de su SB</span><b>{pct(hover.mark.row.share)}</b></div>{hover.mark.row.location_provisional && <small className="market-quality-note">Ubicación estimada · no es un punto preciso</small>}<small>Haz clic para abrir la ficha.</small></>}
    </div>}
    {!!coincident.length && <div className="market-coincident-picker" role="dialog" aria-label="Elegir establecimiento con ubicación compartida"><div><strong>{num(coincident.length)} puntos en esta ubicación</strong><button type="button" aria-label="Cerrar establecimientos coincidentes" onClick={() => setCoincident([])}><MarketIcon name="close"/></button></div><p>Selecciona una farmacia para abrir su ficha. Las ubicaciones estimadas conservan su aviso.</p><ul>{coincident.map(row => <li key={row.id}><button type="button" onClick={() => { onSelect(row); setCoincident([]); }}><span><b>{row.name}</b><small>{row.location_provisional ? "Ubicación estimada" : "Coordenada validada"}</small></span><strong>{eur(row.potential_eur)}</strong></button></li>)}</ul></div>}
    <div className="market-map-legend">{colorMode === "category" ? <div className="market-category-legend">{Object.entries(CATEGORY_COLORS).map(([category, color]) => <span key={category}><i style={{ background: color }}/>{category}</span>)}<em>Categoría {reference === "national" ? "nacional" : "CCAA"}</em></div> : <div className="market-continuous-legend"><span>0 €</span><i/><span>{eur(scaleMax)}</span><em>Escala nacional logarítmica</em></div>}<span className="market-proxy-legend"><i/> Ubicación estimada</span></div>
  </div>;
}
