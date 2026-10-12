// Official Essity digital palette; letters retain their economic meaning.
export const CATEGORY_COLORS = { A: "#00005A", B: "#E6007A", C: "#00499F", D: "#888B8D" };
export const CATEGORY_TEXT_COLORS = { A: "#00005A", B: "#A90059", C: "#00499F", D: "#51566A" };
export const CATEGORY_DARK_TEXT_COLORS = { A: "#A8B5FF", B: "#FF80C4", C: "#66C2FF", D: "#C2C6D5" };
export function markLabelColor(fill) {
  const channels = fill.startsWith("#") ? [1, 3, 5].map(i => parseInt(fill.slice(i, i + 2), 16)) : fill.match(/[\d.]+/g).slice(0, 3).map(Number);
  const luminance = rgb => rgb.map(c => { const v = c / 255; return v <= .04045 ? v / 12.92 : ((v + .055) / 1.055) ** 2.4; }).reduce((sum, c, i) => sum + c * [.2126, .7152, .0722][i], 0);
  const bg = luminance(channels), navy = luminance([0, 0, 90]);
  const whiteRatio = 1.05 / (bg + .05), navyRatio = (Math.max(bg, navy) + .05) / (Math.min(bg, navy) + .05);
  if (Math.max(whiteRatio, navyRatio) < 4.5) return "#000000";
  return whiteRatio >= navyRatio ? "#ffffff" : "#00005A";
}
const integer = new Intl.NumberFormat("es-ES", { maximumFractionDigits: 0 });
const decimal = new Intl.NumberFormat("es-ES", { maximumFractionDigits: 1 });
const money = new Intl.NumberFormat("es-ES", { style: "currency", currency: "EUR", maximumFractionDigits: 0 });
const preciseMoney = new Intl.NumberFormat("es-ES", { style: "currency", currency: "EUR", minimumFractionDigits: 2, maximumFractionDigits: 2 });
export const num = value => integer.format(Number(value) || 0);
export const pop = value => decimal.format(Number(value) || 0);
export const eur = value => money.format(Number(value) || 0);
export const eurExact = value => preciseMoney.format(Number(value) || 0);
export const pct = value => `${decimal.format((Number(value) || 0) * 100)} %`;
export const shortEur = value => value >= 1e6 ? `${decimal.format(value / 1e6)} M€` : value >= 1000 ? `${decimal.format(value / 1000)} mil €` : eur(value);
export const normalize = value => String(value ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLocaleLowerCase("es");
export function potentialColor(value, max) {
  const t = Math.min(1, Math.max(0, Math.log1p(Number(value) || 0) / Math.log1p(max || 1)));
  const from = [224, 230, 248], to = [0, 0, 90];
  return `rgb(${from.map((c, i) => Math.round(c + (to[i] - c) * t)).join(",")})`;
}
export function exportRows(rows) {
  return exportTable(rows, [["id","Establecimiento"],["name","Nombre"],["iqvia","IQVIA"],["ccaa","CCAA"],["province","Provincia"],["municipality","Municipio"],["postal_code","CP"],["sb_code","SB código"],["sb_name","SB nombre"],["sscc","SSCC"],["category_national","Categoría nacional"],["category_ccaa","Categoría CCAA"],["dual_category_disagreement","Letras distintas"],["sb_priority_national","SB de visita prioritaria nacional"],["national_AB_n","Puntos A/B nacionales en SB completo"],["national_AB_eur","Mercado A/B nacional en SB completo EUR"],["potential_eur","Potencial anual estimado EUR"],["share","Proporción mercado SB ubicación"],["sb_original_eur","Mercado RX agregado SB EUR"],["pop65","Población atribuida 65+"],["pop75","Población atribuida 75+"],["pop85","Población atribuida 85+"],["pop_65_69","Población 65-69"],["pop_70_74","Población 70-74"],["pop_75_79","Población 75-79"],["pop_80_84","Población 80-84"],["pop_85_mas","Población 85+"],["weighted_population","Demanda ponderada"],["sb_weighted_population","Demanda total puntos SB"],["demographic_share","Proporción demanda ponderada SB"],["geographic_scope","Ámbito rural / no rural"],["rurality","DEGURBA original"],["degurba","DEGURBA código"],["visit_eligible","Elegible visita no rural"],["allocation_method","Método de reparto"],["allocation_weight","Peso aplicado"],["allocation_share","Proporción aplicada"],["rural_group_key","Grupo rural municipio-SB"],["rural_group_count","Puntos del grupo rural"],["rural_group_demand","Demanda grupo rural"],["visitable_national_AB_n","A/B no rurales del SB"],["visitable_national_AB_eur","Mercado A/B no rural del SB EUR"],["lat","Latitud visualización"],["lon","Longitud visualización"],["location_provisional","Ubicación estimada"],["provisional","Resultado provisional"],["location_method","Método de ubicación"],["potential_v1_eur","Potencial histórico V1 EUR"],["category_v1","Categoría histórica V1"],["sensitive_v1","Sensibilidad espacial histórica V1"],["sensitivity_variation_v1","Variación espacial histórica V1"],["near_national_cut_5pct","Cerca de corte nacional"],["near_ccaa_cut_5pct","Cerca de corte CCAA"]], `farmacias_seleccion_${new Date().toISOString().slice(0,10)}.csv`);
}
export function exportTable(rows, columns, filename) {
  const cell = value => {
    if (typeof value === "number") return Number.isFinite(value) ? String(value).replace(".", ",") : "";
    let safe = String(value ?? "");
    if (/^[=+@\-\t\r]/.test(safe)) safe = `'${safe}`;
    return `"${safe.replaceAll('"', '""')}"`;
  };
  const csv = [columns.map(([, label]) => cell(label)).join(";"), ...rows.map(row => columns.map(([field]) => cell(row[field])).join(";"))].join("\r\n");
  const url = URL.createObjectURL(new Blob(["\uFEFF", csv], { type: "text/csv;charset=utf-8;" }));
  const anchor = document.createElement("a"); anchor.href = url; anchor.download = filename; anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
