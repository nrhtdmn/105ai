/**
 * MGRS → lat/lon + internet rakım (open-elevation / opentopodata).
 */
import { toPoint } from "./vendor/mgrs.js";

/** 37 + SFA + 19072 + 60022 → 37SFA1907260022 */
export function buildMgrs(bolgeNu, bolgeSayisi, sag, yukari) {
  const zone = String(bolgeNu || "").trim();
  const sq = String(bolgeSayisi || "").trim().toUpperCase();
  const e = String(Math.round(Number(sag))).padStart(5, "0");
  const n = String(Math.round(Number(yukari))).padStart(5, "0");
  if (!zone || !sq || e.length > 5 || n.length > 5) {
    throw new Error("MGRS için bölge / sağ / yukarı eksik veya hatalı.");
  }
  return `${zone}${sq}${e}${n}`;
}

export function mgrsToLatLon(mgrsStr) {
  const [lon, lat] = toPoint(String(mgrsStr).replace(/\s/g, ""));
  return { lat, lon };
}

async function fetchOpenElevation(lat, lon) {
  const url = `https://api.open-elevation.com/api/v1/lookup?locations=${lat},${lon}`;
  const res = await fetch(url);
  if (!res.ok) throw new Error("open-elevation HTTP " + res.status);
  const data = await res.json();
  const elev = data?.results?.[0]?.elevation;
  if (elev == null || Number.isNaN(Number(elev))) throw new Error("Rakım sonucu yok");
  return Math.round(Number(elev));
}

async function fetchOpenTopo(lat, lon) {
  const url = `https://api.opentopodata.org/v1/aster30m?locations=${lat},${lon}`;
  const res = await fetch(url);
  if (!res.ok) throw new Error("opentopodata HTTP " + res.status);
  const data = await res.json();
  const elev = data?.results?.[0]?.elevation;
  if (elev == null || Number.isNaN(Number(elev))) throw new Error("Rakım sonucu yok");
  return Math.round(Number(elev));
}

/**
 * Koordinattan rakım (metre). Önce open-elevation, olmazsa opentopodata.
 */
export async function fetchElevationForTarget({ bolgeNu, bolgeSayisi, sag, yukari }) {
  const mgrs = buildMgrs(bolgeNu, bolgeSayisi, sag, yukari);
  const { lat, lon } = mgrsToLatLon(mgrs);
  try {
    const elev = await fetchOpenElevation(lat, lon);
    return { elev, lat, lon, mgrs, source: "open-elevation" };
  } catch (e1) {
    try {
      const elev = await fetchOpenTopo(lat, lon);
      return { elev, lat, lon, mgrs, source: "opentopodata" };
    } catch (e2) {
      throw new Error(
        `Rakım alınamadı (${mgrs}). Ağ / CORS engeli olabilir. Elle girin. (${e1.message}; ${e2.message})`
      );
    }
  }
}

/** Excel uyumlu CSV (BOM) indir */
export function downloadCsv(filename, rows) {
  const escape = (v) => {
    const s = v == null ? "" : String(v);
    if (/[",\n\r;]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
    return s;
  };
  // Excel TR: noktalı virgül ayırıcı daha sorunsuz açılır
  const lines = rows.map((r) => r.map(escape).join(";"));
  const blob = new Blob(["\uFEFF" + lines.join("\r\n")], {
    type: "text/csv;charset=utf-8",
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename.endsWith(".csv") ? filename : filename + ".csv";
  a.click();
  URL.revokeObjectURL(a.href);
}
