/** METRAP metin ayrıştırma — Python satır 1622–1836 ile uyumlu */

const ZONE_OFFSETS = [28, 41, 54, 67, 80, 93, 106, 119, 132, 145, 158];

function decodeTempOrDens(raw3) {
  const s = String(raw3);
  if (s[0] === "0") {
    return parseInt("1" + s, 10) / 10;
  }
  return parseInt(s, 10) / 10;
}

/**
 * @param {string} text
 * @returns {{
 *   stationElevation: number,
 *   surfaceDensity: number,
 *   reportType: string,
 *   zones: Array<{dir:number,speed:number,temp:number,dens:number}>
 * }}
 */
export function parseMetrap(text) {
  const lines = String(text)
    .replace(/\r\n/g, "\n")
    .split("\n")
    .map((l) => l.trim())
    .filter((l) => l.length > 0);

  if (lines.length < 2) {
    throw new Error("METRAP verisi yetersiz");
  }

  // İlk satırı atla, kalanı \n ile birleştir (Python offset'leri buna göre)
  const birlesik = lines.slice(1).join("\n");

  const stationElevation = parseInt(birlesik.slice(19, 22), 10) * 10;
  const surfaceDensity = parseInt(birlesik.slice(22, 25), 10) / 10;
  const reportType = birlesik.slice(0, 5);

  const zones = [];
  for (let z = 0; z <= 10; z++) {
    const s = ZONE_OFFSETS[z];
    const dir = parseInt(birlesik.slice(s, s + 2), 10) * 100;
    const speed = parseInt(birlesik.slice(s + 2, s + 4), 10);
    const temp = decodeTempOrDens(birlesik.slice(s + 4, s + 7));
    const dens = decodeTempOrDens(birlesik.slice(s + 7, s + 10));
    zones.push({ dir, speed, temp, dens });
  }

  return { stationElevation, surfaceDensity, reportType, zones };
}

export function getZoneWeather(parsed, zone) {
  const z = Math.max(0, Math.min(10, Math.round(Number(zone) || 0)));
  return parsed.zones[z];
}
