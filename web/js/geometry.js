/** Geometri: plan mesafe, azimut (milyem), yan, barut hakkı */

export function planRange(sag1, yuk1, sag2, yuk2) {
  return Math.round(Math.sqrt((sag1 - sag2) ** 2 + (yuk1 - yuk2) ** 2));
}

/**
 * Batarya→hedef istikamet açısı (0–6400 milyem).
 * Python barut_hakki_oner ile aynı bölge düzeltmeleri.
 */
export function azimuthMil(gunSag, gunYuk, tgtSag, tgtYuk) {
  const dx = Math.abs(gunSag - tgtSag);
  const dy = Math.abs(gunYuk - tgtYuk);
  let ia = Math.round((Math.atan2(dx, dy) * 3200) / Math.PI);

  if (gunSag > tgtSag && gunYuk > tgtYuk) {
    // 3. bölge
    ia += 3200;
  } else if (gunSag > tgtSag && gunYuk < tgtYuk) {
    // 4. bölge
    ia = 6400 - ia;
  } else if (gunSag < tgtSag && gunYuk < tgtYuk) {
    // 1. bölge — değişmez
  } else if (gunSag < tgtSag && gunYuk > tgtYuk) {
    // 2. bölge
    ia = 3200 - ia;
  }
  return ((ia % 6400) + 6400) % 6400;
}

/**
 * Plan yan = müşterek yan ± |azimut − AHİA|
 * AHİA < azimut → musyan − fark; aksi halde musyan + fark; ≤0 ise +6400
 */
export function baseYan(ahia, musyan, azimut) {
  const fark = Math.abs(azimut - ahia);
  let yan;
  if (ahia < azimut) {
    yan = musyan - fark;
  } else {
    yan = musyan + fark;
  }
  if (yan <= 0) yan = 6400 + yan;
  return yan;
}

/**
 * @param {number} rangeM
 * @param {number[]} breaks — örn. [2348,2766,3354,4218,5363,6694,11000]
 * @returns {number|string}
 */
export function recommendCharge(rangeM, breaks = [2348, 2766, 3354, 4218, 5363, 6694, 11000]) {
  const r = Number(rangeM) || 0;
  if (r === 0) return 0;
  for (let i = 0; i < breaks.length; i++) {
    if (r <= breaks[i]) return i + 1;
  }
  return "Atış Yapılamaz...";
}
