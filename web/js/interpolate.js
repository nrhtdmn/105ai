/** Lineer interpolasyon ve tablo yoğunlaştırma (Python veri_kumesi_interpolasyonu) */

export function linearInterpolate(x, x1, y1, x2, y2) {
  if (x2 === x1) return y1;
  return y1 + ((x - x1) * (y2 - y1)) / (x2 - x1);
}

/**
 * Tam sayı x ekseninde her metre için satır üret.
 * @param {Array<Array<number>>} rows — her satır [x, ...y]
 * @returns {Map<number, number[]>} yuvarlanmış int x → satır dizisi
 */
export function densifyTable(rows) {
  const map = new Map();
  if (!rows || rows.length === 0) return map;

  const sorted = rows.map((r) => r.slice()).sort((a, b) => a[0] - b[0]);

  for (let i = 0; i < sorted.length - 1; i++) {
    const x1 = sorted[i][0];
    const y1 = sorted[i].slice(1);
    const x2 = sorted[i + 1][0];
    const y2 = sorted[i + 1].slice(1);
    const ix1 = Math.round(x1);
    map.set(ix1, [ix1, ...y1]);

    for (let j = ix1 + 1; j < Math.round(x2); j++) {
      const ys = y1.map((v, k) => linearInterpolate(j, x1, v, x2, y2[k]));
      map.set(j, [j, ...ys]);
    }
  }

  const last = sorted[sorted.length - 1];
  const ixLast = Math.round(last[0]);
  map.set(ixLast, [ixLast, ...last.slice(1)]);
  return map;
}

/** İki noktalı (x,y) tabloları için metre yoğunlaştırma (isi/yogunluk/barut). */
export function densifyPairs(pairs) {
  const map = new Map();
  if (!pairs || pairs.length === 0) return map;
  const sorted = pairs.map((p) => [p[0], p[1]]).sort((a, b) => a[0] - b[0]);

  for (let i = 0; i < sorted.length - 1; i++) {
    const [x1, y1] = sorted[i];
    const [x2, y2] = sorted[i + 1];
    const diffX = x2 - x1;
    for (let j = 0; j < diffX; j++) {
      const x = x1 + j;
      const y = y1 + (j * (y2 - y1)) / diffX;
      map.set(Math.round(x), y);
    }
  }
  const last = sorted[sorted.length - 1];
  map.set(Math.round(last[0]), last[1]);
  return map;
}

export function lookupExact(map, x) {
  const key = Math.round(Number(x));
  return map.has(key) ? map.get(key) : null;
}

export function lookupNearest(map, x) {
  const key = Math.round(Number(x));
  if (map.has(key)) return map.get(key);
  let best = null;
  let bestDist = Infinity;
  for (const k of map.keys()) {
    const d = Math.abs(k - key);
    if (d < bestDist) {
      bestDist = d;
      best = map.get(k);
    }
  }
  return best;
}
