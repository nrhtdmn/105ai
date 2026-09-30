/**
 * Ateş esası hesaplama — charge-5 boru hattı, lib JSON ile parametreli.
 * Python atesidare.py ~2598–3730 metro/TAÇ aritmetiği ile hizalı.
 */
import { densifyTable, densifyPairs, lookupExact, lookupNearest } from "./interpolate.js";
import { parseMetrap, getZoneWeather } from "./metro.js";
import { planRange, azimuthMil, baseYan, recommendCharge } from "./geometry.js";

const GUN_IDS = ["1A", "1B", "1C", "1D", "2A", "2B", "2C", "2D"];
const GROUP1 = ["1A", "1B", "1C", "1D"];
const GROUP2 = ["2A", "2B", "2C", "2D"];
const MIL_CONST = 1.0186;
const STD_TEMP = 100;
const STD_DENS = 100;
const EARTH_RANGE_FACTOR = 0.77;

/** Yükseklik farkı → altitudeZone / complementaryRange sütun indeksi (Python özel: 800≡700) */
function heightColumnIndex(heightRound100) {
  const h = heightRound100;
  if (h === -400) return 1;
  if (h === -300) return 2;
  if (h === -200) return 3;
  if (h === -100) return 4;
  if (h === 0 || h === -0) return 5;
  if (h === 100) return 6;
  if (h === 200) return 7;
  if (h === 300) return 8;
  if (h === 400) return 9;
  if (h === 500) return 10;
  if (h === 600) return 11;
  if (h === 700 || h === 800) return 12;
  if (h === 900) return 13;
  if (h === 1000) return 14;
  return null;
}

/** Arz dönüşü yan sütunu (azimut 400'e yuvarlanmış) */
function earthYanCol(az400) {
  const a = az400;
  if (a === 0 || a === 6400) return 1;
  if (a === 400 || a === 6000) return 2;
  if (a === 800 || a === 5600) return 3;
  if (a === 1200 || a === 5200) return 4;
  if (a === 1600 || a === 4800) return 5;
  if (a === 2000 || a === 4400) return 6;
  if (a === 2400 || a === 4000) return 7;
  if (a === 2800 || a === 3600) return 8;
  if (a === 3200) return 9;
  return null;
}

/** Arz dönüşü mesafe sütunu — Python koşulları (çoğu 200'lük adım asla eşleşmez) */
function earthRangeCol(az400) {
  const a = az400;
  if (a === 0 || a === 3200 || a === 6400) return 1;
  if (a === 200 || a === 3000) return 2;
  if (a === 400 || a === 2800) return 3;
  if (a === 600 || a === 2600) return 4;
  if (a === 800 || a === 2400) return 5;
  if (a === 1000 || a === 2200) return 6;
  if (a === 1200 || a === 2000) return 7;
  if (a === 1400 || a === 1800) return 8;
  if (a === 1600) return 9;
  if (a === 3400 || a === 6200) return 11;
  if (a === 3600 || a === 6000) return 12;
  if (a === 3800 || a === 5800) return 13;
  if (a === 4000 || a === 5600) return 14;
  if (a === 4200 || a === 5400) return 15;
  if (a === 4400 || a === 5200) return 16;
  if (a === 4600 || a === 5000) return 17;
  if (a === 4800) return 18;
  return null;
}

function round1(n) {
  return Math.round(Number(n) * 10) / 10;
}

function roundN(n, digits) {
  const f = 10 ** digits;
  return Math.round(Number(n) * f) / f;
}

function pyRound(n) {
  // Python 3 round half-even; basitçe Math.round yeter (çoğu durumda aynı)
  return Math.round(Number(n));
}

const _densCache = new WeakMap();

function getDensified(lib) {
  let c = _densCache.get(lib);
  if (c) return c;
  c = {
    altitude: densifyTable(lib.altitudeZoneTable || []),
    complementary: densifyTable(lib.complementaryRange || []),
    wind: densifyTable(lib.windComponents || []),
    ground: densifyTable(lib.groundFire || []),
    earthYan: densifyTable(lib.earthRotationYan || []),
    earthRange: densifyTable(lib.earthRotationRange || []),
    fuze: densifyTable(lib.fuzeSecondary || []),
    dropTac: densifyTable(lib.dropAngleTac || []),
    oneMil: densifyTable(lib.oneMilChange || []),
    tempAlt: densifyPairs(lib.tempAltFactor || []),
    densAlt: densifyPairs(lib.densityAltFactor || []),
    propMv: densifyPairs(lib.propellantTempMv || []),
  };
  _densCache.set(lib, c);
  return c;
}

function dogalYanMap(shared, charge) {
  const rows = (shared && shared.dogalYan && shared.dogalYan[String(charge)]) || [];
  return densifyPairs(rows);
}

/**
 * Tek grup (1B veya 2B referans) için tam ateş esası.
 */
function computeGroupSolution({
  charge,
  lib,
  shared,
  refGun,
  refId,
  planM,
  azimut,
  baseYanVal,
  targetRakim,
  barutIsisi,
  squareWeight,
  metroText,
  paralanmaYuksekligi = 0,
  mevziHiz = 0,
}) {
  const dens = getDensified(lib);
  const tapasaniyesi = 1;
  const bataryaRakimRaw = Number(refGun.rakim) || 0;
  const ilkHizFarki = Number(refGun.ihf) || 0;
  let bataryaRakim = pyRound(bataryaRakimRaw / 10) * 10;

  const paralanmaRakim = Number(targetRakim) + Number(paralanmaYuksekligi || 0);
  const yukseklikFarki = paralanmaRakim - bataryaRakim;
  const hedefinToptan = pyRound(yukseklikFarki / 100) * 100;
  const col = heightColumnIndex(hedefinToptan);

  const mesafe = Math.round(planM);
  const altRow = lookupNearest(dens.altitude, mesafe);
  let bolge = 0;
  if (altRow && col != null) bolge = Math.round(altRow[col] || 0);

  const parsed = parseMetrap(metroText);
  const wx = getZoneWeather(parsed, bolge);
  let { dir: ruzgarIstikameti, speed: ruzgarHizi, temp: havaSicakligi, dens: havaYogunlugu } = wx;

  const btMetroFark = bataryaRakim - parsed.stationElevation;
  // Exact int match like Python isi_data / yogunluk_data
  let isi_duzeltmesi = dens.tempAlt.has(btMetroFark) ? dens.tempAlt.get(btMetroFark) : 0;
  let yogunluk_duzeltmesi = dens.densAlt.has(btMetroFark) ? dens.densAlt.get(btMetroFark) : 0;

  const duzeltilmisSicaklik = isi_duzeltmesi + havaSicakligi;
  const duzeltilmisYogunluk = yogunluk_duzeltmesi + havaYogunlugu;

  const compRow = lookupNearest(dens.complementary, mesafe);
  let tamamlayici = 0;
  if (compRow && col != null) tamamlayici = Math.trunc(compRow[col] || 0);

  const girisMesafesi = pyRound((mesafe + tamamlayici) / 100) * 100;

  // Rüzgar bileşenleri — atış 100'e yuvarlanır
  let atis100 = pyRound(azimut / 100) * 100;
  let ruzgarDir = ruzgarIstikameti;
  if (ruzgarDir < atis100) ruzgarDir += 6400;
  const ruzgarPlan = ruzgarDir - atis100;

  const windRow = lookupExact(dens.wind, ruzgarPlan) || lookupNearest(dens.wind, ruzgarPlan);
  const yanRuzgarBilesen = windRow ? windRow[1] : 0;
  const mesafeRuzgarBilesen = windRow ? windRow[2] : 0;

  // F cetveli @ giriş mesafesi
  const gf = lookupExact(dens.ground, girisMesafesi) || lookupNearest(dens.ground, girisMesafesi);
  if (!gf) {
    throw new Error(`F cetveli bulunamadı (giriş ${girisMesafesi})`);
  }

  let nisangah = roundN(gf[1], 3);
  let m564ts = roundN(gf[2], 3);
  const deltaTS = roundN(gf[3], 3);
  const ucussuresi = roundN(gf[4], 3);
  let dogalyandz = roundN(gf[5], 3);
  const yanRuzgarFaktor = round1(gf[6]);
  const ilkhizdzl_eksilme = roundN(gf[7], 3);
  const ilkhizdzl_martma = roundN(gf[8], 3);
  const mesaferuzgaridzl_bas = roundN(gf[9], 3);
  const mesaferuzgaridzl_arka = roundN(gf[10], 3);
  const havaisisi_eksilme = roundN(gf[11], 3);
  const havaisisi_artma = roundN(gf[12], 3);
  const havayogunlugu_eksilme = roundN(gf[13], 3);
  const havayogunlugu_artma = roundN(gf[14], 3);
  const mermikareagirligi_eksilme = roundN(gf[15], 3);
  const mermikareagirligi_artma = roundN(gf[16], 3);

  const mesafe_ruzgari = ruzgarHizi * mesafeRuzgarBilesen;
  const ruzgar_yan_duzeltmesi = round1(ruzgarHizi * yanRuzgarBilesen * yanRuzgarFaktor);
  dogalyandz = round1(dogalyandz);

  // Arz dönüşü — azimut 400'e
  const atis400 = pyRound(azimut / 400) * 400;
  let arzindonusuyanduzeltmesi = 0;
  let arzindonusumesafeduzeltmesi = 0;

  const eyCol = earthYanCol(atis400);
  const eyRow = lookupExact(dens.earthYan, girisMesafesi) || lookupNearest(dens.earthYan, girisMesafesi);
  if (eyRow && eyCol != null) arzindonusuyanduzeltmesi = round1(eyRow[eyCol]);

  const erCol = earthRangeCol(atis400);
  const erRow = lookupExact(dens.earthRange, girisMesafesi) || lookupNearest(dens.earthRange, girisMesafesi);
  if (erRow && erCol != null) arzindonusumesafeduzeltmesi = round1(erRow[erCol]);

  let metro_yan_duzeltmesi = round1(arzindonusuyanduzeltmesi + dogalyandz + ruzgar_yan_duzeltmesi);
  metro_yan_duzeltmesi = pyRound(metro_yan_duzeltmesi / 1) * 1;

  // Metro mesafe düzeltmeleri — standart 100
  let mesafe_ruzgari_fark = round1(mesafe_ruzgari - 0);
  let mesafe_ruzgari_durum = 0;
  if (mesafe_ruzgari_fark > 0) mesafe_ruzgari_durum = "BAŞ";
  if (mesafe_ruzgari_fark < 0) mesafe_ruzgari_durum = "ARKA";
  else if (!(mesafe_ruzgari_fark > 0)) mesafe_ruzgari_durum = 0;

  let hava_sicakligi_fark = round1(duzeltilmisSicaklik - STD_TEMP);
  let hava_sicakligi_durum = 0;
  if (hava_sicakligi_fark < 0) hava_sicakligi_durum = "EKSİLME";
  if (hava_sicakligi_fark > 0) hava_sicakligi_durum = "ARTMA";
  else if (!(hava_sicakligi_fark < 0)) hava_sicakligi_durum = 0;

  let hava_yogunlugu_fark = round1(duzeltilmisYogunluk - STD_DENS);
  let hava_yogunlugu_durum = 0;
  if (hava_yogunlugu_fark < 0) hava_yogunlugu_durum = "EKSİLME";
  if (hava_yogunlugu_fark > 0) hava_yogunlugu_durum = "ARTMA";
  else if (!(hava_yogunlugu_fark < 0)) hava_yogunlugu_durum = 0;

  const baselineSq = (shared && shared.meta && shared.meta.squareWeightBaseline) || 2;
  let mermi_kare_agirligi_fark = round1(Number(squareWeight) - baselineSq);
  let merkar_durum = 0;
  if (mermi_kare_agirligi_fark < 0) merkar_durum = "EKSİLME";
  if (mermi_kare_agirligi_fark > 0) merkar_durum = "ARTMA";
  if (mermi_kare_agirligi_fark === 0) merkar_durum = 0;

  let mesafe_ruzgari_duzeltme_birimi =
    mesafe_ruzgari_fark < 0 ? mesaferuzgaridzl_arka : mesaferuzgaridzl_bas;
  if (mesafe_ruzgari_fark === 0) mesafe_ruzgari_duzeltme_birimi = 0;

  let hava_sicakligi_duzeltme_birimi =
    hava_sicakligi_fark < 0 ? havaisisi_eksilme : havaisisi_artma;
  if (hava_sicakligi_fark === 0) hava_sicakligi_duzeltme_birimi = 0;

  let hava_yogunlugu_duzeltme_birimi =
    hava_yogunlugu_fark < 0 ? havayogunlugu_eksilme : havayogunlugu_artma;
  if (hava_yogunlugu_fark === 0) hava_yogunlugu_duzeltme_birimi = 0;

  let mermi_kare_agirligi_duzeltme_birimi = 0;
  if (mermi_kare_agirligi_fark < 0) mermi_kare_agirligi_duzeltme_birimi = mermikareagirligi_eksilme;
  else if (mermi_kare_agirligi_fark > 0) mermi_kare_agirligi_duzeltme_birimi = mermikareagirligi_artma;

  mesafe_ruzgari_fark = Math.abs(mesafe_ruzgari_fark);
  hava_sicakligi_fark = Math.abs(hava_sicakligi_fark);
  hava_yogunlugu_fark = Math.abs(hava_yogunlugu_fark);

  const mes_ruz_dzl = mesafe_ruzgari_fark * mesafe_ruzgari_duzeltme_birimi;
  const hav_sic_dzl = round1(hava_sicakligi_fark * hava_sicakligi_duzeltme_birimi);
  const hav_yog_dzl = round1(hava_yogunlugu_fark * hava_yogunlugu_duzeltme_birimi);
  const mer_kare_dzl = round1(mermi_kare_agirligi_fark * mermi_kare_agirligi_duzeltme_birimi);
  const dun_don_mes_dzl = round1(arzindonusumesafeduzeltmesi * EARTH_RANGE_FACTOR);

  const metro_mesafe_duzeltmesi = pyRound(
    mes_ruz_dzl + hav_sic_dzl + hav_yog_dzl + mer_kare_dzl + dun_don_mes_dzl
  );

  // Barut ısısı → ΔV
  const barutKey = Math.round(Number(barutIsisi));
  let barut__isisi_dzl2 = dens.propMv.has(barutKey) ? round1(dens.propMv.get(barutKey)) : 3;

  const hiz_degisikligi = round1(ilkHizFarki + mevziHiz);
  const deltaV_hiz_farki = round1(hiz_degisikligi + barut__isisi_dzl2);

  let ilk_hiz_duzeltme_birimi = 0;
  if (deltaV_hiz_farki > 0) ilk_hiz_duzeltme_birimi = ilkhizdzl_martma;
  else if (deltaV_hiz_farki < 0) ilk_hiz_duzeltme_birimi = ilkhizdzl_eksilme;

  let deltaV_mesafe_duzeltmesi = Math.abs(round1(deltaV_hiz_farki * ilk_hiz_duzeltme_birimi));
  if (deltaV_hiz_farki > 0) {
    deltaV_mesafe_duzeltmesi = Math.abs(round1(deltaV_hiz_farki * ilk_hiz_duzeltme_birimi)) * -1;
  } else if (deltaV_hiz_farki < 0) {
    deltaV_mesafe_duzeltmesi = Math.abs(round1(deltaV_hiz_farki * ilk_hiz_duzeltme_birimi));
  } else {
    deltaV_mesafe_duzeltmesi = 0;
  }

  const toplam_mesafe_duzeltmesi = pyRound(metro_mesafe_duzeltmesi + deltaV_mesafe_duzeltmesi);

  // Tapa ikincil (tapasaniyesi=1) — Python hesaplar; ana mesafe/yan'a eklenmez
  const fuzeRow = lookupExact(dens.fuze, tapasaniyesi) || lookupNearest(dens.fuze, tapasaniyesi);
  const fuzeSecondary = fuzeRow
    ? {
        ihz: deltaV_hiz_farki < 0 ? fuzeRow[1] : deltaV_hiz_farki > 0 ? fuzeRow[2] : 0,
        wind: mesafe_ruzgari_durum === "BAŞ" ? fuzeRow[3] : mesafe_ruzgari_durum === "ARKA" ? fuzeRow[4] : 0,
        temp: hava_sicakligi_durum === "EKSİLME" ? fuzeRow[5] : hava_sicakligi_durum === "ARTMA" ? fuzeRow[6] : 0,
        dens: hava_yogunlugu_durum === "EKSİLME" ? fuzeRow[7] : hava_yogunlugu_durum === "ARTMA" ? fuzeRow[8] : 0,
        sq: merkar_durum === "EKSİLME" ? fuzeRow[9] : merkar_durum === "ARTMA" ? fuzeRow[10] : 0,
      }
    : null;

  const gac_mesafe = mesafe + toplam_mesafe_duzeltmesi;

  // F cetveli yeniden @ GAC — nişangah / doğal yan
  const gf2 = lookupExact(dens.ground, gac_mesafe) || lookupNearest(dens.ground, gac_mesafe);
  if (gf2) {
    nisangah = round1(gf2[1]);
    dogalyandz = roundN(gf2[5], 3);
    m564ts = roundN(gf2[2], 3);
  }

  const metro_yan_duzeltmesi2 = round1(arzindonusuyanduzeltmesi + dogalyandz + ruzgar_yan_duzeltmesi);
  const toplam_yan_duzeltmesi = pyRound(metro_yan_duzeltmesi2 + mevziHiz);
  const gac_yan_duzeltmesi = pyRound(toplam_yan_duzeltmesi - dogalyandz);

  // TAÇ — düşüş açısı faktörleri plan mesafesinde
  let dtac_arti1 = 0;
  let dtac_eksi1 = 0;
  let dusus_acisi = null;
  let tepe_yuksekligi = null;
  const tacAtGac = lookupExact(dens.dropTac, gac_mesafe) || lookupNearest(dens.dropTac, gac_mesafe);
  if (tacAtGac) {
    dusus_acisi = roundN(tacAtGac[1], 3);
    tepe_yuksekligi = roundN(tacAtGac[2], 3);
  }
  const tacAtPlan = lookupExact(dens.dropTac, mesafe) || lookupNearest(dens.dropTac, mesafe);
  if (tacAtPlan) {
    dtac_arti1 = roundN(tacAtPlan[3], 3);
    dtac_eksi1 = roundN(tacAtPlan[4], 3);
  }

  const hedef_batarya_rakim_farki = Number(targetRakim) - bataryaRakimRaw;
  const dtac = (hedef_batarya_rakim_farki / (mesafe / 1000)) * MIL_CONST;

  let tac = 0;
  let ttac = 0;
  if (dtac < 0) {
    ttac = Math.trunc(dtac * dtac_arti1);
    tac = dtac + ttac;
  } else if (dtac > 0) {
    ttac = Math.trunc(dtac * dtac_eksi1);
    tac = dtac + ttac;
  }

  const oneMilRow = lookupExact(dens.oneMil, gac_mesafe) || lookupNearest(dens.oneMil, gac_mesafe);
  const birmil = oneMilRow ? round1(oneMilRow[1]) : 1;

  const yukselis = round1(nisangah + tac);
  const yanFinal = pyRound(baseYanVal + toplam_yan_duzeltmesi);

  // shared dogalYan (bilgi / Python dogal_yan_1A)
  const dyMap = dogalYanMap(shared, charge);
  const dogalYanShared = dyMap.has(mesafe) ? roundN(dyMap.get(mesafe), 2) : null;

  return {
    refId,
    charge,
    planMesafe: mesafe,
    azimut,
    bolge,
    tamamlayici,
    girisMesafesi,
    gacMesafe: gac_mesafe,
    nisangah,
    yukselis,
    yan: yanFinal,
    yanPlan: baseYanVal,
    tapa: round1(m564ts),
    birmil,
    tac: roundN(tac, 3),
    dtac: roundN(dtac, 3),
    dogalyandz,
    dogalYanShared,
    ucussuresi,
    deltaTS,
    dusus_acisi,
    tepe_yuksekligi,
    metro_yan: metro_yan_duzeltmesi2,
    metro_mesafe: metro_mesafe_duzeltmesi,
    toplam_yan: toplam_yan_duzeltmesi,
    toplam_mesafe: toplam_mesafe_duzeltmesi,
    gac_yan: gac_yan_duzeltmesi,
    ruzgar_yan: ruzgar_yan_duzeltmesi,
    arz_yan: arzindonusuyanduzeltmesi,
    arz_mesafe: arzindonusumesafeduzeltmesi,
    weather: {
      stationElevation: parsed.stationElevation,
      bolge,
      dir: ruzgarIstikameti,
      speed: ruzgarHizi,
      temp: duzeltilmisSicaklik,
      dens: duzeltilmisYogunluk,
    },
    fuzeSecondary,
    deltaV: deltaV_hiz_farki,
    hedefinToptan,
  };
}

function fanGuns(gunIds, guns, plans, yans, refPlan, refYukselis, yanDuzeltme, birmil) {
  const out = {};
  const bm = Number(birmil) || 1;
  for (const id of gunIds) {
    const plan = plans[id];
    const yan0 = yans[id];
    const yan = pyRound(Number(yan0) + Number(yanDuzeltme));
    const dy = (Number(plan) - Number(refPlan)) / bm;
    const yukselis = round1(Number(refYukselis) + dy);
    out[id] = {
      id,
      plan,
      yan,
      yukselis,
      azimut: null,
      charge: null,
    };
  }
  return out;
}

/**
 * Ana giriş: her iki grup için ateş esası.
 */
export function computeFireSolution(opts) {
  const {
    charge,
    lib,
    shared,
    guns,
    target,
    barutIsisi = 75,
    squareWeight = 2,
    metroText,
    paralanmaYuksekligi = 0,
  } = opts;

  if (![5, 6, 7].includes(Number(charge))) {
    throw new Error("Tam atış esası yalnızca 5, 6 ve 7. barut hakkı için hesaplanır.");
  }
  if (!lib) throw new Error("Barut kütüphanesi yok");
  if (!metroText) throw new Error("METRAP metni yok");

  const tgtSag = Number(target.sag);
  const tgtYuk = Number(target.yukari);
  const tgtRakim = Number(target.rakim);

  const plans = {};
  const yans = {};
  const azimuts = {};

  for (const id of GUN_IDS) {
    const g = guns[id];
    plans[id] = planRange(g.sag, g.yukari, tgtSag, tgtYuk);
  }

  // Grup azimutları 1B / 2B üzerinden
  azimuts.g1 = azimuthMil(guns["1B"].sag, guns["1B"].yukari, tgtSag, tgtYuk);
  azimuts.g2 = azimuthMil(guns["2B"].sag, guns["2B"].yukari, tgtSag, tgtYuk);

  for (const id of GROUP1) {
    const g = guns[id];
    yans[id] = baseYan(g.ahia, g.musyan, azimuts.g1);
  }
  for (const id of GROUP2) {
    const g = guns[id];
    yans[id] = baseYan(g.ahia, g.musyan, azimuts.g2);
  }

  const sol1 = computeGroupSolution({
    charge: Number(charge),
    lib,
    shared,
    refGun: guns["1B"],
    refId: "1B",
    planM: plans["1B"],
    azimut: azimuts.g1,
    baseYanVal: yans["1B"],
    targetRakim: tgtRakim,
    barutIsisi,
    squareWeight,
    metroText,
    paralanmaYuksekligi,
  });

  const sol2 = computeGroupSolution({
    charge: Number(charge),
    lib,
    shared,
    refGun: guns["2B"],
    refId: "2B",
    planM: plans["2B"],
    azimut: azimuts.g2,
    baseYanVal: yans["2B"],
    targetRakim: tgtRakim,
    barutIsisi,
    squareWeight,
    metroText,
    paralanmaYuksekligi,
  });

  const fans1 = fanGuns(GROUP1, guns, plans, yans, sol1.planMesafe, sol1.yukselis, sol1.toplam_yan, sol1.birmil);
  const fans2 = fanGuns(GROUP2, guns, plans, yans, sol2.planMesafe, sol2.yukselis, sol2.toplam_yan, sol2.birmil);

  // Referans kartlarını çözüm değerleriyle güncelle
  fans1["1B"].yan = sol1.yan;
  fans1["1B"].yukselis = sol1.yukselis;
  fans2["2B"].yan = sol2.yan;
  fans2["2B"].yukselis = sol2.yukselis;

  const results = { ...fans1, ...fans2 };
  for (const id of GROUP1) {
    results[id].azimut = azimuts.g1;
    results[id].charge = charge;
    results[id].tapa = sol1.tapa;
    results[id].plan = plans[id];
  }
  for (const id of GROUP2) {
    results[id].azimut = azimuts.g2;
    results[id].charge = charge;
    results[id].tapa = sol2.tapa;
    results[id].plan = plans[id];
  }

  return {
    charge: Number(charge),
    guns: results,
    group1: sol1,
    group2: sol2,
    esaslar: {
      plan_yan: sol1.yanPlan,
      yukselis: sol1.yukselis,
      mesafe: sol1.planMesafe,
      barut: charge,
      istikamet: sol1.azimut,
      ucus: sol1.ucussuresi,
      nisangah: sol1.nisangah,
      dogal_yan: sol1.dogalyandz,
      yuz_m: sol1.planMesafe / 1000,
      yirmi_m: sol1.birmil,
      dusus: sol1.dusus_acisi,
      tepe: sol1.tepe_yuksekligi,
      dtac: sol1.dtac,
      tac: sol1.tac,
      gac_yan: sol1.gac_yan,
      toplam_yan: sol1.toplam_yan,
      toplam_mesafe: sol1.toplam_mesafe,
      toplam_tapa: sol1.deltaTS,
      metro_yan: sol1.metro_yan,
      metro_mesafe: sol1.metro_mesafe,
      gac_mesafe: sol1.gacMesafe,
      tapa: sol1.tapa,
    },
  };
}

/** barut_hakki_oner eşdeğeri */
export function recommendAndGeometry(inputs) {
  const { guns, target, shared } = inputs;
  const breaks = (shared && shared.meta && shared.meta.chargeRecommendBreaks) || [
    2348, 2766, 3354, 4218, 5363, 6694, 11000,
  ];
  const tgtSag = Number(target.sag);
  const tgtYuk = Number(target.yukari);

  const plans = {};
  for (const id of GUN_IDS) {
    const g = guns[id];
    plans[id] = planRange(g.sag, g.yukari, tgtSag, tgtYuk);
  }

  const charge1 = recommendCharge(plans["1A"], breaks);
  const charge2 = recommendCharge(plans["2A"], breaks);
  const az1 = azimuthMil(guns["1B"].sag, guns["1B"].yukari, tgtSag, tgtYuk);
  const az2 = azimuthMil(guns["2B"].sag, guns["2B"].yukari, tgtSag, tgtYuk);

  const yans = {};
  for (const id of GROUP1) {
    yans[id] = baseYan(guns[id].ahia, guns[id].musyan, az1);
  }
  for (const id of GROUP2) {
    yans[id] = baseYan(guns[id].ahia, guns[id].musyan, az2);
  }

  return {
    plans,
    yans,
    azimut1: az1,
    azimut2: az2,
    charge1,
    charge2,
    recommend: charge1,
  };
}

/**
 * Gözlemci düzeltmeleri (Python düzeltme ~15917–15947).
 * oneMil = 1 milyemlik mesafe değişimi (nişangah_dz).
 */
export function applyObserverCorrections({
  mesafe,
  yan,
  yukselis,
  tapa,
  oneMil,
  sola = 0,
  saga = 0,
  uzat = 0,
  kisalt = 0,
  kaldir = 0,
  indir = 0,
}) {
  const m = Number(mesafe) || 1;
  const bm = Number(oneMil) || 1;
  const ts = Number(tapa) || 1;

  const yan_topla = ((Math.atan(Number(sola) / m) * 180) / Math.PI) * (6400 / 360);
  const yan_cikar = ((Math.atan(Number(saga) / m) * 180) / Math.PI) * (6400 / 360);
  const yukselis_topla = Number(uzat) / bm;
  const yukselis_cikar = Number(kisalt) / bm;
  const paralan_topla = (2 / ts) * (Number(indir) / 10);
  const paralan_cikar = (2 / ts) * (Number(kaldir) / 10);

  return {
    yan: pyRound(Number(yan) + yan_topla - yan_cikar),
    yukselis: round1(Number(yukselis) + yukselis_topla - yukselis_cikar),
    tapa: round1(Number(tapa) + paralan_topla - paralan_cikar),
    deltas: { yan_topla, yan_cikar, yukselis_topla, yukselis_cikar, paralan_topla, paralan_cikar },
  };
}

export { GUN_IDS, GROUP1, GROUP2, recommendCharge, planRange, azimuthMil, baseYan };
