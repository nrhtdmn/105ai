/**
 * Ateş İdare PWA — UI bağlama (top seçimi, müşterek yan bandı, atıldı, görev bitir)
 */
import {
  computeFireSolution,
  recommendAndGeometry,
  applyObserverCorrections,
  GUN_IDS,
} from "./calc.js";
import { fetchElevationForTarget, downloadCsv } from "./elevation.js";

const DEFAULT_MEVZI = {
  "1A": { sag: 19034, yukari: 65603, rakim: 364, ahia: 2400, ihf: 10.0, musyan: 2600 },
  "1B": { sag: 19034, yukari: 65603, rakim: 364, ahia: 3000, ihf: 10.0, musyan: 2600 },
  "1C": { sag: 19034, yukari: 65603, rakim: 364, ahia: 3450, ihf: 10.0, musyan: 2600 },
  "1D": { sag: 19034, yukari: 65603, rakim: 364, ahia: 3700, ihf: 10.0, musyan: 2600 },
  "2A": { sag: 19054, yukari: 65616, rakim: 364, ahia: 1000, ihf: 9.0, musyan: 2600 },
  "2B": { sag: 19054, yukari: 65616, rakim: 364, ahia: 1800, ihf: 9.0, musyan: 2600 },
  "2C": { sag: 19054, yukari: 65616, rakim: 364, ahia: 1800, ihf: 9.0, musyan: 2600 },
  "2D": { sag: 19054, yukari: 65616, rakim: 364, ahia: 200, ihf: 9.0, musyan: 2600 },
};

const ESAS_LABELS = [
  ["plan_yan", "Plan yan"],
  ["yukselis", "Yükseliş"],
  ["mesafe", "Mesafe"],
  ["barut", "Barut hakkı"],
  ["istikamet", "İstikamet"],
  ["ucus", "Uçuş süresi"],
  ["nisangah", "Nişangah"],
  ["dogal_yan", "Doğal yan dzl"],
  ["yuz_m", "100 M"],
  ["yirmi_m", "1 ml değişim"],
  ["dusus", "Düşüş açısı"],
  ["tepe", "Tepe yüksekliği"],
  ["dtac", "DTAÇ"],
  ["tac", "TAÇ"],
  ["gac_yan", "GAC yan"],
  ["toplam_yan", "Toplam yan"],
  ["toplam_mesafe", "Toplam mesafe dzl"],
  ["toplam_tapa", "Toplam tapa dzl"],
  ["metro_yan", "Metro yan"],
  ["metro_mesafe", "Metro mesafe"],
  ["gac_mesafe", "GAC mesafe"],
  ["tapa", "Tapa sn"],
];

const WEAPONS = {
  obus_bh567: {
    id: "obus_bh567",
    label: "Obüs — BH 5 / 6 / 7",
    charges: [5, 6, 7],
    hint: "Mevcut balistik kütüphaneler (charge_5 / 6 / 7).",
  },
};

const DEFAULT_AYARLAR = {
  silah: "obus_bh567",
  birlik: "",
  mermi: "TD",
  barutIsisi: 75,
  kare: 2,
  bolgeNu: "37",
  bolgeSayisi: "SFA",
  musterek: 2600,
  bandHalf: 400,
  solHudut: "",
  sagHudut: "",
  aktifToplar: [...GUN_IDS],
};

const AYAR_LS = "ates_ayarlar";

const state = {
  shared: null,
  libs: {},
  metroText: "",
  metroName: "",
  metroSource: "default",
  last: null,
  /** Canlı düzeltme değerleri (seçili top) */
  live: null,
  shots: [],
  selectedGun: "1B",
  /** Hedef listesinde düzenlenen satır indeksi */
  hlEditIndex: null,
  rakimBusy: false,
  ayarlar: { ...DEFAULT_AYARLAR, aktifToplar: [...GUN_IDS] },
};

const $ = (id) => document.getElementById(id);

function showError(msg) {
  const el = $("calc-error");
  if (!msg) {
    el.hidden = true;
    el.textContent = "";
    return;
  }
  el.hidden = false;
  el.textContent = msg;
}

function fmt(v) {
  if (v == null || Number.isNaN(v)) return "—";
  if (typeof v === "number") return String(Math.round(v * 1000) / 1000);
  return String(v);
}

function bandHalf() {
  const n = Number(state.ayarlar?.bandHalf);
  return Number.isFinite(n) && n > 0 ? n : 400;
}

function musterekYan() {
  const n = Number($("musterek-yan").value);
  return Number.isFinite(n) ? n : Number(state.ayarlar.musterek) || 2600;
}

function yanBand() {
  const m = musterekYan();
  const half = bandHalf();
  return { lo: m - half, hi: m + half, mid: m };
}

function updateBandLabel() {
  const { lo, hi } = yanBand();
  $("yan-bandi").value = `${lo} – ${hi}`;
  $("band-legend").textContent = `(müşterek yan bandı ${lo}–${hi})`;
  if ($("dzl-band-label")) $("dzl-band-label").textContent = `${lo} – ${hi}`;
  refreshGunHighlight();
  updateMissionLabels();
}

function inYanBand(yan) {
  const y = Number(yan);
  if (!Number.isFinite(y)) return false;
  const { lo, hi } = yanBand();
  return y > lo && y < hi;
}

function selectedId() {
  return $("secili-top").value || state.selectedGun || "1B";
}

function setSelectedGun(id) {
  if (!GUN_IDS.includes(id)) return;
  if (!isGunActive(id)) {
    const fallback = activeGunIds().find((g) => g === "1B") || activeGunIds()[0];
    if (fallback && fallback !== id) {
      setSelectedGun(fallback);
      return;
    }
  }
  state.selectedGun = id;
  if ($("secili-top")) $("secili-top").value = id;
  document.querySelectorAll(".gun-card").forEach((c) => {
    c.classList.toggle("selected", c.getAttribute("data-gun") === id);
    c.setAttribute("aria-pressed", c.getAttribute("data-gun") === id ? "true" : "false");
  });
  // Seçilen topun MUSYAN değerini müşterek yana taşı
  const box = document.querySelector(`[data-mevzi="${id}"]`);
  const mus = box?.querySelector('[data-f="musyan"]');
  if (mus && mus.value) {
    $("musterek-yan").value = mus.value;
    updateBandLabel();
  }
  syncLiveFromResult();
  paintDuzeltmeBase();
  updateMissionLabels();
  updateActionButtons();
}

function isGunActive(id) {
  const list = state.ayarlar?.aktifToplar;
  if (!list || !list.length) return true;
  return list.includes(id);
}

function activeGunIds() {
  const list = (state.ayarlar?.aktifToplar || GUN_IDS).filter((id) => GUN_IDS.includes(id));
  // Referanslar her zaman açık
  const set = new Set(list);
  set.add("1B");
  set.add("2B");
  return GUN_IDS.filter((id) => set.has(id));
}

function buildGunPlaceholders() {
  const host = $("gun-results");
  const order = ["1B", "2B", "1A", "1C", "1D", "2A", "2C", "2D"].filter(isGunActive);
  host.innerHTML = order
    .map((id) => {
      const ref = id === "1B" || id === "2B";
      return `<article class="gun-card ${ref ? "ref" : ""}" data-gun="${id}" tabindex="0" role="button" aria-pressed="false">
        <header>
          <strong>${id}</strong>
          <span class="card-badges">
            <span class="band-badge" hidden>BANDA</span>
            <span class="hudut-badge" hidden>HUDUT DIŞI</span>
          </span>
        </header>
        <dl>
          <div><dt>Yan</dt><dd data-k="yan" class="yan-val">—</dd></div>
          <div><dt>Yükseliş</dt><dd data-k="yukselis">—</dd></div>
          <div><dt>BH</dt><dd data-k="charge">—</dd></div>
          <div><dt>Tapa</dt><dd data-k="tapa">—</dd></div>
          <div><dt>İstikamet</dt><dd data-k="azimut">—</dd></div>
          <div><dt>Mesafe</dt><dd data-k="plan">—</dd></div>
        </dl>
      </article>`;
    })
    .join("");

  host.querySelectorAll(".gun-card").forEach((card) => {
    const pick = () => setSelectedGun(card.getAttribute("data-gun"));
    card.addEventListener("click", pick);
    card.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter" || ev.key === " ") {
        ev.preventDefault();
        pick();
      }
    });
  });

  rebuildSeciliTopOptions();
  const pickId = isGunActive(state.selectedGun) ? state.selectedGun : activeGunIds()[0] || "1B";
  setSelectedGun(pickId);
}

function updateMissionLabels() {
  const id = selectedId();
  const topcu = ($("topcu").value || "—").trim() || "—";
  if ($("dzl-top-label")) $("dzl-top-label").textContent = id;
  if ($("dzl-topcu-label")) $("dzl-topcu-label").textContent = topcu;
  const { lo, hi } = yanBand();
  if ($("dzl-band-label")) $("dzl-band-label").textContent = `${lo} – ${hi}`;
}

function updateActionButtons() {
  const ready = !!state.last;
  $("btn-atildi").disabled = !ready;
  $("btn-gorev-bitir").disabled = !ready && state.shots.length === 0;
}

function rebuildSeciliTopOptions() {
  const sel = $("secili-top");
  if (!sel) return;
  const cur = sel.value;
  sel.innerHTML = activeGunIds()
    .map((id) => `<option value="${id}">${id}</option>`)
    .join("");
  if (activeGunIds().includes(cur)) sel.value = cur;
}

function paintGuns(guns) {
  document.querySelectorAll(".gun-card").forEach((card) => {
    const id = card.getAttribute("data-gun");
    const g = guns[id];
    if (!g) return;
    card.querySelectorAll("[data-k]").forEach((dd) => {
      const k = dd.getAttribute("data-k");
      dd.textContent = fmt(g[k]);
    });
  });
  refreshGunHighlight();
}

function refreshGunHighlight() {
  document.querySelectorAll(".gun-card").forEach((card) => {
    const yanEl = card.querySelector('[data-k="yan"]');
    const azEl = card.querySelector('[data-k="azimut"]');
    const yan = Number(yanEl?.textContent);
    const az = Number(azEl?.textContent);
    const ok = inYanBand(yan);
    card.classList.toggle("in-band", ok);
    const badge = card.querySelector(".band-badge");
    if (badge) badge.hidden = !ok;
    if (yanEl) yanEl.classList.toggle("in-band", ok);

    const out = Number.isFinite(az) && !inHudut(az);
    card.classList.toggle("out-hudut", out);
    const hb = card.querySelector(".hudut-badge");
    if (hb) hb.hidden = !out;
  });
  updateHudutHint();
}

function milNorm(v) {
  const n = Number(v);
  if (!Number.isFinite(n)) return NaN;
  return ((Math.round(n) % 6400) + 6400) % 6400;
}

function readHudut() {
  const solRaw = $("sol-hudut")?.value;
  const sagRaw = $("sag-hudut")?.value;
  const sol =
    solRaw === "" || solRaw == null ? null : milNorm(solRaw);
  const sag =
    sagRaw === "" || sagRaw == null ? null : milNorm(sagRaw);
  return { sol, sag };
}

/** Sol→sağ artan milyem yönünde (0’ı dolanabilir) sektör kontrolü */
function inHudut(azimut) {
  const { sol, sag } = readHudut();
  if (sol == null || sag == null || Number.isNaN(sol) || Number.isNaN(sag)) return true;
  const az = milNorm(azimut);
  if (Number.isNaN(az)) return true;
  if (sol === sag) return az === sol;
  if (sol < sag) return az >= sol && az <= sag;
  return az >= sol || az <= sag;
}

function saveHudut() {
  const { sol, sag } = readHudut();
  localStorage.setItem(
    "ates_hudut",
    JSON.stringify({
      sol: sol == null || Number.isNaN(sol) ? "" : sol,
      sag: sag == null || Number.isNaN(sag) ? "" : sag,
    })
  );
  // Ayarlar formunu senkronla
  if ($("ay-sol-hudut")) $("ay-sol-hudut").value = $("sol-hudut").value;
  if ($("ay-sag-hudut")) $("ay-sag-hudut").value = $("sag-hudut").value;
  state.ayarlar.solHudut = $("sol-hudut").value;
  state.ayarlar.sagHudut = $("sag-hudut").value;
  try {
    localStorage.setItem(AYAR_LS, JSON.stringify(state.ayarlar));
  } catch {
    /* */
  }
  refreshGunHighlight();
}

function loadHudut() {
  try {
    const h = JSON.parse(localStorage.getItem("ates_hudut") || "null");
    if (!h) return;
    if ($("sol-hudut") && h.sol !== "" && h.sol != null) $("sol-hudut").value = h.sol;
    if ($("sag-hudut") && h.sag !== "" && h.sag != null) $("sag-hudut").value = h.sag;
  } catch {
    /* ignore */
  }
}

function updateHudutHint() {
  const el = $("hudut-hint");
  if (!el) return;
  const { sol, sag } = readHudut();
  if (sol == null || sag == null || Number.isNaN(sol) || Number.isNaN(sag)) {
    el.textContent = "Sol / sağ hudut boşsa sınır kontrolü yapılmaz. Değerler tarayıcıda saklanır.";
    return;
  }
  el.textContent = `Hudut: ${sol} – ${sag} mil (sol→sağ). Hedef istikameti bu aralıkta olmalı.`;
}

const METRO_LS_KEY = "ates_metro";

function saveMetro(text, name, source) {
  state.metroText = text;
  state.metroName = name || "METRAP";
  state.metroSource = source || "custom";
  localStorage.setItem(
    METRO_LS_KEY,
    JSON.stringify({
      text,
      name: state.metroName,
      source: state.metroSource,
      savedAt: new Date().toISOString(),
    })
  );
  updateMetroLabel();
}

function loadSavedMetro() {
  try {
    const raw = localStorage.getItem(METRO_LS_KEY);
    if (!raw) return null;
    const o = JSON.parse(raw);
    if (!o?.text) return null;
    return o;
  } catch {
    return null;
  }
}

function updateMetroLabel() {
  const el = $("metro-label");
  const ayEl = $("ay-metro-label");
  if (!state.metroText) {
    if (el) el.textContent = "METRAP yok";
    if (ayEl) ayEl.textContent = "METRAP yok";
    return;
  }
  const lines = state.metroText.trim().split("\n").length;
  const when = (() => {
    try {
      const o = JSON.parse(localStorage.getItem(METRO_LS_KEY) || "{}");
      if (o.savedAt) return " · " + new Date(o.savedAt).toLocaleString("tr-TR");
    } catch {
      /* */
    }
    return "";
  })();
  let text;
  if (state.metroSource === "default") {
    text = `Varsayılan METRAP (${lines} satır)${when}`;
  } else {
    text = `Kayıtlı METRAP: ${state.metroName || "dosya"} (${lines} satır) — yenisi yüklenene kadar geçerli${when}`;
  }
  if (el) el.textContent = text;
  if (ayEl) ayEl.textContent = text;
}

async function loadDefaultMetroText() {
  const base = new URL("../data/", import.meta.url);
  return fetch(new URL("METRAP.txt", base)).then((r) => r.text());
}

function loadAyarlar() {
  try {
    const raw = localStorage.getItem(AYAR_LS);
    if (!raw) {
      state.ayarlar = { ...DEFAULT_AYARLAR, aktifToplar: [...GUN_IDS] };
      return state.ayarlar;
    }
    const o = JSON.parse(raw);
    state.ayarlar = {
      ...DEFAULT_AYARLAR,
      ...o,
      aktifToplar: Array.isArray(o.aktifToplar) ? o.aktifToplar.filter((id) => GUN_IDS.includes(id)) : [...GUN_IDS],
    };
    if (!state.ayarlar.aktifToplar.includes("1B")) state.ayarlar.aktifToplar.push("1B");
    if (!state.ayarlar.aktifToplar.includes("2B")) state.ayarlar.aktifToplar.push("2B");
  } catch {
    state.ayarlar = { ...DEFAULT_AYARLAR, aktifToplar: [...GUN_IDS] };
  }
  return state.ayarlar;
}

function fillAyarlarForm() {
  const a = state.ayarlar;
  if ($("ay-silah")) $("ay-silah").value = a.silah || "obus_bh567";
  if ($("ay-birlik")) $("ay-birlik").value = a.birlik || "";
  if ($("ay-mermi")) $("ay-mermi").value = a.mermi || "TD";
  if ($("ay-barut-isisi")) $("ay-barut-isisi").value = a.barutIsisi ?? 75;
  if ($("ay-kare")) $("ay-kare").value = a.kare ?? 2;
  if ($("ay-bolge-nu")) $("ay-bolge-nu").value = a.bolgeNu || "37";
  if ($("ay-bolge-sayisi")) $("ay-bolge-sayisi").value = a.bolgeSayisi || "SFA";
  if ($("ay-musterek")) $("ay-musterek").value = a.musterek ?? 2600;
  if ($("ay-band-half")) $("ay-band-half").value = a.bandHalf ?? 400;
  if ($("ay-sol-hudut")) $("ay-sol-hudut").value = a.solHudut ?? "";
  if ($("ay-sag-hudut")) $("ay-sag-hudut").value = a.sagHudut ?? "";
  buildAktifToplarToggles();
  const w = WEAPONS[a.silah] || WEAPONS.obus_bh567;
  if ($("ay-silah-hint")) $("ay-silah-hint").textContent = w.hint || "";
}

function buildAktifToplarToggles() {
  const host = $("ay-aktif-toplar");
  if (!host) return;
  const aktif = new Set(activeGunIds());
  host.innerHTML = GUN_IDS.map((id) => {
    const locked = id === "1B" || id === "2B";
    const checked = aktif.has(id) ? "checked" : "";
    const dis = locked ? "disabled" : "";
    return `<label class="gun-toggle ${locked ? "locked" : ""}">
      <input type="checkbox" data-gun-toggle="${id}" ${checked} ${dis} />
      <span>${id}${locked ? " (ref)" : ""}</span>
    </label>`;
  }).join("");
}

function readAyarlarForm() {
  const aktif = GUN_IDS.filter((id) => {
    if (id === "1B" || id === "2B") return true;
    const el = document.querySelector(`[data-gun-toggle="${id}"]`);
    return el ? el.checked : true;
  });
  const silah = $("ay-silah")?.value || "obus_bh567";
  return {
    silah,
    birlik: ($("ay-birlik")?.value || "").trim(),
    mermi: $("ay-mermi")?.value || "TD",
    barutIsisi: Number($("ay-barut-isisi")?.value) || 75,
    kare: Number($("ay-kare")?.value) || 2,
    bolgeNu: ($("ay-bolge-nu")?.value || "37").trim() || "37",
    bolgeSayisi: ($("ay-bolge-sayisi")?.value || "SFA").trim().toUpperCase() || "SFA",
    musterek: Number($("ay-musterek")?.value) || 2600,
    bandHalf: Number($("ay-band-half")?.value) || 400,
    solHudut: $("ay-sol-hudut")?.value ?? "",
    sagHudut: $("ay-sag-hudut")?.value ?? "",
    aktifToplar: aktif,
  };
}

function applyAyarlar(a, { rebuildGuns = true } = {}) {
  state.ayarlar = a;
  localStorage.setItem(AYAR_LS, JSON.stringify(a));

  if ($("barut-isisi")) $("barut-isisi").value = a.barutIsisi;
  if ($("kare-agirlik")) $("kare-agirlik").value = a.kare;
  if ($("mermi-tip")) $("mermi-tip").value = a.mermi;
  if ($("musterek-yan")) $("musterek-yan").value = a.musterek;
  if ($("hedef-bolge-nu") && !$("hedef-sag").value) $("hedef-bolge-nu").value = a.bolgeNu;
  if ($("hedef-bolge-sayisi") && !$("hedef-sag").value) $("hedef-bolge-sayisi").value = a.bolgeSayisi;
  if ($("hl-bolge-nu")) $("hl-bolge-nu").value = a.bolgeNu;
  if ($("hl-bolge-sayisi")) $("hl-bolge-sayisi").value = a.bolgeSayisi;
  if ($("sol-hudut")) $("sol-hudut").value = a.solHudut ?? "";
  if ($("sag-hudut")) $("sag-hudut").value = a.sagHudut ?? "";
  localStorage.setItem(
    "ates_hudut",
    JSON.stringify({ sol: a.solHudut ?? "", sag: a.sagHudut ?? "" })
  );

  // Barut hakkı seçeneklerini silaha göre
  const w = WEAPONS[a.silah] || WEAPONS.obus_bh567;
  const bh = $("barut-hakki");
  if (bh && w.charges) {
    const cur = bh.value;
    bh.innerHTML = w.charges.map((c) => `<option value="${c}">${c}</option>`).join("");
    bh.value = w.charges.map(String).includes(cur) ? cur : String(w.charges[0]);
  }

  const birlik = a.birlik ? ` · ${a.birlik}` : "";
  if ($("app-tagline")) $("app-tagline").textContent = `${w.label}${birlik}`;

  updateBandLabel();
  updateHudutHint();
  if (rebuildGuns) {
    buildGunPlaceholders();
    if (state.last) paintGuns(state.last.guns);
  }
}

function kaydetAyarlar() {
  const a = readAyarlarForm();
  applyAyarlar(a);
  fillAyarlarForm();
  if ($("ay-status")) $("ay-status").textContent = "Ayarlar kaydedildi · " + new Date().toLocaleTimeString("tr-TR");
}

function sifirlaAyarlar() {
  state.ayarlar = { ...DEFAULT_AYARLAR, aktifToplar: [...GUN_IDS] };
  applyAyarlar(state.ayarlar);
  fillAyarlarForm();
  if ($("ay-status")) $("ay-status").textContent = "Varsayılan ayarlara dönüldü.";
}

async function loadData() {
  const base = new URL("../data/", import.meta.url);
  const [shared, c5, c6, c7] = await Promise.all([
    fetch(new URL("shared.json", base)).then((r) => r.json()),
    fetch(new URL("charge_5.json", base)).then((r) => r.json()),
    fetch(new URL("charge_6.json", base)).then((r) => r.json()),
    fetch(new URL("charge_7.json", base)).then((r) => r.json()),
  ]);
  state.shared = shared;
  state.libs = { 5: c5, 6: c6, 7: c7 };

  const saved = loadSavedMetro();
  if (saved?.text) {
    state.metroText = saved.text;
    state.metroName = saved.name || "METRAP";
    state.metroSource = saved.source || "custom";
    updateMetroLabel();
  } else {
    const metro = await loadDefaultMetroText();
    saveMetro(metro, "varsayılan METRAP.txt", "default");
  }
}

function loadMevzi() {
  try {
    return JSON.parse(localStorage.getItem("ates_mevzi")) || { ...DEFAULT_MEVZI };
  } catch {
    return { ...DEFAULT_MEVZI };
  }
}

function buildMevziForms() {
  const data = loadMevzi();
  const host = $("mevzi-forms");
  host.innerHTML = GUN_IDS.map((id) => {
    const d = data[id] || DEFAULT_MEVZI[id];
    return `<div class="card mevzi-card" data-mevzi="${id}">
      <h3>${id}</h3>
      <div class="form-grid dense">
        <label>Sağ <input data-f="sag" type="number" value="${d.sag}" /></label>
        <label>Yukarı <input data-f="yukari" type="number" value="${d.yukari}" /></label>
        <label>Rakım <input data-f="rakim" type="number" value="${d.rakim}" /></label>
        <label>AHIA <input data-f="ahia" type="number" value="${d.ahia}" /></label>
        <label>İHF <input data-f="ihf" type="number" step="0.1" value="${d.ihf}" /></label>
        <label>MUSYAN <input data-f="musyan" type="number" value="${d.musyan}" /></label>
      </div>
    </div>`;
  }).join("");
}

function readGuns() {
  const guns = {};
  document.querySelectorAll("[data-mevzi]").forEach((box) => {
    const id = box.getAttribute("data-mevzi");
    const val = (f) => Number(box.querySelector(`[data-f="${f}"]`).value);
    guns[id] = {
      sag: val("sag"),
      yukari: val("yukari"),
      rakim: val("rakim"),
      ahia: val("ahia"),
      ihf: val("ihf"),
      musyan: val("musyan"),
    };
  });
  return guns;
}

function saveMevzi() {
  localStorage.setItem("ates_mevzi", JSON.stringify(readGuns()));
  showError("");
  $("mission-hint").textContent = "Mevziler kaydedildi.";
}

function resetMevzi() {
  localStorage.removeItem("ates_mevzi");
  buildMevziForms();
}

function readTarget() {
  const sag = Number($("hedef-sag").value);
  const yukari = Number($("hedef-yukari").value);
  const rakim = Number($("hedef-rakim").value);
  if (!Number.isFinite(sag) || !Number.isFinite(yukari)) {
    throw new Error("Hedef sağ ve yukarı zorunlu.");
  }
  if (!Number.isFinite(rakim)) {
    throw new Error("Hedef rakımını girin (PWA’da yükseklik API’si yok).");
  }
  return {
    sag,
    yukari,
    rakim,
    bolgeNu: $("hedef-bolge-nu").value || "37",
    bolgeSayisi: $("hedef-bolge-sayisi").value || "SFA",
  };
}

function paintEsaslar(esas) {
  $("esaslar-grid").innerHTML = ESAS_LABELS.map(
    ([k, label]) =>
      `<div class="kpi-tile"><span class="lab">${label}</span><span class="val">${fmt(
        esas[k]
      )}</span></div>`
  ).join("");
}

function syncLiveFromResult() {
  if (!state.last) {
    state.live = null;
    return;
  }
  const id = selectedId();
  const g = state.last.guns[id];
  if (!g) return;
  state.live = {
    gun: id,
    yan: g.yan,
    yukselis: g.yukselis,
    tapa: state.last.esaslar.tapa,
    plan: g.plan,
    charge: state.last.charge,
    birmil: state.last.esaslar.yirmi_m,
    yuzm: state.last.esaslar.yuz_m,
  };
}

function paintDuzeltmeBase() {
  const L = state.live;
  if (!L) {
    ["dzl-yan", "dzl-yukselis", "dzl-barut", "dzl-tapa", "dzl-mesafe", "dzl-birmil", "dzl-yuzm"].forEach(
      (id) => {
        $(id).textContent = "—";
      }
    );
    return;
  }
  $("dzl-yan").textContent = fmt(L.yan);
  $("dzl-yukselis").textContent = fmt(L.yukselis);
  $("dzl-barut").textContent = fmt(L.charge);
  $("dzl-tapa").textContent = fmt(L.tapa);
  $("dzl-mesafe").textContent = fmt(L.plan);
  $("dzl-birmil").textContent = fmt(L.birmil);
  $("dzl-yuzm").textContent = fmt(L.yuzm);
  updateMissionLabels();
}

function oner() {
  try {
    showError("");
    const guns = readGuns();
    const target = readTarget();
    const geo = recommendAndGeometry({ guns, target, shared: state.shared });
    if (geo.charge1 == null) throw new Error("Atış yapılamaz — menzil dışı.");
    $("barut-oner").value = String(geo.charge1);
    $("barut-hakki").value = String(Math.min(7, Math.max(5, geo.charge1)));
    const preview = {};
    for (const id of GUN_IDS) {
      preview[id] = {
        yan: Math.round(geo.yans[id]),
        yukselis: "—",
        charge: id.startsWith("1") ? geo.charge1 : geo.charge2,
        tapa: $("mermi-tip").value,
        azimut: id.startsWith("1") ? geo.azimut1 : geo.azimut2,
        plan: geo.plans[id],
      };
    }
    paintGuns(preview);
    // Band içindeki ilk topu öner
    const inBand = activeGunIds().find((id) => inYanBand(preview[id].yan));
    if (inBand) setSelectedGun(inBand);
  } catch (e) {
    showError(e.message || String(e));
  }
}

function hesapla() {
  try {
    showError("");
    if (!state.shared) throw new Error("Kütüphaneler yüklenmedi.");
    const guns = readGuns();
    const target = readTarget();
    const charge = Number($("barut-hakki").value);
    if (![5, 6, 7].includes(charge)) {
      throw new Error("Tam çözüm için barut hakkı 5, 6 veya 7 olmalı.");
    }
    const result = computeFireSolution({
      charge,
      lib: state.libs[charge],
      shared: state.shared,
      guns,
      target,
      barutIsisi: Number($("barut-isisi").value) || 75,
      squareWeight: Number($("kare-agirlik").value) || 2,
      metroText: state.metroText,
      paralanmaYuksekligi: Number($("paralanma").value) || 0,
    });
    state.last = result;
    paintGuns(result.guns);
    paintEsaslar(result.esaslar);

    // Band içindeki topu otomatik seç (yoksa mevcut seçim)
    const inBand = activeGunIds().find((id) => inYanBand(result.guns[id]?.yan));
    if (inBand) setSelectedGun(inBand);
    else {
      syncLiveFromResult();
      paintDuzeltmeBase();
    }
    updateActionButtons();
    const az1 = result.guns["1B"]?.azimut;
    const az2 = result.guns["2B"]?.azimut;
    const hudutDisi =
      (Number.isFinite(Number(az1)) && !inHudut(az1)) ||
      (Number.isFinite(Number(az2)) && !inHudut(az2));
    let hint = inBand
      ? `${inBand} müşterek yan bandında — seçildi. ATILDI ile kaydedin.`
      : "Hiçbir top bandında değil; listeden top seçin.";
    if (hudutDisi) hint += " ⚠ Hedef istikameti sol/sağ hudut dışında.";
    $("mission-hint").textContent = hint;
  } catch (e) {
    console.error(e);
    showError(e.message || String(e));
  }
}

function dzlUygula() {
  if (!state.live) {
    showError("Önce HESAPLA çalıştırın ve top seçin.");
    return;
  }
  showError("");
  const L = state.live;
  const out = applyObserverCorrections({
    mesafe: L.plan,
    yan: L.yan,
    yukselis: L.yukselis,
    tapa: L.tapa,
    oneMil: L.birmil,
    sola: Number($("dzl-sola").value) || 0,
    saga: Number($("dzl-saga").value) || 0,
    uzat: Number($("dzl-uzat").value) || 0,
    kisalt: Number($("dzl-kisalt").value) || 0,
    kaldir: Number($("dzl-kaldir").value) || 0,
    indir: Number($("dzl-indir").value) || 0,
  });
  state.live = { ...L, yan: out.yan, yukselis: out.yukselis, tapa: out.tapa };
  // Kartı da güncelle
  if (state.last?.guns[L.gun]) {
    state.last.guns[L.gun].yan = out.yan;
    state.last.guns[L.gun].yukselis = out.yukselis;
    paintGuns(state.last.guns);
  }
  paintDuzeltmeBase();
}

function dzlTemizle() {
  ["sola", "saga", "uzat", "kisalt", "kaldir", "indir"].forEach((k) => {
    $(`dzl-${k}`).value = "0";
  });
  syncLiveFromResult();
  paintDuzeltmeBase();
  if (state.last) paintGuns(state.last.guns);
}

function renderShots() {
  const tb = $("shot-tbody");
  if (!state.shots.length) {
    tb.innerHTML = `<tr class="empty"><td colspan="8">Henüz atış yok</td></tr>`;
    return;
  }
  tb.innerHTML = state.shots
    .map(
      (s, i) => `<tr>
      <td>${i + 1}</td>
      <td>${s.saat}</td>
      <td>${s.top}</td>
      <td>${s.topcu}</td>
      <td>${fmt(s.yan)}</td>
      <td>${fmt(s.yukselis)}</td>
      <td>${fmt(s.charge)}</td>
      <td>${fmt(s.tapa)}</td>
    </tr>`
    )
    .join("");
}

function atildi() {
  if (!state.live && !state.last) {
    showError("Önce hesaplayın.");
    return;
  }
  showError("");
  const id = selectedId();
  const L = state.live || {
    yan: state.last.guns[id].yan,
    yukselis: state.last.guns[id].yukselis,
    tapa: state.last.esaslar.tapa,
    charge: state.last.charge,
  };
  const now = new Date();
  const saat = now.toLocaleTimeString("tr-TR", { hour: "2-digit", minute: "2-digit", second: "2-digit" });
  state.shots.push({
    saat,
    top: id,
    topcu: ($("topcu").value || "").trim() || "—",
    yan: L.yan,
    yukselis: L.yukselis,
    charge: L.charge ?? state.last?.charge,
    tapa: L.tapa,
    musterekYan: musterekYan(),
  });
  renderShots();
  updateActionButtons();
  $("mission-hint").textContent = `${id} atıldı kaydedildi (${state.shots.length}. atış).`;
}

function gorevBitir() {
  try {
    let target = null;
    try {
      target = readTarget();
    } catch {
      target = {
        sag: $("hedef-sag").value,
        yukari: $("hedef-yukari").value,
        rakim: $("hedef-rakim").value,
        bolgeNu: $("hedef-bolge-nu").value,
        bolgeSayisi: $("hedef-bolge-sayisi").value,
      };
    }
    const kayit = {
      bitis: new Date().toISOString(),
      bitisYerel: new Date().toLocaleString("tr-TR"),
      topcu: ($("topcu").value || "").trim(),
      seciliTop: selectedId(),
      musterekYan: musterekYan(),
      yanBand: yanBand(),
      hedef: target,
      barutHakki: Number($("barut-hakki").value) || null,
      shots: state.shots.slice(),
      sonuc: state.last
        ? {
            charge: state.last.charge,
            guns: state.last.guns,
            esaslar: state.last.esaslar,
          }
        : null,
    };
    const list = JSON.parse(localStorage.getItem("ates_gorevler") || "[]");
    list.unshift(kayit);
    localStorage.setItem("ates_gorevler", JSON.stringify(list.slice(0, 50)));
    gorevListesi();

    // Başa dön — görev alanlarını sıfırla
    state.last = null;
    state.live = null;
    state.shots = [];
    renderShots();
    $("hedef-sag").value = "";
    $("hedef-yukari").value = "";
    $("hedef-rakim").value = "";
    $("barut-oner").value = "";
    $("topcu").value = "";
    ["sola", "saga", "uzat", "kisalt", "kaldir", "indir"].forEach((k) => {
      $(`dzl-${k}`).value = "0";
    });
    buildGunPlaceholders();
    updateBandLabel();
    paintDuzeltmeBase();
    $("esaslar-grid").innerHTML = "";
    updateActionButtons();
    showError("");
    $("mission-hint").textContent = "Görev kaydedildi. Yeni hedef girebilirsiniz.";
    // Atış görevine dön
    document.querySelector('.tab[data-tab="atis"]')?.click();
  } catch (e) {
    showError(e.message || String(e));
  }
}

function gorevListesi() {
  const list = JSON.parse(localStorage.getItem("ates_gorevler") || "[]");
  const ul = $("gorev-listesi");
  if (!ul) return;
  ul.innerHTML = list.length
    ? list
        .map((g, i) => {
          const h = g.hedef || {};
          const n = (g.shots || []).length;
          const name = h.isim ? escapeHtml(h.isim) + " · " : "";
          return `<li>
            <div class="target-meta">
              <span class="name">${g.bitisYerel || g.bitis}</span>
              <span class="coords">${name}${h.sag || "?"}/${h.yukari || "?"} · ${n} atış · ${
                g.seciliTop || "?"
              } · ${escapeHtml(g.topcu || "—")}</span>
            </div>
            <div class="target-actions">
              <button type="button" class="btn-icon" data-gorev="${i}" title="Sil">✕</button>
            </div>
          </li>`;
        })
        .join("")
    : "<li class='hint'>Bitmiş görev yok</li>";
  ul.querySelectorAll("[data-gorev]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const i = Number(btn.getAttribute("data-gorev"));
      const arr = JSON.parse(localStorage.getItem("ates_gorevler") || "[]");
      arr.splice(i, 1);
      localStorage.setItem("ates_gorevler", JSON.stringify(arr));
      gorevListesi();
    });
  });
}

function hlStatus(msg, isError = false) {
  const st = $("hl-status");
  const er = $("hl-error");
  if (isError) {
    if (er) {
      er.hidden = false;
      er.textContent = msg;
    }
    return;
  }
  if (er) {
    er.hidden = true;
    er.textContent = "";
  }
  if (st) st.textContent = msg || "🔥 = atış görevine yükle · ✎ = düzenle · ✕ = sil";
}

function loadHedefler() {
  try {
    return JSON.parse(localStorage.getItem("ates_hedefler") || "[]");
  } catch {
    return [];
  }
}

function saveHedefler(list) {
  localStorage.setItem("ates_hedefler", JSON.stringify(list));
}

function readHlForm() {
  const isim = ($("hl-isim").value || "").trim();
  const bolgeNu = ($("hl-bolge-nu").value || "").trim() || "37";
  const bolgeSayisi = ($("hl-bolge-sayisi").value || "").trim().toUpperCase() || "SFA";
  const sag = Number($("hl-sag").value);
  const yukari = Number($("hl-yukari").value);
  const rakim = Number($("hl-rakim").value);
  const not = ($("hl-not").value || "").trim();
  if (!isim) throw new Error("Hedef ismi girin.");
  if (!Number.isFinite(sag) || !Number.isFinite(yukari)) {
    throw new Error("Sağ / yukarı koordinatı gerekli.");
  }
  if (!Number.isFinite(rakim)) throw new Error("Rakım gerekli (Rakım al veya elle girin).");
  return {
    id: crypto.randomUUID ? crypto.randomUUID() : String(Date.now()),
    isim,
    bolgeNu,
    bolgeSayisi,
    sag,
    yukari,
    rakim,
    not,
  };
}

function clearHlForm() {
  state.hlEditIndex = null;
  $("hl-isim").value = "";
  $("hl-sag").value = "";
  $("hl-yukari").value = "";
  $("hl-rakim").value = "";
  $("hl-not").value = "";
  if (!$("hl-bolge-nu").value) $("hl-bolge-nu").value = "37";
  if (!$("hl-bolge-sayisi").value) $("hl-bolge-sayisi").value = "SFA";
  $("btn-hl-kaydet").textContent = "Listeye ekle";
  $("btn-hl-iptal").hidden = true;
}

function fillHlForm(h, index) {
  state.hlEditIndex = index;
  $("hl-isim").value = h.isim || "";
  $("hl-bolge-nu").value = h.bolgeNu || "37";
  $("hl-bolge-sayisi").value = h.bolgeSayisi || "SFA";
  $("hl-sag").value = h.sag ?? "";
  $("hl-yukari").value = h.yukari ?? "";
  $("hl-rakim").value = h.rakim ?? "";
  $("hl-not").value = h.not || "";
  $("btn-hl-kaydet").textContent = "Güncelle";
  $("btn-hl-iptal").hidden = false;
}

function applyTargetToAtis(h) {
  $("hedef-bolge-nu").value = h.bolgeNu || "37";
  $("hedef-bolge-sayisi").value = h.bolgeSayisi || "SFA";
  $("hedef-sag").value = h.sag ?? "";
  $("hedef-yukari").value = h.yukari ?? "";
  $("hedef-rakim").value = h.rakim ?? "";
  document.querySelector('.tab[data-tab="atis"]')?.click();
  $("mission-hint").textContent = h.isim
    ? `Hedef yüklendi: ${h.isim}`
    : "Hedef atış görevine yüklendi.";
}

function hedefListesi() {
  const list = loadHedefler();
  const ul = $("hedef-listesi");
  if (!ul) return;
  ul.innerHTML = list.length
    ? list
        .map((h, i) => {
          const name = h.isim || `Hedef ${i + 1}`;
          const note = h.not ? ` · ${escapeHtml(h.not)}` : "";
          return `<li>
            <div class="target-meta">
              <span class="name">${escapeHtml(name)}</span>
              <span class="coords">${escapeHtml(String(h.bolgeNu || ""))}${escapeHtml(
                String(h.bolgeSayisi || "")
              )} · ${h.sag}/${h.yukari} · ${h.rakim} m${note}</span>
            </div>
            <div class="target-actions">
              <button type="button" class="btn-icon fire" data-fire="${i}" title="Atış görevine yükle">🔥</button>
              <button type="button" class="btn-icon" data-edit="${i}" title="Düzenle">✎</button>
              <button type="button" class="btn-icon" data-del="${i}" title="Sil">✕</button>
            </div>
          </li>`;
        })
        .join("")
    : "<li class='hint'>Liste boş — isim ve koordinat girip ekleyin.</li>";

  ul.querySelectorAll("[data-fire]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const h = list[Number(btn.getAttribute("data-fire"))];
      if (h) applyTargetToAtis(h);
    });
  });
  ul.querySelectorAll("[data-edit]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const i = Number(btn.getAttribute("data-edit"));
      fillHlForm(list[i], i);
      hlStatus(`Düzenleniyor: ${list[i].isim || i + 1}`);
    });
  });
  ul.querySelectorAll("[data-del]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const i = Number(btn.getAttribute("data-del"));
      const name = list[i]?.isim || String(i + 1);
      if (!confirm(`"${name}" listeden silinsin mi?`)) return;
      list.splice(i, 1);
      saveHedefler(list);
      if (state.hlEditIndex === i) clearHlForm();
      else if (state.hlEditIndex != null && state.hlEditIndex > i) state.hlEditIndex -= 1;
      hedefListesi();
      hlStatus(`Silindi: ${name}`);
    });
  });
}

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function hedefKaydetHl() {
  try {
    const t = readHlForm();
    const list = loadHedefler();
    if (state.hlEditIndex != null && state.hlEditIndex >= 0 && state.hlEditIndex < list.length) {
      t.id = list[state.hlEditIndex].id || t.id;
      list[state.hlEditIndex] = t;
      hlStatus(`Güncellendi: ${t.isim}`);
    } else {
      list.push(t);
      hlStatus(`Eklendi: ${t.isim}`);
    }
    saveHedefler(list);
    clearHlForm();
    hedefListesi();
  } catch (e) {
    hlStatus(e.message || String(e), true);
  }
}

function hlMevcutAl() {
  $("hl-bolge-nu").value = $("hedef-bolge-nu").value || "37";
  $("hl-bolge-sayisi").value = $("hedef-bolge-sayisi").value || "SFA";
  $("hl-sag").value = $("hedef-sag").value;
  $("hl-yukari").value = $("hedef-yukari").value;
  $("hl-rakim").value = $("hedef-rakim").value;
  if (!$("hl-isim").value.trim()) {
    $("hl-isim").value = `Hedef ${$("hedef-sag").value || ""}`;
  }
  hlStatus("Atış görevi koordinatları forma alındı — kaydedin.");
}

async function fetchRakimInto(fields, statusEl) {
  const { bolgeNuId, bolgeSayisiId, sagId, yukariId, rakimId, busyBtn } = fields;
  const bolgeNu = $(bolgeNuId).value;
  const bolgeSayisi = $(bolgeSayisiId).value;
  const sag = Number($(sagId).value);
  const yukari = Number($(yukariId).value);
  if (!Number.isFinite(sag) || !Number.isFinite(yukari)) {
    throw new Error("Önce sağ / yukarı girin.");
  }
  if (busyBtn) busyBtn.classList.add("rakim-busy");
  state.rakimBusy = true;
  try {
    const r = await fetchElevationForTarget({ bolgeNu, bolgeSayisi, sag, yukari });
    $(rakimId).value = String(r.elev);
    if (statusEl === "calc") {
      showError("");
      $("mission-hint").textContent = `Rakım ${r.elev} m (${r.source}) · ${r.lat.toFixed(5)}, ${r.lon.toFixed(5)}`;
    } else if (statusEl === "hl") {
      hlStatus(`Rakım ${r.elev} m (${r.source})`);
    }
    return r.elev;
  } finally {
    state.rakimBusy = false;
    if (busyBtn) busyBtn.classList.remove("rakim-busy");
  }
}

async function autoRakimAtis() {
  if (state.rakimBusy) return;
  const sag = Number($("hedef-sag").value);
  const yukari = Number($("hedef-yukari").value);
  if (!Number.isFinite(sag) || !Number.isFinite(yukari)) return;
  try {
    await fetchRakimInto(
      {
        bolgeNuId: "hedef-bolge-nu",
        bolgeSayisiId: "hedef-bolge-sayisi",
        sagId: "hedef-sag",
        yukariId: "hedef-yukari",
        rakimId: "hedef-rakim",
        busyBtn: $("btn-rakim-al"),
      },
      "calc"
    );
  } catch (e) {
    showError(e.message || String(e));
  }
}

function recommendChargeForTarget(target) {
  const guns = readGuns();
  const geo = recommendAndGeometry({ guns, target, shared: state.shared });
  if (geo.charge1 == null) return null;
  return Math.min(7, Math.max(5, Number(geo.charge1)));
}

async function tumHedeflereHesaplaExcel() {
  try {
    hlStatus("Hesaplanıyor…");
    if (!state.shared || !state.metroText) throw new Error("Kütüphaneler / METRAP yüklenmedi.");
    const list = loadHedefler();
    if (!list.length) throw new Error("Hedef listesi boş.");
    const guns = readGuns();
    const barutIsisi = Number($("barut-isisi").value) || 75;
    const squareWeight = Number($("kare-agirlik").value) || 2;
    const paralanma = Number($("paralanma").value) || 0;

    const header = [
      "İsim",
      "Bölge",
      "Sağ",
      "Yukarı",
      "Rakım",
      "Not",
      "Barut",
      "Plan yan",
      "Yükseliş",
      "Mesafe",
      "İstikamet",
      "Uçuş",
      "Nişangah",
      "Tapa",
      "1B yan",
      "1B yükseliş",
      "2B yan",
      "2B yükseliş",
      "Durum",
    ];
    const rows = [header];

    for (const h of list) {
      const target = {
        sag: Number(h.sag),
        yukari: Number(h.yukari),
        rakim: Number(h.rakim),
        bolgeNu: h.bolgeNu,
        bolgeSayisi: h.bolgeSayisi,
      };
      const base = [
        h.isim || "",
        `${h.bolgeNu || ""}${h.bolgeSayisi || ""}`,
        h.sag,
        h.yukari,
        h.rakim,
        h.not || "",
      ];
      try {
        if (!Number.isFinite(target.sag) || !Number.isFinite(target.yukari) || !Number.isFinite(target.rakim)) {
          throw new Error("Eksik koordinat/rakım");
        }
        const charge = recommendChargeForTarget(target);
        if (charge == null) throw new Error("Menzil dışı");
        const result = computeFireSolution({
          charge,
          lib: state.libs[charge],
          shared: state.shared,
          guns,
          target,
          barutIsisi,
          squareWeight,
          metroText: state.metroText,
          paralanmaYuksekligi: paralanma,
        });
        const e = result.esaslar;
        rows.push([
          ...base,
          charge,
          e.plan_yan,
          e.yukselis,
          e.mesafe,
          e.istikamet,
          e.ucus,
          e.nisangah,
          e.tapa,
          result.guns["1B"]?.yan,
          result.guns["1B"]?.yukselis,
          result.guns["2B"]?.yan,
          result.guns["2B"]?.yukselis,
          "OK",
        ]);
      } catch (err) {
        rows.push([...base, "", "", "", "", "", "", "", "", "", "", "", "", err.message || String(err)]);
      }
    }

    const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
    downloadCsv(`ates_hedefler_${stamp}.csv`, rows);
    hlStatus(`${list.length} hedef hesaplandı · Excel (CSV) indirildi.`);
  } catch (e) {
    hlStatus(e.message || String(e), true);
  }
}

function bindTabs() {
  document.querySelectorAll(".tab").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab").forEach((b) => b.classList.remove("active"));
      document.querySelectorAll(".panel").forEach((p) => {
        p.classList.remove("active");
        p.hidden = true;
      });
      btn.classList.add("active");
      const panel = $("tab-" + btn.dataset.tab);
      panel.classList.add("active");
      panel.hidden = false;
    });
  });
}

async function main() {
  loadAyarlar();
  fillAyarlarForm();
  applyAyarlar(state.ayarlar, { rebuildGuns: false });

  buildGunPlaceholders();
  buildMevziForms();
  bindTabs();
  loadHudut();
  // Hudut: ayarlar öncelikli
  if (state.ayarlar.solHudut !== "" || state.ayarlar.sagHudut !== "") {
    if ($("sol-hudut")) $("sol-hudut").value = state.ayarlar.solHudut ?? "";
    if ($("sag-hudut")) $("sag-hudut").value = state.ayarlar.sagHudut ?? "";
  }
  hedefListesi();
  gorevListesi();
  updateBandLabel();
  updateHudutHint();
  renderShots();
  updateActionButtons();

  $("btn-oner").addEventListener("click", oner);
  $("btn-hesapla").addEventListener("click", hesapla);
  $("btn-mevzi-kaydet").addEventListener("click", saveMevzi);
  $("btn-mevzi-sifirla").addEventListener("click", resetMevzi);
  const restoreDefaultMetro = async () => {
    try {
      const metro = await loadDefaultMetroText();
      saveMetro(metro, "varsayılan METRAP.txt", "default");
      if ($("metro-file")) $("metro-file").value = "";
      if ($("ay-metro-file")) $("ay-metro-file").value = "";
      showError("");
    } catch (e) {
      showError("Varsayılan METRAP yüklenemedi: " + (e.message || e));
    }
  };
  $("btn-metro-default")?.addEventListener("click", restoreDefaultMetro);
  $("ay-metro-default")?.addEventListener("click", restoreDefaultMetro);
  $("btn-dzl-hesapla").addEventListener("click", dzlUygula);
  $("btn-dzl-temizle").addEventListener("click", dzlTemizle);
  $("btn-atildi").addEventListener("click", atildi);
  $("btn-gorev-bitir").addEventListener("click", gorevBitir);
  $("btn-gorev-bitir-dzl")?.addEventListener("click", gorevBitir);
  $("secili-top").addEventListener("change", () => setSelectedGun($("secili-top").value));
  $("musterek-yan").addEventListener("input", updateBandLabel);
  $("topcu").addEventListener("input", updateMissionLabels);
  $("sol-hudut")?.addEventListener("change", saveHudut);
  $("sol-hudut")?.addEventListener("input", updateHudutHint);
  $("sag-hudut")?.addEventListener("change", saveHudut);
  $("sag-hudut")?.addEventListener("input", updateHudutHint);

  $("btn-ay-kaydet")?.addEventListener("click", kaydetAyarlar);
  $("btn-ay-sifirla")?.addEventListener("click", sifirlaAyarlar);
  $("ay-silah")?.addEventListener("change", () => {
    const w = WEAPONS[$("ay-silah").value] || WEAPONS.obus_bh567;
    if ($("ay-silah-hint")) $("ay-silah-hint").textContent = w.hint || "";
  });

  const onMetroFile = async (input) => {
    const f = input.files?.[0];
    if (!f) return;
    try {
      const text = await f.text();
      if (!text.trim()) throw new Error("Boş METRAP dosyası");
      saveMetro(text, f.name, "custom");
      showError("");
    } catch (e) {
      showError("METRAP yüklenemedi: " + (e.message || e));
    }
  };
  $("metro-file")?.addEventListener("change", (ev) => onMetroFile(ev.target));
  $("ay-metro-file")?.addEventListener("change", (ev) => onMetroFile(ev.target));

  $("btn-hl-kaydet")?.addEventListener("click", hedefKaydetHl);
  $("btn-hl-iptal")?.addEventListener("click", () => {
    clearHlForm();
    hlStatus("");
  });
  $("btn-hl-mevcut")?.addEventListener("click", hlMevcutAl);
  $("btn-hl-toplu")?.addEventListener("click", () => tumHedeflereHesaplaExcel());

  $("btn-rakim-al")?.addEventListener("click", async () => {
    try {
      await fetchRakimInto(
        {
          bolgeNuId: "hedef-bolge-nu",
          bolgeSayisiId: "hedef-bolge-sayisi",
          sagId: "hedef-sag",
          yukariId: "hedef-yukari",
          rakimId: "hedef-rakim",
          busyBtn: $("btn-rakim-al"),
        },
        "calc"
      );
    } catch (e) {
      showError(e.message || String(e));
    }
  });
  $("btn-hl-rakim")?.addEventListener("click", async () => {
    try {
      await fetchRakimInto(
        {
          bolgeNuId: "hl-bolge-nu",
          bolgeSayisiId: "hl-bolge-sayisi",
          sagId: "hl-sag",
          yukariId: "hl-yukari",
          rakimId: "hl-rakim",
          busyBtn: $("btn-hl-rakim"),
        },
        "hl"
      );
    } catch (e) {
      hlStatus(e.message || String(e), true);
    }
  });

  let rakimTimer = null;
  const scheduleAutoRakim = () => {
    clearTimeout(rakimTimer);
    rakimTimer = setTimeout(autoRakimAtis, 600);
  };
  ["hedef-sag", "hedef-yukari", "hedef-bolge-nu", "hedef-bolge-sayisi"].forEach((id) => {
    $(id)?.addEventListener("change", scheduleAutoRakim);
  });

  // Seçili topun MUSYAN'ını müşterek yana öner
  document.querySelectorAll("[data-mevzi]").forEach((box) => {
    const mus = box.querySelector('[data-f="musyan"]');
    if (mus) {
      mus.addEventListener("change", () => {
        if (box.getAttribute("data-mevzi") === selectedId()) {
          $("musterek-yan").value = mus.value;
          updateBandLabel();
        }
      });
    }
  });

  document.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter" && ev.target.matches("input") && !ev.target.closest(".mission-card")) {
      ev.preventDefault();
      hesapla();
    }
  });

  try {
    await loadData();
  } catch (e) {
    showError("Veri yüklenemedi: " + e.message);
  }

  if ("serviceWorker" in navigator) {
    try {
      const reg = await navigator.serviceWorker.register("./sw.js");
      $("sw-status").title = reg.active ? "Çevrimdışı hazır" : "SW kaydı";
    } catch {
      $("sw-status").style.color = "#c45c26";
    }
  }
}

main();
