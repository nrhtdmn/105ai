#!/usr/bin/env python3
"""Extract ballistic library tables from atesidare.py into editable JSON files."""

from __future__ import annotations

import ast
import json
import re
import shutil
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "atesidare.py"
OUT_DIR = ROOT / "web" / "data"
METRAP_SRC = ROOT / "METRAP.txt"

# First (group-1) occurrence line ranges (1-based, inclusive end for slicing)
CHARGE_RANGES = {
    5: (1292, 3737),
    6: (3739, 6019),
    7: (6021, 8214),
}

SHARED_RANGE = (775, 1176)

CHARGE_RECOMMEND_BREAKS = [2348, 2766, 3354, 4218, 5363, 6694, 11000]

ASSIGN_RE = re.compile(
    r"^\s*(?P<name>data1|data2|isi_dzl|yogunluk_dzl|brt_isisi_ilk_hiz_dzl|"
    r"a3bh|liste4bh|liste5bh|liste6bh|liste7bh)\s*=\s*[\[(]",
    re.MULTILINE,
)


def lines_slice(text: str, start: int, end: int) -> str:
    """Return source lines start..end (1-based inclusive)."""
    lines = text.splitlines(keepends=True)
    return "".join(lines[start - 1 : end])


def extract_bracket_literal(src: str, assign_match: re.Match[str]) -> Any:
    """Extract and literal_eval the list/tuple starting at an assignment match."""
    eq = src.find("=", assign_match.start())
    i = eq + 1
    while i < len(src) and src[i].isspace():
        i += 1
    if i >= len(src) or src[i] not in "[(":
        raise ValueError(f"No opener after assignment at {assign_match.start()}")

    opener = src[i]
    closer = "]" if opener == "[" else ")"
    depth = 0
    in_str: str | None = None
    escape = False
    j = i
    while j < len(src):
        ch = src[j]
        if in_str:
            if escape:
                escape = False
            elif ch == "\\":
                escape = True
            elif ch == in_str:
                in_str = None
        else:
            if ch in ("'", '"'):
                in_str = ch
            elif ch == opener:
                depth += 1
            elif ch == closer:
                depth -= 1
                if depth == 0:
                    literal = src[i : j + 1]
                    # Normalize odd spacing; keep numeric fidelity via literal_eval
                    return ast.literal_eval(literal)
        j += 1
    raise ValueError(f"Unclosed literal for {assign_match.group('name')}")


def find_assignments(block: str) -> list[tuple[str, Any, int]]:
    """Return (name, value, char_offset) for each matching assignment in block."""
    results: list[tuple[str, Any, int]] = []
    for m in ASSIGN_RE.finditer(block):
        name = m.group("name")
        value = extract_bracket_literal(block, m)
        results.append((name, value, m.start()))
    return results


def nrows(value: Any) -> int:
    if isinstance(value, (list, tuple)):
        return len(value)
    return 0


def ncols(row: Any) -> int:
    if isinstance(row, (list, tuple)):
        return len(row)
    return 0


def to_jsonable(value: Any) -> Any:
    """Convert tuples to lists recursively for JSON."""
    if isinstance(value, tuple):
        return [to_jsonable(v) for v in value]
    if isinstance(value, list):
        return [to_jsonable(v) for v in value]
    return value


def classify_charge_tables(assignments: list[tuple[str, Any, int]]) -> dict[str, Any]:
    """Map raw assignments inside one charge block to semantic table names."""
    tables: dict[str, Any] = {}
    failed: list[str] = []

    # Named assignments (unique expected)
    named_once = {
        "data2": "complementaryRange",
        "isi_dzl": "tempAltFactor",
        "yogunluk_dzl": "densityAltFactor",
        "brt_isisi_ilk_hiz_dzl": "propellantTempMv",
    }
    for name, value, _ in assignments:
        key = named_once.get(name)
        if key and key not in tables:
            tables[key] = value

    data1s = [(v, off) for n, v, off in assignments if n == "data1"]

    # First data1 = B-cetveli / altitude zone
    if data1s:
        tables["altitudeZoneTable"] = data1s[0][0]
    else:
        failed.append("altitudeZoneTable")

    wind = [v for v, _ in data1s if nrows(v) and ncols(v[0]) == 3]
    if wind:
        tables["windComponents"] = wind[0]
    else:
        failed.append("windComponents")

    f17 = [v for v, _ in data1s if nrows(v) and ncols(v[0]) == 17]
    if f17:
        tables["groundFire"] = f17[0]
        if len(f17) > 1:
            tables["extras"] = f17[1]
    else:
        failed.append("groundFire")

    yan = [v for v, _ in data1s if nrows(v) and ncols(v[0]) == 10]
    if yan:
        tables["earthRotationYan"] = yan[0]
    else:
        failed.append("earthRotationYan")

    rng = [v for v, _ in data1s if nrows(v) and ncols(v[0]) == 18]
    if rng:
        tables["earthRotationRange"] = rng[0]
    else:
        failed.append("earthRotationRange")

    fuze = [v for v, _ in data1s if nrows(v) and ncols(v[0]) == 11]
    if fuze:
        tables["fuzeSecondary"] = fuze[0]
    # optional — do not mark failed if absent

    drop = [v for v, _ in data1s if nrows(v) and ncols(v[0]) == 5]
    if drop:
        tables["dropAngleTac"] = drop[0]
    else:
        failed.append("dropAngleTac")

    onemil = [v for v, _ in data1s if nrows(v) and ncols(v[0]) == 2]
    if onemil:
        tables["oneMilChange"] = onemil[0]
    # BH7 grup-1 bloğunda 2 sütunlu tablo yok; dzl_hesapla liste7bh kullanılır (aşağıda)

    for expected in (
        "complementaryRange",
        "tempAltFactor",
        "densityAltFactor",
        "propellantTempMv",
    ):
        if expected not in tables:
            failed.append(expected)

    return tables, failed


def extract_shared(block: str) -> dict[str, Any]:
    assignments = find_assignments(block)
    by_name = {}
    for name, value, _ in assignments:
        if name not in by_name:
            by_name[name] = value

    mapping = {
        "a3bh": "3",
        "liste4bh": "4",
        "liste5bh": "5",
        "liste6bh": "6",
        "liste7bh": "7",
    }
    dogal: dict[str, Any] = {}
    failed: list[str] = []
    for src_name, charge_key in mapping.items():
        if src_name in by_name:
            dogal[charge_key] = by_name[src_name]
        else:
            failed.append(f"dogalYan[{charge_key}] ({src_name})")
    return {
        "dogalYan": dogal,
        "meta": {
            "chargeRecommendBreaks": CHARGE_RECOMMEND_BREAKS,
            "squareWeightBaseline": 2,
            "milCircle": 6400,
        },
    }, failed


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
        f.write("\n")


def summarize_tables(tables: dict[str, Any]) -> dict[str, int]:
    return {k: nrows(v) for k, v in tables.items() if isinstance(v, (list, tuple))}


def write_readme(path: Path) -> None:
    text = """# Balistik kütüphane verileri

Bu klasördeki JSON dosyaları `atesidare.py` içindeki ateş idare tablolarından üretilmiştir.
PWA yeniden yazımında doğrudan düzenlenebilir kaynak olarak kullanılır.

## Dosyalar

| Dosya | İçerik |
|-------|--------|
| `charge_5.json` | 5. barut hakkı tabloları |
| `charge_6.json` | 6. barut hakkı tabloları |
| `charge_7.json` | 7. barut hakkı tabloları |
| `shared.json` | Ortak doğal yan tabloları (3–7) ve meta |
| `METRAP.txt` | Metro/atmosfer girdi dosyası |

## Tablo anahtarları (charge_*.json)

- `altitudeZoneTable` — B cetveli (irtifa / bölge)
- `complementaryRange` — tamamlayıcı mesafe (data2)
- `tempAltFactor` — sıcaklık–irtifa düzeltme (`isi_dzl`)
- `densityAltFactor` — yoğunluk–irtifa düzeltme (`yogunluk_dzl`)
- `windComponents` — azimut rüzgar dairesi (3 sütun)
- `groundFire` — F cetveli (17 sütun)
- `earthRotationYan` — arz dönüşü yan
- `earthRotationRange` — arz dönüşü mesafe
- `propellantTempMv` — barut ısısı / ilk hız (`brt_isisi_ilk_hiz_dzl`)
- `fuzeSecondary` — tapa ikincil düzeltmeler (varsa)
- `dropAngleTac` — düşüş açısı / TAÇ
- `extras` — ek F cetveli / MSO benzeri (varsa)
- `oneMilChange` — 1 milyemlik değişim (varsa)

Her satır bir dizi: ilk eleman genelde mesafe (m) veya girdi anahtarıdır; sonraki elemanlar ilgili katsayılardır.

## shared.json

- `dogalYan["3"…"7"]` — doğal yan (mesafe → değer) çiftleri
- `meta.chargeRecommendBreaks` — barut hakkı öneri eşikleri
- `meta.squareWeightBaseline` — kare ağırlık tabanı (2)
- `meta.milCircle` — milyem dairesi (6400)

## Nasıl düzenlenir?

1. İlgili JSON dosyasını bir metin editöründe açın.
2. Sayısal değerleri **ondalık nokta** ile değiştirin (ör. `12.5`). Satır yapısını (sütun sayısını) bozmayın.
3. Yeni satır eklerken aynı sütun sayısını koruyun; mesafeleri artan sırada tutun.
4. Kaydettikten sonra uygulamayı yenileyin / yeniden derleyin.
5. Kaynak Python değiştiyse tabloları yeniden üretmek için:

```bash
python tools/extract_libraries.py
```

> Uyarı: Bu betik `atesidare.py` içindeki **birinci grup** (ilk charge blokları) değerlerini yazar. Elle yaptığınız JSON düzenlemeleri üzerine yazılır.
"""
    path.write_text(text, encoding="utf-8")


def main() -> int:
    if not SOURCE.is_file():
        raise SystemExit(f"Source not found: {SOURCE}")

    text = SOURCE.read_text(encoding="utf-8", errors="replace")
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    all_failed: dict[str, list[str]] = {}
    summaries: dict[str, dict[str, int]] = {}

    for charge, (start, end) in CHARGE_RANGES.items():
        block = lines_slice(text, start, end)
        assignments = find_assignments(block)
        tables, failed = classify_charge_tables(assignments)
        payload = {
            "charge": charge,
            "sourceRange": [start, end],
            "group": 1,
            **{k: to_jsonable(v) for k, v in tables.items()},
        }
        out = OUT_DIR / f"charge_{charge}.json"
        write_json(out, payload)
        summaries[f"charge_{charge}"] = summarize_tables(tables)
        if failed:
            all_failed[f"charge_{charge}"] = failed
        print(f"\n=== charge {charge} ({out}) ===")
        for name, count in summaries[f"charge_{charge}"].items():
            print(f"  {name}: {count} rows")
        if failed:
            print(f"  FAILED: {', '.join(failed)}")

    # BH7 oneMilChange: dzl_hesapla içindeki liste7bh (0,42)...(11000,0)
    c7_path = OUT_DIR / "charge_7.json"
    if c7_path.is_file():
        c7 = json.loads(c7_path.read_text(encoding="utf-8"))
        if "oneMilChange" not in c7:
            dzl = lines_slice(text, 15344, 15455)
            m = re.search(r"liste7bh\s*=\s*\[", dzl)
            if m:
                rows = extract_bracket_literal(dzl, m)
                c7["oneMilChange"] = to_jsonable(rows)
                write_json(c7_path, c7)
                print(f"\nPatched charge_7 oneMilChange: {nrows(rows)} rows")
                summaries.setdefault("charge_7", {})["oneMilChange"] = nrows(rows)

    shared_block = lines_slice(text, *SHARED_RANGE)
    shared_payload, shared_failed = extract_shared(shared_block)
    shared_out = OUT_DIR / "shared.json"
    write_json(
        shared_out,
        {
            "dogalYan": {
                k: to_jsonable(v) for k, v in shared_payload["dogalYan"].items()
            },
            "meta": shared_payload["meta"],
        },
    )
    dogal_counts = {
        f"dogalYan[{k}]": nrows(v) for k, v in shared_payload["dogalYan"].items()
    }
    summaries["shared"] = dogal_counts
    if shared_failed:
        all_failed["shared"] = shared_failed
    print(f"\n=== shared ({shared_out}) ===")
    for name, count in dogal_counts.items():
        print(f"  {name}: {count} rows")
    print(f"  meta.chargeRecommendBreaks: {CHARGE_RECOMMEND_BREAKS}")
    if shared_failed:
        print(f"  FAILED: {', '.join(shared_failed)}")

    if METRAP_SRC.is_file():
        dest = OUT_DIR / "METRAP.txt"
        shutil.copy2(METRAP_SRC, dest)
        print(f"\nCopied METRAP.txt -> {dest}")
    else:
        all_failed.setdefault("METRAP", []).append("METRAP.txt missing at repo root")
        print("\nWARNING: METRAP.txt not found at repo root")

    readme = OUT_DIR / "README.md"
    write_readme(readme)
    print(f"Wrote {readme}")

    print("\n========== SUMMARY ==========")
    for key, counts in summaries.items():
        print(f"{key}: {counts}")
    if all_failed:
        print("\nTables that failed / missing:")
        for key, items in all_failed.items():
            print(f"  {key}: {items}")
    else:
        print("\nAll expected tables extracted.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
