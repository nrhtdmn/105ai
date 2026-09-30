# Balistik kütüphane verileri

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
