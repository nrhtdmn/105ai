# Ateş İdare

Obüs atış esası hesaplama uygulaması. Masaüstü PyQt sürümü (`atesidare.py`) ile aynı balistik tabloları kullanan **GitHub Pages PWA** arayüzü `web/` klasöründedir.

## GitHub Pages ile yayınlama

`web/` içeriğini site kökü olarak yayınlayın (önerilen):

1. Bu depoyu GitHub’a itin.
2. **Settings → Pages → Build and deployment**
3. Source: **Deploy from a branch**
4. Branch: `main` (veya kullandığınız dal), folder: **`/web`**
5. Kaydedin; birkaç dakika sonra `https://<kullanıcı>.github.io/<repo>/` adresinde açılır.

Alternatif: yalnızca `web/` içeriğini `gh-pages` dalının köküne kopyalayın.

> `web/.nojekyll` Jekyll işlemesini kapatır; `_` ile başlayan yollar bozulmaz.

PWA için HTTPS gerekir (GitHub Pages sağlar). İlk yüklemeden sonra çevrimdışı çalışır (service worker).

## Yerel önizleme

ES modülleri için basit bir HTTP sunucusu kullanın (file:// yeterli değil):

```bash
cd web
python -m http.server 8080
```

Tarayıcıda: http://localhost:8080/

## Veri dosyaları (`web/data/`)

| Dosya | Açıklama |
|-------|----------|
| `charge_5.json` … `charge_7.json` | Barut hakkı tabloları |
| `shared.json` | Doğal yan + meta (öneri eşikleri) |
| `METRAP.txt` | Varsayılan metro raporu |

Tabloları düzenlemek için ilgili JSON’u açın; sütun sayısını bozmayın, ondalık ayırıcı olarak nokta kullanın. Kaynak Python değiştiyse:

```bash
python tools/extract_libraries.py
```

Ayrıntılar: [`web/data/README.md`](web/data/README.md).

## Uygulama sekmeleri

- **Atış Görevi** — hedef, barut hakkı, HESAPLA, 8 obüs kartı (1B/2B referans)
- **Bt.Sb. Raporu** — 8 mevzi formu (localStorage), METRAP yükleme
- **Düzeltme** — gözlemci sola/sağa/uzat/kısalt/kaldır/indir
- **Esaslar** — TAÇ, nişangah, metro KPI’ları
- **Hedefler** — localStorage hedef listesi
- **Ayarlar** — silah, aktif toplar, hudut, METRAP, varsayılanlar

## Teknik

- Vanilla JS ES modules (`web/js/`)
- Hesap çekirdeği: `calc.js` (charge-5 boru hattı, lib ile parametreli)
- Görünüm: lacivert başlık `#1E3A5F`, zemin `#E6EAEF`
