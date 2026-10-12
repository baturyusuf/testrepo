# İstanbul Rumeli Üniversitesi — Yoklama Oluşturucu

[Uygulamayı aç](https://baturyusuf.github.io/testrepo/) · [GitHub Pages kaynak dalı](https://github.com/baturyusuf/testrepo/tree/gh-pages)

OBS `.xlsx` sınıf listelerinden **A4 yatay Excel, PDF ve ZIP** yoklamaları oluşturan sunucusuz tarayıcı uygulaması. Öğrenci bilgileri yalnızca cihazdaki Pyodide (WebAssembly) sanal dosya sisteminde işlenir.

## Kullanım

1. OBS `.xlsx` öğrenci listelerini yükleyin; birden çok liste seçilebilir.
2. Eğitim yılı, Güz/Bahar dönemi, ders kodu, ders adı ve öğretim üyesini girin.
3. **Hafta sayısını 10–20**, **haftalık imza sayısını 1–5**, **vize ve final haftalarını** seçin. Varsayılanlar sırasıyla **16, 1, 8 ve 16**. Vize haftası finalden önce olmalıdır. Ayrıca asgari imza genişliğini, yazı tipini/puntosunu ve Ad Soyad azami genişliğini seçin.
4. Gösterilen A4 planını kontrol edin ve **Yoklama Dokümanlarını Oluştur** seçeneğine basın.
5. Excel veya PDF dosyalarını ayrı ayrı ya da hepsini ZIP olarak indirin.

Eğitim yılı ve yarıyıl bilgisinin varsayılan değeri tarayıcının yerel tarihine bağlıdır: Eylül–Ocak için Güz, Şubat–Ağustos için Bahar. Alanlar manuel değiştirilebilir.

## Uyarlanabilir sayfalama algoritması

`adaptive_layout.py`, tüm seçimler için Excel çalışma sayfalarını doğrudan oluşturur. Önceki 1/2/3 imzalı sabit şablonlar arşivde tutulur ancak yeni çalışma mantığında **kullanılmaz**. Vize ve final haftaları normal imza alanları yerine birer sınav sütunu oluşturur.

Her A4 yatay sayfası **297 × 210 mm** olup kâğıt kenarlarında 4 mm yatay boşluk ayrılır. Algoritma şu ölçüleri kullanır:

| Alan | Varsayılan / sınır |
| --- | ---: |
| Sıra numarası | **7,5 mm** (eski 10 mm) |
| Öğrenci numarası | **21 mm** (eski 26 mm) |
| Ad Soyad | **En uzun öğrenci adına göre 28–58 mm**; 58 mm varsayılan üst sınır, kullanıcı 30–80 mm seçebilir |
| Asgari imza/sınav hücresi | **15 mm**, kullanıcı 8–20 mm seçebilir |
| Boş yer varsa imza hücresi büyütülebilir | **18 mm'ye kadar** (asgari genişlik 18 mm'den büyükse seçilen değer geçerlidir) |
| Ad Soyad varsayılan puntosu | **8 pt** (eski 10 pt), kullanıcı 6–14 pt seçebilir |
| Yazı tipi | Aptos, Arial, Calibri, Times New Roman |
| Öğrenci satırı yüksekliği | **23,5 pt** (~8,3 mm) |
| Bir dikey A4 öğrenci alanı | **17 satır** |

Ad Soyad genişliği, listedeki tüm öğrenci isimlerinin tahmini fiziksel metin uzunlukları ölçülerek tek seferde hesaplanır. Satır numarası ile öğrenci numarası sütunlarının yeni sabit değerleri de A4 hesabına dahil edilir. Ad Soyad sütunu seçilen maksimumdan daha geniş olamaz; daha kısa isimli sınıflarda daralır. Çok uzun bir isim iki satıra sığmıyorsa sadece ilgili hücredeki punto otomatik düşürülür (en az 6 pt). PDF aynı otomatik noktasal değeri kullanır.

Haftalar sıralı biçimde ve **hafta grupları bölünmeden** çalışma sayfalarına yerleştirilir. Bir hafta için `imza_sayısı` sütunu, sınav haftaları için 1 sütun hesaplanır. Bir sonraki hafta eklendiğinde herhangi bir imza hücresinin genişliği kullanıcının seçtiği alt sınırın altına düşecekse sonraki çalışma sayfasına geçilir. Sütun genişlikleri kalan alana göre hesaplanır; seçilen minimum genişlik korunur ve yer varsa 18 mm’ye kadar genişleyebilir.

Örnek: **20 hafta + 5 imza + vize 8 + final 16** ve **15 mm asgari imza genişliği** → öğrenci isimlerine göre yaklaşık **8–9 yatay bölüm**. Birden çok öğrenci sayfası gerekiyorsa her bölüm ayrıca 17 öğrencilik dikey A4 sayfalarıyla devam eder; öğrenci listesi her yatay bölümde tekrar edilir. Aynı sayfa planı PDF oluşturucusuna aktarılır.

Excel için `openpyxl` ile A4 yönü, baskı alanları, tekrarlanan başlık satırları, satır kesmeleri, hizalama ve sınav sütunları yazılır. PDF için aynı manifest `html2canvas` + `jsPDF` üzerinden işlenir. PDF ile Microsoft Excel'in kendi baskı motorunun piksel düzeyinde birebir eşdeğerliği garanti edilmez; PDF metni raster olarak oluşturulur.

## Mimari

- `index.html`, `assets/app.css`: üniversitenin kırmızı-beyaz temalı duyarlı arayüzü
- `assets/gelenekten-gelecege.png`, `assets/bilimle-daha-ileriye.png`: kullanıcı tarafından sağlanan **saydam PNG** sloganlar
- `assets/helpers.js`: akademik varsayılanlar, ZIP isimleri ve A4 sayfa planı önizlemesi
- `assets/app.js`: dosya seçimi, Pyodide yükleme, Excel/PDF/ZIP indirme işlemleri
- `adaptive_layout.py`: genel amaçlı Python Excel/yatay-dikey sayfalama motoru
- `web_bridge.py`: tarayıcı içi Python köprüsü ve PDF sayfa manifesti
- `yoklama_olustur.py`: OBS dosyası okuma ve Türkçe metin normalleştirme yardımcıları; eski komut satırı akışı arşiv niteliğindedir
- `templates/`: eski 1, 2 ve 3 imzalı Excel şablonları, geriye dönük referans için bulunur

İlk açılışta Pyodide, openpyxl, jsPDF, html2canvas, JSZip CDNs üzerinden yüklenir. Başlangıç için internet bağlantısı gereklidir; son kullanıcıya Python/Excel kurulumu gerekmez.

## Testler

```sh
python tests/test_adaptive_layout.py
node tests/test_helpers.cjs
node tests/test_layout_helpers.cjs
node --check assets/app.js
```

- **880 kombinasyon**: hafta 10–20, imza 1–5 ve çeşitli vize/final haftaları
- **10 entegrasyon çıktısı**: farklı hafta/imza/sınav kombinasyonlarında 1 ve 105 öğrencilik Excel + PDF manifest doğrulaması
- A4 sütun genişliği sınırı, hafta sırası, sınav alanları, baskı satır kesmeleri, ve yeniden açılabilir Excel yapısı test edilir

## GitHub Pages

Bu depo `gh-pages` dalının `/(root)` klasöründen yayınlanmaktadır. Kod bu dala pushlandığında Pages yeniden yayımlanır.

## Güvenlik ve sınırlamalar

Öğrenci listeleri GitHub'a yüklenmemelidir. Uygulama veriyi tarayıcı belleğinde işler. İşlem maliyeti dosya boyutuna ve kullanıcı cihazının CPU/RAM kapasitesine bağlıdır. Çok sayıda hafta/imza kombinasyonunda çıktı sayısı artabileceği için PDF oluşturma süresi uzayabilir.

Geliştiren: **Arş. Gör. Hüseyin Eren Demirtaş**.

## Yeni baskı ayarları (v6)

- `minMarkWidthMm`: 8–20 mm, başlangıç **15 mm**.
- `fontSizePt`: 6–14 pt, başlangıç **8 pt** (önceki 10 pt bazına göre −2).
- `fontFamily`: Aptos (varsayılan), Arial, Calibri, Times New Roman.
- `nameMaxWidthMm`: 30–80 mm, başlangıç **58 mm**.

Tarayıcıdaki A4 planı hesaplaması, henüz OBS listesi okunmadığı için ad sütununu seçilen **azami** genişlikte varsayar. Excel üretildiğinde gerçek en uzun öğrenci adına göre hesaplanıp aynı fiziksel ölçüler PDF manifestine aktarılır; bu nedenle üretilen yatay bölüm sayısı önizlemeden daha az olabilir.

Excel ve PDF için yazı punto bilgisi aynı ayardan beslenir. Chrome üzerinde Aptos bulunmuyorsa PDF önizlemesinde tarayıcının yedek fontu kullanılabilir; Excel dosyası seçilen aileyi font meta verisi olarak taşır.

## Canlı A4 önizlemesi

- Dosya seçmeden çalışır. Tarayıcıda 17 sentetik öğrenci satırı (`9999999 / Öğrenci 1` … `Öğrenci 17`) ile ilk dikey A4 sayfasını gösterir.
- Yıl, dönem, ders kodu, ders adı, öğretim elemanı, hafta, imza sayısı, sınav haftaları, en küçük imza genişliği, font ailesi ve punto ayarlarını anlık dinler.
- Bölüm değiştirici ile tüm yatay çalışma sayfaları gezilebilir. **%100 göster** ile örnek 96 DPI CSS ölçüsünde detaylı inceleme, **Ekrana sığdır** ile ekrana uygun ölçekleme yapılır.
- Önizleme ile gerçek PDF aynı `YoklamaPrintPage` renderer fonksiyonunu kullanır. Sütun genişlikleri, A4 yan kenar boşlukları ve font boyları Excel üretimindeki mm/pt ölçülerinden hesaplanır.
- Gerçek Excel çıktısının adı-soyadı sütunu listedeki en uzun gerçek isme göre daralabileceğinden, dosyasız örnek önizlemesi seçili **azami** Ad Soyad genişliğini varsayar. Yazıcı sürücüsü ve Excel font ölçümü küçük sapmalar yaratabilir. Bu sebeple fiziksel baskının birebir piksel eşitliği garanti edilmez.