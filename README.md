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
## Tam ekran A4 önizleme ve dengeli bölümler (v9)

- Önizlemenin **Tam ekran** düğmesi önce tarayıcının Fullscreen API'sini kullanır; engellenirse ekranı kaplayan yerleşime geçer. **Küçült** düğmesi veya Escape ile çıkılır.
- Masaüstünde ayarlar / önizleme genişlikleri yaklaşık **%42 / %58** oranındadır. Dar ekranlarda kartlar tek sütuna dönüşür.
- `page_plan` (Python) ve `paperLayout` (JavaScript) artık açgözlü sıralama yerine **önce en az sayfayı, ardından sütun dağılımındaki kareler toplamını küçülten dinamik programlamayı** kullanır. Haftalar bölünmez, kullanıcı tarafından belirlenen asgari imza alanı korunur.
- Örnek: 16 hafta, bir imza, 8. hafta vize ve 16. hafta final, asgari 15 mm ile önceden oluşan 13+3 sütun yerine **8+8 sütun** oluşturulur.
- Başlık bölgeleri, dar yatay bölümlerde birbirine girmesin diye Excel'de fiziksel sütun genişliğine göre tek sütunlu açıklama satırlarına dönüşür; PDF ve önizleme de bu durumu destekler.
- Yerleşim regresyonları `.github/workflows/validate-layout.yml` ile otomatik çalıştırılır.

## Canlı kaydırıcı ayarları (v10)

"Yazdırma ve Sütun Ayarları" alanındaki **Asgari İmza Genişliği** (8–20 mm, 15 mm varsayılan), **Yazı Büyüklüğü** (6–14 pt, 8 pt varsayılan) ve **Ad Soyad Azami Genişliği** (30–80 mm, 58 mm varsayılan; ikişer mm adımlarla) artık erişilebilir HTML range kaydırıcılarıdır. Güncel değer, kaydırıcının yanında görüntülenir. Her `input` olayında ölçü önizlemesi ve üretim geçerliliği yenilenir; canlı A4 çizimi animasyon karesine göre güncellenir. Yazı tipi listesi kategorik bir tercih olduğu için açılır listede kalır.

## Tek imzalı hafta başlıkları (v11)

Haftalık imza sayısı **1** seçildiğinde Excel'de hafta başlığı 5. ve 6. satır boyunca birleştirilir; ikinci satıra `1. İMZA` yazılmaz. PDF ve canlı A4 önizlemesi aynı başlığı tek ve iki satır yüksekliğinde gösterir. Haftalık imza sayısı **2–5** ise numaralı imza alt başlıkları korunur.


## Belirli haftadan başlatma (v12)

Yoklama listesini dönem ortasında yeniden oluştururken **Başlangıç Haftası** seçilebilir (varsayılan: 1).
Seçilen haftadan **önceki imza haftaları**, her A4 sayfasının 17 öğrenci satırını kapsayan tek gri
blokla imzaya kapatılır. Gri blokta gerekçe yazılıdır; dar bloklarda açıklama dikeydir.
Hafta ve imza başlıkları soluk ve üstü çizili gösterilir. Vize/final sütunları değişmez;
sınav haftası bir kapanan aralığı iki ayrı bloğa böler. İmza atılacak haftası kalmayan
yatay bölümler Excel, PDF ve önizlemeden çıkarılır, ancak diğer bölümlerin adları,
sütun genişlikleri ve A4 sayfa planı değişmez. Başlangıç haftasını değiştirmek
oluşturulmuş sonuçları geçersiz kılar.


## Kapalı haftalarda sütun tasarrufu (v13)

Başlangıç haftası 1'den büyük olduğunda, daha önce işlenmiş **her normal hafta**,
haftalık imza sayısı 1–5 arasında ne olursa olsun **9 mm genişliğinde tek bir sütun**
kullanır. Sütun başlıkları haftayı göstermeye devam eder; bitişik kapalı sütunların
17 öğrenci satırındaki gövde alanı, v12'deki uyarı metnini içeren tek gri blok olarak
birleştirilir. Vize/final **sınav** sütunları daraltılmaz; ayrıca aradaki vize
gri bölgeyi ayrı bloklara böler.

A4 yatay sayfa planı fiziksel genişliklerle **yeniden hesaplanır**:
kapalı normal hafta = 9 mm, açık haftanın her imzası ve sınavlar = kullanıcı
tarafından belirlenen en az 8–20 mm. Açık imzalar, mümkünse 18 mm'ye
kadar genişletilir. Önizleme, Excel ve PDF aynı grup/sütun genişliğini
kullanır; tamamen kapalı yatay bölümler v12'deki gibi atlanır.
Yazdırma alanı, 17 öğrenci/sayfa, başlık ve vize/final ayrımı korunur.
Başlangıç haftası **1** olduğunda v12 Excel XML'i ve PDF HTML'i değişmez.

Bu davranış [PR #2 geri bildirimi](https://github.com/baturyusuf/testrepo/pull/2#issuecomment-6578850105)
üzerine eklendi. Regresyonlar: `tests/test_compact_weeks.py`,
`tests/test_compact_weeks.cjs`, `tests/test_compact_parity.cjs`.

## PDF baskı netliği ve Excel renk uyumu (v14)

PDF dışa aktarımı artık önceki JPEG %89 / 1,7× (~163 DPI) yerine
**kayıpsız PNG / 3× (~288 DPI)** kullanır. `canvas.toBlob` ve
`Uint8Array` bellek tüketimini sınırlamak amacıyla base64 oluşturmadan
jsPDF'ye geçirilir. Seçilen font yüklenmeden ekran görüntüsü alınmaz;
her A4 sayfası ayrı işlenir, canvas işi bitince bellekten bırakılır ve hata
durumunda gizli PDF sahnesi temizlenir. PDF hâlâ görüntü tabanlıdır,
metin seçilebilir değildir; tam vektörel PDF ayrı bir geliştirmedir.

`assets/pdf.css` dosyasında tek bir `v14 print fidelity` bölümünde
`adaptive_layout.py` renk sabitleriyle aynı değerler kullanılır
(bordo `9B0D22`, açık pembe `F7E6E9`, satır grisi `F0F1F3`,
sınav grisi `D3D4DA`, kenarlık `A6A6AD`). Varsayılan siyah/kalın
çizgiler inceltildi; sıra ve numara Excel'deki gibi normal, öğrenci ismi
kalın ağırlıkta tutuldu. Aynı CSS hem A4 önizlemede hem PDF üretiminde
kullanılır. **Excel oluşturma, hafta planı, 9 mm kapalı sütun ve
17 öğrenci/sayfa mantığında değişiklik yapılmadı.**

Not: Yaklaşık 288 DPI PNG dosyaları, önceki JPEG çıktılardan daha
büyük olabilir ve PDF üretimi düşük bellekli cihazlarda daha uzun
sürebilir. Tarayıcı yazdırmasında "arka plan grafikleri" seçeneğinin
açılması ve gerçek yazıcı baskısının manuel kontrolü önerilir.
Gerçek öğrenci bilgileri otomatik test deposuna eklenmez.
