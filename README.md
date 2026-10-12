# OBS Yoklama Oluşturucu — Web

OBS şube öğrenci listelerinden **Excel ve A4 yatay PDF** yoklama listeleri üreten, sunucusuz GitHub Pages uygulaması.

## Kullanım

1. Web sitesini açın; OBS üzerinden indirdiğiniz `.xlsx` öğrenci listelerini sürükleyip bırakın.
2. Haftalık imza sayısını **1, 2 veya 3** seçin.
3. Eğitim öğretim yılını ve **Güz/Bahar** dönemini kontrol edin; tarayıcının yerel tarihine göre otomatik seçilir (Eylül–Ocak Güz, Şubat–Ağustos Bahar), elle değiştirilebilir.
4. Ders kodu ve adı ile öğretim elemanını girin.
5. **Yoklamaları oluştur** düğmesini seçin.
6. Çıktıları ayrı ayrı Excel/PDF veya tek ZIP olarak indirin (ör. `BLG103 - Programlama I Yoklamaları.zip`).

Dosyalar tarayıcının içindeki sanal dosya sisteminde işlenir. Öğrenci kayıtları GitHub'a veya uygulamanın bir sunucusuna yüklenmez.

## Mimari

- `index.html` / `assets/app.css`: duyarlı web arayüzü
- `assets/app.js`: dosya seçimi, Pyodide yükleyici, istemci tarafında PDF/ZIP üretimi
- `assets/pdf.css`: A4 yatay yoklama tablosu stilleri
- `yoklama_olustur.py`: orijinal OBS Excel okuma, şablon doldurma ve sayfalama algoritması
- `web_bridge.py`: Pyodide dosya sistemi adaptörü ve PDF sayfaları için şablon meta verisi
- `templates/`: orijinal 1, 2 ve 3 imzalı Excel şablonları

İlk açılışta Pyodide (0.27.7), openpyxl (3.1.5), jsPDF, html2canvas ve JSZip CDNs üzerinden indirilir. Başlangıç için internet bağlantısı gerekir; uygulama backend gerektirmez.

## Orijinal Python programına göre uyarlamalar

Excel çıktısı mevcut Python ve `openpyxl` algoritmasını kullanır; öğrenci sayıları, boş satırlar, Türkçe büyük harf, program alanı, ad sığdırma, sayfa sonları ve birleştirilmiş sınav hücreleri korunur. PDF ise Windows Excel COM yerine ayrı bir **HTML → canvas → PDF** motoruyla üretilir. PDF, A4 yatay ve şablonların sayfa başına öğrenci kapasitesini kullanır; Excel'in çıktı motorunun piksel-piksel aynısı olduğu iddia edilmez. PDF içinde metin görsel olarak bulunur, doğrudan seçilebilir değildir.

### Şablon sayfalama kapasitesi

| İmza | Sayfada öğrenci satırı | Çalışma sayfası |
| --- | ---: | --- |
| 1 | 17 | main |
| 2 | 16 | Vize Öncesi, Vize Sonrası |
| 3 | 23 | Vize Öncesi, Vize Sonrası |

### OBS girdi sütunları

Başlık satırında `Öğrenci No`, `Adı`, `Soyadı`, `Program` aranır. OBS'nin `_XXXX` gibi rastgele başlık son ekleri desteklenir. Hatalı dosyalar bağımsız raporlanır, sağlam dosyalar işlenmeye devam eder.

## GitHub Pages yayımlama

Bu proje `baturyusuf/testrepo` deposunun `gh-pages` dalına aktarılmıştır. Yayımlamak için bir kere şu ayarı yapın:

1. [Repo → Settings → Pages](https://github.com/baturyusuf/testrepo/settings/pages) sayfasını açın.
2. **Build and deployment → Source**: `Deploy from a branch`.
3. **Branch**: `gh-pages`, klasör: `/(root)` ve **Save**.
4. Yayımlama tamamlandığında adres: `https://baturyusuf.github.io/testrepo/`.

`.github/workflows/pages.yml`, uygulama daha sonra kendi bağımsız deposuna taşınır ve varsayılan `main` dalında GitHub Actions ile yayımlanırsa kullanılabilecek isteğe bağlı iş akışıdır. Mevcut `gh-pages` dağıtımı için Actions seçmeyin.

Yerel önizleme:

```sh
python -m http.server 8000
# http://localhost:8000
```

> Tarayıcıda `file://` ile doğrudan açmayın; `fetch` için HTTP gerekir.

## Güvenlik ve sınırlamalar

- Gerçek öğrenci listelerini hiçbir zaman repoya eklemeyin. Sadece örnek, anonimleştirilmiş test verisi kullanın.
- Dosya boyutu ve işleme süresi kullanıcı cihazının RAM/CPU kapasitesine bağlıdır.
- PDF HTML tablosunun görüntüsünden üretildiğinden metin tabanlı/erişilebilir PDF değildir.
- Kütüphanelerin ilk kurulumu için internet bağlantısı gerekir.
- Bütün öğrenci bilgileri yalnızca kullanıcının tarayıcısında işlenir; hata ayıklama için tarayıcı konsoluna tam öğrenci verileri basılmaz.

## Lisans

Bu proje kuruma özel yoklama oluşturma aracının web uyarlamasıdır. Dağıtım ve şablon kullanımı için kurum politikalarını uygulayın.
## Akademik dönem ve dosya adları

- Eğitim yılı Eylül ayında değişir; Ocak **Güz**, Şubat–Ağustos **Bahar** varsayılanıdır. Tarih kullanıcı cihazından alınır, seçim elle değiştirilebilir.
- Seçilen eğitim yılı ve dönem, Excel/PDF başlıklarına aktarılır.
- Sol menüde seçilen eğitim yılı ve geliştirici bilgisi (Arş. Gör. Hüseyin Eren Demirtaş) görüntülenir.
- İndirilen ZIP adı işlem sırasında girilmiş ders kodundan/adından oluşturulur (örn. `BLG103 - Programlama I Yoklamaları.zip`). ZIP içindeki Excel/PDF yapısı değişmez.

## Testler

```sh
node tests/test_helpers.cjs
python tests/test_web_bridge.py
```

Python testi için `openpyxl` yüklü olmalıdır.

## Kurumsal görsel kimlik

- Sol menü ve ana ekran İstanbul Rumeli Üniversitesi'nin özgün bordo-kırmızı ve beyaz renkleriyle düzenlenmiştir.
- `assets/rumeli-logo.png`, kullanıcı tarafından sağlanan üniversite amblemidir; hem sayfada hem sekme favicon'unda kullanılır.
- Üretim sürerken animasyonlu işaret gösterilir; tamamlanan işlerde sabit yeşil tik, hatalarda uyarı işareti görünür.

## Referans görseline uygun arayüz (v4)

- Beyaz kurumsal yan panel, kampüs arka planı ve sloganlar
- Bordo/kırmızı tanıtım bandı ve üç aşamalı kullanım akışı
- Yan yana OBS dosya yükleme ve ders/dönem bilgi kartları
- Ders kodu ve ders adı ayrı yazılır; mevcut Python algoritmasına birleşik ders metni gönderilir
- Haftalık imza sayısı 1/2/3 olarak seçilir. Hafta sayısı, orijinal şablonların **16 haftalık** yapısı nedeniyle sabittir.
- Excel, PDF ve tüm dosyalar için ayrı indirme kontrolleri (birden fazla çıktıda ZIP)
- Sağlam çalışma için yardım penceresi, sürükle bırak, dosya kaldırma ve işlem durumları
- Geliştirici bilgisi: Arş. Gör. Hüseyin Eren Demirtaş

Kullanıcı tarafından iletilen arayüz referansındaki dekoratif görsellerin kırpılmış halleri `assets/` klasöründe bulunmaktadır.
