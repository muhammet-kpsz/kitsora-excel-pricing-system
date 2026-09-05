# 🦊 Kitsora Excel Fiyatlandırma ve Barkod Yönetim Sistemi

![Kitsora Logo](assets/icon.png)

**Kitsora**, e-ticaret satıcıları, pazaryeri entegratörleri ve toptancılar için geliştirilmiş; yüksek performanslı, esnek ve modern bir Excel fiyatlandırma ve barkod yönetim sistemidir. Ürün maliyetlerinizi, kategori bazlı indirimlerinizi, kâr marjlarınızı, varyant gruplarınızı ve ürün barkodlarınızı tek merkezden yöneterek saniyeler içinde on binlerce ürünü güncelleyebilirsiniz.

---

## 🚀 Öne Çıkan Özellikler

### 🎯 Akıllı Sütun Eşleştirme & Hızlı İçe Aktarma
- **Otomatik Sütun Algılama:** Excel dosyanızı yüklediğiniz anda Stok Kodu, Ürün Adı, Kategori, Alış, Satış, İndirimli, Piyasa, Varyant, Stok ve Barkod sütunlarını otomatik olarak tanır.
- **Büyük Veri Desteği:** 100.000+ satırlık devasa Excel dosyalarını akışlı (streaming) motoru ile bellek sorunu yaşamadan son derece hızlı okur ve işler.

### 🏷️ Gelişmiş Barkod Üretim ve Yönetim Sistemi
- **Özel Barkod Ön Eki:** Firmanıza veya markanıza özel barkod ön eki tanımlayabilme (örn: `HFGYM`, üretilen: `HFGYM41393551`).
- **İki Farklı Çalışma Modu:**
  - **Tüm Ürünlere Barkod Ekle/Güncelle:** Dosyadaki tüm satırlara otomatik olarak yeni ve benzersiz barkodlar üretir.
  - **Sadece Seçili Ürünlere Barkod Ekle/Güncelle:** Yalnızca önizleme ekranında işaretlediğiniz ürünlere barkod atar; işaretlenmeyen ürünlerin mevcut barkodları korunur.
- **Varyant Grubu Entegrasyonu:**
  - Bir ürünü seçtiğinizde, aynı **Varyant Grup Kodu**'na (Varyant ID) bağlı olan tüm varyasyonlar (renk, beden vb.) otomatik olarak tespit edilir ve işleme dahil edilir.
  - Her bir varyasyona birbirinden tamamen farklı ve benzersiz bir barkod atanır.
- **Benzersizlik Güvencesi (Global Çakışma Önleyici):**
  - Dosyada önceden mevcut olan tüm barkodlar taranır. Yeni üretilen barkodların hem kendi aralarında hem de mevcut barkodlarla çakışması kesin olarak engellenir.
  - İşlem sonunda otomatik analiz mekanizması çalışarak üretilen barkodların benzersizliğini denetler ve loglar.

### 👁️ Dinamik Ürün Önizleme Ekranı
- **Akıllı Tablo Görünümü:** Hesaplanan fiyatları, indirimleri, stok adetlerini ve mevcut barkodları işlem öncesinde kontrol edebilme.
- **Barkodsuz Ürün Vurgusu:** Barkodu olmayan ürünler dikkat çeken kırmızı **"Barkod Yok"** uyarısıyla gösterilir.
- **Pratik Seçim Araçları:**
  - Checkbox'lar varsayılan olarak boş gelir; kullanıcı tam kontrole sahiptir.
  - Tablodaki **"Seçim" sütun başlığına** tıklayarak sayfadaki ürünleri tek hamlede seçip kaldırabilirsiniz.
  - Sağ tık menüsünden **"Tüm Barkodları Seç"**, **"Sadece Barkodu Olmayanları Seç"** ve **"Tüm Barkod Seçimlerini Kaldır"** kısayolları.
  - Sayfa değiştirdiğinizde veya arama yaptığınızda seçimleriniz hafızada güvenle korunur.
- **Varyant Grubu Detayları:** Tabloda varyant koduna tıklayarak gruptaki tüm varyasyonları açılan pencerede listeleyebilme.

### 🌳 Kategori Ağacı ve İndirim Yönetimi
- **Görsel Kategori Ağacı:** Excel'deki kategorileri derinlemesine ayrıştırarak hiyerarşik bir ağaç yapısında sunar (`Elektronik > Bilgisayar > Donanım`).
- **Kategori Bazlı İndirim:** Her kategoriye özel "Varsayılan İndirim Oranı" tanımlayabilirsiniz.
- **Gelişmiş Kategori Filtreleme:** İster tekil ister tüm ağaç alt dallarını seçerek sadece belirlediğiniz kategorileri güncelleyebilirsiniz.

### 💰 Fiyatlandırma ve Kâr Marjı Motoru
- **Kademeli Fiyat Segmentleri:** Farklı fiyat aralıklarına farklı kâr marjları ekleyin (Örn: 0 - 100 TL arası %50, 100 - 500 TL arası %30).
- **Zarar Önleme (Minimum Kâr Limiti):** Ürünlerin hiçbir koşulda belirlenen taban kârın altında satılmasını engelleyen güvenlik sınırı.
- **Esnek Baz Fiyat Kaynağı:** Hesaplamayı "Alış Fiyatı", "Satış Fiyatı" veya "Piyasa Fiyatı" üzerinden yapabilme.
- **Fiyat Yuvarlama:** Psikolojik fiyatlandırma (örn: .90 veya .99 bitişli yuvarlama) seçenekleri.

### 📦 Stok ve Filtreleme Seçenekleri
- **Stok Filtresi:** Belirlediğiniz stok sütununa göre stoku sıfır veya negatif olan ürünleri otomatik olarak çıktıdan eleyebilme.

### 🎨 Modern Arayüz ve Temalar
- **Kitsora (Fox Orange) Teması:** Turuncu-krem tonlarında özel tasarım.
- **Açık (White) Mod:** Yüksek kontrastlı, net ve ferah çalışma ortamı.
- **Koyu (Dark - Beta) Mod:** Gece çalışmaları için göz yormayan karanlık tema.
- **Sistem Teması:** Windows tema tercihinize (Açık/Koyu) otomatik uyum.

### 💾 Çıktı & Kayıt Kolaylığı
- **Bölümlü Dışa Aktarma:** Dosyaları pazaryerlerinin yükleme sınırlarına uygun olarak belirli satır sayılarına (örn: 5.000 satır) otomatik parçalayarak kaydeder (`output_part_1.xlsx`, `output_part_2.xlsx` vb.).
- **Şablon Sistemi:** Hazırladığınız tüm sütun eşleştirmelerini, segmentleri ve kâr oranlarını şablon dosyası olarak kaydedip dilediğiniz an tek tıkla geri yükleyebilirsiniz.
- **Ayrıntılı İşlem Günlükleri (Log):** Yapılan her adımı ve oluşturulan barkod sayılarını `.log` ve `.txt` formatında dışa aktarma imkanı.

---

## 🛠️ Kurulum

1. **Python Kurulumu:**
   Sistemin çalışması için bilgisayarınızda [Python 3.10+](https://www.python.org/downloads/) kurulu olmalıdır. Kurulum sırasında *"Add Python to PATH"* seçeneğini işaretlemeyi unutmayın.

2. **Otomatik Kurulum & Çalıştırma (Önerilen):**
   - Proje dizinindeki **`run.bat`** dosyasına çift tıklayın.
   - Gerekli tüm kütüphaneler (`PySide6`, `openpyxl` vb.) otomatik olarak kontrol edilir, eksikler varsa kurulur ve uygulama doğrudan açılır.

3. **Manuel Kurulum:**
   ```bash
   pip install -r requirements.txt
   python main.py
   ```

---

## 📖 Adım Adım Kullanım Rehberi

1. **Dosya Seç:** Ana sekmeden işlem yapmak istediğiniz Excel (`.xlsx`) dosyasını seçin.
2. **Sütunları Eşleştir:**
   - Stok Kodu, Ürün Adı, Kategori ve Baz Fiyat sütunlarını belirleyin.
   - Varyantlı ürünleriniz varsa *Varyant Desteği* kutusunu işaretleyip Varyant Grup Kodu ve Varyasyon sütunlarını eşleştirin.
   - Barkod üretimi yapacaksanız *Barkod Sütunu* eşleştirmesini yapın ve dilediğiniz *Barkod Ön Eki*ni girin.
3. **Kategorileri & Kâr Oranlarını Belirleyin:**
   - İlgili sekmelerden kategori ağacını inceleyin ve kâr segmentlerinizi tanımlayın.
4. **Ürün Önizleme:**
   - "Ürün Önizleme" sekmesinden fiyatların nasıl şekillendiğini ve mevcut barkod durumlarını kontrol edin.
   - Eğer sadece belirli ürünlere barkod atayacaksanız, tablodan dilediğiniz ürünleri işaretleyin (Sağ tık menüsünden *"Sadece Barkodu Olmayanları Seç"* diyerek hızlıca seçim yapabilirsiniz).
5. **Çıktı & Barkod Ayarları:**
   - "Çıktı İşlemleri / Ayarlar" sekmesinden güncellenmesini istediğiniz fiyat sütunlarını ve *Barkodları Güncelle/Oluştur* seçeneğini işaretleyin.
   - Barkod güncelleme ayarlarından çalışma modunuzu seçin:
     - *Tüm barkodları güncelle* veya *Seçili barkodları güncelle*.
6. **İşlemi Başlat:**
   - "İşlemi Başlat" butonuna tıklayın. İlerleme çubuğunu ve detaylı logları canlı olarak takip edin.
   - İşlem bittiğinde yeni Excel dosyalarınız belirlediğiniz çıktı klasöründe hazır olacaktır!

---

## 📄 Lisans
Bu yazılım [MIT Lisansı](LICENSE) kapsamında lisanslanmıştır.
