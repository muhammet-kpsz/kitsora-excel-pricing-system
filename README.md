# 🦊 Kitsora Excel Fiyatlandırma Sistemi

![Kitsora Logo](assets/icon.png)

**Kitsora**, e-ticaret satıcıları ve işletmeler için geliştirilmiş, güçlü ve esnek bir Excel fiyat yönetim sistemidir. Ürünlerinizi, maaliyetlerinizi, kar marjlarınızı ve varyantlarınızı tek bir noktadan yöneterek saniyeler içinde binlerce ürünü güncelleyebilirsiniz.

---

## 🚀 Özellikler

### 🎯 Akıllı Eşleştirme & İçe Aktarma

- Excel dosyalarınızı otomatik algılar.
- Akıllı sütun eşleştirme ile (Stok Kodu, Ürün Adı, Kategori, Fiyatlar vb.) hızlı kurulum.
- 100.000+ satırlık büyük dosyaları yüksek performansla okur.

### 🌳 Gelişmiş Kategori Yönetimi

- **Kategori Ağacı:** Excel'den kategorileri çeker ve görsel bir ağaç yapısında sunar.
- **İndirim Yönetimi:** Her kategoriye özel "Varsayılan İndirim Oranı" tanımlayabilirsiniz.
- **Alt Kategori Filtreleme:** Sadece seçtiğiniz alt kategorilerdeki ürünleri güncelleyebilirsiniz.

### 💰 Kâr Marjı & Fiyatlandırma Motoru

- **Fiyat Segmentleri:** Farklı fiyat aralıklarına farklı kâr marjları ekleyin (Örn: 0-100 TL arası %50, 100-500 TL arası %30).
- **Global Kâr Limiti:** Zarar etmenizi önleyen "Minimum Kâr" koruması.
- **Baz Fiyat Seçimi:** İster "Alış Fiyatı", ister "Piyasa Fiyatı" üzerinden hesaplama yapın.

### 🎨 Görsel Özelleştirme ve Diğer Özellikler

- **Barkod Üretici:** Ürünlere benzersiz ön ekli rastgele barkodlar (örn: HFGYM33136552) üretebilir ve güvenle yazdırabilirsiniz.
- **Kitsora Teması:** Turuncu-krem tonlarında özel tasarlanmış modern arayüz.
- **Açık/Koyu Mod:** Göz yormayan tema seçenekleri.
- **Varyant Desteği:** Varyantlı ürünleri gruplayarak veya tekil olarak yönetme.

### 💾 Çıktı & Kayıt

- **Otomatik Bölümleme:** Çıktı dosyalarını belirli satır sayılarına (örn. 5000) bölerek kaydedin.
- **Şablonlar:** Sık kullandığınız ayarları şablon olarak kaydedin ve dilediğiniz zaman geri yükleyin.

---

## 🛠️ Kurulum

1. **Python Kurulumu:**
   Sistemin çalışması için bilgisayarınızda [Python 3.10+](https://www.python.org/downloads/) yüklü olmalıdır.

2. **Gereksinimleri Yükle:**
   `run.bat` dosyasını çalıştırdığınızda gerekli kütüphaneler otomatik olarak yüklenecektir.
   Manuel kurulum için:

   ```bash
   pip install -r requirements.txt
   ```

3. **Çalıştırma:**
   - **Windows:** `run.bat` dosyasına çift tıklayın.

---

## 📖 Nasıl Kullanılır?

1. **Dosya Seç:** Ana ekranda güncellemek istediğiniz Excel dosyasını seçin.
2. **Sütunları Eşleştir:** Programın veriyi tanıması için sütun başlıklarını seçin.
3. **Kategorileri Ayarla:** Kategori sekmesinden çalışmak istediğiniz ürün gruplarını seçin.
4. **Kâr Ekle:** Fiyat segmentlerine göre kâr oranlarınızı girin.
5. **Önizle:** "Ürün Önizleme" sekmesinden fiyatların nasıl değiştiğini kontrol edin.
6. **Barkod Oluşturma (Opsiyonel):** "Güncellenecek Sütunlar" kısmından barkod oluşturmayı aktifleştirebilir ve ön ekinizi belirleyebilirsiniz.
7. **Dışa Aktar:** Sonucu yeni bir Excel dosyası olarak kaydedin.

---
