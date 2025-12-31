# PRG

Bu proje, işletmenin ERP süreçlerini, depo yönetimini, sevkiyat planlamasını ve finansal operasyonlarını tek bir çatı altında toplayan kapsamlı, modüler bir masaüstü uygulamasıdır. Python ve PyQt5 kullanılarak geliştirilmiş olup, modern bir yazılım mimarisine (Event-Driven, Command Pattern) sahiptir.

## 🏗️ Mimari Yapı

Proje, sürdürülebilirlik ve genişletilebilirlik için sağlam bir temel üzerine inşa edilmiştir:

- **`core_architecture.py`**: Uygulamanın kalbidir.
  - **EventBus**: Modüller arası gevşek bağlı (decoupled) iletişim sağlar. Olay tabanlı (Event-Driven) bir yapı kurar.
  - **AppState & ThemeManager**: Uygulama durumunu ve tema (Dark/Light) tercihlerini yönetir.
  - **Command Pattern**: İşlemleri (örn. sayfa geçişleri, tema değişimi) nesneleştirerek "Geri Al/Yinele" (Undo/Redo) altyapısı sunar.
  - **ModuleRegistry**: Modüllerin dinamik olarak yüklenmesini ve yönetilmesini sağlar.

## 🧩 Modüller ve İşlevleri

Uygulama, her biri belirli bir iş alanına odaklanan bağımsız modüllerden oluşur:

### ⚙️ Yönetim ve Ayarlar

- **`ayar_module.py`**:
  - Uygulamanın tüm konfigürasyonunun (API anahtarları, Veritabanı bağlantıları, Sabitler) yönetildiği merkezdir.
  - Google Sheets (`PRGsheet`) ile senkronize çalışarak ayarları buluttan çeker ve yerel önbellekte saklar.
  - E-posta sunucu ayarları ve Risk parametreleri buradan yapılandırılır.

### 📦 Stok ve Ürün Yönetimi

- **`stok_module.py`**:
  - SQL Server (Mikro ERP) ve Google Sheets verilerini birleştirerek gerçek zamanlı stok analizi yapar.
  - **Özellikler:** Stok kartı oluşturma, pasif stok yönetimi (3A -> 2A dönüşümü), sepet oluşturma ve WhatsApp üzerinden satış ekibiyle paylaşma.
  - Kritik stok seviyelerini, bekleyen siparişleri ve depo durumunu tek ekranda sunar.
- **`fiyat_module.py`**:
  - Ürünlerin farklı fiyat listelerindeki (Toptan, Perakende, Kampanyalı) durumlarını analiz eder ve karşılaştırır.

### 🚚 Sevkiyat ve Lojistik

- **`sevkiyat_module.py`**:
  - Müşteri siparişlerinin sevkiyat planlamasını yapar.
  - **Özellikler:** Müşteri borç/risk kontrolü, araç planlama, "Sevke Hazır" ve "Açık Sipariş" bilgilendirme mailleri gönderme.
  - WhatsApp entegrasyonu ile müşterilere randevu ve bilgilendirme mesajları gönderir.
  - Eksik ürünleri ve tedarik süreçlerini "Bekleyenler" havuzunda yönetir.
- **`irsaliye_module.py`**:
  - Kesilen irsaliyelerin takibi ve ERP sistemiyle entegrasyonu.

### 💰 Finans ve Muhasebe

- **`risk_module.py`**:
  - Müşterilerin finansal risklerini (Açık Çek/Senet, Gecikmiş Bakiye) analiz eder ve sevkiyat onayı için "Kırmızı/Yeşil" ışık yakar.
- **`kasa_module.py`**:
  - Günlük nakit akışı, kasa giriş-çıkış hareketleri.
- **`sanalpos_module.py`**:
  - Sanal POS tahsilatlarının banka kayıtları ile ERP kayıtlarını otomatik olarak karşılaştırır (Mutabakat).
- **`okc_module.py`**:
  - Ödeme Kaydedici Cihaz (Yazar Kasa) verilerinin analizi.
- **`virman_module.py`**:
  - Hesaplar arası para transferleri (Virman) ve EFT işlemlerinin yönetimi.

### 🛠️ Satış Sonrası ve Operasyon

- **`ssh_module.py`** (Satış Sonrası Hizmetler):
  - Müşteri şikayetleri, teknik servis talepleri ve yedek parça süreçlerinin takibi.
- **`sozlesme_module.py`**:
  - Müşteri satış sözleşmelerinin dijital takibi ve yönetimi.

## 📂 Diğer Önemli Dosyalar

- **`main.py` & `run.py`**: Uygulamanın başlatıcı dosyalarıdır. Gerekli kütüphaneleri kontrol eder ve ana pencereyi ayağa kaldırır.
- **`ui_components.py`**: Uygulama genelinde kullanılan yeniden kullanılabilir arayüz bileşenlerini (Butonlar, Tablolar, Kartlar) içerir.
- **`embedded_resources.py`**: Uygulamanın ikon, logo gibi görsel kaynaklarını ve statik verilerini barındırır.

## Yazar

<div data-spark-custom-html="true">
    <table cellspacing="0" cellpadding="0" border="0" style="border-collapse: collapse; border: none; font-family: sans-serif;">
        <tbody>
            <tr>
                <td style="vertical-align: top; border: none; padding: 0 8px 0 0;">
                     <img src="https://res.spikenow.com/c/?id=576ji8df6q7d6eq2&amp;s=48&amp;m=c&amp;_ts=1xc0n1" width="27" height="27" style="border-radius: 50%; display: block;">
                </td>
                <td style="vertical-align: top; border: none; padding: 0;">
                    <div style="line-height: 1.2;"><a href="https://twitter.com/Guneslsmail" style="text-decoration: none !important; color: #0084ff !important; font-size: 13px; font-weight: bold;">İsmail Güneş</a></div>
                    <div style="line-height: 1.2; margin-top: 2px;"><a href="https://www.instagram.com/dogtasbatman/" style="text-decoration: none !important; color: #0084ff !important; font-size: 12px; font-weight: bold;">Güneşler Elektronik<br>Mühendislik Mobilya</a></div>
                </td>
            </tr>
        </tbody>
    </table>
</div>

## **Proje Başlangıç Tarihi:** 15.11.2024
