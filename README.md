# PRG - Kurumsal Yönetim Sistemi

[![Python](https://img.shields.io/badge/Python-3.13-blue.svg)](https://www.python.org/)
[![PyQt5](https://img.shields.io/badge/PyQt5-5.15+-green.svg)](https://www.riverbankcomputing.com/software/pyqt/)
[![License](https://img.shields.io/badge/license-Private-red.svg)]()

**PRG**, PyQt5 ile geliştirilmiş kapsamlı bir kurumsal yönetim sistemidir. Stok yönetimi, sözleşmeler, sevkiyat, finansal işlemler ve daha fazlası için tasarlanmıştır.

## 📋 İçindekiler

- [Genel Bakış](#genel-bakış)
- [Özellikler](#özellikler)
- [Mimari](#mimari)
- [Modüller](#modüller)
- [Kurulum](#kurulum)
- [Yapılandırma](#yapılandırma)
- [Kullanım](#kullanım)
- [Geliştirme](#geliştirme)
- [Son Güncellemeler](#son-güncellemeler)
- [Teknoloji Yığını](#teknoloji-yığını)

## 🌟 Genel Bakış

PRG, aşağıdaki sistemlerle entegre çalışan modüler bir kurumsal yönetim uygulamasıdır:
- **Google Sheets** (Service Account ile) - Veri depolama ve senkronizasyon
- **Microsoft SQL Server** (Mikro ERP) - Mali ve stok verileri
- **Email servisleri** - Otomatik bildirimler
- **WhatsApp** - Müşteri iletişimi

**İstatistikler:**
- **22 Python dosyası**
- **22.456+ satır kod**
- **12 fonksiyonel modül**
- **Modern PyQt5 arayüzü (koyu/açık temalar)**

## ✨ Özellikler

### Temel Özellikler
- 🔐 **Merkezi Yapılandırma** - Service Account tabanlı kimlik doğrulama
- 🎨 **Modern Arayüz** - Temiz, duyarlı PyQt5 arayüzü
- 💾 **Global Veri Cache** - Önbellekleme ile verimli veri yönetimi
- 📊 **Gerçek Zamanlı Senkronizasyon** - Google Sheets ile çift yönlü senkronizasyon
- 🔄 **Lazy Loading** - İhtiyaç anında veri yükleme ile optimize performans
- 🎯 **Focus Border Free** - Temiz tablo seçimi ile gelişmiş kullanıcı deneyimi
- 📱 **Çoklu Platform Desteği** - Windows için EXE paketleme ile optimize edilmiş

### İş Özellikleri
- 📦 **Stok Yönetimi** - Komple stok takibi ve yönetimi
- 📝 **Sözleşme Yönetimi** - Tam sözleşme yaşam döngüsü yönetimi
- 🚚 **Sevkiyat İşlemleri** - Kapsamlı sevkiyat ve lojistik
- 💰 **Mali Takip** - Kasa, virman ve POS işlemleri
- ⚠️ **Risk Yönetimi** - Müşteri risk analizi ve izleme
- 🔐 **SSH Yönetimi** - Güvenli kabuk erişimi ve yönetimi
- 💳 **Ödeme İşlemleri** - Sanal POS ve ödeme takibi
- 📄 **Doküman Yönetimi** - İrsaliyeler ve faturalar

## 🏗️ Mimari

### Temel Bileşenler

#### `run.py` - Uygulama Giriş Noktası
PRG uygulamasının giriş noktası. İşlevler:
- Python yolu yapılandırması
- Modül başlatma
- Hata yönetimi ve teşhis
- Service Account kurulum doğrulaması

#### `main.py` - Ana Uygulama Mantığı
Ana uygulama penceresi ve temel mantık:
- **GlobalDataCache** - Merkezi veri önbellekleme sistemi
- **PRGMainWindow** - Sekmeli arayüz ile ana pencere
- Modül entegrasyonu ve yaşam döngüsü yönetimi
- Global veri yenileme mekanizması

#### `core_architecture.py` - Mimari Temel
Modern mimari desenleri:
- **EventType & ModuleType** - Olay güdümlü mimari
- **Theme** - Arayüz tema sistemi
- **EventBus** - Modüller arası iletişim
- **ModuleRegistry** - Dinamik modül yükleme

#### `ui_components.py` - UI Bileşenleri
Yeniden kullanılabilir arayüz bileşenleri ve widget'lar

#### `embedded_resources.py` - Kaynak Yönetimi
Uygulama ikonları ve gömülü kaynaklar

## 📦 Modüller

### 1. **Stok Modülü** (`stok_module.py`)
**Envanter ve Stok Yönetimi**

Kapsamlı stok yönetim sistemi:
- Gerçek zamanlı stok seviyeleri (DEPO, EXCLUSIVE, SUBE)
- Alışveriş sepeti (Sepet) yönetimi
- Gelişmiş filtreleme ve arama
- Mikro ERP verileri için SQL Server entegrasyonu
- KDV ve marj ile fiyat hesaplamaları
- Excel içe/dışa aktarma
- Hızlı işlemler için sağ tık menüsü

**Ana Özellikler:**
- Çoklu depo desteği
- Otomatik fiyat hesaplamaları
- Gerçek zamanlı stok güncellemeleri
- Bulanık eşleştirme ile akıllı arama
- Düzenlenebilir alışveriş sepeti
- Temiz kullanıcı deneyimi için focus border kaldırıldı

---

### 2. **Sevkiyat Modülü** (`sevkiyat_module.py`)
**Sevkiyat ve Lojistik Yönetimi**

Komple sevkiyat operasyonları yönetimi:
- Otomatik tamamlama ile müşteri arama
- Çoklu sekme sevkiyat verileri (Sevkiyat, Bekleyenler, Araç, Malzeme)
- Bildirimler için WhatsApp entegrasyonu
- Email bildirimleri
- Tüm sekmeler için Excel dışa aktarma
- Risk analizi entegrasyonu
- Mikro ERP entegrasyonu

**Ana Özellikler:**
- Bulanık müşteri adı eşleştirme
- Sözleşme ürün sorgulama
- Araç ve malzeme takibi
- Otomatik email/WhatsApp mesajlaşma
- Çoklu görünüm veri filtreleme
- Özel tarih aralığı filtreleme
- Müşteri listesinden focus border kaldırıldı

---

### 3. **Sözleşme Modülü** (`sozlesme_module.py`)
**Sözleşme Yönetimi**

Gelişmiş sözleşme yaşam döngüsü yönetimi:
- Sözleşme detaylarını görüntüleme
- Ürün kalem yönetimi
- Müşteri ve sipariş bilgileri
- Mikro ERP entegrasyonu (Cari, Stok, Sipariş)
- IPT durum takibi
- Header bilgi yönetimi
- Çoklu tablo veri görünümü

**Ana Özellikler:**
- Sözleşme arama ve filtreleme
- Müşteri seçim diyalogu
- Ürün tablosu düzenleme
- SAP/ERP aktarım işlemleri
- Stok kartı oluşturma
- Sipariş transferi
- 3 tablodan focus border kaldırıldı

---

### 4. **Risk Modülü** (`risk_module.py`)
**Müşteri Risk Analizi**

Müşteri kredisi ve risk yönetimi:
- Risk seviyesi izleme
- Kredi limiti takibi
- Ödeme geçmişi analizi
- Mikro ERP veri entegrasyonu
- Excel dışa aktarma
- Otomatik risk güncellemeleri

**Ana Özellikler:**
- Gerçek zamanlı risk hesaplamaları
- Renkli risk göstergeleri
- Eşik tabanlı uyarılar
- Geçmiş risk takibi
- Temiz tablolar için focus border kaldırıldı

---

### 5. **OKC Modülü** (`okc_module.py`)
**OKC YazarKasa Yönetimi**

Yazar kasa ve ödeme yönetimi:
- Fatura takibi
- Ödeme tutarı filtreleme
- Tarih formatlama (00:00 saat gösterimi kaldırıldı)
- Excel dışa aktarma
- Mikro ERP entegrasyonu
- Hızlı navigasyon

**Ana Özellikler:**
- Tutar bazlı filtreleme (1000 TL çarpanı)
- Fatura tarihi yönetimi
- Ödeme takibi
- Renkli durum göstergeleri
- Temiz tarih gösterimi (GG.AA.YYYY)

---

### 6. **SSH Modülü** (`ssh_module.py`)
**Güvenli Kabuk Yönetimi**

SSH bağlantı ve yönetim sistemi:
- Bağlantı yönetimi
- Farklı SSH veri görünümleri için iki tablolu arayüz
- Durum izleme
- Hızlı işlemler
- Yazdırma desteği

**Ana Özellikler:**
- Çoklu tablo SSH veri gösterimi
- Bağlantı durumu takibi
- Yazdırma işlevi
- 2 tablodan focus border kaldırıldı
- Gerçek zamanlı güncellemeler

---

### 7. **Kasa Modülü** (`kasa_module.py`)
**Kasa İşlemleri**

Mali işlem yönetimi:
- Aylık kasa verileri
- Yıl/ay filtreleme
- İşlem kategorilendirme
- Excel dışa aktarma
- Bakiye hesaplamaları

**Ana Özellikler:**
- Güncel tarih varsayılanı ile aylık görünüm
- Renkli işlem tipleri
- Bakiye takibi
- Hızlı navigasyon
- Dışa aktarma yetenekleri

---

### 8. **Sanalpos Modülü** (`sanalpos_module.py`)
**Sanal POS Yönetimi**

Online ödeme işleme ve takip:
- POS işlem izleme
- Ödeme durumu takibi
- Tarih bazlı filtreleme
- Excel dışa aktarma
- Kasa verileri ile entegrasyon

**Ana Özellikler:**
- Gerçek zamanlı POS verileri
- İşlem geçmişi
- Durum göstergeleri
- QApplication import düzeltmesi uygulandı
- Dışa aktarma işlevi

---

### 9. **İrsaliye Modülü** (`irsaliye_module.py`)
**İrsaliye Yönetimi**

Sevkiyat dokümanı yönetimi:
- İrsaliye oluşturma ve takip
- Çoklu sekme arayüzü
- Doküman dışa aktarma
- Müşteri atama
- Tarih takibi

**Ana Özellikler:**
- Sekme tabanlı organizasyon
- Doküman arama
- Excel'e aktarma
- Kopyalama fonksiyonu ile sağ tık menüsü
- Focus border kaldırıldı
- Kalın yazı tipi stili

---

### 10. **Fiyat Modülü** (`fiyat_module.py`)
**Fiyat ve Etiket Yönetimi**

Ürün fiyatlandırma ve etiketleme:
- SAP kodu oluşturma
- Fiyat listesi yönetimi
- Stok veri entegrasyonu
- Etiket yazdırma hazırlığı
- Excel dışa/içe aktarma

**Ana Özellikler:**
- Otomatik SAP kodu oluşturma
- Çoklu kaynak veri entegrasyonu (DEPO, EXC, SUBE)
- Fiyat hesaplama
- Toplu işleme
- Performans için threading

---

### 11. **Virman Modülü** (`virman_module.py`)
**Virman Yönetimi**

Hesaplar arası transfer işlemleri:
- Hesap transferi takibi
- Aylık veri görünümü
- Bakiye doğrulama
- SQL Server entegrasyonu
- İşlem geçmişi

**Ana Özellikler:**
- Ay bazlı filtreleme
- Transfer doğrulama
- Bakiye kontrolü
- İşlem kayıtları
- Gerçek zamanlı güncellemeler

---

### 12. **Ayar Modülü** (`ayar_module.py`)
**Ayarlar ve Yapılandırma**

Sistem yapılandırma yönetimi:
- Çoklu sekme ayarlar (Ayar, Mail, NoRisk)
- Google Sheets entegrasyonu
- Yapılandırma düzenleme
- Ayar kalıcılığı
- Lazy loading optimizasyonu

**Ana Özellikler:**
- Sekme tabanlı organizasyon
- Doğrudan Google Sheets düzenleme
- Yapılandırma doğrulama
- Kaydet/Yeniden yükle işlevleri
- Gerçek zamanlı güncellemeler

## 🚀 Kurulum

### Gereksinimler

```bash
# Python 3.13+
python --version

# Gerekli paketler
pip install -r requirements.txt
```

### Gerekli Bağımlılıklar

```
PyQt5>=5.15.0
pandas>=2.0.0
numpy>=1.24.0
requests>=2.31.0
gspread>=5.0.0
google-auth>=2.0.0
openpyxl>=3.1.0
pyodbc>=4.0.0
python-dotenv>=1.0.0
fuzzywuzzy>=0.18.0
python-levenshtein>=0.21.0
pyperclip>=1.8.0
cryptography>=41.0.0
```

### Service Account Kurulumu

1. Google Cloud projesi oluşturun
2. Google Sheets API'yi etkinleştirin
3. Service Account oluşturun
4. `service_account.json` dosyasını indirin
5. Üst dizine yerleştirin (`D:/GoogleDrive/PRG/OAuth2/`)
6. Google Sheets'i service account email ile paylaşın

### Yapılandırma

Üst dizinde `central_config.py` oluşturun:

```python
class CentralConfigManager:
    MASTER_SPREADSHEET_ID = "spreadsheet_id_buraya"
    # ... diğer yapılandırmalar
```

## 💻 Kullanım

### Uygulamayı Çalıştırma

```bash
# OAuth2 dizininden
cd D:/GoogleDrive/PRG/OAuth2
python PRG/run.py
```

### Çalıştırılabilir Dosya Oluşturma

```bash
# PyInstaller kullanarak
pyinstaller PRG_onefile.spec --clean --noconfirm
```

Çalıştırılabilir dosya `dist/PRG.exe` dizininde oluşturulacaktır (~76MB).

## 🛠️ Geliştirme

### Proje Yapısı

```
PRG/
├── run.py                  # Giriş noktası
├── main.py                 # Ana uygulama
├── core_architecture.py    # Mimari desenler
├── ui_components.py        # UI bileşenleri
├── embedded_resources.py   # Kaynaklar
├── ayar_module.py          # Ayarlar
├── stok_module.py          # Stok
├── sevkiyat_module.py      # Sevkiyat
├── sozlesme_module.py      # Sözleşmeler
├── risk_module.py          # Risk yönetimi
├── okc_module.py           # Yazar kasa
├── ssh_module.py           # SSH yönetimi
├── kasa_module.py          # Kasa işlemleri
├── sanalpos_module.py      # Sanal POS
├── irsaliye_module.py      # İrsaliyeler
├── fiyat_module.py         # Fiyatlandırma
├── virman_module.py        # Virmanlar
├── icon.ico                # Uygulama ikonu
└── icon.jpg                # İkon kaynağı
```

### Kod Stili

- **PEP 8** uyumluluğu
- Uygun yerlerde **type hints**
- Tüm modüller ve sınıflar için **docstrings**
- Yapılandırma değerleri için **sabitler**
- Stylesheet sabitleri ile **merkezi stillendirme**

### Mimari Desenler

- **Lazy Loading** - Veriler sadece gerektiğinde yüklenir
- **Global Cache** - Modüller arası paylaşılan veri önbelleği
- **Event Bus** - Modüller arası iletişim
- **Module Registry** - Dinamik modül yükleme
- **Service Account** - Merkezi kimlik doğrulama

## 🔄 Son Güncellemeler

### UI/UX İyileştirmeleri
- ✅ **Focus Border Kaldırma** - Tüm modüllerde temiz tablo seçimi
  - stok_module.py - Tablo widget'ları
  - sevkiyat_module.py - Müşteri listesi
  - sozlesme_module.py - 3 tablo (products_table, dialog tablosu, ana tablo)
  - risk_module.py - Risk tablosu
  - okc_module.py - OKC tablosu
  - ssh_module.py - 2 SSH tablosu
  - irsaliye_module.py - Doküman tabloları
  - CSS: `QTableWidget::item:focus { outline: none; border: none; }`
  - Policy: `setFocusPolicy(Qt.NoFocus)`

### Hata Düzeltmeleri
- ✅ **Tarih Format Düzeltmesi** - okc_module.py
  - `strftime('%d.%m.%Y %H:%M')` yerine `strftime('%d.%m.%Y')` kullanıldı
  - Tarih görünümlerinden "00:00" kaldırıldı
  - Daha temiz tarih sunumu

- ✅ **Import Düzeltmesi** - sanalpos_module.py
  - QApplication import eklendi
  - Pano işlemlerinde NameError düzeltildi

### Stil İyileştirmeleri
- ✅ **Sabitler Mimarisi** - irsaliye_module.py
  - CONFIG CONSTANTS bölümü eklendi
  - STYLESHEET CONSTANTS bölümü eklendi
  - Kalın yazı tipi uygulaması
  - Kopyalama fonksiyonu ile sağ tık menüsü

## 🔧 Teknoloji Yığını

### Ön Yüz
- **PyQt5** - GUI framework'ü
- **QTableWidget** - Veri gösterimi
- **QTabWidget** - Çoklu görünüm arayüzü
- **Özel Stylesheet'ler** - Modern stillendirme

### Arka Yüz
- **pandas** - Veri manipülasyonu
- **numpy** - Sayısal işlemler
- **requests** - HTTP istekleri
- **pyodbc** - SQL Server bağlantısı

### Entegrasyon
- **gspread** - Google Sheets API
- **google-auth** - Service Account kimlik doğrulama
- **cryptography** - Güvenli veri işleme

### Araçlar
- **PyInstaller** - Çalıştırılabilir paketleme
- **openpyxl** - Excel dosya işleme
- **fuzzywuzzy** - Bulanık string eşleştirme

## 📝 Lisans

Bu özel bir yazılımdır. Tüm hakları saklıdır.

## 👥 Yazar

**İsmail Güneş**

## 🤝 Katkıda Bulunma

Bu özel bir projedir. Katkılar dahili olarak yönetilmektedir.

## 📞 Destek

Dahili destek için geliştirme ekibiyle iletişime geçin.

---

**by İsmail Güneş**

Son Güncelleme: 24 Kasım 2025
