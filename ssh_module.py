"""
SSH Modülü - Google Sheets SSH verilerini yönetme ve UI
"""

import os
import re
import logging
import sys
from datetime import datetime, timedelta
from typing import List, Dict, Optional, Any
import pandas as pd
import gspread
from dotenv import load_dotenv
import requests

# Üst dizini Python path'e ekle (central_config için)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Central config import
from central_config import CentralConfigManager

# PyQt5 UI bileşenleri
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QLabel, QPushButton,
                            QTableWidget, QTableWidgetItem, QHBoxLayout,
                            QMessageBox, QProgressBar, QTextEdit, QSplitter,
                            QLineEdit, QStatusBar, QHeaderView, QMenu,
                            QScrollArea, QApplication, QMainWindow, QCheckBox,
                            QProgressDialog, QDialog, QGroupBox)
from PyQt5.QtCore import QThread, pyqtSignal, Qt, QTimer, QSortFilterProxyModel, QMarginsF
from PyQt5.QtGui import QFont, QPixmap, QPainter, QColor
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog
from PyQt5.QtGui import QPageLayout, QPageSize

load_dotenv()

# Logging ayarları - KONSOL İÇİN (UTF-8 encoding)
import sys
logging.basicConfig(
    level=logging.INFO,
    format='%(levelname)s: %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# Force console output
logger.setLevel(logging.INFO)
for handler in logger.handlers:
    handler.setLevel(logging.INFO)


def safe_print(msg):
    """Windows console emoji hatası için güvenli print"""
    try:
        print(msg)
    except:
        # Emoji karakterlerini kaldır
        import re
        clean_msg = re.sub(r'[^\x00-\x7F]+', '', msg)
        print(clean_msg)


class PartStatusChecker:
    """BekleyenFast.py kodunu kullanarak parça durumu kontrol sınıfı"""

    def __init__(self):
        self.token = None
        self.sheets_manager = self._init_sheets()
        self.config_error = None
        self._load_config()

    def _init_sheets(self):
        """Google Sheets API bağlantısını başlatır - Service Account"""
        try:
            config_manager = CentralConfigManager()
            return config_manager.gc
        except Exception:
            return None

    def _load_config(self):
        """Google Sheets Ayar sayfasından API konfigürasyonlarını yükler"""
        try:
            if not self.sheets_manager:
                self.config_error = "Google Sheets bağlantısı kurulamadı"
                self.base_url = ''
                self.endpoint = ''
                self.customer_no = ''
                self.auth_data = {}
                return

            sheet = self.sheets_manager.open("PRGsheet").worksheet('Ayar')

            # Ayar sayfası 4 sütunlu: [App Name, Key, Description, Value]
            # row[1] = Key (base_url, userName, vb.)
            # row[3] = Value (gerçek değer)
            all_rows = sheet.get_all_values()
            config = {}
            for row in all_rows[1:]:  # İlk satır header, atla
                if len(row) >= 4 and row[1]:  # Key (row[1]) ve Value (row[3]) olmalı
                    config[row[1]] = row[3]
                elif len(row) >= 2 and row[1]:  # Eski format uyumluluğu için
                    config[row[1]] = row[1] if len(row) < 3 else ''

            self.base_url = config.get('base_url', '')
            self.endpoint = config.get('bekleyenler', '')
            self.customer_no = config.get('CustomerNo', '')

            self.auth_data = {
                "userName": config.get('userName', ''),
                "password": config.get('password', ''),
                "clientId": config.get('clientId', ''),
                "clientSecret": config.get('clientSecret', ''),
                "applicationCode": config.get('applicationCode', '')
            }

            # Config doğrulaması
            missing_fields = []
            if not self.base_url:
                missing_fields.append('base_url')
            if not self.endpoint:
                missing_fields.append('bekleyenler')
            if not self.customer_no:
                missing_fields.append('CustomerNo')
            if not self.auth_data.get('userName'):
                missing_fields.append('userName')
            if not self.auth_data.get('password'):
                missing_fields.append('password')
            if not self.auth_data.get('clientId'):
                missing_fields.append('clientId')
            if not self.auth_data.get('clientSecret'):
                missing_fields.append('clientSecret')
            if not self.auth_data.get('applicationCode'):
                missing_fields.append('applicationCode')

            if missing_fields:
                self.config_error = f"Eksik config alanları: {', '.join(missing_fields)}"

        except Exception as e:
            self.config_error = f"Config yükleme hatası: {str(e)}"
            self.base_url = ''
            self.endpoint = ''
            self.customer_no = ''
            self.auth_data = {}

    def _get_token(self):
        """API access token alma"""
        try:
            if not self.base_url:
                return False

            token_url = f"{self.base_url}/Authorization/GetAccessToken"
            response = requests.post(
                token_url,
                json=self.auth_data,
                timeout=10
            )

            if response.status_code == 200:
                data = response.json()
                if data.get('isSuccess') and 'data' in data:
                    self.token = data['data']['accessToken']
                    return True
            return False

        except Exception:
            return False

    def check_part_status(self, siparis_no, montaj_tarihi):
        """Tek parça için durum kontrolü"""
        # Config hatası varsa
        if self.config_error:
            return {"error": f"Config hatası: {self.config_error}", "siparis_no": siparis_no}

        # Token al
        if not self.token and not self._get_token():
            return {"error": f"Token alınamadı", "siparis_no": siparis_no}

        try:
            # Sabit başlangıç tarihi ve dinamik bugün tarihi (DD.MM.YYYY formatında)
            start_date = "01.01.2023"
            end_date = datetime.now().strftime("%d.%m.%Y")

            # API payload
            payload = {
                "orderId": siparis_no,
                "CustomerNo": self.customer_no,
                "RegistrationDateStart": start_date,
                "RegistrationDateEnd": end_date,
                "referenceDocumentNo": "",
                "SalesDocumentType": ""
            }

            api_url = f"{self.base_url}{self.endpoint}"

            # API çağrısı
            response = requests.post(
                api_url,
                json=payload,
                headers={
                    'Authorization': f'Bearer {self.token}',
                    'Content-Type': 'application/json'
                },
                timeout=30
            )

            if response.status_code == 200:
                result = response.json()

                if result.get('isSuccess') and isinstance(result.get('data'), list):
                    data = result['data']
                    filtered_data = data

                    return {
                        "success": True,
                        "siparis_no": siparis_no,
                        "data_count": len(filtered_data),
                        "data": filtered_data[:10]
                    }
                elif result.get('isSuccess') == False:
                    message = result.get('message', 'Veri bulunamadı')

                    if result.get('data') is not None:
                        data = result.get('data', [])
                        if isinstance(data, list) and len(data) > 0:
                            filtered_data = data

                            return {
                                "success": True,
                                "siparis_no": siparis_no,
                                "data_count": len(filtered_data),
                                "data": filtered_data[:10]
                            }

                    return {
                        "success": False,
                        "siparis_no": siparis_no,
                        "error": f"Veri bulunamadı: {message}",
                        "data_count": 0
                    }
                else:
                    return {
                        "success": False,
                        "siparis_no": siparis_no,
                        "error": "Beklenmeyen API response yapısı",
                        "raw_response": result
                    }

            return {"error": f"HTTP hatası: {response.status_code}", "siparis_no": siparis_no}

        except Exception as e:
            return {"error": f"Hata: {str(e)}", "siparis_no": siparis_no}


class SshDataLoader(QThread):
    """SSH verilerini Google Sheets'den yükleyen thread"""

    data_loaded = pyqtSignal(list)  # Yüklenen veri listesi
    error_occurred = pyqtSignal(str)  # Hata mesajı
    progress_updated = pyqtSignal(int, str)  # Progress (0-100) ve mesaj

    def __init__(self):
        super().__init__()

    def run(self):
        """Thread ana işlevi - SSH verilerini yükle"""
        try:
            # PRGsheet/Ayar sayfasından SPREADSHEET_ID'yi yükle
            self.progress_updated.emit(0, "📊 Yapılandırma okunuyor...")
            config_manager = CentralConfigManager()
            spreadsheet_id = config_manager.MASTER_SPREADSHEET_ID

            # Google Sheets Excel export URL'si
            import requests
            import io
            gsheets_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"

            # Excel dosyasını indir
            self.progress_updated.emit(20, "🔗 Google Sheets'e bağlanıyor...")
            response = requests.get(gsheets_url, timeout=30)
            response.raise_for_status()

            # Pandas ile SSH sayfasını oku
            self.progress_updated.emit(50, "📥 SSH sayfası indiriliyor...")
            excel_data = pd.read_excel(io.BytesIO(response.content), sheet_name='Ssh')

            # DataFrame'i dict listesine çevir
            self.progress_updated.emit(80, "🔄 Veriler işleniyor...")
            ssh_data = excel_data.fillna('').to_dict('records')

            self.progress_updated.emit(100, f"✅ {len(ssh_data)} kayıt yüklendi")
            self.data_loaded.emit(ssh_data)

        except Exception as e:
            error_msg = f"SSH veri yükleme hatası: {str(e)}"
            logger.error(error_msg)
            self.error_occurred.emit(error_msg)


class ContractCheckWorker(QThread):
    """SAP API çağrılarını paralel olarak yapan worker thread"""

    progress_updated = pyqtSignal(int, str)  # Progress (0-100) ve mesaj
    finished_signal = pyqtSignal(list)  # Filtrelenmiş kayıtlar
    error_occurred = pyqtSignal(str)  # Hata mesajı

    def __init__(self, sozleme_module, kayitlar):
        super().__init__()
        self.sozleme_module = sozleme_module
        self.kayitlar = kayitlar
        self.is_cancelled = False

    def run(self):
        """Thread ana işlevi - paralel API çağrıları"""
        try:
            from concurrent.futures import ThreadPoolExecutor, as_completed

            eslesmeyen_kayitlar = []
            total_records = len(self.kayitlar)
            completed_count = 0

            # ThreadPoolExecutor ile 5 paralel çağrı (çok fazla yapmayalım, API rate limit için)
            with ThreadPoolExecutor(max_workers=5) as executor:
                # Tüm kayıtlar için future'lar oluştur
                future_to_kayit = {
                    executor.submit(self._check_contract, kayit): kayit
                    for kayit in self.kayitlar
                }

                # Tamamlananları işle
                for future in as_completed(future_to_kayit):
                    if self.is_cancelled:
                        break

                    kayit = future_to_kayit[future]
                    try:
                        result = future.result()
                        if result:  # 'Bayi Dış Teslimat' değilse
                            eslesmeyen_kayitlar.append(result)
                    except Exception as e:
                        logger.error(f"Kayıt işlenirken hata: {e}")
                        # Hata durumunda da ekle (güvenli taraf)
                        kayit_copy = kayit.copy()
                        kayit_copy.pop('_sip_belgeno', None)
                        eslesmeyen_kayitlar.append(kayit_copy)

                    # Progress güncelle
                    completed_count += 1
                    progress = 50 + int((completed_count / total_records) * 40)
                    sip_belgeno = kayit.get('_sip_belgeno', '')
                    self.progress_updated.emit(progress, f"🔍 Kontrol ediliyor: {sip_belgeno} ({completed_count}/{total_records})")

            self.finished_signal.emit(eslesmeyen_kayitlar)

        except Exception as e:
            self.error_occurred.emit(f"Thread hatası: {str(e)}")

    def _check_contract(self, kayit):
        """Tek bir sözleşme için SHIPPING_COND kontrolü"""
        sip_belgeno = kayit['_sip_belgeno']

        try:
            contract_data = self.sozleme_module.get_all_contract_info(sip_belgeno)

            # SHIPPING_COND kontrolü
            shipping_cond = ""
            if contract_data and hasattr(contract_data, 'ES_CONTRACT_INFO'):
                contract_info = contract_data.ES_CONTRACT_INFO
                if hasattr(contract_info, 'SHIPPING_COND'):
                    shipping_cond = contract_info.SHIPPING_COND

            # 'Bayi Dış Teslimat' değilse kayıt döndür
            if shipping_cond != 'Bayi Dış Teslimat':
                kayit_copy = kayit.copy()
                kayit_copy.pop('_sip_belgeno', None)
                return kayit_copy

            return None  # 'Bayi Dış Teslimat' ise None döndür

        except Exception as e:
            logger.error(f"Sözleşme {sip_belgeno} kontrol hatası: {e}")
            # Hata durumunda kayıt döndür (güvenli taraf)
            kayit_copy = kayit.copy()
            kayit_copy.pop('_sip_belgeno', None)
            return kayit_copy

    def cancel(self):
        """Thread'i iptal et"""
        self.is_cancelled = True


class SshModule(QMainWindow):
    """SSH Modülü Ana Sınıfı"""

    def __init__(self):
        super().__init__()
        self.ssh_data = []  # Ham veri (Google Sheets'den gelen)
        self.base_filtered_data = []  # Temel filtrelenmiş veriler ("Çözüldü" + "Sorunsuz Teslimat" hariç)
        self.filtered_data = []  # Arama sonrası filtrelenmiş veriler
        self.data_loader = None
        self.ssh_raporu_calisiyor = False  # SSH.exe çalışıyor mu
        self.mikro_calisiyor = False  # Tamamlanan.exe çalışıyor mu
        self.montaj_yukleniyor = False  # Montaj.exe çalışıyor mu
        self._data_loaded = False  # Lazy loading için flag
        self.contract_worker = None  # Contract check worker thread

        # UI bileşenleri
        self.search_input = None
        self.refresh_btn = None
        self.clear_btn = None
        self.status_btn = None
        self.table = None
        self.progress_bar = None
        self.search_timer = None

        self.setup_ui()
        self.setup_auto_refresh()

    def setup_ui(self):
        """
        Kullanıcı arayüzünü oluşturur ve yapılandırır.
        Sozlesme_module ile aynı tasarım
        """
        # Pencere başlığı
        self.setWindowTitle("SSH Veri Yönetimi")

        # Ekran boyutunu al ve pencere boyutunu ayarla
        screen = QApplication.primaryScreen()
        screen_geometry = screen.availableGeometry()
        width = int(screen_geometry.width() * 0.8)  # Ekran genişliğinin %80'i
        height = int(screen_geometry.height() * 0.85)  # Ekran yüksekliğinin %85'i

        # Pencereyi ekranın merkezine yerleştir
        x = (screen_geometry.width() - width) // 2
        y = (screen_geometry.height() - height) // 2
        self.setGeometry(x, y, width, height)

        # Ana widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Beyaz arka plan ayarla (sozlesme_module gibi)
        central_widget.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                color: #000000;
            }
        """)
        # Force white background
        central_widget.setAutoFillBackground(True)
        palette = central_widget.palette()
        palette.setColor(central_widget.backgroundRole(), QColor("#ffffff"))
        central_widget.setPalette(palette)

        # Layout
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Üst kontrol paneli
        self.create_control_panel(layout)


        # Scroll area
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)

        # SSH verileri tablosu
        self.create_ssh_table(scroll_layout)

        scroll_area.setWidget(scroll_widget)
        layout.addWidget(scroll_area)

        # Status Layout (Label + Progress Bar) - stok_module ile aynı
        status_layout = QHBoxLayout()

        # Status Label
        self.status_label = QLabel("Hazır")
        self.status_label.setStyleSheet("""
            QLabel {
                color: #333333;
                padding: 4px 8px;
                background-color: #f5f5f5;
                border-top: 1px solid #d0d0d0;
                font-size: 14px;
                max-height: 20px;
            }
        """)

        # Progress Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #d0d0d0;
                border-radius: 3px;
                background-color: #ffffff;
                color: #333333;
                text-align: center;
                font-weight: bold;
                min-height: 17px;
                max-height: 17px;
                font-size: 17px;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 0, stop: 0 #4CAF50, stop: 1 #45a049);
                border-radius: 3px;
            }
        """)

        status_layout.addWidget(self.status_label, 3)
        status_layout.addWidget(self.progress_bar, 1)
        status_layout.setContentsMargins(0, 0, 0, 0)

        status_widget = QWidget()
        status_widget.setLayout(status_layout)
        status_widget.setStyleSheet("background-color: #f5f5f5; border-top: 1px solid #d0d0d0;")

        layout.addWidget(status_widget)

    def showEvent(self, event):
        """Widget ilk gösterildiğinde veri yükle (lazy loading)"""
        super().showEvent(event)
        if not self._data_loaded:
            self._data_loaded = True
            # UI render olduktan sonra veri yükle
            QTimer.singleShot(100, self.refresh_data)

    def setup_auto_refresh(self):
        """77 dakikada bir otomatik veri yenileme"""
        self.auto_refresh_timer = QTimer()
        self.auto_refresh_timer.timeout.connect(self.auto_refresh_data)
        # 77 dakika = 77 * 60 * 1000 = 4,620,000 milisaniye
        self.auto_refresh_timer.start(4620000)

    def auto_refresh_data(self):
        """Otomatik veri yenileme fonksiyonu"""
        # Hiçbir işlem çalışmıyorsa otomatik yenileme yap
        if not self.ssh_raporu_calisiyor and not self.mikro_calisiyor and not self.montaj_yukleniyor:
            from PyQt5.QtCore import QDateTime
            current_time = QDateTime.currentDateTime().toString("hh:mm:ss")
            self.status_label.setText(f"🔄 Otomatik veri yenileme başlatıldı ({current_time})")
            QApplication.processEvents()
            self.refresh_data()

    def run_mikro(self):
        """Tamamlanan.exe dosyasını çalıştır"""
        try:
            exe_path = r"D:/GoogleDrive/PRG/EXE/Tamamlanan.exe"
            if not os.path.exists(exe_path):
                self.status_label.setText(f"❌ Tamamlanan.exe bulunamadı: {exe_path}")
                return

            self.status_label.setText("🔄 Tamamlanan.exe çalıştırılıyor...")
            self.mikro_button.setEnabled(False)
            self.mikro_calisiyor = True
            QApplication.processEvents()

            os.startfile(exe_path)

            # Tamamlanan.exe'nin çalışması için bekleme (7 saniye)
            QTimer.singleShot(7000, self.on_mikro_finished)

        except Exception as e:
            error_msg = f"Program çalıştırma hatası: {str(e)}"
            logger.error(error_msg)
            self.status_label.setText(f"❌ {error_msg}")
            self.mikro_button.setEnabled(True)
            self.mikro_calisiyor = False

    def on_mikro_finished(self):
        """Mikro program bittikten sonra"""
        self.mikro_button.setEnabled(True)
        self.mikro_calisiyor = False
        self.status_label.setText("✅ Tamamlanan.exe tamamlandı, Google Sheets güncelleme bekleniyor...")

        # Google Sheets'e kaydedilmesi için ek bekleme (5 saniye)
        QTimer.singleShot(5000, self.delayed_data_refresh)

    def run_montaj_yukle(self):
        """Montaj.exe dosyasını çalıştır"""
        try:
            exe_path = r"D:/GoogleDrive/PRG/EXE/Montaj.exe"
            if not os.path.exists(exe_path):
                self.status_label.setText(f"❌ Montaj.exe bulunamadı: {exe_path}")
                return

            self.status_label.setText("🔄 Montaj.exe çalıştırılıyor...")
            self.montaj_yukle_btn.setEnabled(False)
            self.montaj_yukleniyor = True
            QApplication.processEvents()

            os.startfile(exe_path)

            # Montaj.exe'nin çalışması için bekleme (7 saniye)
            QTimer.singleShot(7000, self.on_montaj_yukle_finished)

        except Exception as e:
            error_msg = f"Program çalıştırma hatası: {str(e)}"
            logger.error(error_msg)
            self.status_label.setText(f"❌ {error_msg}")
            self.montaj_yukle_btn.setEnabled(True)
            self.montaj_yukleniyor = False

    def on_montaj_yukle_finished(self):
        """Montaj yükleme bittikten sonra"""
        self.montaj_yukle_btn.setEnabled(True)
        self.montaj_yukleniyor = False
        self.status_label.setText("✅ Montaj.exe tamamlandı, Google Sheets güncelleme bekleniyor...")

        # Google Sheets'e kaydedilmesi için ek bekleme (5 saniye)
        QTimer.singleShot(5000, self.delayed_data_refresh)

    def create_control_panel(self, layout):
        """Üst kontrol panelini oluştur"""
        control_panel = QWidget()
        control_layout = QHBoxLayout(control_panel)
        control_layout.setContentsMargins(10, 5, 10, 5)

        # Mikro butonu
        self.mikro_button = QPushButton("🔧 Mikro")
        self.mikro_button.setStyleSheet("""            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 15px;
                font-size: 12px;
                font-weight: bold;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
            QPushButton:pressed {
                background-color: #909090;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #888888;
            }
        """)
        self.mikro_button.clicked.connect(self.run_mikro_ssh)
        control_layout.addWidget(self.mikro_button)

        # Montaj Belgesi Yükle butonu
        self.montaj_yukle_btn = QPushButton("📤 Montaj Belgesi Yükle")
        self.montaj_yukle_btn.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 15px;
                font-size: 12px;
                font-weight: bold;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
            QPushButton:pressed {
                background-color: #909090;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #888888;
            }
        """)
        self.montaj_yukle_btn.clicked.connect(self.run_montaj_yukle)
        control_layout.addWidget(self.montaj_yukle_btn)

        # Montaj Raporu butonu
        self.montaj_belgesi_btn = QPushButton("📋 Montaj Raporu")
        self.montaj_belgesi_btn.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 15px;
                font-size: 12px;
                font-weight: bold;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
            QPushButton:pressed {
                background-color: #909090;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #888888;
            }
        """)
        self.montaj_belgesi_btn.clicked.connect(self.show_montaj_belgesi_dialog)
        control_layout.addWidget(self.montaj_belgesi_btn)

        # Arama alanı (dinamik genişleyebilir)
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Müşteri Adı, Ürün Adı veya Yedek Parça Ürün Tanımı ile ara...")
        self.search_input.setMinimumWidth(250)
        self.search_input.setStyleSheet("""
            QLineEdit {
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                padding: 5px;
                font-size: 12px;
            }
            QLineEdit:focus {
                border-color: #3498db;
            }
        """)
        # Arama için timer ekle (sozlesme_module gibi)
        self.search_timer = QTimer()
        self.search_timer.setSingleShot(True)
        self.search_timer.timeout.connect(self.schedule_filter)
        self.search_input.textChanged.connect(self.on_search_text_changed)
        control_layout.addWidget(self.search_input, 1)  # Dinamik genişleme için stretch factor 1

        # Temizle butonu
        self.clear_btn = QPushButton("🗑️ Temizle")
        self.clear_btn.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 15px;
                font-size: 12px;
                font-weight: bold;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
            QPushButton:pressed {
                background-color: #909090;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #888888;
            }
        """)
        self.clear_btn.clicked.connect(self.clear_search)
        control_layout.addWidget(self.clear_btn)

        # Parça Durumu butonu
        self.status_btn = QPushButton("📊 Parça Durumu")
        self.status_btn.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 15px;
                font-size: 12px;
                font-weight: bold;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
            QPushButton:pressed {
                background-color: #909090;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #888888;
            }
        """)
        self.status_btn.clicked.connect(self.check_part_status)
        control_layout.addWidget(self.status_btn)

        # SSH Raporu butonu
        self.montaj_raporu_btn = QPushButton("📋 SSH Raporu")
        self.montaj_raporu_btn.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 15px;
                font-size: 12px;
                font-weight: bold;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
            QPushButton:pressed {
                background-color: #909090;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #888888;
            }
        """)
        self.montaj_raporu_btn.clicked.connect(self.run_montaj_raporu)
        control_layout.addWidget(self.montaj_raporu_btn)

        # Verileri Yenile butonu
        self.refresh_btn = QPushButton("🔄 Verileri Yenile")
        self.refresh_btn.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 15px;
                font-size: 12px;
                font-weight: bold;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
            QPushButton:pressed {
                background-color: #909090;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #888888;
            }
        """)
        self.refresh_btn.clicked.connect(self.refresh_data)
        control_layout.addWidget(self.refresh_btn)

        # Yazdır butonu (başta inaktif)
        self.print_btn = QPushButton("🖨️ Yazdır")
        self.print_btn.setEnabled(False)
        self.print_btn.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 15px;
                font-size: 12px;
                font-weight: bold;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
            QPushButton:pressed {
                background-color: #909090;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #888888;
            }
        """)
        self.print_btn.clicked.connect(self.print_customer_info)
        control_layout.addWidget(self.print_btn)

        layout.addWidget(control_panel)

    def create_ssh_table(self, layout):
        """SSH verileri tablosunu oluştur"""
        # Tablo widget'ı - "SSH Veri Listesi" başlığı kaldırıldı
        self.table = QTableWidget()
        # Light theme - risk_module.py ile aynı
        self.table.setStyleSheet("""
            QTableWidget {
                font-size: 15px;
                font-weight: bold;
                background-color: #ffffff;
                alternate-background-color: #f5f5f5;
                gridline-color: #d0d0d0;
                border: 1px solid #d0d0d0;
                color: #000000;
            }
            QTableWidget::item {
                padding: 5px;
                border-bottom: 1px solid #e0e0e0;
                color: #000000;
            }
            QTableWidget::item:selected {
                background-color: #b3d9ff;
                color: #000000;
            }
            QTableWidget::item:focus {
                outline: none;
                border: none;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                color: #000000;
                padding: 8px;
                border: 1px solid #d0d0d0;
                font-weight: bold;
                font-size: 15px;
            }
        """)

        # Tablo ayarları
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSortingEnabled(True)
        self.table.setFocusPolicy(Qt.NoFocus)  # Focus border'ı kaldır (risk_module.py gibi)

        # Satır yüksekliği
        self.table.verticalHeader().setDefaultSectionSize(35)
        self.table.verticalHeader().setVisible(False)

        # Context menu
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)

        layout.addWidget(self.table)

    def clear_search(self):
        """Arama alanını temizle"""
        self.search_input.clear()

    # ================== CONTEXT MENU ==================
    def show_context_menu(self, position):
        """Sağ tık menüsü - Sadece hücre kopyalama"""
        item = self.table.itemAt(position)
        if not item:
            return

        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                background-color: #ffffff;
                border: 1px solid #d0d0d0;
                border-radius: 5px;
                padding: 4px;
                color: #000000;
            }
            QMenu::item {
                padding: 6px 12px;
                border-radius: 3px;
            }
            QMenu::item:selected {
                background-color: #e3f2fd;
                color: #000000;
            }
        """)

        copy_action = menu.addAction("Kopyala")

        action = menu.exec_(self.table.viewport().mapToGlobal(position))

        if action == copy_action:
            self.copy_cell(item)

    def copy_cell(self, item: QTableWidgetItem):
        """Tıklanan hücreyi kopyala"""
        if item and item.text():
            QApplication.clipboard().setText(item.text())
            self.status_label.setText("✅ Kopyalandı")
        else:
            self.status_label.setText("⚠️ Boş hücre")

    def check_print_button_state(self):
        """Yazdır butonunun durumunu kontrol et - tüm seçili satırların Parça Durumu 'FATR' içermeli"""
        has_checked = False
        all_fatr = True

        for row in range(self.table.rowCount()):
            checkbox_widget = self.table.cellWidget(row, 0)
            if checkbox_widget:
                checkbox = checkbox_widget.findChild(QCheckBox)
                if checkbox and checkbox.isChecked():
                    has_checked = True

                    # Parça Durumu sütununu kontrol et (index 1)
                    parca_durumu_item = self.table.item(row, 1)
                    if parca_durumu_item:
                        parca_durumu = parca_durumu_item.text().strip().upper()
                        if "FATR" not in parca_durumu:
                            all_fatr = False
                            break
                    else:
                        all_fatr = False
                        break

        # Yazdır butonunu aktif/inaktif yap
        if has_checked and all_fatr:
            self.print_btn.setEnabled(True)
        else:
            self.print_btn.setEnabled(False)

    def print_customer_info(self):
        """Müşteri bilgilerini yazdır - seçili satırların sözleşme numarasını ve müşteri adını kontrol et"""
        # İşaretli satırları bul ve sözleşme numaraları + müşteri adlarını topla
        selected_contracts = []
        selected_customers = []
        selected_rows_data = []

        for row in range(self.table.rowCount()):
            checkbox_widget = self.table.cellWidget(row, 0)
            if checkbox_widget:
                checkbox = checkbox_widget.findChild(QCheckBox)
                if checkbox and checkbox.isChecked():
                    sozlesme_no = None
                    musteri_adi = None
                    row_data = {}

                    # Sözleşme Numarası ve Müşteri Adı sütunlarını bul
                    for col in range(2, self.table.columnCount()):
                        header = self.table.horizontalHeaderItem(col)
                        if header:
                            header_text = header.text()
                            item = self.table.item(row, col)
                            if item:
                                if "Sözleşme Numarası" in header_text:
                                    sozlesme_no = item.text().strip()
                                elif "Müşteri Adı" in header_text:
                                    musteri_adi = item.text().strip()

                                # Tüm satır verisini sakla
                                row_data[header_text] = item.text().strip()

                    if sozlesme_no:
                        if sozlesme_no not in selected_contracts:
                            selected_contracts.append(sozlesme_no)
                        if musteri_adi and musteri_adi not in selected_customers:
                            selected_customers.append(musteri_adi)
                        selected_rows_data.append(row_data)

        if not selected_contracts:
            QMessageBox.warning(self, "Uyarı", "Lütfen en az bir satır seçin.")
            return

        # Sözleşme numaralarının hepsi aynı mı kontrol et
        first_contract = selected_contracts[0]
        first_customer = selected_customers[0] if selected_customers else ""

        if len(selected_contracts) > 1:
            # Farklı sözleşme numaraları var - önce müşteri adını kontrol et
            if len(selected_customers) > 1:
                # Müşteri adları da farklı - devam etme
                QMessageBox.warning(
                    self,
                    "Uyarı",
                    "Sözleşme Numaraları ve Müşteri Adları farklı. İşlem iptal edildi."
                )
                return

            # Müşteri adları aynı ama sözleşme numaraları farklı - kullanıcıya sor
            reply = QMessageBox.question(
                self,
                "Sözleşme Numaraları Farklı",
                "Sözleşme Numaraları aynı değildir. Devam etmek ister misiniz?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )

            if reply == QMessageBox.No:
                return

        # İlk satırın sözleşme numarasını kullan
        contract_id = first_contract

        # Sözleşme bilgilerini al ve SSH Arıza Formu göster
        self.fetch_and_show_ssh_form(contract_id, selected_rows_data)

    def fetch_and_show_ssh_form(self, contract_id, selected_rows_data):
        """Sözleşme bilgilerini API'den al ve SSH Arıza Formu göster"""
        try:
            # Sözleşme numarası validasyonu
            if len(contract_id) != 10 or not contract_id.startswith('15'):
                QMessageBox.warning(
                    self,
                    "Uyarı",
                    "Lütfen doğru Sözleşme Numarası giriniz...\n\nSözleşme numarası 10 karakter olmalı ve '15' ile başlamalıdır."
                )
                return

            # Loading mesajı göster
            self.status_label.setText(f"🔍 Sözleşme {contract_id} sorgulanıyor...")
            self.print_btn.setEnabled(False)
            QApplication.processEvents()

            # Sozleme.py modülünü import et - Static import (PyInstaller uyumlu)
            try:
                from PRG import Sozleme as sozleme_module
            except ImportError:
                try:
                    import Sozleme as sozleme_module
                except ImportError as import_error:
                    QMessageBox.warning(self, "Uyarı", f"Sozleme.py yüklenirken hata: {str(import_error)}")
                    self.print_btn.setEnabled(True)
                    self.status_label.setText("❌ Sozleme.py yüklenemedi")
                    return

            # Sözleşme bilgilerini al
            contract_data = sozleme_module.get_all_contract_info(contract_id)

            if contract_data:
                self.status_label.setText(f"✅ Sözleşme {contract_id} başarıyla alındı")
                # SSH Arıza Formu penceresini göster
                self.show_ssh_print_dialog(contract_data, selected_rows_data)
            else:
                self.status_label.setText(f"❌ Sözleşme {contract_id} bulunamadı")
                QMessageBox.warning(self, "Uyarı", f"Sözleşme {contract_id} bulunamadı veya hata oluştu.")

            self.print_btn.setEnabled(True)

        except Exception as e:
            logger.error(f"SSH Arıza Formu hatası: {str(e)}")
            self.status_label.setText(f"❌ SSH Arıza Formu hatası: {str(e)}")
            QMessageBox.critical(self, "Hata", f"SSH Arıza Formu hatası: {str(e)}")
            self.print_btn.setEnabled(True)

    def show_ssh_print_dialog(self, contract_data, selected_rows_data):
        """SSH Arıza Formu penceresini göster"""
        try:
            dialog = SSHPrintDialog(contract_data, selected_rows_data, self)
            dialog.show()
        except Exception as e:
            logger.error(f"SSH Arıza Formu penceresi hatası: {str(e)}")
            QMessageBox.critical(self, "Hata", f"SSH Arıza Formu penceresi hatası: {str(e)}")

    def fetch_and_show_customer_info(self, contract_id):
        """Sözleşme bilgilerini API'den al ve müşteri bilgilerini göster"""
        try:
            # Sözleşme numarası validasyonu
            if len(contract_id) != 10 or not contract_id.startswith('15'):
                QMessageBox.warning(
                    self,
                    "Uyarı",
                    "Lütfen doğru Sözleşme Numarası giriniz...\n\nSözleşme numarası 10 karakter olmalı ve '15' ile başlamalıdır."
                )
                return

            # Loading mesajı göster
            self.status_label.setText(f"🔍 Sözleşme {contract_id} sorgulanıyor...")
            self.print_btn.setEnabled(False)
            QApplication.processEvents()

            # Sozleme.py modülünü import et - Static import (PyInstaller uyumlu)
            try:
                from PRG import Sozleme as sozleme_module
            except ImportError:
                try:
                    import Sozleme as sozleme_module
                except ImportError as import_error:
                    QMessageBox.warning(self, "Uyarı", f"Sozleme.py yüklenirken hata: {str(import_error)}")
                    self.print_btn.setEnabled(True)
                    self.status_label.setText("❌ Sozleme.py yüklenemedi")
                    return

            # Sözleşme bilgilerini al
            contract_data = sozleme_module.get_all_contract_info(contract_id)

            if contract_data:
                self.status_label.setText(f"✅ Sözleşme {contract_id} başarıyla alındı")
                # Müşteri bilgileri penceresini göster
                self.show_customer_info_window(contract_data, contract_id)
            else:
                self.status_label.setText(f"❌ Sözleşme {contract_id} bulunamadı")
                QMessageBox.warning(self, "Uyarı", f"Sözleşme {contract_id} bulunamadı veya hata oluştu.")

            self.print_btn.setEnabled(True)

        except Exception as e:
            logger.error(f"Sözleşme sorgulama hatası: {str(e)}")
            self.status_label.setText(f"❌ Sözleşme sorgulama hatası: {str(e)}")
            QMessageBox.critical(self, "Hata", f"Sözleşme sorgulama hatası: {str(e)}")
            self.print_btn.setEnabled(True)

    def schedule_filter(self):
        """Filtreleme işlemini zamanlı olarak başlat (sozlesme_module gibi)"""
        self.filter_data()

    def check_part_status(self):
        """İşaretli parçaların durumunu kontrol et"""
        # İşaretli satırları bul
        selected_parts = []
        checked_count = 0

        for row in range(self.table.rowCount()):
            checkbox_widget = self.table.cellWidget(row, 0)
            if checkbox_widget:
                checkbox = checkbox_widget.findChild(QCheckBox)
                if checkbox and checkbox.isChecked():
                    checked_count += 1
                    # Yedek Parça Sipariş No ve Montaj Belgesi Tarihi sütunlarını bul
                    siparis_no = None
                    montaj_tarihi = None

                    for col in range(2, self.table.columnCount()):  # 2'den başla (Seç ve Parça Durumu atla)
                        header = self.table.horizontalHeaderItem(col)
                        if header:
                            header_text = header.text()
                            item = self.table.item(row, col)
                            if item:
                                if "Yedek Parça Sipariş No" in header_text:
                                    siparis_no = item.text().strip()
                                elif "Montaj Belgesi Tarihi" in header_text:
                                    montaj_tarihi = item.text().strip()

                    if siparis_no and montaj_tarihi:
                        selected_parts.append({
                            'siparis_no': siparis_no,
                            'montaj_tarihi': montaj_tarihi
                        })

        if not selected_parts:
            if checked_count == 0:
                QMessageBox.information(self, "Bilgi", "Lütfen en az bir satır seçin.")
            else:
                QMessageBox.warning(
                    self,
                    "Uyarı",
                    f"{checked_count} satır seçili ama geçerli sipariş bilgisi bulunamadı.\n\n"
                    "Lütfen seçili satırlarda 'Yedek Parça Sipariş No' ve 'Montaj Belgesi Tarihi' sütunlarının dolu olduğundan emin olun."
                )
            return

        # Parça durum kontrolü başlat
        self.start_part_status_check(selected_parts)

    def start_part_status_check(self, selected_parts):
        """Seçili parçalar için durum kontrolünü başlat"""
        if not selected_parts:
            return

        # PartStatusChecker instance oluştur
        checker = PartStatusChecker()

        # Config hatası kontrolü
        if checker.config_error:
            error_msg = f"Config hatası: {checker.config_error}\n\nLütfen PRGsheet/Ayar sayfasını kontrol edin."
            QMessageBox.critical(self, "Config Hatası", error_msg)
            return

        # Progress dialog oluştur
        progress = QProgressDialog("Parça durumları kontrol ediliyor...", "İptal", 0, len(selected_parts), self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        progress.show()

        results = []
        errors = []

        for i, part in enumerate(selected_parts):
            if progress.wasCanceled():
                break

            progress.setLabelText(f"Kontrol ediliyor: {part['siparis_no']}")
            progress.setValue(i)

            # Her parça için durum kontrolü yap
            result = checker.check_part_status(part['siparis_no'], part['montaj_tarihi'])
            result['siparis_no'] = part['siparis_no']
            result['montaj_tarihi'] = part['montaj_tarihi']
            results.append(result)

            # Hata varsa kaydet
            if 'error' in result:
                errors.append(f"Sipariş {part['siparis_no']}: {result['error']}")

            # UI'yi güncel tut
            QApplication.processEvents()

        progress.setValue(len(selected_parts))
        progress.close()

        # Hataları göster
        if errors:
            error_summary = "\n".join(errors[:5])  # İlk 5 hata
            if len(errors) > 5:
                error_summary += f"\n... ve {len(errors) - 5} hata daha"

            QMessageBox.warning(
                self,
                "Parça Durum Hataları",
                f"{len(errors)} sipariş için hata oluştu:\n\n{error_summary}"
            )

        # Sonuçları tabloya yaz
        if results:
            # Tabloyu güncelle
            for result in results:
                siparis_no = result.get('siparis_no')

                # orderStatus bilgisini çıkar
                order_status = ""
                if result.get("success") == True:
                    if result.get("data") and len(result["data"]) > 0:
                        first_record = result["data"][0]
                        order_status = first_record.get("orderStatus", "")

                # ssh_data ve filtered_data'yı güncelle
                for data_row in self.ssh_data:
                    if str(data_row.get("Yedek Parça Sipariş No", "")).strip() == str(siparis_no).strip():
                        data_row["Parça Durumu"] = order_status
                        break

                for data_row in self.filtered_data:
                    if str(data_row.get("Yedek Parça Sipariş No", "")).strip() == str(siparis_no).strip():
                        data_row["Parça Durumu"] = order_status
                        break

                # Tabloda ilgili satırı bul ve güncelle
                for row in range(self.table.rowCount()):
                    # Yedek Parça Sipariş No sütununu bul
                    for col in range(2, self.table.columnCount()):  # 2'den başla (Seç ve Parça Durumu atla)
                        header = self.table.horizontalHeaderItem(col)
                        if header and "Yedek Parça Sipariş No" in header.text():
                            item = self.table.item(row, col)
                            if item and item.text().strip() == str(siparis_no).strip():
                                # Parça Durumu sütununu güncelle (index 1)
                                parca_durumu_item = self.table.item(row, 1)
                                if parca_durumu_item:
                                    parca_durumu_item.setText(order_status)
                                else:
                                    new_item = QTableWidgetItem(order_status)
                                    font = QFont("Segoe UI", 12)
                                    font.setBold(True)
                                    new_item.setFont(font)
                                    self.table.setItem(row, 1, new_item)
                                break

            self.status_label.setText(f"✅ {len(results)} parça durumu güncellendi")

            # Yazdır butonu durumunu kontrol et
            self.check_print_button_state()
        else:
            self.status_label.setText("ℹ️ Hiç sonuç bulunamadı")

    def refresh_data(self):
        """SSH verilerini yenile"""
        if self.data_loader and self.data_loader.isRunning():
            return

        self.refresh_btn.setEnabled(False)

        # Progress bar'ı göster ve başlat
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.status_label.setText("🔄 SSH verileri yükleniyor...")

        # Data loader thread'ini başlat
        self.data_loader = SshDataLoader()
        self.data_loader.data_loaded.connect(self.on_data_loaded)
        self.data_loader.error_occurred.connect(self.on_error_occurred)
        self.data_loader.finished.connect(self.on_loading_finished)
        self.data_loader.progress_updated.connect(self.on_progress_updated)
        self.data_loader.start()

    def on_progress_updated(self, progress, message):
        """Progress güncellemesi"""
        self.progress_bar.setValue(progress)
        self.status_label.setText(message)
        QApplication.processEvents()

    def on_data_loaded(self, data):
        """Veri yükleme tamamlandığında çağrılır"""
        try:
            # Ham veriyi kaydet
            self.ssh_data = data

            # Temel filtreleme: "Çözüldü" + "Sorunsuz Teslimat" (veya boş) olanları gizle
            self.status_label.setText(f"🔄 {len(data)} kayıt filtreleniyor...")
            self.progress_bar.setValue(0)
            QApplication.processEvents()

            self.base_filtered_data = []
            total = len(data)
            for i, row in enumerate(data):
                parca_durumu = str(row.get("Parça Durumu", "")).strip()
                belge_durum_nedeni = str(row.get("Belge Durum Nedeni", "")).strip()

                # Eğer Parça Durumu "Çözüldü" VE Belge Durum Nedeni boş veya "Sorunsuz Teslimat" ise atla
                if parca_durumu == "Çözüldü" and (belge_durum_nedeni == "" or belge_durum_nedeni == "Sorunsuz Teslimat"):
                    continue

                self.base_filtered_data.append(row)

                # Progress güncelle (her 100 kayıtta bir)
                if i % 100 == 0:
                    progress = int((i / total) * 100)
                    self.progress_bar.setValue(progress)
                    QApplication.processEvents()

            # Başlangıçta arama filtresi olmadığı için filtered_data = base_filtered_data
            self.filtered_data = self.base_filtered_data.copy()

            self.progress_bar.setValue(100)
            self.status_label.setText(f"🔄 Tablo oluşturuluyor ({len(self.filtered_data)} kayıt)...")
            QApplication.processEvents()

            self.populate_table()

            # Progress bar'ı gizle
            QTimer.singleShot(1000, lambda: self.progress_bar.setVisible(False))
            self.status_label.setText(f"✅ {len(self.filtered_data)} kayıt yüklendi (Toplam: {len(data)})")
        except Exception as e:
            logger.error(f"Veri yükleme hatası: {str(e)}")
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"❌ Veri işleme hatası: {str(e)}")

    def on_error_occurred(self, error_message):
        """Hata oluştuğunda çağrılır"""
        QMessageBox.critical(self, "SSH Veri Yükleme Hatası", error_message)

    def on_loading_finished(self):
        """Yükleme işlemi bittiğinde çağrılır"""
        self.refresh_btn.setEnabled(True)

    def populate_table(self):
        """Tabloyu verilerle doldur - ultra optimize sürümü kullan"""
        self.populate_table_ultra_optimized()

    def populate_table_optimized(self):
        """Tabloyu optimize edilmiş şekilde doldur - sozlesme_module stili"""
        if not self.filtered_data:
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            self.record_count_label.setText("Toplam: 0 kayıt")
            return

        # Tablo güncellemesini hızlandırmak için sinyalleri geçici olarak kapat
        self.table.blockSignals(True)
        self.table.setSortingEnabled(False)

        try:
            # Başlıkları hazırla - İstenilen sıralama
            desired_order = [
                'Belge Durum Nedeni',
                'Montaj Belgesi Tarihi',
                'Müşteri Adı',
                'Ürün Adı',
                'Yedek Parça Ürün Tanımı',
                'Yedek Parça Ürün Miktarı',
                'Sözleşme Numarası',
                'Servis Bakım ID',
                'Yedek Parça Sipariş No',
                'Ürün ID',
                'Yedek Parça Ürün ID'
            ]

            # Mevcut tüm sütunları al
            all_headers = list(self.filtered_data[0].keys())

            # Parça Durumu'nu çıkar (zaten ilk sütun olacak)
            if "Parça Durumu" in all_headers:
                all_headers.remove("Parça Durumu")

            # İstenilen sıralamaya göre sütunları düzenle
            original_headers = []
            for header in desired_order:
                if header in all_headers:
                    original_headers.append(header)

            # Listede olmayan sütunları sonuna ekle
            for header in all_headers:
                if header not in original_headers:
                    original_headers.append(header)

            headers = ["Seç", "Parça Durumu"] + original_headers  # Checkbox ve Parça Durumu sütunları

            # Tablo boyutlarını ayarla
            self.table.setRowCount(len(self.filtered_data))
            self.table.setColumnCount(len(headers))
            self.table.setHorizontalHeaderLabels(headers)

            # Verileri tabloya ekle
            for row_idx, row_data in enumerate(self.filtered_data):
                # İlk sütuna checkbox ekle
                checkbox_widget = QWidget()
                checkbox_layout = QHBoxLayout(checkbox_widget)
                checkbox_layout.setContentsMargins(0, 0, 0, 0)
                checkbox_layout.setAlignment(Qt.AlignCenter)

                checkbox = QCheckBox()
                checkbox.setChecked(True)  # Varsayılan olarak işaretli
                checkbox.setStyleSheet("""
                    QCheckBox {
                        font-size: 14px;
                        font-weight: bold;
                    }
                    QCheckBox::indicator {
                        width: 18px;
                        height: 18px;
                    }
                """)
                checkbox_layout.addWidget(checkbox)
                self.table.setCellWidget(row_idx, 0, checkbox_widget)

                # Parça Durumu sütunu (index 1)
                parca_durumu = row_data.get("Parça Durumu", "")
                item = QTableWidgetItem(str(parca_durumu))
                font = QFont("Segoe UI", 12)
                font.setBold(True)
                item.setFont(font)
                self.table.setItem(row_idx, 1, item)

                # Diğer sütunlara veri ekle
                for col_idx, header in enumerate(original_headers):
                    value = row_data.get(header, "")
                    item = QTableWidgetItem(str(value))

                    # risk_module stili font
                    font = QFont("Segoe UI", 12)
                    font.setBold(True)
                    item.setFont(font)

                    self.table.setItem(row_idx, col_idx + 2, item)  # +2 çünkü ilk iki sütun checkbox ve Parça Durumu

            # Sütun genişliklerini ayarla
            header = self.table.horizontalHeader()

            # Checkbox sütunu için sabit genişlik
            header.setSectionResizeMode(0, QHeaderView.Fixed)
            self.table.setColumnWidth(0, 60)

            # Parça Durumu sütunu için sabit genişlik
            header.setSectionResizeMode(1, QHeaderView.Fixed)
            self.table.setColumnWidth(1, 120)

            # Diğer sütunlar için otomatik boyutlandırma
            for i in range(2, len(headers)):
                header.setSectionResizeMode(i, QHeaderView.ResizeToContents)

            # Satır yüksekliğini artır (sozlesme_module gibi)
            self.table.verticalHeader().setDefaultSectionSize(
                self.table.verticalHeader().defaultSectionSize() + 5
            )

        finally:
            # Sinyalleri tekrar aç
            self.table.setSortingEnabled(True)
            self.table.blockSignals(False)

        # Kayıt sayısı bilgisi kaldırıldı

    def on_search_text_changed(self):
        """Arama metni değiştiğinde çağrılır - debounce için"""
        self.search_timer.stop()
        self.search_timer.start(500)  # 500ms bekle (daha uzun debounce)

    def filter_data(self):
        """Arama filtreleme işlemi - stok_module regex bazlı"""
        search_text = self.search_input.text().strip().lower()

        # UI'yi blokla
        self.table.setUpdatesEnabled(False)

        try:
            if not search_text:
                # Arama yoksa base_filtered_data'yı kullan (zaten temel filtreleme yapılmış)
                self.filtered_data = self.base_filtered_data.copy()
            else:
                # Regex pattern oluştur (her kelime için AND operasyonu)
                parts = [re.escape(part) for part in search_text.split() if part]
                pattern = r'(?=.*?{})'.format(')(?=.*?'.join(parts))

                # Arama yapılacak sütunlar
                search_columns = ["Müşteri Adı", "Ürün Adı", "Yedek Parça Ürün Tanımı", "Sözleşme Numarası"]

                # Regex bazlı filtreleme - base_filtered_data üzerinde ara
                self.filtered_data = []
                for row in self.base_filtered_data:
                    # Tüm arama sütunlarını birleştir
                    combined_text = " ".join([
                        str(row.get(column, "")).lower()
                        for column in search_columns
                        if column in row
                    ])

                    # Pattern ile eşleşme kontrolü
                    if re.search(pattern, combined_text):
                        self.filtered_data.append(row)

            # Tabloyu güncelle
            self.populate_table_ultra_optimized()

            # Sonuç bilgisi
            self.status_label.setText(f"✅ {len(self.filtered_data)} kayıt gösteriliyor (Toplam: {len(self.base_filtered_data)})")

        finally:
            # UI'yi tekrar aktif et
            self.table.setUpdatesEnabled(True)

    def populate_table_ultra_optimized(self):
        """Ultra optimize edilmiş tablo doldurma - büyük veriler için"""
        if not self.filtered_data:
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            return

        # Tüm sinyalleri kapat ve UI güncellemelerini durdur
        self.table.setVisible(False)  # Tablo görünmezken doldur
        self.table.blockSignals(True)
        self.table.setSortingEnabled(False)
        self.table.setUpdatesEnabled(False)

        try:
            # Başlıkları hazırla - İstenilen sıralama
            desired_order = [
                'Belge Durum Nedeni',
                'Montaj Belgesi Tarihi',
                'Müşteri Adı',
                'Ürün Adı',
                'Yedek Parça Ürün Tanımı',
                'Yedek Parça Ürün Miktarı',
                'Sözleşme Numarası',
                'Servis Bakım ID',
                'Yedek Parça Sipariş No',
                'Ürün ID',
                'Yedek Parça Ürün ID'
            ]

            # Mevcut tüm sütunları al
            all_headers = list(self.filtered_data[0].keys())

            # Parça Durumu'nu çıkar (zaten ilk sütun olacak)
            if "Parça Durumu" in all_headers:
                all_headers.remove("Parça Durumu")

            # İstenilen sıralamaya göre sütunları düzenle
            original_headers = []
            for header in desired_order:
                if header in all_headers:
                    original_headers.append(header)

            # Listede olmayan sütunları sonuna ekle
            for header in all_headers:
                if header not in original_headers:
                    original_headers.append(header)

            headers = ["Seç", "Parça Durumu"] + original_headers

            # Tablo boyutunu ayarla
            row_count = len(self.filtered_data)
            col_count = len(headers)

            # Tablo boyutunu sadece gerekirse değiştir
            if self.table.rowCount() != row_count:
                self.table.setRowCount(row_count)
            if self.table.columnCount() != col_count:
                self.table.setColumnCount(col_count)
                self.table.setHorizontalHeaderLabels(headers)

                # Sütun genişliklerini sadece yeni tablo için ayarla
                header = self.table.horizontalHeader()
                header.setSectionResizeMode(0, QHeaderView.Fixed)
                self.table.setColumnWidth(0, 60)

                # Parça Durumu sütunu için sabit genişlik
                header.setSectionResizeMode(1, QHeaderView.Fixed)
                self.table.setColumnWidth(1, 120)

                for i in range(2, col_count):
                    header.setSectionResizeMode(i, QHeaderView.ResizeToContents)

            # Batch şeklinde veri ekle
            items_to_set = []
            widgets_to_set = []

            for row_idx, row_data in enumerate(self.filtered_data):
                # Checkbox widget'ını hazırla
                if not self.table.cellWidget(row_idx, 0):
                    checkbox_widget = QWidget()
                    checkbox_layout = QHBoxLayout(checkbox_widget)
                    checkbox_layout.setContentsMargins(0, 0, 0, 0)
                    checkbox_layout.setAlignment(Qt.AlignCenter)

                    checkbox = QCheckBox()
                    checkbox.setChecked(True)
                    checkbox.setStyleSheet("""
                        QCheckBox {
                            font-size: 14px;
                            font-weight: bold;
                        }
                        QCheckBox::indicator {
                            width: 18px;
                            height: 18px;
                        }
                    """)
                    # Checkbox değişiminde Yazdır butonu durumunu kontrol et
                    checkbox.stateChanged.connect(self.check_print_button_state)
                    checkbox_layout.addWidget(checkbox)
                    widgets_to_set.append((row_idx, 0, checkbox_widget))

                # Parça Durumu sütunu (index 1)
                parca_durumu = row_data.get("Parça Durumu", "")
                if isinstance(parca_durumu, (int, float)):
                    if isinstance(parca_durumu, float) and parca_durumu.is_integer():
                        display_value = str(int(parca_durumu))
                    else:
                        display_value = str(parca_durumu)
                elif pd.isna(parca_durumu) or parca_durumu is None:
                    display_value = ""
                else:
                    display_value = str(parca_durumu)

                existing_item = self.table.item(row_idx, 1)
                if existing_item:
                    existing_item.setText(display_value)
                else:
                    item = QTableWidgetItem(display_value)
                    font = QFont("Segoe UI", 12)
                    font.setBold(True)
                    item.setFont(font)
                    items_to_set.append((row_idx, 1, item))

                # Veri item'larını hazırla
                for col_idx, header in enumerate(original_headers):
                    value = row_data.get(header, "")

                    # Sayısal değerlerde .0 ifadesini kaldır
                    if isinstance(value, (int, float)):
                        # Eğer float ama tam sayı ise, int olarak göster
                        if isinstance(value, float) and value.is_integer():
                            display_value = str(int(value))
                        else:
                            display_value = str(value)
                    elif pd.isna(value) or value is None:
                        display_value = ""
                    else:
                        display_value = str(value)

                    existing_item = self.table.item(row_idx, col_idx + 2)
                    if existing_item:
                        existing_item.setText(display_value)
                    else:
                        item = QTableWidgetItem(display_value)
                        font = QFont("Segoe UI", 12)
                        font.setBold(True)
                        item.setFont(font)
                        items_to_set.append((row_idx, col_idx + 2, item))

            # Batch insert işlemleri
            for widget_data in widgets_to_set:
                self.table.setCellWidget(*widget_data)

            for item_data in items_to_set:
                self.table.setItem(*item_data)

        finally:
            # Her şeyi tekrar aç
            self.table.setUpdatesEnabled(True)
            self.table.setSortingEnabled(True)
            self.table.blockSignals(False)
            self.table.setVisible(True)  # Tablo dolduktan sonra göster

        # Kayıt sayısı bilgisi kaldırıldı

    def export_to_excel(self):
        """Filtrelenmiş SSH verilerini Excel'e aktar"""
        if not self.filtered_data:
            QMessageBox.warning(self, "Uyarı", "Dışa aktarılacak veri yok.")
            return

        try:
            # Excel'e aktarılacak DataFrame oluştur
            df = pd.DataFrame(self.filtered_data)

            # Çıktı dosya yolu
            output_path = "D:/GoogleDrive/~ SSH_Export.xlsx"
            df.to_excel(output_path, index=False, engine='openpyxl')

            self.status_label.setText(f"✅ Veriler dışa aktarıldı: {output_path}")
            QMessageBox.information(self, "Başarılı", f"Veriler başarıyla dışa aktarıldı:\n{output_path}")

        except Exception as e:
            error_msg = f"Dışa aktarma hatası: {str(e)}"
            logger.error(error_msg)
            self.status_label.setText(f"❌ {error_msg}")
            QMessageBox.critical(self, "Hata", error_msg)

    def run_montaj_raporu(self):
        """SSH.exe dosyasını çalıştır ve tabloyu güncelle"""
        try:
            exe_path = r"D:/GoogleDrive/PRG/EXE/SSH.exe"
            if not os.path.exists(exe_path):
                self.status_label.setText(f"❌ SSH.exe bulunamadı: {exe_path}")
                return

            self.status_label.setText("🔄 SSH.exe çalıştırılıyor...")
            self.montaj_raporu_btn.setEnabled(False)
            self.ssh_raporu_calisiyor = True
            QApplication.processEvents()

            # SSH.exe'yi çalıştır
            os.startfile(exe_path)

            # SSH.exe'nin çalışması için bekleme (7 saniye)
            QTimer.singleShot(7000, self.on_ssh_exe_finished)

        except Exception as e:
            error_msg = f"Program çalıştırma hatası: {str(e)}"
            logger.error(error_msg)
            self.status_label.setText(f"❌ {error_msg}")
            self.montaj_raporu_btn.setEnabled(True)
            self.ssh_raporu_calisiyor = False

    def on_ssh_exe_finished(self):
        """SSH.exe bittikten sonra çağrılır"""
        self.montaj_raporu_btn.setEnabled(True)
        self.ssh_raporu_calisiyor = False
        self.status_label.setText("✅ SSH.exe tamamlandı, Google Sheets güncelleme bekleniyor...")

        # Google Sheets'e kaydedilmesi için ek bekleme (5 saniye)
        # Sonra verileri yenile
        QTimer.singleShot(5000, self.delayed_data_refresh)

    def delayed_data_refresh(self):
        """Gecikmeli veri yenileme"""
        self.status_label.setText("🔄 Google Sheets'ten güncel veriler alınıyor...")
        QApplication.processEvents()
        self.refresh_data()

    def run_mikro_ssh(self):
        """Tamamlanan.exe dosyasını çalıştır ve tabloyu güncelle"""
        try:
            exe_path = r"D:/GoogleDrive/PRG/EXE/Tamamlanan.exe"
            if not os.path.exists(exe_path):
                self.status_label.setText(f"❌ Tamamlanan.exe bulunamadı: {exe_path}")
                return

            self.status_label.setText("🔄 Tamamlanan.exe çalıştırılıyor...")
            self.mikro_button.setEnabled(False)
            QApplication.processEvents()

            # Tamamlanan.exe'yi çalıştır
            os.startfile(exe_path)

            # Tamamlanan.exe'nin çalışması için bekleme (7 saniye)
            QTimer.singleShot(7000, self.on_mikro_guncelle_finished)

        except Exception as e:
            error_msg = f"Program çalıştırma hatası: {str(e)}"
            logger.error(error_msg)
            self.status_label.setText(f"❌ {error_msg}")
            self.mikro_button.setEnabled(True)

    def on_mikro_guncelle_finished(self):
        """Tamamlanan.exe bittikten sonra çağrılır"""
        self.mikro_button.setEnabled(True)
        self.status_label.setText("✅ Tamamlanan.exe tamamlandı, Google Sheets güncelleme bekleniyor...")

        # Google Sheets'e kaydedilmesi için ek bekleme (5 saniye)
        QTimer.singleShot(5000, self.delayed_data_refresh)

    def run_montaj_yukle(self):
        """Montaj.exe dosyasını çalıştır ve tabloyu güncelle"""
        try:
            exe_path = r"D:/GoogleDrive/PRG/EXE/Montaj.exe"
            if not os.path.exists(exe_path):
                self.status_label.setText(f"❌ Montaj.exe bulunamadı: {exe_path}")
                QMessageBox.warning(self, "Uyarı", f"Montaj.exe bulunamadı:\n{exe_path}")
                return

            self.status_label.setText("🔄 Montaj.exe çalıştırılıyor...")
            self.montaj_yukle_btn.setEnabled(False)
            QApplication.processEvents()

            # Montaj.exe'yi çalıştır
            os.startfile(exe_path)

            # Montaj.exe'nin çalışması için bekleme (7 saniye)
            QTimer.singleShot(7000, self.on_montaj_yukle_finished)

        except Exception as e:
            error_msg = f"Program çalıştırma hatası: {str(e)}"
            logger.error(error_msg)
            self.status_label.setText(f"❌ {error_msg}")
            QMessageBox.critical(self, "Hata", error_msg)
            self.montaj_yukle_btn.setEnabled(True)

    def on_montaj_yukle_finished(self):
        """Montaj.exe bittikten sonra çağrılır"""
        self.montaj_yukle_btn.setEnabled(True)
        self.status_label.setText("✅ Montaj.exe tamamlandı, Google Sheets güncelleme bekleniyor...")

        # Google Sheets'e kaydedilmesi için ek bekleme (5 saniye)
        QTimer.singleShot(5000, self.delayed_data_refresh)

    def show_montaj_belgesi_dialog(self):
        """Montaj Belgesi oluşturulmayan sözleşmeleri göster"""
        try:
            # Progress bar'ı göster
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(0)
            self.status_label.setText("📊 Google Sheets bağlantısı kuruluyor...")
            QApplication.processEvents()

            # Google Sheets'ten Tamamlanan, Montaj ve Siparisler sayfalarını oku - Service Account
            config_manager = CentralConfigManager()
            sheets_manager = config_manager.gc

            # PRGsheet dosyasını aç
            spreadsheet = sheets_manager.open("PRGsheet")

            # 3 sayfayı paralel thread'lerle oku (daha hızlı!)
            self.status_label.setText("📥 Tüm sayfalar okunuyor...")
            self.progress_bar.setValue(10)
            QApplication.processEvents()

            from concurrent.futures import ThreadPoolExecutor, as_completed

            def read_worksheet(sheet_name):
                """Bir worksheet'i oku"""
                try:
                    worksheet = spreadsheet.worksheet(sheet_name)
                    return sheet_name, worksheet.get_all_values()
                except Exception as e:
                    logger.error(f"{sheet_name} okuma hatası: {e}")
                    return sheet_name, []

            # 3 sayfayı paralel oku
            with ThreadPoolExecutor(max_workers=3) as executor:
                futures = {
                    executor.submit(read_worksheet, 'Tamamlanan'): 'Tamamlanan',
                    executor.submit(read_worksheet, 'Montaj'): 'Montaj',
                    executor.submit(read_worksheet, 'Siparisler'): 'Siparisler'
                }

                results = {}
                for future in as_completed(futures):
                    sheet_name, data = future.result()
                    results[sheet_name] = data

            tamamlanan_data = results.get('Tamamlanan', [])
            montaj_data = results.get('Montaj', [])
            siparisler_data = results.get('Siparisler', [])

            self.progress_bar.setValue(30)
            QApplication.processEvents()

            if len(tamamlanan_data) <= 1:
                self.progress_bar.setVisible(False)
                self.status_label.setText("⚠️ Tamamlanan sayfasında veri bulunamadı")
                return

            if len(montaj_data) <= 1:
                self.progress_bar.setVisible(False)
                self.status_label.setText("⚠️ Montaj sayfasında veri bulunamadı")
                return

            if len(siparisler_data) <= 1:
                self.progress_bar.setVisible(False)
                self.status_label.setText("⚠️ Siparisler sayfasında veri bulunamadı")
                return

            # Tamamlanan verilerini DataFrame'e çevir
            tamamlanan_df = pd.DataFrame(tamamlanan_data[1:], columns=tamamlanan_data[0])

            # Montaj verilerini DataFrame'e çevir
            montaj_df = pd.DataFrame(montaj_data[1:], columns=montaj_data[0])

            # Siparisler verilerini DataFrame'e çevir
            siparisler_df = pd.DataFrame(siparisler_data[1:], columns=siparisler_data[0])

            # Sütun kontrolü
            self.status_label.setText("🔍 Sütunlar kontrol ediliyor...")
            self.progress_bar.setValue(40)
            QApplication.processEvents()

            if 'sip_belgeno' not in tamamlanan_df.columns:
                self.progress_bar.setVisible(False)
                self.status_label.setText("❌ 'sip_belgeno' sütunu Tamamlanan sayfasında bulunamadı")
                return

            if 'sip_musteri_kod' not in tamamlanan_df.columns:
                self.progress_bar.setVisible(False)
                self.status_label.setText("❌ 'sip_musteri_kod' sütunu Tamamlanan sayfasında bulunamadı")
                return

            if 'Sözleşme Numarası' not in montaj_df.columns:
                self.progress_bar.setVisible(False)
                self.status_label.setText("❌ 'Sözleşme Numarası' sütunu Montaj sayfasında bulunamadı")
                return

            if 'Cari Kod' not in siparisler_df.columns:
                self.progress_bar.setVisible(False)
                self.status_label.setText("❌ 'Cari Kod' sütunu Siparisler sayfasında bulunamadı")
                return

            if 'Cari Adi' not in siparisler_df.columns:
                self.progress_bar.setVisible(False)
                self.status_label.setText("❌ 'Cari Adi' sütunu Siparisler sayfasında bulunamadı")
                return

            # Siparisler lookup dictionary oluştur (Cari Kod -> Cari Adi)
            cari_lookup = {}
            for _, row in siparisler_df.iterrows():
                cari_kod = str(row['Cari Kod']).strip()
                cari_adi = str(row['Cari Adi']).strip()
                if cari_kod and cari_kod.lower() not in ['nan', 'none', '']:
                    cari_lookup[cari_kod] = cari_adi

            # Montaj sayfasındaki Sözleşme Numaralarını set'e çevir
            montaj_sozlesme_set = set()
            for sozlesme in montaj_df['Sözleşme Numarası']:
                sozlesme_str = str(sozlesme).strip()
                if sozlesme_str and sozlesme_str.lower() not in ['nan', 'none', '']:
                    if '.' in sozlesme_str:
                        try:
                            sozlesme_str = str(int(float(sozlesme_str)))
                        except:
                            pass
                    montaj_sozlesme_set.add(sozlesme_str)

            # Tamamlanan sayfasından Montaj'da olmayan kayıtları filtrele ve Cari Adı ekle
            # Önce API ile SHIPPING_COND kontrolü yapılacak kayıtları topla
            # DİNAMİK TARIH FİLTRESİ (PRGsheet/Ayar sayfasından okunuyor)
            from datetime import datetime, timedelta

            # Gün sayısını PRGsheet/Ayar'dan al (- cache yenile!)
            # NOT: Her Montaj Raporu çağrısında ayarları yeniden okur (PRGsheet değişmişse günceller)
            try:
                # Cache'i yenile ve güncel ayarları al
                fresh_settings = config_manager.get_settings(use_cache=False)
                gun_sayisi_str = fresh_settings.get('MONTAJ_FILTRE_GUN_SAYISI', '365')
                gun_sayisi = int(gun_sayisi_str)
                # MONTAJ_FILTRE_GUN_SAYISI ayarı okundu
            except Exception as e:
                logger.warning(f"Ayar okuma hatası: {e}, varsayılan 365 kullanılıyor")
                gun_sayisi = 365

            one_year_ago = datetime.now() - timedelta(days=gun_sayisi)

            self.status_label.setText(f"🔍 Kayıtlar filtreleniyor (Son {gun_sayisi} gün)...")
            self.progress_bar.setValue(45)
            QApplication.processEvents()

            temp_eslesmeyen_kayitlar = []
            for _, row in tamamlanan_df.iterrows():
                sip_belgeno = str(row['sip_belgeno']).strip()
                # .0 varsa kaldır
                if '.' in sip_belgeno:
                    try:
                        sip_belgeno = str(int(float(sip_belgeno)))
                    except:
                        pass

                # Tarih kontrolü (msg_S_1072 sütunu)
                try:
                    tarih_str = str(row.get('msg_S_1072', '')).strip()
                    if tarih_str and tarih_str.lower() not in ['nan', 'none', '']:
                        # Tarih formatı: "DD.MM.YYYY" veya "YYYY-MM-DD"
                        if '.' in tarih_str:
                            tarih = datetime.strptime(tarih_str, '%d.%m.%Y')
                        elif '-' in tarih_str:
                            tarih = datetime.strptime(tarih_str.split()[0], '%Y-%m-%d')
                        else:
                            tarih = one_year_ago  # Parse edilemezse dahil et

                        # Son 1 yıldan eskiyse atla
                        if tarih < one_year_ago:
                            continue
                except:
                    # Tarih parse edilemezse kayıt dahil edilsin (güvenli taraf)
                    pass

                # Boş değilse ve eşleşmiyorsa geçici listeye ekle
                if sip_belgeno and sip_belgeno.lower() not in ['nan', 'none', ''] and sip_belgeno not in montaj_sozlesme_set:
                    kayit = row.to_dict()

                    # sip_musteri_kod ile Cari Adi'ni bul (Siparisler sayfasından)
                    sip_musteri_kod = str(row['sip_musteri_kod']).strip()

                    # Cari Adi'ni ekle - Siparisler sayfasından
                    kayit['Cari Adı'] = cari_lookup.get(sip_musteri_kod, '')
                    kayit['_sip_belgeno'] = sip_belgeno  # API sorgusu için
                    temp_eslesmeyen_kayitlar.append(kayit)

            # Sozleme.py modülünü import et - Static import (PyInstaller uyumlu)
            self.status_label.setText("📦 Sozleme.py modülü yükleniyor...")
            self.progress_bar.setValue(50)
            QApplication.processEvents()

            try:
                from PRG import Sozleme as sozleme_module
            except ImportError:
                try:
                    import Sozleme as sozleme_module
                except ImportError as import_error:
                    self.progress_bar.setVisible(False)
                    self.status_label.setText(f"❌ Sozleme.py yüklenirken hata: {str(import_error)}")
                    return

            # ContractCheckWorker thread'i ile paralel API çağrıları
            self.contract_worker = ContractCheckWorker(sozleme_module, temp_eslesmeyen_kayitlar)
            self.contract_worker.progress_updated.connect(self.on_contract_check_progress)
            self.contract_worker.finished_signal.connect(self.on_contract_check_finished)
            self.contract_worker.error_occurred.connect(self.on_contract_check_error)
            self.contract_worker.start()

            # Worker thread çalışırken fonksiyon dönüyor (non-blocking)
            # Sonuçlar on_contract_check_finished callback'inde işlenecek
            return

        except Exception as e:
            error_msg = f"Montaj belgesi kontrol hatası: {str(e)}"
            logger.error(error_msg)
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"❌ {error_msg}")

    def on_contract_check_progress(self, progress, message):
        """Contract check progress güncellemesi"""
        self.progress_bar.setValue(progress)
        self.status_label.setText(message)
        QApplication.processEvents()

    def on_contract_check_finished(self, eslesmeyen_kayitlar):
        """Contract check tamamlandığında çağrılır"""
        try:
            self.progress_bar.setValue(90)

            if not eslesmeyen_kayitlar:
                self.progress_bar.setValue(100)
                QTimer.singleShot(1000, lambda: self.progress_bar.setVisible(False))
                self.status_label.setText("✅ Tüm sözleşmelerde montaj belgesi oluşturulmuş")
                return

            # Dialog penceresini göster
            self.progress_bar.setValue(100)
            self.status_label.setText(f"✅ {len(eslesmeyen_kayitlar)} kayıt montaj belgesi eksik")
            QApplication.processEvents()

            # Progress bar'ı gizle
            QTimer.singleShot(500, lambda: self.progress_bar.setVisible(False))

            dialog = MontajBelgesiDialog(eslesmeyen_kayitlar, self)
            dialog.exec_()

        except Exception as e:
            error_msg = f"Callback hatası: {str(e)}"
            logger.error(error_msg)
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"❌ {error_msg}")

    def on_contract_check_error(self, error_msg):
        """Contract check hatası oluştuğunda çağrılır"""
        logger.error(error_msg)
        self.progress_bar.setVisible(False)
        self.status_label.setText(f"❌ {error_msg}")

    def create_customer_info_group(self, title, info_dict):
        """
        Müşteri bilgilerini gösteren grup kutusu oluşturur.

        Args:
            title (str): Grup başlığı (örn: "MÜŞTERİ BİLGİLERİ")
            info_dict (dict): Müşteri bilgilerini içeren sözlük

        Returns:
            QGroupBox: Grid düzeninde müşteri bilgileri grubu
        """
        group_box = QGroupBox(title)
        group_box.setStyleSheet("""
            QGroupBox {
                font-size: 16px;
                font-weight: bold;
                border: 3px solid #3498db;
                border-radius: 8px;
                margin-top: 15px;
                padding-top: 15px;
                background-color: #f8f9fa;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 5px 10px 5px 10px;
                color: #2c3e50;
                background-color: #3498db;
                color: white;
                border-radius: 4px;
                font-size: 16px;
            }
        """)

        # Ana layout
        main_layout = QVBoxLayout()

        # Grid layout - 3 satır x 2 sütun
        grid_layout = QHBoxLayout()

        # Sol sütun
        left_layout = QVBoxLayout()
        left_items = [
            ("Ad Soyad:", info_dict.get('ad_soyad', '')),
            ("Telefon 1:", info_dict.get('telefon1', '')),
            ("Telefon 2:", info_dict.get('telefon2', ''))
        ]

        for label_text, value in left_items:
            if value and str(value).strip() and str(value) != 'N/A':
                item_layout = QHBoxLayout()

                label = QLabel(label_text)
                label.setStyleSheet("font-weight: bold; color: #34495e; font-size: 16px;")

                value_label = QLabel(str(value))
                value_label.setStyleSheet("color: #2c3e50; font-size: 16px; font-weight: bold;")
                value_label.setWordWrap(True)

                item_layout.addWidget(label)
                item_layout.addWidget(value_label, 1)
                left_layout.addLayout(item_layout)

        # Sağ sütun
        right_layout = QVBoxLayout()
        right_items = [
            ("TCKN No:", info_dict.get('vergi_no', '')),
            ("Şehir:", info_dict.get('sehir_ilce', ''))
        ]

        for label_text, value in right_items:
            if value and str(value).strip() and str(value) != 'N/A':
                item_layout = QHBoxLayout()

                label = QLabel(label_text)
                label.setStyleSheet("font-weight: bold; color: #34495e; font-size: 16px;")

                value_label = QLabel(str(value))
                value_label.setStyleSheet("color: #2c3e50; font-size: 16px; font-weight: bold;")
                value_label.setWordWrap(True)

                item_layout.addWidget(label)
                item_layout.addWidget(value_label, 1)
                right_layout.addLayout(item_layout)

        # Grid'e sütunları ekle
        grid_layout.addLayout(left_layout)
        grid_layout.addSpacing(20)  # Sütunlar arası boşluk
        grid_layout.addLayout(right_layout)

        main_layout.addLayout(grid_layout)

        # Adres bilgisini en alta ekle (tam genişlikte)
        adres = info_dict.get('adres', '')
        if adres and str(adres).strip() and str(adres) != 'N/A':
            adres_layout = QHBoxLayout()

            adres_label = QLabel("Adres:")
            adres_label.setStyleSheet("font-weight: bold; color: #34495e; font-size: 16px;")

            adres_value = QLabel(str(adres))
            adres_value.setStyleSheet("color: #2c3e50; font-size: 16px; font-weight: bold;")
            adres_value.setWordWrap(True)

            adres_layout.addWidget(adres_label)
            adres_layout.addWidget(adres_value, 1)

            main_layout.addLayout(adres_layout)

        group_box.setLayout(main_layout)
        return group_box

    def show_customer_info_window(self, contract_data, contract_id):
        """Müşteri bilgileri penceresini göster"""
        try:
            # CustomerInfoWindow sınıfını oluştur ve göster
            info_window = CustomerInfoWindow(contract_data, contract_id, self)
            info_window.show()
        except Exception as e:
            logger.error(f"Müşteri bilgileri penceresi hatası: {str(e)}")
            QMessageBox.critical(self, "Hata", f"Müşteri bilgileri penceresi hatası: {str(e)}")


class SSHPrintDialog(QDialog):
    """SSH Arıza Formu yazdırma penceresi"""

    def __init__(self, contract_data, selected_rows_data, parent=None):
        super().__init__(parent)
        self.setWindowTitle("SSH - ARIZA FORMU")
        self.setGeometry(100, 100, 1200, 800)
        self.setStyleSheet("background-color: white;")

        self.layout = QVBoxLayout(self)

        # Metin alanı
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        self.text_edit.setStyleSheet("background-color: white;")
        self.layout.addWidget(self.text_edit)

        # Butonlar için yatay layout
        button_layout = QHBoxLayout()

        # Yazıcıya Gönder butonu
        self.btn_print = QPushButton("Yazıcıya Gönder")
        self.btn_print.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 15px;
                font-size: 12px;
                font-weight: bold;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
            QPushButton:pressed {
                background-color: #909090;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #888888;
            }
        """)
        self.btn_print.clicked.connect(self.print_data)
        button_layout.addWidget(self.btn_print)

        # Sorun Çözüldü butonu
        self.btn_montor = QPushButton("Sorun Çözüldü")
        self.btn_montor.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 15px;
                font-size: 12px;
                font-weight: bold;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
            QPushButton:pressed {
                background-color: #909090;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #888888;
            }
        """)
        self.btn_montor.clicked.connect(self.update_to_montor)
        button_layout.addWidget(self.btn_montor)

        self.layout.addLayout(button_layout)

        # Verileri sakla
        self.contract_data = contract_data
        self.selected_rows_data = selected_rows_data

        # Verileri metin alanına yükle
        self.load_data()

    def load_data(self):
        """Verileri metin alanına yükler"""
        # HTML başlangıcı
        html_content = """
        <!DOCTYPE html>
        <html>
        <head>
            <style>
                @page {
                    size: A4;
                    margin: 10mm;
                }
                body {
                    font-family: 'Segoe UI', Arial, sans-serif;
                    margin: 0;
                    padding: 0;
                    font-size: 14pt !important;
                    background-color: white;
                    color: black;
                }
                .header {
                    text-align: center;
                    margin-bottom: 20px;
                    color: black;
                }
                .customer-info {
                    margin-bottom: 20px;
                    font-size: 18pt !important;
                    color: black;
                }
                .print-table {
                    width: 100%;
                    border-collapse: collapse;
                    margin-top: 20px;
                    font-size: 16pt !important;
                    background-color: white;
                }
                .print-table th, .print-table td {
                    border: 2px solid black;
                    padding: 10px;
                    text-align: left;
                    background-color: white;
                    color: black;
                }
                .print-table th {
                    font-weight: bold;
                }
                h1 {
                    font-size: 24pt !important;
                    margin-bottom: 15px;
                    color: black;
                }
                p {
                    font-size: 18pt !important;
                    margin: 8px 0;
                    color: black;
                }
                strong {
                    font-weight: bold;
                    color: black;
                }
                .footer {
                    margin-top: 40px;
                    font-size: 16pt !important;
                    color: black;
                }
                .signature-container {
                    display: flex;
                    justify-content: space-between;
                    margin-top: 5px;
                    align-items: baseline;
                    color: black;
                }
                .customer-section {
                    text-align: left;
                    width: 45%;
                    color: black;
                }
                .technician-section {
                    text-align: right;
                    width: 45%;
                    color: black;
                }
                .service-info {
                    margin-bottom: 60px;
                    color: black;
                }
                .problem-solved {
                    font-weight: bold;
                    margin-bottom: 5px;
                    color: black;
                }
            </style>
        </head>
        <body>
        """

        # Başlık
        html_content += """
        <div class="header">
            <h1>SSH - ARIZA FORMU</h1>
        </div>
        """

        # Müşteri bilgilerini contract_data'dan al
        if hasattr(self.contract_data, 'ES_CONTRACT_INFO'):
            contract_info = self.contract_data.ES_CONTRACT_INFO

            def safe_get(obj, attr, default=''):
                if obj is None:
                    return default
                return getattr(obj, attr, default) if hasattr(obj, attr) else default

            customer_name = f"{safe_get(contract_info, 'CUSTOMER_NAMEFIRST')} {safe_get(contract_info, 'CUSTOMER_NAMELAST')}".strip()
            phone1 = safe_get(contract_info, 'CUSTOMER_PHONE1')
            phone2 = safe_get(contract_info, 'CUSTOMER_PHONE2')
            address = safe_get(contract_info, 'CUSTOMER_ADDRESS')
            city = safe_get(contract_info, 'CUSTOMER_CITY')

            # Telefon formatı
            telefon_str = f"{phone1} - {phone2}" if phone1 and phone2 else (phone1 or phone2 or '')

            html_content += f"""
            <div class="customer-info">
                <p><strong>MÜŞTERİ ADI & SOYADI:</strong> {customer_name}</p>
                <p><strong>TELEFON:</strong> {telefon_str}</p>
                <p><strong>ADRES:</strong> {address}</p>
                <p><strong>ŞEHİR:</strong> {city}</p>
            </div>
            """
        else:
            html_content += """
            <div class="customer-info">
                <p><strong>MÜŞTERİ BİLGİLERİ:</strong> Bilgi alınamadı</p>
            </div>
            """

        # Tablo başlıkları
        headers = ["YEDEK PARÇA ID", "ÜRÜN ADI", "YEDEK PARÇA", "MİKTAR"]

        # Tablo oluşturma
        html_content += """
        <table class="print-table">
            <thead>
                <tr>
        """

        # Başlıkları ekle
        for header in headers:
            html_content += f"<th>{header}</th>"
        html_content += """
                </tr>
            </thead>
            <tbody>
        """

        # Veri satırlarını ekle
        for row_data in self.selected_rows_data:
            html_content += "<tr>"
            # Yedek Parça Ürün ID, ÜRÜN ADI, YEDEK PARÇA, MİKTAR
            yedek_parca_id = row_data.get("Yedek Parça Ürün ID", "")
            urun_adi = row_data.get("Ürün Adı", "")
            yedek_parca = row_data.get("Yedek Parça Ürün Tanımı", "")
            miktar = row_data.get("Yedek Parça Ürün Miktarı", "")

            for item in [yedek_parca_id, urun_adi, yedek_parca, miktar]:
                # .0 formatını temizle
                cleaned_item = str(item).replace('.0', '') if item else ""
                html_content += f"<td>{cleaned_item}</td>"
            html_content += "</tr>"

        html_content += """
            </tbody>
        </table>
        """

        # Müşteri adını al
        customer_name = ""
        if hasattr(self.contract_data, 'ES_CONTRACT_INFO'):
            contract_info = self.contract_data.ES_CONTRACT_INFO
            customer_name = f"{getattr(contract_info, 'CUSTOMER_NAMEFIRST', '')} {getattr(contract_info, 'CUSTOMER_NAMELAST', '')}".strip()

        # Açıklama ve imza bölümü
        html_content += f"""
        <div class="footer">
            <div class="service-info">
                <p>Her türlü arıza için DOĞTAŞ SERVİS 0850 800 34 87 numarasını arayabilirsiniz.</p>
                <br />
                <br />
            </div>

            <div class="signature-container">
                <div class="customer-section">
                    <p class="problem-solved">Sorun Giderildi.</p>
                    <p>{customer_name}</p>
                </div>

                <div class="technician-section">
                    <p>Montör Adı</p>
                </div>
            </div>
        </div>
        """

        html_content += """
        </body>
        </html>
        """

        self.text_edit.setHtml(html_content)

    def print_data(self):
        """A4 kağıdına tam sığacak şekilde optimize yazdırma"""
        # Önce Parça Durumu'nu "Çözüldü" olarak güncelle
        self.update_parca_durumu("Çözüldü")

        printer = QPrinter(QPrinter.HighResolution)

        # A4 boyutunu ve kenar boşluklarını ayarla
        page_layout = QPageLayout(
            QPageSize(QPageSize.A4),
            QPageLayout.Portrait,
            QMarginsF(10, 10, 10, 10),
            QPageLayout.Millimeter
        )
        printer.setPageLayout(page_layout)

        # Yazdırma diyaloğunu aç
        dialog = QPrintDialog(printer, self)
        if dialog.exec_() == QPrintDialog.Accepted:
            painter = QPainter(printer)

            # Doküman boyutlarını al
            doc = self.text_edit.document()
            doc_width = doc.size().width()
            doc_height = doc.size().height()

            # Yazdırılabilir alan boyutlarını al
            page_rect = printer.pageRect(QPrinter.DevicePixel)
            page_width = page_rect.width()
            page_height = page_rect.height()

            # Ölçeklendirme faktörünü hesapla
            x_scale = page_width / doc_width
            y_scale = page_height / doc_height
            scale_factor = min(x_scale, y_scale) * 0.95

            # Ölçeklendirme ve çizim
            painter.save()
            painter.scale(scale_factor, scale_factor)
            doc.drawContents(painter)
            painter.restore()

            painter.end()

    def update_google_sheets_after_print(self):
        """Yazdırma sonrası Google Sheets'i güncelle - Artık kullanılmıyor"""
        pass

    def update_parca_durumu(self, durum_degeri):
        """Seçili satırların Parça Durumu sütununu günceller"""
        try:
            # Service Account ile Google Sheets client'ı al
            config_manager = CentralConfigManager()
            sheets_manager = config_manager.gc

            # PRGsheet dosyasını aç
            spreadsheet = sheets_manager.open("PRGsheet")
            worksheet = spreadsheet.worksheet('Ssh')

            # Tüm veriyi al
            all_values = worksheet.get_all_values()
            headers = all_values[0]

            # Sütun indekslerini bul
            yedek_parca_siparis_col = None
            parca_durumu_col = None

            for idx, header in enumerate(headers):
                if "Yedek Parça Sipariş No" in header:
                    yedek_parca_siparis_col = idx
                elif "Parça Durumu" in header:
                    parca_durumu_col = idx

            if yedek_parca_siparis_col is None or parca_durumu_col is None:
                return

            # selected_rows_data'dan Yedek Parça Sipariş No listesini al
            siparis_no_list = []
            for row_data in self.selected_rows_data:
                siparis_no = row_data.get("Yedek Parça Sipariş No", "")
                if siparis_no:
                    siparis_no_list.append(str(siparis_no).strip())

            # Her satırı kontrol et ve güncelle
            for row_idx, row in enumerate(all_values[1:], start=2):  # 1. satır header, 2'den başla
                if yedek_parca_siparis_col < len(row):
                    cell_value = str(row[yedek_parca_siparis_col]).strip()

                    if cell_value in siparis_no_list:
                        # Parça Durumu sütununu güncelle
                        worksheet.update_cell(row_idx, parca_durumu_col + 1, durum_degeri)

        except Exception as e:
            logger.error(f"Parça Durumu güncelleme hatası: {str(e)}")

    def update_to_montor(self):
        """Seçilen kayıtları 'Sorun Çözüldü' olarak günceller"""
        self.update_parca_durumu("Çözüldü")

        # Font rengini düzelt
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("Bilgi")
        msg.setText("Parça Durumu 'Çözüldü' olarak güncellendi.")
        msg.setStyleSheet("""
            QMessageBox {
                background-color: white;
            }
            QMessageBox QLabel {
                color: black;
                font-size: 14px;
            }
            QMessageBox QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                padding: 5px 15px;
                font-size: 12px;
                min-width: 80px;
            }
        """)
        msg.exec_()


class MontajBelgesiDialog(QDialog):
    """Montaj Belgesi oluşturulmayan sözleşmeleri gösteren dialog"""

    def __init__(self, eslesmeyen_kayitlar, parent=None):
        super().__init__(parent)
        self.eslesmeyen_kayitlar = eslesmeyen_kayitlar
        self.setup_ui()

    def setup_ui(self):
        """UI'yi oluştur"""
        self.setWindowTitle("Montaj Belgesi Oluşturulmayan Sözleşmeler")

        # Light theme background
        self.setStyleSheet("""
            QDialog {
                background-color: #ffffff;
                color: #000000;
            }
            QScrollBar:vertical {
                background-color: #f0f0f0;
                width: 15px;
                border: 1px solid #d0d0d0;
            }
            QScrollBar::handle:vertical {
                background-color: #c0c0c0;
                min-height: 20px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #a0a0a0;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                background-color: #f0f0f0;
                height: 0px;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background-color: #f0f0f0;
            }
            QScrollBar:horizontal {
                background-color: #f0f0f0;
                height: 15px;
                border: 1px solid #d0d0d0;
            }
            QScrollBar::handle:horizontal {
                background-color: #c0c0c0;
                min-width: 20px;
                border-radius: 4px;
            }
            QScrollBar::handle:horizontal:hover {
                background-color: #a0a0a0;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                background-color: #f0f0f0;
                width: 0px;
            }
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                background-color: #f0f0f0;
            }
        """)

        # Ana layout
        layout = QVBoxLayout(self)

        # Başlık ve Excel butonu aynı satırda
        title_layout = QHBoxLayout()
        
        title_label = QLabel("Montaj Belgesi Oluşturulmayan Sözleşmeler")
        title_label.setStyleSheet("""
            QLabel {
                font-size: 18px;
                font-weight: bold;
                color: #000000;
                padding: 10px;
                background-color: #f0f0f0;
                border-radius: 4px;
            }
        """)
        title_layout.addWidget(title_label)
        title_layout.addStretch()
        
        # Excel butonu başlığın yanında
        self.excel_btn = QPushButton("Excel")
        self.excel_btn.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 15px;
                font-size: 12px;
                font-weight: bold;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
            QPushButton:pressed {
                background-color: #909090;
            }
        """)
        self.excel_btn.clicked.connect(self.export_to_excel)
        title_layout.addWidget(self.excel_btn)
        
        layout.addLayout(title_layout)

        # Tablo - Light theme (risk_module.py ile aynı)
        self.table = QTableWidget()
        self.table.setStyleSheet("""
            QTableWidget {
                font-size: 15px;
                font-weight: bold;
                background-color: #ffffff;
                alternate-background-color: #f5f5f5;
                gridline-color: #d0d0d0;
                border: 1px solid #d0d0d0;
                color: #000000;
            }
            QTableWidget::item {
                padding: 5px;
                border-bottom: 1px solid #e0e0e0;
                color: #000000;
            }
            QTableWidget::item:selected {
                background-color: #b3d9ff;
                color: #000000;
            }
            QTableWidget::item:focus {
                outline: none;
                border: none;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                color: #000000;
                padding: 8px;
                border: 1px solid #d0d0d0;
                font-weight: bold;
                font-size: 15px;
            }
            QAbstractScrollArea {
                background-color: #ffffff;
            }
            QAbstractScrollArea > QWidget {
                background-color: #ffffff;
            }
            QAbstractScrollArea::corner {
                background-color: #f0f0f0;
                border: 1px solid #d0d0d0;
            }
            QScrollBar:vertical {
                background-color: #f0f0f0;
                width: 15px;
                border: 1px solid #d0d0d0;
            }
            QScrollBar::handle:vertical {
                background-color: #c0c0c0;
                min-height: 20px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #a0a0a0;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                background-color: #f0f0f0;
                height: 0px;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background-color: #f0f0f0;
            }
            QScrollBar:horizontal {
                background-color: #f0f0f0;
                height: 15px;
                border: 1px solid #d0d0d0;
            }
            QScrollBar::handle:horizontal {
                background-color: #c0c0c0;
                min-width: 20px;
                border-radius: 4px;
            }
            QScrollBar::handle:horizontal:hover {
                background-color: #a0a0a0;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                background-color: #f0f0f0;
                width: 0px;
            }
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                background-color: #f0f0f0;
            }
            QTableCornerButton::section {
                background-color: #f0f0f0;
                border: 1px solid #d0d0d0;
            }
        """)
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.setFocusPolicy(Qt.NoFocus)  # Focus border'ı kaldır (risk_module.py gibi)

        # Verileri tabloya doldur
        if self.eslesmeyen_kayitlar:
            df = pd.DataFrame(self.eslesmeyen_kayitlar)

            # Sütun isimlerini değiştir
            rename_dict = {
                'sip_belgeno': 'Sözleşme No',
                'msg_S_1072': 'Tarih',
                'msg_S_0789': 'Mikro Sip. No',
                'sip_musteri_kod': 'Cari Kod'
            }
            df = df.rename(columns=rename_dict)

            self.table.setRowCount(len(df))
            self.table.setColumnCount(len(df.columns))
            self.table.setHorizontalHeaderLabels(df.columns.tolist())

            for i, row in df.iterrows():
                for j, value in enumerate(row):
                    # Tarih sütunundaki datetime değerlerini formatla
                    column_name = df.columns[j]
                    if column_name == 'Tarih' and pd.notna(value):
                        # Datetime ise sadece tarih kısmını al
                        if isinstance(value, pd.Timestamp) or isinstance(value, str):
                            try:
                                # String ise datetime'a çevir
                                if isinstance(value, str):
                                    dt = pd.to_datetime(value)
                                else:
                                    dt = value
                                # Sadece tarih kısmını göster (YYYY-MM-DD formatında)
                                value = dt.strftime('%Y-%m-%d')
                            except:
                                pass  # Hata olursa orijinal değeri kullan

                    item = QTableWidgetItem(str(value) if pd.notna(value) else "")
                    item.setFlags(item.flags() ^ Qt.ItemIsEditable)  # Make non-editable

                    # Set font properties - risk_module.py ile aynı
                    font = QFont('Segoe UI', 12)
                    font.setBold(True)
                    item.setFont(font)

                    self.table.setItem(i, j, item)

            # Sütun genişliklerini ayarla
            header = self.table.horizontalHeader()
            header.setSectionResizeMode(QHeaderView.Interactive)
            header.setStretchLastSection(False)

            # Minimum sütun genişliklerini ayarla
            for i in range(self.table.columnCount()):
                self.table.setColumnWidth(i, max(150, self.table.columnWidth(i)))

            # Sütunları içeriğe göre boyutlandır
            self.table.resizeColumnsToContents()

            # Satır yüksekliğini ayarla - daha kompakt görünüm
            row_height = 36
            for i in range(self.table.rowCount()):
                self.table.setRowHeight(i, row_height)

        layout.addWidget(self.table)

        # Status bar ekle
        from PyQt5.QtWidgets import QStatusBar
        self.status_bar = QStatusBar()
        self.status_bar.setStyleSheet("""
            QStatusBar {
                background-color: #f5f5f5;
                color: #000000;
                font-size: 14px;
                font-weight: bold;
                border-top: 1px solid #d0d0d0;
            }
        """)
        self.status_bar.showMessage(f"Toplam {len(self.eslesmeyen_kayitlar)} kayıt bulundu")
        layout.addWidget(self.status_bar)

        # Dinamik boyutlandırma
        self.adjust_size_to_content()

    def adjust_size_to_content(self):
        """Dialog boyutunu içeriğe göre ayarla"""
        if not self.eslesmeyen_kayitlar:
            self.setGeometry(100, 100, 800, 400)
            return

        # Tablo genişliğini hesapla
        table_width = 0
        for i in range(self.table.columnCount()):
            table_width += self.table.columnWidth(i)

        # Tablo yüksekliğini hesapla
        row_count = self.table.rowCount()
        row_height = 36  # Satır yüksekliği
        header_height = 40  # Başlık yüksekliği
        table_height = (row_count * row_height) + header_height

        # Başlık, kayıt sayısı ve butonlar için ekstra alan
        extra_height = 180

        # Dialog boyutlarını hesapla - tüm verilerin gözükeceği genişlikte
        # Scroll bar (15px) + border (2px) + padding (60px) = ~80px ekstra alan
        dialog_width = table_width + 80
        dialog_height = table_height + extra_height

        # Ekran boyutunu al
        from PyQt5.QtWidgets import QDesktopWidget
        screen = QDesktopWidget().screenGeometry()

        # Minimum ve maksimum sınırlar - ekran boyutuna göre
        min_width = 800
        max_width = int(screen.width() * 0.95)  # Ekranın %95'i
        min_height = 400
        max_height = 900

        dialog_width = max(min_width, min(dialog_width, max_width))
        dialog_height = max(min_height, min(dialog_height, max_height))

        # Ekran merkezine konumlandır
        x = (screen.width() - dialog_width) // 2
        y = (screen.height() - dialog_height) // 2

        self.setGeometry(x, y, dialog_width, dialog_height)

    def export_to_excel(self):
        """Montaj belgesi olmayan kayıtları Excel'e aktar"""
        try:
            if not self.eslesmeyen_kayitlar:
                self.status_bar.showMessage("❌ Dışa aktarılacak veri yok")
                return

            # DataFrame oluştur
            df = pd.DataFrame(self.eslesmeyen_kayitlar)

            # Sütun isimlerini değiştir
            rename_dict = {
                'sip_belgeno': 'Sözleşme No',
                'msg_S_1072': 'Tarih',
                'msg_S_0789': 'Mikro Sip. No',
                'sip_musteri_kod': 'Cari Kod',
                'Cari Adı': 'Cari Adı'
            }
            df = df.rename(columns=rename_dict)

            # Çıktı dosya yolu
            output_path = "D:/GoogleDrive/~ MontajBelgesiOlmayan.xlsx"

            # Excel'e aktar
            self.status_bar.showMessage("📊 Excel dosyası oluşturuluyor...")
            QApplication.processEvents()

            df.to_excel(output_path, index=False, engine='openpyxl')

            self.status_bar.showMessage(f"✅ {len(df)} kayıt Excel'e aktarıldı: {output_path}")

            # Dosyayı aç
            import os
            os.startfile(output_path)

        except Exception as e:
            error_msg = f"Excel export hatası: {str(e)}"
            logger.error(error_msg)
            self.status_bar.showMessage(f"❌ {error_msg}")


class CustomerInfoWindow(QMainWindow):
    """Müşteri bilgilerini gösteren pencere"""

    def __init__(self, contract_data, contract_id, parent=None):
        super().__init__(parent)
        self.contract_data = contract_data
        self.contract_id = contract_id
        self.setup_ui()

    def setup_ui(self):
        """UI'yi oluştur"""
        self.setWindowTitle(f"Müşteri Bilgileri - Sözleşme: {self.contract_id}")

        # Pencere boyutu
        self.setGeometry(100, 100, 900, 400)

        # Ana widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Layout
        layout = QVBoxLayout(central_widget)

        # Başlık
        title_label = QLabel("MÜŞTERİ BİLGİLERİ")
        title_label.setStyleSheet("""
            QLabel {
                font-size: 20px;
                font-weight: bold;
                color: #2c3e50;
                padding: 10px;
                background-color: #ecf0f1;
                border-radius: 4px;
                margin-bottom: 10px;
            }
        """)
        layout.addWidget(title_label)

        # Contract info'dan müşteri bilgilerini çıkar
        if hasattr(self.contract_data, 'ES_CONTRACT_INFO'):
            contract_info = self.contract_data.ES_CONTRACT_INFO

            def safe_get(obj, attr, default='N/A'):
                """Güvenli attribute alma"""
                if obj is None:
                    return default
                return getattr(obj, attr, default) if hasattr(obj, attr) else default

            # Müşteri bilgileri
            customer_name = f"{safe_get(contract_info, 'CUSTOMER_NAMEFIRST')} {safe_get(contract_info, 'CUSTOMER_NAMELAST')}".strip()

            customer_group = self.create_customer_info_group("MÜŞTERİ BİLGİLERİ", {
                'ad_soyad': customer_name,
                'telefon1': safe_get(contract_info, 'CUSTOMER_PHONE1'),
                'telefon2': safe_get(contract_info, 'CUSTOMER_PHONE2'),
                'vergi_no': safe_get(contract_info, 'CUSTOMER_TAXNR'),
                'sehir_ilce': safe_get(contract_info, 'CUSTOMER_CITY'),
                'adres': safe_get(contract_info, 'CUSTOMER_ADDRESS')
            })
            layout.addWidget(customer_group)
        else:
            error_label = QLabel("Sözleşme bilgileri alınamadı.")
            error_label.setStyleSheet("color: red; font-size: 14px;")
            layout.addWidget(error_label)

        # Kapat butonu
        close_btn = QPushButton("Kapat")
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 15px;
                font-size: 12px;
                font-weight: bold;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
            QPushButton:pressed {
                background-color: #909090;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #888888;
            }
        """)
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn)

    def create_customer_info_group(self, title, info_dict):
        """
        Müşteri bilgilerini gösteren grup kutusu oluşturur.

        Args:
            title (str): Grup başlığı (örn: "MÜŞTERİ BİLGİLERİ")
            info_dict (dict): Müşteri bilgilerini içeren sözlük

        Returns:
            QGroupBox: Grid düzeninde müşteri bilgileri grubu
        """
        group_box = QGroupBox(title)
        group_box.setStyleSheet("""
            QGroupBox {
                font-size: 16px;
                font-weight: bold;
                border: 3px solid #3498db;
                border-radius: 8px;
                margin-top: 15px;
                padding-top: 15px;
                background-color: #f8f9fa;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 5px 10px 5px 10px;
                color: #2c3e50;
                background-color: #3498db;
                color: white;
                border-radius: 4px;
                font-size: 16px;
            }
        """)

        # Ana layout
        main_layout = QVBoxLayout()

        # Grid layout - 3 satır x 2 sütun
        grid_layout = QHBoxLayout()

        # Sol sütun
        left_layout = QVBoxLayout()
        left_items = [
            ("Ad Soyad:", info_dict.get('ad_soyad', '')),
            ("Telefon 1:", info_dict.get('telefon1', '')),
            ("Telefon 2:", info_dict.get('telefon2', ''))
        ]

        for label_text, value in left_items:
            if value and str(value).strip() and str(value) != 'N/A':
                item_layout = QHBoxLayout()

                label = QLabel(label_text)
                label.setStyleSheet("font-weight: bold; color: #34495e; font-size: 16px;")

                value_label = QLabel(str(value))
                value_label.setStyleSheet("color: #2c3e50; font-size: 16px; font-weight: bold;")
                value_label.setWordWrap(True)

                item_layout.addWidget(label)
                item_layout.addWidget(value_label, 1)
                left_layout.addLayout(item_layout)

        # Sağ sütun
        right_layout = QVBoxLayout()
        right_items = [
            ("TCKN No:", info_dict.get('vergi_no', '')),
            ("Şehir:", info_dict.get('sehir_ilce', ''))
        ]

        for label_text, value in right_items:
            if value and str(value).strip() and str(value) != 'N/A':
                item_layout = QHBoxLayout()

                label = QLabel(label_text)
                label.setStyleSheet("font-weight: bold; color: #34495e; font-size: 16px;")

                value_label = QLabel(str(value))
                value_label.setStyleSheet("color: #2c3e50; font-size: 16px; font-weight: bold;")
                value_label.setWordWrap(True)

                item_layout.addWidget(label)
                item_layout.addWidget(value_label, 1)
                right_layout.addLayout(item_layout)

        # Grid'e sütunları ekle
        grid_layout.addLayout(left_layout)
        grid_layout.addSpacing(20)  # Sütunlar arası boşluk
        grid_layout.addLayout(right_layout)

        main_layout.addLayout(grid_layout)

        # Adres bilgisini en alta ekle (tam genişlikte)
        adres = info_dict.get('adres', '')
        if adres and str(adres).strip() and str(adres) != 'N/A':
            adres_layout = QHBoxLayout()

            adres_label = QLabel("Adres:")
            adres_label.setStyleSheet("font-weight: bold; color: #34495e; font-size: 16px;")

            adres_value = QLabel(str(adres))
            adres_value.setStyleSheet("color: #2c3e50; font-size: 16px; font-weight: bold;")
            adres_value.setWordWrap(True)

            adres_layout.addWidget(adres_label)
            adres_layout.addWidget(adres_value, 1)

            main_layout.addLayout(adres_layout)

        group_box.setLayout(main_layout)
        return group_box

