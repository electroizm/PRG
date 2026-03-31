import os
import sys
import logging
from pathlib import Path
from io import BytesIO
from datetime import datetime
from typing import Optional

import pandas as pd
import requests

# Üst dizini Python path'e ekle (central_config için)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Central config import
from central_config import CentralConfigManager # type: ignore

from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, 
                             QAbstractItemView, QMenu, QProgressBar, QLabel, QApplication, QShortcut,
                             QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox, QFrame)
from PyQt5.QtGui import QFont, QKeySequence

# Logger
logger = logging.getLogger(__name__)


# ================== CONFIG CONSTANTS ==================
# Paths
MIKRO_EXE_PATH = Path("D:/GoogleDrive/PRG/EXE/Risk.exe")
EXPORT_PATH = Path("D:/GoogleDrive/~ Risk_Export.xlsx")

# Cache
CACHE_KEY_RISK = "Risk"
SHEET_NAME_RISK = "Risk"

# Timing (milliseconds)
LAZY_LOAD_DELAY_MS = 100
MIKRO_EXECUTION_TIMEOUT_MS = 7000
SHEETS_UPDATE_DELAY_MS = 5000
PROGRESS_BAR_HIDE_DELAY_MS = 1000

# Network
REQUEST_TIMEOUT_SEC = 30

# UI
MIN_COLUMN_WIDTH = 150
ROW_HEIGHT = 35
FONT_FAMILY = "Segoe UI"
FONT_SIZE = 12


# ================== STYLESHEET CONSTANTS ==================
BUTTON_STYLE = """
    QPushButton {
        background-color: #dfdfdf;
        color: black;
        border: 1px solid #444;
        padding: 8px 16px;
        border-radius: 5px;
        font-size: 14px;
        font-weight: bold;
        min-width: 80px;
    }
    QPushButton:hover {
        background-color: #a0a5a2;
    }
"""

TOTAL_RISK_LABEL_STYLE = """
    QLabel {
        color: #d0d0d0;
        font-size: 16px;
        font-weight: bold;
        padding: 8px;
        background-color: #f8f8f8;
        border: 1px solid #dddddd;
        border-radius: 3px;
        margin: 2px;
    }
    QLabel:hover {
        color: #000000;
    }
"""


TABLE_STYLE = """
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
"""


CONTEXT_MENU_STYLE = """
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
"""


# ================== DATA LOADER THREAD ==================
class DataLoaderThread(QThread):
    """Arka planda veri yükleme için thread"""

    # Signals
    progress_updated = pyqtSignal(int, str)  # (progress_value, status_message)
    data_loaded = pyqtSignal(pd.DataFrame)  # Yüklenen DataFrame
    error_occurred = pyqtSignal(str)  # Hata mesajı

    def __init__(self, gsheets_url: str):
        super().__init__()
        self.gsheets_url = gsheets_url
        self.is_cancelled = False

    def run(self):
        """Thread ana işlevi - arka planda veri yükleme"""
        try:
            if not self.gsheets_url:
                self.error_occurred.emit("PRGsheet/Ayar sayfasında SPREADSHEET_ID bulunamadı")
                return

            self.progress_updated.emit(10, "🔗 Google Sheets'e bağlanıyor...")

            # URL'den Excel dosyasını oku
            response = requests.get(
                self.gsheets_url,
                timeout=REQUEST_TIMEOUT_SEC,
                verify=True
            )

            if self.is_cancelled:
                return

            self.progress_updated.emit(30, "✅ Google Sheets'e bağlantı başarılı")

            # Status code kontrolü
            if response.status_code == 401:
                self.error_occurred.emit("Google Sheets erişim hatası: Dosya özel veya izin gerekli")
                return
            elif response.status_code != 200:
                self.error_occurred.emit(f"HTTP Hatası: {response.status_code} - {response.reason}")
                return

            response.raise_for_status()

            if self.is_cancelled:
                return

            self.progress_updated.emit(50, "📋 Risk sayfası yükleniyor...")

            # Risk sayfasını oku
            df = pd.read_excel(BytesIO(response.content), sheet_name=SHEET_NAME_RISK)

            if self.is_cancelled:
                return

            self.progress_updated.emit(70, "🔄 Veriler işleniyor...")

            # DataFrame işleme
            df = self._reorder_columns(df)
            df = self._format_date_columns(df)

            self.progress_updated.emit(90, "📋 Tablo dolduruluyor...")

            # Veri yüklendi sinyali
            self.data_loaded.emit(df)

            self.progress_updated.emit(100, f"✅ {len(df)} kayıt başarıyla yüklendi (Risk sayfası)")

        except requests.exceptions.Timeout:
            self.error_occurred.emit("Bağlantı zaman aşımı - Google Sheets'e erişilemiyor")
        except requests.exceptions.RequestException as e:
            self.error_occurred.emit(f"Bağlantı hatası: {str(e)}")
        except Exception as e:
            logger.exception("Veri yükleme hatası")
            self.error_occurred.emit(f"Veri yükleme hatası: {str(e)}")

    def cancel(self):
        """Thread'i iptal et"""
        self.is_cancelled = True

    @staticmethod
    def _reorder_columns(df: pd.DataFrame) -> pd.DataFrame:
        """
        Sütunları şu sıraya göre düzenle: Cari hesap adı -> Risk -> Cari hesap adı 2 -> diğer

        Args:
            df: İşlenecek DataFrame

        Returns:
            Sütunları yeniden sıralanmış DataFrame
        """
        if df.empty:
            return df

        columns = list(df.columns)
        new_order = []

        # İlk olarak "Cari hesap adı" sütununu ekle 
        cari_adi_cols = [col for col in columns
                        if 'cari hesap adı' in col.lower() and '2' not in col.lower()]
        new_order.extend(cari_adi_cols)

        # Sonra "Risk" sütununu ekle
        risk_cols = [col for col in columns if 'risk' in col.lower()]
        new_order.extend(risk_cols)

        # Sonra "Cari hesap adı 2" sütununu ekle
        cari_adi2_cols = [col for col in columns if 'cari hesap adı 2' in col.lower()]
        new_order.extend(cari_adi2_cols)

        # Kalan sütunları ekle
        remaining_cols = [col for col in columns if col not in new_order]
        new_order.extend(remaining_cols)

        return df[new_order]

    @staticmethod
    def _format_date_columns(df: pd.DataFrame) -> pd.DataFrame:
        """
        Tarih sütunlarını 'YYYY-MM-DD' formatına çevir

        Args:
            df: İşlenecek DataFrame

        Returns:
            Tarih sütunları formatlanmış DataFrame
        """
        for col in df.columns:
            if 'tarih' in col.lower():
                try:
                    df[col] = pd.to_datetime(df[col], errors='coerce')
                    df[col] = df[col].dt.strftime('%Y-%m-%d')
                except Exception as e:
                    logger.warning(f"Tarih formatlama hatası ({col}): {e}")
        return df


# ================== ANA UYGULAMA ==================
class RiskApp(QWidget):
    """Risk analizi ana uygulama widget'ı"""

    def __init__(self):
        super().__init__()
        self.veri_cercevesi = pd.DataFrame()
        self.mikro_calisiyor = False
        self.gsheets_url = self._load_gsheets_url()
        self.data_loader_thread: Optional[DataLoaderThread] = None

        # Lazy loading için flag
        self._data_loaded = False

        # Cache'i başta yükle
        self._cache = None
        try:
            if 'main' in sys.modules:
                from main import GlobalDataCache
                self._cache = GlobalDataCache()
        except Exception:
            pass

        self.setup_ui()
        self.setup_connections()

    def showEvent(self, event):
        """Widget ilk gösterildiğinde veri yükle (lazy loading)"""
        super().showEvent(event)
        if not self._data_loaded:
            self._data_loaded = True
            # UI render olduktan sonra veri yükle
            QTimer.singleShot(LAZY_LOAD_DELAY_MS,
                            lambda: self.load_data(force_reload=False))

    def _load_gsheets_url(self) -> Optional[str]:
        """
        Google Sheets SPREADSHEET_ID'sini yükle - Service Account

        Returns:
            Google Sheets export URL veya None
        """
        try:
            config_manager = CentralConfigManager()
            spreadsheet_id = config_manager.MASTER_SPREADSHEET_ID
            return f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"
        except Exception as e:
            logger.error(f"PRGsheet yüklenirken hata: {e}")
            if hasattr(self, 'status_label'):
                self.status_label.setText(f"❌ PRGsheet yüklenirken hata: {str(e)}")
            return None

    # ================== UI SETUP ==================
    def setup_ui(self):
        """UI'ı oluştur"""
        self._setup_widget_style()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Header (butonlar + toplam risk)
        header_widget = self._create_header()

        # Progress Bar
        self.progress_bar = self._create_progress_bar()

        # Table
        self.table = self._create_table()

        # Status Bar
        status_widget = self._create_status_bar()

        # Layout'a ekle
        layout.addWidget(header_widget)
        layout.addWidget(self.table, 1)
        layout.addWidget(status_widget)

    def _setup_widget_style(self):
        """Widget arka plan stilini ayarla"""
        self.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                color: #000000;
            }
        """)

    def _create_header(self) -> QWidget:
        """Header widget'ını oluştur (butonlar + toplam risk)"""
        header_layout = QHBoxLayout()

        # Butonlar
        self.mikro_button = QPushButton("Mikro")
        self.mikro_button.setStyleSheet(BUTTON_STYLE)

        self.refresh_button = QPushButton("Yenile")
        self.refresh_button.setStyleSheet(BUTTON_STYLE)

        self.export_button = QPushButton("Excel")
        self.export_button.setStyleSheet(BUTTON_STYLE)

        # Toplam Risk Label
        self.total_risk_label = QLabel("Toplam Risk: 0 ₺")
        self.total_risk_label.setStyleSheet(TOTAL_RISK_LABEL_STYLE)

        header_layout.addWidget(self.mikro_button)
        header_layout.addWidget(self.refresh_button)
        header_layout.addWidget(self.export_button)
        header_layout.addStretch()
        header_layout.addWidget(self.total_risk_label)

        # Widget olarak sar
        header_widget = QWidget()
        header_widget.setLayout(header_layout)
        header_widget.setStyleSheet("""
            background-color: #ffffff;
            margin-bottom: 0px;
        """)
        header_layout.setContentsMargins(10, 10, 10, 10)

        return header_widget

    def _create_progress_bar(self) -> QProgressBar:
        """Progress bar oluştur"""
        progress_bar = QProgressBar()
        progress_bar.setVisible(False)
        progress_bar.setTextVisible(True)
        progress_bar.setAlignment(Qt.AlignCenter)
        progress_bar.setFormat("%p%")
        progress_bar.setStyleSheet("""
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
        return progress_bar

    def _create_table(self) -> QTableWidget:
        """Tablo widget'ını oluştur"""
        table = QTableWidget()
        table.setContextMenuPolicy(Qt.CustomContextMenu)
        table.setStyleSheet(TABLE_STYLE)
        table.setAlternatingRowColors(True)
        table.setShowGrid(True)
        return table

    def _create_status_bar(self) -> QWidget:
        """Status bar widget'ını oluştur"""
        status_layout = QHBoxLayout()

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

        status_layout.addWidget(self.status_label, 3)
        status_layout.addWidget(self.progress_bar, 1)
        status_layout.setContentsMargins(0, 0, 0, 0)

        status_widget = QWidget()
        status_widget.setLayout(status_layout)
        status_widget.setStyleSheet("background-color: #f5f5f5; border-top: 1px solid #d0d0d0;")

        return status_widget

    def setup_connections(self):
        """Signal-slot bağlantılarını kur"""
        self.mikro_button.clicked.connect(self.run_mikro)
        self.refresh_button.clicked.connect(lambda: self.load_data(force_reload=True))
        self.export_button.clicked.connect(self.export_to_excel)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        
        # Ctrl+C kısayolu - self üzerine bağlanır ki focus nerede olursa olsun çalışsın
        self.copy_shortcut = QShortcut(QKeySequence("Ctrl+C"), self)
        self.copy_shortcut.setContext(Qt.WindowShortcut)
        self.copy_shortcut.activated.connect(self.handle_ctrl_c)

    # ================== DATA LOADING ==================
    def load_data(self, force_reload: bool = False):
        """
        Risk sayfasından verileri yükle (cache-aware)

        Args:
            force_reload: True ise cache'i bypass et, Google Sheets'ten çek
        """
        # Cache kontrolü (force_reload değilse)
        if not force_reload and self._try_load_from_cache():
            return

        # Cache yoksa veya force_reload ise: Google Sheets'ten çek (thread ile)
        self._load_from_sheets()

    def _try_load_from_cache(self) -> bool:
        """
        Cache'den veri yüklemeyi dene

        Returns:
            True ise cache'den yüklendi, False ise cache yok
        """
        try:
            # Cache kontrolü
            if not self._cache:
                return False

            if not self._cache.has(CACHE_KEY_RISK):
                return False

            # Cache'den yükle
            self.veri_cercevesi = self._cache.get(CACHE_KEY_RISK)

            # Sütun sıralaması ve tarih formatlama
            if not self.veri_cercevesi.empty:
                self.veri_cercevesi = DataLoaderThread._reorder_columns(self.veri_cercevesi)
                self.veri_cercevesi = DataLoaderThread._format_date_columns(self.veri_cercevesi)

            self.populate_table()
            self.update_total_risk()
            self.status_label.setText(
                f"✅ {len(self.veri_cercevesi)} kayıt yüklendi (Cache'den - anında)"
            )
            return True

        except Exception as e:
            logger.exception("Cache'den veri yükleme hatası")
            return False

    def _load_from_sheets(self):
        """Google Sheets'ten veri yükle (QThread ile arka planda)"""
        # Eğer çalışan bir thread varsa iptal et
        if self.data_loader_thread and self.data_loader_thread.isRunning():
            self.data_loader_thread.cancel()
            self.data_loader_thread.wait()

        # Progress bar'ı göster
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.status_label.setText("📊 Risk sayfasından veriler yükleniyor...")
        self.set_buttons_enabled(False)

        # Thread oluştur ve başlat
        self.data_loader_thread = DataLoaderThread(self.gsheets_url)
        self.data_loader_thread.progress_updated.connect(self._on_progress_updated)
        self.data_loader_thread.data_loaded.connect(self._on_data_loaded)
        self.data_loader_thread.error_occurred.connect(self._on_error_occurred)
        self.data_loader_thread.finished.connect(self._on_thread_finished)
        self.data_loader_thread.start()

    def _on_progress_updated(self, progress: int, message: str):
        """Progress güncellemesi"""
        self.progress_bar.setValue(progress)
        self.status_label.setText(message)

    def _on_data_loaded(self, df: pd.DataFrame):
        """Veri yükleme başarılı"""
        self.veri_cercevesi = df
        self.populate_table()
        self.update_total_risk()

        # Cache'e kaydet
        try:
            if self._cache:
                self._cache.set(CACHE_KEY_RISK, self.veri_cercevesi)
        except Exception as e:
            logger.warning(f"Cache'e kaydetme hatası: {e}")

    def _on_error_occurred(self, error_message: str):
        """Hata oluştu"""
        self.veri_cercevesi = pd.DataFrame()
        self.populate_table()
        self.progress_bar.setVisible(False)
        self.status_label.setText(f"❌ {error_message}")

    def _on_thread_finished(self):
        """Thread tamamlandı"""
        self.set_buttons_enabled(True)
        # Progress bar'ı 1 saniye sonra gizle
        QTimer.singleShot(PROGRESS_BAR_HIDE_DELAY_MS,
                         lambda: self.progress_bar.setVisible(False))

    # ================== TABLE OPERATIONS ==================
    def populate_table(self):
        """
        Tabloyu verilerle doldur (optimized)

        Performance improvements:
        - setUpdatesEnabled(False) kullanımı
        - Batch processing hazır
        """
        if self.veri_cercevesi.empty:
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            return

        # Performans: UI güncellemelerini durdur
        self.table.setUpdatesEnabled(False)
        self.table.setSortingEnabled(False)

        try:
            self.table.setRowCount(len(self.veri_cercevesi))
            self.table.setColumnCount(len(self.veri_cercevesi.columns))
            self.table.setHorizontalHeaderLabels(self.veri_cercevesi.columns.tolist())

            # Tablo özellikleri
            self.table.setAlternatingRowColors(True)
            self.table.setSelectionBehavior(QAbstractItemView.SelectItems)
            self.table.setSelectionMode(QAbstractItemView.SingleSelection)
            self.table.setFocusPolicy(Qt.NoFocus)

            # Tablo doldur
            for i, row in self.veri_cercevesi.iterrows():
                for j, value in enumerate(row):
                    item = self._create_table_item(value, j)
                    self.table.setItem(i, j, item)

            # Header styling
            self._configure_table_header()

            # Satır yüksekliği
            for i in range(self.table.rowCount()):
                self.table.setRowHeight(i, ROW_HEIGHT)

        finally:
            # Performans: UI güncellemelerini tekrar aç
            self.table.setUpdatesEnabled(True)
            self.table.setSortingEnabled(True)

    @staticmethod
    def _format_phone_number(value) -> str:
        """
        Telefon numarasını formatla (float → int → str)

        Args:
            value: Telefon numarası değeri

        Returns:
            Formatlanmış telefon numarası string'i
        """
        try:
            return str(int(float(value)))
        except (ValueError, TypeError):
            return str(value)

    def _create_table_item(self, value, column_index: int) -> QTableWidgetItem:
        """
        Tablo item'ı oluştur (formatlama ve renklendirme ile)

        Args:
            value: Hücre değeri
            column_index: Sütun indeksi

        Returns:
            QTableWidgetItem
        """
        # NaN değerlerini boş string yap
        if pd.isna(value) or str(value).lower() == 'nan':
            display_value = ""
        elif column_index < len(self.veri_cercevesi.columns) and \
             'telefon' in self.veri_cercevesi.columns[column_index].lower():
            # Telefon sütunu için özel formatlama
            display_value = self._format_phone_number(value)
        else:
            display_value = str(value)

        item = QTableWidgetItem(display_value)
        item.setFlags(item.flags() ^ Qt.ItemIsEditable)  # Non-editable

        # Font
        font = QFont(FONT_FAMILY, FONT_SIZE)
        font.setBold(True)
        item.setFont(font)

        return item

    def _configure_table_header(self):
        """Tablo header'ını yapılandır"""
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setStretchLastSection(False)

        # Minimum sütun genişlikleri
        for i in range(self.table.columnCount()):
            self.table.setColumnWidth(i, max(MIN_COLUMN_WIDTH, self.table.columnWidth(i)))

        # İçeriğe göre boyutlandır
        self.table.resizeColumnsToContents()

    def update_total_risk(self):
        """Toplam risk hesapla ve güncelle"""
        if not self.veri_cercevesi.empty:
            # Risk sütunu ara
            risk_columns = [col for col in self.veri_cercevesi.columns if 'risk' in col.lower()]
            if risk_columns:
                try:
                    total_risk = self.veri_cercevesi[risk_columns[0]].astype(float).sum()
                    self.total_risk_label.setText(f"Toplam Risk: {total_risk:,.0f} ₺")
                except (ValueError, TypeError) as e:
                    logger.warning(f"Risk hesaplama hatası: {e}")
                    self.total_risk_label.setText("Toplam Risk: Hesaplanamadı")
            else:
                self.total_risk_label.setText("Toplam Risk: Risk sütunu bulunamadı")
        else:
            self.total_risk_label.setText("Toplam Risk: 0 ₺")

    # ================== MIKRO EXECUTION ==================
    def run_mikro(self):
        """Risk.exe dosyasını çalıştır"""
        try:
            exe_path = MIKRO_EXE_PATH

            # Path kontrolü
            if not exe_path.exists():
                self.status_label.setText(f"❌ Risk.exe bulunamadı: {exe_path}")
                return

            if not exe_path.is_file():
                self.status_label.setText(f"❌ Risk.exe bir dosya değil: {exe_path}")
                return

            if not str(exe_path).lower().endswith('.exe'):
                self.status_label.setText(f"❌ Geçersiz dosya türü: {exe_path}")
                return

            self.status_label.setText("🔄 Risk.exe çalıştırılıyor...")
            self.mikro_button.setEnabled(False)
            self.mikro_calisiyor = True

            os.startfile(str(exe_path))

            # Risk.exe'nin çalışması için bekleme
            QTimer.singleShot(MIKRO_EXECUTION_TIMEOUT_MS, self.on_mikro_finished)

        except Exception as e:
            logger.exception("Program çalıştırma hatası")
            self.status_label.setText(f"❌ Program çalıştırma hatası: {str(e)}")
            self.mikro_button.setEnabled(True)
            self.mikro_calisiyor = False

    def on_mikro_finished(self):
        """Mikro program bittikten sonra"""
        self.mikro_button.setEnabled(True)
        self.mikro_calisiyor = False
        self.status_label.setText("✅ Risk.exe tamamlandı, Google Sheets güncelleme bekleniyor...")

        # Google Sheets'e kaydedilmesi için ek bekleme, sonra verileri yenile
        QTimer.singleShot(SHEETS_UPDATE_DELAY_MS, self.delayed_data_refresh)

    def delayed_data_refresh(self):
        """Gecikmeli veri yenileme"""
        self.status_label.setText("🔄 Google Sheets'ten güncel veriler alınıyor...")
        self.load_data(force_reload=True)

    # ================== EXPORT ==================
    def export_to_excel(self):
        """Verileri Excel'e aktar"""
        if self.veri_cercevesi.empty:
            self.status_label.setText("⚠️ Dışa aktarılacak veri yok")
            return

        try:
            output_path = EXPORT_PATH
            self.veri_cercevesi.to_excel(str(output_path), index=False, engine='openpyxl')
            self.status_label.setText(f"✅ Veriler dışa aktarıldı: {output_path}")
            logger.info(f"Excel export başarılı: {output_path}")
        except Exception as e:
            logger.exception("Excel export hatası")
            self.status_label.setText(f"❌ Dışa aktarma hatası: {str(e)}")

    # ================== CONTEXT MENU ==================
    def show_context_menu(self, position):
        """Sağ tık menüsü - Sadece hücre kopyalama"""
        item = self.table.itemAt(position)
        if not item:
            return

        menu = QMenu(self)
        menu.setStyleSheet(CONTEXT_MENU_STYLE)

        copy_action = menu.addAction("Kopyala")

        action = menu.exec_(self.table.viewport().mapToGlobal(position))

        if action == copy_action:
            self.copy_cell(item)

    def copy_cell(self, item: QTableWidgetItem):
        """Tıklanan hücreyi kopyala"""
        if item and item.text():
            QApplication.clipboard().setText(item.text())
            old_text = self.status_label.text()
            self.status_label.setText("✅ Kopyalandı")
            QTimer.singleShot(1500, lambda t=old_text: self.status_label.setText(t))
        else:
            self.status_label.setText("⚠️ Boş hücre")

    def handle_ctrl_c(self):
        """Ctrl+C ile kopyalama işlemi"""
        item = self.table.currentItem()
        if item:
            self.copy_cell(item)

    # ================== UTILITY ==================
    def set_buttons_enabled(self, enabled: bool):
        """Butonları aktif/pasif yap"""
        self.mikro_button.setEnabled(enabled)
        self.refresh_button.setEnabled(enabled)
        self.export_button.setEnabled(enabled)
