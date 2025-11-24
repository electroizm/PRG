"""
OKC YazarKasa Modülü - Optimized Version
"""

import os
import sys
import logging
from pathlib import Path
from io import BytesIO
from typing import Optional

import pandas as pd
import requests

# Üst dizini Python path'e ekle (central_config için)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Central config import
from central_config import CentralConfigManager # pyright: ignore[reportMissingImports]

from PyQt5.QtCore import Qt, QTimer, pyqtSignal
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLineEdit,
                             QLabel, QPushButton, QTableWidget, QTableWidgetItem,
                             QHeaderView, QAbstractItemView, QMenu, QProgressBar,
                             QMessageBox, QDialog, QDialogButtonBox, QApplication)
from PyQt5.QtGui import QIntValidator, QFont

# Logger
logging.basicConfig(level=logging.WARNING)
logger = logging.getLogger(__name__)


# ================== CONFIG CONSTANTS ==================
# Paths
OKC_EXE_PATH = Path("D:/GoogleDrive/PRG/EXE/OKC.exe")

# Sheet Names
SHEET_NAME_OKC = "OKC"

# Timing (milliseconds)
LAZY_LOAD_DELAY_MS = 100
OKC_EXECUTION_TIMEOUT_MS = 7000
SHEETS_UPDATE_DELAY_MS = 5000
PROGRESS_BAR_HIDE_DELAY_MS = 1000

# Network
REQUEST_TIMEOUT_SEC = 30

# UI
MIN_COLUMN_WIDTH = 150
ROW_HEIGHT = 35
FONT_FAMILY = "Segoe UI"
FONT_SIZE = 12

# Filter
FILTER_MULTIPLIER = 1000  # Bin TL çarpanı


# ================== STYLESHEET CONSTANTS ==================
WIDGET_STYLE = """
    QWidget {
        background-color: #ffffff;
        color: #000000;
    }
"""

SEARCH_INPUT_STYLE = """
    QLineEdit {
        font-size: 16px;
        padding: 8px;
        border-radius: 5px;
        border: 2px solid #cccccc;
        font-weight: bold;
        min-width: 50px;
        max-width: 300px;
        background-color: #ffffff;
        color: #000000;
    }
"""

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

SEARCH_WIDGET_STYLE = """
    background-color: #ffffff;
    margin-bottom: 0px;
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

DIALOG_STYLE = """
    QDialog {
        background-color: #ffffff;
    }
"""

DIALOG_VALUE_STYLE = """
    font-size: 48px;
    font-weight: bold;
    color: #000000;
    background-color: transparent;
"""

DIALOG_WARNING_STYLE = """
    font-size: 24px;
    color: #e67e22;
    font-weight: bold;
    margin-top: 20px;
    margin-bottom: 20px;
"""

DIALOG_CANCEL_BTN_STYLE = """
    QPushButton {
        background-color: #95a5a6;
        color: white;
        padding: 18px 35px;
        font-size: 24px;
        border-radius: 8px;
        min-width: 150px;
    }
    QPushButton:hover {
        background-color: #7f8c8d;
    }
"""

DIALOG_CONFIRM_BTN_STYLE = """
    QPushButton {
        background-color: #27ae60;
        color: white;
        padding: 18px 35px;
        font-size: 24px;
        border-radius: 8px;
        min-width: 150px;
    }
    QPushButton:hover {
        background-color: #219653;
    }
"""


# ================== ANA UYGULAMA ==================
class OKCYazarKasaApp(QWidget):
    """OKC YazarKasa ana uygulama widget'ı"""

    def __init__(self):
        super().__init__()
        self.veri_cercevesi = pd.DataFrame()
        self.mikro_calisiyor = False
        self.gsheets_url = self._load_gsheets_url()
        self.full_df: Optional[pd.DataFrame] = None
        self.original_df: Optional[pd.DataFrame] = None
        self.current_df: Optional[pd.DataFrame] = None

        # Lazy loading için flag
        self._data_loaded = False

        self.setup_ui()
        self.setup_connections()

    def showEvent(self, event):
        """Widget ilk gösterildiğinde veri yükle (lazy loading)"""
        super().showEvent(event)
        if not self._data_loaded:
            self._data_loaded = True
            # UI render olduktan sonra veri yükle
            QTimer.singleShot(LAZY_LOAD_DELAY_MS, self.load_data)

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

        # Search Widget (Fatura Tutarı + Temizle + e-Arşiv)
        search_widget = self._create_search_widget()

        # Progress Bar
        self.progress_bar = self._create_progress_bar()

        # Table
        self.table = self._create_table()

        # Status Bar
        status_widget = self._create_status_bar()

        # Layout'a ekle
        layout.addWidget(search_widget)
        layout.addWidget(self.table, 1)
        layout.addWidget(status_widget)

    def _setup_widget_style(self):
        """Widget arka plan stilini ayarla"""
        self.setStyleSheet(WIDGET_STYLE)

    def _create_search_widget(self) -> QWidget:
        """Arama widget'ını oluştur (Fatura Tutarı + Temizle + e-Arşiv)"""
        search_layout = QHBoxLayout()

        # Fatura Tutarı etiketi
        fatura_label = QLabel("Fatura Tutarı :")
        fatura_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #000000;")
        search_layout.addWidget(fatura_label)

        # Arama Kutusu
        self.search_input = QLineEdit()
        self.search_input.setValidator(QIntValidator())
        self.search_input.textChanged.connect(self.filter_data)
        self.search_input.setStyleSheet(SEARCH_INPUT_STYLE)
        search_layout.addWidget(self.search_input, 1)

        # Temizle butonu
        self.clear_btn = QPushButton("Temizle")
        self.clear_btn.setStyleSheet(BUTTON_STYLE)
        self.clear_btn.clicked.connect(self.clear_search)
        search_layout.addWidget(self.clear_btn)

        # e-Arşiv butonu
        self.e_arsiv_btn = QPushButton("e-Arşiv")
        self.e_arsiv_btn.setStyleSheet(BUTTON_STYLE)
        self.e_arsiv_btn.clicked.connect(self.run_e_arsiv)
        search_layout.addWidget(self.e_arsiv_btn)

        # Sağa doğru esnek boşluk
        search_layout.addStretch()

        # Widget olarak sar
        search_widget = QWidget()
        search_widget.setLayout(search_layout)
        search_widget.setStyleSheet(SEARCH_WIDGET_STYLE)
        search_layout.setContentsMargins(10, 10, 10, 10)

        return search_widget

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
        table.cellDoubleClicked.connect(self.on_row_double_click)
        table.customContextMenuRequested.connect(self.show_context_menu)
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
        pass

    # ================== DATA LOADING ==================
    def load_data(self):
        """OKC sayfasından verileri yükle"""
        try:
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(0)
            self.status_label.setText("📊 OKC sayfasından veriler yükleniyor...")
            self.set_buttons_enabled(False)

            QApplication.processEvents()

            if not self.gsheets_url:
                self.veri_cercevesi = pd.DataFrame()
                self.populate_table()
                self.progress_bar.setVisible(False)
                self.status_label.setText("❌ PRGsheet/Ayar sayfasında SPREADSHEET_ID bulunamadı")
                return

            self.progress_bar.setValue(10)
            self.status_label.setText("🔗 Google Sheets'e bağlanıyor...")
            QApplication.processEvents()

            # URL'den Excel dosyasını oku
            response = requests.get(self.gsheets_url, timeout=REQUEST_TIMEOUT_SEC, verify=True)

            if response.status_code == 401:
                self.veri_cercevesi = pd.DataFrame()
                self.populate_table()
                self.progress_bar.setVisible(False)
                self.status_label.setText("❌ Google Sheets erişim hatası: Dosya özel veya izin gerekli")
                return
            elif response.status_code != 200:
                self.veri_cercevesi = pd.DataFrame()
                self.populate_table()
                self.progress_bar.setVisible(False)
                self.status_label.setText(f"❌ HTTP Hatası: {response.status_code} - {response.reason}")
                return

            response.raise_for_status()

            self.progress_bar.setValue(30)
            self.status_label.setText("📥 Excel dosyası indiriliyor...")
            QApplication.processEvents()

            # OKC sayfasını oku
            self.full_df = pd.read_excel(BytesIO(response.content), sheet_name=SHEET_NAME_OKC)

            self.progress_bar.setValue(50)
            self.status_label.setText("🔍 Veriler işleniyor...")
            QApplication.processEvents()

            # Orijinal index'leri yeni sütuna kaydet
            self.full_df['_original_index_'] = self.full_df.index

            # Filtreleme yaparken bu sütunu koru
            if 'YazarKasa' in self.full_df.columns:
                self.original_df = self.full_df[self.full_df['YazarKasa'] != 'OK'].copy()
            else:
                self.original_df = self.full_df.copy()

            self.progress_bar.setValue(70)
            self.status_label.setText("📊 Veriler sıralanıyor...")
            QApplication.processEvents()

            # Sıralama yap - Önce tarihe göre (yeni olanlar üstte), sonra tutara göre
            if 'Fatura Düzenlenme Tarihi' in self.original_df.columns and 'Ödenecek Tutar' in self.original_df.columns:
                self.original_df = self.original_df.sort_values(
                    ['Fatura Düzenlenme Tarihi', 'Ödenecek Tutar'],
                    ascending=[False, True]
                )

            self.progress_bar.setValue(90)
            self.status_label.setText("📋 Tablo dolduruluyor...")
            QApplication.processEvents()

            self.veri_cercevesi = self.original_df.copy()
            self.populate_table()

            self.progress_bar.setValue(100)
            self.status_label.setText(f"✅ {len(self.veri_cercevesi)} kayıt başarıyla yüklendi (OKC sayfası)")

            # Progress bar'ı 1 saniye sonra gizle
            QTimer.singleShot(PROGRESS_BAR_HIDE_DELAY_MS, lambda: self.progress_bar.setVisible(False))

        except requests.exceptions.Timeout:
            self.veri_cercevesi = pd.DataFrame()
            self.populate_table()
            self.status_label.setText("❌ Bağlantı zaman aşımı - Google Sheets'e erişilemiyor")
        except requests.exceptions.RequestException as e:
            self.veri_cercevesi = pd.DataFrame()
            self.populate_table()
            self.status_label.setText(f"❌ Bağlantı hatası: {str(e)}")
        except Exception as e:
            logger.exception("Veri yükleme hatası")
            self.veri_cercevesi = pd.DataFrame()
            self.populate_table()
            self.status_label.setText(f"❌ Veri yükleme hatası: {str(e)}")
        finally:
            self.progress_bar.setVisible(False)
            self.set_buttons_enabled(True)

    # ================== TABLE OPERATIONS ==================
    def populate_table(self):
        """
        Tabloyu verilerle doldur (optimized)

        Performance improvements:
        - setUpdatesEnabled(False) kullanımı
        - setSortingEnabled yönetimi
        """
        if self.veri_cercevesi.empty:
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            return

        # Performans: UI güncellemelerini durdur
        self.table.setUpdatesEnabled(False)
        self.table.setSortingEnabled(False)

        try:
            # Görüntülenecek sütunlar (YazarKasa, Alıcı Unvanı / Adı Soyadı ve geçici sütunlar hariç)
            visible_columns = [col for col in self.veri_cercevesi.columns
                            if col not in ['YazarKasa', '_original_index_', 'index', 'Alıcı Unvanı /Adı Soyadı']]

            self.table.setRowCount(len(self.veri_cercevesi))
            self.table.setColumnCount(len(visible_columns))
            self.table.setHorizontalHeaderLabels(visible_columns)

            # Tablo özellikleri
            self.table.setAlternatingRowColors(True)
            self.table.setSelectionBehavior(QAbstractItemView.SelectItems)
            self.table.setSelectionMode(QAbstractItemView.SingleSelection)
            self.table.setFocusPolicy(Qt.NoFocus)

            # Tablo doldur
            for row_idx in range(len(self.veri_cercevesi)):
                row_data = self.veri_cercevesi.iloc[row_idx]

                for col_idx, col_name in enumerate(visible_columns):
                    item = self._create_table_item(row_data, col_name, col_idx)
                    self.table.setItem(row_idx, col_idx, item)

            # Header styling
            self._configure_table_header()

            # Satır yüksekliği
            for i in range(self.table.rowCount()):
                self.table.setRowHeight(i, ROW_HEIGHT)

        finally:
            # Performans: UI güncellemelerini tekrar aç
            self.table.setUpdatesEnabled(True)
            self.table.setSortingEnabled(True)

    def _create_table_item(self, row_data, col_name: str, col_idx: int) -> QTableWidgetItem:
        """
        Tablo item'ı oluştur (formatlama ve renklendirme ile)

        Args:
            row_data: Satır verisi
            col_name: Sütun adı
            col_idx: Sütun indeksi

        Returns:
            QTableWidgetItem
        """
        value = str(row_data[col_name])

        # Özel formatlamalar
        if col_name == 'Alıcı VKN/TCKN' and value.endswith('.0'):
            value = value[:-2]
        elif col_name == 'Ödenecek Tutar':
            try:
                value = f"{float(value):,.0f} TL".replace(",", "X").replace(".", ",").replace("X", ".")
            except:
                pass
        elif col_name == 'Fatura Düzenlenme Tarihi':
            try:
                if pd.notna(row_data[col_name]):
                    value = pd.to_datetime(row_data[col_name]).strftime('%d.%m.%Y')
            except:
                pass

        # NaN değerlerini boş string yap
        if pd.isna(row_data[col_name]) or str(value).lower() == 'nan':
            display_value = ""
        else:
            display_value = value

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

    # ================== ROW OPERATIONS ==================
    def on_row_double_click(self, row, column):
        """Satır çift tıklandığında onay dialog'unu göster"""
        self.show_confirmation_dialog(row)

    def show_confirmation_dialog(self, row_idx):
        """Seçilen satır için onay penceresi açar"""
        try:
            # Filtrelenmiş tablodan seçilen satır
            selected_row = self.veri_cercevesi.iloc[row_idx]

            # Excel'deki tam karşılığını bulmak için eşsiz kombinasyon
            mask = (
                (self.full_df['Alıcı VKN/TCKN'] == selected_row['Alıcı VKN/TCKN']) &
                (self.full_df['Fatura Numarası'] == selected_row['Fatura Numarası'])
            )

            matching_rows = self.full_df[mask]

            if len(matching_rows) != 1:
                # Eğer sütun isimleri farklıysa alternatif kontrol
                if 'Fatura No' in selected_row and 'Alıcı VKN/TCKN' in selected_row:
                    mask = (
                        (self.full_df['Alıcı VKN/TCKN'] == selected_row['Alıcı VKN/TCKN']) &
                        (self.full_df['Fatura No'] == selected_row['Fatura No'])
                    )
                    matching_rows = self.full_df[mask]

                if len(matching_rows) != 1:
                    raise ValueError("Eşleşen fatura bulunamadı veya çoklu eşleşme var")

            original_index = matching_rows.index[0]
            data = self.full_df.loc[original_index]

            # Dialog penceresi oluştur
            dialog = QDialog(self)
            dialog.setWindowTitle("Fiş Onayı")
            dialog.setFixedSize(900, 700)

            # Pencereyi sol tarafa yakın konumlandır
            parent_rect = self.geometry()
            dialog.move(parent_rect.left() + 100, parent_rect.top() + 100)

            layout = QVBoxLayout()
            layout.setContentsMargins(20, 20, 20, 20)
            layout.setSpacing(30)

            # Dialog'un arka plan rengini ayarla
            dialog.setStyleSheet(DIALOG_STYLE)

            # Büyük fontla veri gösterme fonksiyonu
            def add_large_value(value):
                widget = QWidget()
                hbox = QHBoxLayout(widget)
                hbox.setAlignment(Qt.AlignCenter)

                val = QLabel(str(value))
                val.setStyleSheet(DIALOG_VALUE_STYLE)

                hbox.addWidget(val)
                return widget

            # VKN/TCKN
            vkn_key = 'Alıcı VKN/TCKN' if 'Alıcı VKN/TCKN' in data else 'VKN/TCKN'
            vkn = str(data[vkn_key]).replace('.0', '')
            if len(vkn) >= 10:
                formatted_vkn = ' '.join([vkn[:3], vkn[3:6], vkn[6:9], vkn[9:]])
            else:
                formatted_vkn = vkn
            layout.addWidget(add_large_value(formatted_vkn))

            # Fatura No
            fno_key = 'Fatura Numarası' if 'Fatura Numarası' in data else 'Fatura No'
            fno = str(data[fno_key])
            if len(fno) >= 16:
                fno_parts = [fno[:3], fno[3:7], fno[7:10], fno[10:13], fno[13:]]
                layout.addWidget(add_large_value(' '.join(fno_parts)))
            else:
                layout.addWidget(add_large_value(fno))

            # Uyarı
            warning = QLabel("Fişi kestikten sonra onaylayın!")
            warning.setStyleSheet(DIALOG_WARNING_STYLE)
            layout.addWidget(warning, alignment=Qt.AlignCenter)

            # Butonlar
            btn_box = QWidget()
            btn_layout = QHBoxLayout(btn_box)
            btn_layout.setContentsMargins(0, 30, 0, 0)
            btn_layout.setSpacing(30)

            cancel_btn = QPushButton("İptal")
            cancel_btn.setStyleSheet(DIALOG_CANCEL_BTN_STYLE)
            cancel_btn.clicked.connect(dialog.close)
            btn_layout.addWidget(cancel_btn)

            confirm_btn = QPushButton("Onayla")
            confirm_btn.setStyleSheet(DIALOG_CONFIRM_BTN_STYLE)
            confirm_btn.clicked.connect(lambda: self.mark_as_processed(dialog, original_index))
            btn_layout.addWidget(confirm_btn)

            layout.addWidget(btn_box)
            dialog.setLayout(layout)

            # Klavye kısayolları
            confirm_btn.setDefault(True)
            cancel_btn.setShortcut(Qt.Key_Escape)

            dialog.exec_()

        except Exception as e:
            logger.exception("Onay penceresi açma hatası")
            QMessageBox.critical(self, "Hata", f"Onay penceresi açılamadı: {str(e)}")

    def mark_as_processed(self, dialog, original_index):
        """Google Sheets'te ilgili satırı 'OK' olarak işaretler"""
        try:
            # Progress bar göster
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 0)  # Indeterminate
            self.status_label.setText("📝 Google Sheets güncelleniyor...")
            QApplication.processEvents()

            # Güncellenecek satırın fatura numarasını al
            selected_row = self.full_df.iloc[original_index]
            fatura_no_key = 'Fatura Numarası' if 'Fatura Numarası' in selected_row else 'Fatura No'
            fatura_no = str(selected_row[fatura_no_key])

            # Google Sheets API ile güncelleme yap
            success = self._update_google_sheets(fatura_no)

            if success:
                # Lokal verileri güncelle
                self.full_df.loc[original_index, 'YazarKasa'] = 'OK'

                # Filtreleme yaparken YazarKasa != 'OK' olanları göster
                if 'YazarKasa' in self.full_df.columns:
                    self.original_df = self.full_df[self.full_df['YazarKasa'] != 'OK'].copy()
                else:
                    self.original_df = self.full_df.copy()

                # Sıralama yap
                if 'Fatura Düzenlenme Tarihi' in self.original_df.columns and 'Ödenecek Tutar' in self.original_df.columns:
                    self.original_df = self.original_df.sort_values(
                        ['Fatura Düzenlenme Tarihi', 'Ödenecek Tutar'],
                        ascending=[False, True]
                    )

                self.veri_cercevesi = self.original_df.copy()
                self.populate_table()

                # Başarı mesajı
                QMessageBox.information(
                    self,
                    "Başarılı",
                    f"Fiş onaylandı!\n\nFiş numarası: {fatura_no}\n\nKayıt tablodan kaldırıldı."
                )

                self.status_label.setText("✅ Fiş başarıyla onaylandı ve güncellendi")
                dialog.close()
            else:
                raise Exception("Google Sheets güncelleme işlemi başarısız")

        except Exception as e:
            logger.exception("Kayıt güncelleme hatası")
            QMessageBox.critical(
                self,
                "Güncelleme Hatası",
                f"Kayıt güncellenirken hata oluştu:\n{str(e)}"
            )
        finally:
            self.progress_bar.setVisible(False)

    def _update_google_sheets(self, fatura_no):
        """GoogleSheetsManager yapısını kullanarak OKC sayfasında YazarKasa güncelleme"""
        try:
            # GoogleSheetsManager oluştur
            sheets_manager = self._create_google_sheets_manager()

            if not sheets_manager:
                return False

            # PRGsheet dosyasını aç
            spreadsheet = sheets_manager.gc.open("PRGsheet")
            okc_worksheet = spreadsheet.worksheet(SHEET_NAME_OKC)

            # Tüm veriyi al
            all_values = okc_worksheet.get_all_values()

            if not all_values:
                raise Exception("OKC sayfasında veri bulunamadı")

            # Header satırını al
            headers = all_values[0]

            # Fatura Numarası ve YazarKasa sütunlarının indekslerini bul
            fatura_col_idx = None
            yazarkasa_col_idx = None

            for i, header in enumerate(headers):
                if header in ['Fatura Numarası', 'Fatura No']:
                    fatura_col_idx = i
                elif header == 'YazarKasa':
                    yazarkasa_col_idx = i

            if fatura_col_idx is None:
                raise Exception("Fatura Numarası sütunu bulunamadı")
            if yazarkasa_col_idx is None:
                raise Exception("YazarKasa sütunu bulunamadı")

            # Eşleşen fatura numarasını bul
            target_row = None
            for row_idx, row in enumerate(all_values[1:], start=2):  # 2'den başla (1 header, gspread 1-based)
                if len(row) > fatura_col_idx and str(row[fatura_col_idx]) == fatura_no:
                    target_row = row_idx
                    break

            if target_row is None:
                raise Exception(f"Fatura numarası '{fatura_no}' bulunamadı")

            # YazarKasa hücresini güncelle
            okc_worksheet.update_cell(target_row, yazarkasa_col_idx + 1, "OK")  # gspread 1-based indexing

            return True

        except Exception as e:
            logger.exception("Google Sheets güncelleme hatası")
            QMessageBox.critical(
                self,
                "Google Sheets Güncelleme Hatası",
                f"Google Sheets güncellenirken hata:\n{str(e)}"
            )
            return False

    def _create_google_sheets_manager(self):
        """Service Account kullanan CentralConfigManager ile Google Sheets bağlantısı oluştur"""
        try:
            # Service Account kullanan merkezi config manager'ı kullan
            config_manager = CentralConfigManager()

            # Basit wrapper class oluştur
            class GoogleSheetsManager:
                def __init__(self, config_manager):
                    self.gc = config_manager.gc  # Service Account ile yetkilendirilmiş client

            return GoogleSheetsManager(config_manager)

        except ImportError as e:
            logger.exception("Merkezi config import hatası")
            QMessageBox.critical(
                self,
                "Kütüphane Hatası",
                f"Merkezi config için gerekli modül yüklenemedi:\n{str(e)}\n\n"
                "central_config.py dosyasının mevcut olduğundan emin olun."
            )
            return None
        except Exception as e:
            logger.exception("Google Sheets Manager oluşturma hatası")
            QMessageBox.critical(
                self,
                "Google Sheets Manager Hatası",
                f"Service Account ile Google Sheets bağlantısı kurulamadı:\n{str(e)}\n\n"
                "service_account.json dosyasının geçerli olduğundan emin olun."
            )
            return None

    # ================== FILTER OPERATIONS ==================
    def filter_data(self, text):
        """Fatura tutarına göre filtreleme"""
        if text:
            try:
                filter_value = int(text) * FILTER_MULTIPLIER
                if 'Ödenecek Tutar' in self.original_df.columns:
                    filtered_df = self.original_df[self.original_df['Ödenecek Tutar'] >= filter_value].copy()
                    # Filtrelenmiş veriyi de aynı şekilde sırala
                    if 'Fatura Düzenlenme Tarihi' in filtered_df.columns and 'Ödenecek Tutar' in filtered_df.columns:
                        filtered_df = filtered_df.sort_values(
                            ['Fatura Düzenlenme Tarihi', 'Ödenecek Tutar'],
                            ascending=[False, True]
                        )
                    self.veri_cercevesi = filtered_df
                    self.populate_table()
            except ValueError:
                pass
        else:
            self.veri_cercevesi = self.original_df.copy()
            self.populate_table()

    def clear_search(self):
        """Arama kutusunu temizle"""
        self.search_input.clear()
        self.veri_cercevesi = self.original_df.copy()
        self.populate_table()

    # ================== E-ARSIV EXECUTION ==================
    def run_e_arsiv(self):
        """e-Arşiv programını çalıştır"""
        try:
            program_path = OKC_EXE_PATH

            # Path kontrolü
            if not program_path.exists():
                self.status_label.setText(f"❌ OKC.exe bulunamadı")
                QMessageBox.critical(self, "Hata", f"OKC.exe bulunamadı:\n{program_path}\n\nLütfen dosyanın var olduğundan emin olun.")
                return

            if not program_path.is_file():
                self.status_label.setText(f"❌ OKC.exe bir dosya değil")
                QMessageBox.critical(self, "Hata", f"OKC.exe bir dosya değil: {program_path}")
                return

            # Dosya boyutu kontrolü
            file_size = program_path.stat().st_size
            if file_size == 0:
                self.status_label.setText("❌ OKC.exe dosyası bozuk")
                QMessageBox.critical(self, "Hata", f"OKC.exe dosyası bozuk (0 byte):\n{program_path}")
                return

            self.status_label.setText("🔄 OKC.exe çalıştırılıyor...")
            self.e_arsiv_btn.setEnabled(False)
            self.clear_btn.setEnabled(False)

            QApplication.processEvents()

            # Programı başlat
            os.startfile(str(program_path))

            # OKC.exe'nin çalışması için bekleme
            QTimer.singleShot(OKC_EXECUTION_TIMEOUT_MS, self.on_e_arsiv_finished)

        except PermissionError as e:
            logger.exception("İzin hatası")
            self.status_label.setText("❌ İzin hatası")
            QMessageBox.critical(self, "İzin Hatası", f"OKC.exe çalıştırma izni yok:\n{str(e)}\n\nDosyayı yönetici olarak çalıştırmayı deneyin.")
            self.e_arsiv_btn.setEnabled(True)
            self.clear_btn.setEnabled(True)
        except OSError as e:
            logger.exception("Sistem hatası")
            self.status_label.setText("❌ Dosya çalıştırma hatası")
            QMessageBox.critical(self, "Sistem Hatası", f"OKC.exe çalıştırılamadı:\n{str(e)}\n\nDosya bozuk veya uyumlu değil olabilir.")
            self.e_arsiv_btn.setEnabled(True)
            self.clear_btn.setEnabled(True)
        except Exception as e:
            logger.exception("Program çalıştırma hatası")
            self.status_label.setText(f"❌ Program çalıştırma hatası: {str(e)}")
            QMessageBox.critical(self, "Hata", f"Beklenmeyen hata:\n{str(e)}\n\nDetay: {type(e).__name__}")
            self.e_arsiv_btn.setEnabled(True)
            self.clear_btn.setEnabled(True)

    def on_e_arsiv_finished(self):
        """e-Arşiv program bittikten sonra"""
        self.e_arsiv_btn.setEnabled(True)
        self.clear_btn.setEnabled(True)
        self.status_label.setText("✅ OKC.exe tamamlandı, Google Sheets güncelleme bekleniyor...")

        # Google Sheets'e kaydedilmesi için ek bekleme
        QTimer.singleShot(SHEETS_UPDATE_DELAY_MS, self.delayed_data_refresh)

    def delayed_data_refresh(self):
        """Gecikmeli veri yenileme"""
        self.status_label.setText("🔄 Google Sheets'ten güncel veriler alınıyor...")
        QApplication.processEvents()
        self.load_data()

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
            self.status_label.setText("✅ Kopyalandı")
        else:
            self.status_label.setText("⚠️ Boş hücre")

    # ================== UTILITY ==================
    def set_buttons_enabled(self, enabled: bool):
        """Butonları aktif/pasif yap"""
        self.clear_btn.setEnabled(enabled)
        self.e_arsiv_btn.setEnabled(enabled)
