"""
İrsaliye Modülü
"""

import os
import sys
import pandas as pd
import requests
from io import BytesIO
from datetime import datetime
from pathlib import Path

# Üst dizini Python path'e ekle (central_config için)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Central config import
from central_config import CentralConfigManager

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
                             QMenu, QProgressBar, QLabel, QTabWidget, QFileDialog, QMessageBox, QApplication)
from PyQt5.QtGui import QFont, QColor, QKeySequence


# ================== CONFIG CONSTANTS ==================
# UI
MIN_COLUMN_WIDTH = 150
ROW_HEIGHT = 35
FONT_FAMILY = "Segoe UI"
FONT_SIZE = 12


# ================== STYLESHEET CONSTANTS ==================
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


class CopyableTableWidget(QTableWidget):
    """Ctrl+C ile kopyalama destekli QTableWidget"""
    def keyPressEvent(self, event):
        if event.matches(QKeySequence.Copy):
            self.copy_selection()
        else:
            super().keyPressEvent(event)

    def copy_selection(self):
        """Seçili satırları tab-separated format ile clipboard'a kopyala"""
        selection = self.selectedIndexes()
        if not selection:
            return

        # Satır ve sütunları grupla
        rows = sorted(set(index.row() for index in selection))
        cols = sorted(set(index.column() for index in selection))

        # Veriyi topla
        clipboard_data = []
        for row in rows:
            row_data = []
            for col in range(self.columnCount()):
                item = self.item(row, col)
                row_data.append(item.text() if item else "")
            clipboard_data.append("\t".join(row_data))

        # Clipboard'a kopyala
        QApplication.clipboard().setText("\n".join(clipboard_data))


class IrsaliyeWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("İrsaliye Modülü - PRG v2.0")
        self.setMinimumSize(1200, 800)
        self.mikro_calisiyor = False
        self.bagkodu_calisiyor = False
        self.veri_cercevesi_fatura = pd.DataFrame()
        self.veri_cercevesi_irsaliye = pd.DataFrame()
        self.gsheets_url = self._load_gsheets_url()
        
        # Apply main window styling - Light theme
        self.setStyleSheet("""
            QMainWindow {
                background-color: #ffffff;
                color: #000000;
            }
        """)

        self.setup_ui()
        self.setup_irsaliye_connections()

        # Lazy loading için flag
        self._data_loaded = False

    def showEvent(self, event):
        """Widget ilk gösterildiğinde veri yükle (lazy loading)"""
        super().showEvent(event)
        if not self._data_loaded:
            self._data_loaded = True
            # UI render olduktan sonra veri yükle
            QTimer.singleShot(100, self.load_data)
    
    def _load_gsheets_url(self):
        """Google Sheets SPREADSHEET_ID'sini yükle - Service Account"""
        try:
            config_manager = CentralConfigManager()
            spreadsheet_id = config_manager.MASTER_SPREADSHEET_ID
            return f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"
        except Exception as e:
            if hasattr(self, 'status_label'):
                self.status_label.setText(f"❌ PRGsheet yüklenirken hata: {str(e)}")
            return None
    
    
    def setup_ui(self):
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        # Light theme - Force white background
        self.central_widget.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                color: #000000;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
        """)
        self.central_widget.setAutoFillBackground(True)
        palette = self.central_widget.palette()
        palette.setColor(self.central_widget.backgroundRole(), QColor("#ffffff"))
        self.central_widget.setPalette(palette)
        
        layout = QVBoxLayout(self.central_widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # Header Layout
        header_layout = QHBoxLayout()
        
        # Buttons - Light theme
        self.mikro_button = QPushButton("Mikro")
        self.mikro_button.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 16px;
                border-radius: 5px;
                font-weight: bold;
                font-size: 14px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #888888;
            }
        """)
        
        self.refresh_button = QPushButton("Yenile")
        self.refresh_button.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 16px;
                border-radius: 5px;
                font-weight: bold;
                font-size: 14px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
        """)
        
        
        
        self.stok_aktar_button = QPushButton("Stok Aktar")
        self.stok_aktar_button.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 16px;
                border-radius: 5px;
                font-weight: bold;
                font-size: 14px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
        """)
        
        self.irsaliye_kaydet_button = QPushButton("İrsaliye Kaydet")
        self.irsaliye_kaydet_button.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 16px;
                border-radius: 5px;
                font-weight: bold;
                font-size: 14px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
        """)
        
        header_layout.addWidget(self.mikro_button)
        header_layout.addWidget(self.refresh_button)
        header_layout.addStretch()
        header_layout.addWidget(self.stok_aktar_button)
        header_layout.addWidget(self.irsaliye_kaydet_button)
        
        # Header layout'u widget olarak sar - beyaz arka plan için
        header_widget = QWidget()
        header_widget.setLayout(header_layout)
        header_widget.setStyleSheet("""
            background-color: #ffffff;
            margin-bottom: 0px;
        """)
        header_layout.setContentsMargins(10, 10, 10, 10)
        
        # Progress Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)  # Yuzde metnini goster
        self.progress_bar.setAlignment(Qt.AlignCenter)  # Metni ortala
        self.progress_bar.setFormat("%p%")  # Yuzde formati
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
        
        # Tab Widget - Light theme
        self.tab_widget = QTabWidget()
        self.tab_widget.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #d0d0d0;
                background-color: #ffffff;
                border-radius: 0px;
                margin-top: -1px;
            }
            QTabWidget::tab-bar {
                alignment: left;
            }
            QTabBar::tab {
                background-color: #f0f0f0;
                color: #666666;
                border: 1px solid #d0d0d0;
                border-bottom: none;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                padding: 10px 28px;
                margin-right: 4px;
                font-size: 14px;
                font-weight: 600;
                font-family: 'Segoe UI', Arial, sans-serif;
                min-width: 150px;
            }
            QTabBar::tab:selected {
                background-color: #ffffff;
                border-color: #4CAF50;
                border-bottom: 2px solid #ffffff;
                color: #000000;
            }
            QTabBar::tab:hover:!selected {
                background-color: #e0e0e0;
                color: #000000;
            }
        """)
        
        # Status Layout (Label + Progress Bar)
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
        
        # Status layout'a ekle (kasa_module.py gibi 3:1 oranında)
        status_layout.addWidget(self.status_label, 3)
        status_layout.addWidget(self.progress_bar, 1)
        status_layout.setContentsMargins(0, 0, 0, 0)
        
        # Status layout'u widget olarak sar
        status_widget = QWidget()
        status_widget.setLayout(status_layout)
        status_widget.setStyleSheet("background-color: #f5f5f5; border-top: 1px solid #d0d0d0;")
        
        layout.addWidget(header_widget)
        layout.addWidget(self.tab_widget, 1)
        layout.addWidget(status_widget)
    
    def setup_irsaliye_connections(self):
        self.mikro_button.clicked.connect(self.run_mikro)
        self.refresh_button.clicked.connect(self.load_data)
        self.stok_aktar_button.clicked.connect(self.run_bagkodu)
        self.irsaliye_kaydet_button.clicked.connect(self.irsaliye_kaydet)
    
    def load_data(self):
        """Google Sheets'ten Fatura ve Irsaliye sayfalarından verileri yükle"""
        try:
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 100)  # Yuzde bazli
            self.progress_bar.setValue(0)  # 0%
            self.status_label.setText("📊 Veriler yükleniyor...")
            self.set_buttons_enabled(False)

            from PyQt5.QtWidgets import QApplication
            QApplication.processEvents()

            # Progress: 10%
            self.progress_bar.setValue(10)
            QApplication.processEvents()

            if not self.gsheets_url:
                self.status_label.setText("❌ PRGsheet/Ayar sayfasında SPREADSHEET_ID bulunamadı")
                return

            # Progress: 20% - Google Sheets'e baglaniliyor
            self.progress_bar.setValue(20)
            QApplication.processEvents()

            # URL'den Excel dosyasını oku
            response = requests.get(self.gsheets_url, timeout=30)

            if response.status_code == 401:
                self.status_label.setText("❌ Google Sheets erişim hatası: Dosya özel veya izin gerekli.")
                return
            elif response.status_code != 200:
                self.status_label.setText(f"❌ HTTP Hatası: {response.status_code} - {response.reason}")
                return

            response.raise_for_status()

            from io import BytesIO

            # Geçici değişkenler
            temp_fatura = None
            temp_irsaliye = None
            fatura_error = None
            irsaliye_error = None

            # Progress: 40% - Fatura sayfasi okunuyor
            self.progress_bar.setValue(40)
            QApplication.processEvents()

            # Fatura sayfasını oku
            try:
                temp_fatura = pd.read_excel(BytesIO(response.content), sheet_name="Fatura")
            except Exception as e:
                fatura_error = f"Sayfa bulunamadı veya okunamadı ({str(e)})"
                temp_fatura = pd.DataFrame()

            # Progress: 60% - Irsaliye sayfasi okunuyor
            self.progress_bar.setValue(60)
            QApplication.processEvents()

            # İrsaliye sayfasını oku
            try:
                temp_irsaliye = pd.read_excel(BytesIO(response.content), sheet_name="Irsaliye")
            except Exception as e:
                irsaliye_error = f"Sayfa bulunamadı veya okunamadı ({str(e)})"
                temp_irsaliye = pd.DataFrame()

            # Başarılı okumaları instance değişkenlerine ata
            self.veri_cercevesi_fatura = temp_fatura
            self.veri_cercevesi_irsaliye = temp_irsaliye

            # Progress: 80% - Tablolar dolduruluyor
            self.progress_bar.setValue(80)
            QApplication.processEvents()

            self.populate_tables()

            # Progress: 90%
            self.progress_bar.setValue(90)
            QApplication.processEvents()

            # Sonuç mesajı oluştur
            messages = []
            if not temp_fatura.empty:
                messages.append(f"Fatura: {len(temp_fatura)} kayıt")
            elif fatura_error:
                messages.append(f"Fatura: ❌ {fatura_error}")

            if not temp_irsaliye.empty:
                messages.append(f"İrsaliye: {len(temp_irsaliye)} kayıt")
            elif irsaliye_error:
                messages.append(f"İrsaliye: ❌ {irsaliye_error}")

            # Progress: 100% - Tamamlandi
            self.progress_bar.setValue(100)
            QApplication.processEvents()

            if messages:
                self.status_label.setText(f"✅ Veriler yüklendi - {', '.join(messages)}")
            else:
                self.status_label.setText("❌ Hiçbir sayfa yüklenemedi - Google Sheets'te 'Fatura' ve 'Irsaliye' sayfalarını kontrol edin")

            # Progress bar'i 1 saniye sonra gizle
            QTimer.singleShot(1000, lambda: self.progress_bar.setVisible(False))

        except requests.exceptions.Timeout:
            self.status_label.setText("❌ Bağlantı zaman aşımı - Google Sheets'e erişilemiyor")
            self.progress_bar.setVisible(False)  # Hata durumunda hemen gizle
        except requests.exceptions.RequestException as e:
            self.status_label.setText(f"❌ Bağlantı hatası: {str(e)}")
            self.progress_bar.setVisible(False)  # Hata durumunda hemen gizle
        except Exception as e:
            self.status_label.setText(f"❌ Veri yükleme hatası: {str(e)}")
            self.progress_bar.setVisible(False)  # Hata durumunda hemen gizle
        finally:
            self.set_buttons_enabled(True)
    
    def populate_tables(self):
        """Tabloları verilerle doldur"""
        self.tab_widget.clear()

        # İrsaliye tablosunu ÖNCE oluştur (ilk açılacak sekme)
        if not self.veri_cercevesi_irsaliye.empty:
            self.create_table(self.veri_cercevesi_irsaliye, "İrsaliye Kayıt")

        # Fatura tablosunu SONRA oluştur
        if not self.veri_cercevesi_fatura.empty:
            self.create_table(self.veri_cercevesi_fatura, "Fatura Hata")

        # Eğer hiç veri yoksa boş tab ekle
        if self.veri_cercevesi_fatura.empty and self.veri_cercevesi_irsaliye.empty:
            empty_widget = QWidget()
            empty_layout = QVBoxLayout(empty_widget)
            empty_label = QLabel("Veri bulunamadı")
            empty_label.setAlignment(Qt.AlignCenter)
            empty_label.setStyleSheet("font-size: 24px; color: #666666; margin: 50px;")
            empty_layout.addWidget(empty_label)
            self.tab_widget.addTab(empty_widget, "Veri Yok")
    
    def create_table(self, dataframe, title):
        """Tablo oluştur"""
        table = CopyableTableWidget()  # Kopyalama destekli tablo
        table.setRowCount(dataframe.shape[0])
        
        table.setColumnCount(dataframe.shape[1])
        table.setHorizontalHeaderLabels(dataframe.columns)

        # Tablo stilini uygula - Light theme (risk_module.py gibi)
        table.setStyleSheet(TABLE_STYLE)
        
        # Tablo özelliklerini ayarla
        table.setAlternatingRowColors(True)
        table.setShowGrid(True)
        table.setSortingEnabled(False)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)  # Satır seçimi
        table.setSelectionMode(QAbstractItemView.ExtendedSelection)  # Çoklu seçim (Ctrl/Shift ile)
        table.setFocusPolicy(Qt.NoFocus)  # Focus border'ı kaldır (risk_module.py gibi)

        # Context menu (sağ tık) için custom policy - risk_module.py gibi
        table.setContextMenuPolicy(Qt.CustomContextMenu)
        table.customContextMenuRequested.connect(lambda pos, t=table: self.show_context_menu(pos, t))
        
        # Tabloyu verilerle doldur
        for i, row in dataframe.iterrows():
            for j, value in enumerate(row):
                # NaN değerlerini boş string yap
                if pd.isna(value) or str(value).lower() == 'nan':
                    display_value = ""
                else:
                    display_value = str(value)
                
                item = QTableWidgetItem(display_value)
                item.setFlags(item.flags() ^ Qt.ItemIsEditable)  # Make non-editable

                # Set font properties - bold (risk_module.py gibi)
                font = QFont(FONT_FAMILY, FONT_SIZE)
                font.setBold(True)
                item.setFont(font)
                item.setForeground(QColor("#000000"))

                table.setItem(i, j, item)
        
        # Header stillerini uygula
        header = table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setStretchLastSection(False)
        
        # Minimum sütun genişliklerini ayarla
        for i in range(table.columnCount()):
            table.setColumnWidth(i, max(MIN_COLUMN_WIDTH, table.columnWidth(i)))

        # Sütunları içeriğe göre boyutlandır
        table.resizeColumnsToContents()

        # Satır yüksekliğini ayarla - daha kompakt görünüm
        for i in range(table.rowCount()):
            table.setRowHeight(i, ROW_HEIGHT)
        
        
        # Tablonu tab widget'a ekle
        self.tab_widget.addTab(table, title)
    
    
    def run_mikro(self):
        """Irsaliye.exe dosyasını çalıştır"""
        try:
            exe_path = r"D:/GoogleDrive/PRG/EXE/Irsaliye.exe"
            if not os.path.exists(exe_path):
                self.status_label.setText(f"❌ Irsaliye.exe bulunamadı: {exe_path}")
                return
            
            self.status_label.setText("🔄 Irsaliye.exe çalıştırılıyor...")
            self.mikro_button.setEnabled(False)
            self.mikro_calisiyor = True
            
            os.startfile(exe_path)
            
            # 7 saniye sonra program bitmiş sayıp kontrol et
            QTimer.singleShot(7000, self.on_mikro_finished)
            
        except Exception as e:
            self.status_label.setText(f"❌ Program çalıştırma hatası: {str(e)}")
            self.mikro_button.setEnabled(True)
            self.mikro_calisiyor = False
    
    def on_mikro_finished(self):
        """Mikro program bittikten sonra"""
        self.mikro_button.setEnabled(True)
        self.mikro_calisiyor = False
        self.status_label.setText("✅ Irsaliye.exe tamamlandı, Google Sheets güncelleme bekleniyor...")
        
        # Google Sheets'e kaydedilmesi için ek bekleme (5 saniye)
        # Sonra verileri yenile
        QTimer.singleShot(5000, self.delayed_data_refresh)
    
    def delayed_data_refresh(self):
        """Gecikmeli veri yenileme"""
        self.status_label.setText("🔄 Google Sheets'ten güncel veriler alınıyor...")
        from PyQt5.QtWidgets import QApplication
        QApplication.processEvents()
        self.load_data()
    
    def run_bagkodu(self):
        """BagKodu.exe dosyasını çalıştır"""
        try:
            exe_path = r"D:/GoogleDrive/PRG/EXE/BagKodu.exe"
            if not os.path.exists(exe_path):
                self.status_label.setText(f"❌ BagKodu.exe bulunamadı: {exe_path}")
                return
            
            self.status_label.setText("🔄 BagKodu.exe çalıştırılıyor...")
            self.stok_aktar_button.setEnabled(False)
            self.bagkodu_calisiyor = True
            
            os.startfile(exe_path)
            
            # 7 saniye sonra program bitmiş sayıp kontrol et
            QTimer.singleShot(7000, self.on_bagkodu_finished)
            
        except Exception as e:
            self.status_label.setText(f"❌ BagKodu çalıştırma hatası: {str(e)}")
            self.stok_aktar_button.setEnabled(True)
            self.bagkodu_calisiyor = False
    
    def on_bagkodu_finished(self):
        """BagKodu program bittikten sonra"""
        self.stok_aktar_button.setEnabled(True)
        self.bagkodu_calisiyor = False
        self.status_label.setText("✅ BagKodu.exe tamamlandı, Google Sheets güncelleme bekleniyor...")
        
        # BagKodu işlemi tamamlandı
        self.status_label.setText("✅ BagKodu verileri Google Sheets'e kaydedildi")
    
    
    
    def convert_float_to_clean_string(self, value):
        """Float değerleri temiz string formatına dönüştür (.0 sorunu için)"""
        try:
            if pd.notna(value) and str(value).replace('.', '').replace('-', '').isdigit():
                # Eğer sayısal bir değerse ve ondalık kısmı varsa
                if '.' in str(value):
                    return str(int(float(value)))
                else:
                    return str(value)
            else:
                return str(value) if pd.notna(value) else ""
        except (ValueError, TypeError):
            return str(value) if pd.notna(value) else ""

    def find_column_name(self, headers, target_column):
        """Sütun adını esnek şekilde bul - case insensitive ve alternatif isimlerle"""
        # Olası sütun adı alternatifleri
        alternatives = {
            "Prosap Sas Kalem no": [
                "Prosap Sas Kalem no", "Prosap SAS Kalem No", "Prosap SAS Kalem No.", 
                "SAS Kalem No", "Kalem No", "prosap sas kalem no", "PROSAP SAS KALEM NO"
            ],
            "Fatura No": [
                "Fatura No", "Fatura Numarası", "FaturaNo", "FATURA NO", "fatura no"
            ],
            "Fatura Tarihi": [
                "Fatura Tarihi", "Fatura Tarih", "FaturaTarihi", "FATURA TARIHI", "fatura tarihi"
            ]
        }
        
        # Hedef sütun için alternatifleri al
        possible_names = alternatives.get(target_column, [target_column])
        
        # Önce tam eşleşme ara
        for alt_name in possible_names:
            if alt_name in headers:
                return alt_name
        
        # Sonra case-insensitive eşleşme ara
        for alt_name in possible_names:
            for header in headers:
                if alt_name.lower() == header.lower():
                    return header
        
        # Son olarak kısmi eşleşme ara
        for alt_name in possible_names:
            for header in headers:
                if alt_name.lower() in header.lower() or header.lower() in alt_name.lower():
                    return header
        
        return None

    def show_context_menu(self, position, table):
        """Sağ tık menüsü - Sadece hücre kopyalama (risk_module.py gibi)"""
        item = table.itemAt(position)
        if not item:
            return

        menu = QMenu(self)
        menu.setStyleSheet(CONTEXT_MENU_STYLE)

        copy_action = menu.addAction("Kopyala")

        action = menu.exec_(table.viewport().mapToGlobal(position))

        if action == copy_action:
            self.copy_cell(item)

    def copy_cell(self, item: QTableWidgetItem):
        """Tıklanan hücreyi kopyala (risk_module.py gibi)"""
        if item and item.text():
            QApplication.clipboard().setText(item.text())
            self.status_label.setText("✅ Kopyalandı")
        else:
            self.status_label.setText("⚠️ Boş hücre")

    def irsaliye_kaydet(self):
        """İrsaliye Kaydet işlevi - PRG.py'daki export_to_excel fonksiyonunu uygular"""
        try:
            # 1. Kullanıcıdan Excel dosyası seçmesini iste
            file_path, _ = QFileDialog.getOpenFileName(
                self, 
                "Excel Dosyası Seç", 
                "D:/GoogleDrive/", 
                "Excel Files (*.xlsx *.xls)"
            )
            
            if not file_path:
                return  # Kullanıcı iptal etti
                
            # 2. Excel dosyasını oku
            excel_df = pd.read_excel(file_path)
            
            # 3. Fatura sütununu bul
            fatura_col = None
            for col in ['Faturalama belgesi', 'SAP Fatura No']:
                if col in excel_df.columns:
                    fatura_col = col
                    break
                    
            if fatura_col is None:
                QMessageBox.warning(self, "Uyarı", "Excel dosyasında fatura bilgisi içeren sütun bulunamadı!")
                return
                
            # 4. 900 ile başlayan fatura numaralarını bul
            # Float değerleri temiz string formatına dönüştür (.0 sorunu için)
            excel_df[fatura_col] = excel_df[fatura_col].apply(self.convert_float_to_clean_string)
            fatura_nolari = excel_df[
                excel_df[fatura_col].astype(str).str.startswith('900', na=False)
            ][fatura_col].unique()

            if len(fatura_nolari) == 0:
                QMessageBox.warning(self, "Uyarı", "900 ile başlayan fatura numarası bulunamadı!")
                return
                
            # 5. Bulunan faturaları kullanıcıya göster
            # Fatura numaralarını düzgün formatta göster (.0 olmadan)
            fatura_listesi = "\n".join([self.convert_float_to_clean_string(num) for num in fatura_nolari])
            msg_box = QMessageBox()
            msg_box.setWindowTitle("Bulunan Faturalar")
            msg_box.setIcon(QMessageBox.Information)
            msg_box.setText(
                f"Excel'de bulunan 900 ile başlayan faturalar:\n\n{fatura_listesi}\n\n"
                f"Toplam {len(fatura_nolari)} adet fatura bulundu.\n"
                "Bu faturaları işlemek için Devam Et butonuna basın."
            )
            msg_box.addButton("Devam Et", QMessageBox.AcceptRole)
            msg_box.addButton("İptal", QMessageBox.RejectRole)
            
            if msg_box.exec_() == QMessageBox.RejectRole:
                return

            # 6. Mevcut tablodan verileri al
            current_tab_index = self.tab_widget.currentIndex()
            current_table = self.tab_widget.widget(current_tab_index)
            current_tab_name = self.tab_widget.tabText(current_tab_index) if current_tab_index >= 0 else "Bilinmeyen"
            
            if not isinstance(current_table, QTableWidget):
                QMessageBox.warning(self, "Uyarı", f"Aktif tab bir tablo değil!\n\nSeçili tab: '{current_tab_name}'\n\nLütfen 'İrsaliye Kayıt' tabını seçin.")
                return
            
            # Tab kontrolü - doğru tab seçili mi?
            if current_tab_name not in ["İrsaliye Kayıt"]:
                response = QMessageBox.question(
                    self,
                    "Tab Kontrolü",
                    f"Seçili tab: '{current_tab_name}'\n\nİrsaliye Kaydet işlemi için 'İrsaliye Kayıt' tabını seçmeniz önerilir.\n\nDevam etmek istiyor musunuz?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No
                )
                if response == QMessageBox.No:
                    return
            
            # Sütun başlıklarını ve indekslerini bul
            headers = []
            for j in range(current_table.columnCount()):
                header = current_table.horizontalHeaderItem(j).text()
                headers.append(header)
            
            # Debug: Mevcut sütunları göster
            self.status_label.setText(f"📋 Mevcut sütunlar: {', '.join(headers[:5])}..." if len(headers) > 5 else f"📋 Mevcut sütunlar: {', '.join(headers)}")
            from PyQt5.QtWidgets import QApplication
            QApplication.processEvents()
            
            # Esnek sütun eşleştirme
            fatura_no_column = self.find_column_name(headers, "Fatura No")
            if not fatura_no_column:
                QMessageBox.warning(self, "Uyarı", f"Tabloda 'Fatura No' sütunu bulunamadı!\n\nSeçili tab: '{current_tab_name}'\nMevcut sütunlar:\n{chr(10).join(headers)}\n\nLütfen doğru tab'ı seçtiğinizden emin olun.")
                return
            
            fatura_no_col_idx = headers.index(fatura_no_column)

            # Tablo verilerini al
            data = []
            for i in range(current_table.rowCount()):
                row_data = []
                for j in range(current_table.columnCount()):
                    item = current_table.item(i, j)
                    row_data.append(item.text() if item and item.text() else "")
                data.append(row_data)
            
            df_output = pd.DataFrame(data, columns=headers)

            # 7. Eşleşen satırları filtrele (Fatura No'ya göre)
            # Fatura numaralarını temiz string formatına dönüştür (.0 sorunu için)
            fatura_nolari_str = [self.convert_float_to_clean_string(num) for num in fatura_nolari]
            # df_output'daki fatura numaralarını da aynı formatta dönüştür
            df_output['Fatura No'] = df_output['Fatura No'].apply(self.convert_float_to_clean_string)
            filtered_data = df_output[df_output['Fatura No'].isin(fatura_nolari_str)].copy()

            if filtered_data.empty:
                QMessageBox.warning(self, "Uyarı", "Eşleşen fatura kaydı bulunamadı!")
                return

            # 8. Prosap Sas Kalem no sütununu esnek şekilde bul
            prosap_column = self.find_column_name(list(filtered_data.columns), "Prosap Sas Kalem no")
            
            if not prosap_column:
                # Daha detaylı hata mesajı
                column_list = "\n• ".join(filtered_data.columns.tolist())
                QMessageBox.critical(self, "Sütun Hatası", 
                    f"❌ 'Prosap Sas Kalem no' sütunu bulunamadı!\n\n"
                    f"📍 Aktif Tab: '{current_tab_name}'\n"
                    f"📊 Toplam sütun sayısı: {len(filtered_data.columns)}\n"
                    f"📋 Mevcut sütunlar:\n• {column_list}\n\n"
                    f"💡 Çözüm önerileri:\n"
                    f"   1. 'İrsaliyeler' tabını seçin\n"
                    f"   2. 'Mikro Güncelle' ile verileri yenileyin\n"
                    f"   3. Google Sheets bağlantısını kontrol edin\n"
                    f"   4. Sütun adlarının doğru olduğundan emin olun"
                )
                return
            
            # 9. Sadece "1" ile başlayan satırları tut
            filtered_data = filtered_data[filtered_data[prosap_column].astype(str).str.startswith("1", na=False)].copy()
            
            if filtered_data.empty:
                QMessageBox.warning(self, "Uyarı", f"'{prosap_column}' sütununda '1' ile başlayan kayıt bulunamadı!")
                return

            # 10. BagKodu verilerini Google Sheets'ten oku
            self.status_label.setText("🔄 BagKodu verileri Google Sheets'ten alınıyor...")
            from PyQt5.QtWidgets import QApplication
            QApplication.processEvents()
            
            try:
                if not self.gsheets_url:
                    QMessageBox.critical(self, "Konfigürasyon Hatası", 
                        f"❌ Google Sheets bağlantısı kurulamadı!\n\n"
                        f"💡 Çözüm önerileri:\n"
                        f"   1. PRGsheet/Ayar sayfasında SPREADSHEET_ID'nin tanımlı olduğundan emin olun\n"
                        f"   2. Google Sheets erişim izinlerini kontrol edin\n"
                        f"   3. Internet bağlantınızı kontrol edin"
                    )

            # Google Sheets'ten BagKodu sayfasını oku
                response = requests.get(self.gsheets_url, timeout=30)
                response.raise_for_status()
                
                bagKodu_df = pd.read_excel(BytesIO(response.content), sheet_name="BagKodu")
                
                if bagKodu_df.empty:
                    QMessageBox.warning(self, "Veri Hatası", 
                        f"❌ Google Sheets'teki 'BagKodu' sayfası boş!\n\n"
                        f"💡 Çözüm önerileri:\n"
                        f"   1. Google Sheets'te 'BagKodu' sayfasında veri olduğundan emin olun\n"
                        f"   2. Sayfa adının tam olarak 'BagKodu' olduğunu kontrol edin\n"
                        f"   3. Google Sheets'te gerekli sütunların mevcut olduğunu kontrol edin"
                    )
                    return
                    
                self.status_label.setText(f"✅ BagKodu verileri yüklendi - {len(bagKodu_df)} kayıt")
                from PyQt5.QtWidgets import QApplication
                QApplication.processEvents()
                
            except requests.exceptions.RequestException as e:
                QMessageBox.critical(self, "Bağlantı Hatası", 
                    f"❌ Google Sheets'e bağlanırken hata oluştu!\n\n"
                    f"🔍 Hata: {str(e)}\n\n"
                    f"💡 Çözüm önerileri:\n"
                    f"   1. Internet bağlantınızı kontrol edin\n"
                    f"   2. Google Sheets URL'sinin doğru olduğundan emin olun\n"
                    f"   3. Google Sheets'in herkese açık olduğunu kontrol edin\n"
                    f"   4. Güvenlik duvarı ayarlarınızı kontrol edin"
                )
                return
            except Exception as e:
                QMessageBox.critical(self, "Veri Okuma Hatası", 
                    f"❌ Google Sheets'ten BagKodu verileri okunurken hata oluştu!\n\n"
                    f"🔍 Hata: {str(e)}\n\n"
                    f"💡 Çözüm önerileri:\n"
                    f"   1. Google Sheets'te 'BagKodu' adlı bir sayfa olduğundan emin olun\n"
                    f"   2. Sayfa adının tam olarak 'BagKodu' olduğunu kontrol edin\n"
                    f"   3. Sayfada gerekli sütunların (bagKodum, Malzeme, malzemeKodu) mevcut olduğunu kontrol edin\n"
                    f"   4. Google Sheets formatının Excel uyumlu olduğunu kontrol edin"
                )
                return

            # 11. Prosap Sas Kalem no'yu parçala
            split_data = filtered_data[prosap_column].str.split("-", n=1, expand=True)
            filtered_data["Satış belgesi"] = split_data[0]
            filtered_data["Kalem"] = split_data[1]

            # 12. Kalem bilgilerini işle
            filtered_data['Kalem'] = filtered_data['Kalem'].fillna('0')
            filtered_data['Kalem - Metin'] = filtered_data['Kalem'].astype(str)
            filtered_data['Kalem'] = pd.to_numeric(filtered_data['Kalem'], errors='coerce').fillna(0).astype(int)

            # 13. BagKoduBekleyen oluştur
            filtered_data['BagKoduBekleyen'] = filtered_data.apply(
                lambda row: f"{row['Satış belgesi']}00{row['Kalem - Metin']}" if row['Kalem'] >= 1000 else (
                    f"{row['Satış belgesi']}0000{row['Kalem - Metin']}" if row['Kalem'] < 100 else
                    f"{row['Satış belgesi']}000{row['Kalem - Metin']}"), axis=1)

            # 14. BagKodu ile birleştir
            # Sayısal değerleri temiz string formatına dönüştür (.0 sorunu için)
            filtered_data['BagKoduBekleyen'] = filtered_data['BagKoduBekleyen'].apply(self.convert_float_to_clean_string)
            bagKodu_df['bagKodum'] = bagKodu_df['bagKodum'].apply(self.convert_float_to_clean_string)
            merged_df = filtered_data.merge(bagKodu_df, left_on='BagKoduBekleyen', right_on='bagKodum', how='left')

            # 15. Malzeme Kodu oluştur
            # Malzeme kodunu temiz string formatına dönüştür (.0 sorunu için)
            merged_df['Malzeme'] = merged_df['Malzeme'].apply(self.convert_float_to_clean_string)
            merged_df['Malzeme Kodu'] = merged_df.apply(
                lambda row: f"{row['Malzeme']}-0" if pd.isna(row['malzemeKodu']) else row['malzemeKodu'], axis=1)

            # 16. Fatura No'ya göre filtrele (esnek sütun adı kullan)
            merged_df = merged_df.dropna(subset=[fatura_no_column])

            # 17. Sayısal sütunları dönüştür - esnek sütun eşleştirme ile
            vergi_column = self.find_column_name(list(merged_df.columns), "Vergi")
            net_tutar_column = self.find_column_name(list(merged_df.columns), "Net Tutar")
            miktar_column = self.find_column_name(list(merged_df.columns), "Miktar")
            fatura_tarihi_column = self.find_column_name(list(merged_df.columns), "Fatura Tarihi")
            
            if vergi_column:
                merged_df['vergi_oran'] = merged_df[vergi_column].astype(str).str.replace('%', '').str.replace(',', '.').astype(float)

            # Net Tutar ve Miktar'ı sayısal formata çevir
            if net_tutar_column:
                # Net Tutar sütununu float'a çevir
                merged_df['Net Tutar'] = pd.to_numeric(merged_df[net_tutar_column], errors='coerce')

                # ÖNEMLI: Tam sayılara 0.01111 ekle (CSV formatı için)
                def add_offset(x):
                    if pd.notna(x) and x != 0:
                        val = float(x)
                        # Tam sayı mı kontrol et
                        if val == int(val):
                            return val + 0.01111
                        return val
                    return x

                merged_df['Net Tutar'] = merged_df['Net Tutar'].apply(add_offset)

            if miktar_column:
                # Miktar sütununu float'a çevir
                merged_df['Miktar'] = pd.to_numeric(merged_df[miktar_column], errors='coerce')

                # Birim Fiyat hesapla: Net Tutar / Miktar (round kaldırıldı)
                if net_tutar_column:
                    merged_df['Birim Fiyat'] = merged_df['Net Tutar'] / merged_df['Miktar']

            # 18. CSV dosyalarını oluştur
            output_dir = "D:/GoogleDrive/"
            os.makedirs(output_dir, exist_ok=True)

            created_files = []
            for fatura_no, group in merged_df.groupby(fatura_no_column):
                fatura_tarihi = group[fatura_tarihi_column].iloc[0] if fatura_tarihi_column and fatura_tarihi_column in group.columns else None
                tarih_str = "tarihyok"
                if pd.notna(fatura_tarihi):
                    try:
                        tarih_obj = pd.to_datetime(fatura_tarihi)
                        tarih_str = tarih_obj.strftime('%d %m %Y')
                    except:
                        pass
                
                filename = f"~ {tarih_str} - {fatura_no}.csv"
                full_path = os.path.join(output_dir, filename)
                
                # Gerekli sütunları kontrol et
                required_columns = ['Malzeme Kodu', 'Miktar', 'Birim Fiyat']
                available_columns = [col for col in required_columns if col in group.columns]
                available_columns.append(prosap_column)  # Prosap sütunu da ekle
                
                output_data = group[available_columns].copy()
                for i in range(1, 6):
                    output_data[f'BosSutun{i}'] = ''

                # Sütun sırasını düzenle
                final_columns = ['Malzeme Kodu', 'Miktar', 'Birim Fiyat', 'BosSutun1', 'BosSutun2', 'BosSutun3', 'BosSutun4', 'BosSutun5', prosap_column]
                output_data = output_data[[col for col in final_columns if col in output_data.columns]]

                # decimal=',' ile . → , değişimi, float_format='%.4f' ile 4 ondalık
                # 13410.01111 → "13410,0111", 3256.345 → "3256,3450"
                output_data.to_csv(full_path, index=False, sep=';', encoding='utf-8', decimal=',', float_format='%.4f', header=False)
                created_files.append(filename)

            self.status_label.setText("✅ CSV dosyaları başarıyla oluşturuldu")
            QMessageBox.information(self, "İşlem Tamamlandı", f"Toplam {len(created_files)} fatura için CSV dosyaları oluşturuldu:\n\n" + "\n".join(created_files))
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            QMessageBox.critical(self, "İşlem Hatası", 
                f"❌ İrsaliye Kaydet işlemi sırasında beklenmeyen bir hata oluştu!\n\n"
                f"🔍 Hata detayı: {str(e)}\n\n"
                f"💡 Çözüm önerileri:\n"
                f"   1. Internet bağlantınızı kontrol edin\n"
                f"   2. Google Sheets erişim izinlerini kontrol edin\n"
                f"   3. Excel dosyasının formatını kontrol edin\n"
                f"   4. Veriler yüklenene kadar bekleyin ve tekrar deneyin\n"
                f"   5. Programı yeniden başlatmayı deneyin\n\n"
                f"📧 Hata devam ederse, bu detayları sistem yöneticinizle paylaşın:\n{str(e)}"
            )
            self.status_label.setText(f"❌ İşlem hatası: {str(e)}")
            print(f"İrsaliye Kaydet Hatası - Detaylı Log:\n{error_details}")
    
    def set_buttons_enabled(self, enabled: bool):
        """Butonları aktif/pasif yap"""
        self.mikro_button.setEnabled(enabled)
        self.refresh_button.setEnabled(enabled)
        self.stok_aktar_button.setEnabled(enabled)
        self.irsaliye_kaydet_button.setEnabled(enabled)