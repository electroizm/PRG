"""
Ayarlar Modülü - PRGsheet Ayar ve Mail sayfası düzenleme
"""

import os
import sys
import pandas as pd
import requests
from io import BytesIO

# Üst dizini Python path'e ekle (central_config için)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Central config import
from central_config import CentralConfigManager

# Google Sheets API - Service Account
try:
    import gspread
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False
    print("⚠️ Google Sheets API paketleri yüklü değil. Kaydetme özelliği çalışmayacak.")
    print("Yüklemek için: pip install gspread google-auth")

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QTableWidget, QTableWidgetItem, QHeaderView,
                             QAbstractItemView, QLabel, QMessageBox, QTabWidget, QShortcut)
from PyQt5.QtGui import QFont, QColor, QKeySequence


class AyarlarApp(QWidget):
    def __init__(self):
        super().__init__()
        self.spreadsheet_id = None
        self.gsheets_url = self._load_gsheets_url()

        # Ayarlar sekmesi için veriler
        self.ayar_df = pd.DataFrame()
        self.ayar_original = None

        # Mail sekmesi için veriler
        self.mail_df = pd.DataFrame()
        self.mail_original = None

        # NoRisk sekmesi için veriler
        self.norisk_df = pd.DataFrame()
        self.norisk_original = None

        # Lazy loading için flag'ler
        self._ayar_loaded = False
        self._mail_loaded = False
        self._norisk_loaded = False

        self.setup_ui()

    def _load_gsheets_url(self):
        """Google Sheets SPREADSHEET_ID'sini yükle - Service Account"""
        try:
            config_manager = CentralConfigManager()
            self.spreadsheet_id = config_manager.MASTER_SPREADSHEET_ID
            if not self.spreadsheet_id:
                return None
            return f"https://docs.google.com/spreadsheets/d/{self.spreadsheet_id}/export?format=xlsx"
        except Exception as e:
            print(f"URL yükleme hatası: {e}")
            return None

    def showEvent(self, event):
        """Widget ilk gösterildiğinde aktif sekmenin verilerini yükle (lazy loading)"""
        super().showEvent(event)
        # İlk açılışta aktif sekmeyi yükle
        QTimer.singleShot(100, self._load_active_tab)

        # Tab değişikliklerini dinle
        self.tab_widget.currentChanged.connect(self._on_tab_changed)

    def _on_tab_changed(self, index):
        """Sekme değiştiğinde ilgili veriyi yükle"""
        if index == 0 and not self._ayar_loaded:
            self._ayar_loaded = True
            QTimer.singleShot(50, self.load_ayar_data)
        elif index == 1 and not self._mail_loaded:
            self._mail_loaded = True
            QTimer.singleShot(50, self.load_mail_data)
        elif index == 2 and not self._norisk_loaded:
            self._norisk_loaded = True
            QTimer.singleShot(50, self.load_norisk_data)

    def _load_active_tab(self):
        """Aktif sekmenin verilerini yükle"""
        current_index = self.tab_widget.currentIndex()
        self._on_tab_changed(current_index)

    def setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(5)

        # QTabWidget oluştur
        self.tab_widget = QTabWidget()
        self.tab_widget.setStyleSheet("""
            QTabWidget::pane {
                border: 2px solid #cccccc;
                background: white;
                border-radius: 8px;
            }
            QTabBar::tab {
                background: #e0e0e0;
                border: 2px solid #cccccc;
                padding: 10px 35px;
                font-size: 18px;
                font-weight: bold;
                min-width: 150px;
                min-height: 30px;
                color: #666666;
                margin-right: 4px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
            }
            QTabBar::tab:selected {
                background: #2196F3;
                border-bottom-color: #2196F3;
                color: #ffffff;
                border: 2px solid #2196F3;
            }
            QTabBar::tab:hover {
                background: #c0c0c0;
                color: #333333;
            }
        """)

        # Ayarlar sekmesi
        self.ayarlar_tab = self._create_ayarlar_tab()
        self.tab_widget.addTab(self.ayarlar_tab, "Ayarlar")

        # Mail sekmesi
        self.mail_tab = self._create_mail_tab()
        self.tab_widget.addTab(self.mail_tab, "Mail")

        # NoRisk sekmesi
        self.norisk_tab = self._create_norisk_tab()
        self.tab_widget.addTab(self.norisk_tab, "NoRisk")

        main_layout.addWidget(self.tab_widget)

    def _create_ayarlar_tab(self):
        """Ayarlar sekmesini oluştur"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Başlık ve Butonlar
        header_layout = QHBoxLayout()

        title = QLabel("Ayar Verileri")
        title.setStyleSheet("font-size: 16px; font-weight: bold; color: #333333;")

        self.ayar_refresh_btn = QPushButton("Yenile")
        self.ayar_refresh_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
        """)

        self.ayar_save_btn = QPushButton("Kaydet")
        self.ayar_save_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)

        header_layout.addWidget(title)
        header_layout.addStretch()
        header_layout.addWidget(self.ayar_refresh_btn)
        header_layout.addWidget(self.ayar_save_btn)

        # Tablo
        self.ayar_table = QTableWidget()
        self.ayar_table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)
        self.ayar_table.verticalHeader().setDefaultSectionSize(self.ayar_table.verticalHeader().defaultSectionSize() + 2)
        self.ayar_table.setStyleSheet("""
            QTableWidget {
                font-size: 15px;
                font-weight: bold;
                background-color: white;
                gridline-color: #d0d0d0;
                selection-background-color: #e3f2fd;
                selection-color: #000000;
            }
            QTableWidget::item:selected {
                background-color: #e3f2fd;
                color: #000000;
            }
            QHeaderView::section {
                background-color: #1a1a1a;
                color: #ffffff;
                padding: 8px;
                border: 1px solid #404040;
                font-weight: bold;
                font-size: 14px;
            }
            QTableCornerButton::section {
                background-color: #1a1a1a;
                border: 1px solid #404040;
            }
        """)

        # Ctrl+C kısayolu - Ayar tablosu
        self.copy_shortcut_ayar = QShortcut(QKeySequence("Ctrl+C"), self.ayar_table)
        self.copy_shortcut_ayar.activated.connect(lambda: self.copy_table_selection(self.ayar_table))

        # Status Label
        self.ayar_status = QLabel("Hazır")
        self.ayar_status.setStyleSheet("""
            QLabel {
                color: #666666;
                padding: 8px;
                background-color: #f5f5f5;
                border-top: 2px solid #cccccc;
                font-size: 13px;
            }
        """)

        layout.addLayout(header_layout)
        layout.addWidget(self.ayar_table, 1)
        layout.addWidget(self.ayar_status)

        # Sinyaller
        self.ayar_refresh_btn.clicked.connect(self.load_ayar_data)
        self.ayar_save_btn.clicked.connect(self.save_ayar_changes)

        return tab

    def _create_mail_tab(self):
        """Mail sekmesini oluştur"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Başlık ve Butonlar
        header_layout = QHBoxLayout()

        title = QLabel("e-Posta Bilgileri")
        title.setStyleSheet("font-size: 16px; font-weight: bold; color: #333333;")

        self.mail_refresh_btn = QPushButton("Yenile")
        self.mail_refresh_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
        """)

        self.mail_save_btn = QPushButton("Kaydet")
        self.mail_save_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)

        header_layout.addWidget(title)
        header_layout.addStretch()
        header_layout.addWidget(self.mail_refresh_btn)
        header_layout.addWidget(self.mail_save_btn)

        # Tablo
        self.mail_table = QTableWidget()
        self.mail_table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)
        self.mail_table.verticalHeader().setDefaultSectionSize(self.mail_table.verticalHeader().defaultSectionSize() + 2)
        self.mail_table.setStyleSheet("""
            QTableWidget {
                font-size: 15px;
                font-weight: bold;
                background-color: white;
                gridline-color: #d0d0d0;
                selection-background-color: #e3f2fd;
                selection-color: #000000;
            }
            QTableWidget::item:selected {
                background-color: #e3f2fd;
                color: #000000;
            }
            QHeaderView::section {
                background-color: #1a1a1a;
                color: #ffffff;
                padding: 8px;
                border: 1px solid #404040;
                font-weight: bold;
                font-size: 14px;
            }
            QTableCornerButton::section {
                background-color: #1a1a1a;
                border: 1px solid #404040;
            }
        """)

        # Ctrl+C kısayolu - Mail tablosu
        self.copy_shortcut_mail = QShortcut(QKeySequence("Ctrl+C"), self.mail_table)
        self.copy_shortcut_mail.activated.connect(lambda: self.copy_table_selection(self.mail_table))

        # Status Label
        self.mail_status = QLabel("Hazır")
        self.mail_status.setStyleSheet("""
            QLabel {
                color: #666666;
                padding: 8px;
                background-color: #f5f5f5;
                border-top: 2px solid #cccccc;
                font-size: 13px;
            }
        """)

        layout.addLayout(header_layout)
        layout.addWidget(self.mail_table, 1)
        layout.addWidget(self.mail_status)

        # Sinyaller
        self.mail_refresh_btn.clicked.connect(self.load_mail_data)
        self.mail_save_btn.clicked.connect(self.save_mail_changes)

        return tab

    def load_ayar_data(self):
        """Ayar sayfasından verileri yükle"""
        try:
            self.ayar_status.setText("📊 Ayar sayfası yükleniyor...")
            self.ayar_refresh_btn.setEnabled(False)
            self.ayar_save_btn.setEnabled(False)

            from PyQt5.QtWidgets import QApplication
            QApplication.processEvents()

            if not self.gsheets_url:
                self.ayar_df = pd.DataFrame()
                self.populate_ayar_table()
                self.ayar_status.setText("❌ PRGsheet/Ayar sayfasında SPREADSHEET_ID bulunamadı")
                return

            response = requests.get(self.gsheets_url, timeout=30)

            if response.status_code != 200:
                self.ayar_df = pd.DataFrame()
                self.populate_ayar_table()
                self.ayar_status.setText(f"❌ HTTP Hatası: {response.status_code}")
                return

            # Ayar sayfasını oku
            self.ayar_df = pd.read_excel(BytesIO(response.content), sheet_name="Ayar")
            self.ayar_original = self.ayar_df.copy()

            self.populate_ayar_table()
            self.ayar_status.setText(f"✅ {len(self.ayar_df)} ayar yüklendi")

        except Exception as e:
            self.ayar_df = pd.DataFrame()
            self.populate_ayar_table()
            self.ayar_status.setText(f"❌ Yükleme hatası: {str(e)}")
        finally:
            self.ayar_refresh_btn.setEnabled(True)
            self.ayar_save_btn.setEnabled(True)

    def populate_ayar_table(self):
        """Ayar tablosunu doldur - Key kilitli, Value ve Description düzenlenebilir"""
        if self.ayar_df.empty:
            self.ayar_table.setRowCount(0)
            self.ayar_table.setColumnCount(0)
            return

        # Ekstra boş satırlar ekle (yeni satır eklemek için)
        extra_rows = 50
        total_rows = len(self.ayar_df) + extra_rows

        self.ayar_table.setRowCount(total_rows)
        self.ayar_table.setColumnCount(len(self.ayar_df.columns))
        column_names = self.ayar_df.columns.tolist()
        self.ayar_table.setHorizontalHeaderLabels(column_names)

        # Satır numaralarını göster
        self.ayar_table.verticalHeader().setVisible(True)

        self.ayar_table.setAlternatingRowColors(False)  # Alternating colors kapatıldı
        self.ayar_table.setSortingEnabled(False)
        self.ayar_table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.ayar_table.setSelectionMode(QAbstractItemView.SingleSelection)

        # "App Name" sütununu gizle (zaten hep "Global")
        if "App Name" in column_names:
            app_name_idx = column_names.index("App Name")
            self.ayar_table.hideColumn(app_name_idx)

        for i, row in self.ayar_df.iterrows():
            # İlk sütunun değerini kontrol et (# ile başlıyor mu - ayırıcı satır)
            first_col_value = "" if pd.isna(row.iloc[0]) else str(row.iloc[0])
            is_separator = first_col_value.startswith("#")

            for j, value in enumerate(row):
                display_value = "" if pd.isna(value) else str(value)
                item = QTableWidgetItem(display_value)
                col_name = column_names[j]

                font = QFont('Segoe UI', 12)
                if is_separator:
                    font.setBold(True)
                item.setFont(font)

                # Ayırıcı satır ise (# ile başlayan) - Tüm satır sarı ve kilitli
                if is_separator:
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                    item.setBackground(QColor("#ffeb3b"))  # Sarı arka plan
                    item.setForeground(QColor("#000000"))  # Siyah yazı
                # "Key" sütunu kilitli
                elif col_name == "Key":
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                    item.setBackground(QColor("#f5f5f5"))  # Açık gri arka plan
                    item.setForeground(QColor("#666666"))  # Koyu gri yazı
                # "Value" ve "Description" sütunları düzenlenebilir
                elif col_name in ["Value", "Description"]:
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsEditable | Qt.ItemIsSelectable)
                    item.setBackground(QColor("#ffffff"))  # Beyaz arka plan
                    item.setForeground(QColor("#000000"))  # Siyah yazı
                # Diğer sütunlar (App Name gibi) - kilitli
                else:
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                    item.setBackground(QColor("#f5f5f5"))  # Açık gri arka plan
                    item.setForeground(QColor("#666666"))  # Koyu gri yazı

                self.ayar_table.setItem(i, j, item)

        # Boş satırları doldur (ekstra satırlar - düzenlenebilir)
        for i in range(len(self.ayar_df), total_rows):
            for j, col_name in enumerate(column_names):
                item = QTableWidgetItem("")
                font = QFont('Segoe UI', 12)
                item.setFont(font)

                # "Key" sütunu kilitli
                if col_name == "Key":
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                    item.setBackground(QColor("#f5f5f5"))
                    item.setForeground(QColor("#666666"))
                # "Value" ve "Description" sütunları düzenlenebilir
                elif col_name in ["Value", "Description"]:
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsEditable | Qt.ItemIsSelectable)
                    item.setBackground(QColor("#ffffff"))
                    item.setForeground(QColor("#000000"))
                # Diğer sütunlar
                else:
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                    item.setBackground(QColor("#f5f5f5"))
                    item.setForeground(QColor("#666666"))

                self.ayar_table.setItem(i, j, item)

        # Header ayarları - Dinamik genişlik
        header = self.ayar_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        header.setStretchLastSection(True)

        for i in range(self.ayar_table.rowCount()):
            self.ayar_table.setRowHeight(i, 40)

    def save_ayar_changes(self):
        """Ayar değişikliklerini Google Sheets'e Kaydet"""
        # Tablodan tüm verileri al (boş satırlar dahil)
        all_rows = []
        for i in range(self.ayar_table.rowCount()):
            row_data = []
            is_empty_row = True
            for j in range(self.ayar_table.columnCount()):
                item = self.ayar_table.item(i, j)
                value = item.text() if item else ""
                row_data.append(value if value else None)
                if value.strip():  # Eğer boş değilse
                    is_empty_row = False

            # Sadece dolu satırları ekle
            if not is_empty_row:
                all_rows.append(row_data)

        if not all_rows:
            QMessageBox.warning(self, "Uyarı", "Kaydedilecek veri yok!")
            return

        # Yeni DataFrame oluştur
        column_names = self.ayar_df.columns.tolist() if not self.ayar_df.empty else [self.ayar_table.horizontalHeaderItem(i).text() for i in range(self.ayar_table.columnCount())]
        updated_df = pd.DataFrame(all_rows, columns=column_names)

        # Değişiklik kontrolü
        if self.ayar_original is not None:
            if updated_df.equals(self.ayar_original):
                QMessageBox.information(self, "Bilgi", "Herhangi bir değişiklik yapılmadı.")
                return

        # Onay iste
        reply = QMessageBox.question(
            self,
            "Kaydet",
            "Yaptığınız değişiklikler Google Sheets'in 'Ayar' sayfasına kaydedilecek.\n\n"
            "⚠️ DİKKAT: Bu işlem geri alınamaz!\n\n"
            "Devam etmek istiyor musunuz?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self._save_to_gsheets("Ayar", updated_df, self.ayar_status)
            # Başarılı kaydedilirse, güncellenen DataFrame'i kaydet
            self.ayar_df = updated_df
            self.ayar_original = updated_df.copy()

    def load_mail_data(self):
        """Mail sayfasından verileri yükle"""
        try:
            self.mail_status.setText("📊 Mail sayfası yükleniyor...")
            self.mail_refresh_btn.setEnabled(False)
            self.mail_save_btn.setEnabled(False)

            from PyQt5.QtWidgets import QApplication
            QApplication.processEvents()

            if not self.gsheets_url:
                self.mail_df = pd.DataFrame()
                self.populate_mail_table()
                self.mail_status.setText("❌ PRGsheet/Ayar sayfasında SPREADSHEET_ID bulunamadı")
                return

            response = requests.get(self.gsheets_url, timeout=30)

            if response.status_code != 200:
                self.mail_df = pd.DataFrame()
                self.populate_mail_table()
                self.mail_status.setText(f"❌ HTTP Hatası: {response.status_code}")
                return

            # Mail sayfasını oku
            self.mail_df = pd.read_excel(BytesIO(response.content), sheet_name="Mail")
            self.mail_original = self.mail_df.copy()

            self.populate_mail_table()
            self.mail_status.setText(f"✅ {len(self.mail_df)} mail ayarı yüklendi")

        except Exception as e:
            self.mail_df = pd.DataFrame()
            self.populate_mail_table()
            self.mail_status.setText(f"❌ Yükleme hatası: {str(e)}")
        finally:
            self.mail_refresh_btn.setEnabled(True)
            self.mail_save_btn.setEnabled(True)

    def populate_mail_table(self):
        """Mail tablosunu doldur - Belirli sütunlar kilitli"""
        if self.mail_df.empty:
            self.mail_table.setRowCount(0)
            self.mail_table.setColumnCount(0)
            return

        # Kilitli sütun adları
        locked_columns = ['sender_mail', 'smtp_server', 'password', 'bcc_email']

        # Ekstra boş satırlar ekle (yeni satır eklemek için)
        extra_rows = 50
        total_rows = len(self.mail_df) + extra_rows

        self.mail_table.setRowCount(total_rows)
        self.mail_table.setColumnCount(len(self.mail_df.columns))
        self.mail_table.setHorizontalHeaderLabels(self.mail_df.columns.tolist())

        # Satır numaralarını göster
        self.mail_table.verticalHeader().setVisible(True)

        self.mail_table.setAlternatingRowColors(False)  # Alternating colors kapatıldı
        self.mail_table.setSortingEnabled(False)
        self.mail_table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.mail_table.setSelectionMode(QAbstractItemView.SingleSelection)

        for i, row in self.mail_df.iterrows():
            for j, value in enumerate(row):
                display_value = "" if pd.isna(value) else str(value)
                item = QTableWidgetItem(display_value)

                font = QFont('Segoe UI', 12)
                item.setFont(font)

                # Sütun adını al
                column_name = self.mail_df.columns[j]

                # A sütunu (j=0) veya kilitli sütunlar düzenlenemez
                if j == 0 or column_name in locked_columns:
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                    item.setBackground(QColor("#f5f5f5"))  # Açık gri arka plan
                    item.setForeground(QColor("#666666"))  # Koyu gri yazı
                else:
                    # Diğer sütunlar düzenlenebilir
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsEditable | Qt.ItemIsSelectable)
                    item.setBackground(QColor("#ffffff"))  # Beyaz arka plan
                    item.setForeground(QColor("#000000"))  # Siyah yazı

                self.mail_table.setItem(i, j, item)

        # Boş satırları doldur (ekstra satırlar - düzenlenebilir)
        for i in range(len(self.mail_df), total_rows):
            for j in range(len(self.mail_df.columns)):
                item = QTableWidgetItem("")
                font = QFont('Segoe UI', 12)
                item.setFont(font)

                column_name = self.mail_df.columns[j]

                # A sütunu (j=0) veya kilitli sütunlar düzenlenemez
                if j == 0 or column_name in locked_columns:
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                    item.setBackground(QColor("#f5f5f5"))
                    item.setForeground(QColor("#666666"))
                else:
                    # Diğer sütunlar düzenlenebilir
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsEditable | Qt.ItemIsSelectable)
                    item.setBackground(QColor("#ffffff"))
                    item.setForeground(QColor("#000000"))

                self.mail_table.setItem(i, j, item)

        # Header ayarları - Dinamik genişlik
        header = self.mail_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        header.setStretchLastSection(True)

        for i in range(self.mail_table.rowCount()):
            self.mail_table.setRowHeight(i, 40)

    def save_mail_changes(self):
        """Mail değişikliklerini Google Sheets'e kaydet"""
        # Tablodan tüm verileri al (boş satırlar dahil)
        all_rows = []
        for i in range(self.mail_table.rowCount()):
            row_data = []
            is_empty_row = True
            for j in range(self.mail_table.columnCount()):
                item = self.mail_table.item(i, j)
                value = item.text() if item else ""
                row_data.append(value if value else None)
                if value.strip():  # Eğer boş değilse
                    is_empty_row = False

            # Sadece dolu satırları ekle
            if not is_empty_row:
                all_rows.append(row_data)

        if not all_rows:
            QMessageBox.warning(self, "Uyarı", "Kaydedilecek veri yok!")
            return

        # Yeni DataFrame oluştur
        column_names = self.mail_df.columns.tolist() if not self.mail_df.empty else [self.mail_table.horizontalHeaderItem(i).text() for i in range(self.mail_table.columnCount())]
        updated_df = pd.DataFrame(all_rows, columns=column_names)

        # Değişiklik kontrolü
        if self.mail_original is not None:
            if updated_df.equals(self.mail_original):
                QMessageBox.information(self, "Bilgi", "Herhangi bir değişiklik yapılmadı.")
                return

        # Onay iste
        reply = QMessageBox.question(
            self,
            "Kaydet",
            "Yaptığınız değişiklikler Google Sheets'in 'Mail' sayfasına kaydedilecek.\n\n"
            "⚠️ DİKKAT: Bu işlem geri alınamaz!\n\n"
            "Devam etmek istiyor musunuz?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self._save_to_gsheets("Mail", updated_df, self.mail_status)
            # Başarılı kaydedilirse, güncellenen DataFrame'i kaydet
            self.mail_df = updated_df
            self.mail_original = updated_df.copy()

    def _create_norisk_tab(self):
        """NoRisk sekmesini oluştur"""
        tab = QWidget()
        layout = QVBoxLayout(tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Başlık ve Butonlar
        header_layout = QHBoxLayout()

        title = QLabel("NoRisk Verileri")
        title.setStyleSheet("font-size: 16px; font-weight: bold; color: #333333;")

        self.norisk_refresh_btn = QPushButton("Verileri Yenile")
        self.norisk_refresh_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
        """)

        self.norisk_save_btn = QPushButton("Değişiklikleri Kaydet")
        self.norisk_save_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)

        header_layout.addWidget(title)
        header_layout.addStretch()
        header_layout.addWidget(self.norisk_refresh_btn)
        header_layout.addWidget(self.norisk_save_btn)

        # Tablo
        self.norisk_table = QTableWidget()
        self.norisk_table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)
        self.norisk_table.verticalHeader().setDefaultSectionSize(self.norisk_table.verticalHeader().defaultSectionSize() + 2)
        self.norisk_table.setStyleSheet("""
            QTableWidget {
                font-size: 15px;
                font-weight: bold;
                background-color: white;
                gridline-color: #d0d0d0;
                selection-background-color: #e3f2fd;
                selection-color: #000000;
            }
            QTableWidget::item:selected {
                background-color: #e3f2fd;
                color: #000000;
            }
            QHeaderView::section {
                background-color: #1a1a1a;
                color: #ffffff;
                padding: 8px;
                border: 1px solid #404040;
                font-weight: bold;
                font-size: 14px;
            }
            QTableCornerButton::section {
                background-color: #1a1a1a;
                border: 1px solid #404040;
            }
        """)

        # Ctrl+C kısayolu - NoRisk tablosu
        self.copy_shortcut_norisk = QShortcut(QKeySequence("Ctrl+C"), self.norisk_table)
        self.copy_shortcut_norisk.activated.connect(lambda: self.copy_table_selection(self.norisk_table))

        # Status Label
        self.norisk_status = QLabel("Hazır")
        self.norisk_status.setStyleSheet("""
            QLabel {
                color: #666666;
                padding: 8px;
                background-color: #f5f5f5;
                border-top: 2px solid #cccccc;
                font-size: 13px;
            }
        """)

        layout.addLayout(header_layout)
        layout.addWidget(self.norisk_table, 1)
        layout.addWidget(self.norisk_status)

        # Sinyaller
        self.norisk_refresh_btn.clicked.connect(self.load_norisk_data)
        self.norisk_save_btn.clicked.connect(self.save_norisk_changes)

        return tab

    def load_norisk_data(self):
        """NoRisk sayfasından verileri yükle"""
        try:
            self.norisk_status.setText("📊 NoRisk sayfası yükleniyor...")
            self.norisk_refresh_btn.setEnabled(False)
            self.norisk_save_btn.setEnabled(False)

            from PyQt5.QtWidgets import QApplication
            QApplication.processEvents()

            if not self.gsheets_url:
                self.norisk_df = pd.DataFrame()
                self.populate_norisk_table()
                self.norisk_status.setText("❌ PRGsheet/Ayar sayfasında SPREADSHEET_ID bulunamadı")
                return

            response = requests.get(self.gsheets_url, timeout=30)

            if response.status_code != 200:
                self.norisk_df = pd.DataFrame()
                self.populate_norisk_table()
                self.norisk_status.setText(f"❌ HTTP Hatası: {response.status_code}")
                return

            # NoRisk sayfasını oku
            self.norisk_df = pd.read_excel(BytesIO(response.content), sheet_name="NoRisk")
            self.norisk_original = self.norisk_df.copy()

            self.populate_norisk_table()
            self.norisk_status.setText(f"✅ {len(self.norisk_df)} kayıt yüklendi")

        except Exception as e:
            self.norisk_df = pd.DataFrame()
            self.populate_norisk_table()
            self.norisk_status.setText(f"❌ Yükleme hatası: {str(e)}")
        finally:
            self.norisk_refresh_btn.setEnabled(True)
            self.norisk_save_btn.setEnabled(True)

    def populate_norisk_table(self):
        """NoRisk tablosunu doldur - B sütunu kilitli, A ve diğerleri düzenlenebilir"""
        if self.norisk_df.empty:
            self.norisk_table.setRowCount(0)
            self.norisk_table.setColumnCount(0)
            return

        # Ekstra boş satırlar ekle (yeni satır eklemek için)
        extra_rows = 50
        total_rows = len(self.norisk_df) + extra_rows

        self.norisk_table.setRowCount(total_rows)
        self.norisk_table.setColumnCount(len(self.norisk_df.columns))
        self.norisk_table.setHorizontalHeaderLabels(self.norisk_df.columns.tolist())

        # Satır numaralarını göster
        self.norisk_table.verticalHeader().setVisible(True)

        self.norisk_table.setAlternatingRowColors(False)
        self.norisk_table.setSortingEnabled(False)
        self.norisk_table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.norisk_table.setSelectionMode(QAbstractItemView.SingleSelection)

        for i, row in self.norisk_df.iterrows():
            for j, value in enumerate(row):
                display_value = "" if pd.isna(value) else str(value)
                item = QTableWidgetItem(display_value)

                font = QFont('Segoe UI', 12)
                item.setFont(font)

                # Tüm sütunlar düzenlenebilir
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsEditable | Qt.ItemIsSelectable)
                item.setBackground(QColor("#ffffff"))  # Beyaz arka plan
                item.setForeground(QColor("#000000"))  # Siyah yazı

                self.norisk_table.setItem(i, j, item)

        # Boş satırları doldur (ekstra satırlar - tüm sütunlar düzenlenebilir)
        for i in range(len(self.norisk_df), total_rows):
            for j in range(len(self.norisk_df.columns)):
                item = QTableWidgetItem("")
                font = QFont('Segoe UI', 12)
                item.setFont(font)

                # Tüm sütunlar düzenlenebilir
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsEditable | Qt.ItemIsSelectable)
                item.setBackground(QColor("#ffffff"))  # Beyaz arka plan
                item.setForeground(QColor("#000000"))  # Siyah yazı

                self.norisk_table.setItem(i, j, item)

        # Header ayarları - Dinamik genişlik
        header = self.norisk_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setStretchLastSection(True)

        # İlk yükleme için sütunları içeriğe göre boyutlandır
        self.norisk_table.resizeColumnsToContents()

        for i in range(self.norisk_table.rowCount()):
            self.norisk_table.setRowHeight(i, 40)

    def save_norisk_changes(self):
        """NoRisk değişikliklerini Google Sheets'e kaydet"""
        # Tablodan tüm verileri al (boş satırlar dahil)
        all_rows = []
        for i in range(self.norisk_table.rowCount()):
            row_data = []
            is_empty_row = True
            for j in range(self.norisk_table.columnCount()):
                item = self.norisk_table.item(i, j)
                value = item.text() if item else ""
                row_data.append(value if value else None)
                if value.strip():  # Eğer boş değilse
                    is_empty_row = False

            # Sadece dolu satırları ekle
            if not is_empty_row:
                all_rows.append(row_data)

        if not all_rows:
            QMessageBox.warning(self, "Uyarı", "Kaydedilecek veri yok!")
            return

        # Yeni DataFrame oluştur
        column_names = self.norisk_df.columns.tolist() if not self.norisk_df.empty else [self.norisk_table.horizontalHeaderItem(i).text() for i in range(self.norisk_table.columnCount())]
        updated_df = pd.DataFrame(all_rows, columns=column_names)

        # Değişiklik kontrolü
        if self.norisk_original is not None:
            if updated_df.equals(self.norisk_original):
                QMessageBox.information(self, "Bilgi", "Herhangi bir değişiklik yapılmadı.")
                return

        # Onay iste
        reply = QMessageBox.question(
            self,
            "Değişiklikleri Kaydet",
            "Yaptığınız değişiklikler Google Sheets'in 'NoRisk' sayfasına kaydedilecek.\n\n"
            "⚠️ DİKKAT: Bu işlem geri alınamaz!\n\n"
            "Devam etmek istiyor musunuz?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self._save_to_gsheets("NoRisk", updated_df, self.norisk_status)
            # Başarılı kaydedilirse, güncellenen DataFrame'i kaydet
            self.norisk_df = updated_df
            self.norisk_original = updated_df.copy()

    def _save_to_gsheets(self, sheet_name, dataframe, status_label):
        """Google Sheets'e kaydet"""
        try:
            # Google Sheets API kontrol
            if not GSPREAD_AVAILABLE:
                QMessageBox.critical(
                    self,
                    "Hata",
                    "Google Sheets API paketleri yüklü değil!\n\n"
                    "Aşağıdaki komutu çalıştırın:\n"
                    "pip install gspread google-auth google-auth-oauthlib google-auth-httplib2"
                )
                status_label.setText("❌ gspread paketi bulunamadı")
                return

            status_label.setText(f"💾 {sheet_name} sayfasına kaydediliyor...")

            # Google Sheets API kullanarak kaydet - Service Account
            config_manager = CentralConfigManager()
            client = config_manager.gc

            if not client:
                QMessageBox.critical(
                    self,
                    "Hata",
                    "Google Sheets bağlantısı kurulamadı!\n\n"
                    "Service Account credentials kontrolü yapın: service_account.json"
                )
                status_label.setText("❌ Google Sheets bağlantı hatası")
                return

            # Spreadsheet'i aç
            spreadsheet = client.open_by_key(self.spreadsheet_id)
            worksheet = spreadsheet.worksheet(sheet_name)

            # Sayfayı temizle
            worksheet.clear()

            # Yeni verileri yaz (header dahil)
            data_to_write = [dataframe.columns.tolist()] + dataframe.values.tolist()
            worksheet.update('A1', data_to_write)

            QMessageBox.information(
                self,
                "Başarılı",
                f"Değişiklikler Google Sheets '{sheet_name}' sayfasına başarıyla kaydedildi!"
            )

            status_label.setText(f"✅ {sheet_name} sayfası güncellendi")

        except Exception as e:
            QMessageBox.critical(
                self,
                "Kayıt Hatası",
                f"{sheet_name} sayfasına kaydedilirken hata oluştu:\n{str(e)}\n\n"
                "Google Sheets API credentials kontrolünü yapın."
            )
            status_label.setText(f"❌ Kayıt hatası: {str(e)}")

    def copy_table_selection(self, table):
        """Tablodaki seçili hücreyi/hücreleri kopyala"""
        from PyQt5.QtWidgets import QApplication
        
        selected_items = table.selectedItems()
        if not selected_items:
            return

        # Sadece ilk seçili öğeyi kopyala (basitlik için)
        text = selected_items[0].text()
        if text:
            QApplication.clipboard().setText(text)
            # Find status label if possible, or just pass silently
            # self.status_label equivalent is separate for each tab, so we might skip status update 
            # or try to find parent tab's status label.
            # Simplified: just copy. 
        else:
            pass
