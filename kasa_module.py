"""
Kasa Modülü
"""

import os
import sys
import pandas as pd
import requests
from io import BytesIO
from datetime import datetime

# Üst dizini Python path'e ekle (central_config için)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Central config import
from central_config import CentralConfigManager

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton, 
                             QTableWidget, QTableWidgetItem, QHeaderView, 
                             QAbstractItemView, QMenu, QProgressBar, QLabel,
                             QCheckBox, QComboBox, QInputDialog, QLineEdit,
                             QMessageBox, QApplication, QShortcut)
from PyQt5.QtGui import QFont, QColor, QIntValidator, QKeySequence


class KasaApp(QWidget):
    def __init__(self):
        super().__init__()
        self.df = pd.DataFrame()
        self.veri_cercevesi = pd.DataFrame()
        self.mikro_calisiyor = False
        self.gsheets_url = self._load_gsheets_url()

        # Mevcut tarihi al
        now = datetime.now()
        self.current_year = now.year
        self.current_month = now.month

        self.setup_ui()
        self.setup_connections()

        # Lazy loading için flag
        self._data_loaded = False

    def showEvent(self, event):
        """Widget ilk gösterildiğinde veri yükle (lazy loading)"""
        super().showEvent(event)
        if not self._data_loaded:
            self._data_loaded = True
            # UI render olduktan sonra veri yükle
            QTimer.singleShot(100, lambda: self.load_data(force_reload=False))

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
    
    def _load_password_from_Pass(self):
        """Pass sayfasından KasaApp için şifreyi yükle"""
        try:
            if not self.gsheets_url:
                return None
                
            response = requests.get(self.gsheets_url, timeout=30)
            if response.status_code != 200:
                return None
                
            from io import BytesIO
            Pass_df = pd.read_excel(BytesIO(response.content), sheet_name="Pass")
            
            # KasaApp için şifreyi bul
            kasa_row = Pass_df[Pass_df['Modul'] == 'KasaApp']
            if not kasa_row.empty:
                return str(kasa_row.iloc[0]['Password'])
            return None
            
        except Exception as e:
            print(f"Pass şifre yükleme hatası: {str(e)}")
            return None
    
    def _show_password_dialog(self):
        """Şifre doğrulama dialog'unu göster"""
        password, ok = QInputDialog.getText(
            self, 
            'Şifre Gerekli', 
            'Bu seçeneği değiştirmek için şifre giriniz:',
            QLineEdit.Password
        )
        
        if ok and password:
            correct_password = self._load_password_from_Pass()
            if correct_password and password == correct_password:
                return True
            else:
                # Yanlış şifre
                reply = QMessageBox.question(
                    self, 
                    'Yanlış Şifre', 
                    'Şifre yanlış! Tekrar denemek istiyor musunuz?',
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes
                )
                if reply == QMessageBox.Yes:
                    return self._show_password_dialog()  # Recursive çağrı
                else:
                    return False
        return False
    
    def _on_dekont_checkbox_clicked(self, checked):
        """Dekont checkbox tıklandığında şifre kontrolü"""
        if checked:
            # Şifre kontrolü yap
            if self._show_password_dialog():
                # Şifre doğru, filtreleme yap
                self.filter_table()
            else:
                # Şifre yanlış veya iptal, checkbox'ı kaldır (signal engellemeden)
                self.dekont_checkbox.blockSignals(True)
                self.dekont_checkbox.setChecked(False)
                self.dekont_checkbox.blockSignals(False)
        else:
            # Checkbox kaldırılıyor, filtreleme yap
            self.filter_table()
    
    def _on_alacak_checkbox_clicked(self, checked):
        """Alacak checkbox tıklandığında şifre kontrolü"""
        if checked:
            # Şifre kontrolü yap
            if self._show_password_dialog():
                # Şifre doğru, filtreleme yap
                self.filter_table()
            else:
                # Şifre yanlış veya iptal, checkbox'ı kaldır (signal engellemeden)
                self.alacak_checkbox.blockSignals(True)
                self.alacak_checkbox.setChecked(False)
                self.alacak_checkbox.blockSignals(False)
        else:
            # Checkbox kaldırılıyor, filtreleme yap
            self.filter_table()
    
    def setup_ui(self):
        # Light theme - Force white background
        self.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                color: #000000;
            }
        """)
        self.setAutoFillBackground(True)
        palette = self.palette()
        palette.setColor(self.backgroundRole(), QColor("#ffffff"))
        self.setPalette(palette)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # Yıl seçimi için checkbox'lar
        self.year_checkbox_layout = QHBoxLayout()
        
        # Yılları checkbox olarak ekle
        self.year_checkboxes = []
        for year in range(self.current_year - 2, self.current_year + 1):  # Son 2 yıl ve bulunduğumuz yıl
            checkbox = QCheckBox(str(year))
            checkbox.setChecked(year == self.current_year)  # Varsayılan olarak bulunduğumuz yıl seçili
            checkbox.setStyleSheet("""                QCheckBox { 
                    font-size: 20px; 
                    color: #000000;
                    font-weight: bold;
                }
                QCheckBox::indicator {
                    width: 20px;
                    height: 20px;
                    border: 2px solid #d0d0d0;
                    border-radius: 4px;
                    background-color: #ffffff;
                }
                QCheckBox::indicator:checked {
                    background-color: #007acc;
                    border-color: #007acc;
                }
            """)
            self.year_checkbox_layout.addWidget(checkbox)
            self.year_checkboxes.append(checkbox)
        
        # Sağ tarafa butonlar ekle
        self.year_checkbox_layout.addStretch()
        
        # Butonları tanımla - Header Layout'tan buraya taşındı
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
        self.year_checkbox_layout.addWidget(self.mikro_button)
        
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
            QPushButton:disabled {
                background-color: #cccccc;
                color: #888888;
            }
        """)
        self.year_checkbox_layout.addWidget(self.refresh_button)
        
        self.export_button = QPushButton("Excel")
        self.export_button.setStyleSheet("""
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
        self.year_checkbox_layout.addWidget(self.export_button)
        
        # Year checkbox layout'u widget'a sar
        year_widget = QWidget()
        year_widget.setLayout(self.year_checkbox_layout)
        year_widget.setStyleSheet("""
            background-color: #ffffff;
            padding: 10px;
        """)
        self.year_checkbox_layout.setContentsMargins(10, 10, 10, 10)
        layout.addWidget(year_widget)

        # Ay ve KASA ADI için QComboBox'ları yan yana yerleştir
        self.combo_layout = QHBoxLayout()

        # Aylar için QComboBox ekle
        self.TURKCE_AYLAR = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran", 
                            "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
        
        self.ay_combo = QComboBox()
        self.ay_combo.addItem("Tüm Aylar")  # Tüm ayları göster seçeneği
        self.ay_combo.addItems(self.TURKCE_AYLAR)  # Türkçe ay isimlerini ekle
        self.ay_combo.setCurrentIndex(self.current_month)  # Varsayılan olarak bulunduğumuz ay seçili
        self.ay_combo.setStyleSheet("""            QComboBox { 
                font-size: 20px;
                min-width: 150px;
                background-color: #ffffff;
                color: #000000;
                border: 2px solid #d0d0d0;
                border-radius: 6px;
                padding: 8px;
                font-weight: bold;
            }
            QComboBox::drop-down {
                border: none;
                background-color: #f0f0f0;
            }
            QComboBox::down-arrow {
                image: none;
                border: none;
            }
            QComboBox QAbstractItemView { 
                font-size: 18px;
                background-color: #ffffff;
                color: #000000;
                selection-background-color: #b3d9ff;
                border: 1px solid #d0d0d0;
            }
        """)
        self.combo_layout.addWidget(self.ay_combo)

        # KASA ADI için QComboBox ekle
        self.kasa_adi_combo = QComboBox()
        self.kasa_adi_combo.setEditable(False)  # Düzenlenebilir özelliği kapalı
        self.kasa_adi_combo.setInsertPolicy(QComboBox.NoInsert)  # Yeni öğe eklenmesini engelle
        self.kasa_adi_combo.setStyleSheet("""            QComboBox { 
                font-size: 20px;
                min-width: 200px;
                background-color: #ffffff;
                color: #000000;
                border: 2px solid #d0d0d0;
                border-radius: 6px;
                padding: 8px;
                font-weight: bold;
            }
            QComboBox::drop-down {
                border: none;
                background-color: #f0f0f0;
            }
            QComboBox::down-arrow {
                image: none;
                border: none;
            }
            QComboBox QAbstractItemView { 
                font-size: 18px;
                background-color: #ffffff;
                color: #000000;
                selection-background-color: #b3d9ff;
                border: 1px solid #d0d0d0;
            }
        """)
        self.combo_layout.addWidget(self.kasa_adi_combo)

        # Combo layout'u widget'a sar
        combo_widget = QWidget()
        combo_widget.setLayout(self.combo_layout)
        combo_widget.setStyleSheet("""
            background-color: #ffffff;
            padding: 10px;
        """)
        self.combo_layout.setContentsMargins(10, 10, 10, 10)
        layout.addWidget(combo_widget)

        # Nakit / Dekont ve Alacak / Borç için checkbox'ları aynı satırda göster
        self.filter_checkbox_layout = QHBoxLayout()

        # Nakit / Dekont için checkbox'lar
        self.nakit_checkbox = QCheckBox("Nakit")
        self.nakit_checkbox.setChecked(True)  # Varsayılan olarak "Nakit" seçili
        self.nakit_checkbox.setStyleSheet("""                QCheckBox { 
                    font-size: 18px; 
                    color: #000000;
                    font-weight: bold;
                }
                QCheckBox::indicator {
                    width: 18px;
                    height: 18px;
                    border: 2px solid #d0d0d0;
                    border-radius: 4px;
                    background-color: #ffffff;
                }
                QCheckBox::indicator:checked {
                    background-color: #4CAF50;
                    border-color: #4CAF50;
                }
        """)
        
        self.dekont_checkbox = QCheckBox("Dekont")
        self.dekont_checkbox.setChecked(False)  # Başlangıçta işaretli değil
        self.dekont_checkbox.setStyleSheet("""                QCheckBox { 
                    font-size: 18px; 
                    color: #000000;
                    font-weight: bold;
                }
                QCheckBox::indicator {
                    width: 18px;
                    height: 18px;
                    border: 2px solid #d0d0d0;
                    border-radius: 4px;
                    background-color: #ffffff;
                }
                QCheckBox::indicator:checked {
                    background-color: #4CAF50;
                    border-color: #4CAF50;
                }
        """)
        
        self.filter_checkbox_layout.addWidget(self.nakit_checkbox)
        self.filter_checkbox_layout.addWidget(self.dekont_checkbox)
        self.filter_checkbox_layout.addSpacing(30)

        # Alacak / Borç için checkbox'lar
        self.alacak_checkbox = QCheckBox("Alacak")
        self.alacak_checkbox.setChecked(False)  # Başlangıçta işaretli değil
        self.alacak_checkbox.setStyleSheet("""                QCheckBox { 
                    font-size: 18px; 
                    color: #000000;
                    font-weight: bold;
                }
                QCheckBox::indicator {
                    width: 18px;
                    height: 18px;
                    border: 2px solid #d0d0d0;
                    border-radius: 4px;
                    background-color: #ffffff;
                }
                QCheckBox::indicator:checked {
                    background-color: #2196F3;
                    border-color: #2196F3;
                }
        """)
        
        self.borc_checkbox = QCheckBox("Borç")
        self.borc_checkbox.setChecked(True)  # Varsayılan olarak "Borç" seçili
        self.borc_checkbox.setStyleSheet("""                QCheckBox { 
                    font-size: 18px; 
                    color: #000000;
                    font-weight: bold;
                }
                QCheckBox::indicator {
                    width: 18px;
                    height: 18px;
                    border: 2px solid #d0d0d0;
                    border-radius: 4px;
                    background-color: #ffffff;
                }
                QCheckBox::indicator:checked {
                    background-color: #2196F3;
                    border-color: #2196F3;
                }
        """)
        
        self.filter_checkbox_layout.addWidget(self.alacak_checkbox)
        self.filter_checkbox_layout.addWidget(self.borc_checkbox)
        self.filter_checkbox_layout.addStretch()

        # Filter checkbox layout'u widget'a sar
        filter_widget = QWidget()
        filter_widget.setLayout(self.filter_checkbox_layout)
        filter_widget.setStyleSheet("""
            background-color: #ffffff;
            padding: 10px;
        """)
        self.filter_checkbox_layout.setContentsMargins(10, 10, 10, 10)
        layout.addWidget(filter_widget)

        # Progress Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.progress_bar.setFormat("%p%")
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
        
        # Table - Light theme (risk_module.py gibi)
        self.table = QTableWidget()
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
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
            QHeaderView::section {
                background-color: #f0f0f0;
                color: #000000;
                padding: 8px;
                border: 1px solid #d0d0d0;
                font-weight: bold;
                font-size: 15px;
            }
            QScrollBar:vertical {
                background: #2d2d2d;
                width: 15px;
                border-radius: 7px;
            }
            QScrollBar::handle:vertical {
                background: #007acc;
                border-radius: 7px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background: #0099ff;
            }
            QScrollBar:horizontal {
                background: #2d2d2d;
                height: 15px;
                border-radius: 7px;
            }
            QScrollBar::handle:horizontal {
                background: #007acc;
                border-radius: 7px;
                min-width: 20px;
            }
            QScrollBar::handle:horizontal:hover {
                background: #0099ff;
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

        status_layout.addWidget(self.status_label, 3)
        status_layout.addWidget(self.progress_bar, 1)
        status_layout.setContentsMargins(0, 0, 0, 0)

        # Status layout'u widget olarak sar
        status_widget = QWidget()
        status_widget.setLayout(status_layout)
        status_widget.setStyleSheet("background-color: #f5f5f5; border-top: 1px solid #d0d0d0;")
        
        # header_widget kaldırıldı
        layout.addWidget(self.table, 1)
        
        # Dip toplamı için etiket
        self.total_label = QLabel("Toplam: 0 ₺")
        self.total_label.setStyleSheet("""
            QLabel {
                color: #000000;
                padding: 8px;
                font-size: 14px;
                font-weight: bold;
                background-color: #ffffff;
            }
        """)
        
        layout.addWidget(self.total_label)
        layout.addWidget(status_widget)
        
        # Yıl, ay, KASA ADI değişimini dinle
        for checkbox in self.year_checkboxes:
            checkbox.stateChanged.connect(self.filter_table)
        self.ay_combo.currentTextChanged.connect(self.filter_table)
        self.kasa_adi_combo.currentTextChanged.connect(self.filter_table)
        
        # Nakit ve Borç checkbox'ları normal bağlantı (şifre gerekmez)
        self.nakit_checkbox.stateChanged.connect(self.filter_table)
        self.borc_checkbox.stateChanged.connect(self.filter_table)
    
    def setup_connections(self):
        self.mikro_button.clicked.connect(self.run_mikro)
        # Verileri Yenile butonu: cache'i bypass et, Google Sheets'ten çek
        self.refresh_button.clicked.connect(lambda: self.load_data(force_reload=True))
        self.export_button.clicked.connect(self.export_to_excel)
        self.table.customContextMenuRequested.connect(self.show_context_menu)

        # Dekont ve Alacak checkbox'ları şifre kontrolü ile bağlantı
        self.dekont_checkbox.clicked.connect(self._on_dekont_checkbox_clicked)
        self.alacak_checkbox.clicked.connect(self._on_alacak_checkbox_clicked)

        # Ctrl+C kısayolu - self üzerine bağlanır ki focus nerede olursa olsun çalışsın
        self.copy_shortcut = QShortcut(QKeySequence("Ctrl+C"), self)
        self.copy_shortcut.setContext(Qt.WindowShortcut)
        self.copy_shortcut.activated.connect(self.handle_ctrl_c)

    def load_data(self, force_reload=False):
        """
        Kasa sayfasından verileri yükle (cache-aware)

        Args:
            force_reload: True ise cache'i bypass et, Google Sheets'ten çek
        """
        try:
            # Global cache'i import et
            import sys
            if 'main' in sys.modules:
                from main import GlobalDataCache
                cache = GlobalDataCache()
            else:
                cache = None

            # Cache kontrolü (force_reload değilse)
            if not force_reload and cache and cache.has("Kasa"):
                self.df = cache.get("Kasa")
                self.veri_cercevesi = self.df.copy()

                # Tarih sütununu datetime formatına çevir
                if "Tarih" in self.df.columns:
                    self.df["Tarih"] = pd.to_datetime(self.df["Tarih"], format="%Y-%m-%d", errors='coerce')

                # TUTAR sütununu int'e çevir
                if "TUTAR" in self.df.columns:
                    self.df["TUTAR"] = pd.to_numeric(self.df["TUTAR"], errors='coerce').fillna(0).astype(int)

                # KASA ADI combobox'ını doldur
                if "KASA ADI" in self.df.columns:
                    if "KASA KODU" in self.df.columns:
                        self.df = self.df.sort_values(by="KASA KODU")
                    kasa_adlari = self.df["KASA ADI"].dropna().unique()
                    self.kasa_adi_combo.clear()
                    self.kasa_adi_combo.addItem("Tüm Kasa Adları")
                    self.kasa_adi_combo.addItems(kasa_adlari)

                self.filter_table()
                self.status_label.setText(f"✅ {len(self.df)} kayıt yüklendi (Cache'den - anında)")
                return

            # Cache yoksa veya force_reload ise: Google Sheets'ten çek
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(0)
            self.status_label.setText("📊 Kasa sayfasından veriler yükleniyor...")
            self.set_buttons_enabled(False)

            from PyQt5.QtWidgets import QApplication
            QApplication.processEvents()

            if not self.gsheets_url:
                self.progress_bar.setVisible(False)
                self.df = pd.DataFrame()
                self.veri_cercevesi = pd.DataFrame()
                self.update_table(self.df)
                self.status_label.setText("❌ PRGsheet/Ayar sayfasında SPREADSHEET_ID bulunamadı")
                return

            self.progress_bar.setValue(10)
            self.status_label.setText(f"🔗 Google Sheets'e bağlanıyor...")
            QApplication.processEvents()

            # URL'den Excel dosyasını oku
            response = requests.get(self.gsheets_url, timeout=30)

            self.progress_bar.setValue(30)
            self.status_label.setText("✅ Google Sheets'e bağlantı başarılı")
            QApplication.processEvents()

            if response.status_code == 401:
                self.progress_bar.setVisible(False)
                self.df = pd.DataFrame()
                self.veri_cercevesi = pd.DataFrame()
                self.update_table(self.df)
                self.status_label.setText("❌ Google Sheets erişim hatası: Dosya özel veya izin gerekli. Dosyayı 'Anyone with the link can view' yapmayı deneyin.")
                return
            elif response.status_code != 200:
                self.progress_bar.setVisible(False)
                self.df = pd.DataFrame()
                self.veri_cercevesi = pd.DataFrame()
                self.update_table(self.df)
                self.status_label.setText(f"❌ HTTP Hatası: {response.status_code} - {response.reason}")
                return
            
            response.raise_for_status()
            
            from io import BytesIO

            self.progress_bar.setValue(50)
            self.status_label.setText("📋 Kasa sayfası yükleniyor...")
            QApplication.processEvents()

            # Kasa sayfasını oku
            self.df = pd.read_excel(BytesIO(response.content), sheet_name="Kasa")
            self.veri_cercevesi = self.df.copy()

            self.progress_bar.setValue(70)
            self.status_label.setText("🔄 Veriler işleniyor...")
            QApplication.processEvents()

            # Tarih sütununu datetime formatına çevir
            if "Tarih" in self.df.columns:
                self.df["Tarih"] = pd.to_datetime(self.df["Tarih"], format="%Y-%m-%d", errors='coerce')

            # TUTAR sütununu int'e çevir (NaN değerlerini 0 ile doldur)
            if "TUTAR" in self.df.columns:
                self.df["TUTAR"] = pd.to_numeric(self.df["TUTAR"], errors='coerce').fillna(0).astype(int)

            self.progress_bar.setValue(85)
            self.status_label.setText("🔄 Kasa listesi hazırlanıyor...")
            QApplication.processEvents()

            # KASA ADI combobox'ını doldur
            if "KASA ADI" in self.df.columns:
                if "KASA KODU" in self.df.columns:
                    # KASA KODU'na göre sırala
                    self.df = self.df.sort_values(by="KASA KODU")
                kasa_adlari = self.df["KASA ADI"].dropna().unique()  # NaN değerleri atla ve benzersiz değerleri al
                self.kasa_adi_combo.clear()  # Combobox'ı temizle
                self.kasa_adi_combo.addItem("Tüm Kasa Adları")  # Tümünü göster seçeneği
                self.kasa_adi_combo.addItems(kasa_adlari)  # Benzersiz KASA ADI değerlerini ekle

            self.progress_bar.setValue(95)
            self.status_label.setText("🔄 Tablo güncelleniyor...")
            QApplication.processEvents()

            # İlk açılışta filtreleme yap
            self.filter_table()

            # Cache'e kaydet
            if cache:
                cache.set("Kasa", self.df)

            # Tüm işlemler tamamlandı
            self.progress_bar.setValue(100)
            QApplication.processEvents()

            # Progress bar'ı 1 saniye sonra gizle
            QTimer.singleShot(1000, lambda: self.progress_bar.setVisible(False))

            self.status_label.setText(f"✅ {len(self.df)} kayıt başarıyla yüklendi (Kasa sayfası)")

        except requests.exceptions.Timeout:
            self.progress_bar.setVisible(False)
            self.df = pd.DataFrame()
            self.veri_cercevesi = pd.DataFrame()
            self.update_table(self.df)
            self.status_label.setText("❌ Bağlantı zaman aşımı - Google Sheets'e erişilemiyor")
        except requests.exceptions.RequestException as e:
            self.progress_bar.setVisible(False)
            self.df = pd.DataFrame()
            self.veri_cercevesi = pd.DataFrame()
            self.update_table(self.df)
            self.status_label.setText(f"❌ Bağlantı hatası: {str(e)}")
        except Exception as e:
            self.progress_bar.setVisible(False)
            self.df = pd.DataFrame()
            self.veri_cercevesi = pd.DataFrame()
            self.update_table(self.df)
            self.status_label.setText(f"❌ Veri yükleme hatası: {str(e)}")
        finally:
            self.set_buttons_enabled(True)
    
    def filter_table(self):
        """Filtreleme fonksiyonu"""
        try:
            if self.df.empty:
                return
            
            filtered_df = self.df.copy()
            
            # Yıl filtresi
            selected_years = []
            for checkbox in self.year_checkboxes:
                if checkbox.isChecked():
                    selected_years.append(int(checkbox.text()))

            if selected_years and "Tarih" in self.df.columns:
                filtered_df = filtered_df[filtered_df["Tarih"].dt.year.isin(selected_years)]

            # Ay filtresi
            selected_ay = self.ay_combo.currentText()
            if selected_ay != "Tüm Aylar" and "Tarih" in filtered_df.columns:
                ay_numarasi = self.TURKCE_AYLAR.index(selected_ay) + 1  # Ay numarasını bul (1-12)
                filtered_df = filtered_df[filtered_df["Tarih"].dt.month == ay_numarasi]

            # KASA ADI filtresi
            selected_kasa_adi = self.kasa_adi_combo.currentText()
            if selected_kasa_adi != "Tüm Kasa Adları" and "KASA ADI" in filtered_df.columns:
                filtered_df = filtered_df[filtered_df["KASA ADI"] == selected_kasa_adi]

            # Nakit / Dekont filtresi
            if "Nakit / Dekont" in filtered_df.columns:
                if self.nakit_checkbox.isChecked() and not self.dekont_checkbox.isChecked():
                    filtered_df = filtered_df[filtered_df["Nakit / Dekont"] == "Nakit"]
                elif not self.nakit_checkbox.isChecked() and self.dekont_checkbox.isChecked():
                    filtered_df = filtered_df[filtered_df["Nakit / Dekont"] == "Dekont"]
                elif not self.nakit_checkbox.isChecked() and not self.dekont_checkbox.isChecked():
                    filtered_df = filtered_df[filtered_df["Nakit / Dekont"].isna()]  # Hiçbiri seçilmediyse boş veri göster

            # Alacak / Borç filtresi
            if "Alacak / Borç" in filtered_df.columns:
                if self.alacak_checkbox.isChecked() and not self.borc_checkbox.isChecked():
                    filtered_df = filtered_df[filtered_df["Alacak / Borç"] == "Alacak"]
                elif not self.alacak_checkbox.isChecked() and self.borc_checkbox.isChecked():
                    filtered_df = filtered_df[filtered_df["Alacak / Borç"] == "Borç"]
                elif not self.alacak_checkbox.isChecked() and not self.borc_checkbox.isChecked():
                    filtered_df = filtered_df[filtered_df["Alacak / Borç"].isna()]  # Hiçbiri seçilmediyse boş veri göster

            # Tabloyu güncelle
            self.update_table(filtered_df)
            
        except Exception as e:
            self.status_label.setText(f"❌ Filtreleme hatası: {str(e)}")

    def update_table(self, df):
        """Tabloyu verilerle güncelle"""
        if df.empty:
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            self.total_label.setText("Toplam: 0 ₺")
            return
        
        # Verileri tarihe göre sırala (yeniden eskiye)
        if "Tarih" in df.columns:
            df = df.sort_values(by="Tarih", ascending=False)
            # Tarih sütununu "YYYY-MM-DD" formatına çevir (saat bilgisini kaldır)
            df_display = df.copy()
            df_display["Tarih"] = df_display["Tarih"].dt.strftime("%Y-%m-%d")
        else:
            df_display = df.copy()
        
        self.table.setRowCount(len(df_display))
        self.table.setColumnCount(len(df_display.columns))
        self.table.setHorizontalHeaderLabels(df_display.columns.tolist())
        
        # Set table properties for better appearance
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setFocusPolicy(Qt.NoFocus)  # Remove focus policy to eliminate dotted borders
        
        # Fill table with data and apply enhanced formatting
        for i in range(len(df_display)):
            for j in range(len(df_display.columns)):
                value = df_display.iat[i, j]
                
                if pd.isna(value) or str(value).lower() == 'nan':
                    display_value = ""
                elif j < len(df_display.columns) and 'telefon' in df_display.columns[j].lower():
                    try:
                        display_value = str(int(float(value)))
                    except (ValueError, TypeError):
                        display_value = str(value)
                else:
                    display_value = str(value)
                
                item = QTableWidgetItem(display_value)
                item.setFlags(item.flags() ^ Qt.ItemIsEditable)  # Make non-editable
                
                # Set font properties for better readability
                font = QFont('Segoe UI', 12)
                font.setBold(True)
                item.setFont(font)
                
                # Color coding for specific columns
                column_name = df_display.columns[j]
                if 'tutar' in column_name.lower() or 'miktar' in column_name.lower():
                    try:
                        numeric_value = float(str(display_value).replace(',', ''))
                        if numeric_value > 0:
                            item.setForeground(QColor("#4CAF50"))  # Green for positive
                        elif numeric_value < 0:
                            item.setForeground(QColor("#f44336"))  # Red for negative
                        else:
                            item.setForeground(QColor("#000000"))  # White for zero
                    except:
                        item.setForeground(QColor("#000000"))
                else:
                    item.setForeground(QColor("#000000"))
                
                self.table.setItem(i, j, item)
        
        # Enhanced header styling
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setStretchLastSection(False)
        
        # Set minimum column widths
        for i in range(self.table.columnCount()):
            self.table.setColumnWidth(i, max(150, self.table.columnWidth(i)))

        # Resize columns to content but with minimum width
        self.table.resizeColumnsToContents()
        
        # Set row height for better readability
        for i in range(self.table.rowCount()):
            self.table.setRowHeight(i, 35)
        
        # TUTAR sütunu toplamını hesapla ve göster
        if "TUTAR" in df.columns:
            total_tutar = df["TUTAR"].sum()
            formatted_total = "{:,.0f} ₺".format(total_tutar).replace(",", ".")  # Binlik ayraçları nokta ile göster
            self.total_label.setText(f"Toplam: {formatted_total}")
        else:
            self.total_label.setText("Toplam: 0 ₺")
    
    def run_mikro(self):
        """Kasa.exe dosyasını çalıştır"""
        try:
            exe_path = r"D:/GoogleDrive/PRG/EXE/Kasa.exe"
            if not os.path.exists(exe_path):
                self.status_label.setText(f"❌ Kasa.exe bulunamadı: {exe_path}")
                return
            
            self.status_label.setText("🔄 Kasa.exe çalıştırılıyor...")
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
        self.status_label.setText("✅ Kasa.exe tamamlandı, Google Sheets güncelleme bekleniyor...")
        
        # Google Sheets'e kaydedilmesi için ek bekleme (5 saniye)
        QTimer.singleShot(5000, self.delayed_data_refresh)
    
    def delayed_data_refresh(self):
        """Gecikmeli veri yenileme"""
        self.status_label.setText("🔄 Google Sheets'ten güncel veriler alınıyor...")
        from PyQt5.QtWidgets import QApplication
        QApplication.processEvents()
        self.load_data()
    
    def export_to_excel(self):
        """Filtrelenmiş verileri Excel'e aktar"""
        if self.df.empty:
            self.status_label.setText("⚠️ Dışa aktarılacak veri yok")
            return
        
        try:
            # Seçili KASA ADI'nı al
            selected_kasa_adi = self.kasa_adi_combo.currentText()

            # Dosya adını oluştur
            if selected_kasa_adi == "Tüm Kasa Adları":
                file_name = "Tum_Kasa_Adlari"
            else:
                file_name = selected_kasa_adi.replace(" ", "_")  # Boşlukları alt çizgi ile değiştir

            # Dosya yolunu belirle
            output_path = f"D:/GoogleDrive/~ {file_name}.xlsx"

            # Filtrelenmiş veriyi al
            filtered_df = self.get_filtered_data()

            # Excel'e kaydet
            filtered_df.to_excel(output_path, index=False, engine='openpyxl')
            self.status_label.setText(f"✅ Veriler dışa aktarıldı: {output_path}")
        except Exception as e:
            self.status_label.setText(f"❌ Dışa aktarma hatası: {str(e)}")
    
    def get_filtered_data(self):
        """Filtrelenmiş veriyi döndür"""
        if self.df.empty:
            return pd.DataFrame()
            
        # Yıl filtresi
        selected_years = []
        for checkbox in self.year_checkboxes:
            if checkbox.isChecked():
                selected_years.append(int(checkbox.text()))

        if selected_years and "Tarih" in self.df.columns:
            filtered_df = self.df[self.df["Tarih"].dt.year.isin(selected_years)]
        else:
            filtered_df = self.df

        # Ay filtresi
        selected_ay = self.ay_combo.currentText()
        if selected_ay != "Tüm Aylar" and "Tarih" in filtered_df.columns:
            ay_numarasi = self.TURKCE_AYLAR.index(selected_ay) + 1
            filtered_df = filtered_df[filtered_df["Tarih"].dt.month == ay_numarasi]

        # KASA ADI filtresi
        selected_kasa_adi = self.kasa_adi_combo.currentText()
        if selected_kasa_adi != "Tüm Kasa Adları" and "KASA ADI" in filtered_df.columns:
            filtered_df = filtered_df[filtered_df["KASA ADI"] == selected_kasa_adi]

        # Nakit / Dekont filtresi
        if "Nakit / Dekont" in filtered_df.columns:
            if self.nakit_checkbox.isChecked() and not self.dekont_checkbox.isChecked():
                filtered_df = filtered_df[filtered_df["Nakit / Dekont"] == "Nakit"]
            elif not self.nakit_checkbox.isChecked() and self.dekont_checkbox.isChecked():
                filtered_df = filtered_df[filtered_df["Nakit / Dekont"] == "Dekont"]
            elif not self.nakit_checkbox.isChecked() and not self.dekont_checkbox.isChecked():
                filtered_df = filtered_df[filtered_df["Nakit / Dekont"].isna()]

        # Alacak / Borç filtresi
        if "Alacak / Borç" in filtered_df.columns:
            if self.alacak_checkbox.isChecked() and not self.borc_checkbox.isChecked():
                filtered_df = filtered_df[filtered_df["Alacak / Borç"] == "Alacak"]
            elif not self.alacak_checkbox.isChecked() and self.borc_checkbox.isChecked():
                filtered_df = filtered_df[filtered_df["Alacak / Borç"] == "Borç"]
            elif not self.alacak_checkbox.isChecked() and not self.borc_checkbox.isChecked():
                filtered_df = filtered_df[filtered_df["Alacak / Borç"].isna()]

        return filtered_df
    
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
    
    def set_buttons_enabled(self, enabled: bool):
        """Butonları aktif/pasif yap"""
        self.mikro_button.setEnabled(enabled)
        self.refresh_button.setEnabled(enabled)
        self.export_button.setEnabled(enabled)