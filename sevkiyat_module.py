"""
Sevkiyat Modülü - Sevkiyat işlemlerini yönetir
"""

import os
import sys
import time
import numpy as np
import pandas as pd
import requests
import subprocess
from typing import List
from pathlib import Path
from io import BytesIO
from datetime import datetime, timedelta, date
from dateutil.relativedelta import relativedelta
from fuzzywuzzy import process
from dataclasses import dataclass
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import urllib.parse
import webbrowser
import pyperclip
from dotenv import load_dotenv

# Üst dizini Python path'e ekle (central_config için)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Central config import
from central_config import CentralConfigManager
from PyQt5.QtCore import Qt, QTimer, QDateTime, QThread, pyqtSignal
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QLineEdit, QTextEdit, 
                             QTableWidget, QTableWidgetItem, QListWidget, QScrollArea, QHeaderView,
                             QAbstractItemView, QMenu, QAction, QMessageBox, QProgressBar, QApplication)
from PyQt5.QtGui import QFont, QColor


from PyQt5.QtGui import QFont, QColor


class MikroUpdateThread(QThread):
    """Mikro güncelleme işlemlerini sırayla yürüten thread"""
    status_update = pyqtSignal(str)
    progress_update = pyqtSignal(int)
    finished_signal = pyqtSignal()
    error_signal = pyqtSignal(str)

    def run(self):
        try:
            exe_list = [
                ("BagKodu.exe", r"D:/GoogleDrive/PRG/EXE/BagKodu.exe"),
                ("BekleyenAPI.exe", r"D:/GoogleDrive/PRG/EXE/BekleyenAPI.exe"),
                ("Risk.exe", r"D:/GoogleDrive/PRG/EXE/Risk.exe"),
                ("Stok.exe", r"D:/GoogleDrive/PRG/EXE/Stok.exe"),
                ("Sevkiyat.exe", r"D:/GoogleDrive/PRG/EXE/Sevkiyat.exe")
            ]
            
            total_steps = len(exe_list)
            
            for i, (name, path) in enumerate(exe_list):
                # Özel karakter temizliği (örn: görünmez unicode karakterleri)
                clean_path = path.replace('\u202a', '').replace('\u202c', '').strip()
                
                if not os.path.exists(clean_path):
                    self.error_signal.emit(f"Dosya bulunamadı: {name}")
                    continue
                    
                self.status_update.emit(f"🔄 {name} çalıştırılıyor... ({i+1}/{total_steps})")
                self.progress_update.emit(int((i / total_steps) * 100))
                
                # EXE'yi çalıştır ve bitmesini bekle
                try:
                    # creationflags=0x08000000 (CREATE_NO_WINDOW) konsol penceresini gizlemek için opsiyonel kullanılabilir
                    # ancak kullanıcı görsün istiyorsa varsayılan haliyle bırakıyoruz.
                    subprocess.run(clean_path, check=True, shell=False)
                except subprocess.CalledProcessError as e:
                    self.error_signal.emit(f"{name} hatayla sonlandı: {e}")
                except Exception as e:
                    self.error_signal.emit(f"{name} çalıştırılamadı: {e}")
            
            self.progress_update.emit(100)
            self.finished_signal.emit()
            
        except Exception as e:
            self.error_signal.emit(f"Beklenmedik hata: {str(e)}")


class SevkiyatModule(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sevkiyat Yönetimi")
        self.setGeometry(200, 200, 1200, 800)
        
        # Google Sheets integration
        self.gsheets_url = self._load_gsheets_url()
        
        # Data frames
        self.cari_df = pd.DataFrame()
        self.sevkiyat_df = pd.DataFrame()
        self.bekleyenler_df = pd.DataFrame()
        self.arac_df = pd.DataFrame()
        self.mail_info_df = pd.DataFrame()
        self.mail_sevk_info_df = pd.DataFrame()
        self.risk_df = pd.DataFrame()
        
        # Filtered data
        self.sevkiyat_filtered_data = pd.DataFrame()
        self.sevkiyat_filtered_again = pd.DataFrame()  # Eksik değişken eklendi
        self.bekleyenler_filtered_data = pd.DataFrame()
        self.arac_filtered_data = pd.DataFrame()
        self.mail_data = pd.DataFrame()
        self.mail_sevk_data = pd.DataFrame()
        
        # Customer data
        self.customer_names = []
        self.cari_column_name = None  # Dinamik sütun adı
        self.cari_adi = None
        self.cari_telefon = None
        self.depo = None
        
        # Mikro güncelleme için
        self.mikro_calisiyor = False

        # Lazy loading için flag
        self._data_loaded = False

        self.init_ui()
        self.setup_connections()

    def showEvent(self, event):
        """Widget ilk gösterildiğinde veri yükle (lazy loading)"""
        super().showEvent(event)
        if not self._data_loaded:
            self._data_loaded = True
            # UI render olduktan sonra veri yükle
            QTimer.singleShot(100, self.load_all_data)

    def _load_gsheets_url(self):
        """Google Sheets SPREADSHEET_ID'sini yükle - Service Account"""
        try:
            config_manager = CentralConfigManager()
            spreadsheet_id = config_manager.MASTER_SPREADSHEET_ID
            return f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"
        except Exception as e:
            return None
    
    def init_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        
        # Progress Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #2d2d2d;
                border-radius: 3px;
                background-color: #1a1a1a;
                color: #ffffff;
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
        
        # Ana layout - Sol ve sağ kısım
        main_layout = QHBoxLayout()
        
        # Sol kısım - Arama çubuğu ve arama sonuçları
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 5, 0)
        
        # Arama çubuğu
        self.search_bar = QLineEdit(self)
        self.search_bar.setPlaceholderText("Müşteri isim ve soyismini girin (örn. GÜNEŞ*)")
        self.search_bar.setFixedHeight(50)  # 50px yükseklik
        self.search_bar.setFont(QFont("Arial Bold", 14))
        self.search_bar.setStyleSheet("""
            QLineEdit {
                background-color: #1a1a1a;
                color: #ffffff;
                border: 2px solid #404040;
                border-radius: 8px;
                padding: 8px;
                font-size: 17px;
                font-weight: bold;
            }
            QLineEdit:focus {
                border-color: #007acc;
            }
        """)
        
        # Arama sonuçları listesi (arama kutusu genişliğinde)
        self.result_list = QListWidget(self)
        self.result_list.setFont(QFont("Arial Bold", 14))
        self.result_list.setFixedHeight(7 * 37) 
        self.result_list.setStyleSheet("""
            QListWidget {
                background-color: #1a1a1a;
                color: #ffffff;
                border: 2px solid #404040;
                border-radius: 8px;
                selection-background-color: #007acc;
                selection-color: #ffffff;
                font-size: 14px;
                font-weight: bold;
            }
            QListWidget::item {
                padding: 8px;
                border-bottom: 1px solid #404040;
                color: #ffffff;
            }
            QListWidget::item:selected {
                background-color: #007acc;
                color: #ffffff;
            }
            QListWidget::item:hover {
                background-color: #333333;
                color: #ffffff;
            }
            QListWidget::item:focus {
                outline: none;
                border: none;
            }
        """)

        # Bağlam menüsü
        self.result_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.result_list.setFocusPolicy(Qt.NoFocus)
        
        # Sol layout'a widget'ları ekle
        left_layout.addWidget(self.search_bar)
        left_layout.addWidget(self.result_list)
        left_layout.addStretch()  # Boş alan ekle
        
        # Sağ kısım - Müşteri bilgi butonu ve diğer butonlar
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(5, 0, 0, 0)
        
        # Müşteri bilgi butonu  
        self.sozlesmedeki_urunler_button = QPushButton(" ", self)
        self.sozlesmedeki_urunler_button.setFixedHeight(50)  # 50px yükseklik
        self.sozlesmedeki_urunler_button.setFont(QFont("Arial Bold", 16))
        self.sozlesmedeki_urunler_button.setEnabled(False)
        self.sozlesmedeki_urunler_button.setStyleSheet("""
            QPushButton {
                background-color: #2d2d2d;
                color: #ffffff;
                border: 2px solid #404040;
                border-radius: 8px;
                padding: 8px;
                font-weight: bold;
                text-align: center;
            }
        """)
        
        # Butonlar için layout
        buttons_layout = QVBoxLayout()  # Dikey düzen
        
        # İlk satır butonlar - Verileri Yenile + WhatsApp
        first_row_layout = QHBoxLayout()
        
        # Refresh Button
        self.refresh_button = QPushButton("Verileri Yenile")
        self.refresh_button.setFixedHeight(40)
        self.refresh_button.setStyleSheet("""
            QPushButton {
                background-color: #000000;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 17px;
            }
            QPushButton:hover {
                background-color: #333333;
            }
        """)
        
        # Mikro Güncelle Button
        self.mikro_button = QPushButton("Mikro Güncelle")
        self.mikro_button.setFixedHeight(40)
        self.mikro_button.setStyleSheet("""
            QPushButton {
                background-color: #000000;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 17px;
            }
            QPushButton:hover {
                background-color: #333333;
            }
            QPushButton:disabled {
                background-color: #666666;
            }
        """)
        
        first_row_layout.addWidget(self.refresh_button)
        first_row_layout.addWidget(self.mikro_button)
        
        # WhatsApp Buttons - ikinci satıra taşınacak
        whatsapp_row_layout = QHBoxLayout()
        
        self.whatsapp_randevu_button = QPushButton("📆   WhatsApp - Randevu Al", self)
        self.whatsapp_randevu_button.setFixedHeight(40)
        self.whatsapp_randevu_button.setStyleSheet("""
            QPushButton {
                background-color: #000000;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 17px;
            }
            QPushButton:hover {
                background-color: #25D366;
            }
        """)
        
        self.whatsapp_bilgi_button = QPushButton("📩   WhatsApp - Bilgi", self)
        self.whatsapp_bilgi_button.setFixedHeight(40)
        self.whatsapp_bilgi_button.setStyleSheet("""
            QPushButton {
                background-color: #000000;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 17px;
            }
            QPushButton:hover {
                background-color: #25D366;
            }
        """)
        
        whatsapp_row_layout.addWidget(self.whatsapp_randevu_button)
        whatsapp_row_layout.addWidget(self.whatsapp_bilgi_button)
        
        # İkinci satır butonlar - Sevkiyatı Dışa Aktar + Bekleyenleri Dışa Aktar
        second_row_layout = QHBoxLayout()
        
        self.export_button = QPushButton("Dışa Aktar - Sevkiyat")
        self.export_button.setFixedHeight(40)
        self.export_button.setStyleSheet("""
            QPushButton {
                background-color: #000000;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 17px;
            }
            QPushButton:hover {
                background-color: #ff9800;
            }
        """)
        
        self.export_bekleyenler_button = QPushButton("Dışa Aktar - Bekleyenler", self)
        self.export_bekleyenler_button.setFixedHeight(40)
        self.export_bekleyenler_button.setStyleSheet("""
            QPushButton {
                background-color: #000000;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 17px;
            }
            QPushButton:hover {
                background-color: #ff9800;
            }
        """)
        
        second_row_layout.addWidget(self.export_button)
        second_row_layout.addWidget(self.export_bekleyenler_button)
        
        # Üçüncü satır butonlar - Planlanan Aracı Dışa Aktar + Malzeme Bazlı Dışa Aktar
        third_row_layout = QHBoxLayout()
        
        self.export_arac_button = QPushButton("Dışa Aktar - Plan Araç", self)
        self.export_arac_button.setFixedHeight(40)
        self.export_arac_button.setStyleSheet("""
            QPushButton {
                background-color: #000000;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 17px;
            }
            QPushButton:hover {
                background-color: #ff9800;
            }
        """)
        
        self.export_malzeme_button = QPushButton("Dışa Aktar - Malzeme Borç", self)
        self.export_malzeme_button.setFixedHeight(40)
        self.export_malzeme_button.setStyleSheet("""
            QPushButton {
                background-color: #000000;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 17px;
            }
            QPushButton:hover {
                background-color: #ff9800;
            }
        """)
        
        third_row_layout.addWidget(self.export_arac_button)
        third_row_layout.addWidget(self.export_malzeme_button)
        
        # Dördüncü satır butonlar - Açık Sipariş Mail Gönder + Sevke Hazır Mail Gönder
        fourth_row_layout = QHBoxLayout()
        
        self.mail_gonder_button = QPushButton("Mail Gönder - Açık Sipariş", self)
        self.mail_gonder_button.setFixedHeight(40)
        self.mail_gonder_button.setStyleSheet("""
            QPushButton {
                background-color: #000000;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 17px;
            }
            QPushButton:hover {
                background-color: #1a73e8;
            }
        """)
        
        self.sevk_button = QPushButton("Mail Gönder - Sevke Hazır", self)
        self.sevk_button.setFixedHeight(40)
        self.sevk_button.setStyleSheet("""
            QPushButton {
                background-color: #000000;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 17px;
            }
            QPushButton:hover {
                background-color: #34a853;
            }
        """)
        
        fourth_row_layout.addWidget(self.mail_gonder_button)
        fourth_row_layout.addWidget(self.sevk_button)
        
        # Beşinci satır - Stok Analizi butonu
        fifth_row_layout = QHBoxLayout()
        
        self.stok_analizi_button = QPushButton("Stok Analizi", self)
        self.stok_analizi_button.setFixedHeight(40)
        self.stok_analizi_button.setStyleSheet("""
            QPushButton {
                background-color: #000000;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 17px;
            }
            QPushButton:hover {
                background-color: #ff6b35;
            }
        """)
        
        fifth_row_layout.addWidget(self.stok_analizi_button)
        
        # Buton layout'larını ana buton layout'una ekle
        buttons_layout.addLayout(first_row_layout)
        buttons_layout.addLayout(whatsapp_row_layout)
        buttons_layout.addLayout(second_row_layout)
        buttons_layout.addLayout(third_row_layout)
        buttons_layout.addLayout(fourth_row_layout)
        buttons_layout.addLayout(fifth_row_layout)
        
        # Sağ layout'a widget'ları ekle
        right_layout.addWidget(self.sozlesmedeki_urunler_button)
        right_layout.addLayout(buttons_layout)
        right_layout.addStretch()  # Boş alan ekle
        
        # Ana layout'a sol ve sağ widget'ları ekle
        main_layout.addWidget(left_widget, 1)   # Sol kısım
        main_layout.addWidget(right_widget, 2)  # Sağ kısım (daha geniş)
        
        # QScrollArea bileşenini ekleyelim
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area_widget = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_area_widget)
        
        self.filtered_label = QLabel(self)
        self.filtered_label.setWordWrap(True)
        self.filtered_label.setAlignment(Qt.AlignTop)
        self.filtered_label.setStyleSheet("""
            QLabel {
                background-color: #ffffff;
                color: #000000;
                border: none;
                border-radius: 8px;
                padding: 10px;
            }
        """)
        self.scroll_layout.addWidget(self.filtered_label)
        
        self.scroll_area.setWidget(self.scroll_area_widget)
        self.scroll_area.setStyleSheet("""
            QScrollArea {
                border: none;
                background-color: #1a1a1a;
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
        """)
        
        # Status Layout
        status_layout = QHBoxLayout()
        
        self.status_label = QLabel("Hazır")
        self.status_label.setStyleSheet("""
            QLabel {
                color: #cccccc;
                padding: 4px 8px;
                background-color: #2d2d2d;
                border-top: 1px solid #404040;
                font-size: 14px;
                max-height: 20px;
            }
        """)
        
        status_layout.addWidget(self.status_label, 3)
        status_layout.addWidget(self.progress_bar, 1)  
        status_layout.setContentsMargins(0, 0, 0, 0)
        
        status_widget = QWidget()
        status_widget.setLayout(status_layout)
        status_widget.setStyleSheet("background-color: #2d2d2d; border-top: 1px solid #404040;")
        
        # Ana layout'a widget'ları ekle
        layout.addLayout(main_layout)          # Sol (arama+sonuçlar) + Sağ (müşteri bilgi+butonlar)
        layout.addWidget(self.scroll_area, 1)  # Tablo alanı (genişleyebilir)
        layout.addWidget(status_widget)        # Durum çubuğu
        
        # Widget'ın genel stilini ayarla
        self.setStyleSheet("""
            QWidget {
                background-color: #1a1a1a;
                color: #ffffff;
            }
        """)

    def setup_connections(self):
        """Bağlantıları kur"""
        self.refresh_button.clicked.connect(self.load_all_data)
        self.mikro_button.clicked.connect(self.run_mikro)
        self.export_button.clicked.connect(self.export_sevkiyat_to_excel)
        self.export_bekleyenler_button.clicked.connect(self.export_bekleyenler_to_excel)
        self.export_arac_button.clicked.connect(self.export_arac_to_excel)
        self.export_malzeme_button.clicked.connect(self.export_malzeme_to_excel)
        self.search_bar.textChanged.connect(self.update_search)
        self.result_list.itemClicked.connect(self.filter_by_selected_customer)
        self.result_list.customContextMenuRequested.connect(self.show_context_menu)
        self.whatsapp_randevu_button.clicked.connect(self.whatsapp_randevu_gonder)
        self.whatsapp_bilgi_button.clicked.connect(self.whatsapp_bilgi_gonder)
        self.mail_gonder_button.clicked.connect(self.mail_gonder_button_clicked)
        self.sevk_button.clicked.connect(self.sevk_button_clicked)
        self.stok_analizi_button.clicked.connect(self.stok_analizi_goster)

    def format_kalem_no(self, df):
        """Kalem No sütununu formatla: 11-13. karakterdeki '000' yerine '-' koy"""
        if 'Kalem No' in df.columns:
            def transform_kalem_no(value):
                # Değeri string'e çevir
                kalem_str = str(value)
                # Bilimsel notasyonu temizle
                if 'E+' in kalem_str or 'e+' in kalem_str:
                    kalem_str = str(int(float(kalem_str)))
                # Eğer uzunluk yeterli ise 11-13. karakterdeki '000' yerine '-' koy
                if len(kalem_str) >= 13:
                    return kalem_str[:10] + '-' + kalem_str[13:]
                return kalem_str

            df['Kalem No'] = df['Kalem No'].apply(transform_kalem_no)
        return df

    def load_depo_settings(self):
        """
        Ayar sayfasından depo bilgilerini yükle

        Returns:
            {'17': 'BİGA', '16': 'İNEGÖL', '48': 'KAYSERİ'}
        """
        try:
            config_manager = CentralConfigManager()
            # Cache kullan (HIZLI) - "Veri Yenile" butonuna basıldığında cache temizlenir
            settings = config_manager.get_settings(use_cache=True)

            # Depo_ ile başlayan tüm ayarları bul
            depolar = {}
            for key, value in settings.items():
                if key.startswith('Depo_'):
                    # Depo_17 -> 17
                    depo_plaka = key.replace('Depo_', '')
                    depolar[depo_plaka] = value

            # Eğer hiç depo bulunamadıysa uyarı ver
            if not depolar:
                QMessageBox.warning(self, "Depo Ayarları Bulunamadı",
                                   "PRGsheet → Ayar sayfasında 'Depo_' ile başlayan ayarlar bulunamadı!\n\n"
                                   "Örnek format:\n"
                                   "App Name: Global\n"
                                   "Key: Depo_17\n"
                                   "Value: BİGA\n\n"
                                   "Ayarları ekledikten sonra 'Veri Yenile' butonuna basın.\n\n"
                                   "Şimdilik varsayılan depo ayarları kullanılacak.")
                return {
                    "17": "BİGA",
                    "16": "İNEGÖL",
                    "48": "KAYSERİ"
                }

            return depolar
        except Exception as e:
            # Hata durumunda kullanıcıya bilgi ver ve varsayılan değerleri kullan
            QMessageBox.warning(self, "Depo Ayarları Yükleme Hatası",
                               f"Depo ayarları yüklenirken hata oluştu:\n{str(e)}\n\n"
                               "'Veri Yenile' butonuna basarak tekrar deneyin.\n\n"
                               "Şimdilik varsayılan depo ayarları kullanılacak.")
            return {
                "17": "BİGA",
                "16": "İNEGÖL",
                "48": "KAYSERİ"
            }

    def load_all_data(self):
        """Tüm Google Sheets sayfalarından verileri yükle"""
        try:
            # Ayar cache'ini temizle - böylece güncel ayarlar yüklenecek
            try:
                config_manager = CentralConfigManager()
                config_manager.refresh_config()
            except Exception as e:
                pass  # Cache temizleme hatası önemli değil, devam et

            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(0)
            self.status_label.setText("📊 Google Sheets'ten veriler yükleniyor...")
            self.set_buttons_enabled(False)

            QApplication.processEvents()

            if not self.gsheets_url:
                self.progress_bar.setVisible(False)
                self.status_label.setText("❌ PRGsheet/Ayar sayfasında SPREADSHEET_ID bulunamadı")
                return

            # URL'den Excel dosyasını oku
            self.progress_bar.setValue(10)
            self.status_label.setText("🔗 Google Sheets'e bağlanıyor...")
            QApplication.processEvents()

            response = requests.get(self.gsheets_url, timeout=30)

            self.progress_bar.setValue(20)
            self.status_label.setText("✅ Google Sheets'e bağlantı başarılı")
            QApplication.processEvents()

            if response.status_code == 401:
                self.progress_bar.setVisible(False)
                self.status_label.setText("❌ Google Sheets erişim hatası: Dosya özel veya izin gerekli")
                return
            elif response.status_code != 200:
                self.progress_bar.setVisible(False)
                self.status_label.setText(f"❌ HTTP Hatası: {response.status_code} - {response.reason}")
                return
            
            response.raise_for_status()

            # Tüm sayfaları yükle
            self.progress_bar.setValue(30)
            self.status_label.setText("📋 Cari sayfası yükleniyor...")
            QApplication.processEvents()
            self.cari_df = pd.read_excel(BytesIO(response.content), sheet_name="Cari")

            self.progress_bar.setValue(45)
            self.status_label.setText("📋 Sevkiyat sayfası yükleniyor...")
            QApplication.processEvents()
            self.sevkiyat_df = pd.read_excel(BytesIO(response.content), sheet_name="Sevkiyat")

            self.progress_bar.setValue(55)
            self.status_label.setText("📋 Bekleyenler sayfası yükleniyor...")
            QApplication.processEvents()
            self.bekleyenler_df = pd.read_excel(BytesIO(response.content), sheet_name="Bekleyenler")

            self.progress_bar.setValue(65)
            self.status_label.setText("📋 Plan sayfası yükleniyor...")
            QApplication.processEvents()
            self.arac_df = pd.read_excel(BytesIO(response.content), sheet_name="Plan")

            self.progress_bar.setValue(75)
            self.status_label.setText("📋 Mail sayfası yükleniyor...")
            QApplication.processEvents()
            mail_df = pd.read_excel(BytesIO(response.content), sheet_name="Mail")
            self.mail_info_df = mail_df[mail_df['fonksiyon'] == 'mail_gonder'].copy()
            self.mail_sevk_info_df = mail_df[mail_df['fonksiyon'] == 'mail_sevk_gonder'].copy()

            self.progress_bar.setValue(85)
            self.status_label.setText("📋 Risk sayfası yükleniyor...")
            QApplication.processEvents()
            self.risk_df = pd.read_excel(BytesIO(response.content), sheet_name="Risk")
            
            # Müşteri adlarını güncelle
            self.progress_bar.setValue(95)
            self.status_label.setText("🔄 Müşteri listesi hazırlanıyor...")
            QApplication.processEvents()

            if not self.cari_df.empty:
                # Sütun isimlerini kontrol et - farklı olasılıkları dene
                cari_column = None
                possible_names = ['Cari Adi', 'Cari Adı', 'CariAdi', 'Cari_Adi', 'cari_adi', 'CARI ADI', 'Müşteri Adı', 'Musteri Adi']

                for col_name in possible_names:
                    if col_name in self.cari_df.columns:
                        cari_column = col_name
                        break

                if cari_column:
                    self.cari_column_name = cari_column  # Sütun adını sakla
                    # Boş değerleri ve null değerleri filtrele
                    self.customer_names = self.cari_df[cari_column].dropna().astype(str).tolist()
                    # Boş string'leri de filtrele
                    self.customer_names = [name.strip() for name in self.customer_names if name.strip()]
                    pass
                else:
                    self.customer_names = []
                    self.cari_column_name = None
                    pass
            else:
                self.customer_names = []
                pass

            # Tüm işlemler tamamlandı
            self.progress_bar.setValue(100)
            QApplication.processEvents()

            # Progress bar'ı 1 saniye sonra gizle
            QTimer.singleShot(1000, lambda: self.progress_bar.setVisible(False))

            self.status_label.setText(f"✅ Veriler başarıyla yüklendi (Cari: {len(self.cari_df)}, Sevkiyat: {len(self.sevkiyat_df)})")
                
        except requests.exceptions.Timeout:
            self.progress_bar.setVisible(False)
            self.status_label.setText("❌ Bağlantı zaman aşımı - Google Sheets'e erişilemiyor")
        except requests.exceptions.RequestException as e:
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"❌ Bağlantı hatası: {str(e)}")
        except Exception as e:
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"❌ Veri yükleme hatası: {str(e)}")
        finally:
            self.set_buttons_enabled(True)
    
    def update_search(self):
        """Arama çubuğunu güncelle"""
        if not self.customer_names:
            return
            
        input_text = self.search_bar.text().strip()
        
        # Eğer arama metni boşsa listeyi temizle
        if not input_text:
            self.result_list.clear()
            return
        
        # Arama metni çok kısaysa (2 karakterden az) arama yapma
        if len(input_text) < 2:
            self.result_list.clear()
            return
            
        try:
            # fuzzywuzzy ile arama yap
            matches = process.extract(input_text, self.customer_names, scorer=process.fuzz.partial_ratio, limit=7)
            self.result_list.clear()
            
            for match in matches:
                if match[1] >= 50:  # Eşik değerini 60'tan 50'ye düşürdüm
                    self.result_list.addItem(match[0])
                    
        except Exception as e:
            self.result_list.clear()
    
    def show_context_menu(self, pos):
        """Bağlam menüsü oluştur"""
        context_menu = QMenu(self)
        context_menu.setStyleSheet("""
            QMenu {
                background-color: #2d2d2d;
                border: 1px solid #404040;
                border-radius: 8px;
                padding: 8px;
                color: #ffffff;
            }
            QMenu::item {
                padding: 8px 16px;
                border-radius: 4px;
            }
            QMenu::item:selected {
                background-color: #007acc;
            }
        """)
        copy_action = QAction("Kopyala", self)
        context_menu.addAction(copy_action)
        copy_action.triggered.connect(self.copy_selected_item)
        context_menu.exec_(self.result_list.mapToGlobal(pos))
    
    def copy_selected_item(self):
        """Seçili öğeyi kopyala"""
        selected_item = self.result_list.currentItem()
        if selected_item:
            clipboard = QApplication.clipboard()
            clipboard.setText(selected_item.text())
            self.status_label.setText("✅ Müşteri adı panoya kopyalandı")
    
    def filter_by_selected_customer(self, item):
        """Seçili müşteriye göre filtreleme yap"""
        try:
            selected_customer = item.text()
            
            if self.cari_df.empty:
                QMessageBox.warning(self, "Hata", "Cari veriler yüklenmemiş!")
                return
            
            # Müşteri bilgilerini al
            if not self.cari_column_name:
                QMessageBox.warning(self, "Hata", "Cari sütun adı bulunamadı!")
                return
                
            customer_rows = self.cari_df[self.cari_df[self.cari_column_name].str.strip() == selected_customer]
            if customer_rows.empty:
                QMessageBox.warning(self, "Hata", "Müşteri bulunamadı!")
                return
            
            customer_row = customer_rows.iloc[0]
            cari_kodu = customer_row['Cari Kodu']
            self.cari_telefon = str(customer_row.get('Telefon', ''))
            self.cari_adi = selected_customer
            
            # Sevkiyat verilerini filtrele
            self.sevkiyat_filtered_data = self.sevkiyat_df[self.sevkiyat_df['Cari Kodu'] == cari_kodu].copy()
            
            if not self.sevkiyat_filtered_data.empty:
                # Veri işleme
                self.sevkiyat_filtered_data["Açıklama"] = self.sevkiyat_filtered_data["Açıklama"].fillna("")
                self.sevkiyat_filtered_data["Kalan Siparis"] = self.sevkiyat_filtered_data["Kalan Siparis"].astype(int).astype(str)
                self.sevkiyat_filtered_data["Toplam Stok"] = self.sevkiyat_filtered_data["Toplam Stok"].astype(int).astype(str)
                
                # Tarih formatı
                if "Tarih" in self.sevkiyat_filtered_data.columns:
                    self.sevkiyat_filtered_data["Tarih"] = pd.to_datetime(self.sevkiyat_filtered_data["Tarih"], errors='coerce')
                    self.sevkiyat_filtered_data["Tarih"] = self.sevkiyat_filtered_data["Tarih"].apply(
                        lambda x: x.strftime("%d.%m.%Y") if pd.notnull(x) and hasattr(x, 'strftime') else "")
                
                if "SPEC" in self.sevkiyat_filtered_data.columns:
                    self.sevkiyat_filtered_data["SPEC"] = self.sevkiyat_filtered_data["SPEC"].fillna("")
            
            

            # Bekleyen verilerini filtrele
            if not self.sevkiyat_filtered_data.empty:
                malzeme_kodlari = self.sevkiyat_filtered_data['Malzeme Kodu'].tolist()
                self.bekleyenler_filtered_data = self.bekleyenler_df[self.bekleyenler_df['Malzeme Kodu'].isin(malzeme_kodlari)].copy()

                if not self.bekleyenler_filtered_data.empty:
                    # Bekleyen verilerini işle
                    self.bekleyenler_filtered_data["Bekleyen Adet"] = self.bekleyenler_filtered_data["Bekleyen Adet"].astype(int).astype(str)
                    
                    # Tarih formatları
                    for date_col in ["Sipariş Tarihi", "Teslimat Tarihi"]:
                        if date_col in self.bekleyenler_filtered_data.columns:
                            self.bekleyenler_filtered_data[date_col] = pd.to_datetime(self.bekleyenler_filtered_data[date_col], errors='coerce')
                            self.bekleyenler_filtered_data[date_col] = self.bekleyenler_filtered_data[date_col].apply(
                                lambda x: x.strftime("%d.%m.%Y") if pd.notnull(x) and hasattr(x, 'strftime') else "")
                    
                    if "Depo Yeri Plaka" in self.bekleyenler_filtered_data.columns:
                        self.bekleyenler_filtered_data["Depo Yeri Plaka"] = self.bekleyenler_filtered_data["Depo Yeri Plaka"].astype(int).astype(str)
                        self.bekleyenler_filtered_data["Depo Yeri Plaka"] = self.bekleyenler_filtered_data["Depo Yeri Plaka"].replace(
                            {"300": "48", "2": "17", "200": "16"})
                    
                    if "Spec Adı" in self.bekleyenler_filtered_data.columns:
                        self.bekleyenler_filtered_data["Spec Adı"] = self.bekleyenler_filtered_data["Spec Adı"].fillna("")
                
                # Araç verilerini filtrele
                if not self.arac_df.empty and 'Malzeme Kodu' in self.arac_df.columns:
                    self.arac_filtered_data = self.arac_df[self.arac_df['Malzeme Kodu'].isin(malzeme_kodlari)].copy()
                else:
                    self.arac_filtered_data = pd.DataFrame()
                
                if not self.arac_filtered_data.empty:
                    self.arac_filtered_data["Adet"] = self.arac_filtered_data["Adet"].astype(int).astype(str)
                    
                    # Tarih formatları
                    for date_col in ["Sipariş Tarihi", "Sevk Tarihi"]:
                        if date_col in self.arac_filtered_data.columns:
                            self.arac_filtered_data[date_col] = pd.to_datetime(self.arac_filtered_data[date_col], errors='coerce')
                            self.arac_filtered_data[date_col] = self.arac_filtered_data[date_col].apply(
                                lambda x: x.strftime("%d.%m.%Y") if pd.notnull(x) and hasattr(x, 'strftime') else "")
                    
                    if "Depo Yeri" in self.arac_filtered_data.columns:
                        self.arac_filtered_data["Depo Yeri"] = self.arac_filtered_data["Depo Yeri"].astype(int).astype(str)
                        self.arac_filtered_data["Depo Yeri"] = self.arac_filtered_data["Depo Yeri"].replace(
                            {"300": "48", "2": "17", "200": "16"})
                    
                    if "Nakliye Numarası" in self.arac_filtered_data.columns:
                        self.arac_filtered_data["Nakliye Numarası"] = self.arac_filtered_data["Nakliye Numarası"].astype(int).astype(str)
                    
                    if "Spec" in self.arac_filtered_data.columns:
                        self.arac_filtered_data["Spec"] = self.arac_filtered_data["Spec"].fillna("")
            
            # Mail verilerini hazırla
            if not self.bekleyenler_filtered_data.empty:
                one_month_ago = datetime.now() - timedelta(days=30)
                if "Sipariş Tarihi" in self.bekleyenler_filtered_data.columns:
                    # Tarih sütununu datetime'a çevir ve karşılaştır
                    bekleyen_mail_data = self.bekleyenler_filtered_data.copy()
                    bekleyen_mail_data["Sipariş_Tarihi_dt"] = pd.to_datetime(bekleyen_mail_data["Sipariş Tarihi"], format="%d.%m.%Y", errors='coerce')
                    bekleyen_mail_data = bekleyen_mail_data[bekleyen_mail_data["Sipariş_Tarihi_dt"] <= one_month_ago]
                    
                    self.mail_data = bekleyen_mail_data[bekleyen_mail_data["Durum"] == "Açık"].copy() if "Durum" in bekleyen_mail_data.columns else pd.DataFrame()
                    self.mail_sevk_data = self.bekleyenler_filtered_data[self.bekleyenler_filtered_data["Durum"] == "Sevke Hazır"].copy() if "Durum" in self.bekleyenler_filtered_data.columns else pd.DataFrame()
                else:
                    self.mail_data = pd.DataFrame()
                    self.mail_sevk_data = pd.DataFrame()
            else:
                self.mail_data = pd.DataFrame()
                self.mail_sevk_data = pd.DataFrame()
            
            # Risk bilgisini al
            risk_tutari = 0
            if not self.risk_df.empty and 'Cari hesap kodu' in self.risk_df.columns:
                cari_riskli = self.risk_df[self.risk_df['Cari hesap kodu'] == cari_kodu]
                if not cari_riskli.empty and 'Risk' in cari_riskli.columns:
                    risk_tutari = cari_riskli["Risk"].sum()
            
            # Sevkiyat verilerini tekrar Malzeme Kodu'na göre filtreleme yap
            if not self.sevkiyat_filtered_data.empty:
                malzeme_kodlari = self.sevkiyat_filtered_data['Malzeme Kodu'].tolist()
                self.sevkiyat_filtered_again = self.sevkiyat_df[self.sevkiyat_df['Malzeme Kodu'].isin(malzeme_kodlari)].copy()
                
                if not self.sevkiyat_filtered_again.empty:
                    self.sevkiyat_filtered_again["SPEC"] = self.sevkiyat_filtered_again["SPEC"].fillna("")
                    self.sevkiyat_filtered_again["Açıklama"] = self.sevkiyat_filtered_again["Açıklama"].fillna("")
                    self.sevkiyat_filtered_again["Kalan Siparis"] = self.sevkiyat_filtered_again["Kalan Siparis"].astype(int)
                    self.sevkiyat_filtered_again["Toplam Stok"] = self.sevkiyat_filtered_again["Toplam Stok"].astype(int)
                    self.sevkiyat_filtered_again["Kalan Siparis"] = self.sevkiyat_filtered_again["Kalan Siparis"].astype(str)
                    self.sevkiyat_filtered_again["Toplam Stok"] = self.sevkiyat_filtered_again["Toplam Stok"].astype(str)
                    self.sevkiyat_filtered_again = self.sevkiyat_filtered_again.sort_values(by=["Malzeme Adı", "Tarih"])
                    self.sevkiyat_filtered_again["Tarih"] = pd.to_datetime(self.sevkiyat_filtered_again["Tarih"], errors='coerce')
                    self.sevkiyat_filtered_again["Tarih"] = self.sevkiyat_filtered_again["Tarih"].apply(lambda x: x.strftime("%d.%m.%Y") if pd.notnull(x) and hasattr(x, 'strftime') else "")
            
            # Bekleyen veriler için Sipariş Tarihi formatlaması - sadece datetime objelerini format et
            if not self.bekleyenler_filtered_data.empty and "Sipariş Tarihi" in self.bekleyenler_filtered_data.columns:
                self.bekleyenler_filtered_data["Sipariş Tarihi"] = self.bekleyenler_filtered_data["Sipariş Tarihi"].apply(lambda x: x.strftime("%d.%m.%Y") if pd.notnull(x) and hasattr(x, 'strftime') else (str(x) if pd.notnull(x) else ""))
            
            # Risk bilgisini al
            risk_tutari = 0
            if not self.risk_df.empty and 'Cari hesap kodu' in self.risk_df.columns:
                cari_riskli = self.risk_df[self.risk_df['Cari hesap kodu'] == cari_kodu]
                if not cari_riskli.empty and 'Risk' in cari_riskli.columns:
                    risk_tutari = cari_riskli["Risk"].sum()
            
            # Müşteri bilgi butonunu güncelle - HTML ile sola/sağa yaslama
            # Müşteri bilgi butonunu güncelle - HTML ile sola/sağa yaslama
            # Telefon numarasındaki .0'ı temizle
            formatted_phone = str(self.cari_telefon).replace('.0', '') if self.cari_telefon else ""

            if risk_tutari == 0:
                button_text = f"{self.cari_adi} : {formatted_phone}"
            else:
                # Sol tarafa müşteri bilgisi, sağ tarafa risk tutarı
                spaces_needed = max(0, 60 - len(f"{self.cari_adi} : {formatted_phone}") - len(f"Risk: {risk_tutari}"))
                button_text = f"{self.cari_adi} : {formatted_phone}{' ' * spaces_needed}Risk: {risk_tutari}"
            
            self.sozlesmedeki_urunler_button.setText(button_text)

            # HTML içeriği oluştur
            self.create_html_content()
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            QMessageBox.critical(self, "Hata", f"Filtreleme sırasında hata oluştu: {str(e)}\n\nDetay:\n{error_details}")
            self.status_label.setText(f"❌ Filtreleme hatası: {str(e)}")
    
    def create_html_content(self):
        """HTML içeriği oluştur"""
        try:
            style = """
            <style>
                body {
                    background-color: #c00c0c;
                    color: #26b47e;
                    font-family: 'Segoe UI', Arial, sans-serif;
                    font-size: 22px;
                    margin: 20px;
                    padding: 0;
                }
                
                h2 {
                    color: #000000;
                    font-weight: 600;
                    margin: 25px 0 15px 0;
                    padding: 0;
                    font-size: 16px;
                }
                
                table {
                    width: 100%;
                    border-collapse: collapse;
                    margin: 15px 0;
                    font-size: 11px;
                    background-color: #000000;
                    border: 1px solid #000000;
                    border-radius: 4px;
                    overflow: hidden;

                }
                
                th {
                    padding: 8px 6px;
                    border: 1px solid #000000;
                    text-align: center;
                    background-color: #000000;
                    color: #ffffff;
                    font-weight: 600;
                    font-size: 13px;
                    font-weight: bold;
                    letter-spacing: 0.5px;
                }
                
                td {
                    padding: 6px;
                    border: 1px solid #000000;
                    text-align: left;
                    background-color: #ffffff;
                    color: #000000;
                    vertical-align: middle;
                    white-space: nowrap;
                    font-size: 13px;
                    font-weight: bold;
                }

                tr.stok-yetersiz td {
                    background-color: #f8d7da !important;
                }

                tr.sevke-hazir td {
                    background-color: #d4edda !important;
                }
                
                tr.secilen-cari td {
                    background-color: #c8e6c9 !important;
                }
                
            </style>
            """
            
            html_content = style + "<body>"
            
            # Sevkiyat tablosu - Stok karşılaştırması ile renklendirme
            if not self.sevkiyat_filtered_data.empty:
                html_content += "<h2>Sevkiyat Bilgileri</h2>"
                html_content += '<table>'
                
                # Header
                html_content += '<thead><tr>'
                for col in self.sevkiyat_filtered_data.columns:
                    html_content += f'<th>{col}</th>'
                html_content += '</tr></thead><tbody>'
                
                # Data rows with conditional formatting
                for _, row in self.sevkiyat_filtered_data.iterrows():
                    try:
                        toplam_stok = int(str(row.get('Toplam Stok', 0)).replace(',', '')) if str(row.get('Toplam Stok', 0)).replace(',', '').isdigit() else 0
                        kalan_siparis = int(str(row.get('Kalan Siparis', 0)).replace(',', '')) if str(row.get('Kalan Siparis', 0)).replace(',', '').isdigit() else 0

                        if toplam_stok < kalan_siparis:
                            row_class = 'stok-yetersiz'
                        else:
                            row_class = ''
                    except:
                        row_class = ''

                    html_content += f'<tr class="{row_class}">'
                    for col in self.sevkiyat_filtered_data.columns:
                        html_content += f'<td>{row.get(col, "")}</td>'
                    html_content += '</tr>'
                
                html_content += '</tbody></table>'
            
            # Bekleyen ürünler tablosu - Sevke Hazır renklendirmesi
            if not self.bekleyenler_filtered_data.empty:
                html_content += "<h2>Bekleyen Ürünler</h2>"
                bekleyenler_display = self.bekleyenler_filtered_data.copy()
                if "KDV(%)" in bekleyenler_display.columns:
                    bekleyenler_display = bekleyenler_display.drop(columns=["KDV(%)"])
                
                html_content += '<table>'
                
                # Header
                html_content += '<thead><tr>'
                for col in bekleyenler_display.columns:
                    html_content += f'<th>{col}</th>'
                html_content += '</tr></thead><tbody>'
                
                # Data rows with conditional formatting
                for _, row in bekleyenler_display.iterrows():
                    try:
                        if str(row.get('Durum', '')) == 'Sevke Hazır':
                            row_class = 'sevke-hazir'
                        else:
                            row_class = ''
                    except:
                        row_class = ''
                    
                    html_content += f'<tr class="{row_class}">'
                    for col in bekleyenler_display.columns:
                        html_content += f'<td>{row.get(col, "")}</td>'
                    html_content += '</tr>'
                
                html_content += '</tbody></table>'
            
            # Araç tablosu
            if not self.arac_filtered_data.empty:
                html_content += "<h2>Planlanan Araç Bilgileri</h2>"
                html_content += '<table>'
                
                # Header
                html_content += '<thead><tr>'
                for col in self.arac_filtered_data.columns:
                    html_content += f'<th>{col}</th>'
                html_content += '</tr></thead><tbody>'
                
                # Data rows
                for _, row in self.arac_filtered_data.iterrows():
                    html_content += '<tr>'
                    for col in self.arac_filtered_data.columns:
                        html_content += f'<td>{row.get(col, "")}</td>'
                    html_content += '</tr>'
                
                html_content += '</tbody></table>'
            
            # Malzeme Bazlı Kalan Sevkiyatlar tablosu - Gelişmiş renklendirme sistemi
            if not self.sevkiyat_filtered_again.empty:
                html_content += "<h2>Malzeme Bazlı Kalan Sevkiyatlar</h2>"
                html_content += '<table>'
                
                # Header
                html_content += '<thead><tr>'
                for col in self.sevkiyat_filtered_again.columns:
                    html_content += f'<th>{col}</th>'
                html_content += '</tr></thead><tbody>'
                
                # Malzeme kodlarına göre grupla ve kümülatif hesaplama için
                cumulative_tracker = {}
                
                # Data rows with conditional formatting
                for _, row in self.sevkiyat_filtered_again.iterrows():
                    row_class = ''
                    
                    try:
                        malzeme_kodu = str(row.get('Malzeme Kodu', ''))
                        kalan_siparis = int(str(row.get('Kalan Siparis', 0)).replace(',', '')) if str(row.get('Kalan Siparis', 0)).replace(',', '').isdigit() else 0
                        toplam_stok = int(str(row.get('Toplam Stok', 0)).replace(',', '')) if str(row.get('Toplam Stok', 0)).replace(',', '').isdigit() else 0
                        cari_adi = str(row.get('Cari Adi', ''))
                        
                        # Her malzeme kodu için kümülatif toplamı takip et
                        if malzeme_kodu not in cumulative_tracker:
                            cumulative_tracker[malzeme_kodu] = 0
                        
                        # Bu satırdaki Kalan Sipariş'i kümülatif toplama ekle
                        cumulative_tracker[malzeme_kodu] += kalan_siparis
                        
                        # Sadece seçili cari için renklendirme yap
                        if self.cari_adi and cari_adi == self.cari_adi:
                            # Kümülatif Kalan Sipariş ile Toplam Stok karşılaştırması
                            if cumulative_tracker[malzeme_kodu] <= toplam_stok:
                                row_class = 'secilen-cari'  # Yeşil
                            else:
                                row_class = 'stok-yetersiz'  # Açık kırmızı
                                
                    except Exception as e:
                        row_class = ''
                    
                    html_content += f'<tr class="{row_class}">'
                    for col in self.sevkiyat_filtered_again.columns:
                        html_content += f'<td>{row.get(col, "")}</td>'
                    html_content += '</tr>'
                
                html_content += '</tbody></table>'

            html_content += "</body>"
            self.filtered_label.setText(html_content)
            
        except Exception as e:
            self.filtered_label.setText(f"<h3 style='color: #ff6b6b;'>HTML oluşturma hatası: {str(e)}</h3>")
    
    def whatsapp_randevu_gonder(self):
        """WhatsApp randevu mesajı gönder - Stok yetersizliği kontrolü ile"""
        if not self.cari_adi or not self.cari_telefon:
            QMessageBox.warning(self, "Hata", "Lütfen önce bir müşteri seçin!")
            return
        
        # Stok yetersizliği kontrolü yap
        if not self.sevkiyat_filtered_again.empty:
            problematic_products = self._check_stock_insufficiency()
            
            if problematic_products:
                # Uyarı mesajı oluştur
                warning_message = f"{self.cari_adi} için stok yetersizdir.\n"
                
                for product_info in problematic_products:
                    warning_message += f"🔴 {product_info['malzeme_adi']}   :  {product_info['toplam_stok']} adet\n"
                    
                    for cari_info in product_info['other_customers']:
                        warning_message += f"     • {cari_info['kalan_siparis']} : {cari_info['cari_adi']}\n"
                    warning_message += "\n"
                
                warning_message += "Yine de randevu vermek istermisiniz?"
                
                # İçerik uzunluğuna göre dinamik boyut hesapla
                line_count = warning_message.count('\n') + 1
                max_line_length = max(len(line) for line in warning_message.split('\n'))
                
                # Daha akıllı genişlik hesaplaması (karakter başına pixel)
                estimated_width = max_line_length * 9  # Daha gerçekçi karakter genişliği
                estimated_height = line_count * 22     # Satır yüksekliği
                
                # Ekran boyutlarına göre maksimum sınırlar
                screen_width = 1200  # Makul maksimum genişlik
                screen_height = 800  # Makul maksimum yükseklik
                
                # Dinamik boyutları hesapla
                content_width = min(max(estimated_width + 50, 400), screen_width - 200)
                content_height = min(max(estimated_height + 50, 150), screen_height - 300)
                
                # Pencere boyutları (content + butonlar + padding)
                window_width = content_width + 80
                window_height = content_height + 150
                
                # Custom message box oluştur
                msg_box = QMessageBox(self)
                msg_box.setWindowTitle("Stok Yetersizliği Uyarısı")
                msg_box.resize(window_width, window_height)
                # Icon'u kaldır - ünlem işareti görünmesin
                
                # QTextBrowser kullanarak scrollable metin alanı oluştur
                from PyQt5.QtWidgets import QTextBrowser, QVBoxLayout, QWidget
                
                # İçerik widget'ı oluştur
                content_widget = QWidget()
                layout = QVBoxLayout(content_widget)
                
                # Scrollable text browser - dinamik boyutlar
                text_browser = QTextBrowser()
                text_browser.setPlainText(warning_message)
                text_browser.setFixedSize(content_width, content_height)
                text_browser.setStyleSheet("""
                    QTextBrowser {
                        background-color: #1a1a1a;
                        color: #ffffff;
                        font-size: 15px;
                        font-weight: bold;
                        border: 1px solid #404040;
                        border-radius: 8px;
                        padding: 10px;
                        selection-background-color: #0078d4;
                    }
                    QScrollBar:vertical {
                        background: #2d2d2d;
                        width: 16px;
                        border-radius: 8px;
                        margin: 0px;
                    }
                    QScrollBar::handle:vertical {
                        background: #555555;
                        border-radius: 8px;
                        min-height: 25px;
                        margin: 2px;
                    }
                    QScrollBar::handle:vertical:hover {
                        background: #777777;
                    }
                    QScrollBar::handle:vertical:pressed {
                        background: #888888;
                    }
                    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                        border: none;
                        background: none;
                        height: 0px;
                    }
                    QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                        background: none;
                    }
                """)
                
                layout.addWidget(text_browser)
                msg_box.layout().addWidget(content_widget, 1, 0, 1, msg_box.layout().columnCount())
                
                # MessageBox için temel styling
                msg_box.setStyleSheet("""
                    QMessageBox {
                        background-color: #2d2d2d;
                        color: #ffffff;
                    }
                    QPushButton {
                        font-size: 13px;
                        font-weight: bold;
                        padding: 12px 20px;
                        margin: 8px;
                        border-radius: 6px;
                        min-width: 120px;
                    }
                    QPushButton[text="İptal"] {
                        background-color: #dc3545;
                        color: white;
                        border: 2px solid #dc3545;
                    }
                    QPushButton[text="İptal"]:hover {
                        background-color: #c82333;
                        border-color: #bd2130;
                    }
                    QPushButton[text="Randevu Al"] {
                        background-color: #28a745;
                        color: white;
                        border: 2px solid #28a745;
                    }
                    QPushButton[text="Randevu Al"]:hover {
                        background-color: #218838;
                        border-color: #1e7e34;
                    }
                """)
                
                # Butonları ekle
                iptal_btn = msg_box.addButton("İptal", QMessageBox.RejectRole)
                randevu_btn = msg_box.addButton("Randevu Al", QMessageBox.AcceptRole)
                
                msg_box.exec_()
                
                # Kullanıcının seçimine göre işlem yap
                if msg_box.clickedButton() == randevu_btn:
                    self._send_randevu_message()
                else:
                    self.status_label.setText("❌ Randevu gönderimi iptal edildi")
                    return
            else:
                # Stok problemi yoksa direkt gönder
                self._send_randevu_message()
        else:
            # Veri yoksa direkt gönder
            self._send_randevu_message()
    
    def _check_stock_insufficiency(self):
        """Stok yetersizliği kontrolü yap"""
        problematic_products = []
        cumulative_tracker = {}
        
        try:
            for _, row in self.sevkiyat_filtered_again.iterrows():
                malzeme_kodu = str(row.get('Malzeme Kodu', ''))
                malzeme_adi = str(row.get('Malzeme Adı', ''))
                kalan_siparis = int(str(row.get('Kalan Siparis', 0)).replace(',', '')) if str(row.get('Kalan Siparis', 0)).replace(',', '').isdigit() else 0
                toplam_stok = int(str(row.get('Toplam Stok', 0)).replace(',', '')) if str(row.get('Toplam Stok', 0)).replace(',', '').isdigit() else 0
                cari_adi = str(row.get('Cari Adi', ''))
                
                # Her malzeme kodu için kümülatif toplamı takip et
                if malzeme_kodu not in cumulative_tracker:
                    cumulative_tracker[malzeme_kodu] = {
                        'malzeme_adi': malzeme_adi,
                        'toplam_stok': toplam_stok,
                        'cumulative_sum': 0,
                        'customers': []
                    }
                
                # Bu satırdaki Kalan Sipariş'i kümülatif toplama ekle
                cumulative_tracker[malzeme_kodu]['cumulative_sum'] += kalan_siparis
                cumulative_tracker[malzeme_kodu]['customers'].append({
                    'cari_adi': cari_adi,
                    'kalan_siparis': kalan_siparis
                })
                
                # Seçili cari için stok yetersizliği kontrolü
                if (self.cari_adi and cari_adi == self.cari_adi and 
                    cumulative_tracker[malzeme_kodu]['cumulative_sum'] > toplam_stok):
                    
                    # Bu ürün için problemli durumu kaydet (sadece bir kez)
                    already_added = any(p['malzeme_kodu'] == malzeme_kodu for p in problematic_products)
                    if not already_added:
                        # Bu ürünü alan diğer carileri topla (seçili cari hariç)
                        other_customers = []
                        for customer in cumulative_tracker[malzeme_kodu]['customers']:
                            if customer['cari_adi'] != self.cari_adi:
                                other_customers.append(customer)
                        
                        problematic_products.append({
                            'malzeme_kodu': malzeme_kodu,
                            'malzeme_adi': malzeme_adi,
                            'toplam_stok': toplam_stok,
                            'other_customers': other_customers
                        })
            
            return problematic_products
            
        except Exception as e:
            self.status_label.setText(f"❌ Stok kontrol hatası: {str(e)}")
            return []
    
    def _send_randevu_message(self):
        """Randevu mesajını gönder"""
        message = f"""Merhaba {self.cari_adi}, Batman Doğtaş Mobilya'dan aldığınız ürünlerin teslimatı montaj ekibimiz tarafından "YARIN GÜN İÇİNDE" yapılacaktır. Müsaitlik durumunuz hakkında lütfen bilgi verebilir misiniz?
            
    Evet. Onaylıyorum. 
    Hayır. Müsait değilim."""
        
        self._send_whatsapp_message(message)
    
    def stok_analizi_goster(self):
        """Stok analizi penceresini göster"""
        if not self.cari_adi:
            QMessageBox.warning(self, "Hata", "Lütfen önce bir müşteri seçin!")
            return
        
        # Stok yetersizliği kontrolü yap
        if not self.sevkiyat_filtered_again.empty:
            problematic_products = self._check_stock_insufficiency()
            
            # Analiz mesajı oluştur
            analysis_message = f"{self.cari_adi} için detaylı stok analizi:\n\n"
            
            if problematic_products:
                for product_info in problematic_products:
                    analysis_message += f"🔴 {product_info['malzeme_adi']}  :  Stok miktarı {product_info['toplam_stok']} adet\n"
                    
                    # Bu ürünü sipariş veren TÜM carileri bul (seçili cari dahil)
                    all_customers_for_product = []
                    
                    # Sevkiyat tablosundan bu ürün için tüm carileri topla
                    for _, row in self.sevkiyat_filtered_again.iterrows():
                        if str(row.get('Malzeme Adı', '')) == product_info['malzeme_adi']:
                            cari_adi = str(row.get('Cari Adi', ''))
                            kalan_siparis = int(str(row.get('Kalan Siparis', 0)).replace(',', '')) if str(row.get('Kalan Siparis', 0)).replace(',', '').isdigit() else 0
                            tarih = str(row.get('Tarih', ''))

                            # Aynı cari birden fazla kez yoksa ekle
                            existing_customer = next((c for c in all_customers_for_product if c['cari_adi'] == cari_adi), None)
                            if not existing_customer:
                                all_customers_for_product.append({
                                    'cari_adi': cari_adi,
                                    'kalan_siparis': kalan_siparis,
                                    'tarih': tarih
                                })
                    
                    # Carileri tarihe göre sırala (eskiden yeniye)
                    # Tarih formatı datetime'a çevir ve sırala
                    def parse_date(tarih_str):
                        try:
                            from datetime import datetime
                            # Tarih formatı: DD.MM.YYYY
                            return datetime.strptime(tarih_str, "%d.%m.%Y")
                        except:
                            # Tarih parse edilemezse çok ileride bir tarih döndür (en sona atsın)
                            return datetime(2099, 12, 31)

                    all_customers_for_product.sort(key=lambda x: parse_date(x['tarih']))
                    
                    # Tüm carileri göster - seçili cariyi yeşil renkte vurgula
                    for cari_info in all_customers_for_product:
                        if self.cari_adi and cari_info['cari_adi'] == self.cari_adi:
                            # HTML formatında yeşil renk (QTextBrowser HTML destekler)
                            analysis_message += f"     • <span style='color: #28a745; font-weight: bold;'>{cari_info['tarih']}  :  {cari_info['kalan_siparis']}  :  {cari_info['cari_adi']}</span>\n"
                        else:
                            analysis_message += f"     • {cari_info['tarih']}  :  {cari_info['kalan_siparis']}  :  {cari_info['cari_adi']}\n"
                    analysis_message += "\n"
            
            # Eğer stok yetersizliği yoksa bilgi ver
            if not problematic_products:
                analysis_message += "✅ Tüm ürünlerde stok yeterli! Herhangi bir sorun tespit edilmedi."
            
            # Büyük analiz penceresi göster (2 kat büyük)
            self._show_analysis_window(analysis_message, title="Detaylı Stok Analizi")
        else:
            QMessageBox.information(self, "Bilgi", "Analiz edilecek stok verisi bulunamadı!")
    
    def _show_analysis_window(self, message, title="Analiz"):
        """Büyük analiz penceresini göster (randevu uyarısının 2 katı)"""
        # İçerik uzunluğuna göre dinamik boyut hesapla (2 kat büyük)
        line_count = message.count('\n') + 1
        max_line_length = max(len(line) for line in message.split('\n'))
        
        # 2 katına çıkarılmış boyut hesaplaması
        estimated_width = max_line_length * 18  # 9*2
        estimated_height = line_count * 44      # 22*2
        
        # Ekran boyutlarına göre maksimum sınırlar (2 kat büyük)
        screen_width = 1600  # 1200 * 1.33
        screen_height = 1000 # 800 * 1.25
        
        # Dinamik boyutları hesapla (2 kat büyük)
        content_width = min(max(estimated_width + 100, 800), screen_width - 200)  # min 800
        content_height = min(max(estimated_height + 100, 300), screen_height - 300)  # min 300
        
        # Pencere boyutları
        window_width = content_width + 160  # 80*2
        window_height = content_height + 300 # 150*2
        
        # Custom message box oluştur
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle(title)
        msg_box.resize(window_width, window_height)
        
        # QTextBrowser kullanarak scrollable metin alanı oluştur
        from PyQt5.QtWidgets import QTextBrowser, QVBoxLayout, QWidget
        
        # İçerik widget'ı oluştur
        content_widget = QWidget()
        layout = QVBoxLayout(content_widget)
        
        # Scrollable text browser - 2 kat büyük
        text_browser = QTextBrowser()
        # HTML formatını desteklemek için setHtml kullan
        html_message = message.replace('\n', '<br>')
        text_browser.setHtml(f"<div style='color: white; font-family: monospace; white-space: pre;'>{html_message}</div>")
        text_browser.setFixedSize(content_width, content_height)
        text_browser.setStyleSheet("""
            QTextBrowser {
                background-color: #1a1a1a;
                color: #ffffff;
                font-size: 16px;
                font-weight: bold;
                border: 1px solid #404040;
                border-radius: 8px;
                padding: 15px;
                selection-background-color: #0078d4;
            }
            QScrollBar:vertical {
                background: #2d2d2d;
                width: 20px;
                border-radius: 10px;
                margin: 0px;
            }
            QScrollBar::handle:vertical {
                background: #555555;
                border-radius: 10px;
                min-height: 30px;
                margin: 2px;
            }
            QScrollBar::handle:vertical:hover {
                background: #777777;
            }
            QScrollBar::handle:vertical:pressed {
                background: #888888;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                border: none;
                background: none;
                height: 0px;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none;
            }
        """)
        
        layout.addWidget(text_browser)
        msg_box.layout().addWidget(content_widget, 1, 0, 1, msg_box.layout().columnCount())
        
        # MessageBox için temel styling
        msg_box.setStyleSheet("""
            QMessageBox {
                background-color: #2d2d2d;
                color: #ffffff;
            }
            QPushButton {
                font-size: 16px;
                font-weight: bold;
                padding: 15px 25px;
                margin: 10px;
                border-radius: 8px;
                min-width: 150px;
            }
            QPushButton[text="Tamam"] {
                background-color: #007acc;
                color: white;
                border: 2px solid #007acc;
            }
            QPushButton[text="Tamam"]:hover {
                background-color: #005a9e;
                border-color: #005a9e;
            }
            QPushButton[text="Mail Gönder"] {
                background-color: #28a745;
                color: white;
                border: 2px solid #28a745;
            }
            QPushButton[text="Mail Gönder"]:hover {
                background-color: #218838;
                border-color: #1e7e34;
            }
        """)
        
        # Tamam ve Mail Gönder butonları
        tamam_btn = msg_box.addButton("Tamam", QMessageBox.AcceptRole)
        mail_gonder_btn = msg_box.addButton("Mail Gönder", QMessageBox.ActionRole)
        
        msg_box.exec_()
        
        # Kullanıcının seçimine göre işlem yap
        if msg_box.clickedButton() == mail_gonder_btn:
            # Önce açık sipariş maili gönder, sonra sevke hazır maili gönder
            self._sequential_mail_send()
    
    def _sequential_mail_send(self):
        """Sırayla mail gönder: önce açık sipariş, sonra sevke hazır"""
        try:
            # İlk önce açık sipariş mailini gönder
            if hasattr(self, 'mail_data') and not self.mail_data.empty:
                self.status_label.setText("📧 Açık sipariş maili gönderiliyor...")
                QApplication.processEvents()
                self.mail_gonder(self.mail_data, self.cari_adi)
                
                # Kısa bir bekleme
                QTimer.singleShot(1000, self._send_sevk_mail)
            else:
                # Açık sipariş verisi yoksa direkt sevke hazır gönder
                self._send_sevk_mail()
                
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Mail gönderme hatası: {str(e)}")
            self.status_label.setText(f"❌ Mail gönderme hatası: {str(e)}")
    
    def _send_sevk_mail(self):
        """Sevke hazır mailini gönder"""
        try:
            if hasattr(self, 'mail_sevk_data') and not self.mail_sevk_data.empty:
                self.status_label.setText("📧 Sevke hazır maili gönderiliyor...")
                QApplication.processEvents()
                self.mail_sevk_gonder(self.mail_sevk_data, self.cari_adi)
            else:
                self.status_label.setText("ℹ️ Sevke hazır gönderilecek veri bulunamadı")
                
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Sevke hazır mail hatası: {str(e)}")
            self.status_label.setText(f"❌ Sevke hazır mail hatası: {str(e)}")
    
    def whatsapp_bilgi_gonder(self):
        """WhatsApp bilgi mesajı gönder"""
        if not self.cari_adi or not self.cari_telefon:
            QMessageBox.warning(self, "Hata", "Lütfen önce bir müşteri seçin!")
            return
        
        message = f"Merhaba {self.cari_adi}"
        
        self._send_whatsapp_message(message)
    
    def _send_whatsapp_message(self, message):
        """WhatsApp mesajı gönderme ortak fonksiyonu"""
        try:
            # 1. Ham veriyi stringe çevir ve temizle
            phone = str(self.cari_telefon).strip()
            
            # 2. Eğer sayı sonu .0 ile bitiyorsa (Float hatası), o kısmı sil
            if phone.endswith(".0"):
                phone = phone[:-2]
            
            # 3. Sadece rakamları tut (boşluk, tire, + gibi karakterleri temizler)
            phone = "".join(filter(str.isdigit, phone))
            
            # 4. Türkiye formatına getir (Hedef: 905321234567)
            if phone.startswith("0"):
                phone = "90" + phone[1:]
            elif len(phone) == 10: # 532... formatındaysa
                phone = "90" + phone
            
            # 5. Validasyon (Türkiye numaraları 12 hanedir)
            if len(phone) != 12:
                QMessageBox.warning(self, "Hata", f"Geçersiz telefon numarası!\nNumara: {phone}\nLütfen 10 haneli (532...) olarak kontrol edin.")
                return
            
            # Mesaj hazırlama ve gönderme
            pyperclip.copy(message)
            encoded_message = urllib.parse.quote(message)
            
            # Daha stabil olan wa.me linkini kullanmanızı öneririm
            url = f"whatsapp://send?phone={phone}&text={encoded_message}"
            webbrowser.open(url)
            
            self.status_label.setText("✅ WhatsApp mesajı hazırlandı")
            
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Bir hata oluştu: {str(e)}")
    
    def mail_gonder_button_clicked(self):
        """Mail gönder butonuna tıklandığında"""
        self.mail_gonder(self.mail_data, self.cari_adi)
    
    def sevk_button_clicked(self):
        """Sevk butonu tıklandığında"""
        self.mail_sevk_gonder(self.mail_sevk_data, self.cari_adi)
    
    def mail_gonder(self, mail_data, cari_adi):
        """Mail gönder"""
        if mail_data.empty:
            QMessageBox.information(self, "Veri Bulunamadı", f"{cari_adi}\nGönderilecek veri bulunamadı.\nSipariş tarihi 1 ay önce olan kalemler için mail gönderilmektedir.")
            return
        
        if self.mail_info_df.empty:
            QMessageBox.warning(self, "Hata", "Mail bilgileri yüklenmemiş!")
            return
        
        try:
            # Veri işleme
            processed_mail_data = mail_data.drop_duplicates(subset=['Malzeme Kodu'], keep='first').copy()

            # Mail için gereksiz sütunları çıkar
            columns_to_remove = ["Malzeme Kodu", "Prosap Sözleşme Ad Soyad", "Sipariş_Tarihi_dt", "KDV(%)"]
            for col in columns_to_remove:
                if col in processed_mail_data.columns:
                    processed_mail_data = processed_mail_data.drop(columns=[col])

            # Kalem No formatını düzenle
            processed_mail_data = self.format_kalem_no(processed_mail_data)

            if processed_mail_data.empty:
                QMessageBox.information(self, "Bilgi", "Filtreleme sonrası gönderilecek veri bulunamadı.")
                return
            
            # Mail bilgilerini al
            mail_info = self.mail_info_df.iloc[0]
            sender_email = mail_info["sender_email"]
            receiver_email = mail_info["receiver_email"]
            receiver_name = mail_info["receiver_name"]
            cc_emails = str(mail_info["cc_email"]).split(',') if pd.notna(mail_info["cc_email"]) else []
            bcc_email = str(mail_info["bcc_email"]) if pd.notna(mail_info["bcc_email"]) else ""
            password = mail_info["password"]
            smtp_server = mail_info["smtp_server"]
            
            subject = f"Güneşler - {cari_adi} bekleyen ürünleri hk."
            
            body = f"""
            <p>Merhaba {receiver_name},</p>
            <p>Ekteki ürünlerin sevk tarihi konusunda yardımcı olabilir misiniz?</p>
            {processed_mail_data.to_html(index=False)}
            <p>İyi çalışmalar diliyorum.</p>
            """
            
            # E-posta oluşturma
            msg = MIMEMultipart()
            msg["From"] = str(Header(sender_email, "utf-8"))
            msg["To"] = str(Header(receiver_email, "utf-8"))
            msg["Cc"] = ', '.join(cc_emails)
            msg["Subject"] = str(Header(subject, "utf-8"))
            msg.attach(MIMEText(body, "html", "utf-8"))

            to_addrs = [receiver_email] + cc_emails + ([bcc_email] if bcc_email else [])
            
            # Kullanıcıdan onay alma
            reply = QMessageBox.question(self, "E-posta Gönderimi", 
                                       f"{cari_adi} için e-posta göndermek istediğinizden emin misiniz?", 
                                       QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                with smtplib.SMTP(smtp_server, 587) as server:
                    server.starttls()
                    server.login(sender_email, password)
                    server.sendmail(sender_email, to_addrs, msg.as_string())
                    QMessageBox.information(self, "E-posta Gönderildi", 
                                          f"{cari_adi}\n\nE-posta başarıyla gönderildi.\nKime : {receiver_email}\nBilgi : {', '.join(cc_emails)}\n")
                    self.status_label.setText("✅ E-posta başarıyla gönderildi")
            else:
                QMessageBox.information(self, "E-posta Gönderilmedi", "E-posta gönderimi iptal edildi.")
                
        except Exception as e:
            QMessageBox.critical(self, "E-posta Gönderme Hatası", f"E-posta gönderme hatası: {e}")
            self.status_label.setText(f"❌ E-posta gönderme hatası: {str(e)}")
    
    def mail_sevk_gonder(self, mail_sevk_data, cari_adi):
        """Sevk maili gönder"""
        if mail_sevk_data.empty:
            QMessageBox.information(self, "Veri Bulunamadı", f"{cari_adi}\nGönderilecek veri bulunamadı.")
            return

        if self.mail_sevk_info_df.empty:
            QMessageBox.warning(self, "Hata", "Mail sevk bilgileri yüklenmemiş!")
            return

        # Depo bilgilerini ayar sayfasından yükle
        depolar = self.load_depo_settings()

        for plaka, depo in depolar.items():
            if "Depo Yeri Plaka" in mail_sevk_data.columns and plaka in mail_sevk_data["Depo Yeri Plaka"].values:
                self.depo = depo
                mail_sevk_govde_data = mail_sevk_data[mail_sevk_data["Depo Yeri Plaka"] == plaka]
                self.mail_sevk_govde_fonk(mail_sevk_govde_data, self.depo, cari_adi)
    
    def mail_sevk_govde_fonk(self, mail_sevk_govde, depo, cari_adi):
        """Sevk mail gövdesi fonksiyonu"""
        try:
            # Mail sevk bilgilerini al
            mail_sevk_info = self.mail_sevk_info_df.iloc[0]
            sender_email = mail_sevk_info["sender_email"]
            receiver_email = mail_sevk_info["receiver_email"]
            cc_email = str(mail_sevk_info["cc_email"]) if pd.notna(mail_sevk_info["cc_email"]) else ""
            bcc_email = str(mail_sevk_info["bcc_email"]) if pd.notna(mail_sevk_info["bcc_email"]) else ""
            password = mail_sevk_info["password"]
            smtp_server = mail_sevk_info["smtp_server"]
            
            subject = f"{depo} BAYİ SEVK"

            # Mail için gereksiz sütunları çıkar (orijinal veriyi korumak için kopya oluştur)
            mail_display_data = mail_sevk_govde.copy()
            columns_to_remove = ["Malzeme Kodu", "KDV(%)", "Prosap Sözleşme Ad Soyad"]
            for col in columns_to_remove:
                if col in mail_display_data.columns:
                    mail_display_data = mail_display_data.drop(columns=[col])

            # Kalem No formatını düzenle
            mail_display_data = self.format_kalem_no(mail_display_data)

            body = f"""
            <p>Merhaba,</p>
            <p>Ekteki ürünlerin ilk sevkiyat planına alınması için yardımcı olabilir misiniz?</p>
            {mail_display_data.to_html(index=False)}
            <p>İyi çalışmalar diliyorum.</p>
            """
            
            # E-posta oluşturma
            msg = MIMEMultipart()
            msg["From"] = str(Header(sender_email, "utf-8"))
            msg["To"] = str(Header(receiver_email, "utf-8"))
            msg["Cc"] = str(Header(cc_email, "utf-8"))
            msg["Subject"] = str(Header(subject, "utf-8"))
            msg.attach(MIMEText(body, "html", "utf-8"))

            to_addrs = [receiver_email] + ([cc_email] if cc_email else []) + ([bcc_email] if bcc_email else [])
            
            # Kullanıcıdan onay alma
            reply = QMessageBox.question(self, "E-posta Gönderimi", 
                                       f"\n{cari_adi} için e-posta göndermek istediğinizden emin misiniz?\n{depo} Depodaki ürünler için gönderilecektir.", 
                                       QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            
            if reply == QMessageBox.Yes:
                with smtplib.SMTP(smtp_server, 587) as server:
                    server.starttls()
                    server.login(sender_email, password)
                    server.sendmail(sender_email, to_addrs, msg.as_string())
                    QMessageBox.information(self, "E-posta Gönderildi", 
                                          f"{cari_adi}\n\nE-posta başarıyla gönderildi.\nKime : {receiver_email}\nBilgi : {cc_email}\n")
                    self.status_label.setText("✅ Sevk e-postası başarıyla gönderildi")
            else:
                QMessageBox.information(self, "E-posta Gönderilmedi", "E-posta gönderimi iptal edildi.")
                
        except Exception as e:
            QMessageBox.critical(self, "E-posta Gönderme Hatası", f"E-posta gönderme hatası: {e}")
            self.status_label.setText(f"❌ Sevk e-posta gönderme hatası: {str(e)}")
    
    def export_sevkiyat_to_excel(self):
        """Sevkiyat verilerini Excel'e aktar"""
        try:
            if not self.cari_adi:
                QMessageBox.warning(self, "Hata", "Önce bir müşteri seçin!")
                return
            
            if self.sevkiyat_filtered_data.empty:
                QMessageBox.warning(self, "Hata", "Sevkiyat verisi bulunamadı!")
                return
            
            output_path = f"D:/GoogleDrive/~ {self.cari_adi}_Sevkiyat.xlsx"
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                self.sevkiyat_filtered_data.to_excel(writer, sheet_name='Sevkiyat', index=False)
            
            self.status_label.setText(f"✅ Sevkiyat verileri dışa aktarıldı: {output_path}")
            QMessageBox.information(self, "Başarılı", f"Sevkiyat verileri başarıyla dışa aktarıldı:\n{output_path}")
            
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Sevkiyat dışa aktarma hatası: {str(e)}")
            self.status_label.setText(f"❌ Sevkiyat dışa aktarma hatası: {str(e)}")
    
    def export_bekleyenler_to_excel(self):
        """Bekleyenler verilerini Excel'e aktar"""
        try:
            if not self.cari_adi:
                QMessageBox.warning(self, "Hata", "Önce bir müşteri seçin!")
                return
            
            if self.bekleyenler_filtered_data.empty:
                QMessageBox.warning(self, "Hata", "Bekleyenler verisi bulunamadı!")
                return
            
            output_path = f"D:/GoogleDrive/~ {self.cari_adi}_Bekleyenler.xlsx"
            
            # KDV sütununu kaldır
            bekleyenler_export = self.bekleyenler_filtered_data.copy()
            if "KDV(%)" in bekleyenler_export.columns:
                bekleyenler_export = bekleyenler_export.drop(columns=["KDV(%)"])

            # Kalem No formatını düzenle
            bekleyenler_export = self.format_kalem_no(bekleyenler_export)

            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                bekleyenler_export.to_excel(writer, sheet_name='Bekleyenler', index=False)
            
            self.status_label.setText(f"✅ Bekleyenler verileri dışa aktarıldı: {output_path}")
            QMessageBox.information(self, "Başarılı", f"Bekleyenler verileri başarıyla dışa aktarıldı:\n{output_path}")
            
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Bekleyenler dışa aktarma hatası: {str(e)}")
            self.status_label.setText(f"❌ Bekleyenler dışa aktarma hatası: {str(e)}")
    
    def export_arac_to_excel(self):
        """Planlanan araç verilerini Excel'e aktar"""
        try:
            if not self.cari_adi:
                QMessageBox.warning(self, "Hata", "Önce bir müşteri seçin!")
                return
            
            if self.arac_filtered_data.empty:
                QMessageBox.warning(self, "Hata", "Planlanan araç verisi bulunamadı!")
                return
            
            output_path = f"D:/GoogleDrive/~ {self.cari_adi}_Planlanan_Arac.xlsx"
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                self.arac_filtered_data.to_excel(writer, sheet_name='Planlanan_Arac', index=False)
            
            self.status_label.setText(f"✅ Planlanan araç verileri dışa aktarıldı: {output_path}")
            QMessageBox.information(self, "Başarılı", f"Planlanan araç verileri başarıyla dışa aktarıldı:\n{output_path}")
            
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Planlanan araç dışa aktarma hatası: {str(e)}")
            self.status_label.setText(f"❌ Planlanan araç dışa aktarma hatası: {str(e)}")
    
    def export_malzeme_to_excel(self):
        """Malzeme bazlı verileri Excel'e aktar"""
        try:
            if not self.cari_adi:
                QMessageBox.warning(self, "Hata", "Önce bir müşteri seçin!")
                return
            
            if self.sevkiyat_filtered_again.empty:
                QMessageBox.warning(self, "Hata", "Malzeme bazlı veri bulunamadı!")
                return
            
            output_path = f"D:/GoogleDrive/~ {self.cari_adi}_Malzeme_Bazli.xlsx"
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                self.sevkiyat_filtered_again.to_excel(writer, sheet_name='Malzeme_Bazli', index=False)
            
            self.status_label.setText(f"✅ Malzeme bazlı veriler dışa aktarıldı: {output_path}")
            QMessageBox.information(self, "Başarılı", f"Malzeme bazlı veriler başarıyla dışa aktarıldı:\n{output_path}")
            
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Malzeme bazlı dışa aktarma hatası: {str(e)}")
            self.status_label.setText(f"❌ Malzeme bazlı dışa aktarma hatası: {str(e)}")
    
    def run_mikro(self):
        """Mikro güncelleme işlemlerini başlat"""
        try:
            # Progress bar'ı göster ve butonları devre dışı bırak
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 100)
            self.status_label.setText("� Mikro güncelleme işlemi başlatılıyor...")
            self.set_buttons_enabled(False)
            self.mikro_calisiyor = True
            
            # Thread'i oluştur ve başlat
            self.update_thread = MikroUpdateThread()
            self.update_thread.status_update.connect(self.status_label.setText)
            self.update_thread.progress_update.connect(self.progress_bar.setValue)
            self.update_thread.finished_signal.connect(self.on_mikro_sequence_finished)
            self.update_thread.error_signal.connect(lambda msg: self.status_label.setText(f"⚠️ {msg}"))
            self.update_thread.start()
            
        except Exception as e:
            self.status_label.setText(f"❌ Başlatma hatası: {str(e)}")
            self.progress_bar.setVisible(False)
            self.set_buttons_enabled(True)
            self.mikro_calisiyor = False
    
    def on_mikro_sequence_finished(self):
        """Tüm EXE'ler tamamlandığında"""
        self.status_label.setText("✅ Tüm güncellemeler tamamlandı, veriler yenileniyor...")
        self.progress_bar.setValue(100)
        
        # Google Sheets'e verinin gitmesi için kısa bir bekleme
        QTimer.singleShot(2000, self.on_mikro_finished)
    
    def on_mikro_finished(self):
        """Mikro program bittikten sonra"""
        self.mikro_calisiyor = False
        self.status_label.setText("✅ Sevkiyat.exe tamamlandı, Google Sheets güncelleme bekleniyor...")
        
        # Google Sheets'e kaydedilmesi için ek bekleme (risk modülü ile aynı süre: 5 saniye)
        # Sonra verileri yenile
        QTimer.singleShot(5000, self.delayed_data_refresh)
    
    def delayed_data_refresh(self):
        """Gecikmeli veri yenileme"""
        self.status_label.setText("🔄 Google Sheets'ten güncel veriler alınıyor...")
        QApplication.processEvents()
        self.load_all_data()
        
        # Veri yenileme tamamlandıktan sonra progress bar'ı gizle ve butonları aktif et
        self.progress_bar.setVisible(False)
    
    def set_buttons_enabled(self, enabled: bool):
        """Butonları aktif/pasif yap"""
        self.refresh_button.setEnabled(enabled)
        self.mikro_button.setEnabled(enabled)
        self.whatsapp_randevu_button.setEnabled(enabled)
        self.whatsapp_bilgi_button.setEnabled(enabled)
        self.export_button.setEnabled(enabled)
        self.export_bekleyenler_button.setEnabled(enabled)
        self.export_arac_button.setEnabled(enabled)
        self.export_malzeme_button.setEnabled(enabled)
        self.mail_gonder_button.setEnabled(enabled)
        self.sevk_button.setEnabled(enabled)
        self.stok_analizi_button.setEnabled(enabled)