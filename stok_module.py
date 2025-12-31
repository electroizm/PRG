import sys
import os
import pandas as pd
import numpy as np

# Üst dizini Python path'e ekle (central_config için)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import pyodbc
import requests
import time
import re
import urllib.parse
import webbrowser
import logging
from datetime import datetime
from central_config import CentralConfigManager

def ayar_verilerini_al(response_content):
    """
    PRGsheet/Ayar sayfasından ayar verilerini çeker
    
    Args:
        response_content: Excel dosyasının BytesIO içeriği
        
    Returns:
        dict: Ayar verileri {'KDV': value, 'Ön Ödeme İskonto': value}
    """
    try:
        from io import BytesIO
        
        # Ayar sayfasını oku
        ayar_df = pd.read_excel(BytesIO(response_content), sheet_name="Ayar")
        
        ayar_dict = {}
        
        # Verileri satır satır oku ve ayar değerlerini bul
        for _, row in ayar_df.iterrows():
            if 'Ayar' in row and 'Değer' in row:
                key = str(row['Ayar']).strip()
                value = str(row['Değer']).strip()
                
                if key == 'KDV':
                    try:
                        # Virgülü noktaya çevir ve float'a dönüştür
                        value_clean = value.replace(',', '.')
                        ayar_dict['KDV'] = float(value_clean)
                    except (ValueError, TypeError) as e:
                        logging.warning(f"KDV değeri okunamadı ({value}): {e}")
                        ayar_dict['KDV'] = 1.10  # Varsayılan değer
                        
                elif key == 'Ön Ödeme İskonto':
                    try:
                        # Virgülü noktaya çevir ve float'a dönüştür
                        value_clean = value.replace(',', '.')
                        ayar_dict['Ön Ödeme İskonto'] = float(value_clean)
                    except (ValueError, TypeError) as e:
                        logging.warning(f"Ön Ödeme İskonto değeri okunamadı ({value}): {e}")
                        ayar_dict['Ön Ödeme İskonto'] = 0.90  # Varsayılan değer
                        
                elif key == 'Sepet_Marj':
                    try:
                        # Virgülü noktaya çevir ve float'a dönüştür
                        value_clean = value.replace(',', '.')
                        ayar_dict['Sepet_Marj'] = float(value_clean)
                    except (ValueError, TypeError) as e:
                        logging.warning(f"Sepet_Marj değeri okunamadı ({value}): {e}")
                        ayar_dict['Sepet_Marj'] = 1.35  # Varsayılan değer
                        
        
        # Eğer değerler bulunamadıysa varsayılan değerleri ata
        if 'KDV' not in ayar_dict:
            ayar_dict['KDV'] = 1.10
        if 'Ön Ödeme İskonto' not in ayar_dict:
            ayar_dict['Ön Ödeme İskonto'] = 0.90
        if 'Sepet_Marj' not in ayar_dict:
            ayar_dict['Sepet_Marj'] = 1.35
            
        return ayar_dict
        
    except Exception as e:
        logging.warning(f"Ayar sayfası okunamadı: {e}")
        # Varsayılan değerler döndür
        return {'KDV': 1.10, 'Ön Ödeme İskonto': 0.90, 'Sepet_Marj': 1.35}

class StokApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.original_df = pd.DataFrame()
        self.filtered_df = pd.DataFrame()
        self.sepet_marj = 1.35  # Varsayılan değer, load_data'da güncellenecek
        self.kar_marji_column_name = "1.35"  # Dinamik sütun adı
        self.mikro_calisiyor = False  # Mikro program çalışma durumu
        self.bekleyen_calisiyor = False  # Bekleyen program çalışma durumu
        self._data_loaded = False  # Lazy loading için flag
        self.setup_ui()
        self.show()

    def showEvent(self, event):
        """Widget ilk gösterildiğinde veri yükle (lazy loading)"""
        super().showEvent(event)
        if not self._data_loaded:
            self._data_loaded = True
            # UI render olduktan sonra veri yükle
            QTimer.singleShot(100, self.load_data)
    
    def setup_ui(self):
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)

        # Ana pencere arka planını beyaz yap
        self.setStyleSheet("""
            QMainWindow {
                background-color: #ffffff;
                color: #000000;
            }
        """)

        # Central widget arka planını beyaz yap
        self.central_widget.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                color: #000000;
            }
        """)

        # Arama ve Temizleme Alanı
        search_layout = QHBoxLayout()

        # Mikro Butonu
        self.micro_btn = QPushButton("Mikro")
        self.micro_btn.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 16px;
                border-radius: 5px;
                font-size:14px;
                font-weight: bold;
                min-width: 50px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
        """)
        self.micro_btn.setToolTip("Mikro programını çalıştırır")
        self.micro_btn.clicked.connect(self.run_mikro)
        search_layout.addWidget(self.micro_btn)

        # 3A -> 2A Butonu
        self.pasif_btn = QPushButton("3A -> 2A")
        self.pasif_btn.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 16px;
                border-radius: 5px;
                font-size:14px;
                font-weight: bold;
                min-width: 50px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
        """)
        self.pasif_btn.setToolTip("Kullanılmayan aktif stokları pasif yapar")
        self.pasif_btn.clicked.connect(self.pasif_yap)
        search_layout.addWidget(self.pasif_btn)

        # Bekleyen Butonu
        self.bekleyen_btn = QPushButton("Bekleyen")
        self.bekleyen_btn.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 16px;
                border-radius: 5px;
                font-size:14px;
                font-weight: bold;
                min-width: 50px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
        """)
        self.bekleyen_btn.setToolTip("BekleyenFast.exe programını çalıştırır")
        self.bekleyen_btn.clicked.connect(self.run_bekleyen)
        search_layout.addWidget(self.bekleyen_btn)

        # Malzeme Butonu
        self.malzeme_btn = QPushButton("Malzeme")
        self.malzeme_btn.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 16px;
                border-radius: 5px;
                font-size:14px;
                font-weight: bold;
                min-width: 50px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
        """)
        self.malzeme_btn.setToolTip("Bekleyen Stok kartı oluşturur")
        self.malzeme_btn.clicked.connect(self.create_malzeme_karti)
        search_layout.addWidget(self.malzeme_btn)

        # Arama Kutusu
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Malzeme Adı...")
        self.search_box.setStyleSheet("""
            font-size:16px; 
            padding:14px;
            border-radius: 5px;
            border: 1px solid #444;                          
            font-weight: bold;
        """)
        self.search_box.textChanged.connect(self.schedule_filter)
        search_layout.addWidget(self.search_box, 1)

        # Temizle Butonu
        self.clear_btn = QPushButton("Temizle")
        self.clear_btn.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                border-radius: 5px;
                padding: 8px 16px;
                font-size:14px;
                font-weight: bold;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
        """)
        self.clear_btn.setToolTip("Arama kutusunu ve satış verilerini temizler")
        self.clear_btn.clicked.connect(self.clear_all)
        search_layout.addWidget(self.clear_btn)

        # Liste Butonu
        self.list_btn = QPushButton("Liste")
        self.list_btn.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 16px;
                font-size:14px;
                font-weight: bold;
                border-radius: 5px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
        """)
        self.list_btn.setToolTip("Satış Listesi")
        self.list_btn.clicked.connect(self.filter_by_Sepet)
        search_layout.addWidget(self.list_btn)

        # Sepet Butonu
        self.Sepet_btn = QPushButton("Sepet")
        self.Sepet_btn.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 16px;
                font-size:14px;
                font-weight: bold;
                border-radius: 5px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
        """)
        self.Sepet_btn.setToolTip("Sepet verilerini kaydeder")
        self.Sepet_btn.clicked.connect(self.save_Sepet)
        search_layout.addWidget(self.Sepet_btn)

        self.team_btn = QPushButton("Ekip")
        self.team_btn.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 16px;
                font-size:14px;
                font-weight: bold;
                border-radius: 5px;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
        """)
        self.team_btn.setToolTip("Listeyi WhatsApp Satış Ekibi Grubuna gönderir")
        self.team_btn.clicked.connect(self.filter_and_send_to_whatsapp)
        search_layout.addWidget(self.team_btn)

        # Excel Butonu
        self.excel_btn = QPushButton("Excel")
        self.excel_btn.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 16px;
                border-radius: 5px;
                font-size:14px;
                font-weight: bold;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
        """)
        self.excel_btn.setToolTip("Listeyi Excel'e aktarır")
        self.excel_btn.clicked.connect(self.Stoklistesi)
        search_layout.addWidget(self.excel_btn)

        # Sipariş Butonu
        self.order_btn = QPushButton("Sipariş")
        self.order_btn.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 16px;
                font-size:14px;
                border-radius: 5px;
                font-weight: bold;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #a0a5a2;
            }
        """)
        self.order_btn.setToolTip("Siparişleri kaydeder")
        self.order_btn.clicked.connect(self.save_order)
        search_layout.addWidget(self.order_btn)

        self.main_layout.addLayout(search_layout)

        # Tablo
        self.table = QTableWidget()
        self.table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        
        # Ctrl+C kısayolu
        self.copy_shortcut = QShortcut(QKeySequence("Ctrl+C"), self.table)
        self.copy_shortcut.activated.connect(self.handle_ctrl_c)

        self.table.itemChanged.connect(self.handle_Sepet_edit)
        self.table.verticalHeader().setDefaultSectionSize(self.table.verticalHeader().defaultSectionSize() + 2)
        self.table.setStyleSheet("""
            QTableWidget {
                font-size: 15px;
                font-weight: bold;
            }
            QTableWidget::item:focus {
                outline: none;
                border: none;
            }
        """)
        self.table.setFocusPolicy(Qt.NoFocus)
        self.main_layout.addWidget(self.table)

        # Toplamlar Paneli
        self.totals_widget = QWidget()
        self.totals_layout = QGridLayout(self.totals_widget)

        # Hesaplama Etiketleri
        self.calculation_labels = {
            'Sepet': QLabel("Sepet: 0.00 ₺"),
            'EXC': QLabel("EXC : 0.00 ₺"),
            'Plan': QLabel("Plan : 0.00 ₺"),
            'Borç': QLabel("Borç : 0.00 ₺"),
            'Perakende': QLabel("Perakende : 0.00 ₺"),            
            'SUBE': QLabel("SUBE : 0.00 ₺"),
            'Bekleyen': QLabel("Bekleyen : 0.00 ₺"),
            'Fazla': QLabel("Fazla : 0.00 ₺"),
            'Liste': QLabel("Liste : 0.00 ₺"),
            'DEPO': QLabel("DEPO : 0.00 ₺"),
            'Ver': QLabel("Ver : 0.00 ₺"),
            'Toplam': QLabel("Toplam : 0.00 ₺")
        }

        # Etiketleri yerleştirme ve stil ayarlama
        row, col = 0, 0
        for key, label in self.calculation_labels.items():
            if key not in ['Perakende', 'Liste']:  # Sadece bu iki etiket hariç
                label.setStyleSheet("""
                    QLabel {
                        font-size:11pt; 
                        color:#f5f5f5;  /* Normalde sönük gri */
                        font-weight:bold;
                    }
                    QLabel:hover {
                        color:black;    /* Üzerine gelindiğinde siyah */
                    }
                """)
            else:
                label.setStyleSheet("""
                    QLabel {
                        font-size:11pt; 
                        color:black;    /* Her zaman siyah */
                        font-weight:bold;
                    }
                """)
            self.totals_layout.addWidget(label, row, col)
            col += 1
            if col > 3:
                col = 0
                row += 1

        self.main_layout.addWidget(self.totals_widget)

        # Progress Bar ve Status Label - Yan yana ve en altta (Sevkiyat modülü ile aynı)
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

        self.main_layout.addWidget(status_widget)

        # Performans Optimizasyonları
        self.filter_timer = QTimer()
        self.filter_timer.setSingleShot(True)
        self.filter_timer.timeout.connect(self.filter_data)

    def pasif_yap(self):
        try:
            # PRGsheet dosyasının Stok sayfasından veri oku
            import requests
            from io import BytesIO

            # Service Account ile PRGsheet'e erişim
            config_manager = CentralConfigManager()
            spreadsheet_id = config_manager.MASTER_SPREADSHEET_ID
            # Google Sheets URL'sini oluştur
            gsheets_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"
            
            # URL'den Excel dosyasını oku
            response = requests.get(gsheets_url, timeout=30)
            
            if response.status_code == 401:
                QMessageBox.warning(self, "Uyarı", "Google Sheets erişim hatası: Dosya özel veya izin gerekli")
                return
            elif response.status_code != 200:
                QMessageBox.warning(self, "Uyarı", f"HTTP Hatası: {response.status_code} - {response.reason}")
                return
            
            response.raise_for_status()
            
            # Stok sayfasını oku
            df_stok = pd.read_excel(BytesIO(response.content), sheet_name="Stok")
            
            # Pasif yapılmayacak özel durumları filtrele
            exclude_conditions = (
                df_stok['Malzeme Adı'].str.contains('MHZ %0|MHZ %10|MHZ %20', na=False, regex=True) |
                df_stok['Malzeme Kodu'].str.startswith('XX', na=False)
            )
            
            # Pasif yapılacak adayları belirle
            df_to_check = df_stok[~exclude_conditions].copy()
            
            # Sayısal sütunları float'a çevir ve NaN'ları 0 yap
            numeric_cols = ['Fazla', 'Borç', 'DEPO', 'Bekleyen', 'Plan', 'EXC', 'SUBE']
            for col in numeric_cols:
                if col in df_to_check.columns:
                    df_to_check[col] = pd.to_numeric(df_to_check[col], errors='coerce').fillna(0)
            
            # Stok hareketleri toplamı 0 olanları bul
            df_to_check['Toplam'] = (
                df_to_check['Fazla'] + 
                df_to_check['Borç'] + 
                df_to_check['DEPO'] + 
                df_to_check['Bekleyen'] + 
                df_to_check['Plan'] + 
                df_to_check['EXC'] + 
                df_to_check['SUBE']
            )
            
            # Toplamı tam olarak 0 olanları seç (küçük ondalıkları önlemek için round kullan)
            df_to_deactivate = df_to_check[np.round(df_to_check['Toplam'], 10) == 0]
            
            if df_to_deactivate.empty:
                QMessageBox.information(self, "Bilgi", "Pasif yapılacak stok bulunamadı.")
                return
            
            # Veritabanı bağlantısını oluştur
            conn = self.get_connection()
            cursor = conn.cursor()
            
            # Pasif yapma işlemi
            deactivated_count = 0
            deactivated_items = []

            
            for _, row in df_to_deactivate.iterrows():
                try:
                    malzeme_kodu = row['Malzeme Kodu']
                    malzeme_adi = row['Malzeme Adı']
                    
                    cursor.execute("UPDATE STOKLAR SET sto_pasif_fl = 1 WHERE sto_kod = ?", (malzeme_kodu,))
                    conn.commit()
                    deactivated_count += 1
                    deactivated_items.append(f"{malzeme_kodu} - {malzeme_adi}")
                        
                except Exception as e:
                    logging.error(f"Pasif yapma hatası: {malzeme_kodu} - {str(e)}")
                    continue
            
            conn.close()
            
            # Sonuç mesajını oluştur
            message = f"Pasif yapılan stok sayısı: {deactivated_count}"
            
            # Eğer pasif yapılan stoklar varsa detayları göster
            if deactivated_count > 0:
                message += "\n\nPasif yapılan stoklar:\n" + "\n".join(deactivated_items[:70])
                if deactivated_count > 10:
                    message += f"\n...ve {deactivated_count - 70} adet daha"
            
            QMessageBox.information(self, "Sonuç", message)
            
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Pasif yapma işlemi sırasında hata: {str(e)}")
            logging.error(f"Pasif yapma hatası: {str(e)}")

    def create_malzeme_karti(self):
            try:
                # PRGsheet dosyasının Bekleyenler sayfasından veri oku
                import requests
                from io import BytesIO

                # Service Account ile PRGsheet'e erişim
                config_manager = CentralConfigManager()
                spreadsheet_id = config_manager.MASTER_SPREADSHEET_ID
                # Google Sheets URL'sini oluştur
                gsheets_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"
                
                # URL'den Excel dosyasını oku
                response = requests.get(gsheets_url, timeout=30)
                
                if response.status_code == 401:
                    QMessageBox.warning(self, "Uyarı", "Google Sheets erişim hatası: Dosya özel veya izin gerekli")
                    return
                elif response.status_code != 200:
                    QMessageBox.warning(self, "Uyarı", f"HTTP Hatası: {response.status_code} - {response.reason}")
                    return
                
                response.raise_for_status()
                
                # Bekleyenler sayfasını oku
                df_bekleyen = pd.read_excel(BytesIO(response.content), sheet_name="Bekleyenler")
                
                # 1. Malzeme Kodu başlangıcı "3" sonu "-0" ile bitenleri filtrele
                df_bekleyen = df_bekleyen[
                    (df_bekleyen['Malzeme Kodu'].str.startswith('3', na=False)) & 
                    (df_bekleyen['Malzeme Kodu'].str.endswith('-0', na=False))
                ]

                # 2. Aynı Malzeme Koduna sahip tekrar eden satırları sil (son olanı korur)
                df_bekleyen = df_bekleyen.drop_duplicates(subset=['Malzeme Kodu'], keep='last')
                
                if df_bekleyen.empty:
                    QMessageBox.information(self, "Bilgi", "İşlenecek malzeme bulunamadı.")
                    return
                
                # Veritabanı bağlantısını oluştur
                conn = self.get_connection()
                cursor = conn.cursor()
                
                # Başarılı ve başarısız işlem sayıları
                success_count = 0
                fail_count = 0
                
                for index, row in df_bekleyen.iterrows():
                    try:
                        sto_kod = row['Malzeme Kodu']
                        sto_isim = row['Ürün Adı'] if 'Ürün Adı' in row else ''
                        #sto_yabanci_isim = row['Spec Adı'] if 'Spec Adı' in row else ''
                        sto_yabanci_isim = '' if pd.isna(row.get('Spec Adı', np.nan)) else str(row['Spec Adı'])
                        
                        # KDV oranını işle (10.0 -> 10)
                        kdv_oran = row['KDV(%)'] if 'KDV(%)' in row else 10.0
                        try:
                            kdv_oran = int(float(kdv_oran))
                        except:
                            kdv_oran = 10
                        
                        vergi_kodu = self.convert_kdv_to_vergi_kodu(kdv_oran)
                        
                        # Stok verilerini hazırla
                        stok_data = {
                            'sto_kod': sto_kod,
                            'sto_isim': sto_isim,
                            'sto_yabanci_isim': sto_yabanci_isim,
                            'sto_perakende_vergi': vergi_kodu,
                            'sto_toptan_vergi': vergi_kodu,
                            'sto_oto_barkod_kod_yapisi': '0'
                        }
                        
                        # Stok kartının varlığını kontrol et
                        cursor.execute("SELECT sto_pasif_fl FROM STOKLAR WHERE sto_kod = ?", (sto_kod,))
                        result = cursor.fetchone()
                        
                        if result:
                            # Stok kartı varsa ve pasifse aktif yap
                            if result[0] == 1:
                                cursor.execute("UPDATE STOKLAR SET sto_pasif_fl = 0 WHERE sto_kod = ?", (sto_kod,))
                                conn.commit()
                                success_count += 1
                        else:
                            # Stok kartı yoksa oluştur
                            if self.create_stok_karti(stok_data):
                                success_count += 1
                            else:
                                fail_count += 1
                                
                    except Exception as e:
                        logging.error(f"Malzeme işlenirken hata: {str(e)}")
                        fail_count += 1
                
                conn.close()
                
                QMessageBox.information(
                    self, 
                    "Sonuç", 
                    f"Bekleyen Stokların Aktarımı?\n\n"
                    f"Başarılı: {success_count}\n"
                    f"Başarısız: {fail_count}"
                )
                
            except Exception as e:
                QMessageBox.critical(self, "Hata", f"Malzeme kartı oluşturma hatası: {str(e)}")
                logging.error(f"Malzeme kartı oluşturma hatası: {str(e)}")

    def convert_kdv_to_vergi_kodu(self, kdv_oran):
        """KDV oranını Mikro vergi koduna dönüştür"""
        try:
            kdv_oran = float(kdv_oran)
            kdv_mapping = {
                1: 2,    # %1 KDV → Kod 2
                8: 3,     # %8 KDV → Kod 3
                10: 7,    # %10 KDV → Kod 7 (varsayılan)
                18: 4,    # %18 KDV → Kod 4
                20: 8     # %20 KDV → Kod 8
            }
            return kdv_mapping.get(kdv_oran, 7)  # Varsayılan %10 (kod 7)
        except:
            return 7  # Hata durumunda varsayılan %10 (kod 7)

    def create_stok_karti(self, stok_data):
        """Mikro DB'de stok kartı oluştur"""
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # STOKLAR için son RECno'yu al
            cursor.execute("SELECT MAX(sto_RECid_RECno) FROM STOKLAR")
            last_sto_recid = cursor.fetchone()[0] or 36177
            new_sto_recid = last_sto_recid + 1

            # BARKOD_TANIMLARI için son RECno'yu al
            cursor.execute("SELECT MAX(bar_RECid_RECno) FROM BARKOD_TANIMLARI")
            last_bar_recid = cursor.fetchone()[0] or 0
            new_bar_recid = last_bar_recid + 1

            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S.000")
            stok_kod = str(stok_data['sto_kod'])
            barkod = str(stok_data.get('sto_oto_barkod_kod_yapisi', '0')).strip()

            try:
                # 1. Adım: STOKLAR tablosuna ekleme
                sql_stok = """
                INSERT INTO [dbo].[STOKLAR] (
                    [sto_RECid_DBCno], [sto_RECid_RECno], [sto_SpecRECno], 
                    [sto_iptal], [sto_fileid], [sto_hidden], [sto_kilitli], 
                    [sto_degisti], [sto_checksum], [sto_create_user], 
                    [sto_create_date], [sto_lastup_user], [sto_lastup_date],
                    [sto_special1], [sto_special2], [sto_special3], [sto_kod],
                    [sto_isim], [sto_yabanci_isim], [sto_kisa_ismi], [sto_sat_cari_kod],
                    [sto_cins], [sto_doviz_cinsi], [sto_detay_takip], [sto_birim1_ad],
                    [sto_birim1_katsayi], [sto_birim1_agirlik], [sto_birim1_en],
                    [sto_birim1_boy], [sto_birim1_yukseklik], [sto_birim1_dara],
                    [sto_perakende_vergi], [sto_toptan_vergi], [sto_oto_barkod_kod_yapisi]
                ) VALUES (
                    0, ?, 0, 
                    0, 13, 0, 0, 
                    0, 0, 1, 
                    ?, 1, ?,
                    '', '', '', ?,
                    ?, ?, '', '',
                    10, 0, 0, 'Adet',
                    1, 0, 0, 0, 0, 0,
                    ?, ?, ?
                )
                """
                cursor.execute(sql_stok, (
                    new_sto_recid,
                    current_time, current_time,
                    stok_kod,
                    str(stok_data.get('sto_isim', ''))[:50],
                    str(stok_data.get('sto_yabanci_isim', ''))[:50],
                    int(stok_data.get('sto_perakende_vergi', 0)),
                    int(stok_data.get('sto_toptan_vergi', 0)),
                    barkod
                ))

                # 2. Adım: BARKOD_TANIMLARI tablosuna ekleme
                sql_barkod = """
                INSERT INTO [dbo].[BARKOD_TANIMLARI] (
                    [bar_RECid_DBCno], [bar_RECid_RECno], [bar_SpecRECno],
                    [bar_iptal], [bar_fileid], [bar_hidden], [bar_kilitli],
                    [bar_degisti], [bar_checksum], [bar_create_user],
                    [bar_create_date], [bar_lastup_user], [bar_lastup_date],
                    [bar_special1], [bar_special2], [bar_special3],
                    [bar_kodu], [bar_stokkodu], [bar_partikodu],
                    [bar_lotno], [bar_serino_veya_bagkodu], [bar_barkodtipi],
                    [bar_icerigi], [bar_birimpntr], [bar_master],
                    [bar_bedenpntr], [bar_renkpntr], [bar_baglantitipi],
                    [bar_harrecid_dbcno], [bar_harrecid_recno]
                ) VALUES (
                    0, ?, 0,
                    0, 15, 0, 0,
                    0, 0, 1,
                    ?, 1, ?,
                    '', '', '',
                    ?, ?, '',
                    0, ?, 0,
                    0, 0, 0,
                    0, 0, 0,
                    0, 0
                )
                """
                cursor.execute(sql_barkod, (
                    new_bar_recid,
                    current_time, current_time,
                    stok_kod, stok_kod, barkod
                ))

                conn.commit()
                return True

            except pyodbc.IntegrityError:
                conn.rollback()
                return False
            except Exception as e:
                conn.rollback()
                logging.error(f"Stok kartı oluşturulamadı: {str(e)}")
                return False
            finally:
                conn.close()

        except Exception as e:
            logging.error(f"Veritabanı bağlantı hatası: {str(e)}")
            return False

    def get_connection(self):

        # PRGsheet/Ayar sayfasından SQL bağlantı bilgilerini yükle

        server = os.getenv('SQL_SERVER')
        database = os.getenv('SQL_DATABASE')
        username = os.getenv('SQL_USERNAME')
        password = os.getenv('SQL_PASSWORD')
        
        if not all([server, database, username, password]):
            raise Exception("PRGsheet/Ayar sayfasında SQL bağlantı bilgileri eksik")

        # Bağlantı dizesini oluşturun
        connection_string = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'

        # Veritabanına bağlanın
        return pyodbc.connect(connection_string)

    def load_data(self):
        try:
            # Progress bar ve status label'ı göster
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(0)
            self.status_label.setVisible(True)
            self.status_label.setText("📊 Stok sayfasından veriler yükleniyor...")
            QApplication.processEvents()

            # PRGsheet dosyasından hem Stok hem de Fiyat sayfalarını yükle

            from io import BytesIO

            # PRGsheet/Ayar sayfasından SPREADSHEET_ID'yi yükle

            spreadsheet_id = CentralConfigManager().MASTER_SPREADSHEET_ID

            # Google Sheets URL'sini oluştur
            gsheets_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"
            
            if hasattr(self, 'status_label'):
                self.status_label.setText("🔗 Google Sheets'e bağlanıyor...")
                QApplication.processEvents()

            # URL'den Excel dosyasını oku
            response = requests.get(gsheets_url, timeout=30)

            # Google Sheets bağlantısı başarılı
            self.progress_bar.setValue(10)
            self.status_label.setText("✅ Google Sheets'e bağlantı başarılı")
            QApplication.processEvents()

            if response.status_code == 401:
                self.progress_bar.setVisible(False)
                if hasattr(self, 'status_label'):
                    self.status_label.setText("❌ Google Sheets erişim hatası: Dosya özel veya izin gerekli")
                logging.error("Google Sheets erişim hatası: Dosya özel veya izin gerekli")
                return
            elif response.status_code != 200:
                self.progress_bar.setVisible(False)
                if hasattr(self, 'status_label'):
                    self.status_label.setText(f"❌ HTTP Hatası: {response.status_code} - {response.reason}")
                logging.error(f"HTTP Hatası: {response.status_code} - {response.reason}")
                return
            
            response.raise_for_status()
            
            # 1. Adım: Stok sayfasını oku
            if hasattr(self, 'status_label'):
                self.status_label.setText("📋 Stok sayfası işleniyor...")
                QApplication.processEvents()
            self.original_df = pd.read_excel(BytesIO(response.content), sheet_name="Stok")

            # Kritik kontrol: Stok sayfası boş mu?
            if self.original_df.empty:
                error_msg = "Stok sayfası boş! Veriler yüklenemedi."
                logging.error(error_msg)
                self.progress_bar.setVisible(False)
                if hasattr(self, 'status_label'):
                    self.status_label.setText(f"❌ {error_msg}")
                QMessageBox.warning(self, "Uyarı", error_msg)
                return

            # Kritik sütunların varlığını kontrol et
            required_columns = ['SAP Kodu', 'Malzeme Adı', 'Malzeme Kodu']
            missing_columns = [col for col in required_columns if col not in self.original_df.columns]
            if missing_columns:
                error_msg = f"Stok sayfasında gerekli sütunlar eksik: {', '.join(missing_columns)}"
                logging.error(error_msg)
                self.progress_bar.setVisible(False)
                if hasattr(self, 'status_label'):
                    self.status_label.setText(f"❌ {error_msg}")
                QMessageBox.warning(self, "Uyarı", error_msg)
                return

            # Sepet_Timestamp sütunu ekle (değiştirilme zamanı için)
            if 'Sepet_Timestamp' not in self.original_df.columns:
                self.original_df['Sepet_Timestamp'] = pd.NaT

            # Stok sayfası başarıyla yüklendi
            self.progress_bar.setValue(30)
            self.status_label.setText("✅ Stok sayfası başarıyla yüklendi")
            QApplication.processEvents()

            # Tüm sayısal sütunları işle
            int_columns = ['Ver', 'Sepet', 'DEPO', 'Fazla', 'Borç', 'Bekleyen', 'Plan', 'EXC', 'SUBE', 'ID1', 'ID2', 'Miktar', 'TOPTAN', 'PERAKENDE', 'LISTE', self.kar_marji_column_name, 'INDIRIM']
            for col in int_columns:
                if col in self.original_df.columns:
                    self.original_df[col] = pd.to_numeric(self.original_df[col], errors='coerce').fillna(0).astype(int)

            # 2. Adım: Fiyat sayfasını yükle
            try:
                if hasattr(self, 'status_label'):
                    self.status_label.setText("💰 Fiyat sayfası işleniyor...")
                    QApplication.processEvents()
                fiyat_df = pd.read_excel(BytesIO(response.content), sheet_name="Fiyat")

                # Sayısal sütunları işle
                for col in ['TOPTAN', 'PERAKENDE', 'LISTE', self.kar_marji_column_name, 'INDIRIM']:
                    if col in fiyat_df.columns:
                        fiyat_df[col] = pd.to_numeric(fiyat_df[col], errors='coerce').fillna(0).astype(int)

                # SAP Kodu sütununu string'e çevir (merge uyumluluğu için)
                fiyat_df['SAP Kodu'] = fiyat_df['SAP Kodu'].astype(str)

                # Malzeme Kodu oluştur
                fiyat_df['Malzeme Kodu'] = fiyat_df['SAP Kodu'] + '-0'

                # Eksik sütunları varsayılan 0 değeriyle ekle
                for col in ['Ver', 'Sepet', 'Fazla', 'Borç', 'DEPO', 'Bekleyen', 'Plan', 'EXC', 'SUBE', 'ID1', 'ID2', 'Miktar']:
                    if col not in fiyat_df.columns:
                        fiyat_df[col] = 0

                # Stok sayfasındaki ürünlerin fiyat bilgilerini güncelle
                if 'SAP Kodu' in self.original_df.columns:
                    # SAP Kodu sütununu string'e çevir (merge uyumluluğu için)
                    self.original_df['SAP Kodu'] = self.original_df['SAP Kodu'].astype(str)

                    # Fiyat sütunlarını başlangıçta sadece yoksa 0 olarak ekle
                    for col in ['PERAKENDE', 'LISTE', self.kar_marji_column_name, 'INDIRIM']:
                        if col not in self.original_df.columns:
                            self.original_df[col] = 0
                                        
                    # Fiyat sayfasında mevcut olan sütunları belirle
                    available_price_cols = ['SAP Kodu', 'TOPTAN', 'PERAKENDE', 'LISTE']
                    for col in [self.kar_marji_column_name, 'INDIRIM', 'DOSYA']:
                        if col in fiyat_df.columns:
                            available_price_cols.append(col)
                    
                    # Sonra fiyat_df'den gelen değerlerle güncelle
                    price_updates = fiyat_df[available_price_cols].drop_duplicates(subset=['SAP Kodu'])
                    
                    self.original_df = self.original_df.merge(
                        price_updates,
                        on='SAP Kodu',
                        how='left',
                        suffixes=('', '_new')
                    )
                    # Yeni değerlerle güncelleme
                    for col in ['TOPTAN', 'PERAKENDE', 'LISTE', self.kar_marji_column_name, 'INDIRIM', 'DOSYA']:
                        if f'{col}_new' in self.original_df.columns:
                            # Sadece eşleşenlerin değerlerini güncelle, diğerleri 0 olarak kalacak (Dosya için boş string)
                            if col == 'DOSYA':
                                self.original_df[col] = self.original_df[f'{col}_new'].fillna('')
                            else:
                                self.original_df[col] = self.original_df[f'{col}_new'].fillna(0)
                            self.original_df.drop(f'{col}_new', axis=1, inplace=True)
                
                # Stok sayfasında olmayan kayıtları ekle (SAP Kodu bazında kontrol)
                if 'SAP Kodu' in self.original_df.columns:
                    existing_sap_codes = set(self.original_df['SAP Kodu'])
                    new_items = fiyat_df[~fiyat_df['SAP Kodu'].isin(existing_sap_codes)].copy()

                    if not new_items.empty:
                        # Tüm eksik sütunları önce ekle (NaN'larla)
                        for col in self.original_df.columns:
                            if col not in new_items.columns:
                                new_items[col] = 0 if col in int_columns else ''

                        # Sütun sırasını aynı yap
                        new_items = new_items[self.original_df.columns]

                        # Concat işlemi
                        self.original_df = pd.concat([self.original_df, new_items], ignore_index=True)

                # Fiyat sayfası başarıyla yüklendi
                self.progress_bar.setValue(50)
                self.status_label.setText("✅ Fiyat sayfası başarıyla yüklendi")
                QApplication.processEvents()

            except Exception as e:
                # Fiyat sayfası opsiyonel - hata loglansın ama uygulama durmasın
                logging.warning(f"Fiyat sayfası yüklenirken hata oluştu: {str(e)}")
                if hasattr(self, 'status_label'):
                    self.status_label.setText(f"⚠️ Fiyat sayfası yüklenemedi: {str(e)}")

            # 3. Adım: Ayar sayfasından değerleri yükle
            if hasattr(self, 'status_label'):
                self.status_label.setText("⚙️ Ayar sayfası işleniyor...")
                QApplication.processEvents()
            ayar = ayar_verilerini_al(response.content)
            
            kdv = ayar.get('KDV', 1.10)
            on_odeme_iskonto = ayar.get('Ön Ödeme İskonto', 0.90)
            self.sepet_marj = ayar.get('Sepet_Marj', 1.35)  # Instance variable olarak kaydet
            self.kar_marji_column_name = str(self.sepet_marj)  # Sütun adını güncelle

            # Ayar sayfası başarıyla yüklendi
            self.progress_bar.setValue(70)
            self.status_label.setText("✅ Ayar sayfası başarıyla yüklendi")
            QApplication.processEvents()

            # 4. Adım: Hesaplamalar
            # ID2 = TOPTAN * KDV * Ön Ödeme İskonto hesaplaması
            if 'TOPTAN' in self.original_df.columns:
                self.original_df['ID2'] = self.original_df['TOPTAN'].apply(
                    lambda x: int(x * kdv * on_odeme_iskonto) if pd.notna(x) and x != 0 else 0
                )
            else:
                self.original_df['ID2'] = 0
            
            # Kar marjı hesaplaması: (1-(ID2/PERAKENDE))*100 (PERAKENDE=0 ise 0)
            if 'PERAKENDE' in self.original_df.columns and 'ID2' in self.original_df.columns:
                self.original_df[self.kar_marji_column_name] = self.original_df.apply(
                    lambda row: int((1 - (row['ID2'] / row['PERAKENDE'])) * 100) 
                    if row['PERAKENDE'] != 0 
                    else 0, 
                    axis=1
                )
            else:
                self.original_df[self.kar_marji_column_name] = 0
            
            # INDIRIM hesaplaması: (1-(PERAKENDE/LISTE))*100
            if 'PERAKENDE' in self.original_df.columns and 'LISTE' in self.original_df.columns:
                self.original_df['INDIRIM'] = self.original_df.apply(
                    lambda row: int((1 - (row['PERAKENDE'] / row['LISTE'])) * 100) 
                    if row['LISTE'] != 0 
                    else 0, 
                    axis=1
                )
            else:
                self.original_df['INDIRIM'] = 0

            # Hesaplamalar tamamlandı
            self.progress_bar.setValue(80)
            self.status_label.setText("✅ Hesaplamalar tamamlandı")
            QApplication.processEvents()

            # 5. Adım: Fiyat_Mikro sayfasından ID1 güncelleme (opsiyonel)
            try:
                if hasattr(self, 'status_label'):
                    self.status_label.setText("🔄 Fiyat_Mikro sayfası işleniyor...")
                    QApplication.processEvents()
                fiyat_mikro_df = pd.read_excel(BytesIO(response.content), sheet_name="Fiyat_Mikro")

                # Sütun isimlerini temizle (baştaki/sondaki boşlukları kaldır)
                fiyat_mikro_df.columns = fiyat_mikro_df.columns.str.strip()

                # SAP Kodu sütununu kontrol et - farklı isimlendirmeleri de kontrol et
                sap_column = None
                for col in fiyat_mikro_df.columns:
                    if col.upper().replace(' ', '').replace('_', '') == 'SAPKODU':
                        sap_column = col
                        break

                if sap_column and 'SAP Kodu' in self.original_df.columns:
                    # Sütun adını standartlaştır
                    if sap_column != 'SAP Kodu':
                        fiyat_mikro_df.rename(columns={sap_column: 'SAP Kodu'}, inplace=True)

                    fiyat_mikro_df['SAP Kodu'] = fiyat_mikro_df['SAP Kodu'].astype(str)
                    self.original_df['SAP Kodu'] = self.original_df['SAP Kodu'].astype(str)

                    # TOPTAN sütununu sayısal değere çevir
                    if 'TOPTAN' in fiyat_mikro_df.columns:
                        fiyat_mikro_df['TOPTAN'] = pd.to_numeric(fiyat_mikro_df['TOPTAN'], errors='coerce').fillna(0).astype(int)

                        # SAP Kodu eşleşmesi olan satırlar için ID1'i TOPTAN değeri ile güncelle
                        for idx, row in self.original_df.iterrows():
                            sap_kodu = row['SAP Kodu']
                            matching_row = fiyat_mikro_df[fiyat_mikro_df['SAP Kodu'] == sap_kodu]
                            if not matching_row.empty:
                                toptan_value = matching_row.iloc[0]['TOPTAN']
                                self.original_df.at[idx, 'ID1'] = toptan_value
                    else:
                        logging.warning("Fiyat_Mikro sayfasında 'TOPTAN' sütunu bulunamadı")
                else:
                    logging.warning(f"Fiyat_Mikro sayfasında SAP Kodu sütunu bulunamadı. Mevcut sütunlar: {list(fiyat_mikro_df.columns)}")

            except Exception as e:
                logging.warning(f"Fiyat_Mikro sayfası işlenirken hata: {str(e)}")

            # Sütun sıralamasını ayarla
            column_order = [
                'SAP Kodu', 'Ver', 'Sepet', 'Sepet_Timestamp', 'Malzeme Adı', 'DEPO', 'Fazla', 'Borç',
                'Bekleyen', 'Plan', 'EXC', 'SUBE', 'ID1', 'ID2', self.kar_marji_column_name, 'INDIRIM',
                'Malzeme Kodu', 'Miktar', 'PERAKENDE', 'LISTE', 'DOSYA'
            ]
            
            # Mevcut sütunları kontrol et ve sıralama yap
            available_columns = [col for col in column_order if col in self.original_df.columns]
            # Sıralamada olmayan sütunları sona ekle (TOPTAN hariç)
            remaining_columns = [col for col in self.original_df.columns if col not in column_order and col != 'TOPTAN']
            final_columns = available_columns + remaining_columns
            
            self.original_df = self.original_df[final_columns]

            # Veri işleme tamamlandı
            self.progress_bar.setValue(90)
            self.status_label.setText("✅ Veri işleme tamamlandı")
            QApplication.processEvents()

            if hasattr(self, 'status_label'):
                self.status_label.setText("🔄 Tablo güncelleniyor...")
                QApplication.processEvents()

            # Final güvenlik kontrolü
            if self.original_df.empty:
                error_msg = "Veri yükleme tamamlandı ancak hiç kayıt bulunamadı!"
                logging.error(error_msg)
                self.progress_bar.setVisible(False)
                if hasattr(self, 'status_label'):
                    self.status_label.setText(f"❌ {error_msg}")
                QMessageBox.warning(self, "Uyarı", error_msg)
                return

            # Kritik sütunların son kontrolü
            final_required = ['SAP Kodu', 'Malzeme Adı', 'Malzeme Kodu']
            final_missing = [col for col in final_required if col not in self.original_df.columns]
            if final_missing:
                error_msg = f"Kritik sütunlar eksik: {', '.join(final_missing)}"
                logging.error(error_msg)
                self.progress_bar.setVisible(False)
                if hasattr(self, 'status_label'):
                    self.status_label.setText(f"❌ {error_msg}")
                QMessageBox.warning(self, "Uyarı", error_msg)
                return

            self.filtered_df = self.original_df.copy()
            self.update_table()
            self.update_totals()

            # Tüm işlemler tamamlandı
            self.progress_bar.setValue(100)
            QApplication.processEvents()

            # Progress bar'ı gizle
            QTimer.singleShot(1000, lambda: self.progress_bar.setVisible(False))

            if hasattr(self, 'status_label'):
                self.status_label.setText(f"✅ {len(self.original_df)} kayıt başarıyla yüklendi")
                
        except Exception as e:
            logging.error(f"Hata: {str(e)}")
            # Hata durumunda progress bar'ı gizle
            self.progress_bar.setVisible(False)
            if hasattr(self, 'status_label'):
                self.status_label.setText(f"❌ Veri yükleme hatası: {str(e)}")
            QMessageBox.critical(self, "Hata", f"Veri yükleme hatası: {str(e)}")

    def filter_by_Sepet(self):
        try:
            if 'Sepet' in self.original_df.columns:
                # Sepet > 0 olanları filtrele
                self.filtered_df = self.original_df[self.original_df['Sepet'] > 0].copy()

                # Sepet_Timestamp'e göre sırala (Eski → Yeni)
                if 'Sepet_Timestamp' in self.filtered_df.columns:
                    self.filtered_df = self.filtered_df.sort_values(
                        by='Sepet_Timestamp',
                        ascending=True,  # Eski → Yeni
                        na_position='last'  # Timestamp olmayanlar en sonda
                    )

                self.update_table()
                self.update_totals()
            else:
                QMessageBox.warning(self, "Uyarı", "Sepet sütunu bulunamadı!")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Filtreleme hatası: {str(e)}")
            logging.error(f"Filtreleme hatası: {str(e)}")

    def clear_all(self):
        self.search_box.clear()
        if 'Sepet' in self.original_df.columns:
            self.original_df['Sepet'] = 0
            self.filtered_df['Sepet'] = 0
        # Sepet temizlendiğinde timestamp'leri de sil
        if 'Sepet_Timestamp' in self.original_df.columns:
            self.original_df['Sepet_Timestamp'] = pd.NaT
            self.filtered_df['Sepet_Timestamp'] = pd.NaT
        self.filtered_df = self.original_df.copy()
        self.update_table()
        self.update_totals()

    def schedule_filter(self):
        self.filter_timer.stop()
        self.filter_timer.start(200)

    def filter_data(self):
        try:
            search_text = self.search_box.text().strip().lower()

            if not search_text:
                self.filtered_df = self.original_df.copy()
            else:
                parts = [re.escape(part) for part in search_text.split() if part]
                pattern = r'(?=.*?{})'.format(')(?=.*?'.join(parts))
                mask = self.original_df['Malzeme Adı'].str.lower().str.contains(pattern, regex=True)
                self.filtered_df = self.original_df[mask].copy()

            if 'Ver' in self.filtered_df.columns:
                self.filtered_df['Ver'] = self.filtered_df['Ver'].fillna('')
                self.filtered_df = self.filtered_df.sort_values(
                    by=['Ver', 'Malzeme Adı'],
                    ascending=[False, True],
                    na_position='last'
                )

            self.update_table()
            self.update_totals()

        except Exception as e:
            logging.error(f"Filtreleme hatası: {str(e)}")

    def update_table(self):
        self.table.blockSignals(True)
        self.table.clearContents()

        # Sütun sıralaması: SAP Kodu, Ver, Sepet, Malzeme Adı, DEPO, Fazla, Borç, Bekleyen, Plan, EXC, SUBE, ID1, ID2, ###, INDIRIM, Malzeme Kodu, Miktar, PERAKENDE, LISTE, Dosya
        column_order = [
            'SAP Kodu', 'Ver', 'Sepet', 'Malzeme Adı', 'DEPO', 'Fazla', 'Borç','Bekleyen', 'Plan', 'EXC', 'SUBE','Miktar','ID1', 'ID2', 'PERAKENDE', 'LISTE', self.kar_marji_column_name, 'INDIRIM', 
'Malzeme Kodu', 'DOSYA'
        ]
        
        # Set operations ile performans odaklı sütun düzenleme
        # TOPTAN, ###, Sepet_Timestamp sütunlarını gizle
        available_cols = set(self.filtered_df.columns) - {'TOPTAN', '###', 'Sepet_Timestamp'}
        ordered_part = [col for col in column_order if col in available_cols]
        extra_part = list(available_cols - set(column_order))
        columns_to_show = ordered_part + extra_part
        
        self.filtered_df_display = self.filtered_df[columns_to_show]

        rows, cols = self.filtered_df_display.shape
        self.table.setRowCount(rows)
        self.table.setColumnCount(cols)
        self.table.setHorizontalHeaderLabels(self.filtered_df_display.columns)
        
        # Tooltip'leri ekle
        for i, col_name in enumerate(self.filtered_df_display.columns):
            header_item = QTableWidgetItem(col_name)
            if col_name == "ID1":
                header_item.setToolTip("ProSAP Toptan Tutar")
            elif col_name == "ID2":
                header_item.setToolTip("Excel Dosyaları Toptan Tutar")
            elif col_name == self.kar_marji_column_name:
                header_item.setToolTip(f"Kar marjı {self.sepet_marj}")
            elif col_name == "INDIRIM":
                header_item.setToolTip("Perakende yapılan % indirim")
            self.table.setHorizontalHeaderItem(i, header_item)

        editable_columns = ['Sepet']

        for i in range(rows):
            for j in range(cols):
                col_name = self.filtered_df_display.columns[j]
                value = self.filtered_df_display.iat[i, j]
                
                if pd.isna(value) or value is None:
                    display_value = ""
                else:
                    display_value = str(int(value)) if isinstance(value, (int, float)) and not pd.isna(value) else str(value)
                
                item = QTableWidgetItem(display_value)

                if col_name in editable_columns:
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsEditable)
                else:
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)

                if col_name == 'Fazla' and not pd.isna(value) and value > 0:
                    item.setBackground(QColor(144, 238, 144))
                    malzeme_adi_item = self.table.item(i, self.filtered_df_display.columns.get_loc('Malzeme Adı'))
                    if malzeme_adi_item:
                        malzeme_adi_item.setBackground(QColor(144, 238, 144))

                self.table.setItem(i, j, item)

        # Stok hareketleri olan satırların belirli sütunlarını açık gri renkle renklendirme
        stok_hareket_columns = ['DEPO', 'Fazla', 'Borç', 'Bekleyen', 'Plan', 'EXC', 'SUBE']
        renklendirilecek_columns = ['ID1', 'ID2', self.kar_marji_column_name, 'INDIRIM', 'PERAKENDE', 'LISTE']
        
        for i in range(rows):
            # Bu satırda stok hareketi var mı kontrol et
            stok_hareketi_var = False
            for stok_col in stok_hareket_columns:
                if stok_col in self.filtered_df_display.columns:
                    col_index = self.filtered_df_display.columns.get_loc(stok_col)
                    value = self.filtered_df_display.iat[i, col_index]
                    if not pd.isna(value) and value > 0:
                        stok_hareketi_var = True
                        break
            
            # Eğer stok hareketi varsa belirli sütunları açık gri renkle boyar
            if stok_hareketi_var:
                for renkli_col in renklendirilecek_columns:
                    if renkli_col in self.filtered_df_display.columns:
                        col_index = self.filtered_df_display.columns.get_loc(renkli_col)
                        item = self.table.item(i, col_index)
                        if item:
                            item.setBackground(QColor(236, 236, 231))  # Açık gri

        self.table.resizeColumnsToContents()
        self.table.blockSignals(False)

    def handle_Sepet_edit(self, item):
        col_name = self.filtered_df_display.columns[item.column()]
        if col_name == 'Sepet':
            try:
                new_value = float(item.text())
                original_index = self.filtered_df.index[item.row()]

                # Sepet değerini güncelle
                self.original_df.at[original_index, 'Sepet'] = new_value
                self.filtered_df.at[original_index, 'Sepet'] = new_value

                # Timestamp güncelle
                if new_value > 0:
                    # Sepet > 0 ise timestamp kaydet
                    current_time = datetime.now()
                    self.original_df.at[original_index, 'Sepet_Timestamp'] = current_time
                    self.filtered_df.at[original_index, 'Sepet_Timestamp'] = current_time
                else:
                    # Sepet = 0 ise timestamp'i sil
                    self.original_df.at[original_index, 'Sepet_Timestamp'] = pd.NaT
                    self.filtered_df.at[original_index, 'Sepet_Timestamp'] = pd.NaT

                self.update_totals()
            except ValueError:
                # Hatalı giriş - eski değere geri dön
                original_value = self.filtered_df.iat[item.row(), self.filtered_df.columns.get_loc('Sepet')]
                item.setText(str(int(original_value)) if original_value else '0')

    def update_totals(self):
        try:
            # Gerekli sütunların varlığını kontrol et
            if self.filtered_df.empty:
                # DataFrame boşsa tüm değerleri 0 yap
                for key in self.calculation_labels.keys():
                    if key == 'Toplam':
                        self.calculation_labels[key].setText(f"Toplam: 0 ₺")
                    elif key == 'Sepet':
                        self.calculation_labels[key].setText(f"Sepet: 0 ₺")
                    else:
                        self.calculation_labels[key].setText(f"{key} : 0 ₺")
                return
            
            # Gerekli sütunları kontrol et ve yoksa ekle
            required_columns = ['Sepet', 'Ver', 'Fazla', 'Borç', 'DEPO', 'Bekleyen', 'Plan', 'EXC', 'SUBE', 'ID1', 'ID2', 'PERAKENDE', 'LISTE']
            for col in required_columns:
                if col not in self.filtered_df.columns:
                    self.filtered_df[col] = 0
                if col not in self.original_df.columns:
                    self.original_df[col] = 0
            # PERAKENDE sütunundaki boş ve geçersiz değerleri 0 ile doldur
            if 'PERAKENDE' in self.filtered_df.columns:
                self.filtered_df['PERAKENDE'] = pd.to_numeric(self.filtered_df['PERAKENDE'], errors='coerce').fillna(0)
                self.filtered_df['LISTE'] = pd.to_numeric(self.filtered_df['LISTE'], errors='coerce').fillna(0)
                
                # Sepet > 0 olan satırları filtrele
                filtered = self.filtered_df[self.filtered_df['Sepet'] > 0]
                
                # Eğer Sepet > 0 olan satırlarda PERAKENDE 0 ise, toplamı direkt 0 yap
                if not filtered.empty and (filtered['PERAKENDE'] == 0).any():
                    perakende_total = 0
                    liste_total = 0
                else:
                    # Tüm Sepet > 0 olan satırlarda PERAKENDE sıfır değilse normal hesaplama yap
                    perakende_total = (filtered['Sepet'] * filtered['PERAKENDE']).sum()
                    liste_total = (filtered['Sepet'] * filtered['LISTE']).sum()
            else:
                perakende_total = 0
                liste_total = 0

            # ID1 sütununu kullan (eski ID yerine)
            id_col = 'ID1' if 'ID1' in self.filtered_df.columns else 'ID2'
            
            # Hesaplamalar
            calculations = {
                'Sepet': (self.filtered_df['Sepet'] * self.filtered_df[id_col] * self.sepet_marj).sum() if id_col in self.filtered_df.columns else 0,
                'Perakende': perakende_total,  # Yukarıda hesaplanan değeri kullan
                'Liste': liste_total,  # Yukarıda hesaplanan değeri kullan
                'Ver': (self.filtered_df['Ver'] * self.filtered_df[id_col]).sum() if id_col in self.filtered_df.columns else 0,
                'Fazla': (self.filtered_df['Fazla'] * self.filtered_df[id_col]).sum() if id_col in self.filtered_df.columns else 0,
                'EXC': (self.filtered_df['EXC'] * self.filtered_df[id_col]).sum() if id_col in self.filtered_df.columns else 0,
                'SUBE': (self.filtered_df['SUBE'] * self.filtered_df[id_col]).sum() if id_col in self.filtered_df.columns else 0,
                'Borç': (self.filtered_df['Borç'] * self.filtered_df[id_col]).sum() if id_col in self.filtered_df.columns else 0,
                'DEPO': (self.filtered_df['DEPO'] * self.filtered_df[id_col]).sum() if id_col in self.filtered_df.columns else 0,
                'Bekleyen': (self.filtered_df['Bekleyen'] * self.filtered_df[id_col]).sum() if id_col in self.filtered_df.columns else 0,
                'Plan': (self.filtered_df['Plan'] * self.filtered_df[id_col]).sum() if id_col in self.filtered_df.columns else 0
            }
            # Toplam hesaplama
            calculations['Toplam'] = (
                calculations['EXC'] + calculations['SUBE'] + calculations['DEPO'] +
                calculations['Bekleyen'] + calculations['Plan'] -
                calculations['Borç'] - calculations['Ver']
            )

            # Etiketleri güncelle
            for key, value in calculations.items():
                display_text = f"{value:,.0f} ₺"
                if key == 'Toplam':
                    self.calculation_labels[key].setText(f"Toplam: {display_text}")
                elif key == 'Sepet':
                    self.calculation_labels[key].setText(f"Sepet: {display_text}")
                else:
                    self.calculation_labels[key].setText(f"{key} : {display_text}")

        except Exception as e:
            logging.error(f"Hesaplama hatası: {str(e)}")
            QMessageBox.warning(self, "Hesaplama Hatası", f"Toplamlar hesaplanırken bir hata oluştu: {str(e)}")

    def show_context_menu(self, pos):
        menu = QMenu()
        copy_action = QAction("Kopyala", self)
        copy_action.triggered.connect(self.copy_selection)
        menu.addAction(copy_action)
        menu.exec_(self.table.viewport().mapToGlobal(pos))

    def copy_selection(self):
        selected = self.table.selectedItems()
        if selected:
            rows = sorted({item.row() for item in selected})
            cols = sorted({item.column() for item in selected})

            clipboard_text = []
            for row in rows:
                row_data = []
                for col in cols:
                    item = self.table.item(row, col)
                    row_data.append(item.text() if item else "")
                clipboard_text.append("\t".join(row_data))

            QApplication.clipboard().setText("\n".join(clipboard_text))

    def handle_ctrl_c(self):
        """Ctrl+C ile kopyalama işlemi"""
        # Seçili öğeler varsa kopyala (çoklu seçim desteği)
        selected_items = self.table.selectedItems()
        if selected_items:
            # Sadece tek hücre seçiliyse direkt metni kopyala
            if len(selected_items) == 1:
                QApplication.clipboard().setText(selected_items[0].text())
            else:
                # Çoklu seçim varsa tablor formatında kopyala
                # (copy_selection mantığına benzer ama seçili olanları alır)
                # Basitçe ilk öğeyi veya tümünü kopyalayabiliriz, kullanıcı
                # 'Kopyala' dediğinde sağ tık menüsünde ne oluyorsa onu yapmalı
                # Sağ tık menüsünde 'copy_selection' çağrılıyor mu?
                # Hayır, kodda show_context_menu içinde ne vardı? 
                pass
                # show_context_menu içinde ne olduğunu göremedim tam olarak
                
                # Basit yaklaşımı kullanıyorum: sağ tık menüsündeki gibi
                # Sağ tık menüsünde ne olduğunu tekrar kontrol etmem gerekebilirdi
                # Ancak burada güvenli bir 'copy text' yapıyorum
                
                # Eğer tek satır kopyalama varsa:
                rows = sorted(list(set(item.row() for item in selected_items)))
                cols = sorted(list(set(item.column() for item in selected_items)))
                
                text_data = ""
                for r in rows:
                    row_items = []
                    for c in cols:
                        item = self.table.item(r, c)
                        if item and item.isSelected():
                            row_items.append(item.text())
                    if row_items:
                        text_data += "\t".join(row_items) + "\n"
                
                if text_data:
                    QApplication.clipboard().setText(text_data.strip())

    def save_order(self):
        try:
            filtered_data = self.filtered_df[self.filtered_df['Ver'] > 0]

            if filtered_data.empty:
                QMessageBox.information(self, "Bilgi", "Kaydedilecek Sipariş verisi bulunamadı.")
                return

            filtered_data = filtered_data.sort_values(by='SAP Kodu', ascending=True)
            filtered_data = filtered_data[['SAP Kodu', 'Ver', 'Malzeme Adı']].copy()
            # SAP Kodu sütununu integer'a çevir
            filtered_data['SAP Kodu'] = pd.to_numeric(filtered_data['SAP Kodu'], errors='coerce').astype('Int64')
            
            save_path = r"D:/GoogleDrive"
            if not os.path.exists(save_path):
                os.makedirs(save_path)

            days_tr = {
                0: "Pazartesi",
                1: "Salı",
                2: "Çarşamba",
                3: "Perşembe",
                4: "Cuma",
                5: "Cumartesi",
                6: "Pazar"
            }
            
            current_day = datetime.now().weekday()
            day_name = days_tr[current_day]
            time_part = datetime.now().strftime("%H - %M")
            current_time = f"{day_name} - {time_part}"
            file_name = f"~ Sipariş {current_time}.xlsx"
            file_path = os.path.join(save_path, file_name)

            filtered_data.to_excel(file_path, index=False, header=False)
            QMessageBox.information(self, "Başarılı", f"Sipariş dosyası kaydedildi: {file_path}")
        
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Sipariş kaydetme hatası: {str(e)}")

    def save_Sepet(self):
        try:
            filtered_data = self.filtered_df[self.filtered_df['Sepet'] > 0].copy()

            if filtered_data.empty:
                QMessageBox.information(self, "Bilgi", "Kaydedilecek veri bulunamadı.")
                return

            # Sepet_Timestamp'e göre sırala (Eski → Yeni)
            if 'Sepet_Timestamp' in filtered_data.columns:
                filtered_data = filtered_data.sort_values(
                    by='Sepet_Timestamp',
                    ascending=True,  # Eski → Yeni
                    na_position='last'  # Timestamp olmayanlar en sonda
                )

            filtered_data = filtered_data[['SAP Kodu', 'Sepet', 'Malzeme Adı']].copy()
            # SAP Kodu sütununu integer'a çevir
            filtered_data['SAP Kodu'] = pd.to_numeric(filtered_data['SAP Kodu'], errors='coerce').astype('Int64')
            
            save_path = r"D:/GoogleDrive"
            if not os.path.exists(save_path):
                os.makedirs(save_path)

            days_tr = {
                0: "Pazartesi",
                1: "Salı",
                2: "Çarşamba",
                3: "Perşembe",
                4: "Cuma",
                5: "Cumartesi",
                6: "Pazar"
            }
            
            current_day = datetime.now().weekday()
            day_name = days_tr[current_day]
            time_part = datetime.now().strftime("%H - %M")
            current_time = f"{day_name} - {time_part}"
            file_name = f"~ Sepet {current_time}.xlsx"
            file_path = os.path.join(save_path, file_name)

            filtered_data.to_excel(file_path, index=False, header=False)
            QMessageBox.information(self, "Başarılı", f"Sepet dosyası kaydedildi: {file_path}")
        
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Sepet kaydetme hatası: {str(e)}")

    def run_mikro(self):
        """Stok.exe dosyasını çalıştır ve verileri yenile"""
        try:
            exe_path = r"D:/GoogleDrive/PRG/EXE/Stok.exe"
            if not os.path.exists(exe_path):
                QMessageBox.warning(self, "Uyarı", f"Stok.exe bulunamadı: {exe_path}")
                return
            
            # Statusbar için label yoksa oluştur
            if not hasattr(self, 'status_label'):
                self.status_label = QLabel("Hazır")
                self.status_label.setStyleSheet("""
                    QLabel {
                        color: #000000;
                        padding: 8px;
                        background-color: #f0f0f0;
                        border-top: 1px solid #cccccc;
                        font-size: 13px;
                        font-weight: bold;
                    }
                """)
                self.main_layout.addWidget(self.status_label)
            
            self.status_label.setText("🔄 Stok.exe çalıştırılıyor...")
            self.micro_btn.setEnabled(False)
            self.mikro_calisiyor = True
            
            QApplication.processEvents()
            
            os.startfile(exe_path)
            
            # Stok.exe'nin çalışması için bekleme
            # 7 saniye sonra program bitmiş sayıp kontrol et
            QTimer.singleShot(10000, self.on_mikro_finished)
            
        except Exception as e:
            if hasattr(self, 'status_label'):
                self.status_label.setText(f"❌ Program çalıştırma hatası: {str(e)}")
            QMessageBox.critical(self, "Hata", f"Program çalıştırma hatası: {str(e)}")
            self.micro_btn.setEnabled(True)
            self.mikro_calisiyor = False

    def on_mikro_finished(self):
        """Mikro program bittikten sonra"""
        self.micro_btn.setEnabled(True)
        self.mikro_calisiyor = False
        if hasattr(self, 'status_label'):
            self.status_label.setText("✅ Stok.exe tamamlandı, Google Sheets güncelleme bekleniyor...")
        
        # Google Sheets'e kaydedilmesi için ek bekleme (5 saniye)
        # Sonra verileri yenile
        QTimer.singleShot(7000, self.delayed_data_refresh)
    
    def delayed_data_refresh(self):
        """Gecikmeli veri yenileme"""
        if hasattr(self, 'status_label'):
            self.status_label.setText("🔄 Google Sheets'ten güncel veriler alınıyor...")
        QApplication.processEvents()
        self.load_data()
        if hasattr(self, 'status_label'):
            self.status_label.setText("✅ Veriler başarıyla güncellendi")

    def run_bekleyen(self):
        """BekleyenFast.exe dosyasını çalıştır ve verileri yenile"""
        try:
            exe_path = r"D:\GoogleDrive\PRG\EXE\BekleyenFast.exe"
            if not os.path.exists(exe_path):
                QMessageBox.warning(self, "Uyarı", f"BekleyenFast.exe bulunamadı: {exe_path}")
                return
            
            # Statusbar için label yoksa oluştur
            if not hasattr(self, 'status_label'):
                self.status_label = QLabel("Hazır")
                self.status_label.setStyleSheet("""
                    QLabel {
                        color: #000000;
                        padding: 8px;
                        background-color: #f0f0f0;
                        border-top: 1px solid #cccccc;
                        font-size: 13px;
                        font-weight: bold;
                    }
                """)
                self.main_layout.addWidget(self.status_label)
            
            self.status_label.setText("🔄 BekleyenFast.exe çalıştırılıyor...")
            self.bekleyen_btn.setEnabled(False)
            self.bekleyen_calisiyor = True
            
            QApplication.processEvents()
            
            os.startfile(exe_path)
            
            # BekleyenFast.exe'nin çalışması için bekleme
            # 10 saniye sonra program bitmiş sayıp kontrol et
            QTimer.singleShot(10000, self.on_bekleyen_finished)
            
        except Exception as e:
            if hasattr(self, 'status_label'):
                self.status_label.setText(f"❌ Program çalıştırma hatası: {str(e)}")
            QMessageBox.critical(self, "Hata", f"Program çalıştırma hatası: {str(e)}")
            self.bekleyen_btn.setEnabled(True)
            self.bekleyen_calisiyor = False

    def on_bekleyen_finished(self):
        """Bekleyen program bittikten sonra"""
        self.bekleyen_btn.setEnabled(True)
        self.bekleyen_calisiyor = False
        if hasattr(self, 'status_label'):
            self.status_label.setText("✅ BekleyenFast.exe tamamlandı, Google Sheets güncelleme bekleniyor...")
        
        # Google Sheets'e kaydedilmesi için ek bekleme (7 saniye)
        # Sonra verileri yenile
        QTimer.singleShot(7000, self.delayed_data_refresh_bekleyen)
    
    def delayed_data_refresh_bekleyen(self):
        """Bekleyen için gecikmeli veri yenileme"""
        if hasattr(self, 'status_label'):
            self.status_label.setText("🔄 Google Sheets'ten güncel veriler alınıyor...")
        QApplication.processEvents()
        self.load_data()
        if hasattr(self, 'status_label'):
            self.status_label.setText("✅ Veriler başarıyla güncellendi - 5 saniye sonra Mikro çalışacak...")

        # 5 saniye bekledikten sonra run_mikro fonksiyonunu çalıştır
        QTimer.singleShot(5000, self.run_mikro)

    def filter_and_send_to_whatsapp(self):
        try:
            filtered_data = self.filtered_df[self.filtered_df['Sepet'] > 0]
            filtered_data = filtered_data[['SAP Kodu', 'Sepet', 'Malzeme Adı']].copy()
            # SAP Kodu sütununu integer'a çevir
            filtered_data['SAP Kodu'] = pd.to_numeric(filtered_data['SAP Kodu'], errors='coerce').astype('Int64')
            
            if filtered_data.empty:
                QMessageBox.information(self, "Bilgi", "Satış Ekibine gönderilecek liste yok.")
                return
            
            message = "Gün içinde gelen müşterinin fiyat sorduğu ürünler;\n\n"
            for index, row in filtered_data.iterrows():
                message += f"{row['Sepet']} x {row['Malzeme Adı']}\n"
            
            encoded_message = urllib.parse.quote(message)
            whatsapp_url = f"whatsapp://send?text={encoded_message}"
            webbrowser.open(whatsapp_url)
            
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"WhatsApp bağlantı hatası: {str(e)}")

    def Stoklistesi(self):
        try:
            save_path = r"D:/GoogleDrive"
            if not os.path.exists(save_path):
                os.makedirs(save_path)

            file_name = "~ Stok_Listesi.xlsx"
            file_path = os.path.join(save_path, file_name)

            # DataFrame kopyasını oluştur ve SAP Kodu sütununu integer'a çevir
            df_to_save = self.filtered_df.copy()
            if 'SAP Kodu' in df_to_save.columns:
                df_to_save['SAP Kodu'] = pd.to_numeric(df_to_save['SAP Kodu'], errors='coerce').astype('Int64')

            df_to_save.to_excel(file_path, index=False)
            QMessageBox.information(self, "Başarılı", f"Stok listesi Excel'e kaydedildi: {file_path}")
        
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Excel'e kaydetme hatası: {str(e)}")