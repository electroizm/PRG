# -*- coding: utf-8 -*-
"""
Sözleşme Modülü
=============

Bu modül, sözleşme (contract) yönetimi için kullanılan PyQt5 tabanlı bir GUI uygulamasıdır.

Ana Özellikler:
    - Sözleşme detaylarını görüntüleme
    - Müşteri ve sipariş bilgilerini gösterme
    - Ürün listelerini tablo formatında gösterme
    - Cari hesaplara aktarım yapma
    - Stok kartı oluşturma
    - Sipariş transferi yapma
    - Excel import/export işlemleri

Sınıflar:
    - ContractDetailsWindow: Sözleşme detaylarını gösteren ana pencere
    - TableUpdateDialog: Tablo güncelleme dialog'u
    - MusteriBilgileriDialog: Müşteri bilgileri gösterme dialog'u
    - CariSelectionDialog: Cari seçim dialog'u
    - SozlesmeApp: Ana sözleşme uygulama penceresi

Yazar: [Proje Ekibi]
Tarih: 2024
"""

# Standart kütüphane importları
import os
import sys
import re
import logging
from datetime import datetime, timedelta
from io import BytesIO
import importlib.util

# Parent directory'yi Python path'e ekle (central_config için)
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

# Üçüncü parti kütüphane importları
import pandas as pd
import pyodbc
import requests

# PyQt5 importları
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *

# QShortcut ve QKeySequence'in kesin olarak import edildiğinden emin olalım
from PyQt5.QtWidgets import QShortcut, QApplication
from PyQt5.QtGui import QKeySequence

# Logging konfigürasyonu - ERROR mesajlarını sustur
logging.basicConfig(level=logging.WARNING)

# Central config import
from central_config import CentralConfigManager

class ContractDetailsWindow(QMainWindow):
    """
    Sözleşme Detayları Penceresi

    Bu sınıf, bir sözleşmenin tüm detaylarını görüntülemek için kullanılan ana penceredir.
    Müşteri bilgileri, sipariş bilgileri, ürün listesi ve çeşitli işlem butonlarını içerir.

    Özellikler:
        - Sözleşme bilgilerini görsel olarak düzenlenmiş şekilde gösterir
        - Müşteri ve sipariş bilgilerini gruplar halinde sunar
        - Ürünleri düzenlenebilir tablo formatında listeler
        - Cari aktarım, stok aktarım ve sipariş aktarım işlemlerini destekler
        - IPT durumu ve header bilgisi kontrolü yapar
        - Tablo kopyalama özelliği sunar

    Attributes:
        contract_data: Sözleşme verileri içeren obje
        contract_id: Sözleşme kimlik numarası
        has_ipt_status (bool): Sözleşmede IPT durumu olup olmadığı
        header_empty_status (bool): Header bilgisinin boş olup olmadığı
        selected_cari_kod: Seçili cari hesap kodu
        create_order_btn: Stok aktarma butonu
        transfer_order_btn: Sipariş aktarma butonu
    """

    def __init__(self, contract_data, contract_id, parent=None):
        """
        ContractDetailsWindow sınıfının yapıcı metodu.

        Args:
            contract_data: Sözleşme bilgilerini içeren veri objesi
            contract_id: Sözleşmenin benzersiz kimlik numarası
            parent (QWidget, optional): Üst pencere widget'ı. Varsayılan None.
        """
        super().__init__(parent)
        self.contract_data = contract_data  # Sözleşme verilerini sakla
        self.contract_id = contract_id  # Sözleşme ID'sini sakla
        self.has_ipt_status = False  # Varsayılan olarak IPT yok (İptal durumu)
        self.header_empty_status = False  # Varsayılan olarak Header dolu
        self.setup_ui()  # Kullanıcı arayüzünü oluştur
        
    def get_product_items(self):
        """
        Contract_data'dan gerçek ürün listesini döndürür.
        """
        try:
            # SOAP response structure: contract_data.ES_CONTRACT_INFO.ITEMS.item
            if hasattr(self.contract_data, 'ES_CONTRACT_INFO'):
                contract_info = self.contract_data.ES_CONTRACT_INFO
                if hasattr(contract_info, 'ITEMS'):
                    items = contract_info.ITEMS
                    if hasattr(items, 'item'):
                        item_list = items.item
                        if isinstance(item_list, list):
                            return item_list
                        else:
                            return [item_list]
            return []
        except Exception as e:
            return []

    def check_cari_aktar_eligibility(self):
        """
        Cari Aktar butonunun aktif olup olmadığını kontrol eder.
        Tablodan direkt güncel verileri okur.

        Returns:
            tuple: (bool, str) - True/False ve uyarı mesajı
        """
        try:
            # IPT durumu veya Header boş ise inaktif
            if (hasattr(self, 'has_ipt_status') and self.has_ipt_status) or \
               (hasattr(self, 'header_empty_status') and self.header_empty_status):
                return False, "IPT veya Header durumu nedeniyle aktif değil"

            # Tabloyu kontrol et
            if not hasattr(self, 'products_table') or self.products_table is None:
                return False, "Ürünler tablosu bulunamadı"

            products_table = self.products_table
            row_count = products_table.rowCount()

            # Tüm satırları kontrol et - hepsinin uygun olması gerekiyor
            for row_index in range(row_count):
                # Tablodaki sütun indeksleri:
                # 0: Satır, 1: SAP Kodu, 2: Malzeme Adı, 3: SPEC,
                # 4: Miktar, 5: Birim Fiyat, 6: Net Tutar, 7: KDV,
                # 8: Sipariş No, 9: Sip Kalem No

                sap_kod_item = products_table.item(row_index, 1)
                spec_item = products_table.item(row_index, 3)
                siparis_no_item = products_table.item(row_index, 8)
                sip_kalem_no_item = products_table.item(row_index, 9)

                sap_kod = sap_kod_item.text().strip() if sap_kod_item else ""
                spec_value = spec_item.text().strip() if spec_item else ""
                siparis_no = siparis_no_item.text().strip() if siparis_no_item else ""
                sip_kalem_no = sip_kalem_no_item.text().strip() if sip_kalem_no_item else ""

                # SPEC değerinin gerçekten dolu olup olmadığını kontrol et
                spec_lower = spec_value.lower()
                is_spec_filled = (spec_value and
                                spec_lower != '' and
                                'none' not in spec_lower and
                                'null' not in spec_lower)

                # SPEC dolu ise sipariş bilgileri zorunlu
                if is_spec_filled:
                    if not siparis_no or not sip_kalem_no:
                        return False, f"SAP Kodu {sap_kod}: SPEC dolu ama Sipariş bilgileri eksik"

            # Tüm satırlar uygunsa buton aktif
            return True, "Tüm gerekli bilgiler tamamlandı"

        except Exception as e:
            logging.error(f"Cari aktar uygunluk kontrolü hatası: {e}")
            return False, f"Kontrol hatası: {e}"

    def update_cari_aktar_button(self):
        """
        Cari Aktar butonunun durumunu güncelleyen fonksiyon.
        Tablo değişiklikleri sonrası çağrılır.
        """
        try:
            # Buton henüz oluşturulmadıysa hiçbir şey yapma
            if not hasattr(self, 'transfer_btn') or self.transfer_btn is None:
                return

            cari_aktar_eligible, message = self.check_cari_aktar_eligibility()
            self.transfer_btn.setEnabled(cari_aktar_eligible)

            if cari_aktar_eligible:
                self.transfer_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #3498db;
                        color: white;
                        border: none;
                        padding: 10px 20px;
                        font-size: 14px;
                        font-weight: bold;
                        border-radius: 5px;
                        min-width: 120px;
                    }
                    QPushButton:hover {
                        background-color: #2980b9;
                    }
                    QPushButton:pressed {
                        background-color: #21618c;
                    }
                """)
            else:
                self.transfer_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #bdc3c7;
                        color: #7f8c8d;
                        border: none;
                        padding: 10px 20px;
                        font-size: 14px;
                        font-weight: bold;
                        border-radius: 5px;
                        min-width: 120px;
                    }
                """)

        except Exception as e:
            logging.error(f"Cari aktar buton güncelleme hatası: {e}")

    def update_contract_data_from_table(self, table):
        """
        Tablodaki değişiklikleri contract_data'ya yansıtır.
        """
        try:
            # Gerçek ürün listesini al
            product_items = self.get_product_items()

            if not product_items:
                return

            for row in range(table.rowCount()):
                # SAP Kodu ile ürün listesinde eşleşen kaydı bul
                sap_item = table.item(row, 1)  # SAP Kodu sütunu
                if not sap_item:
                    continue

                sap_kod = sap_item.text().strip()

                # Ürün listesinde bu SAP kodunu bul
                for i, product_item in enumerate(product_items):
                    try:
                        # SOAP objesi için PRODUCT_CODE attribute'unu kontrol et
                        product_code = str(getattr(product_item, 'PRODUCT_CODE', '')).strip()

                        if product_code == sap_kod:
                            # SPEC güncelle (sütun 3)
                            spec_item = table.item(row, 3)
                            if spec_item:
                                new_spec = spec_item.text().strip()
                                setattr(product_item, 'SPEC', new_spec)

                            # Sipariş No güncelle (sütun 8)
                            siparis_item = table.item(row, 8)
                            if siparis_item:
                                new_siparis = siparis_item.text().strip()
                                setattr(product_item, 'Siparis_No', new_siparis)

                            # Sip Kalem No güncelle (sütun 9)
                            kalem_item = table.item(row, 9)
                            if kalem_item:
                                new_kalem = kalem_item.text().strip()
                                setattr(product_item, 'Sip_Kalem_No', new_kalem)
                            break
                    except Exception as inner_e:
                        continue

        except Exception as e:
            logging.error(f"Contract data güncelleme hatası: {e}")

    def setup_ui(self):
        """
        Kullanıcı arayüzünü oluşturur ve yapılandırır.

        Bu metod:
        - Pencere başlığını ve boyutunu ayarlar
        - Ana widget'ları oluşturur
        - Scroll area ekler
        - Sözleşme bilgi bölümlerini yükler
        - Tablo kopyalama özelliğini aktive eder
        """
        # Pencere başlığını sözleşme ID'si ile ayarla
        self.setWindowTitle(f"Sözleşme - {self.contract_id}")
        
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
        
        # Layout
        layout = QVBoxLayout(central_widget)
        
        # Scroll area
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        
        # Sözleşme bilgilerini göster
        if hasattr(self.contract_data, 'ES_CONTRACT_INFO'):
            contract_info = self.contract_data.ES_CONTRACT_INFO
            self.add_contract_info_sections(scroll_layout, contract_info)
        else:
            error_label = QLabel("Sözleşme bilgileri alınamadı")
            error_label.setStyleSheet("color: red; font-size: 14px; padding: 10px;")
            scroll_layout.addWidget(error_label)
        
        scroll_area.setWidget(scroll_widget)
        layout.addWidget(scroll_area)
        
        # Kapat butonu
        close_btn = QPushButton("Kapat")
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                border: none;
                padding: 10px 20px;
                font-size: 14px;
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        close_btn.clicked.connect(self.close)
        # Kapat butonu artık ürünler grubunda, burayı kaldırıyoruz
        # layout.addWidget(close_btn)

        # ÜRÜNLER tablosuna kopyalama özelliği ekle
        QTimer.singleShot(1000, self.setup_table_copy_functionality)  # 1 saniye sonra setup et

    def setup_table_copy_functionality(self):
        """Ürünler tablosu için Ctrl+C kısayolunu ayarlar"""
        try:
            if hasattr(self, 'products_table') and self.products_table:
                self.copy_shortcut = QShortcut(QKeySequence("Ctrl+C"), self.products_table)
                self.copy_shortcut.activated.connect(lambda: self.copy_table_selection(self.products_table))
        except Exception as e:
            pass

    def copy_table_selection(self, table):
        """Seçili hücreyi panoya kopyalar"""
        selected_items = table.selectedItems()
        if selected_items:
            # Sadece ilk seçili öğeyi kopyala (tek hücre seçimi varsayımıyla)
            text = selected_items[0].text()
            QApplication.clipboard().setText(text)
        
    def add_contract_info_sections(self, layout, contract_info):
        """
        Sözleşme bilgi bölümlerini layout'a ekler.

        Bu metod, sözleşmeye ait tüm bilgi bölümlerini (müşteri, sipariş, ürünler)
        uygun şekilde formatlar ve arayüze ekler.

        Args:
            layout (QVBoxLayout): Bilgilerin ekleneceği layout
            contract_info: Sözleşme bilgilerini içeren veri objesi

        İşlem Adımları:
            1. Müşteri bilgilerini gruplar ve gösterir
            2. Sipariş bilgilerini düzenler ve gösterir
            3. Header ve durum bilgilerini kontrol eder
            4. IPT ve header boşluk durumlarını tespit eder
            5. Ürün listesini tablo formatında gösterir
            6. Alt kontrol butonlarını ekler
        """
        
        def safe_get(obj, attr, default='N/A'):
            if not obj:
                return default
            return getattr(obj, attr, default) if hasattr(obj, attr) else default
        
        # Üst bilgiler için yatay layout
        top_info_layout = QHBoxLayout()
        
        # Müşteri Bilgileri
        customer_name = f"{safe_get(contract_info, 'CUSTOMER_NAMEFIRST')} {safe_get(contract_info, 'CUSTOMER_NAMELAST')}".strip()
        
        customer_group = self.create_customer_info_group("MÜŞTERİ BİLGİLERİ", {
            'ad_soyad': customer_name,
            'telefon1': safe_get(contract_info, 'CUSTOMER_PHONE1'),
            'telefon2': safe_get(contract_info, 'CUSTOMER_PHONE2'),
            'vergi_no': safe_get(contract_info, 'CUSTOMER_TAXNR'),
            'posta_kodu': safe_get(contract_info, 'CUSTOMER_POSTCODE'),
            'sehir_ilce': safe_get(contract_info, 'CUSTOMER_CITY'),
            'adres': safe_get(contract_info, 'CUSTOMER_ADDRESS')
        })
        top_info_layout.addWidget(customer_group)
        
        # Sipariş Bilgileri
        header_text_raw = safe_get(contract_info, 'HEADER_TEXT', '0')

        # Header'ı integer'a çevir
        header_int = self.convert_turkish_to_integer(header_text_raw)

        salesman_name = f"{safe_get(contract_info, 'SALESMAN_NAMEFIRST')} {safe_get(contract_info, 'SALESMAN_NAMELAST')}".strip()

        # Mağaza bilgisini düzenle
        sales_office = safe_get(contract_info, 'SALES_OFFICE')
        if sales_office == 'IM1':
            magaza_display = '1600704 - MERKEZ'
        elif sales_office == 'IM2':
            magaza_display = '1601175 - ŞUBE'
        else:
            magaza_display = sales_office

        # Durum bilgisini kontrol et
        status_text = safe_get(contract_info, 'STATUS_TEXT')
        status_code = safe_get(contract_info, 'STATUS')
        full_status = f"{status_text} ({status_code})"
        has_ipt = '(IPT)' in full_status

        # Header bilgisini kontrol et (boş veya 0 ise)
        header_empty = header_int == 0

        # Header display formatını hazırla
        if header_empty:
            header_display = ""  # Boş string
        else:
            header_display = f"{header_int:,} TL"
        
        order_sales_group = self.create_order_info_group("SİPARİŞ BİLGİLERİ", {
            'siparis_tarihi': safe_get(contract_info, 'ORD_DATE'),
            'fiyat_listesi': safe_get(contract_info, 'PRICE_LIST_TEXT'),
            'header': header_display,
            'magaza': magaza_display,
            'personel': salesman_name,
            'durum': full_status,
            'has_ipt': has_ipt,
            'header_empty': header_empty
        })
        top_info_layout.addWidget(order_sales_group)
        
        # IPT durumu ve header boşluk durumu için transfer buton durumunu sakla
        self.has_ipt_status = has_ipt
        self.header_empty_status = header_empty
        
        layout.addLayout(top_info_layout)
        
        # Ürünler
        if hasattr(contract_info, 'ITEMS') and hasattr(contract_info.ITEMS, 'item'):
            products_group = self.create_products_group(contract_info.ITEMS.item)
            layout.addWidget(products_group)
            
            # Butonlar ve toplamlar - ÜRÜNLER grubunun dışında
            self.add_bottom_controls(layout, contract_info.ITEMS.item)
            
    def create_info_group(self, title, items):
        """
        Genel amaçlı bilgi grubu oluşturur.

        Bu metod, verilen başlık ve öğeler ile stilize edilmiş bir bilgi grubu
        (QGroupBox) oluşturur.

        Args:
            title (str): Grup başlığı
            items (list): (etiket, değer) tuple'larından oluşan liste

        Returns:
            QGroupBox: Oluşturulan bilgi grubu widget'ı

        Not:
            - Sadece boş olmayan ve 'N/A' olmayan değerler gösterilir
            - Mavi kenarlık ve açık mavi arka plan kullanır
            - Etiketler kalın, değerler normal yazı tipi ile gösterilir
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
        
        layout = QVBoxLayout()
        
        for label_text, value in items:
            if value and str(value).strip() and str(value) != 'N/A':
                item_layout = QHBoxLayout()
                
                label = QLabel(f"{label_text}:")
                label.setMinimumWidth(120)
                label.setStyleSheet("font-weight: bold; color: #34495e;")
                
                value_label = QLabel(str(value))
                value_label.setStyleSheet("color: #2c3e50;")
                value_label.setWordWrap(True)
                
                item_layout.addWidget(label)
                item_layout.addWidget(value_label, 1)
                layout.addLayout(item_layout)
        
        group_box.setLayout(layout)
        return group_box
    
    def create_order_info_group(self, title, info_dict):
        """
        Sipariş bilgileri için özel grup oluşturur - 2 sütunlu layout.

        Bu metod, sipariş bilgilerini 2 sütunlu bir düzende gösterir.
        Sol sütunda sipariş tarihi, fiyat listesi ve header; sağ sütunda
        mağaza, personel ve durum bilgileri yer alır.

        Args:
            title (str): Grup başlığı
            info_dict (dict): Sipariş bilgilerini içeren sözlük
                - siparis_tarihi: Sipariş tarihi
                - fiyat_listesi: Kullanılan fiyat listesi
                - header: Header tutarı
                - magaza: Mağaza bilgisi
                - personel: Satış personeli
                - durum: Sipariş durumu
                - has_ipt: IPT durumu (bool)
                - header_empty: Header boş mu (bool)

        Returns:
            QGroupBox: 2 sütunlu sipariş bilgileri grubu

        Not:
            - IPT içeren durumlar kırmızı renkte gösterilir
            - Boş header bilgisi kırmızı renkte vurgulanır
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
        main_layout = QHBoxLayout()
        
        # Sol sütun
        left_layout = QVBoxLayout()
        left_items = [
            ("Sipariş Tarihi:", info_dict['siparis_tarihi']),
            ("Fiyat Listesi:", info_dict['fiyat_listesi']),
            ("Header:", info_dict['header'])
        ]
        
        for i, (label_text, value) in enumerate(left_items):
            if value and str(value).strip() and str(value) != 'N/A':
                item_layout = QHBoxLayout()
                
                label = QLabel(label_text)
                label.setStyleSheet("font-weight: bold; color: #34495e; font-size: 16px;")
                
                value_label = QLabel(str(value))
                
                # Header alanı için özel kontrol (3. item Header)
                if i == 2 and label_text == "Header:" and info_dict.get('header_empty', False):
                    value_label.setStyleSheet("color: #e74c3c; font-size: 16px; font-weight: bold;")
                else:
                    value_label.setStyleSheet("color: #2c3e50; font-size: 16px; font-weight: bold;")
                    
                value_label.setWordWrap(True)
                
                item_layout.addWidget(label)
                item_layout.addWidget(value_label, 1)
                left_layout.addLayout(item_layout)
        
        # Sağ sütun
        right_layout = QVBoxLayout()
        
        # Mağaza
        if info_dict['magaza'] and str(info_dict['magaza']).strip() and str(info_dict['magaza']) != 'N/A':
            magaza_layout = QHBoxLayout()
            magaza_label = QLabel("Mağaza:")
            magaza_label.setStyleSheet("font-weight: bold; color: #34495e; font-size: 16px;")
            magaza_value = QLabel(str(info_dict['magaza']))
            magaza_value.setStyleSheet("color: #2c3e50; font-size: 16px; font-weight: bold;")
            magaza_layout.addWidget(magaza_label)
            magaza_layout.addWidget(magaza_value, 1)
            right_layout.addLayout(magaza_layout)
        
        # Personel
        if info_dict['personel'] and str(info_dict['personel']).strip() and str(info_dict['personel']) != 'N/A':
            personel_layout = QHBoxLayout()
            personel_label = QLabel("Personel:")
            personel_label.setStyleSheet("font-weight: bold; color: #34495e; font-size: 16px;")
            personel_value = QLabel(str(info_dict['personel']))
            personel_value.setStyleSheet("color: #2c3e50; font-size: 16px; font-weight: bold;")
            personel_layout.addWidget(personel_label)
            personel_layout.addWidget(personel_value, 1)
            right_layout.addLayout(personel_layout)
        
        # Durum
        if info_dict['durum'] and str(info_dict['durum']).strip() and str(info_dict['durum']) != 'N/A':
            durum_layout = QHBoxLayout()
            durum_label = QLabel("Durum:")
            durum_label.setStyleSheet("font-weight: bold; color: #34495e; font-size: 16px;")
            durum_value = QLabel(str(info_dict['durum']))
            # IPT içeriyorsa kırmızı yap
            if info_dict['has_ipt']:
                durum_value.setStyleSheet("color: #e74c3c; font-size: 16px; font-weight: bold;")
            else:
                durum_value.setStyleSheet("color: #2c3e50; font-size: 16px; font-weight: bold;")
            durum_layout.addWidget(durum_label)
            durum_layout.addWidget(durum_value, 1)
            right_layout.addLayout(durum_layout)
        
        # Sütunları ana layout'a ekle
        main_layout.addLayout(left_layout)
        main_layout.addSpacing(20)  # Sütunlar arası boşluk
        main_layout.addLayout(right_layout)
        
        group_box.setLayout(main_layout)
        return group_box
    
    def create_customer_info_group(self, title, info_dict):
        """
        Müşteri bilgileri için özel grup oluşturur - 3x2 grid layout.

        Bu metod, müşteri bilgilerini grid (ızgara) düzeninde gösterir.
        Sol sütunda ad soyad ve telefonlar; sağ sütunda TCKN, posta kodu ve şehir;
        alt kısımda tam genişlikte adres bilgisi yer alır.

        Args:
            title (str): Grup başlığı (örn: "MÜŞTERİ BİLGİLERİ")
            info_dict (dict): Müşteri bilgilerini içeren sözlük
                - ad_soyad: Müşteri adı soyadı
                - telefon1: Birinci telefon numarası
                - telefon2: İkinci telefon numarası
                - vergi_no: TCKN/Vergi numarası
                - posta_kodu: Posta kodu
                - sehir_ilce: Şehir ve ilçe bilgisi
                - adres: Tam adres

        Returns:
            QGroupBox: Grid düzeninde müşteri bilgileri grubu

        Not:
            - Adres bilgisi kelime sarma (word wrap) özelliği ile gösterilir
            - Boş veya 'N/A' değerler gösterilmez
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
            ("Ad Soyad:", info_dict['ad_soyad']),
            ("Telefon 1:", info_dict['telefon1']),
            ("Telefon 2:", info_dict['telefon2'])
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
            ("TCKN No:", info_dict['vergi_no']),
            ("Posta Kodu:", info_dict['posta_kodu']),
            ("Şehir:", info_dict['sehir_ilce'])
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
        
        # Adres alanı - tam genişlik, aynı satırda başlayıp gerekirse alt satıra geçsin
        if info_dict['adres'] and str(info_dict['adres']).strip() and str(info_dict['adres']) != 'N/A':
            main_layout.addSpacing(10)  # Grid ile adres arası boşluk
            
            adres_layout = QHBoxLayout()  # Dikey yerine yatay layout
            
            adres_label = QLabel("Adres:")
            adres_label.setStyleSheet("font-weight: bold; color: #34495e; font-size: 16px;")
            adres_label.setAlignment(Qt.AlignTop)  # Label'ı üst hizaya al
            
            adres_value = QLabel(str(info_dict['adres']))
            adres_value.setStyleSheet("color: #2c3e50; font-size: 16px; font-weight: bold;")
            adres_value.setWordWrap(True)  # Kelime sarımı aktif
            adres_value.setAlignment(Qt.AlignTop)  # Değeri de üst hizaya al
            
            adres_layout.addWidget(adres_label)
            adres_layout.addWidget(adres_value, 1)  # Değer alanını genişletsin
            
            main_layout.addLayout(adres_layout)
        
        group_box.setLayout(main_layout)
        return group_box
        
    def create_products_group(self, items):
        """
        Ürünler için tablo grubu oluşturur.

        Bu metod, sözleşmeye ait ürünleri düzenlenebilir bir tablo formatında gösterir.
        10 sütunlu tablo: Satır, SAP Kodu, Malzeme Adı, SPEC, Miktar, Birim Fiyat,
        Net Tutar, KDV, Sipariş No, Sip Kalem No.

        Args:
            items (list): Ürün bilgilerini içeren liste

        Returns:
            QGroupBox: Ürün tablosunu içeren grup

        Özellikler:
            - SPEC, Sipariş No ve Sip Kalem No sütunları düzenlenebilir
            - Diğer sütunlar salt okunur
            - Satır yüksekliği otomatik ayarlanır
            - Hücre değişikliklerinde otomatik validasyon yapılır
            - Toplam net tutar ve KDV hesaplanır

        Not:
            - MODUL bilgisi SPEC'ten filtrelenir
            - Miktar formatı düzenlenir (.0 kaldırılır)
            - Sipariş numaraları ve kalem numaraları formatlanır
        """
        group_box = QGroupBox("ÜRÜNLER")
        group_box.setStyleSheet("""
            QGroupBox {
                font-size: 16px;
                font-weight: bold;
                border: 3px solid #e74c3c;
                border-radius: 8px;
                margin-top: 15px;
                padding-top: 15px;
                background-color: #fdf2f2;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 5px 10px 5px 10px;
                color: white;
                background-color: #e74c3c;
                border-radius: 4px;
                font-size: 16px;
            }
        """)
        
        layout = QVBoxLayout()
        
        def safe_get(obj, attr, default='N/A'):
            if not obj:
                return default
            return getattr(obj, attr, default) if hasattr(obj, attr) else default
        
        total_net = 0
        total_tax = 0
        
        # Ürünler için tablo oluştur
        self.products_table = QTableWidget()
        products_table = self.products_table  # Geriye dönük uyumluluk için
        products_table.setRowCount(len(items))
        products_table.setColumnCount(10)  # Satir, SAP Kodu, Malzeme Adı, SPEC, Miktar, Birim Fiyat, Net Tutar, KDV, Sipariş No, Sip Kalem No
        products_table.setHorizontalHeaderLabels(["Satir", "SAP Kodu", "Malzeme Adı", "SPEC", "Miktar", "Birim Fiyat", "Net Tutar", "KDV", "Sipariş No", "Sip Kalem No"])

        # Satır yüksekliğini stok_module.py gibi ayarla
        products_table.verticalHeader().setDefaultSectionSize(products_table.verticalHeader().defaultSectionSize() + 2)

        # stok_module.py ile aynı stil uygula
        products_table.setStyleSheet("""
            QTableWidget {
                font-size: 15px;
                font-weight: bold;
            }
            QTableWidget::item:focus {
                outline: none;
                border: none;
            }
        """)
        products_table.setFocusPolicy(Qt.NoFocus)

        # Ctrl+C Kısayolu Ekle
        try:
            self.copy_shortcut_products = QShortcut(QKeySequence("Ctrl+C"), products_table)
            self.copy_shortcut_products.activated.connect(lambda: self.copy_table_selection(products_table))
        except Exception:
            pass

        # Hücre değişiklik event handler'ını bağla
        products_table.itemChanged.connect(lambda item: self.handle_products_table_edit(item, products_table))
        
        for i, item in enumerate(items):
            product_code = safe_get(item, 'PRODUCT_CODE')
            description = safe_get(item, 'DESCRIPTION')
            quantity = safe_get(item, 'QUANTITY')
            unit_price = self.convert_turkish_to_integer(safe_get(item, 'UNIT_PRICE', '0'))
            total_price = self.convert_turkish_to_integer(safe_get(item, 'TOTAL_PRICE', '0'))
            net_amount = self.convert_turkish_to_integer(safe_get(item, 'NET_AMOUNT', '0'))
            tax_amount = self.convert_turkish_to_integer(safe_get(item, 'TAX_AMOUNT', '0'))
            tax_rate = safe_get(item, 'TAX_RATE', '0')
            discount = self.convert_turkish_to_integer(safe_get(item, 'TOTAL_DISCOUNT', '0'))
            siparis = safe_get(item, 'SIPARIS')
            sip_kalem_no = safe_get(item, 'SIP_KALEM_NO')
            kalem_no = safe_get(item, 'KALEM_NO')
            
            # Satir (Kalem No) formatını düzenle - tüm sıfırları kaldır
            satir_display = ""
            if kalem_no not in [None, 'None', 'N/A']:
                try:
                    # Sayıya çevir ve sıfırları kaldır
                    satir_num = int(str(kalem_no).lstrip('0') or '0')
                    satir_display = str(satir_num)
                except:
                    satir_display = str(kalem_no)
            
            # Sip Kalem No formatını düzenle - ilk 4 sıfırı kaldır
            sip_kalem_display = ""
            if str(sip_kalem_no) not in ['0000000000', '0', 'None', 'N/A']:
                sip_kalem_str = str(sip_kalem_no)
                if len(sip_kalem_str) >= 4 and sip_kalem_str.startswith('0000'):
                    sip_kalem_display = sip_kalem_str[4:]  # İlk 4 karakteri kaldır
                else:
                    sip_kalem_display = sip_kalem_str
            
            total_net += net_amount
            total_tax += tax_amount
            
            # SPEC bilgilerini topla
            spec_text = ""
            if hasattr(item, 'SPEC') and hasattr(item.SPEC, 'item'):
                spec_list = []
                for spec_item in item.SPEC.item:
                    charc = safe_get(spec_item, 'CHARC')
                    value = safe_get(spec_item, 'VALUE')
                    if charc != 'N/A' and value != 'N/A':
                        # 'MODUL:' ifadesini ve verisini atla
                        if charc != 'MODUL':
                            spec_list.append(f"{charc}: {value}")
                spec_text = ", ".join(spec_list)
            
            # Miktar formatını düzenle (.0 kaldır)
            try:
                quantity_num = float(quantity)
                if quantity_num == int(quantity_num):
                    quantity_display = f"{int(quantity_num)}"
                else:
                    quantity_display = f"{quantity_num}"
            except:
                quantity_display = str(quantity).replace('.0', '') if '.0' in str(quantity) else str(quantity)
            
            # Sipariş bilgilerini düzenle (None ve sıfırları boş bırak)
            siparis_display = "" if siparis in [None, 'None', 'N/A'] else str(siparis)
            
            # Tabloya verileri ekle
            products_table.setItem(i, 0, QTableWidgetItem(satir_display))
            products_table.setItem(i, 1, QTableWidgetItem(str(product_code)))
            products_table.setItem(i, 2, QTableWidgetItem(str(description)))
            products_table.setItem(i, 3, QTableWidgetItem(spec_text))
            products_table.setItem(i, 4, QTableWidgetItem(quantity_display))
            products_table.setItem(i, 5, QTableWidgetItem(f"{unit_price:,} TL"))
            products_table.setItem(i, 6, QTableWidgetItem(f"{net_amount:,} TL"))
            products_table.setItem(i, 7, QTableWidgetItem(f"{tax_rate}%"))  # Sadece yüzde
            products_table.setItem(i, 8, QTableWidgetItem(siparis_display))
            products_table.setItem(i, 9, QTableWidgetItem(sip_kalem_display))
            
            # SPEC (3), Sipariş No (8), Sip Kalem No (9) sütunları düzenlenebilir, diğerleri düzenlenemez
            editable_columns = [3, 8, 9]  # SPEC, Sipariş No, Sip Kalem No
            for j in range(10):
                item_widget = products_table.item(i, j)
                if item_widget:
                    if j in editable_columns:
                        # Düzenlenebilir hücreler
                        item_widget.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable)
                    else:
                        # Düzenlenemez hücreler
                        item_widget.setFlags(item_widget.flags() & ~Qt.ItemIsEditable)
        
        # Tablo boyutunu ayarla
        products_table.resizeColumnsToContents()
        products_table.setMaximumHeight(200 + (len(items) * 30))  # Dinamik yükseklik
        
        # Sütun genişliklerini dinamik olarak ayarla
        products_table.resizeColumnsToContents()
        
        # Header ayarları
        header = products_table.horizontalHeader()
        header.setStretchLastSection(True)  # Son sütun genişlensin
        header.setSectionResizeMode(QHeaderView.Interactive)  # Kullanıcı sütunları yeniden boyutlandırabilsin
        
        layout.addWidget(products_table)
        
        group_box.setLayout(layout)
        return group_box

    def convert_turkish_to_integer(self, value):
        """
        Türkçe sayı formatını integer'a çevirir.

        Türkçe format: 12.345,67 (nokta binlik ayırıcı, virgül ondalık ayırıcı)

        Args:
            value: String, float veya int değer

        Returns:
            int: Yuvarlanmış integer değer

        Örnekler:
            "12.345,67" -> 12346
            "12.345" -> 12345
            "12345,67" -> 12346
            12345.67 -> 12346
        """
        import re

        if not value or value in ['None', 'N/A', '', 0]:
            return 0

        # String'e çevir
        text_str = str(value).strip()

        # Para birimi sembollerini ve gereksiz karakterleri temizle
        text_str = text_str.replace('₺', '').replace('TL', '').strip()

        # Sayısal kısmı bul
        match = re.search(r'([\d.,]+)', text_str)

        if match:
            numeric_part = match.group(1)

            try:
                # Türkçe format kontrolü: virgül varsa ondalık ayırıcıdır
                if ',' in numeric_part:
                    # Noktaları (binlik ayırıcı) kaldır, virgülü noktaya çevir
                    numeric_part = numeric_part.replace('.', '').replace(',', '.')
                    return int(round(float(numeric_part)))
                else:
                    # Virgül yoksa, noktalar binlik ayırıcı olabilir
                    # Eğer birden fazla nokta varsa binlik ayırıcıdır, kaldır
                    if numeric_part.count('.') > 1:
                        numeric_part = numeric_part.replace('.', '')
                        return int(round(float(numeric_part)))
                    else:
                        # Tek nokta varsa, ondalık kısmının uzunluğuna bak
                        parts = numeric_part.split('.')
                        if len(parts) == 2:
                            # Ondalık kısım 3 haneden fazla ise binlik ayırıcıdır
                            if len(parts[1]) >= 3:
                                numeric_part = numeric_part.replace('.', '')
                                return int(round(float(numeric_part)))
                            else:
                                # 2 hane veya daha az ise ondalık ayırıcıdır
                                return int(round(float(numeric_part)))
                        else:
                            return int(round(float(numeric_part)))
            except:
                return 0

        return 0

    def extract_first_numeric_part(self, text):
        """
        Header metninden sadece ilk sayısal kısmı çıkarır.

        Bu metod, header alanındaki metinden sayısal değeri ayıklar ve
        Türkçe sayı formatını (virgül, nokta) düzgün şekilde işler.

        Args:
            text (str): İşlenecek metin (header alanından gelen değer)

        Returns:
            str: Çıkarılan sayısal değer (string formatında)

        Örnekler:
            "12.345,67 TL" -> "12345,67"
            "1234 TL" -> "1234"
            "" -> "0"
        """
        import re

        # Boş veya geçersiz değerler için 0 dön
        if not text or text in ['None', 'N/A', '']:
            return '0'

        # String'e çevir
        text_str = str(text).strip()
        
        # İlk sayısal kısmı bul (nokta, virgül, rakam içeren)
        # Boşluk veya harf gelene kadar olan kısım
        match = re.match(r'^([\d.,]+)', text_str)
        
        if match:
            numeric_part = str(match.group(1))  # String'e çevir
            # Eğer virgül varsa Türkçe formatında işle
            if ',' in numeric_part:
                # Son virgül ondalik ayirıcısı ise
                parts = numeric_part.split(',')
                if len(parts) == 2 and len(parts[1]) <= 2:
                    # Binlik ayırıcı noktaları kaldır
                    integer_part = parts[0].replace('.', '')
                    return f"{integer_part},{parts[1]}"
                else:
                    # Sadece binlik ayırıcı
                    return numeric_part.replace(',', '')
            else:
                # Sadece nokta varsa binlik ayırıcı olarak işle
                if numeric_part.count('.') > 1:
                    return numeric_part.replace('.', '')
                else:
                    return numeric_part
        else:
            return '0'
    
    def add_bottom_controls(self, parent_layout, items):
        """
        Alt kontrol butonları ve toplam bilgilerini ekler.

        Bu metod, ürün tablosunun altına kontrol butonlarını ve finansal toplam
        bilgilerini ekler. Ayrıca yüzde hesaplaması yaparak iskonto oranını gösterir.

        Args:
            parent_layout (QVBoxLayout): Kontrollerin ekleneceği ana layout
            items (list): Ürün bilgilerini içeren liste

        Layout Yapısı:
            Sol: Butonlar (Cari Aktar, Stok Aktar, Sipariş Aktar, Kapat)
            Sağ: Toplam bilgileri ve yüzde hesabı
        """
        def safe_get(obj, attr, default='N/A'):
            """Güvenli veri çekme yardımcı fonksiyonu"""
            if not obj:
                return default
            return getattr(obj, attr, default) if hasattr(obj, attr) else default
        
        # safe_get metodunu sınıf içinde kullanabilmek için
        self.safe_get = safe_get
        
        # Toplamları hesapla
        total_net = 0
        total_tax = 0
        
        for item in items:
            net_amount = self.convert_turkish_to_integer(safe_get(item, 'NET_AMOUNT', '0'))
            tax_amount = self.convert_turkish_to_integer(safe_get(item, 'TAX_AMOUNT', '0'))
            total_net += net_amount
            total_tax += tax_amount
        
        if total_net > 0 or total_tax > 0:
            # Alt kısım için yatay layout
            bottom_layout = QHBoxLayout()
            bottom_layout.setContentsMargins(0, 10, 0, 0)  # Üst kısımda biraz boşluk
            
            # Sol taraf: Butonlar
            buttons_layout = QVBoxLayout()
            buttons_layout.setSpacing(5)  # Butonlar arası minimal boşluk
            buttons_layout.setContentsMargins(0, 0, 0, 0)
            
            # Cari Aktar butonunun uygunluğunu kontrol et
            cari_aktar_eligible, message = self.check_cari_aktar_eligibility()

            transfer_btn = QPushButton("Cari Aktar")
            if not cari_aktar_eligible:
                transfer_btn.setEnabled(False)
                transfer_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #bdc3c7;
                        color: #7f8c8d;
                        border: none;
                        padding: 10px 20px;
                        font-size: 14px;
                        font-weight: bold;
                        border-radius: 5px;
                        min-width: 120px;
                    }
                """)
            else:
                transfer_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #3498db;
                        color: white;
                        border: none;
                        padding: 10px 20px;
                        font-size: 14px;
                        font-weight: bold;
                        border-radius: 5px;
                        min-width: 120px;
                    }
                    QPushButton:hover {
                        background-color: #2980b9;
                    }
                    QPushButton:pressed {
                        background-color: #21618c;
                    }
                """)

            transfer_btn.clicked.connect(self.transfer_order_to_cari)
            buttons_layout.addWidget(transfer_btn)

            # Butonu sınıf değişkeni olarak sakla (manuel güncellemeler için)
            self.transfer_btn = transfer_btn

            # Stok Aktar butonu - başlangıçta inaktif
            self.create_order_btn = QPushButton("Stok Aktar")
            self.create_order_btn.setEnabled(False)
            self.create_order_btn.setStyleSheet("""
                QPushButton {
                    background-color: #bdc3c7;
                    color: #7f8c8d;
                    border: none;
                    padding: 10px 20px;
                    font-size: 14px;
                    font-weight: bold;
                    border-radius: 5px;
                    min-width: 120px;
                }
            """)
            self.create_order_btn.clicked.connect(self.create_order)
            buttons_layout.addWidget(self.create_order_btn)

            # Sipariş Aktar butonu - başlangıçta inaktif
            self.transfer_order_btn = QPushButton("Sipariş Aktar")
            self.transfer_order_btn.setEnabled(False)
            self.transfer_order_btn.setStyleSheet("""
                QPushButton {
                    background-color: #bdc3c7;
                    color: #7f8c8d;
                    border: none;
                    padding: 10px 20px;
                    font-size: 14px;
                    font-weight: bold;
                    border-radius: 5px;
                    min-width: 120px;
                }
            """)
            self.transfer_order_btn.clicked.connect(self.transfer_order)
            buttons_layout.addWidget(self.transfer_order_btn)

            # Kapat butonu
            close_btn = QPushButton("Kapat")
            close_btn.setStyleSheet("""
                QPushButton {
                    background-color: #e74c3c;
                    color: white;
                    border: none;
                    padding: 10px 20px;
                    font-size: 14px;
                    font-weight: bold;
                    border-radius: 5px;
                    min-width: 120px;
                }
                QPushButton:hover {
                    background-color: #c0392b;
                }
                QPushButton:pressed {
                    background-color: #a93226;
                }
            """)
            close_btn.clicked.connect(self.close)
            buttons_layout.addWidget(close_btn)
            
            bottom_layout.addLayout(buttons_layout)
            
            # Orta kısım boş
            bottom_layout.addStretch()
            
            # Sağ taraf: Yüzde hesaplama ve Toplam bilgileri
            right_section_layout = QHBoxLayout()
            
            # Yüzde hesaplama (sol taraf)
            percentage_layout = QVBoxLayout()
            percentage_layout.setSpacing(0)
            percentage_layout.setContentsMargins(0, 0, 10, 0)  # Sağdan 10px boşluk
            
            # Header değerini al ve yüzde hesapla
            header_value = 0
            if hasattr(self, 'contract_data') and hasattr(self.contract_data, 'ES_CONTRACT_INFO'):
                header_text_raw = self.safe_get(self.contract_data.ES_CONTRACT_INFO, 'HEADER_TEXT', '0')
                header_value = self.convert_turkish_to_integer(header_text_raw)

            genel_toplam = int(round(total_net + total_tax))
            if genel_toplam > 0 and header_value > 0:
                percentage = (1 - (header_value / genel_toplam)) * 100
            else:
                percentage = 0

            # Yüzde değerini tam sayıya yuvarla
            percentage_int = int(round(percentage))
            
            # Renk belirleme
            if percentage_int < 17:
                percentage_color = "#e74c3c"  # Kırmızı
            elif percentage_int >= 27:
                percentage_color = "#27ae60"  # Yeşil
            else:
                percentage_color = "#000000"  # Siyah (ara değer)
            
            # Yüzde label'ları oluştur (3 satır boşluk için)
            for i in range(3):
                if i == 1:  # Orta satırda yüzdeyi göster
                    percentage_label = QLabel(f"%{percentage_int}")
                    percentage_label.setStyleSheet(f"""
                        font-weight: bold; 
                        color: {percentage_color}; 
                        font-size: 32px; 
                        margin: 0; 
                        padding: 2px 2px 2px 8px;
                        border: 2px solid {percentage_color};
                        border-radius: 4px;
                        background-color: #ffffff;
                        min-width: 120px;
                        max-width: 120px;
                    """)
                    percentage_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                else:
                    # Boş label'lar hizalama için
                    percentage_label = QLabel("")
                    percentage_label.setMinimumHeight(30)  # Daha büyük font için yükseklik artırıldı
                
                percentage_layout.addWidget(percentage_label)
            
            # Toplam bilgileri (sağ taraf)
            total_layout = QVBoxLayout()
            total_layout.setSpacing(0)  # Satırlar arası hiç boşluk yok
            total_layout.setContentsMargins(0, 0, 0, 0)
            
            total_info = [
                ("Net Toplam:", f"{int(total_net):,} TL"),
                ("KDV Toplam:", f"{int(total_tax):,} TL"),
                ("Genel Toplam:", f"{int(total_net + total_tax):,} TL")
            ]
            
            for label_text, value in total_info:
                item_layout = QHBoxLayout()
                item_layout.setContentsMargins(0, 0, 0, 0)
                item_layout.setSpacing(5)
                
                label = QLabel(label_text)
                label.setMinimumWidth(100)
                label.setStyleSheet("font-weight: bold; color: #000000; font-size: 16px; margin: 0; padding: 0;")
                
                value_label = QLabel(str(value))
                value_label.setStyleSheet("color: #000000; font-weight: bold; font-size: 16px; margin: 0; padding: 0;")
                
                item_layout.addWidget(label)
                item_layout.addWidget(value_label)
                total_layout.addLayout(item_layout)
            
            # Yüzde ve toplam layout'larını birleştir
            right_section_layout.addLayout(percentage_layout)
            right_section_layout.addLayout(total_layout)
            
            bottom_layout.addLayout(right_section_layout)
            bottom_layout.addSpacing(10)  # Sağ taraftan biraz boşluk

            parent_layout.addLayout(bottom_layout)

            # Buton oluşturulduktan sonra durumu güncelle
            self.update_cari_aktar_button()

    def get_connection(self):
        """
        SQL Server veritabanı bağlantısı oluşturur.

        PRGsheet/Ayar sayfasındaki bağlantı bilgilerini kullanarak pyodbc ile
        veritabanı bağlantısı oluşturur.

        Returns:
            pyodbc.Connection: Veritabanı bağlantı nesnesi

        Raises:
            Exception: PRGsheet/Ayar sayfasında gerekli bilgiler eksikse

        Gerekli PRGsheet/Ayar Değişkenleri:
            SQL_SERVER, SQL_DATABASE, SQL_USERNAME, SQL_PASSWORD
        """
        # PRGsheet/Ayar'dan SQL bağlantı bilgilerini yükle
        config_manager = CentralConfigManager()
        settings = config_manager.get_settings()
        server = settings.get('SQL_SERVER')
        database = settings.get('SQL_DATABASE')
        username = settings.get('SQL_USERNAME')
        password = settings.get('SQL_PASSWORD')

        if not all([server, database, username, password]):
            raise Exception("PRGsheet/Ayar sayfasında SQL bağlantı bilgileri eksik")

        # Bağlantı dizesini oluşturun
        connection_string = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'

        # Veritabanına bağlanın
        return pyodbc.connect(connection_string)

    def check_existing_contract(self):
        """
        Mevcut sözleşmenin daha önce kaydedilip kaydedilmediğini kontrol eder.

        Sözleşme ID'sine göre SIPARISLER tablosunda kayıt olup olmadığını ve
        teslim edilmiş ürün olup olmadığını kontrol eder.

        Returns:
            tuple or None:
                (sip_evrakno_sira, has_teslim_miktar) varsa,
                None yoksa
        """
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Hem evrakno_sira hem de teslim_miktar kontrolü yap
            check_query = """
            SELECT sip_evrakno_sira,
                   CASE WHEN EXISTS(
                       SELECT 1 FROM [dbo].[SIPARISLER]
                       WHERE sip_belgeno = ? AND sip_teslim_miktar > 0
                   ) THEN 1 ELSE 0 END as has_teslim_miktar
            FROM [dbo].[SIPARISLER]
            WHERE sip_belgeno = ?
            GROUP BY sip_evrakno_sira
            """

            cursor.execute(check_query, (self.contract_id, self.contract_id))
            result = cursor.fetchone()

            conn.close()

            if result:
                return (result[0], bool(result[1]))  # (sip_evrakno_sira, has_teslim_miktar)
            else:
                return None  # Kayıt bulunamadı

        except Exception as e:
            print(f"Mevcut sözleşme kontrolünde hata: {str(e)}")
            return None

    def delete_existing_contract_records(self):
        """
        Mevcut sözleşmeye ait kayıtları güvenli şekilde siler.

        GÜVENLİK ÖNLEMİ: Sadece teslim edilmemiş (sip_teslim_miktar = 0)
        kayıtları siler. Teslim edilmiş kayıt varsa işlemi iptal eder.

        İşlem Adımları:
            1. Teslim_miktar > 0 olan kayıt sayısını kontrol et
            2. Teslim edilmiş kayıt varsa uyarı göster ve iptal et
            3. Güvenli ise sadece teslim_miktar = 0 olanları sil
        """
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Güvenlik kontrolü: Önce teslim_miktar > 0 olan kayıt var mı kontrol et
            safety_check_query = """
            SELECT COUNT(*) FROM [dbo].[SIPARISLER]
            WHERE sip_belgeno = ? AND sip_teslim_miktar > 0
            """

            cursor.execute(safety_check_query, (self.contract_id,))
            teslim_count = cursor.fetchone()[0]

            if teslim_count > 0:
                conn.close()
                QMessageBox.critical(
                    self,
                    "GÜVENLİK HATASI",
                    f"UYARI: Sözleşme {self.contract_id} içinde teslim edilmiş ürünler bulundu!\n\n"
                    f"Teslim edilmiş kayıt sayısı: {teslim_count}\n\n"
                    "Güvenlik nedeniyle silme işlemi iptal edildi.\n"
                    "Manuel kontrol gerekiyor!"
                )
                return

            # Güvenli silme işlemi - sadece teslim_miktar = 0 olanları sil
            delete_query = """
            DELETE FROM [dbo].[SIPARISLER]
            WHERE sip_belgeno = ? AND (sip_teslim_miktar = 0 OR sip_teslim_miktar IS NULL)
            """

            cursor.execute(delete_query, (self.contract_id,))
            deleted_count = cursor.rowcount
            conn.commit()
            conn.close()

        except Exception as e:
            logging.error(f"Mevcut kayıtları silerken hata: {str(e)}")
            QMessageBox.critical(self, "Hata", f"Mevcut kayıtları silerken hata oluştu: {str(e)}")

    def run_cari_query(self, query, params):
        """
        Cari hesap sorgularını çalıştırır ve sonuçları döner.

        Args:
            query (str): Çalıştırılacak SQL sorgusu
            params (list): Sorgu parametreleri

        Returns:
            list: Sözlük listesi olarak sonuçlar (her satır bir dict)

        Not:
            Hata durumunda boş liste döner ve kullanıcıya mesaj gösterir
        """
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            # Parametreleri doğru şekilde bind etmek için
            cursor.execute(query, params)

            # Tüm sonuçları al
            columns = [column[0] for column in cursor.description]
            records = [dict(zip(columns, row)) for row in cursor.fetchall()]

            conn.close()
            return records

        except Exception as e:
            QMessageBox.critical(self, "Sorgu Hatası", f"Veritabanı sorgusu sırasında hata oluştu: {str(e)}")
            return []

    def create_new_cari(self, contract_info):
        """
        Yeni cari hesap kaydı oluşturur.

        Sözleşme bilgilerinden müşteri verilerini alıp CARI_HESAPLAR ve
        CARI_HESAP_ADRESLERI tablolarına yeni kayıt ekler.

        Args:
            contract_info: Sözleşme bilgilerini içeren obje

        Cari Kod Formatı:
            - TCKN varsa: "340.{TCKN}"
            - TCKN yoksa: "340.{Telefon1}"

        İşlem Adımları:
            1. Müşteri bilgilerini sözleşmeden al
            2. Telefon numaralarını temizle (+90, 90 öneklerini kaldır)
            3. Cari kod oluştur
            4. Mevcut kayıt varsa güncelle, yoksa yeni ekle
            5. Adres kaydı oluştur
            6. Seçili cari kodu self.selected_cari_kod'a kaydet

        Not:
            - Telefon formatı: "+905321339827" -> "05321339827"
            - Adres 50 karakterden uzunsa bölünür
        """
        try:
            def safe_get(obj, attr, default=''):
                if not obj:
                    return default
                return getattr(obj, attr, default) if hasattr(obj, attr) else default

            # Müşteri bilgilerini al
            customer_name = f"{safe_get(contract_info, 'CUSTOMER_NAMEFIRST')} {safe_get(contract_info, 'CUSTOMER_NAMELAST')}".strip().upper()
            tckn = safe_get(contract_info, 'CUSTOMER_TAXNR')
            telefon1 = safe_get(contract_info, 'CUSTOMER_PHONE1')
            telefon2 = safe_get(contract_info, 'CUSTOMER_PHONE2')
            adres = safe_get(contract_info, 'CUSTOMER_ADDRESS').upper()
            sehir_ilce = safe_get(contract_info, 'CUSTOMER_CITY').upper()

            # TCKN'yi temizle (.0 kaldır)
            if str(tckn).endswith('.0'):
                tckn = str(tckn)[:-2]
            if str(tckn).lower() == 'nan':
                tckn = ''

            # Telefon numaralarını hazırla (None kontrolü ile) - create_new_cari
            telefon1_str = str(telefon1) if telefon1 is not None else ''
            if telefon1_str.endswith('.0'):
                telefon1_str = telefon1_str[:-2]

            # +90 ile başlıyorsa +90'ı kaldır, 90 ile başlıyorsa 90'ı kaldırıp başına 0 ekle
            if telefon1_str.startswith('+90'):
                new_telefon1 = '0' + telefon1_str[3:]  # +905321339827 -> 05321339827
            elif telefon1_str.startswith('90'):
                new_telefon1 = '0' + telefon1_str[2:]  # 905321339827 -> 05321339827
            else:
                new_telefon1 = telefon1_str

            if new_telefon1.lower() in ['nan', 'none']:
                new_telefon1 = ''

            telefon2_str = str(telefon2) if telefon2 is not None else ''
            if telefon2_str.endswith('.0'):
                telefon2_str = telefon2_str[:-2]

            # +90 ile başlıyorsa +90'ı kaldır, 90 ile başlıyorsa 90'ı kaldırıp başına 0 ekle
            if telefon2_str.startswith('+90'):
                new_telefon2 = '0' + telefon2_str[3:]  # +905321339827 -> 05321339827
            elif telefon2_str.startswith('90'):
                new_telefon2 = '0' + telefon2_str[2:]  # 905321339827 -> 05321339827
            else:
                new_telefon2 = telefon2_str

            if new_telefon2.lower() in ['nan', 'none']:
                new_telefon2 = ''

            # Adres formatı
            full_adres = f"{adres} {sehir_ilce}".strip()

            # Bağlantıyı aç
            conn = self.get_connection()
            cursor = conn.cursor()

            # Dinamik değerleri oluştur
            cari_kod = f"340.{tckn}" if tckn else f"340.{new_telefon1}"
            # Yeni oluşturulan cari kodunu self'e kaydet (sipariş oluşturma için)
            self.selected_cari_kod = cari_kod
            cari_vdaire_no = tckn
            cari_vdaire_adi = new_telefon2
            cari_CepTel = new_telefon1
            cari_unvan1 = customer_name

            # Cari kod zaten var mı kontrol et
            cursor.execute("SELECT COUNT(*) FROM CARI_HESAPLAR WHERE cari_kod = ?", [cari_kod])
            existing_count = cursor.fetchone()[0]

            if existing_count > 0:
                # Zaten var olan cari kaydını güncelle
                update_sql = """
                UPDATE CARI_HESAPLAR SET
                    cari_unvan1 = ?,
                    cari_CepTel = ?,
                    cari_vdaire_adi = ?,
                    cari_vdaire_no = ?,
                    cari_lastup_user = 1,
                    cari_lastup_date = GETDATE()
                WHERE cari_kod = ?
                """

                cursor.execute(update_sql, [
                    cari_unvan1,
                    cari_CepTel,
                    cari_vdaire_adi,
                    cari_vdaire_no,
                    cari_kod
                ])

                conn.commit()
                conn.close()

                QMessageBox.information(self, "Başarılı", f"Mevcut cari kaydı güncellendi:\n{cari_kod} - {cari_unvan1}")
                return

            # Son RECno'yu al
            cursor.execute("SELECT MAX(cari_RECid_RECno) FROM CARI_HESAPLAR")
            last_recno = cursor.fetchone()[0] or 0
            new_recno = last_recno + 1

            # CARI_HESAPLAR tablosuna ekleme
            insert_cari_sql = """
            INSERT INTO CARI_HESAPLAR (
                cari_RECid_DBCno, cari_RECid_RECno, cari_SpecRECno,
                cari_iptal, cari_fileid, cari_hidden, cari_kilitli,
                cari_degisti, cari_checksum, cari_create_user,
                cari_create_date, cari_lastup_user, cari_lastup_date,
                cari_kod, cari_unvan1,
                cari_vdaire_adi, cari_vdaire_no,
                cari_CepTel, cari_satis_fk,
                cari_fatura_adres_no, cari_sevk_adres_no,
                cari_EftHesapNum
            ) VALUES (
                0, ?, 0,
                0, 31, 0, 0,
                0, 0, 1,
                GETDATE(), 1, GETDATE(),
                ?, ?,
                ?, ?,
                ?, 1,
                1, 1,
                1
            )
            """

            cursor.execute(insert_cari_sql, [
                new_recno,
                cari_kod, cari_unvan1,
                cari_vdaire_adi, cari_vdaire_no,
                cari_CepTel
            ])

            # Adres RECno'sunu al
            cursor.execute("SELECT MAX(adr_RECid_RECno) FROM CARI_HESAP_ADRESLERI")
            last_adr_recno = cursor.fetchone()[0] or 0
            new_adr_recno = last_adr_recno + 1

            # Adres bilgilerini hazırla (PRG.py formatında)
            full_adres = f"{adres} " if adres else ''
            if len(full_adres) > 50:
                teslimat_adres = full_adres[:50].strip()  # İlk 50 karakter
                teslimat_adres2 = full_adres[50:100].strip()  # Kalan kısım (max 50 karakter)
            else:
                teslimat_adres = full_adres.strip()
                teslimat_adres2 = ''

            teslimat_ilce = ''    # Boş bırak
            teslimat_il = sehir_ilce[:30] if sehir_ilce else ''  # adr_il için limit

            # Bu cari kod için mevcut en yüksek adres numarasını bul
            cursor.execute("SELECT ISNULL(MAX(adr_adres_no), 0) FROM CARI_HESAP_ADRESLERI WHERE adr_cari_kod = ?", [cari_kod])
            max_adres_no = cursor.fetchone()[0]
            new_adres_no = max_adres_no + 1

            # Adres kaydı SQL
            insert_adres_sql = """
            INSERT INTO CARI_HESAP_ADRESLERI (
                adr_RECid_DBCno, adr_RECid_RECno, adr_SpecRECno,
                adr_iptal, adr_fileid, adr_hidden, adr_kilitli,
                adr_degisti, adr_checksum, adr_create_user,
                adr_create_date, adr_lastup_user, adr_lastup_date,
                adr_cari_kod, adr_adres_no,
                adr_cadde, adr_sokak,
                adr_ilce, adr_il,
                adr_aprint_fl, adr_Adres_kodu
            ) VALUES (
                0, ?, 0,
                0, 0, 0, 0,
                0, 0, 1,
                GETDATE(), 1, GETDATE(),
                ?, ?,
                ?, ?,
                ?, ?,
                0, 'ADRS1'
            )
            """

            cursor.execute(insert_adres_sql, [
                new_adr_recno,
                cari_kod,
                new_adres_no,
                teslimat_adres,
                teslimat_adres2,
                teslimat_ilce,
                teslimat_il
            ])

            conn.commit()
            conn.close()

            QMessageBox.information(self, "Başarılı", f"Yeni cari kaydı oluşturuldu:\n{cari_kod} - {cari_unvan1}")

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Yeni cari kaydı oluşturulurken hata oluştu: {str(e)}")

    def update_cari_record(self, selected_record, contract_info):
        """
        Seçili cari kaydını günceller.

        Mevcut bir cari hesabın bilgilerini sözleşme bilgilerine göre günceller.

        Args:
            selected_record (dict): Seçili cari kayıt bilgileri (cariKod içermeli)
            contract_info: Sözleşme bilgilerini içeren obje

        Güncellenen Alanlar:
            - cari_unvan1: Müşteri adı soyadı (büyük harf)
            - cari_CepTel: Birinci telefon numarası
            - cari_vdaire_adi: İkinci telefon numarası
            - cari_vdaire_no: TCKN
            - cari_lastup_user: 1
            - cari_lastup_date: Güncel tarih

        Not:
            - Seçili cari kodu self.selected_cari_kod'a kaydedilir
            - Telefon numaraları temizlenir (+90, 90 önekleri kaldırılır)
        """
        try:
            # Selected cari kodunu self'e kaydet (sipariş oluşturma için)
            self.selected_cari_kod = selected_record['cariKod']
            def safe_get(obj, attr, default=''):
                if not obj:
                    return default
                return getattr(obj, attr, default) if hasattr(obj, attr) else default

            # Müşteri bilgilerini al
            customer_name = f"{safe_get(contract_info, 'CUSTOMER_NAMEFIRST')} {safe_get(contract_info, 'CUSTOMER_NAMELAST')}".strip().upper()
            tckn = safe_get(contract_info, 'CUSTOMER_TAXNR')
            telefon1 = safe_get(contract_info, 'CUSTOMER_PHONE1')
            telefon2 = safe_get(contract_info, 'CUSTOMER_PHONE2')

            # TCKN'yi temizle
            if str(tckn).endswith('.0'):
                tckn = str(tckn)[:-2]
            if str(tckn).lower() == 'nan':
                tckn = ''

            # Telefon numaralarını hazırla (None kontrolü ile) - update_cari_record
            telefon1_str = str(telefon1) if telefon1 is not None else ''
            if telefon1_str.endswith('.0'):
                telefon1_str = telefon1_str[:-2]

            # +90 ile başlıyorsa +90'ı kaldır, 90 ile başlıyorsa 90'ı kaldırıp başına 0 ekle
            if telefon1_str.startswith('+90'):
                new_telefon1 = '0' + telefon1_str[3:]  
            elif telefon1_str.startswith('90'):
                new_telefon1 = '0' + telefon1_str[2:]  
            else:
                new_telefon1 = telefon1_str

            if new_telefon1.lower() in ['nan', 'none']:
                new_telefon1 = ''

            telefon2_str = str(telefon2) if telefon2 is not None else ''
            if telefon2_str.endswith('.0'):
                telefon2_str = telefon2_str[:-2]

            # +90 ile başlıyorsa +90'ı kaldır, 90 ile başlıyorsa 90'ı kaldırıp başına 0 ekle
            if telefon2_str.startswith('+90'):
                new_telefon2 = '0' + telefon2_str[3:] 
            elif telefon2_str.startswith('90'):
                new_telefon2 = '0' + telefon2_str[2:]  
            else:
                new_telefon2 = telefon2_str

            if new_telefon2.lower() in ['nan', 'none']:
                new_telefon2 = ''

            conn = self.get_connection()
            cursor = conn.cursor()

            # Cari kaydını güncelle (TCKN dahil)
            update_sql = """
            UPDATE CARI_HESAPLAR SET
                cari_unvan1 = ?,
                cari_CepTel = ?,
                cari_vdaire_adi = ?,
                cari_vdaire_no = ?,
                cari_lastup_user = 1,
                cari_lastup_date = GETDATE()
            WHERE cari_kod = ?
            """

            cursor.execute(update_sql, [
                customer_name,
                new_telefon1,
                new_telefon2,
                tckn,
                selected_record['cariKod']
            ])

            conn.commit()
            conn.close()

            QMessageBox.information(self, "Başarılı", f"Cari kaydı güncellendi:\n{selected_record['cariKod']} - {customer_name}")

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Cari kaydı güncellenirken hata oluştu: {str(e)}")

    def update_cari_adres(self, cari_kod, contract_info, selected_recno=None):
        """
        Cari hesabın adres kaydını günceller veya yeni oluşturur.

        Args:
            cari_kod (str): Cari hesap kodu
            contract_info: Sözleşme bilgilerini içeren obje
            selected_recno (int, optional): Güncellenecek adres kayıt numarası.
                                           None ise yeni adres oluşturur.

        Adres Formatı:
            - adr_cadde: İlk 50 karakter (adres birinci kısım)
            - adr_sokak: Kalan kısım (max 50 karakter)
            - adr_ilce: Boş bırakılır
            - adr_il: Şehir bilgisi (max 30 karakter)

        İşlem Mantığı:
            - selected_recno varsa: Mevcut adres güncellenir
            - selected_recno None ise: Yeni adres kaydı oluşturulur
            - Adres 50 karakterden uzunsa 2 alana bölünür
        """
        try:
            def safe_get(obj, attr, default=''):
                if not obj:
                    return default
                return getattr(obj, attr, default) if hasattr(obj, attr) else default

            # Adres bilgilerini hazırla (PRG.py formatında)
            adres = safe_get(contract_info, 'CUSTOMER_ADDRESS', '').upper()
            sehir_ilce = safe_get(contract_info, 'CUSTOMER_CITY', '').upper()

            # Adres verilerini böl: 50 karakterden uzunsa böl
            full_adres = f"{adres} " if adres else ''
            if len(full_adres) > 50:
                teslimat_adres = full_adres[:50].strip()  # İlk 50 karakter
                teslimat_adres2 = full_adres[50:100].strip()  # Kalan kısım (max 50 karakter)
            else:
                teslimat_adres = full_adres.strip()
                teslimat_adres2 = ''

            teslimat_ilce = ''    # Boş bırak
            teslimat_il = sehir_ilce[:30] if sehir_ilce else ''  # adr_il için limit

            conn = self.get_connection()
            cursor = conn.cursor()

            if selected_recno:
                # Mevcut adres kaydını güncelle (PRG.py formatında)
                update_adres_sql = """
                UPDATE CARI_HESAP_ADRESLERI SET
                    adr_cadde = ?,
                    adr_sokak = ?,
                    adr_ilce = ?,
                    adr_il = ?,
                    adr_lastup_user = 1,
                    adr_lastup_date = GETDATE()
                WHERE adr_RECno = ?
                """

                cursor.execute(update_adres_sql, [
                    teslimat_adres,
                    teslimat_adres2,
                    teslimat_ilce,
                    teslimat_il,
                    selected_recno
                ])

            else:
                # Yeni adres kaydı oluştur
                cursor.execute("SELECT MAX(adr_RECid_RECno) FROM CARI_HESAP_ADRESLERI")
                last_adr_recno = cursor.fetchone()[0] or 0
                new_adr_recno = last_adr_recno + 1

                # Bu cari kod için mevcut en yüksek adres numarasını bul
                cursor.execute("SELECT ISNULL(MAX(adr_adres_no), 0) FROM CARI_HESAP_ADRESLERI WHERE adr_cari_kod = ?", [cari_kod])
                max_adres_no = cursor.fetchone()[0]
                new_adres_no = max_adres_no + 1

                insert_adres_sql = """
                INSERT INTO CARI_HESAP_ADRESLERI (
                    adr_RECid_DBCno, adr_RECid_RECno, adr_SpecRECno,
                    adr_iptal, adr_fileid, adr_hidden, adr_kilitli,
                    adr_degisti, adr_checksum, adr_create_user,
                    adr_create_date, adr_lastup_user, adr_lastup_date,
                    adr_cari_kod, adr_adres_no,
                    adr_cadde, adr_sokak,
                    adr_ilce, adr_il,
                    adr_aprint_fl, adr_Adres_kodu
                ) VALUES (
                    0, ?, 0,
                    0, 0, 0, 0,
                    0, 0, 1,
                    GETDATE(), 1, GETDATE(),
                    ?, ?,
                    ?, ?,
                    ?, ?,
                    0, 'ADRS1'
                )
                """

                cursor.execute(insert_adres_sql, [
                    new_adr_recno,
                    cari_kod,
                    new_adres_no,
                    teslimat_adres,
                    teslimat_adres2,
                    teslimat_ilce,
                    teslimat_il
                ])

            conn.commit()
            conn.close()

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Adres güncellenirken hata oluştu: {str(e)}")

    def transfer_order_to_cari(self):
        """
        Siparişi cari hesaba aktarma işlemini başlatır.

        Bu metod, "Cari Aktar" butonuna basıldığında çalışır ve müşteri bilgilerine
        göre cari hesap arama/seçme/oluşturma işlemlerini yönetir.

        İşlem Akışı:
            1. Müşteri bilgilerini sözleşmeden al
            2. Telefon numaralarını temizle ve formatla
            3. Üç aşamalı arama yap:
               a) TCKN ile 340.TCKN formatında ara
               b) Bulunamazsa TCKN ile vdaire_no alanında ara
               c) Bulunamazsa telefon numarası ile ara
            4. Bulunan kayıtları CariSelectionDialog'da göster
            5. Kullanıcı seçimine göre:
               - Güncelle: Mevcut cariyi güncelle
               - Yeni Oluştur: Yeni cari oluştur
               - İptal: İşlemi iptal et

        Telefon Formatı:
            "+905321339827" veya "905321339827" -> "05321339827"

        Not:
            - TCKN eşleşmesi varsa öncelikli olarak gösterilir
            - Kayıt bulunamazsa doğrudan yeni oluşturma dialog'ı açılır
        """
        try:
            if not hasattr(self.contract_data, 'ES_CONTRACT_INFO'):
                QMessageBox.warning(self, "Uyarı", "Sözleşme bilgileri alınamadı.")
                return

            contract_info = self.contract_data.ES_CONTRACT_INFO

            def safe_get(obj, attr, default=''):
                if not obj:
                    return default
                return getattr(obj, attr, default) if hasattr(obj, attr) else default

            # Müşteri bilgilerini hazırla
            customer_name = f"{safe_get(contract_info, 'CUSTOMER_NAMEFIRST')} {safe_get(contract_info, 'CUSTOMER_NAMELAST')}".strip()
            tckn = safe_get(contract_info, 'CUSTOMER_TAXNR')
            telefon1 = safe_get(contract_info, 'CUSTOMER_PHONE1')
            telefon2 = safe_get(contract_info, 'CUSTOMER_PHONE2')
            adres = safe_get(contract_info, 'CUSTOMER_ADDRESS')
            sehir_ilce = safe_get(contract_info, 'CUSTOMER_CITY')

            # TCKN'yi temizle
            if str(tckn).endswith('.0'):
                tckn = str(tckn)[:-2]
            if str(tckn).lower() == 'nan':
                tckn = ''

            # Telefon numaralarını temizle (None kontrolü ile) - transfer_order
            telefon1_str = str(telefon1) if telefon1 is not None else ''
            if telefon1_str.endswith('.0'):
                telefon1_str = telefon1_str[:-2]

            # +90 ile başlıyorsa +90'ı kaldır, 90 ile başlıyorsa 90'ı kaldırıp başına 0 ekle
            if telefon1_str.startswith('+90'):
                new_telefon1 = '0' + telefon1_str[3:]  # +905321339827 -> 05321339827
            elif telefon1_str.startswith('90'):
                new_telefon1 = '0' + telefon1_str[2:]  # 905321339827 -> 05321339827
            else:
                new_telefon1 = telefon1_str

            if new_telefon1.lower() in ['nan', 'none']:
                new_telefon1 = ''

            telefon2_str = str(telefon2) if telefon2 is not None else ''
            if telefon2_str.endswith('.0'):
                telefon2_str = telefon2_str[:-2]

            # +90 ile başlıyorsa +90'ı kaldır, 90 ile başlıyorsa 90'ı kaldırıp başına 0 ekle
            if telefon2_str.startswith('+90'):
                new_telefon2 = '0' + telefon2_str[3:]  # +905321339827 -> 05321339827
            elif telefon2_str.startswith('90'):
                new_telefon2 = '0' + telefon2_str[2:]  # 905321339827 -> 05321339827
            else:
                new_telefon2 = telefon2_str

            if new_telefon2.lower() in ['nan', 'none']:
                new_telefon2 = ''

            # Telefon formatı
            telefon_format = new_telefon1
            if new_telefon2:
                telefon_format = f"{new_telefon1}   -   {new_telefon2}"

            # Müşteri adını büyük harf yap
            new_cari_adi = customer_name.upper()

            # Adres formatını hazırla
            adres_parts = []
            if adres and str(adres) != 'N/A':
                adres_parts.append(adres.upper())

            full_adres = ' '.join(adres_parts)
            if sehir_ilce and str(sehir_ilce) != 'N/A':
                if full_adres:
                    full_adres += ' '
                full_adres += sehir_ilce.upper()

            # Müşteri bilgilerini hazırla
            customer_info = {
                'ad': new_cari_adi,
                'tckn': tckn,
                'telefon': telefon_format,
                'adres': full_adres if full_adres else 'Adres bilgisi yok'
            }

            # Veritabanı sorguları
            records = []
            message = ""
            tckn_matched = False

            # 1. Aşama: TCKN ile sorgula
            if tckn:
                records = self.run_cari_query(
                    """SELECT TOP 100 PERCENT
                        c.cari_RECno AS sayac,
                        c.cari_unvan1 AS cariAdi,
                        c.cari_kod AS cariKod,
                        c.cari_vdaire_no AS TCKN,
                        c.cari_CepTel AS cariTelefon,
                        c.cari_unvan2 AS cariAciklama,
                        c.cari_vdaire_adi AS Telefon2,
                        a.adr_RECno AS adres_sayac,
                        CONCAT(COALESCE(a.adr_cadde, ''), ' ', COALESCE(a.adr_sokak, ''), ' ', COALESCE(a.adr_ilce, ''), ' ', COALESCE(a.adr_il, '')) AS Adres,
                        CASE WHEN a.adr_adres_no IS NOT NULL THEN CONCAT(CAST(a.adr_adres_no AS VARCHAR), ' adet adres kaydı') ELSE '' END AS [Adres Bilgisi]
                        FROM dbo.CARI_HESAPLAR c WITH (NOLOCK)
                        LEFT JOIN dbo.CARI_HESAP_ADRESLERI a WITH (NOLOCK) ON c.cari_kod = a.adr_cari_kod
                        WHERE c.cari_kod LIKE N'%' + ? + '%' ORDER BY c.cari_RECno ASC, a.adr_adres_no ASC""",
                        [f"%{tckn}%"]
                )
                message = "340.TCKN ile eşleşen kayıtlar"
                tckn_matched = bool(records)

                # Eğer TCKN ile kayıt bulunamadıysa, vdaire no ile sorgula
                if not records and tckn:
                    records = self.run_cari_query(
                    """SELECT TOP 100 PERCENT
                        c.cari_RECno AS sayac,
                        c.cari_unvan1 AS cariAdi,
                        c.cari_kod AS cariKod,
                        c.cari_vdaire_no AS TCKN,
                        c.cari_CepTel AS cariTelefon,
                        c.cari_unvan2 AS cariAciklama,
                        c.cari_vdaire_adi AS Telefon2,
                        a.adr_RECno AS adres_sayac,
                        CONCAT(COALESCE(a.adr_cadde, ''), ' ', COALESCE(a.adr_sokak, ''), ' ', COALESCE(a.adr_ilce, ''), ' ', COALESCE(a.adr_il, '')) AS Adres,
                        CASE WHEN a.adr_adres_no IS NOT NULL THEN CONCAT(CAST(a.adr_adres_no AS VARCHAR), ' adet adres kaydı') ELSE '' END AS [Adres Bilgisi]
                        FROM dbo.CARI_HESAPLAR c WITH (NOLOCK)
                        LEFT JOIN dbo.CARI_HESAP_ADRESLERI a WITH (NOLOCK) ON c.cari_kod = a.adr_cari_kod
                        WHERE c.cari_vdaire_no LIKE N'%' + ? + '%' ORDER BY c.cari_RECno ASC, a.adr_adres_no ASC""",
                        [f"%{tckn}%"]
                    )
                    message = "340.TCKN ile eşleşme bulunamadı. TCKN No ile eşleşen kayıtlar"

            # 2. Aşama: Telefon ile sorgula
            if not records and new_telefon1:
                records = self.run_cari_query(
                    """SELECT TOP 100 PERCENT
                        c.cari_RECno AS sayac,
                        c.cari_unvan1 AS cariAdi,
                        c.cari_kod AS cariKod,
                        c.cari_vdaire_no AS TCKN,
                        c.cari_CepTel AS cariTelefon,
                        c.cari_unvan2 AS cariAciklama,
                        c.cari_vdaire_adi AS Telefon2,
                        a.adr_RECno AS adres_sayac,
                        CONCAT(COALESCE(a.adr_cadde, ''), ' ', COALESCE(a.adr_sokak, ''), ' ', COALESCE(a.adr_ilce, ''), ' ', COALESCE(a.adr_il, '')) AS Adres,
                        CASE WHEN a.adr_adres_no IS NOT NULL THEN CONCAT(CAST(a.adr_adres_no AS VARCHAR), ' adet adres kaydı') ELSE '' END AS [Adres Bilgisi]
                        FROM dbo.CARI_HESAPLAR c WITH (NOLOCK)
                        LEFT JOIN dbo.CARI_HESAP_ADRESLERI a WITH (NOLOCK) ON c.cari_kod = a.adr_cari_kod
                        WHERE c.cari_CepTel LIKE N'%' + ? + '%' ORDER BY c.cari_RECno ASC, a.adr_adres_no ASC""",
                    [f"%{new_telefon1}%"]
                )
                message = "TCKN ile eşleşen kayıt bulunamadı. Telefon ile eşleşen kayıtlar"

            # 3. Aşama: İsim ile sorgula
            if not records and new_cari_adi:
                records = self.run_cari_query(
                    """SELECT TOP 100 PERCENT
                        c.cari_RECno AS sayac,
                        c.cari_unvan1 AS cariAdi,
                        c.cari_kod AS cariKod,
                        c.cari_vdaire_no AS TCKN,
                        c.cari_CepTel AS cariTelefon,
                        c.cari_unvan2 AS cariAciklama,
                        c.cari_vdaire_adi AS Telefon2,
                        a.adr_RECno AS adres_sayac,
                        CONCAT(COALESCE(a.adr_cadde, ''), ' ', COALESCE(a.adr_sokak, ''), ' ', COALESCE(a.adr_ilce, ''), ' ', COALESCE(a.adr_il, '')) AS Adres,
                        CASE WHEN a.adr_adres_no IS NOT NULL THEN CONCAT(CAST(a.adr_adres_no AS VARCHAR), ' adet adres kaydı') ELSE '' END AS [Adres Bilgisi]
                        FROM dbo.CARI_HESAPLAR c WITH (NOLOCK)
                        LEFT JOIN dbo.CARI_HESAP_ADRESLERI a WITH (NOLOCK) ON c.cari_kod = a.adr_cari_kod
                        WHERE c.cari_unvan1 LIKE N'%' + ? + '%' ORDER BY c.cari_RECno ASC, a.adr_adres_no ASC""",
                    [new_cari_adi]
                )
                message = "TCKN ve Telefon ile kayıt bulunamadı. İsim ile eşleşen kayıtlar"
                tckn_matched = False

            # Kayıt bulunduysa direkt CariSelectionDialog göster
            if records:
                dialog = CariSelectionDialog(records, message, self, customer_info, tckn_matched)
                result = dialog.exec_()

                if result == QDialog.Accepted:
                    if dialog.action == "update":
                        selected_index = dialog.selected_record()
                        if selected_index is not None:
                            selected_record = records[selected_index]
                            # Cari bilgilerini güncelle
                            self.update_cari_record(selected_record, contract_info)
                            # Seçili adres bilgilerini de güncelle
                            selected_address_recno = dialog.get_selected_address_recno()
                            self.update_cari_adres(selected_record['cariKod'], contract_info, selected_address_recno)

                    elif dialog.action == "new":
                        self.create_new_cari(contract_info)
                elif result == QDialog.Rejected:
                    # "Kapat" butonuna basıldı - seçili satırın cari kodunu kullan
                    selected_index = dialog.selected_record()
                    if selected_index is not None:
                        selected_record = records[selected_index]
                        # Seçili cari kodunu set et (sipariş oluşturma için)
                        self.selected_cari_kod = selected_record['cariKod']
                    else:
                        # Hiç seçim yoksa ilk kaydı kullan (varsayılan)
                        if records:
                            self.selected_cari_kod = records[0]['cariKod']
            else:
                # Hiç kayıt bulunamadı, MusteriBilgileriDialog göster
                dialog = MusteriBilgileriDialog(customer_info, self)
                result = dialog.exec_()

                if result == QDialog.Accepted and dialog.action == "new":
                    self.create_new_cari(contract_info)

            # Cari aktar işlemi tamamlandı, Stok Aktar butonunu aktif hale getir
            if hasattr(self, 'create_order_btn'):
                self.create_order_btn.setEnabled(True)
                self.create_order_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #27ae60;
                        color: white;
                        border: none;
                        padding: 10px 20px;
                        font-size: 14px;
                        font-weight: bold;
                        border-radius: 5px;
                        min-width: 120px;
                    }
                    QPushButton:hover {
                        background-color: #229954;
                    }
                    QPushButton:pressed {
                        background-color: #1e8449;
                    }
                """)

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Cari aktar işleminde hata oluştu: {str(e)}")

    def find_products_table(self):
        """
        ÜRÜNLER tablosunu bulur ve döner.

        Bu metod, ana pencere içindeki widget hiyerarşisinde rekursif olarak
        arama yaparak ürünler tablosunu (QTableWidget) bulur.

        Returns:
            QTableWidget or None: Bulunan ürünler tablosu, bulunamazsa None

        Tanımlama Kriterı:
            - QTableWidget tipinde olmalı
            - En az 10 sütun içermeli (Satır, SAP Kodu, Malzeme Adı, SPEC, vb.)

        Not:
            Hata durumunda None döner
        """
        try:
            # Ana widget içindeki tüm QTableWidget'ları ara
            central_widget = self.centralWidget()

            def search_table(widget):
                if isinstance(widget, QTableWidget):
                    # ÜRÜNLER tablosu 10 sütunlu olmalı (Satir, SAP Kodu, Malzeme Adı, SPEC, Miktar, Birim Fiyat, Net Tutar, KDV, Sipariş No, Sip Kalem No)
                    if widget.columnCount() >= 10:
                        return widget

                # Alt widget'larda ara
                for child in widget.children():
                    if hasattr(child, 'children'):
                        result = search_table(child)
                        if result:
                            return result
                return None

            result = search_table(central_widget)
            return result

        except Exception as e:
            return None

    def collect_sap_codes(self, products_table):
        """
        ÜRÜNLER tablosundan SAP kodlarını toplar.

        Tablodaki tüm satırlardan SAP kodlarını çıkarır ve
        tekrarlı değerleri kaldırır.

        Args:
            products_table (QTableWidget): Ürünler tablosu

        Returns:
            list: Benzersiz SAP kodları listesi

        Not:
            - SAP Kodu sütunu indeks 1'de bulunur
            - Boş değerler filtrelenir
            - Dublicate'ler kaldırılır (set kullanılarak)
        """
        try:
            sap_codes = []
            for row in range(products_table.rowCount()):
                sap_kod_item = products_table.item(row, 1)  # SAP Kodu sütunu (indeks 1)
                if sap_kod_item:
                    sap_kod = sap_kod_item.text().strip()
                    if sap_kod and sap_kod != "":
                        sap_codes.append(sap_kod)

            return list(set(sap_codes))  # Dublicate'leri kaldır

        except Exception as e:
            return []

    def process_material_control(self, sap_codes, products_table):
        """
        Ana malzeme kontrol işlemini yürütür.

        Bu metod, SAP kodlarına göre stok kartı kontrolü yapar ve gerektiğinde
        yeni stok kartı oluşturur veya pasif stok kartlarını aktif hale getirir.

        Args:
            sap_codes (list): Kontrol edilecek SAP kodları listesi
            products_table (QTableWidget): Ürünler tablosu

        İşlem Adımları:
            1. Her satır için SAP kodu ve SPEC bilgisini al
            2. Stok kodu oluştur (SAP-0 veya SAP-SPEC formatında)
            3. STOKLAR_CHOOSE_3A'da aktif kayıt var mı kontrol et
            4. Aktif kayıt yoksa STOKLAR_CHOOSE_2A'da pasif kayıt var mı kontrol et
            5. Pasif kayıt varsa aktif hale getir
            6. Hiçbir kayıt yoksa yeni stok kartı oluştur
            7. Sonuçları rapor et (oluşturulan, aktifleştirilen, hatalar)

        Stok Kodu Formatı:
            - SPEC boşsa: "SAP-0"
            - SPEC doluysa: "SAP-SPEC" veya özel format

        Returns:
            Kullanıcıya mesaj kutuları ile sonuç bilgisi verir

        Not:
            - Aynı SAP kodu farklı SPEC'lerle tekrar edebilir
            - Her benzersiz kombinasyon (SAP#SPEC#KALEM) ayrı işlenir
            - Malzeme Kodu sütunu otomatik olarak tabloya eklenir
        """
        try:
            conn = self.get_connection()
            cursor = conn.cursor()

            processed_count = 0
            created_count = 0
            activated_count = 0
            activated_list = []  # Pasif'ten aktif yapılan stok kartları
            created_list = []    # Yeni oluşturulan stok kartları
            error_list = []      # Hata mesajları
            sap_to_stok_map = {}  # SAP kodu -> Stok kodu eşleştirmesi (multiple values per SAP)

            # Her satırı ayrı ayrı işle (SAP kodu tekrarı olsa bile)
            processed_rows = set()  # İşlenmiş satırları takip et

            for row_index in range(products_table.rowCount()):
                try:
                    # Bu satırın SAP kodunu al
                    sap_item = products_table.item(row_index, 1)  # SAP Kodu sütunu
                    if not sap_item:
                        continue

                    sap_kod = sap_item.text().strip()
                    if sap_kod not in sap_codes:
                        continue  # Bu SAP kodu seçili değil

                    # Bu satırı daha önce işledik mi kontrol et
                    spec_item = products_table.item(row_index, 3)  # SPEC sütunu
                    spec_value = spec_item.text().strip() if spec_item else ''
                    kalem_item = products_table.item(row_index, 9)  # Sip Kalem No sütunu
                    kalem_no = kalem_item.text().strip() if kalem_item else ''

                    row_signature = f"{sap_kod}#{spec_value}#{kalem_no}"
                    if row_signature in processed_rows:
                        continue
                    processed_rows.add(row_signature)

                    # print(f"PROCESSING ROW {row_index}: SAP={sap_kod}, SPEC='{spec_value}', KALEM='{kalem_no}'")

                    # Bu satırın verilerini al
                    row_data = self.extract_row_data_by_index(products_table, row_index, has_malzeme_kodu=False)
                    if not row_data:
                        continue

                    # SPEC değerine göre stok kodunu belirle
                    spec_value = row_data.get('SPEC', '').strip()
                    if not spec_value or spec_value == '':
                        # SPEC boş ise SAP-0 formatında kontrol et
                        stok_kod = f"{sap_kod}-0"
                    else:
                        # SPEC dolu ise özel stok kodu oluştur (ama sadece kontrol için)
                        stok_kod, _ = self.generate_stock_code_and_name(sap_kod, row_data)

                    # 1. Aşama: STOKLAR_CHOOSE_3A'da aktif kayıt var mı kontrol et
                    cursor.execute("SELECT * FROM STOKLAR_CHOOSE_3A WHERE msg_S_0870 = ?", (stok_kod,))
                    result = cursor.fetchone()

                    if result:
                        processed_count += 1
                        # Aktif kayıt var, mapping'e ekle
                        if sap_kod not in sap_to_stok_map:
                            sap_to_stok_map[sap_kod] = []
                        sap_to_stok_map[sap_kod].append(stok_kod)
                        continue  # Aktif kayıt var, bir şey yapmaya gerek yok

                    # 2. Aşama: Aktif kayıt yoksa, STOKLAR_CHOOSE_2A'da pasif kayıt var mı kontrol et
                    cursor.execute("SELECT * FROM STOKLAR_CHOOSE_2A WHERE msg_S_0078 = ?", (stok_kod,))
                    result = cursor.fetchone()

                    if result:
                        # Pasif kayıt var, aktif hale getir
                        update_query = "UPDATE STOKLAR SET sto_pasif_fl = 0 WHERE sto_kod = ?"
                        cursor.execute(update_query, (stok_kod,))
                        conn.commit()
                        activated_count += 1
                        # Stok adını al ve listeye ekle
                        stok_adi = self.get_stok_name(cursor, stok_kod)
                        activated_list.append(f"{stok_kod} - {stok_adi}")
                        # SAP kodu -> Stok kodu eşleştirmesi
                        if sap_kod not in sap_to_stok_map:
                            sap_to_stok_map[sap_kod] = []
                        sap_to_stok_map[sap_kod].append(stok_kod)
                    else:
                        # 3. Aşama: Hiç kayıt yok, yeni stok kartı oluştur
                        result = self.handle_new_stok_karti(sap_kod, row_data)
                        if result is not None:
                            success, message = result
                        else:
                            success, message = False, f"SAP {sap_kod}: Stok kartı oluşturma işlemi başarısız"
                        if success:
                            created_count += 1
                            created_list.append(message)
                            # Stok kodunu mesajdan çıkar ve mapping'e ekle (ilk satır stok kodu)
                            if "\n" in message:
                                stok_kod_from_message = message.split("\n")[0].strip()
                                if sap_kod not in sap_to_stok_map:
                                    sap_to_stok_map[sap_kod] = []
                                sap_to_stok_map[sap_kod].append(stok_kod_from_message)
                        else:
                            # Hata durumunda da, eğer "zaten mevcut" mesajı varsa mapping'e ekle
                            if "zaten mevcut" in message or "Aktif stok kartı" in message:
                                # Mevcut stok kartını mapping'e ekle
                                if sap_kod not in sap_to_stok_map:
                                    sap_to_stok_map[sap_kod] = []
                                sap_to_stok_map[sap_kod].append(stok_kod)
                            # Hata mesajını listeye ekle
                            error_list.append(message)

                    processed_count += 1

                except Exception as e:
                    # Sadece log dosyasına yaz, console'a yazdırma
                    import logging
                    logging.getLogger().setLevel(logging.CRITICAL)  # ERROR seviyesini gizle
                    error_list.append(f"SAP {sap_kod}: {str(e)}")
                    continue

            conn.close()

            result = {
                'success': True,
                'processed': processed_count,
                'activated': activated_count,
                'created': created_count,
                'activated_list': activated_list,
                'created_list': created_list,
                'error_list': error_list,
                'sap_to_stok_map': sap_to_stok_map
            }

            return result

        except Exception as e:
            return {'success': False}

    def get_row_data_for_sap_kod(self, products_table, sap_kod):
        """SAP kodu için tablo satır verilerini döndürür"""
        try:
            # SAP Kodu sütununun indeksini belirle
            sap_col_index = 1  # Varsayılan (Malzeme Kodu sütunu eklenmeden önce)

            # Eğer "Malzeme Kodu" sütunu varsa ve SAP Kodu sütunu kaldırıldıysa, malzeme kodundan SAP kodunu bul
            if products_table.columnCount() == 10 and products_table.horizontalHeaderItem(1) and products_table.horizontalHeaderItem(1).text() == "Malzeme Kodu":
                # SAP Kodu sütunu kaldırılmış, Malzeme Kodu'ndan SAP kodunu çıkar
                for row in range(products_table.rowCount()):
                    malzeme_kod_item = products_table.item(row, 1)  # Malzeme Kodu sütunu
                    if malzeme_kod_item:
                        malzeme_kod = malzeme_kod_item.text().strip()
                        # Malzeme kodundan SAP kodunu çıkar (örn: "3120024058-0" -> "3120024058")
                        if malzeme_kod.endswith("-0"):
                            extracted_sap_kod = malzeme_kod[:-2]  # "-0" kısmını çıkar
                            if extracted_sap_kod == sap_kod:
                                return self.extract_row_data_by_index(products_table, row, has_malzeme_kodu=True)
                return None
            elif products_table.columnCount() > 10:  # 11 sütun varsa Malzeme Kodu eklenmiş ama SAP Kodu henüz kaldırılmamış
                sap_col_index = 2

            for row in range(products_table.rowCount()):
                if products_table.item(row, sap_col_index) and products_table.item(row, sap_col_index).text().strip() == sap_kod:
                    # Sütun indekslerini yeni düzene göre ayarla
                    if sap_col_index == 2:  # Malzeme Kodu sütunu varsa
                        malzeme_adi_col = 3
                        spec_col = 4
                        kdv_col = 8
                        siparis_no_col = 9
                        sip_kalem_no_col = 10
                    else:  # Malzeme Kodu sütunu yoksa
                        malzeme_adi_col = 2
                        spec_col = 3
                        kdv_col = 7
                        siparis_no_col = 8
                        sip_kalem_no_col = 9

                    # Güncel tablo verilerini oku
                    spec_value = products_table.item(row, spec_col).text() if products_table.item(row, spec_col) else ''
                    siparis_no = products_table.item(row, siparis_no_col).text() if products_table.item(row, siparis_no_col) else ''
                    sip_kalem_no = products_table.item(row, sip_kalem_no_col).text() if products_table.item(row, sip_kalem_no_col) else ''

                    # Debug için yazdır

                    return {
                        'SAP_Kodu': sap_kod,
                        'Malzeme_Adi': products_table.item(row, malzeme_adi_col).text() if products_table.item(row, malzeme_adi_col) else '',
                        'SPEC': spec_value,
                        'KDV': products_table.item(row, kdv_col).text() if products_table.item(row, kdv_col) else '10%',
                        'Siparis_No': siparis_no,
                        'Sip_Kalem_No': sip_kalem_no
                    }
            return None
        except Exception as e:
            return None

    def extract_row_data_by_index(self, products_table, row, has_malzeme_kodu=False):
        """Satır indeksine göre satır verilerini çıkarır"""
        try:
            if has_malzeme_kodu:
                # Malzeme Kodu sütunu var, SAP Kodu sütunu yok
                malzeme_adi_col = 2
                spec_col = 3
                kdv_col = 7
                siparis_no_col = 8
                sip_kalem_no_col = 9

                # SAP kodunu malzeme kodundan çıkar
                malzeme_kod_item = products_table.item(row, 1)
                malzeme_kod = malzeme_kod_item.text() if malzeme_kod_item else ""
                sap_kod = malzeme_kod[:-2] if malzeme_kod.endswith("-0") else malzeme_kod
            else:
                # Normal durum
                malzeme_adi_col = 2
                spec_col = 3
                kdv_col = 7
                siparis_no_col = 8
                sip_kalem_no_col = 9
                sap_kod = products_table.item(row, 1).text() if products_table.item(row, 1) else ""

            # Güncel tablo verilerini oku
            spec_value = products_table.item(row, spec_col).text() if products_table.item(row, spec_col) else ''
            siparis_no = products_table.item(row, siparis_no_col).text() if products_table.item(row, siparis_no_col) else ''
            sip_kalem_no = products_table.item(row, sip_kalem_no_col).text() if products_table.item(row, sip_kalem_no_col) else ''

            # Debug logging disabled - problem solved
            # print(f"DEBUG extract_row_data_by_index - Row {row}: spec='{spec_value}', kalem='{sip_kalem_no}'")

            return {
                'SAP_Kodu': sap_kod,
                'Malzeme_Adi': products_table.item(row, malzeme_adi_col).text() if products_table.item(row, malzeme_adi_col) else '',
                'SPEC': spec_value,
                'KDV': products_table.item(row, kdv_col).text() if products_table.item(row, kdv_col) else '10%',
                'Siparis_No': siparis_no,
                'Sip_Kalem_No': sip_kalem_no
            }

        except Exception as e:
            return None

    def add_malzeme_kodu_column_and_populate(self, products_table, sap_to_stok_map):
        """ÜRÜNLER tablosuna Malzeme Kodu sütunu ekler ve stok kodlarını doldurur"""
        try:

            # Malzeme Kodu sütunu zaten var mı kontrol et (SAP Kodu sütunu kaldırıldıktan sonra 10 sütun olacak)
            if products_table.columnCount() == 10 and products_table.horizontalHeaderItem(1) and products_table.horizontalHeaderItem(1).text() == "Malzeme Kodu":
                # Sadece stok kodlarını güncelle - artık SAP kodu yok, mapping'den alırız
                for row in range(products_table.rowCount()):
                    # SAP kodunu mapping'den bul (tersini yaparak)
                    current_stok_kod = products_table.item(row, 1).text() if products_table.item(row, 1) else ""
                    for sap_kod, stok_kod_list in sap_to_stok_map.items():
                        if isinstance(stok_kod_list, list) and len(stok_kod_list) > 0:
                            stok_kod = stok_kod_list[0]  # İlk stok kodunu kullan
                        else:
                            stok_kod = stok_kod_list if isinstance(stok_kod_list, str) else ""

                        if stok_kod == current_stok_kod or current_stok_kod == "":
                            stok_kod_item = QTableWidgetItem(stok_kod)
                            products_table.setItem(row, 1, stok_kod_item)
                            break
                return True

            # Yeni sütun ekle
            products_table.insertColumn(1)  # Satir'dan sonra, SAP Kodu'ndan önce ekle

            # SAP kodlarını geçici olarak sakla (SAP Kodu sütununu silmeden önce)
            sap_codes_in_table = {}
            for row in range(products_table.rowCount()):
                sap_item = products_table.item(row, 2)  # SAP kodu 2. sütunda
                if sap_item:
                    sap_codes_in_table[row] = sap_item.text().strip()

            # SAP Kodu sütununu sil (indeks 2)
            products_table.removeColumn(2)

            # Mevcut başlıkları güncelle (SAP Kodu sütunu kaldırıldıktan sonra)
            headers = ["Satir", "Malzeme Kodu", "Malzeme Adı", "SPEC", "Miktar", "Birim Fiyat", "Net Tutar", "KDV", "Sipariş No", "Sip Kalem No"]
            for i, header in enumerate(headers):
                if i < products_table.columnCount():
                    products_table.setHorizontalHeaderItem(i, QTableWidgetItem(header))

            # Her satır için stok kodunu doldur - sadece başarılı olanları
            filled_count = 0
            empty_count = 0

            for row in range(products_table.rowCount()):
                if row in sap_codes_in_table:
                    sap_kod = sap_codes_in_table[row]
                    if sap_kod in sap_to_stok_map and sap_to_stok_map[sap_kod]:
                        # Array formatında stok kodları var - bu satır için doğru olanı bul
                        stok_kod_list = sap_to_stok_map[sap_kod]

                        # Bu satırın SPEC ve Kalem No'suna göre doğru stok kodunu seç
                        selected_stok_kod = ""
                        if isinstance(stok_kod_list, list):
                            # Bu satırın SPEC ve Sip Kalem No verilerini al
                            spec_item = products_table.item(row, 3)  # SPEC sütunu
                            kalem_item = products_table.item(row, 9) # Sip Kalem No sütunu

                            if spec_item and kalem_item:
                                current_spec = spec_item.text().strip()
                                current_kalem = kalem_item.text().strip()

                                # Stok kodları içinde bu kalem no'suna uygun olanı bul
                                for stok_kod in stok_kod_list:
                                    if current_kalem and current_kalem[-3:].zfill(3) in stok_kod:
                                        selected_stok_kod = stok_kod
                                        break

                                # Bulamazsa ilk stok kodunu kullan
                                if not selected_stok_kod and stok_kod_list:
                                    selected_stok_kod = stok_kod_list[0]
                        else:
                            selected_stok_kod = stok_kod_list

                        if selected_stok_kod:
                            stok_kod_item = QTableWidgetItem(selected_stok_kod)
                            products_table.setItem(row, 1, stok_kod_item)
                            filled_count += 1
                        else:
                            empty_item = QTableWidgetItem("")
                            products_table.setItem(row, 1, empty_item)
                            empty_count += 1
                    else:
                        # Başarısız veya boş stok kodu - boş bırak
                        empty_item = QTableWidgetItem("")
                        products_table.setItem(row, 1, empty_item)
                        empty_count += 1

            # Tablo görünümünü yenile
            products_table.resizeColumnsToContents()
            products_table.update()

            return True

        except Exception as e:
            import traceback
            traceback.print_exc()
            return False

    def get_stok_name(self, cursor, stok_kod):
        """Stok koduna göre stok adını döndürür"""
        try:
            cursor.execute("SELECT sto_isim FROM STOKLAR WHERE sto_kod = ?", (stok_kod,))
            result = cursor.fetchone()
            return result[0] if result else "Stok Adı Bulunamadı"
        except Exception as e:
            return "Stok Adı Alınamadı"

    def get_order_date_suffix(self):
        """Sipariş tarihinden GÜN+AY formatında suffix oluşturur (örn: 0920 = 20 Eylül)"""
        try:
            # Contract info'dan sipariş tarihini al
            if hasattr(self, 'contract_info') and self.contract_info:
                order_date = self.contract_info.get('ORD_DATE', '')
                if order_date:
                    # Tarih formatını parse et (örn: "2025-09-20" -> "2009" = 20 Eylül)
                    from datetime import datetime
                    if isinstance(order_date, str):
                        try:
                            # Farklı tarih formatlarını dene
                            for fmt in ['%Y-%m-%d', '%d/%m/%Y', '%d.%m.%Y', '%m/%d/%Y']:
                                try:
                                    date_obj = datetime.strptime(order_date, fmt)
                                    # GÜN + AY formatı (20 Eylül için 2009)
                                    return f"{date_obj.day:02d}{date_obj.month:02d}"
                                except ValueError:
                                    continue
                        except:
                            pass
                    elif hasattr(order_date, 'month') and hasattr(order_date, 'day'):
                        # Datetime objesi ise
                        return f"{order_date.month:02d}{order_date.day:02d}"

            # Varsayılan olarak bugünün tarihini kullan
            from datetime import datetime
            today = datetime.now()
            return f"{today.month:02d}{today.day:02d}"
        except:
            return "0915"  # Fallback

    def generate_stock_code_and_name(self, sap_kod, row_data):
        """Row data'dan stok kodu ve malzeme adı oluşturur"""
        spec_value = row_data.get('SPEC', '').strip()
        siparis_no = row_data.get('Siparis_No', '').strip()
        sip_kalem_no = row_data.get('Sip_Kalem_No', '').strip()

        # Debug logging disabled - problem solved
        # print(f"DEBUG generate_stock_code_and_name: SAP={sap_kod}, kalem='{sip_kalem_no}'")

        if spec_value and siparis_no and sip_kalem_no:
            # SPEC, Sipariş No ve Sip Kalem No dolu ise özel kod oluştur

            # 1. SAP Kodu: Tam hali (3120018724)
            sap_tam = sap_kod

            # 2. Sipariş No: Son 4 hane (1102548493 -> 8493)
            siparis_son4 = siparis_no[-4:] if len(siparis_no) >= 4 else siparis_no.zfill(4)

            # 3. Sip Kalem No: Son 3 hane (000030 -> 030)
            kalem_son3 = sip_kalem_no[-3:] if len(sip_kalem_no) >= 3 else sip_kalem_no.zfill(3)

            # 4. Bugünün tarihi: MMDD formatı (21 Eylül -> 0921)
            from datetime import datetime
            today = datetime.now()
            tarih_suffix = f"{today.month:02d}{today.day:02d}"

            # 5. Format: SAP-SiparişSon4+KalemSon3-Tarih
            # Örnek: 3120018724-8493030-0921
            stok_kod = f"{sap_tam}-{siparis_son4}{kalem_son3}-{tarih_suffix}"
            malzeme_adi = f"{row_data.get('Malzeme_Adi', '')} - {spec_value}"

            return stok_kod, malzeme_adi

        # SPEC boş veya gerekli veriler eksik ise normal kod
        stok_kod = f"{sap_kod}-0"
        malzeme_adi = row_data.get('Malzeme_Adi', '')
        return stok_kod, malzeme_adi

    def check_stok_karti_exists(self, stok_kod):
        """Stok kartının mevcut olup olmadığını kontrol eder - aktif/pasif durumunu da döndürür"""
        try:
            conn = self.get_connection()
            if not conn:
                return {'exists': False, 'active': False, 'passive': False}

            cursor = conn.cursor()

            # 1. STOKLAR_CHOOSE_3A'da aktif kayıt var mı kontrol et
            cursor.execute("SELECT * FROM STOKLAR_CHOOSE_3A WHERE msg_S_0870 = ?", (stok_kod,))
            active_result = cursor.fetchone()

            if active_result:
                conn.close()
                return {'exists': True, 'active': True, 'passive': False}

            # 2. STOKLAR_CHOOSE_2A'da pasif kayıt var mı kontrol et
            cursor.execute("SELECT * FROM STOKLAR_CHOOSE_2A WHERE msg_S_0078 = ?", (stok_kod,))
            passive_result = cursor.fetchone()

            if passive_result:
                conn.close()
                return {'exists': True, 'active': False, 'passive': True}

            # 3. STOKLAR tablosunda genel kontrol
            cursor.execute("SELECT sto_kod, sto_pasif_fl FROM STOKLAR WHERE sto_kod = ?", (stok_kod,))
            general_result = cursor.fetchone()

            conn.close()

            if general_result:
                is_passive = general_result[1] == 1 if len(general_result) > 1 else False
                return {'exists': True, 'active': not is_passive, 'passive': is_passive}
            else:
                return {'exists': False, 'active': False, 'passive': False}

        except Exception as e:
            logging.error(f"ERROR in check_stok_karti_exists: {e}")
            return {'exists': False, 'active': False, 'passive': False}

    def handle_new_stok_karti(self, sap_kod, row_data):
        """Yeni stok kartı oluşturma işlemlerini yürütür"""
        try:
            spec_value = row_data.get('SPEC', '').strip()
            siparis_no = row_data.get('Siparis_No', '').strip()
            sip_kalem_no = row_data.get('Sip_Kalem_No', '').strip()
            kdv_text = row_data.get('KDV', '10%')
            kdv_oran = self.parse_kdv_from_text(kdv_text)
            vergi_kodu = self.convert_kdv_to_vergi_kodu(kdv_oran)

            # Debug logging disabled - problem solved
            # print(f"DEBUG handle_new_stok_karti - SAP: {sap_kod}, kalem: '{sip_kalem_no}'")

            # SPEC boş ise "Sap Kodu-0" formatında stok kartı oluştur
            if not spec_value or spec_value == '':
                # Stok kartı var mı kontrol et
                existing_stok_kod = f"{sap_kod}-0"
                stok_status = self.check_stok_karti_exists(existing_stok_kod)

                if stok_status['active']:
                    return True, f"{existing_stok_kod}\nMalzeme Adı: {row_data.get('Malzeme_Adi', '')} (Mevcut)"
                elif stok_status['passive']:
                    # Pasif stok kartını aktif hale getir
                    try:
                        conn = self.get_connection()
                        cursor = conn.cursor()
                        update_query = "UPDATE STOKLAR SET sto_pasif_fl = 0 WHERE sto_kod = ?"
                        cursor.execute(update_query, (existing_stok_kod,))
                        conn.commit()
                        conn.close()
                        return True, f"{existing_stok_kod}\nMalzeme Adı: {row_data.get('Malzeme_Adi', '')} (Pasiften aktif yapıldı)"
                    except Exception as e:
                        return False, f"SAP Kodu {sap_kod}: Stok kartı aktif yapılamadı: {str(e)}"
                else:
                    # Hiç stok kartı yok, yeni oluştur
                    stok_data = {
                        'sto_kod': existing_stok_kod,
                        'sto_isim': row_data.get('Malzeme_Adi', ''),
                        'sto_perakende_vergi': vergi_kodu,
                        'sto_toptan_vergi': vergi_kodu,
                        'sto_oto_barkod_kod_yapisi': "0"
                    }
                    result = self.create_stok_karti(stok_data)
                    if result:
                        return True, f"{existing_stok_kod}\nMalzeme Adı: {row_data.get('Malzeme_Adi', '')} (Yeni oluşturuldu)"
                    else:
                        return False, f"SAP Kodu {sap_kod}: Yeni stok kartı oluşturulamadı"

            # SPEC dolu ama "None" içeriyorsa uyarı ver
            elif 'none' in spec_value.lower():
                if siparis_no and sip_kalem_no:
                    # Sipariş bilgileri varsa devam et
                    pass
                else:
                    return False, f"SAP Kodu {sap_kod}: SPEC veya Sipariş bilgileri eksik"

            # SPEC dolu ve düzgün ama Sipariş No veya Sip Kalem No boş ise uyarı ver
            elif not siparis_no or not sip_kalem_no:
                return False, f"SAP Kodu {sap_kod}: Sipariş bilgileri eksik"

            # SPEC dolu ve Sipariş No, Sip Kalem No da dolu - normal stok kartı oluştur
            else:
                # GEÇICI ÇÖZÜM: row_data kullanmak yerine doğrudan verileri geç
                fresh_row_data = {
                    'SPEC': spec_value,
                    'Siparis_No': siparis_no,
                    'Sip_Kalem_No': sip_kalem_no,
                    'Malzeme_Adi': row_data.get('Malzeme_Adi', '')
                }
                # print(f"FRESH DATA: SPEC='{spec_value}', Siparis='{siparis_no}', Kalem='{sip_kalem_no}'")

                stok_kod, malzeme_adi = self.generate_stock_code_and_name(sap_kod, fresh_row_data)
                barkod = f"{siparis_no}{sip_kalem_no}"

                # Önce bu stok kodu mevcut mu kontrol et
                existing_status = self.check_stok_karti_exists(stok_kod)

                if existing_status['active']:
                    return True, f"{stok_kod}\nMalzeme Adı: {malzeme_adi} (Mevcut)"
                elif existing_status['passive']:
                    # Pasif stok kartını aktif hale getir
                    try:
                        conn = self.get_connection()
                        cursor = conn.cursor()
                        update_query = "UPDATE STOKLAR SET sto_pasif_fl = 0 WHERE sto_kod = ?"
                        cursor.execute(update_query, (stok_kod,))
                        conn.commit()
                        conn.close()
                        return True, f"{stok_kod}\nMalzeme Adı: {malzeme_adi} (Pasiften aktif yapıldı)"
                    except Exception as e:
                        return False, f"SAP Kodu {sap_kod}: Stok kartı aktif yapılamadı: {str(e)}"
                else:
                    # Yeni stok kartı oluştur
                    stok_data = {
                        'sto_kod': stok_kod,
                        'sto_isim': malzeme_adi,
                        'sto_yabanci_isim': spec_value,
                        'sto_perakende_vergi': vergi_kodu,
                        'sto_toptan_vergi': vergi_kodu,
                        'sto_oto_barkod_kod_yapisi': barkod
                    }
                    result = self.create_stok_karti(stok_data)

                    if result:
                        return True, f"{stok_kod}\nMalzeme Adı: {malzeme_adi} (Yeni oluşturuldu)"
                    else:
                        return False, f"SAP Kodu {sap_kod}: Yeni stok kartı oluşturulamadı"

        except Exception as e:
            return False, f"SAP Kodu {sap_kod}: Stok kartı oluşturma sırasında hata: {str(e)}"

    def parse_kdv_from_text(self, kdv_text):
        """KDV metnini parse eder (örn: '10%' -> 10)"""
        try:
            if isinstance(kdv_text, str):
                # '%' işaretini kaldır ve sayıya dönüştür
                kdv_text = kdv_text.replace('%', '').strip()
            return float(kdv_text)
        except:
            return 10  # Varsayılan %10

    def convert_kdv_to_vergi_kodu(self, kdv_oran):
        """KDV oranını Mikro vergi koduna dönüştür (PRG.py'den uyarlandı)"""
        try:
            kdv_oran = float(kdv_oran)
            kdv_mapping = {
                1: 2,    # %1 KDV → Kod 2
                8: 3,    # %8 KDV → Kod 3
                10: 7,   # %10 KDV → Kod 7 (varsayılan)
                18: 4,   # %18 KDV → Kod 4
                20: 8    # %20 KDV → Kod 8
            }
            return kdv_mapping.get(kdv_oran, 7)  # Varsayılan %10 (kod 7)
        except:
            return 7  # Hata durumunda varsayılan %10 (kod 7)

    def create_stok_karti(self, stok_data):
        """Mikro DB'de stok kartı oluştur (PRG.py'den uyarlandı)"""
        try:

            conn = self.get_connection()
            if not conn:
                return False

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
                # 1. Adım: Stok kodunun varlığını kontrol et
                cursor.execute("SELECT COUNT(*) FROM STOKLAR WHERE sto_kod = ?", (stok_kod,))
                if cursor.fetchone()[0] > 0:
                    conn.close()
                    return False

                # 2. Adım: STOKLAR tablosuna ekleme
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

                # 3. Adım: BARKOD_TANIMLARI tablosuna ekleme
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
                conn.close()
                return True

            except pyodbc.IntegrityError as e:
                QMessageBox.warning(self, "Uyarı", f"Bu stok kodu zaten mevcut: {stok_kod}")
                conn.rollback()
                conn.close()
                return False
            except Exception as e:
                error_msg = f"Stok kartı oluşturulamadı:\n{str(e)}"
                logging.error(f"ERROR create_stok_karti inner: {error_msg}")
                QMessageBox.critical(self, "Hata", error_msg)
                conn.rollback()
                conn.close()
                return False

        except Exception as e:
            error_msg = f"Veritabanı bağlantı hatası: {str(e)}"
            logging.error(f"ERROR create_stok_karti outer: {error_msg}")
            QMessageBox.critical(self, "Hata", error_msg)
            return False

    def create_order(self):
        """Stok Aktar butonuna tıklandığında çalışan fonksiyon"""
        try:
            # İlk önce malzeme kartı kontrolü yap
            #QMessageBox.information(self, "Bilgi", "Malzeme kartı kontrolü başlatılıyor...")

            # ÜRÜNLER tablosunu bul ve SAP kodlarını topla
            products_table = self.find_products_table()
            if not products_table:
                QMessageBox.warning(self, "Uyarı", "ÜRÜNLER tablosu bulunamadı.")
                return

            sap_codes = self.collect_sap_codes(products_table)
            if not sap_codes:
                QMessageBox.information(self, "Bilgi", "Kontrol edilecek SAP kodu bulunamadı.")
                return

            # Malzeme kontrolü işlemini yap
            result = self.process_material_control(sap_codes, products_table)

            if result['success']:
                # Başarı durumunda mesaj gösterme, sadece tablo kontrolü yap
                message = ""

                # Tabloya Malzeme Kodu sütunu ekle - sadece başarılı stok kodları varsa
                successful_stok_count = (result['activated'] + result['created'])
                total_processed = result['processed']

                # Önce tüm satırlar için stok kartı oluşturulup oluşturulmadığını kontrol et
                sap_to_stok_map = result.get('sap_to_stok_map', {})
                total_stok_codes = sum(len(codes) if isinstance(codes, list) else 1 for codes in sap_to_stok_map.values())
                empty_count = total_processed - total_stok_codes

                # print(f"DEBUG: total_processed={total_processed}, total_stok_codes={total_stok_codes}, empty_count={empty_count}")
                # print(f"DEBUG: sap_to_stok_map={sap_to_stok_map}")

                if empty_count > 0:
                    # Bazı satırlar için stok kartı oluşturulamadıysa sadece uyarı ver, tablo oluşturma
                    message = "❌ Tablo oluşturulmadı!"
                    message += f"\n📋 Lütfen aşağıdaki adımları tamamlayın:"

                    if result.get('error_list'):
                        for error in result['error_list']:
                            if "SPEC" in error or "Sipariş" in error:
                                message += f"\n  • {error}"

                    # Hata durumunda mesaj göster
                    QMessageBox.information(self, "Stok Aktar Sonucu", message)

                elif result.get('sap_to_stok_map') and len(result['sap_to_stok_map']) > 0:
                    # Tüm satırlar başarılı - tablo oluştur

                    column_added = self.add_malzeme_kodu_column_and_populate(products_table, result['sap_to_stok_map'])

                    if column_added:
                        # Başarı durumunda mesaj gösterme

                        # Stok Aktar butonunu deaktif hale getir (tekrar kullanımı önlemek için)
                        if hasattr(self, 'create_order_btn'):
                            self.create_order_btn.setEnabled(False)
                            self.create_order_btn.setStyleSheet("""
                                QPushButton {
                                    background-color: #95a5a6;
                                    color: #7f8c8d;
                                    border: none;
                                    padding: 10px 20px;
                                    font-size: 14px;
                                    font-weight: bold;
                                    border-radius: 5px;
                                    min-width: 120px;
                                }
                            """)

                        # Sipariş Aktar butonunu aktif hale getir
                        if hasattr(self, 'transfer_order_btn'):
                            self.transfer_order_btn.setEnabled(True)
                            self.transfer_order_btn.setStyleSheet("""
                                QPushButton {
                                    background-color: #f39c12;
                                    color: white;
                                    border: none;
                                    padding: 10px 20px;
                                    font-size: 14px;
                                    font-weight: bold;
                                    border-radius: 5px;
                                    min-width: 120px;
                                }
                                QPushButton:hover {
                                    background-color: #e67e22;
                                }
                            """)
                    else:
                        # Hata durumunda mesaj göster
                        error_message = "❌ Malzeme Kodu sütunu eklenirken hata oluştu."
                        QMessageBox.information(self, "Stok Aktar Sonucu", error_message)
                else:
                    # Hiç başarılı stok kartı yoksa
                    error_message = "❌ Hiç stok kartı oluşturulamadı."
                    error_message += f"\n📋 Lütfen SPEC ve Sipariş bilgilerini kontrol edin."
                    QMessageBox.information(self, "Stok Aktar Sonucu", error_message)

            else:
                QMessageBox.warning(self, "Uyarı", "Malzeme kontrolü sırasında hata oluştu.")

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Sipariş oluşturma işleminde hata oluştu: {str(e)}")

    def transfer_order(self):
        """Sipariş Aktar butonuna tıklandığında çalışan fonksiyon"""
        try:
            # Önce mevcut sözleşmenin daha önce kaydedilip kaydedilmediğini kontrol et
            existing_contract = self.check_existing_contract()

            if existing_contract is not None:
                existing_evrakno_sira, has_teslim_miktar = existing_contract

                # Eğer teslim_miktar > 0 varsa silmeye izin verme
                if has_teslim_miktar:
                    QMessageBox.critical(
                        self,
                        "UYARI - Silme İşlemi Engellendi",
                        f"Sözleşme No: {self.contract_id} (Evrak No: {existing_evrakno_sira})\n\n"
                        "Bu sözleşmede teslim edilmiş ürünler bulunmaktadır!\n"
                        "(sip_teslim_miktar > 0)\n\n"
                        "GÜVENLİK UYARISI:\n"
                        "Otomatik silme işlemi güvenlik nedeniyle engellendi.\n"
                        "Silme işlemini manuel olarak yapmanız gerekmektedir.\n\n"
                        "Veritabanında şu sorguyu çalıştırın:\n"
                        f"DELETE FROM [dbo].[SIPARISLER] WHERE sip_belgeno = '{self.contract_id}'"
                    )
                    return

                # Teslim_miktar yoksa normal uyarı göster
                reply = QMessageBox.question(
                    self,
                    "Uyarı",
                    f"Sözleşme daha önce kaydedilmiş (Evrak No: {existing_evrakno_sira})\n\n"
                    "Mevcut kayıtlar silinip yeni kayıt yapılacak. Devam etmek istiyor musunuz?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.No
                )

                if reply == QMessageBox.No:
                    return

                # Mevcut kayıtları sil ve aynı evrakno_sira'yı kullan
                self.delete_existing_contract_records()
                self.current_evrakno_sira = existing_evrakno_sira
            else:
                # Yeni kayıt - yeni evrakno_sira al
                self.current_evrakno_sira = self.get_next_evrakno_sira()

            QMessageBox.information(self, "Bilgi", "Sipariş oluşturma başlatıldı...")

            # Ürünler tablosunu bul
            products_table = self.find_products_table()
            if not products_table:
                QMessageBox.critical(self, "Hata", "ÜRÜNLER tablosu bulunamadı!")
                return

            # Tablo sütunlarını kontrol et
            column_headers = []
            for col in range(products_table.columnCount()):
                header_item = products_table.horizontalHeaderItem(col)
                if header_item:
                    column_headers.append(header_item.text())

            # Malzeme Kodu sütunu var mı kontrol et
            malzeme_kodu_col = None
            for col in range(products_table.columnCount()):
                header_item = products_table.horizontalHeaderItem(col)
                if header_item and header_item.text() == "Malzeme Kodu":
                    malzeme_kodu_col = col
                    break

            if malzeme_kodu_col is None:
                QMessageBox.warning(self, "Uyarı", "Malzeme Kodu sütunu bulunamadı! Önce 'Stok Aktar' işlemini tamamlayın.")
                return

            # Sipariş oluşturma işlemini başlat
            order_count = 0
            total_rows = products_table.rowCount()
            failed_orders = []

            # Sütun indekslerini belirle
            column_indices = {}
            for col in range(products_table.columnCount()):
                header_item = products_table.horizontalHeaderItem(col)
                if header_item:
                    column_indices[header_item.text()] = col

            for row in range(total_rows):
                malzeme_kodu_item = products_table.item(row, malzeme_kodu_col)
                if malzeme_kodu_item and malzeme_kodu_item.text().strip():
                    # Bu satır için sipariş verilerini topla
                    order_data = self.collect_order_data(products_table, row, column_indices)
                    order_data['row_index'] = row  # Satır numarasını ekle

                    # Sipariş oluştur
                    success = self.create_order_record(order_data)
                    if success:
                        order_count += 1
                    else:
                        failed_orders.append(f"Satır {row + 1}")

            # Sonuç mesajı
            if order_count > 0:
                message = f"Sipariş oluşturma tamamlandı!\n{order_count} adet sipariş oluşturuldu."
                if failed_orders:
                    message += f"\n\nBaşarısız olanlar: {', '.join(failed_orders)}"
                QMessageBox.information(self, "Başarı", message)

                # Sipariş Aktar butonunu tekrar inaktif yap
                self.transfer_order_btn.setEnabled(False)
                self.transfer_order_btn.setStyleSheet("""
                    QPushButton {
                        background-color: #bdc3c7;
                        color: #7f8c8d;
                        border: none;
                        padding: 10px 20px;
                        font-size: 14px;
                        font-weight: bold;
                        border-radius: 5px;
                        min-width: 120px;
                    }
                """)
            else:
                QMessageBox.warning(self, "Uyarı", "Sipariş oluşturulacak malzeme bulunamadı!")

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Sipariş oluşturma işleminde hata oluştu: {str(e)}")

    def collect_order_data(self, products_table, row, column_indices):
        """Tablodaki satırdan sipariş verilerini toplar"""
        order_data = {}

        # Temel veriler
        for col_name, col_index in column_indices.items():
            item = products_table.item(row, col_index)
            order_data[col_name] = item.text().strip() if item else ""

        return order_data

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

    def get_next_recno(self):
        """SIPARISLER tablosundan bir sonraki RECno değerini alır"""
        try:
            conn = self.get_connection()
            if not conn:
                return 1

            cursor = conn.cursor()
            cursor.execute("SELECT MAX(sip_RECno) FROM SIPARISLER")
            result = cursor.fetchone()
            conn.close()

            if result and result[0]:
                return result[0] + 1
            else:
                return 1
        except:
            return 1

    def get_next_evrakno_sira(self):
        """SIPARISLER tablosundan bir sonraki evrakno_sira değerini alır"""
        try:
            conn = self.get_connection()
            if not conn:
                return 1

            cursor = conn.cursor()
            cursor.execute("SELECT MAX(CAST(sip_evrakno_sira AS INT)) FROM SIPARISLER WHERE ISNUMERIC(sip_evrakno_sira) = 1")
            result = cursor.fetchone()
            conn.close()

            if result and result[0]:
                return result[0] + 1
            else:
                return 1
        except:
            return 1

    def get_musteri_kod(self, safe_get_contract_data):
        """340. ile başlayan cari kodunu alır"""
        # Önce selected_cari_kod'u kontrol et (340. ile başlayan)
        if hasattr(self, 'selected_cari_kod') and self.selected_cari_kod:
            if str(self.selected_cari_kod).startswith('340.'):
                return self.selected_cari_kod

        # Alternatif: Contract_data'dan customer bilgilerini alıp cari kod oluştur
        try:
            if hasattr(self.contract_data, 'ES_CONTRACT_INFO'):
                contract_info = self.contract_data.ES_CONTRACT_INFO
                customer_taxnr = getattr(contract_info, 'CUSTOMER_TAXNR', '')
                customer_phone1 = getattr(contract_info, 'CUSTOMER_PHONE1', '')

                if customer_taxnr:
                    generated_cari_kod = f"340.{customer_taxnr}"
                    self.selected_cari_kod = generated_cari_kod
                    return generated_cari_kod
                elif customer_phone1:
                    generated_cari_kod = f"340.{customer_phone1}"
                    self.selected_cari_kod = generated_cari_kod
                    return generated_cari_kod
        except:
            pass

        return ""

    def get_satici_kod_by_salesman(self, salesman_name):
        """Satış temsilcisi adından satıcı kodunu bulur"""
        try:
            conn = self.get_connection()
            if not conn:
                return ""

            cursor = conn.cursor()
            query = """
            SELECT cari_per_kod AS msg_S_1444
            FROM dbo.CARI_PERSONEL_TANIMLARI WITH (NOLOCK)
            WHERE (cari_per_tip=0)
            AND (cari_per_adi + ' ' + cari_per_soyadi) = ?
            """
            cursor.execute(query, (salesman_name,))
            result = cursor.fetchone()

            if result:
                satici_kodu = result[0]
                conn.close()
                return satici_kodu
            else:
                conn.close()
                return ""
        except:
            return ""

    def extract_first_numeric_part(self, text):
        """Header'dan sadece ilk sayısal kısmı çıkarır"""
        if not text:
            return ""

        # Sadece sayıları al
        import re
        numbers = re.findall(r'\d+', str(text))
        if numbers:
            return numbers[0]
        return ""

    def extract_magaza_kod(self, magaza_display):
        """Mağaza display verisinden sadece 16 ile başlayan int değerini çıkarır"""
        if not magaza_display:
            return 1600704  # Varsayılan

        # "1600704 - MERKEZ" veya "1601175 - ŞUBE" formatından sadece sayısal kısmı al
        if isinstance(magaza_display, str):
            # Boşlukla ayır ve ilk parçayı kontrol et
            parts = magaza_display.split(' ')
            if parts and parts[0].isdigit():
                kod = int(parts[0])
                # 16 ile başlayıp başlamadığını kontrol et
                if str(kod).startswith('16'):
                    return kod

        return 1600704  # Varsayılan

    
    def create_order_record(self, order_data):
        """Veritabanında sipariş kaydı oluşturur"""
        try:
            conn = self.get_connection()
            if not conn:
                return False

            cursor = conn.cursor()

            # Varsayılan değerler
            defaults = {
                'sip_RECid_DBCno': 0,
                'sip_SpecRECno': 0,
                'sip_iptal': 0,
                'sip_fileid': 21,
                'sip_hidden': 0,
                'sip_kilitli': 0,
                'sip_degisti': 0,
                'sip_checksum': 0,
                'sip_create_user': 1,
                'sip_lastup_user': 1,
                'sip_special1': '',
                'sip_special2': '',
                'sip_special3': '',
                'sip_firmano': 0,
                'sip_subeno': 0,
                'sip_tip': 0,
                'sip_cins': 0,
                'sip_evrakno_seri': '',
                'sip_birim_pntr': 1,
                'sip_teslim_miktar': 0,
                'sip_iskonto_1': 0,
                'sip_iskonto_2': 0,
                'sip_iskonto_3': 0,
                'sip_iskonto_4': 0,
                'sip_iskonto_5': 0,
                'sip_iskonto_6': 0,
                'sip_masraf_1': 0,
                'sip_masraf_2': 0,
                'sip_masraf_3': 0,
                'sip_masraf_4': 0,
                'sip_masvergi_pntr': 0,
                'sip_masvergi': 0,
                'sip_opno': 0,
                'sip_aciklama': '',
                'sip_depono': 100,
                'sip_OnaylayanKulNo': 0,
                'sip_vergisiz_fl': 0,
                'sip_kapat_fl': 0,
                'sip_promosyon_fl': 0,
                'sip_cari_grupno': 0,
                'sip_doviz_cinsi': 0,
                'sip_doviz_kuru': 1,
                'sip_alt_doviz_kuru': 1,
                'sip_adresno': 1,
                'sip_teslimturu': '',
                'sip_cagrilabilir_fl': 1,
                'sip_prosiprecDbId': 0,
                'sip_prosiprecrecI': 0,
                'sip_iskonto1': 0,
                'sip_iskonto2': 1,
                'sip_iskonto3': 1,
                'sip_iskonto4': 1,
                'sip_iskonto5': 1,
                'sip_iskonto6': 1,
                'sip_masraf1': 1,
                'sip_masraf2': 1,
                'sip_masraf3': 1,
                'sip_masraf4': 1,
                'sip_isk1': 0,
                'sip_isk2': 0,
                'sip_isk3': 0,
                'sip_isk4': 0,
                'sip_isk5': 0,
                'sip_isk6': 0,
                'sip_mas1': 0,
                'sip_mas2': 0,
                'sip_mas3': 0,
                'sip_mas4': 0,
                'sip_Exp_Imp_Kodu': '',
                'sip_kar_orani': 0,
                'sip_durumu': 0,
                'sip_stalRecId_DBCno': 0,
                'sip_stalRecId_RECno': 0,
                'sip_planlananmiktar': 0,
                'sip_teklifRecId_DBCno': 0,
                'sip_teklifRecId_RECno': 0,
                'sip_parti_kodu': '',
                'sip_lot_no': 0,
                'sip_projekodu': '',
                'sip_fiyat_liste_no': 0,
                'sip_Otv_Pntr': 0,
                'sip_Otv_Vergi': 0,
                'sip_otvtutari': 0,
                'sip_OtvVergisiz_Fl': 0,
                'sip_paket_kod': '',
                'sip_RezRecId_DBCno': 0,
                'sip_RezRecId_RECno': 0,
                'sip_harekettipi': 0,
                'sip_yetkili_recid_dbcno': 0,
                'sip_yetkili_recid_recno': 0,
                'sip_kapatmanedenkod': ''
            }

            # Sadeleştirilmiş INSERT sorgusu - sadece önemli sütunlar
            from datetime import datetime

            insert_query = """
            INSERT INTO SIPARISLER (
                sip_RECid_RECno, sip_create_date, sip_lastup_date, sip_create_user, sip_lastup_user,
                sip_tarih, sip_teslim_tarih, sip_evrakno_sira, sip_evrakno_seri, sip_satirno,
                sip_belgeno, sip_belge_tarih, sip_satici_kod, sip_musteri_kod,
                sip_stok_kod, sip_b_fiyat, sip_birim_pntr, sip_miktar, sip_tutar,
                sip_vergi_pntr, sip_vergi, sip_aciklama, sip_aciklama2,
                sip_cari_sormerk, sip_stok_sormerk, sip_doviz_kuru, sip_alt_doviz_kuru, sip_adresno, sip_cagrilabilir_fl,
                sip_iskonto2, sip_iskonto3, sip_iskonto4, sip_iskonto5, sip_iskonto6,
                sip_masraf1, sip_masraf2, sip_masraf3, sip_masraf4,
                sip_iskonto_1, sip_iskonto_2, sip_iskonto_3, sip_iskonto_4, sip_iskonto_5, sip_iskonto_6,
                sip_masraf_1, sip_masraf_2, sip_masraf_3, sip_masraf_4,
                sip_masvergi_pntr, sip_masvergi, sip_opno, sip_depono,
                sip_OnaylayanKulNo, sip_vergisiz_fl, sip_kapat_fl, sip_promosyon_fl,
                sip_cari_grupno, sip_doviz_cinsi,
                sip_prosiprecDbId, sip_prosiprecrecI, sip_iskonto1,
                sip_isk1, sip_isk2, sip_isk3, sip_isk4, sip_isk5, sip_isk6,
                sip_mas1, sip_mas2, sip_mas3, sip_mas4,
                sip_kar_orani, sip_durumu, sip_stalRecId_DBCno, sip_stalRecId_RECno, sip_planlananmiktar,
                sip_teklifRecId_DBCno, sip_teklifRecId_RECno, sip_lot_no,
                sip_fiyat_liste_no, sip_Otv_Pntr, sip_Otv_Vergi, sip_otvtutari, sip_OtvVergisiz_Fl,
                sip_RezRecId_DBCno, sip_RezRecId_RECno, sip_harekettipi, sip_yetkili_recid_dbcno, sip_yetkili_recid_recno,
                sip_teslimturu, sip_Exp_Imp_Kodu, sip_parti_kodu, sip_projekodu, sip_paket_kod, sip_kapatmanedenkod,
                sip_special1, sip_special2, sip_special3,
                sip_RECid_DBCno, sip_SpecRECno, sip_iptal, sip_fileid, sip_hidden, sip_kilitli,
                sip_degisti, sip_checksum, sip_firmano, sip_subeno, sip_tip, sip_cins, sip_teslim_miktar
            ) VALUES (
                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
            )
            """

            now = datetime.now()

            # Gerekli verileri al
            next_recno = self.get_next_recno()
            # Aynı sözleşme için aynı evrakno_sira kullan
            next_evrakno = getattr(self, 'current_evrakno_sira', self.get_next_evrakno_sira())

            # Contract_data'dan güvenli veri alma fonksiyonu
            def safe_get_contract_data(key, default=''):
                try:
                    # SAP response objesi contract_data.ES_CONTRACT_INFO yapısını kullan
                    if hasattr(self.contract_data, 'ES_CONTRACT_INFO'):
                        contract_info = self.contract_data.ES_CONTRACT_INFO
                        if hasattr(contract_info, key):
                            return getattr(contract_info, key, default)
                    return default
                except:
                    return default

            # Sipariş tarihi - mevcut tarih kullan
            siparis_date = now.date()

            # Satıcı kodu - SALESMAN_NAMEFIRST + SALESMAN_NAMELAST birleştir
            first_name = safe_get_contract_data('SALESMAN_NAMEFIRST', '')
            last_name = safe_get_contract_data('SALESMAN_NAMELAST', '')
            salesman_name = f"{first_name} {last_name}".strip()

            satici_kod = self.get_satici_kod_by_salesman(salesman_name)

            # Header verisini contract_info'dan al
            header_display_text = safe_get_contract_data('HEADER_TEXT', '0')

            # Header'dan sadece sayısal kısım - contract_id kullan (eski kullanım için)
            header_text = self.extract_first_numeric_part(str(self.contract_id))

            # Mağaza kodu - sales_office'den magaza_display oluştur ve extract et
            sales_office = safe_get_contract_data('SALES_OFFICE', '')
            if sales_office == 'IM1':
                magaza_display = '1600704 - MERKEZ'
            elif sales_office == 'IM2':
                magaza_display = '1601175 - ŞUBE'
            else:
                magaza_display = sales_office

            magaza_kod = self.extract_magaza_kod(magaza_display)

            # KDV oranını vergi koduna dönüştür
            kdv_oran = order_data.get('KDV', '10%').replace('%', '')
            vergi_kodu = self.convert_kdv_to_vergi_kodu(kdv_oran)

            # Satır numarası - order_data'dan al
            satirno = order_data.get('row_index', 0)

            # Net Tutar'dan TL ifadesini temizle ve birim fiyat hesapla
            def clean_currency_string(value_str):
                """TL ifadesini temizler ve sayısal değeri döndürür"""
                if not value_str:
                    return 0.0
                # TL ifadesini kaldır, virgülü kaldır (nokta ayırıcı olarak bırak)
                cleaned = str(value_str).replace(' TL', '').replace('TL', '').replace(',', '').strip()
                try:
                    return float(cleaned)
                except ValueError:
                    return 0.0

            # Net Tutar ve Miktar verilerini al
            net_tutar_str = order_data.get('Net Tutar', '0')
            miktar = float(order_data.get('Miktar', 0) or 0)

            # Net Tutar'ı temizle
            net_tutar = clean_currency_string(net_tutar_str)

            # Birim fiyat = Net Tutar / Miktar
            if miktar > 0:
                birim_fiyat = net_tutar / miktar
            else:
                birim_fiyat = 0.0

            # Tutar = sip_miktar * sip_b_fiyat
            tutar = miktar * birim_fiyat

            # Değerleri hazırla
            values = (
                # Ana değerler
                next_recno,  # sip_RECid_RECno
                now,  # sip_create_date (şu anki tarih/saat)
                now,  # sip_lastup_date (şu anki tarih/saat)
                1,  # sip_create_user
                1,  # sip_lastup_user
                siparis_date,  # sip_tarih (sipariş tarihi)
                siparis_date,  # sip_teslim_tarih (sipariş tarihi)
                str(next_evrakno),  # sip_evrakno_sira
                '',  # sip_evrakno_seri (boş string)
                satirno,  # sip_satirno (satır numarası)
                self.contract_id,  # sip_belgeno (contract_id)
                siparis_date,  # sip_belge_tarih (sipariş tarihi)
                satici_kod,  # sip_satici_kod
                self.get_musteri_kod(safe_get_contract_data),  # sip_musteri_kod
                order_data.get('Malzeme Kodu', ''),  # sip_stok_kod
                birim_fiyat,  # sip_b_fiyat
                1,  # sip_birim_pntr
                miktar,  # sip_miktar
                tutar,  # sip_tutar (hesaplanan)
                vergi_kodu,  # sip_vergi_pntr (KDV kodu)
                tutar * (float(kdv_oran or 10) / 100),  # sip_vergi (tutarı, oranı değil)
                '',  # sip_aciklama (boş)
                header_display_text if satirno == 0 else '',  # sip_aciklama2 (Header verisi sadece ilk satırda)
                magaza_kod,  # sip_cari_sormerk
                magaza_kod,  # sip_stok_sormerk
                1,  # sip_doviz_kuru
                1,  # sip_alt_doviz_kuru
                1,  # sip_adresno
                1,  # sip_cagrilabilir_fl
                1,  # sip_iskonto2
                1,  # sip_iskonto3
                1,  # sip_iskonto4
                1,  # sip_iskonto5
                1,  # sip_iskonto6
                1,  # sip_masraf1
                1,  # sip_masraf2
                1,  # sip_masraf3
                1,  # sip_masraf4
                0,  # sip_iskonto_1
                0,  # sip_iskonto_2
                0,  # sip_iskonto_3
                0,  # sip_iskonto_4
                0,  # sip_iskonto_5
                0,  # sip_iskonto_6
                0,  # sip_masraf_1
                0,  # sip_masraf_2
                0,  # sip_masraf_3
                0,  # sip_masraf_4
                0,  # sip_masvergi_pntr
                0,  # sip_masvergi
                0,  # sip_opno
                100,  # sip_depono
                0,  # sip_OnaylayanKulNo
                0,  # sip_vergisiz_fl
                0,  # sip_kapat_fl
                0,  # sip_promosyon_fl
                0,  # sip_cari_grupno
                0,  # sip_doviz_cinsi
                0,  # sip_prosiprecDbId
                0,  # sip_prosiprecrecI
                0,  # sip_iskonto1
                0,  # sip_isk1
                0,  # sip_isk2
                0,  # sip_isk3
                0,  # sip_isk4
                0,  # sip_isk5
                0,  # sip_isk6
                0,  # sip_mas1
                0,  # sip_mas2
                0,  # sip_mas3
                0,  # sip_mas4
                0,  # sip_kar_orani
                0,  # sip_durumu
                0,  # sip_stalRecId_DBCno
                0,  # sip_stalRecId_RECno
                0,  # sip_planlananmiktar
                0,  # sip_teklifRecId_DBCno
                0,  # sip_teklifRecId_RECno
                0,  # sip_lot_no
                0,  # sip_fiyat_liste_no
                0,  # sip_Otv_Pntr
                0,  # sip_Otv_Vergi
                0,  # sip_otvtutari
                0,  # sip_OtvVergisiz_Fl
                0,  # sip_RezRecId_DBCno
                0,  # sip_RezRecId_RECno
                0,  # sip_harekettipi
                0,  # sip_yetkili_recid_dbcno
                0,  # sip_yetkili_recid_recno
                '',  # sip_teslimturu
                '',  # sip_Exp_Imp_Kodu
                '',  # sip_parti_kodu
                '',  # sip_projekodu
                '',  # sip_paket_kod
                '',  # sip_kapatmanedenkod
                '',  # sip_special1 (boş string)
                '',  # sip_special2 (boş string)
                '',  # sip_special3 (boş string)
                # Varsayılan değerler
                0,  # sip_RECid_DBCno
                0,  # sip_SpecRECno
                0,  # sip_iptal
                21,  # sip_fileid
                0,  # sip_hidden
                0,  # sip_kilitli
                0,  # sip_degisti
                0,  # sip_checksum
                0,  # sip_firmano
                0,  # sip_subeno
                0,  # sip_tip
                0,  # sip_cins
                0   # sip_teslim_miktar
            )

            cursor.execute(insert_query, values)
            conn.commit()
            conn.close()
            return True

        except Exception as e:
            logging.error(f"Sipariş oluşturma hatası: {str(e)}")
            return False

    
    def update_products_table_row(self, sap_kod, updated_data):
        """ÜRÜNLER tablosunda belirtilen satırı günceller"""
        try:
            products_table = self.find_products_table()
            if not products_table:
                QMessageBox.critical(self, "Hata", "ÜRÜNLER tablosu bulunamadı!")
                return False

            for row in range(products_table.rowCount()):
                sap_item = products_table.item(row, 1)
                if sap_item and sap_item.text().strip() == sap_kod:
                    # Malzeme Adı güncelle (sütun 2)
                    if 'Malzeme_Adi' in updated_data:
                        item = QTableWidgetItem(updated_data['Malzeme_Adi'])
                        products_table.setItem(row, 2, item)

                    # SPEC güncelle (sütun 3)
                    if 'SPEC' in updated_data:
                        item = QTableWidgetItem(updated_data['SPEC'])
                        products_table.setItem(row, 3, item)

                    # Sipariş No güncelle (sütun 8)
                    if 'Siparis_No' in updated_data:
                        item = QTableWidgetItem(updated_data['Siparis_No'])
                        products_table.setItem(row, 8, item)

                    # Sip Kalem No güncelle (sütun 9)
                    if 'Sip_Kalem_No' in updated_data:
                        item = QTableWidgetItem(updated_data['Sip_Kalem_No'])
                        products_table.setItem(row, 9, item)

                    # Tabloyu yenile
                    products_table.viewport().update()
                    return True

            QMessageBox.warning(self, "Uyarı", f"SAP Kodu {sap_kod} tabloda bulunamadı!")
            return False

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Tablo güncelleme hatası: {str(e)}")
            return False

    def setup_table_copy_functionality(self):
        """ÜRÜNLER tablosuna kopyalama özelliği ekler"""
        try:
            products_table = self.find_products_table()
            if not products_table:
                return

            # Sağ tık menüsü için context menu policy ayarla
            products_table.setContextMenuPolicy(Qt.CustomContextMenu)
            products_table.customContextMenuRequested.connect(self.show_table_context_menu)

            # Seçim modunu ayarla
            products_table.setSelectionMode(QAbstractItemView.ExtendedSelection)

        except Exception as e:
            pass

    def show_table_context_menu(self, position):
        """ÜRÜNLER tablosu için sağ tık menüsü"""
        try:
            products_table = self.find_products_table()
            if not products_table:
                return

            item = products_table.itemAt(position)
            if not item:
                return

            menu = QMenu(self)
            menu.setStyleSheet("""
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

            copy_cell_action = menu.addAction("📋 Hücreyi Kopyala")
            copy_row_action = menu.addAction("📋 Satırı Kopyala")
            copy_column_action = menu.addAction("📋 Sütunu Kopyala")
            copy_selection_action = menu.addAction("📋 Seçimi Kopyala")
            copy_all_action = menu.addAction("📋 Tüm Tabloyu Kopyala")

            action = menu.exec_(products_table.viewport().mapToGlobal(position))

            if action == copy_cell_action:
                self.copy_table_cell(products_table, item)
            elif action == copy_row_action:
                self.copy_table_row(products_table, item.row())
            elif action == copy_column_action:
                self.copy_table_column(products_table, item.column())
            elif action == copy_selection_action:
                self.copy_table_selection(products_table)
            elif action == copy_all_action:
                self.copy_entire_table(products_table)

        except Exception as e:
            pass

    def copy_table_cell(self, table, item):
        """Tek hücreyi kopyala"""
        if item:
            QApplication.clipboard().setText(item.text())

    def copy_table_row(self, table, row):
        """Satırı kopyala"""
        row_data = []
        for col in range(table.columnCount()):
            item = table.item(row, col)
            row_data.append(item.text() if item else '')

        clipboard_text = '\t'.join(row_data)
        QApplication.clipboard().setText(clipboard_text)

    def copy_table_column(self, table, column):
        """Sütunu kopyala"""
        column_data = []
        # Header ekle
        header_item = table.horizontalHeaderItem(column)
        if header_item:
            column_data.append(header_item.text())

        # Tüm satırların o sütununu ekle
        for row in range(table.rowCount()):
            item = table.item(row, column)
            column_data.append(item.text() if item else '')

        clipboard_text = '\n'.join(column_data)
        QApplication.clipboard().setText(clipboard_text)

    def copy_table_selection(self, table):
        """Seçili hücreleri kopyala"""
        selected_items = table.selectedItems()
        if not selected_items:
            return

        # Seçili hücreleri satır/sütun bazında organize et
        selected_ranges = {}
        for item in selected_items:
            row, col = item.row(), item.column()
            if row not in selected_ranges:
                selected_ranges[row] = {}
            selected_ranges[row][col] = item.text()

        # Clipboard formatında düzenle
        data = []
        for row in sorted(selected_ranges.keys()):
            row_data = []
            for col in sorted(selected_ranges[row].keys()):
                row_data.append(selected_ranges[row][col])
            data.append('\t'.join(row_data))

        clipboard_text = '\n'.join(data)
        QApplication.clipboard().setText(clipboard_text)

    def copy_entire_table(self, table):
        """Tüm tabloyu kopyala"""
        data = []

        # Headers ekle
        headers = []
        for col in range(table.columnCount()):
            header_item = table.horizontalHeaderItem(col)
            headers.append(header_item.text() if header_item else f"Sütun {col+1}")
        data.append('\t'.join(headers))

        # Tüm satırları ekle
        for row in range(table.rowCount()):
            row_data = []
            for col in range(table.columnCount()):
                item = table.item(row, col)
                row_data.append(item.text() if item else '')
            data.append('\t'.join(row_data))

        clipboard_text = '\n'.join(data)
        QApplication.clipboard().setText(clipboard_text)

    def handle_products_table_edit(self, item, table):
        """ÜRÜNLER tablosunda hücre düzenlendiğinde çalışır"""
        try:
            if not item:
                return

            row = item.row()
            column = item.column()
            new_value = item.text().strip()

            # Sadece düzenlenebilir sütunlarda değişiklik kabul et
            editable_columns = [3, 8, 9]  # SPEC, Sipariş No, Sip Kalem No
            if column not in editable_columns:
                return

            # SAP Kodu al (sütun 1)
            sap_item = table.item(row, 1)
            if not sap_item:
                return
            sap_kod = sap_item.text().strip()

            # Değişikliği konsola yazdır (debug için)
            column_names = {3: "SPEC", 8: "Sipariş No", 9: "Sip Kalem No"}

            # Tablo değişikliği sonrası contract_data'yı güncelle
            self.update_contract_data_from_table(table)

            # Cari Aktar butonunun durumunu güncelle
            self.update_cari_aktar_button()

        except Exception as e:
            logging.error(f"Tablo düzenleme hatası: {e}")

class TableUpdateDialog(QDialog):
    """
    Tablo Güncelleme Dialog Penceresi

    Bu sınıf, ürün bilgilerini güncellemek için kullanılan bir dialog penceresidir.
    SPEC alanı dolu ama Sipariş No veya Sip Kalem No eksik olan durumlar için
    kullanıcıdan eksik bilgileri tamamlamasını ister.

    Özellikler:
        - Modal dialog (ana pencere bloke edilir)
        - Otomatik validasyon (metin değiştikçe)
        - Form alanları:
            * Malzeme Adı (düzenlenebilir)
            * SPEC (boş bırakılabilir veya 1-20 karakter)
            * Sipariş No (SPEC doluysa zorunlu, 10 hane)
            * Sip Kalem No (SPEC doluysa zorunlu, 6 hane, 000XXX formatı)
        - Kayıt ve İptal butonları

    Validasyon Kuralları:
        - SPEC boş ise: Sipariş No ve Sip Kalem No da boş olabilir
        - SPEC dolu ise:
            * Sipariş No: Tam 10 hane rakam
            * Sip Kalem No: Tam 6 hane (000XXX formatı)
    """

    def __init__(self, parent, sap_kod, malzeme_adi, spec, siparis_no, sip_kalem_no):
        """
        TableUpdateDialog yapıcı metodu.

        Args:
            parent (QWidget): Üst pencere
            sap_kod (str): SAP ürün kodu
            malzeme_adi (str): Malzeme adı
            spec (str): Özellik/Spesifikasyon bilgisi
            siparis_no (str): Sipariş numarası
            sip_kalem_no (str): Sipariş kalem numarası
        """
        super().__init__(parent)
        self.setWindowTitle("🔧 Ürün Bilgileri Güncelle")
        self.setModal(True)  # Modal dialog - ana pencereyi bloke eder
        self.resize(500, 400)  # Pencere boyutu

        # Center the dialog
        screen = QApplication.primaryScreen()
        screen_geometry = screen.availableGeometry()
        x = (screen_geometry.width() - self.width()) // 2
        y = (screen_geometry.height() - self.height()) // 2
        self.move(x, y)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        # Title
        title_label = QLabel(f"📋 SAP Kodu: {sap_kod}")
        title_label.setStyleSheet("""
            QLabel {
                font-size: 16px;
                font-weight: bold;
                color: #2c3e50;
                padding: 10px;
                background-color: #ecf0f1;
                border-radius: 5px;
            }
        """)
        layout.addWidget(title_label)

        # Warning message
        warning_label = QLabel("⚠️ SPEC dolu ama Sipariş No veya Sip Kalem No eksik!\nLütfen eksik bilgileri tamamlayın:")
        warning_label.setStyleSheet("""
            QLabel {
                color: #e74c3c;
                font-weight: bold;
                font-size: 14px;
                padding: 10px;
                background-color: #ffeaa7;
                border-radius: 5px;
                border: 2px solid #fdcb6e;
            }
        """)
        layout.addWidget(warning_label)

        # Form fields
        form_layout = QFormLayout()

        # Malzeme Adı
        self.malzeme_adi_edit = QLineEdit(malzeme_adi)
        self.malzeme_adi_edit.setStyleSheet("padding: 8px; font-size: 14px;")
        form_layout.addRow("📦 Malzeme Adı:", self.malzeme_adi_edit)

        # SPEC
        self.spec_edit = QLineEdit(spec)
        self.spec_edit.setStyleSheet("padding: 8px; font-size: 14px;")
        self.spec_edit.textChanged.connect(self.validate_all)
        form_layout.addRow("🏷️ SPEC:", self.spec_edit)

        # Sipariş No
        self.siparis_no_edit = QLineEdit(siparis_no)
        self.siparis_no_edit.setPlaceholderText("10 hane rakam")
        self.siparis_no_edit.setStyleSheet("padding: 8px; font-size: 14px;")
        self.siparis_no_edit.textChanged.connect(self.validate_all)
        form_layout.addRow("📋 Sipariş No:", self.siparis_no_edit)

        # Sip Kalem No
        self.sip_kalem_no_edit = QLineEdit(sip_kalem_no)
        self.sip_kalem_no_edit.setPlaceholderText("000XXX (6 hane)")
        self.sip_kalem_no_edit.setStyleSheet("padding: 8px; font-size: 14px;")
        self.sip_kalem_no_edit.textChanged.connect(self.validate_all)
        form_layout.addRow("🔢 Sip Kalem No:", self.sip_kalem_no_edit)

        layout.addLayout(form_layout)

        # Validation labels
        self.spec_validation_label = QLabel()
        self.spec_validation_label.setStyleSheet("color: red; font-size: 12px;")
        layout.addWidget(self.spec_validation_label)

        self.siparis_validation_label = QLabel()
        self.siparis_validation_label.setStyleSheet("color: red; font-size: 12px;")
        layout.addWidget(self.siparis_validation_label)

        self.kalem_validation_label = QLabel()
        self.kalem_validation_label.setStyleSheet("color: red; font-size: 12px;")
        layout.addWidget(self.kalem_validation_label)

        # Buttons
        button_layout = QHBoxLayout()

        self.cancel_button = QPushButton("❌ İptal")
        self.cancel_button.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                font-weight: bold;
                padding: 12px 24px;
                border-radius: 5px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        self.cancel_button.clicked.connect(self.reject)

        self.save_button = QPushButton("✅ Güncelle")
        self.save_button.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                font-weight: bold;
                padding: 12px 24px;
                border-radius: 5px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #229954;
            }
            QPushButton:disabled {
                background-color: #95a5a6;
            }
        """)
        self.save_button.clicked.connect(self.accept)

        button_layout.addWidget(self.cancel_button)
        button_layout.addWidget(self.save_button)
        layout.addLayout(button_layout)

        # Initial validation
        self.validate_all()

    def validate_spec(self):
        """SPEC validation: 'None' olamaz ve boş olamaz"""
        text = self.spec_edit.text().strip()

        if text == '' or text.lower() in ['none', 'n/a', 'null']:
            self.spec_validation_label.setText("⚠️ SPEC boş olamaz ve 'None' olamaz")
            return False
        else:
            self.spec_validation_label.setText("")
            return True

    def validate_siparis_no(self):
        """Sipariş No validation: sadece rakam, max 10 hane"""
        text = self.siparis_no_edit.text()

        if not text:
            self.siparis_validation_label.setText("⚠️ Sipariş No boş olamaz")
            return False
        elif not text.isdigit():
            self.siparis_validation_label.setText("⚠️ Sadece rakam girebilirsiniz")
            return False
        elif len(text) > 10:
            self.siparis_validation_label.setText("⚠️ En fazla 10 hane olabilir")
            return False
        else:
            self.siparis_validation_label.setText("")
            return True

    def validate_sip_kalem_no(self):
        """Sip Kalem No validation: 000XXX format, 6 hane"""
        text = self.sip_kalem_no_edit.text()

        if not text:
            self.kalem_validation_label.setText("⚠️ Sip Kalem No boş olamaz")
            return False
        elif len(text) != 6:
            self.kalem_validation_label.setText("⚠️ Tam 6 hane olmalı")
            return False
        elif not text.isdigit():
            self.kalem_validation_label.setText("⚠️ Sadece rakam girebilirsiniz")
            return False
        elif not text.startswith("000"):
            self.kalem_validation_label.setText("⚠️ İlk 3 hane '000' olmalı")
            return False
        else:
            self.kalem_validation_label.setText("")
            return True

    def validate_all(self):
        """Tüm alanları validate et ve buton durumunu güncelle"""
        spec_valid = self.validate_spec()
        siparis_valid = self.validate_siparis_no()
        kalem_valid = self.validate_sip_kalem_no()

        self.save_button.setEnabled(spec_valid and siparis_valid and kalem_valid)

    def get_updated_data(self):
        """Güncellenen verileri döndür"""
        return {
            'Malzeme_Adi': self.malzeme_adi_edit.text().strip(),
            'SPEC': self.spec_edit.text().strip(),
            'Siparis_No': self.siparis_no_edit.text().strip(),
            'Sip_Kalem_No': self.sip_kalem_no_edit.text().strip()
        }

class MusteriBilgileriDialog(QDialog):
    """
    Müşteri Bilgileri Gösterme Dialog Penceresi

    Bu sınıf, sözleşmedeki müşteri bilgilerini görüntüleyen ve
    kullanıcıya yeni cari oluşturma veya iptal etme seçenekleri sunan
    bir dialog penceresidir.

    Özellikler:
        - Müşteri bilgilerini grup halinde gösterir:
            * Ad, TCKN, Telefon
            * Adres bilgileri
        - İki buton seçeneği:
            * Yeni Oluştur: Yeni cari hesap oluşturur
            * İptal: İşlemi iptal eder
        - Ekran boyutuna göre dinamik boyutlandırma (%40 genişlik)

    Attributes:
        action (str): Kullanıcının seçtiği aksiyon ('new' veya 'cancel')
    """

    def __init__(self, customer_info, parent=None):
        """
        MusteriBilgileriDialog yapıcı metodu.

        Args:
            customer_info (dict): Müşteri bilgilerini içeren sözlük
                - ad: Müşteri adı
                - tckn: TCKN/Vergi numarası
                - telefon: Telefon numarası
                - adres: Adres bilgisi
            parent (QWidget, optional): Üst pencere. Varsayılan None.
        """
        super().__init__(parent)
        self.setWindowTitle("Sözleşmedeki Müşteri Bilgileri")

        # Dinamik pencere boyutu
        screen = QApplication.primaryScreen()
        screen_geometry = screen.availableGeometry()
        width = int(screen_geometry.width() * 0.4)  # Ekran genişliğinin %40'ı
        self.resize(width, 400)  # Yükseklik artırıldı

        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(15, 15, 15, 15)
        self.layout.setSpacing(15)  # Boşluklar artırıldı

        # Bilgi Grupları
        info_groups = [
            ("Müşteri Bilgileri", [
                f"<b>Adı:</b> {customer_info.get('ad', '')}",
                f"<b>TCKN:</b> {customer_info.get('tckn', '')}",
                f"<b>Telefon:</b> {customer_info.get('telefon', '')}"
            ]),
            ("Adres Bilgileri", [
                f"<b>Adres:</b> {customer_info.get('adres', '')}"
            ])
        ]

        for title, items in info_groups:
            group_box = QGroupBox(title)
            group_box.setStyleSheet("""
                QGroupBox {
                    font-size: 16px;
                    font-weight: bold;
                    border: 1px solid #aaa;
                    border-radius: 5px;
                    margin-top: 10px;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    left: 10px;
                    padding: 0px 3px;
                }
                QLabel {
                    font-size: 15px;
                }
            """)

            group_layout = QVBoxLayout()
            group_layout.setSpacing(10)  # Boşluk artırıldı

            for item in items:
                label = QLabel(item)
                label.setStyleSheet("font-size: 15px;")  # Boyut artırıldı
                label.setTextFormat(Qt.RichText)
                label.setWordWrap(True)
                group_layout.addWidget(label)

            group_box.setLayout(group_layout)
            self.layout.addWidget(group_box)

        # Butonlar için yatay layout
        button_layout = QHBoxLayout()

        # Kapat butonu
        btn_kapat = QPushButton("Kapat")
        btn_kapat.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                font-weight: bold;
                padding: 10px;
                border-radius: 5px;
                min-width: 120px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
        """)
        btn_kapat.clicked.connect(self.close)

        # Yeni Kayıt butonu
        btn_yeni_kayit = QPushButton("Yeni Kayıt")
        btn_yeni_kayit.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                padding: 10px;
                border-radius: 5px;
                min-width: 120px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #388E3C;
            }
        """)
        btn_yeni_kayit.clicked.connect(lambda: self.accept_with_action("new"))

        button_layout.addWidget(btn_kapat)
        button_layout.addWidget(btn_yeni_kayit)
        self.layout.addLayout(button_layout)

    def accept_with_action(self, action):
        self.action = action
        self.accept()

class CariSelectionDialog(QDialog):
    """
    Cari Hesap Seçim Dialog Penceresi

    Bu sınıf, veritabanında bulunan cari hesapları listeleyen ve kullanıcıya
    seçim yapma, güncelleme veya yeni oluşturma seçenekleri sunan bir dialog
    penceresidir.

    Özellikler:
        - Eşleşen cari hesapları tablo formatında gösterir
        - Sözleşme müşteri bilgilerini üst kısımda gösterir
        - TCKN eşleşmesi varsa özel vurgu yapar
        - Üç seçenek buton:
            * Güncelle: Seçili cariyi günceller
            * Yeni Oluştur: Yeni cari hesap oluşturur
            * İptal: İşlemi iptal eder
        - Tablo sütunları: Cari Kodu, Adı, TCKN, Telefon, Telefon2, Adres

    Attributes:
        records (list): Cari hesap kayıtları listesi
        customer_info (dict): Sözleşme müşteri bilgileri
        tckn_matched (bool): TCKN eşleşmesi var mı?
        action (str): Kullanıcının seçtiği aksiyon ('update', 'new', 'cancel')
        _selected_record: Seçili cari kayıt
        _selected_address_recno: Seçili adres kayıt numarası
    """

    def __init__(self, records, message, parent=None, customer_info=None, tckn_matched=False):
        """
        CariSelectionDialog yapıcı metodu.

        Args:
            records (list): Cari hesap kayıtları listesi (dict formatında)
            message (str): Dialog başlık mesajı (eşleşme tipi bilgisi)
            parent (QWidget, optional): Üst pencere
            customer_info (dict, optional): Sözleşme müşteri bilgileri
            tckn_matched (bool, optional): TCKN eşleşmesi var mı?
        """
        super().__init__(parent)
        self.setWindowTitle("🔍 Müşteri Kayıt Eşleştirme")
        self.records = records
        self.customer_info = customer_info

        # Daha uygun pencere boyutu
        screen = QApplication.primaryScreen()
        screen_geometry = screen.availableGeometry()
        width = min(1200, int(screen_geometry.width() * 0.8))
        height = min(800, int(screen_geometry.height() * 0.8))

        # Pencereyi merkeze yerleştir
        x = (screen_geometry.width() - width) // 2
        y = (screen_geometry.height() - height) // 2
        self.setGeometry(x, y, width, height)

        # Sade ve temiz CSS tema 
        self.setStyleSheet("""
            QDialog {
                background-color: #ffffff;
                color: #000000;
                font-family: 'Segoe UI', Arial, sans-serif;
            }

            QGroupBox {
                font-size: 14px;
                font-weight: bold;
                border: 1px solid #4A90E2;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
                background-color: #ffffff;
                color: #333333;
            }

            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 3px 8px;
                color: white;
                background-color: #4A90E2;
                border-radius: 3px;
                font-size: 14px;
            }

            QLabel {
                font-size: 14px;
                color: #333333;
                background-color: transparent;
                padding: 5px;
            }

            QTableWidget {
                background-color: #ffffff;
                alternate-background-color: #f8f9fa;
                selection-background-color: #4A90E2;
                selection-color: white;
                border: 1px solid #cccccc;
                font-size: 13px;
                gridline-color: #dddddd;
                color: #000000;
            }

            QTableWidget::item {
                padding: 6px 8px;
                border-bottom: 1px solid #eeeeee;
                color: #000000;
                background-color: transparent;
            }

            QTableWidget::item:selected {
                background-color: #4A90E2;
                color: white;
            }

            QTableWidget::item:alternate {
                background-color: #f8f9fa;
                color: #000000;
            }

            QHeaderView::section {
                background-color: #4A5568;
                color: white;
                padding: 8px;
                font-size: 13px;
                font-weight: bold;
                border: 1px solid #3A4A58;
                text-align: center;
            }

            QPushButton {
                font-size: 14px;
                font-weight: bold;
                padding: 10px 20px;
                border: none;
                border-radius: 5px;
                min-width: 140px;
                margin: 3px;
            }

            QPushButton#update {
                background-color: #FF9500;
                color: white;
            }

            QPushButton#update:hover {
                background-color: #E6850E;
            }

            QPushButton#close {
                background-color: #DC3545;
                color: white;
            }

            QPushButton#close:hover {
                background-color: #C82333;
            }

            QPushButton#new {
                background-color: #28A745;
                color: white;
            }

            QPushButton#new:hover {
                background-color: #218838;
            }

            QRadioButton {
                font-size: 14px;
                padding: 5px;
            }

            QRadioButton::indicator {
                width: 16px;
                height: 16px;
            }

            QRadioButton::indicator:unchecked {
                border: 2px solid #999999;
                border-radius: 8px;
                background-color: white;
            }

            QRadioButton::indicator:checked {
                border: 2px solid #4A90E2;
                border-radius: 8px;
                background-color: #4A90E2;
            }
        """)

        # Ana layout
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)

        # Üst bilgi
        info_label = QLabel(message)
        info_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #333333; padding: 10px;")
        info_label.setWordWrap(True)
        main_layout.addWidget(info_label)

        # Müşteri bilgileri grupbox
        if customer_info:
            customer_group = QGroupBox("Sözleşmedeki Müşteri Bilgileri")
            customer_layout = QVBoxLayout(customer_group)

            # Müşteri bilgilerini düz metin olarak göster
            info_text = f"Ad: {customer_info.get('ad', '')}\n"
            info_text += f"TCKN: {customer_info.get('tckn', '')}\n"
            info_text += f"Telefon: {customer_info.get('telefon', '')}\n"
            info_text += f"Adres: {customer_info.get('adres', '')}"

            customer_label = QLabel(info_text)
            customer_label.setStyleSheet("padding: 10px; background-color: #f8f9fa; border-radius: 3px;")
            customer_layout.addWidget(customer_label)

            main_layout.addWidget(customer_group)

        # Tablo
        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setStyleSheet("""
            QTableWidget::item:focus {
                outline: none;
                border: none;
            }
        """)
        self.table.setFocusPolicy(Qt.NoFocus)
        
        # Ctrl+C Kısayolu
        self.copy_shortcut = QShortcut(QKeySequence("Ctrl+C"), self.table)
        self.copy_shortcut.activated.connect(self.copy_selection)
        
        main_layout.addWidget(self.table)

    def copy_selection(self):
        """Seçili hücreyi kopyalar"""
        selected_items = self.table.selectedItems()
        if selected_items:
            text = selected_items[0].text()
            QApplication.clipboard().setText(text)

        # Tabloyu doldur
        self.populate_table(records)

        # Alt buton paneli
        button_layout = QHBoxLayout()

        # Sol tarafta bilgi
        info_text = f"Toplam {len(records) if records else 0} kayıt bulundu"
        info_count = QLabel(info_text)
        info_count.setStyleSheet("color: #666666; font-size: 13px;")
        button_layout.addWidget(info_count)

        button_layout.addStretch()

        # Butonlar - Image #2 tarzında sade
        close_btn = QPushButton("Kapat")
        close_btn.setObjectName("close")
        close_btn.clicked.connect(self.reject)
        button_layout.addWidget(close_btn)

        update_btn = QPushButton("Seçili Kaydı Güncelle")
        update_btn.setObjectName("update")
        update_btn.clicked.connect(lambda: self.accept_with_action("update"))
        button_layout.addWidget(update_btn)

        new_btn = QPushButton("Yeni Kayıt Oluştur")
        new_btn.setObjectName("new")
        new_btn.clicked.connect(lambda: self.accept_with_action("new"))
        button_layout.addWidget(new_btn)

        main_layout.addLayout(button_layout)

    def populate_table(self, records):
        if not records:
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            # Boş tablo mesajı
            empty_label = QLabel("📋 Eşleşen kayıt bulunamadı")
            empty_label.setAlignment(Qt.AlignCenter)
            empty_label.setStyleSheet("font-size: 16px; color: #6c757d; padding: 50px;")
            return

        # Radio button grup oluştur
        self.address_radio_group = QButtonGroup()

        # İstenen sütun sıralaması: Seçim, Müşteri Adı, Adres, Sayaç, Adres Sayaç, Cari Kod, TCKN, Telefon, Telefon2, Açıklama, Adres Bilgisi
        column_order = ['cariAdi', 'Adres', 'sayac', 'adres_sayac', 'cariKod', 'TCKN', 'cariTelefon', 'Telefon2', 'cariAciklama', 'Adres Bilgisi']

        # Sütun isimleri ve genişlikleri için mapping
        column_mapping = {
            'sayac': {'title': 'Sayaç', 'width': 80},
            'cariAdi': {'title': 'Müşteri Adı', 'width': 200},
            'cariKod': {'title': 'Cari Kod', 'width': 120},
            'TCKN': {'title': 'TCKN/VKN', 'width': 120},
            'cariTelefon': {'title': 'Telefon', 'width': 120},
            'cariAciklama': {'title': 'Açıklama', 'width': 150},
            'Telefon2': {'title': 'Telefon2', 'width': 120},
            'adres_sayac': {'title': 'Adres Sayaç', 'width': 100},
            'Adres': {'title': 'Adres', 'width': 300},
            'Adres Bilgisi': {'title': 'Adres Bilgisi', 'width': 150}
        }

        # Kolon sayısını ayarla (Seçim kolonu + sıralanmış kolonlar)
        self.table.setColumnCount(len(column_order) + 1)
        self.table.setRowCount(len(records))

        # Başlıkları istenen sıraya göre ayarla
        headers = ["Seçim"] + [column_mapping.get(key, {'title': key})['title'] for key in column_order]
        self.table.setHorizontalHeaderLabels(headers)

        # Verileri doldur
        for row_idx, record in enumerate(records):
            # Radio button - sade ve küçük
            radio_btn = QRadioButton()
            if row_idx == 0:  # İlk satırı varsayılan seç
                radio_btn.setChecked(True)

            # Radio button'u merkeze yerleştir
            radio_container = QWidget()
            radio_layout = QHBoxLayout(radio_container)
            radio_layout.setContentsMargins(5, 5, 5, 5)
            radio_layout.setAlignment(Qt.AlignCenter)
            radio_layout.addWidget(radio_btn)

            self.address_radio_group.addButton(radio_btn, row_idx)
            self.table.setCellWidget(row_idx, 0, radio_container)

            # Verileri istenen sıraya göre ekle
            for col_idx, key in enumerate(column_order):
                value = record.get(key, "")
                display_value = str(value) if value is not None else ""

                item = QTableWidgetItem(display_value)
                self.table.setItem(row_idx, col_idx + 1, item)

        # Tablo ayarları
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSortingEnabled(False)

        # Sütun genişlik ayarları - kullanıcı spesifikasyonlarına göre
        horizontal_header = self.table.horizontalHeader()

        # Önce tüm sütunları ResizeToContents olarak ayarla ki içerik sığsın
        horizontal_header.setSectionResizeMode(QHeaderView.ResizeToContents)

        # Seçim (0): 55px - küçük sabit
        horizontal_header.setSectionResizeMode(0, QHeaderView.Fixed)
        self.table.setColumnWidth(0, 55)

        # Müşteri Adı (1): tam veri gösterecek dinamik genişlik
        horizontal_header.setSectionResizeMode(1, QHeaderView.ResizeToContents)

        # Adres (2): tam veri gösterecek dinamik genişlik - ÖNEMLİ
        horizontal_header.setSectionResizeMode(2, QHeaderView.ResizeToContents)

        # Sayaç (3): 60px - küçük sabit
        horizontal_header.setSectionResizeMode(3, QHeaderView.Fixed)
        self.table.setColumnWidth(3, 60)

        # Adres Sayaç (4): 60px - küçük sabit
        horizontal_header.setSectionResizeMode(4, QHeaderView.Fixed)
        self.table.setColumnWidth(4, 60)

        # Cari Kod (5): 100px - orta sabit
        horizontal_header.setSectionResizeMode(5, QHeaderView.Fixed)
        self.table.setColumnWidth(5, 100)

        # TCKN (6): 110px - orta sabit
        horizontal_header.setSectionResizeMode(6, QHeaderView.Fixed)
        self.table.setColumnWidth(6, 110)

        # Telefon (7): 100px - orta sabit
        horizontal_header.setSectionResizeMode(7, QHeaderView.Fixed)
        self.table.setColumnWidth(7, 100)

        # Telefon2 (8): 100px - orta sabit
        horizontal_header.setSectionResizeMode(8, QHeaderView.Fixed)
        self.table.setColumnWidth(8, 100)

        # Açıklama (9): 140px - orta, yeniden boyutlandırılabilir
        horizontal_header.setSectionResizeMode(9, QHeaderView.Interactive)
        self.table.setColumnWidth(9, 140)

        # Adres Bilgisi (10): 140px - orta, yeniden boyutlandırılabilir
        horizontal_header.setSectionResizeMode(10, QHeaderView.Interactive)
        self.table.setColumnWidth(10, 140)

        # Satır ayarları - Image #1 tarzında
        self.table.verticalHeader().setDefaultSectionSize(32)
        self.table.verticalHeader().hide()

        # Radio button değişikliklerini dinle
        self.address_radio_group.buttonClicked.connect(self.on_selection_changed)

    def on_selection_changed(self, button):
        """Seçim değiştiğinde tabloyu güncelle"""
        selected_row = self.address_radio_group.id(button)
        self.table.selectRow(selected_row)

    def accept_with_action(self, action):
        self.action = action
        self.accept()

    def selected_record(self):
        if hasattr(self, 'address_radio_group'):
            selected_row = self.address_radio_group.checkedId()
            return selected_row if selected_row >= 0 else None
        return None

    def get_selected_address_recno(self):
        """Seçili adresin RECno'sunu al"""
        if hasattr(self, 'address_radio_group'):
            selected_row = self.address_radio_group.checkedId()
            if selected_row >= 0:
                # adres_sayac sütununu bul ve değeri al - "Adres Sayaç" sütunu 4. indekste (0'dan başlayarak)
                # Sütun düzeni: Seçim(0), Müşteri Adı(1), Adres(2), Sayaç(3), Adres Sayaç(4), Cari Kod(5)...
                recno_item = self.table.item(selected_row, 4)  # adres_sayac kolonu
                if recno_item and recno_item.text() and recno_item.text() != 'None':
                    return int(recno_item.text())
        return None
   

class SozlesmeApp(QMainWindow):
    """
    Sözleşme Yönetim Ana Uygulama Penceresi

    Bu sınıf, sözleşme verilerinin yönetildiği ana penceredir.
    Sözleşmeleri listeleme, filtreleme, arama, import/export ve detay görüntüleme
    özelliklerini içerir.

    Ana Özellikler:
        - Sözleşme listesi tablosu (13 sütun)
        - Ay/Yıl bazlı filtreleme
        - Arama ve temizleme fonksiyonları
        - Excel import/export
        - Sözleşme detay görüntüleme
        - Yüzde hesaplama ve renklendirme
        - Sağ tık menüsü ile kopyalama
        - Hücre düzenlemesi

    Tablo Sütunları:
        1. Satış Kodu (SAP)
        2. Sipariş Tarihi
        3. Fiyat Listesi
        4. Header
        5. Mağaza
        6. Cari Kod
        7. TCKN
        8. Telefon
        9. Telefon2
        10. Açıklama
        11. Adres
        12. %
        13. Durum

    Attributes:
        original_df (DataFrame): Orijinal sözleşme verileri
        filtered_df (DataFrame): Filtrelenmiş sözleşme verileri
        current_month (str): Mevcut ay (İngilizce)
        previous_month (str): Önceki ay (İngilizce)
        month_names_tr (dict): İngilizce-Türkçe ay isimleri eşleştirmesi
        siparis_calisiyor (bool): Sipariş programı çalışıyor mu?
    """

    def __init__(self):
        """
        SozlesmeApp sınıfının yapıcı metodu.

        Uygulamanın başlangıç ayarlarını yapar, arayüzü oluşturur ve
        verileri yükler.
        """
        super().__init__()
        # Veri çerçeveleri
        self.original_df = pd.DataFrame()  # Orijinal veriler
        self.filtered_df = pd.DataFrame()  # Filtrelenmiş veriler
        self.base_filtered_df = pd.DataFrame()  # Ana filtre (ay/yıl) sonrası veriler

        # Tarih bilgileri
        self.current_month = datetime.now().strftime("%B")  # Mevcut ay (İngilizce)
        self.previous_month = (datetime.now().replace(day=1) - timedelta(days=1)).strftime("%B")  # Önceki ay
        
        # Türkçe ay isimleri
        self.month_names_tr = {
            'January': 'Ocak', 'February': 'Şubat', 'March': 'Mart',
            'April': 'Nisan', 'May': 'Mayıs', 'June': 'Haziran',
            'July': 'Temmuz', 'August': 'Ağustos', 'September': 'Eylül',
            'October': 'Ekim', 'November': 'Kasım', 'December': 'Aralık'
        }

        # Ay isimlerini numaraya çevirme için mapping
        self.month_to_num = {
            'Ocak': '01', 'Şubat': '02', 'Mart': '03', 'Nisan': '04',
            'Mayıs': '05', 'Haziran': '06', 'Temmuz': '07', 'Ağustos': '08',
            'Eylül': '09', 'Ekim': '10', 'Kasım': '11', 'Aralık': '12'
        }

        self.siparis_calisiyor = False  # Siparis program çalışma durumu
        self._data_loaded = False  # Lazy loading için flag
        self.setup_ui()
        self.show()

    def showEvent(self, event):
        """Widget ilk gösterildiğinde veri yükle (lazy loading)"""
        super().showEvent(event)
        if not self._data_loaded:
            self._data_loaded = True
            # UI render olduktan sonra veri yükle
            QTimer.singleShot(100, self.load_initial_data)
    
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

        # Arama ve Buton Alanı
        search_layout = QHBoxLayout()

        # Bu Ay Butonu
        month_name = self.month_names_tr.get(self.current_month, self.current_month)
        self.current_month_btn = QPushButton(month_name)
        self.current_month_btn.setStyleSheet("""
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
        self.current_month_btn.setToolTip(f"{self.month_names_tr.get(self.current_month, self.current_month)} ayı siparişlerini gösterir")
        self.current_month_btn.clicked.connect(self.load_current_month_data)
        search_layout.addWidget(self.current_month_btn)

        # Hepsi Butonu
        self.all_btn = QPushButton("Hepsi")
        self.all_btn.setStyleSheet("""
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
        self.all_btn.setToolTip("Tüm siparişleri gösterir")
        self.all_btn.clicked.connect(self.load_data_all)
        search_layout.addWidget(self.all_btn)

        # Yıl Seçimi (dinamik)
        self.year_combo = QComboBox()
        self.year_combo.setEnabled(False)
        self.year_combo.setStyleSheet("""
            QComboBox {
                min-width: 80px;
                padding: 8px;
                font-size: 16px;
                font-weight: bold;
                border: 1px solid #444;
                border-radius: 3px;
            }
            QComboBox::drop-down {
                width: 20px;
            }
            QComboBox::down-arrow {
                width: 10px;
                height: 10px;
            }
        """)
        self.year_combo.currentTextChanged.connect(self.filter_by_year_month)
        search_layout.addWidget(self.year_combo)

        # Ay Seçimi  
        self.month_combo = QComboBox()
        months = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
                 "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
        self.month_combo.addItems(["Tümü"] + months)
        self.month_combo.setEnabled(False)
        self.month_combo.setStyleSheet("""
            QComboBox {
                min-width: 100px;
                padding: 8px;
                font-size: 16px;
                font-weight: bold;
                border: 1px solid #444;
                border-radius: 3px;
            }
            QComboBox::drop-down {
                width: 20px;
            }
            QComboBox::down-arrow {
                width: 10px;
                height: 10px;
            }
        """)
        self.month_combo.currentTextChanged.connect(self.filter_by_year_month)
        search_layout.addWidget(self.month_combo)

        # Aktar Butonu
        self.import_btn = QPushButton("Aktar")
        self.import_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: 1px solid #444;
                border-radius: 5px;
                padding: 8px 16px;
                font-size:14px;
                font-weight: bold;
                min-width: 80px;
            }
            QPushButton:hover {
                background-color: #388E3C;
            }
        """)
        self.import_btn.setToolTip("Sözleşme numarasını sorgular")
        self.import_btn.clicked.connect(self.import_contract_data)
        search_layout.addWidget(self.import_btn)

        # Arama Kutusu
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Müşteri Adı, Malzeme Adı Ara & Şözleşme ID Gir...")
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
        self.clear_btn.setToolTip("Arama kutusunu temizler")
        self.clear_btn.clicked.connect(self.clear_search)
        search_layout.addWidget(self.clear_btn)

        # Toplam Tutar Label
        self.total_amount_label = QLabel("0 ₺")
        self.total_amount_label.setStyleSheet("""
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
        """)
        search_layout.addWidget(self.total_amount_label)
        
        # Header Toplamı Label
        self.header_total_label = QLabel("0 ₺")
        self.header_total_label.setStyleSheet("""
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
        """)
        search_layout.addWidget(self.header_total_label)
        
        # Header Yüzde Label
        self.header_percentage_label = QLabel("%0")
        self.header_percentage_label.setStyleSheet("""
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
        """)
        search_layout.addWidget(self.header_percentage_label)
        
        # Benzersiz Sipariş No Sayısı Label
        self.unique_orders_label = QLabel("0")
        self.unique_orders_label.setStyleSheet("""
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
        """)
        search_layout.addWidget(self.unique_orders_label)

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
        self.excel_btn.clicked.connect(self.export_to_excel)
        search_layout.addWidget(self.excel_btn)

        # Kontrol Butonu
        self.kontrol_btn = QPushButton("Kontrol")
        self.kontrol_btn.setStyleSheet("""
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
        self.kontrol_btn.setToolTip("GoogleDrive/PRG/Kontrol.xlsx dosyasını güncelleyin...")
        self.kontrol_btn.clicked.connect(self.control_excel_data)
        search_layout.addWidget(self.kontrol_btn)

        self.main_layout.addLayout(search_layout)

        # Tablo
        self.table = QTableWidget()
        self.table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        self.table.itemChanged.connect(self.handle_cell_edit)
        self.table.verticalHeader().setDefaultSectionSize(self.table.verticalHeader().defaultSectionSize() + 2)

        # Sütun başlıklarına tıklanabilirlik ekle
        self.table.horizontalHeader().sectionClicked.connect(self.filter_column_header)
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

        # Ctrl+C Kısayolu Ekle
        self.copy_shortcut = QShortcut(QKeySequence("Ctrl+C"), self.table)
        self.copy_shortcut.activated.connect(self.handle_ctrl_c)

        self.main_layout.addWidget(self.table)

        self.main_layout.addWidget(self.table)

        # Performans Optimizasyonları
        self.filter_timer = QTimer()
        self.filter_timer.setSingleShot(True)
        self.filter_timer.timeout.connect(self.filter_data)

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
        self.progress_bar.setTextVisible(True)  # Yüzde metnini göster
        self.progress_bar.setAlignment(Qt.AlignCenter)  # Metni ortala
        self.progress_bar.setFormat("%p%")  # Yüzde formatı
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

    def get_connection(self):
        # PRGsheet/Ayar'dan SQL bağlantı bilgilerini yükle
        config_manager = CentralConfigManager()
        settings = config_manager.get_settings()
        server = settings.get('SQL_SERVER')
        database = settings.get('SQL_DATABASE')
        username = settings.get('SQL_USERNAME')
        password = settings.get('SQL_PASSWORD')

        if not all([server, database, username, password]):
            raise Exception("PRGsheet/Ayar sayfasında SQL bağlantı bilgileri eksik")

        # Bağlantı dizesini oluşturun
        connection_string = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'

        # Veritabanına bağlanın
        return pyodbc.connect(connection_string)

    def load_current_month_data(self):
        """Bu ayın siparişlerini yükle (Siparis.exe çalıştırarak)"""
        try:
            exe_path = r"D:\GoogleDrive\PRG\EXE\Siparis.exe"
            if not os.path.exists(exe_path):
                QMessageBox.warning(self, "Uyarı", f"Siparis.exe bulunamadı: {exe_path}")
                self.load_data_by_month(self.current_month)
                return
            
            self.status_label.setText("🔄 Siparis.exe çalıştırılıyor...")
            self.current_month_btn.setEnabled(False)
            self.siparis_calisiyor = True
            
            QApplication.processEvents()
            
            os.startfile(exe_path)
            
            # Siparis.exe'nin çalışması için bekleme
            # 10 saniye sonra program bitmiş sayıp kontrol et
            QTimer.singleShot(10000, self.on_siparis_finished)
            
        except Exception as e:
            self.status_label.setText(f"❌ Program çalıştırma hatası: {str(e)}")
            QMessageBox.critical(self, "Hata", f"Program çalıştırma hatası: {str(e)}")
            self.current_month_btn.setEnabled(True)
            self.siparis_calisiyor = False
            self.load_data_by_month(self.current_month)

    def on_siparis_finished(self):
        """Siparis program bittikten sonra"""
        self.current_month_btn.setEnabled(True)
        self.siparis_calisiyor = False
        self.status_label.setText("✅ Siparis.exe tamamlandı, Google Sheets güncelleme bekleniyor...")
        
        # Google Sheets'e kaydedilmesi için ek bekleme (7 saniye)
        # Sonra verileri yenile
        QTimer.singleShot(7000, self.delayed_current_month_refresh)
    
    def delayed_current_month_refresh(self):
        """Gecikmeli bu ay veri yenileme"""
        self.status_label.setText("🔄 Google Sheets'ten güncel veriler alınıyor...")
        QApplication.processEvents()
        self.load_data_by_month(self.current_month)
        self.status_label.setText(f"✅ {len(self.original_df)} {self.month_names_tr.get(self.current_month, self.current_month)} siparişi yüklendi")

    def load_data_by_month(self, month_name):
        """Belirli bir ayın siparişlerini yükle (Siparisler sayfasından)"""
        try:
            self.status_label.setText(f"📊 {self.month_names_tr.get(month_name, month_name)} ayı siparişleri yükleniyor...")
            QApplication.processEvents()

            # PRGsheet/Ayar sayfasından SPREADSHEET_ID'yi yükle

            spreadsheet_id = CentralConfigManager().MASTER_SPREADSHEET_ID

            # Google Sheets URL'sini oluştur
            gsheets_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"
            
            self.status_label.setText("🔗 Google Sheets'e bağlanıyor...")
            QApplication.processEvents()
            
            # URL'den Excel dosyasını oku
            response = requests.get(gsheets_url, timeout=30)
            
            if response.status_code == 401:
                self.status_label.setText("❌ Google Sheets erişim hatası: Dosya özel veya izin gerekli")
                logging.error("Google Sheets erişim hatası: Dosya özel veya izin gerekli")
                return
            elif response.status_code != 200:
                self.status_label.setText(f"❌ HTTP Hatası: {response.status_code} - {response.reason}")
                logging.error(f"HTTP Hatası: {response.status_code} - {response.reason}")
                return
            
            response.raise_for_status()
            
            # Siparisler sayfasını oku
            self.status_label.setText(f"📋 {self.month_names_tr.get(month_name, month_name)} siparişleri işleniyor...")
            QApplication.processEvents()
            
            self.original_df = pd.read_excel(BytesIO(response.content), sheet_name="Siparis")
            
            # Veri temizleme işlemleri
            self.clean_data()

            # Tarih sütununu datetime'a çevir
            if 'Tarih' in self.original_df.columns:
                self.original_df['Tarih'] = pd.to_datetime(self.original_df['Tarih'], errors='coerce')
                
                # Belirli ayın verilerini filtrele
                month_filter = self.original_df['Tarih'].dt.strftime('%B') == month_name
                self.original_df = self.original_df[month_filter]
                
                # Tarih sütununu tekrar string formatına döndür (00:00:00 olmadan)
                self.original_df['Tarih'] = self.original_df['Tarih'].dt.strftime('%Y-%m-%d')

            self.status_label.setText("🔄 Tablo güncelleniyor...")
            QApplication.processEvents()

            self.filtered_df = self.original_df.copy()
            self.base_filtered_df = self.filtered_df.copy()  # Ana filtreyi sakla
            self.update_table()

            self.status_label.setText(f"✅ {len(self.original_df)} {self.month_names_tr.get(month_name, month_name)} siparişi yüklendi")
                
        except Exception as e:
            logging.error(f"Hata: {str(e)}")
            self.status_label.setText(f"❌ Veri yükleme hatası: {str(e)}")
            QMessageBox.critical(self, "Hata", f"Veri yükleme hatası: {str(e)}")

    def load_data_all(self):
        """Tüm siparişleri yükle (Siparisler sayfasından)"""
        try:
            self.status_label.setText("📊 Tüm siparişler yükleniyor...")
            QApplication.processEvents()

            # PRGsheet/Ayar sayfasından SPREADSHEET_ID'yi yükle

            spreadsheet_id = CentralConfigManager().MASTER_SPREADSHEET_ID

            # Google Sheets URL'sini oluştur
            gsheets_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"
            
            self.status_label.setText("🔗 Google Sheets'e bağlanıyor...")
            QApplication.processEvents()
            
            # URL'den Excel dosyasını oku
            response = requests.get(gsheets_url, timeout=30)
            
            if response.status_code == 401:
                self.status_label.setText("❌ Google Sheets erişim hatası: Dosya özel veya izin gerekli")
                logging.error("Google Sheets erişim hatası: Dosya özel veya izin gerekli")
                return
            elif response.status_code != 200:
                self.status_label.setText(f"❌ HTTP Hatası: {response.status_code} - {response.reason}")
                logging.error(f"HTTP Hatası: {response.status_code} - {response.reason}")
                return
            
            response.raise_for_status()
            
            # Siparisler sayfasını oku
            self.status_label.setText("📋 Tüm siparişler işleniyor...")
            QApplication.processEvents()
            
            self.original_df = pd.read_excel(BytesIO(response.content), sheet_name="Siparisler")
            
            # Veri temizleme işlemleri
            self.clean_data()
            
            # Yıl ve ay filtrelerini etkinleştir ve yılları doldur
            self.populate_years_from_data()
            self.year_combo.setEnabled(True)
            self.month_combo.setEnabled(True)

            self.status_label.setText("🔄 Tablo güncelleniyor...")
            QApplication.processEvents()

            self.filtered_df = self.original_df.copy()
            self.base_filtered_df = self.filtered_df.copy()  # Ana filtreyi sakla
            self.update_table()

            self.status_label.setText(f"✅ {len(self.original_df)} toplam sipariş yüklendi")
                
        except Exception as e:
            logging.error(f"Hata: {str(e)}")
            self.status_label.setText(f"❌ Veri yükleme hatası: {str(e)}")
            QMessageBox.critical(self, "Hata", f"Veri yükleme hatası: {str(e)}")

    def clear_search(self):
        """Arama kutusunu temizle"""
        self.search_box.clear()
        self.filtered_df = self.original_df.copy()
        self.update_table()

    def update_summary_labels(self):
        """Özet bilgileri güncelle"""
        try:
            if self.filtered_df.empty:
                self.total_amount_label.setText("0 ₺")
                self.header_total_label.setText("0 ₺")
                self.header_percentage_label.setText("%0")
                self.unique_orders_label.setText("0")
                return
                
            # Toplam tutarı hesapla: Birim Fiyat * Miktar + Vergi (integer kullanarak)
            total_amount = 0
            for _, row in self.filtered_df.iterrows():
                try:
                    birim_fiyat = self.convert_turkish_to_integer(row.get('Birim Fiyat', 0))
                    miktar = pd.to_numeric(row.get('Miktar', 0), errors='coerce') or 0
                    vergi = self.convert_turkish_to_integer(row.get('Vergi', 0))

                    satir_toplam = int(round(birim_fiyat * miktar)) + vergi
                    total_amount += satir_toplam
                except:
                    continue
            
            # Header sütunu toplamını hesapla (integer kullanarak)
            header_total = 0
            if 'Header' in self.filtered_df.columns:
                for _, row in self.filtered_df.iterrows():
                    try:
                        header_value = self.convert_turkish_to_integer(row.get('Header', 0))
                        header_total += header_value
                    except (ValueError, TypeError, AttributeError):
                        continue
            
            # Header yüzdesini hesapla
            header_percentage = 0
            if total_amount > 0:
                header_percentage = (1-(header_total / total_amount)) * 100
            
            # Benzersiz Sipariş No sayısını hesapla
            unique_orders = 0
            if 'Sipariş No' in self.filtered_df.columns:
                unique_orders = self.filtered_df['Sipariş No'].nunique()
            
            # Labels'ı güncelle - integer format
            self.total_amount_label.setText(f"{int(total_amount):,} ₺")
            self.header_total_label.setText(f"{int(header_total):,} ₺")
            self.header_percentage_label.setText(f"%{header_percentage:.0f}")
            self.unique_orders_label.setText(f"{unique_orders}")
            
        except Exception as e:
            logging.error(f"Özet bilgi hesaplama hatası: {str(e)}")
            self.total_amount_label.setText("0 ₺")
            self.header_total_label.setText("0 ₺")
            self.header_percentage_label.setText("%0")
            self.unique_orders_label.setText("0")

    def populate_years_from_data(self):
        """Veri setindeki tarihlerden yılları alıp yıl combo'yu doldur"""
        try:
            if 'Tarih' in self.original_df.columns:
                # Tarih sütununu datetime'a çevir
                temp_df = self.original_df.copy()
                temp_df['Tarih'] = pd.to_datetime(temp_df['Tarih'], errors='coerce')
                
                # Benzersiz yılları al ve sırala
                unique_years = sorted(temp_df['Tarih'].dt.year.dropna().unique())
                
                # Yıl combo'yu temizle ve yeniden doldur
                self.year_combo.clear()
                
                # Önce "Tümü" seçeneğini ekle
                self.year_combo.addItem("Tümü")
                
                # Yılları ekle
                year_strings = [str(int(year)) for year in unique_years]
                self.year_combo.addItems(year_strings)
                
                # "Tümü" seçeneğini varsayılan olarak seç
                self.year_combo.setCurrentText("Tümü")
                    
        except Exception as e:
            logging.error(f"Yıl listesi oluşturma hatası: {str(e)}")
            # Hata durumunda varsayılan değerler
            self.year_combo.clear()
            self.year_combo.addItem("Tümü")
            self.year_combo.setCurrentText("Tümü")

    def schedule_filter(self):
        self.filter_timer.stop()
        self.filter_timer.start(200)

    def filter_data(self):
        try:
            search_text = self.search_box.text().strip().lower()

            if not search_text:
                self.filtered_df = self.original_df.copy()
            else:
                # Kelime bazlı arama için parçalara ayır
                parts = [re.escape(part) for part in search_text.split() if part]
                pattern = r'(?=.*?{})'.format(')(?=.*?'.join(parts))
                
                # Belirli sütunlarda ara: Malzeme Adı, Cari Adi, Aciklama
                search_columns = ['Malzeme Adı', 'Cari Adi', 'Aciklama', 'Sozlesme']
                
                # Mask'i original_df'in index yapısını kullanarak oluştur
                mask = pd.Series(False, index=self.original_df.index)
                
                for col in search_columns:
                    if col in self.original_df.columns:
                        mask |= self.original_df[col].astype(str).str.lower().str.contains(pattern, regex=True, na=False)
                
                self.filtered_df = self.original_df[mask].copy()

            self.update_table()

        except Exception as e:
            logging.error(f"Filtreleme hatası: {str(e)}")

    def update_table(self):
        self.table.blockSignals(True)
        self.table.clearContents()

        if self.filtered_df.empty:
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            self.table.blockSignals(False)
            return

        # Sütun sıralaması
        column_order = [
            'Satir', 'Sipariş No','Tarih', 'Sozlesme', '%', 'Cari Adi', 'Malzeme Kodu', 
            'Malzeme Adı', 'Miktar', 'Teslimat','Depo', 'Personel', 'Mağaza', 
            'Birim Fiyat', 'Vergi', 'Iskonto', 'Header', 'Aciklama'
        ]
        
        # Gözükmesini istemediğimiz sütunlar
        hidden_columns = ['VergiKod', 'Cari Kod']
        
        
        # Yüzde sütunu hesapla ve ekle
        temp_df = self.filtered_df.copy()
        temp_df['%'] = temp_df.apply(self.calculate_percentage_for_row, axis=1)
        
        # Mevcut sütunları kontrol et ve sıralama yap
        available_columns = [col for col in column_order if col in temp_df.columns]
        # Sıralamada olmayan sütunları sona ekle (gizli sütunlar hariç)
        remaining_columns = [col for col in temp_df.columns if col not in column_order and col not in hidden_columns]
        final_columns = available_columns + remaining_columns
        
        # DataFrame'i sırala
        display_df = temp_df[final_columns].copy()

        # Sipariş No'ya göre büyükten küçüğe, sonra Satir'a göre küçükten büyüğe sırala
        if 'Sipariş No' in display_df.columns and 'Satir' in display_df.columns:
            display_df = display_df.sort_values(by=['Sipariş No', 'Satir'], ascending=[False, True], na_position='last')

        # Aynı Sipariş No'su olan satırlarda sadece Satir=1 olanların belirli sütunlarını göster
        hide_columns = ['Sipariş No','Tarih', 'Sozlesme', '%','Personel', 'Mağaza', 'Header']
        
        if 'Sipariş No' in display_df.columns and 'Satir' in display_df.columns:
            # Satir != 1 olan satırlar için belirtilen sütunları boşalt
            mask = display_df['Satir'] != 1
            for col in hide_columns:
                if col in display_df.columns:
                    display_df.loc[mask, col] = ""

        rows, cols = display_df.shape
        self.table.setRowCount(rows)
        self.table.setColumnCount(cols)
        self.table.setHorizontalHeaderLabels(display_df.columns)

        for i in range(rows):
            # Miktar = Teslimat kontrolü için satır verilerini al
            row_data = display_df.iloc[i]
            miktar_equals_teslimat = False
            
            if 'Miktar' in display_df.columns and 'Teslimat' in display_df.columns:
                miktar = pd.to_numeric(row_data.get('Miktar', 0), errors='coerce')
                teslimat = pd.to_numeric(row_data.get('Teslimat', 0), errors='coerce')
                miktar_equals_teslimat = (not pd.isna(miktar) and not pd.isna(teslimat) and 
                                        miktar == teslimat and miktar > 0)
            
            for j in range(cols):
                col_name = display_df.columns[j]
                value = display_df.iat[i, j]
                
                if pd.isna(value) or value is None:
                    display_value = ""
                else:
                    # Birim Fiyat, Vergi, Iskonto sütunları için integer formatı
                    if col_name in ['Birim Fiyat', 'Vergi', 'Iskonto'] and isinstance(value, (int, float)):
                        display_value = f"{int(round(value)):,}"
                    # Header sütunu için integer formatı
                    elif col_name == 'Header' and isinstance(value, (int, float)):
                        if value == 0:
                            display_value = ""
                        else:
                            display_value = f"{int(value):,}"
                    else:
                        display_value = str(value)
                
                item = QTableWidgetItem(display_value)

                # Sepet sütunu düzenlenebilir yap
                if col_name == 'Sepet':
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsEditable)
                else:
                    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                
                # Miktar = Teslimat olan satırlarda belirli sütunları yeşil renk yap
                if (miktar_equals_teslimat and 
                    col_name in ['Malzeme Kodu', 'Malzeme Adı', 'Miktar', 'Teslimat']):
                    item.setBackground(QColor(233, 252, 233))  # Açık yeşil
                
                # Yüzde sütunu için renk kodlaması
                if col_name == '%' and display_value and display_value != '':
                    try:
                        percentage_value = int(display_value.replace('%', ''))
                        if percentage_value < 17:
                            item.setForeground(QColor(231, 76, 60))  # Kırmızı
                        elif percentage_value >= 27:
                            item.setForeground(QColor(39, 174, 96))  # Yeşil
                        else:
                            item.setForeground(QColor(0, 0, 0))  # Siyah
                    except:
                        pass
                
                self.table.setItem(i, j, item)

        self.table.resizeColumnsToContents()
        self.table.blockSignals(False)

        # Özet bilgileri güncelle
        self.update_summary_labels()

        # Dialog açılmasını devre dışı bırak - tabloda direkt düzenleme yapılabilir
    
    def calculate_percentage_for_row(self, row):
        """Satır için yüzde hesaplaması yapar"""
        try:
            siparis_no = row.get('Sipariş No')

            if pd.isna(siparis_no) or siparis_no == '' or siparis_no == 0:
                return ''

            # Aynı sipariş numarasına ait tüm satırları bul
            same_order_rows = self.filtered_df[self.filtered_df['Sipariş No'] == siparis_no]

            # Header değerini al (ilk satırdan) - integer kullanarak
            header_value = 0
            if not same_order_rows.empty:
                first_row = same_order_rows.iloc[0]
                header_raw = first_row.get('Header', '0')

                if pd.notna(header_raw) and str(header_raw).strip() != '':
                    try:
                        logging.info(f"Header raw: {header_raw}")
                        header_value = self.convert_turkish_to_integer(header_raw)
                        logging.info(f"Header parsed to integer: {header_raw} -> {header_value}")
                    except Exception as e:
                        logging.error(f"Header parse error: {e}")
                        header_value = 0

            # Genel toplamı hesapla (Birim Fiyat * Miktar + Vergi) - integer kullanarak
            total_amount = 0
            for _, order_row in same_order_rows.iterrows():
                try:
                    birim_fiyat = self.convert_turkish_to_integer(order_row.get('Birim Fiyat', 0))
                    miktar = pd.to_numeric(order_row.get('Miktar', 0), errors='coerce') or 0
                    vergi = self.convert_turkish_to_integer(order_row.get('Vergi', 0))

                    satir_toplam = int(round(birim_fiyat * miktar)) + vergi
                    total_amount += satir_toplam
                except:
                    continue

            logging.info(f"Total amount: {total_amount}, Header value: {header_value}")

            # Yüzde hesapla - Formül: 100 - (header_value / total_amount) * 100
            logging.info(f"Debug: total_amount={total_amount}, header_value={header_value}")

            if total_amount > 0 and header_value > 0:
                # Normal yüzde hesaplama: %100'den header oranını çıkar
                header_ratio = (header_value / total_amount) * 100
                percentage = 100 - header_ratio
                percentage_int = int(round(percentage))
                logging.info(f"Calculated percentage: {percentage_int}% (formula: 100 - ({header_value}/{total_amount}*100) = 100 - {header_ratio:.1f})")
                return f"%{percentage_int}"
            elif total_amount > 0:
                # Header 0 ise %100 olmalı
                logging.info("Header is 0, returning %100")
                return "%100"
            else:
                logging.info("Both values are 0, returning empty")
                return ''

        except Exception as e:
            logging.error(f"Percentage calculation error: {e}")
            return ''
    
    def convert_turkish_to_integer(self, value):
        """
        Türkçe sayı formatını integer'a çevirir.

        Türkçe format: 12.345,67 (nokta binlik ayırıcı, virgül ondalık ayırıcı)

        Args:
            value: String, float veya int değer

        Returns:
            int: Yuvarlanmış integer değer

        Örnekler:
            "12.345,67" -> 12346
            "12.345" -> 12345
            "12345,67" -> 12346
            12345.67 -> 12346
        """
        import re

        if not value or value in ['None', 'N/A', '', 0]:
            return 0

        # String'e çevir
        text_str = str(value).strip()

        # Para birimi sembollerini ve gereksiz karakterleri temizle
        text_str = text_str.replace('₺', '').replace('TL', '').strip()

        # Sayısal kısmı bul
        match = re.search(r'([\d.,]+)', text_str)

        if match:
            numeric_part = match.group(1)

            try:
                # Türkçe format kontrolü: virgül varsa ondalık ayırıcıdır
                if ',' in numeric_part:
                    # Noktaları (binlik ayırıcı) kaldır, virgülü noktaya çevir
                    numeric_part = numeric_part.replace('.', '').replace(',', '.')
                    return int(round(float(numeric_part)))
                else:
                    # Virgül yoksa, noktalar binlik ayırıcı olabilir
                    # Eğer birden fazla nokta varsa binlik ayırıcıdır, kaldır
                    if numeric_part.count('.') > 1:
                        numeric_part = numeric_part.replace('.', '')
                        return int(round(float(numeric_part)))
                    else:
                        # Tek nokta varsa, ondalık kısmının uzunluğuna bak
                        parts = numeric_part.split('.')
                        if len(parts) == 2:
                            # Ondalık kısım 3 haneden fazla ise binlik ayırıcıdır
                            if len(parts[1]) >= 3:
                                numeric_part = numeric_part.replace('.', '')
                                return int(round(float(numeric_part)))
                            else:
                                # 2 hane veya daha az ise ondalık ayırıcıdır
                                return int(round(float(numeric_part)))
                        else:
                            return int(round(float(numeric_part)))
            except:
                return 0

        return 0

    def extract_first_numeric_part(self, text):
        """Header metninden sadece ilk sayısal kısmı çıkarır"""
        import re

        if not text or text in ['None', 'N/A', '']:
            return '0'

        # String'e çevir
        text_str = str(text).strip()

        # İlk sayısal kısmı bul (nokta, virgül, rakam içeren)
        # Boşluk veya harf gelene kadar olan kısım
        match = re.match(r'^([\d.,]+)', text_str)

        if match:
            numeric_part = str(match.group(1))  # String'e çevir
            # Eğer virgül varsa Türkçe formatında işle
            if ',' in numeric_part:
                # Son virgül ondalik ayirıcısı ise
                parts = numeric_part.split(',')
                if len(parts) == 2 and len(parts[1]) <= 2:
                    # Binlik ayırıcı noktaları kaldır
                    integer_part = parts[0].replace('.', '')
                    return f"{integer_part},{parts[1]}"
                else:
                    # Sadece binlik ayırıcı
                    return numeric_part.replace(',', '')
            else:
                # Sadece nokta varsa binlik ayırıcı olarak işle
                if numeric_part.count('.') > 1:
                    return numeric_part.replace('.', '')
                else:
                    return numeric_part
        else:
            return '0'

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
    
    def clean_data(self):
        """Veri temizleme işlemleri"""
        # Tarih sütunundaki 00:00:00 ifadesini kaldır
        if 'Tarih' in self.original_df.columns:
            # Önce string olarak temizle (eğer 00:00:00 varsa)
            self.original_df['Tarih'] = self.original_df['Tarih'].astype(str)
            self.original_df['Tarih'] = self.original_df['Tarih'].str.replace(' 00:00:00', '', regex=False)
            # Sonra datetime'a çevir ve tekrar string formatına döndür
            self.original_df['Tarih'] = pd.to_datetime(self.original_df['Tarih'], errors='coerce')
            self.original_df['Tarih'] = self.original_df['Tarih'].dt.strftime('%Y-%m-%d')
            
        # Sozlesme ve Mağaza sütunlarındaki .0 ifadesini kaldır
        for col in ['Sozlesme', 'Mağaza']:
            if col in self.original_df.columns:
                self.original_df[col] = self.original_df[col].astype(str).str.replace('.0', '', regex=False)
                
        # Satir sütununa 1 ekle (0-based indexing'den 1-based'e çevir)
        if 'Satir' in self.original_df.columns:
            self.original_df['Satir'] = pd.to_numeric(self.original_df['Satir'], errors='coerce').fillna(0) + 1
                
        # Birim Fiyat, Vergi, Iskonto sütunlarını float'a çevir ve 2 ondalık basamakla formatla
        price_columns = ['Birim Fiyat', 'Vergi', 'Iskonto']
        for col in price_columns:
            if col in self.original_df.columns:
                self.original_df[col] = pd.to_numeric(self.original_df[col], errors='coerce').fillna(0)
                self.original_df[col] = self.original_df[col].round(2)

        # Header sütununu integer'a çevir
        if 'Header' in self.original_df.columns:
            # Her satır için convert_turkish_to_integer kullan
            self.original_df['Header'] = self.original_df['Header'].apply(
                lambda x: self.convert_turkish_to_integer(x) if pd.notna(x) else 0
            )

    def filter_by_year_month(self):
        """Yıl ve ay filtrelerine göre veriyi filtrele"""
        if self.original_df.empty:
            return
            
        try:
            filtered_df = self.original_df.copy()
            
            # Tarih sütununu datetime'a çevir
            if 'Tarih' in filtered_df.columns:
                filtered_df['Tarih'] = pd.to_datetime(filtered_df['Tarih'], errors='coerce')
                
                # Yıl filtresi
                selected_year = self.year_combo.currentText()
                if selected_year != "Tümü":
                    try:
                        year_num = int(selected_year)
                        year_filter = filtered_df['Tarih'].dt.year == year_num
                        filtered_df = filtered_df[year_filter]
                    except ValueError:
                        pass  # Geçersiz yıl değeri varsa filtre uygulanmaz
                
                # Ay filtresi
                selected_month = self.month_combo.currentText()
                if selected_month != "Tümü":
                    month_mapping = {
                        "Ocak": 1, "Şubat": 2, "Mart": 3, "Nisan": 4,
                        "Mayıs": 5, "Haziran": 6, "Temmuz": 7, "Ağustos": 8,
                        "Eylül": 9, "Ekim": 10, "Kasım": 11, "Aralık": 12
                    }
                    month_num = month_mapping.get(selected_month)
                    if month_num:
                        month_filter = filtered_df['Tarih'].dt.month == month_num
                        filtered_df = filtered_df[month_filter]
                
                # Tarih sütununu tekrar string formatına çevir
                filtered_df['Tarih'] = filtered_df['Tarih'].dt.strftime('%Y-%m-%d')

            self.filtered_df = filtered_df
            self.base_filtered_df = self.filtered_df.copy()  # Ana filtreyi sakla
            self.update_table()
            
        except Exception as e:
            logging.error(f"Filtreleme hatası: {str(e)}")

    def export_to_excel(self):
        try:
            save_path = r"D:/GoogleDrive"
            if not os.path.exists(save_path):
                os.makedirs(save_path)

            file_name = "~ Sozlesme_Listesi.xlsx"
            file_path = os.path.join(save_path, file_name)

            self.filtered_df.to_excel(file_path, index=False)
            QMessageBox.information(self, "Başarılı", f"Sözleşme listesi Excel'e kaydedildi: {file_path}")
        
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Excel'e kaydetme hatası: {str(e)}")

    def import_contract_data(self):
        """Search box'taki sözleşme numarasını sorgular"""
        try:
            contract_id = self.search_box.text().strip()
            
            if not contract_id:
                QMessageBox.warning(self, "Uyarı", "Lütfen sözleşme numarasını girin.")
                return
            
            # Sözleşme numarası validasyonu
            if len(contract_id) != 10 or not contract_id.startswith('15'):
                QMessageBox.warning(self, "Uyarı", "Lütfen doğru Sözleşme Numarası giriniz...\n\nSözleşme numarası 10 karakter olmalı ve '15' ile başlamalıdır.")
                return
            
            # Loading mesajı göster
            self.status_label.setText(f"🔍 Sözleşme {contract_id} sorgulanıyor...")
            self.import_btn.setEnabled(False)
            QApplication.processEvents()
            
            # Sozleme.py modülünü import et
            try:
                # Normal Python import kullan (PyInstaller uyumlu)
                try:
                    from PRG import Sozleme as sozleme_module
                except ImportError:
                    import Sozleme as sozleme_module
            except Exception as import_error:
                QMessageBox.warning(self, "Uyarı", f"Sozleme.py yüklenirken hata: {str(import_error)}")
                self.import_btn.setEnabled(True)
                self.status_label.setText("❌ Sozleme.py yüklenemedi")
                return

            # Sözleşme bilgilerini al
            contract_data = sozleme_module.get_all_contract_info(contract_id)

            if contract_data:
                self.status_label.setText(f"✅ Sözleşme {contract_id} başarıyla alındı")
                self.show_contract_details(contract_data, contract_id)
            else:
                self.status_label.setText(f"❌ Sözleşme {contract_id} bulunamadı")
                QMessageBox.warning(self, "Uyarı", f"Sözleşme {contract_id} bulunamadı veya hata oluştu.")
            
            self.import_btn.setEnabled(True)
            
        except Exception as e:
            logging.error(f"Sözleşme sorgulama hatası: {str(e)}")
            self.status_label.setText(f"❌ Sözleşme sorgulama hatası: {str(e)}")
            QMessageBox.critical(self, "Hata", f"Sözleşme sorgulama hatası: {str(e)}")
            self.import_btn.setEnabled(True)

    def handle_cell_edit(self, item):
        """Hücre düzenleme işlemini ele alır"""
        try:
            col_name = self.filtered_df.columns[item.column()]

            if col_name == 'Sepet':
                # Sepet sütunu düzenlenebilir, değişikliği kaydet
                try:
                    new_value = float(item.text()) if item.text() else 0
                    original_index = self.filtered_df.index[item.row()]
                    self.original_df.at[original_index, 'Sepet'] = new_value
                    self.filtered_df.at[original_index, 'Sepet'] = new_value

                    # Toplam bilgileri güncelle
                    self.update_summary_labels()
                except ValueError:
                    # Geçersiz değer girildiğinde eski değeri geri yükle
                    original_value = self.filtered_df.iat[item.row(), item.column()]
                    item.setText(str(original_value))
            else:
                # Diğer sütunlar için SPEC kontrolü yap
                row_idx = item.row()

                # Sipariş No ve Sip Kalem No sütunlarının indekslerini bul
                siparis_no_col = None
                sip_kalem_no_col = None
                spec_col = None

                for col_idx, col_name_check in enumerate(self.filtered_df.columns):
                    if 'Sipariş No' in col_name_check or col_name_check == 'Siparis_No':
                        siparis_no_col = col_idx
                    elif 'Sip Kalem No' in col_name_check or col_name_check == 'Sip_Kalem_No':
                        sip_kalem_no_col = col_idx
                    elif col_name_check == 'SPEC':
                        spec_col = col_idx

                # SPEC sütununun değerini kontrol et
                if spec_col is not None:
                    spec_item = self.table.item(row_idx, spec_col)
                    spec_value = spec_item.text() if spec_item else ""

                    # Sipariş No ve Sip Kalem No değerlerini kontrol et
                    siparis_no_value = ""
                    sip_kalem_no_value = ""

                    if siparis_no_col is not None:
                        siparis_no_item = self.table.item(row_idx, siparis_no_col)
                        siparis_no_value = siparis_no_item.text() if siparis_no_item else ""

                    if sip_kalem_no_col is not None:
                        sip_kalem_no_item = self.table.item(row_idx, sip_kalem_no_col)
                        sip_kalem_no_value = sip_kalem_no_item.text() if sip_kalem_no_item else ""

                    # Dialog açılmasını devre dışı bırak - tabloda direkt düzenleme yapılabilir
                    # if (('none' in spec_value.strip().lower() or spec_value.strip() == '' or 'n/a' in spec_value.strip().lower()) and
                    #     siparis_no_value.strip() != '' and
                    #     sip_kalem_no_value.strip() != ''):
                    pass

        except Exception as e:
            logging.error(f"Hücre düzenleme hatası: {str(e)}")

    def show_contract_details(self, contract_data, contract_id):
        """Sözleşme detaylarını yeni pencerede göster"""
        try:
            # Yeni pencere oluştur
            contract_window = ContractDetailsWindow(contract_data, contract_id, self)
            contract_window.show()
            
        except Exception as e:
            logging.error(f"Sözleşme detay penceresi hatası: {str(e)}")
            QMessageBox.critical(self, "Hata", f"Sözleşme detay penceresi hatası: {str(e)}")

    def load_initial_data(self):
        """Başlangıçta mevcut verileri yükle ve bu ay filtrele"""
        try:
            self.progress_bar.setVisible(True)
            self.status_label.setVisible(True)
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(0)
            self.status_label.setText("📊 Mevcut veriler yükleniyor...")
            QApplication.processEvents()

            # PRGsheet/Ayar sayfasından SPREADSHEET_ID'yi yükle
            spreadsheet_id = CentralConfigManager().MASTER_SPREADSHEET_ID

            # Google Sheets URL'sini oluştur
            gsheets_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"

            self.progress_bar.setValue(10)
            self.status_label.setText("🔗 Google Sheets'e bağlanıyor...")
            QApplication.processEvents()

            # URL'den Excel dosyasını oku
            response = requests.get(gsheets_url, timeout=30)

            self.progress_bar.setValue(30)
            self.status_label.setText("✅ Google Sheets'e bağlantı başarılı")
            QApplication.processEvents()

            if response.status_code == 401:
                self.progress_bar.setVisible(False)
                self.status_label.setText("❌ Google Sheets erişim hatası: Dosya özel veya izin gerekli")
                logging.error("Google Sheets erişim hatası: Dosya özel veya izin gerekli")
                return
            elif response.status_code != 200:
                self.progress_bar.setVisible(False)
                self.status_label.setText(f"❌ HTTP Hatası: {response.status_code} - {response.reason}")
                logging.error(f"HTTP Hatası: {response.status_code} - {response.reason}")
                return

            response.raise_for_status()

            # Siparis sayfasını oku
            self.progress_bar.setValue(50)
            self.status_label.setText("📋 Sipariş verileri yükleniyor...")
            QApplication.processEvents()

            self.original_df = pd.read_excel(BytesIO(response.content), sheet_name="Siparis")
            
            # Veri temizleme işlemleri
            self.progress_bar.setValue(70)
            self.status_label.setText("🔄 Veriler temizleniyor...")
            QApplication.processEvents()
            self.clean_data()

            # Bu ayın verilerini filtrele
            if 'Tarih' in self.original_df.columns:
                self.progress_bar.setValue(85)
                self.status_label.setText(f"📅 {self.month_names_tr.get(self.current_month, self.current_month)} ayı filtreleniyor...")
                QApplication.processEvents()

                # Tarih sütununu datetime'a çevir
                self.original_df['Tarih'] = pd.to_datetime(self.original_df['Tarih'], errors='coerce')

                # Bu ayın verilerini filtrele
                month_filter = self.original_df['Tarih'].dt.strftime('%B') == self.current_month
                self.original_df = self.original_df[month_filter]

                # Tarih sütununu tekrar string formatına döndür
                self.original_df['Tarih'] = self.original_df['Tarih'].dt.strftime('%Y-%m-%d')

            self.progress_bar.setValue(95)
            self.status_label.setText("🔄 Tablo güncelleniyor...")
            QApplication.processEvents()

            self.filtered_df = self.original_df.copy()
            self.base_filtered_df = self.filtered_df.copy()  # Ana filtreyi sakla
            self.update_table()

            # Tüm işlemler tamamlandı
            self.progress_bar.setValue(100)
            QApplication.processEvents()

            # Progress bar'ı 1 saniye sonra gizle
            QTimer.singleShot(1000, lambda: self.progress_bar.setVisible(False))

            self.status_label.setText(f"✅ {len(self.original_df)} {self.month_names_tr.get(self.current_month, self.current_month)} siparişi yüklendi")

        except Exception as e:
            logging.error(f"Hata: {str(e)}")
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"❌ Veri yükleme hatası: {str(e)}")
            QMessageBox.critical(self, "Hata", f"Veri yükleme hatası: {str(e)}")

    def filter_column_header(self, index):
        """Sütun başlığına tıklandığında filtreleme menüsü göster"""

        if self.filtered_df.empty:
            return

        column_name = self.filtered_df.columns[index]

        # Filtrelenebilir sütun isimleri (yüklenen tablodaki gerçek sütunlara göre)
        filterable_columns = ["Tarih", "Cari Adi", "Personel", "Mağaza"]
        # Sadece tabloda var olan sütunları kullan
        filterable_columns = [col for col in filterable_columns if col in self.filtered_df.columns]

        if column_name in filterable_columns:
            menu = QMenu(self)
            # Ana filtre (ay/yıl) sonrası verilerden unique değerleri al
            source_df = self.base_filtered_df if not self.base_filtered_df.empty else self.filtered_df
            unique_values = source_df[column_name].dropna().unique()

            if len(unique_values) == 0:
                QMessageBox.warning(self, "Uyarı", "Filtreleme için uygun değer bulunamadı.")
                return

            for value in unique_values:
                action = QAction(str(value), self)
                action.triggered.connect(lambda checked, col=column_name, val=value: self.filter_by_column(col, val))
                menu.addAction(action)

            # Tümünü göster seçeneği ekle
            menu.addSeparator()
            show_all_action = QAction("Tümünü Göster", self)
            show_all_action.triggered.connect(lambda: self.reset_filter())
            menu.addAction(show_all_action)

            # Menüyü tıklanan header sütununun altında göster
            try:
                header = self.table.horizontalHeader()
                # Tıklanan sütunun pozisyonunu al
                section_pos = header.sectionPosition(index)
                section_width = header.sectionSize(index)

                # Header'ın global pozisyonunu al
                header_global_pos = header.mapToGlobal(QPoint(section_pos, header.height()))

                menu.exec_(header_global_pos)

            except Exception as e:
                # Son çare: Mouse pozisyonu
                menu.exec_(QCursor.pos())
        else:
            pass

    def filter_by_column(self, column_name, value):
        """Seçilen sütun ve değere göre filtrele - mevcut filtrelenmiş tablo üzerinden"""
        try:
            # Mevcut filtrelenmiş tablo üzerinden tekrar filtrele
            filtered_df = self.filtered_df[self.filtered_df[column_name] == value]

            if filtered_df.empty:
                QMessageBox.warning(self, "Uyarı", "Filtreleme sonucu hiç veri bulunamadı.")
                return

            self.filtered_df = filtered_df.reset_index(drop=True)
            self.update_table()

            # Status güncellemesi
            self.status_label.setText(f"🔍 {column_name}: {value} - {len(self.filtered_df)} sonuç")

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Filtreleme sırasında hata oluştu: {str(e)}")

    def reset_filter(self):
        """Filtreyi sıfırla ve ana filtreye (ay/yıl) dön"""
        try:
            # Ana filtre varsa ona dön, yoksa orijinal verilere dön
            if not self.base_filtered_df.empty:
                self.filtered_df = self.base_filtered_df.copy()
            else:
                self.filtered_df = self.original_df.copy()
            self.update_table()
            self.status_label.setText(f"✅ Filtre sıfırlandı - {len(self.filtered_df)} sonuç")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Filtre sıfırlama hatası: {str(e)}")

    def control_excel_data(self):
        """Kontrol.xlsx dosyasından veri yükler ve sözleşme numaralarını karşılaştırır"""
        try:
            # Kontrol.xlsx dosyası yolu
            kontrol_file_path = r"D:\GoogleDrive\PRG\Kontrol.xlsx"

            # Dosya varlık kontrolü
            if not os.path.exists(kontrol_file_path):
                QMessageBox.warning(self, "Hata", f"Kontrol.xlsx dosyası bulunamadı:\n{kontrol_file_path}")
                return

            # Kontrol.xlsx dosyasını okuma
            self.status_label.setText("📂 Kontrol.xlsx dosyası okunuyor...")
            QApplication.processEvents()

            df_kontrol = pd.read_excel(kontrol_file_path)

            if df_kontrol.empty:
                QMessageBox.warning(self, "Hata", "Kontrol.xlsx dosyası boş!")
                return

            # Önce "Kullanıcı durumu" "İptal Edildi" olan verileri sil
            if 'Kullanıcı durumu' in df_kontrol.columns:
                original_count = len(df_kontrol)
                df_kontrol = df_kontrol[df_kontrol['Kullanıcı durumu'] != 'İptal Edildi']
                self.status_label.setText(f"🚫 İptal edildi durumu filtrelendi: {original_count} -> {len(df_kontrol)} kayıt")
                QApplication.processEvents()
            else:
                QMessageBox.warning(self, "Uyarı", "'Kullanıcı durumu' sütunu bulunamadı!")

            # 37 gün öncesi verileri filtrele
            if 'Sipariş tarihi' in df_kontrol.columns:
                cutoff_date = datetime.now() - timedelta(days=37)
                df_kontrol['Sipariş tarihi'] = pd.to_datetime(df_kontrol['Sipariş tarihi'], errors='coerce')
                date_filter_count = len(df_kontrol)
                df_kontrol = df_kontrol[df_kontrol['Sipariş tarihi'] >= cutoff_date]
                self.status_label.setText(f"📅 Son 37 gün verileri filtrelendi: {date_filter_count} -> {len(df_kontrol)} kayıt")
                QApplication.processEvents()
            else:
                QMessageBox.warning(self, "Uyarı", "'Sipariş tarihi' sütunu bulunamadı!")

            self.status_label.setText("🔍 Sözleşme numaraları karşılaştırılıyor...")
            QApplication.processEvents()

            # Mevcut dataframe'deki sözleşme numaralarını al
            if 'Sozlesme' not in self.original_df.columns or 'Belge no' not in df_kontrol.columns:
                QMessageBox.warning(self, "Uyarı", "Gerekli sütunlar bulunamadı. 'Sozlesme' (tablo) ve 'Belge no' (Kontrol.xlsx) sütunları olmalı.")
                return

            # Kontrol.xlsx'deki Belge no'ları string formatında al
            kontrol_belge_nos = set(df_kontrol['Belge no'].astype(str).str.strip())

            # Mevcut dataframe'deki Sozlesme'leri string formatında al
            existing_sozlesmes = set(self.original_df['Sozlesme'].astype(str).str.strip())

            # Eşleşmeyen Belge no'ları bul (Kontrol.xlsx'de var ama mevcut dataframe'de yok)
            missing_sozlesmes = kontrol_belge_nos - existing_sozlesmes

            # Eşleşmeyen kayıtları filtrele
            if missing_sozlesmes:
                df_filtered = df_kontrol[df_kontrol['Belge no'].astype(str).str.strip().isin(missing_sozlesmes)]
            else:
                df_filtered = pd.DataFrame()  # Boş dataframe

            # Sonuç bilgilerini hazırla
            total_kontrol_records = len(df_kontrol)
            remaining_records = len(df_filtered)

            if remaining_records == 0:
                message = f"""Kontrol tamamlandı!

Toplam Kontrol.xlsx kaydı: {total_kontrol_records}
Eşleşmeyen kayıt sayısı: 0

Tüm veriler tabloda mevcut."""
                QMessageBox.information(self, "Sonuç", message)
                self.status_label.setText("✅ Kontrol tamamlandı - Tüm veriler eşleşiyor")
            else:
                # Eşleşmeyen kayıtları Kontrol.xlsx dosyasına kaydet
                df_filtered.to_excel(kontrol_file_path, index=False, engine='openpyxl')

                message = f"""Kontrol tamamlandı!

Toplam Kontrol.xlsx kaydı: {total_kontrol_records}
Eşleşmeyen kayıt sayısı: {remaining_records}

Eşleşmeyen veriler Kontrol.xlsx dosyasında güncellendi."""

                QMessageBox.information(self, "Sonuç", message)
                self.status_label.setText(f"✅ Kontrol tamamlandı - {remaining_records} eşleşmeyen kayıt güncellendi")

        except Exception as e:
            error_msg = f"Kontrol işlemi hatası: {str(e)}"
            logging.error(error_msg)
            QMessageBox.critical(self, "Hata", error_msg)
            self.status_label.setText("❌ Kontrol işlemi hatası")

    def handle_ctrl_c(self):
        """Ctrl+C basıldığında seçili hücreyi kopyalar"""
        selected_items = self.table.selectedItems()
        if selected_items:
            # Sadece ilk seçili hücreyi kopyala
            text = selected_items[0].text()
            QApplication.clipboard().setText(text)
            
            # Status bar güncelle
            if hasattr(self, 'status_label'):
                old_text = self.status_label.text()
                self.status_label.setText("📋 Kopyalandı")
                # 3 saniye sonra status'u eski haline getir
                QTimer.singleShot(3000, lambda: self.status_label.setText(old_text))