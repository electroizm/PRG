"""
Sanal Pos Modülü
"""

import os
import sys
import pandas as pd
import requests
from io import BytesIO
from datetime import datetime, date
from pathlib import Path

# Üst dizini Python path'e ekle (central_config için)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Central config import
from central_config import CentralConfigManager

from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, 
                             QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView, 
                             QMenu, QProgressBar, QLabel, QApplication)
from PyQt5.QtGui import QFont, QColor


class SanalPosApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sanal Pos Modülü - PRG v2.0")
        self.setMinimumSize(1200, 800)
        self.mikro_calisiyor = False

        # Load environment variables - Service Account
        config_manager = CentralConfigManager()
        self.spreadsheet_id = config_manager.MASTER_SPREADSHEET_ID

        # Define sheet names for data sources
        self.sanalpos_sheet_name = "SanalPos"
        self.kasa_sheet_name = "Kasa"
        self.sanal_pos_sheet_name = "SanalPos"

        # Apply main window styling - Light theme
        self.setStyleSheet("""
            QMainWindow {
                background-color: #ffffff;
                color: #000000;
            }
        """)

        self.setup_ui()

        # Lazy loading için flag
        self._data_loaded = False

    def showEvent(self, event):
        """Widget ilk gösterildiğinde veri yükle (lazy loading)"""
        super().showEvent(event)
        if not self._data_loaded:
            self._data_loaded = True
            # UI render olduktan sonra veri yükle
            QTimer.singleShot(100, lambda: self.load_sanalpos_data(force_reload=False))

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
        
        self.layout = QVBoxLayout(self.central_widget)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(0)

        # Progress Bar (Risk modülü gibi) - Status bar'dan önce tanımlanmalı
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

        # Header layout with improved spacing
        header_layout = QHBoxLayout()
        header_layout.setSpacing(15)
        header_layout.setContentsMargins(5, 5, 5, 5)

        # Mikro button - Light theme
        self.micro_button = QPushButton("Mikro")
        self.micro_button.setStyleSheet("""
            QPushButton {
                background-color: #dfdfdf;
                color: black;
                border: 1px solid #444;
                padding: 8px 16px;
                border-radius: 5px;
                font-weight: bold;
                font-size: 14px;
                font-family: 'Segoe UI', Arial, sans-serif;
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
        
        # Refresh button - Light theme
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
                font-family: 'Segoe UI', Arial, sans-serif;
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

        header_layout.addWidget(self.micro_button)
        header_layout.addWidget(self.refresh_button)
        header_layout.addStretch()
        
        # Header layout'u widget olarak sar - beyaz arka plan için
        header_widget = QWidget()
        header_widget.setLayout(header_layout)
        header_widget.setStyleSheet("""
            background-color: #ffffff;
            margin-bottom: 0px;
        """)
        header_layout.setContentsMargins(10, 10, 10, 10)
        
        self.layout.addWidget(header_widget)
        
        # Direct Table Widget (no tabs)
        self.table = QTableWidget()
        self.layout.addWidget(self.table)

        # Status Layout (Label + Progress Bar) - risk_module.py gibi
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

        self.layout.addWidget(status_widget)

        # Connect buttons
        self.micro_button.clicked.connect(self.run_mikro)
        # Verileri Yenile butonu: cache'i bypass et, Google Sheets'ten çek
        self.refresh_button.clicked.connect(lambda: self.load_sanalpos_data(force_reload=True))

    def get_google_sheets_url(self, sheet_name, format_type="csv"):
        """Generate Google Sheets export URL for specific sheet"""
        if not self.spreadsheet_id:
            return None
        
        if format_type == "csv":
            # For CSV export, we need to find the sheet GID (this is simplified)
            # In practice, you might need to map sheet names to GIDs
            return f"https://docs.google.com/spreadsheets/d/{self.spreadsheet_id}/export?format=csv&gid=0"
        elif format_type == "xlsx":
            return f"https://docs.google.com/spreadsheets/d/{self.spreadsheet_id}/export?format=xlsx"
        
        return None

    def load_data_from_sheets(self, sheet_name):
        """Load data from Google Sheets"""
        try:
            if not self.spreadsheet_id:
                self.status_label.setText("❌ SPREADSHEET_ID bulunamadı (PRGsheet/Ayar sayfasını kontrol edin)")
                return pd.DataFrame()
            
            # Use Excel format to get all sheets
            export_url = f"https://docs.google.com/spreadsheets/d/{self.spreadsheet_id}/export?format=xlsx"
            
            response = requests.get(export_url, timeout=30)
            response.raise_for_status()
            
            # Read specific sheet
            df = pd.read_excel(BytesIO(response.content), sheet_name=sheet_name)
            return df
            
        except requests.exceptions.RequestException as e:
            self.status_label.setText(f"❌ Google Sheets bağlantı hatası: {str(e)}")
            return pd.DataFrame()
        except Exception as e:
            self.status_label.setText(f"❌ Veri yükleme hatası ({sheet_name}): {str(e)}")
            return pd.DataFrame()

    def run_mikro(self):
        program_path = r"D:/GoogleDrive/PRG/EXE/SanalPos.exe"
        if not os.path.exists(program_path):
            self.status_label.setText(f"Program bulunamadı: {program_path}")
            return

        try:
            self.mikro_calisiyor = True
            self.micro_button.setEnabled(False)
            self.refresh_button.setEnabled(False)
            self.status_label.setText("🔄 SanalPos.exe çalıştırılıyor...")

            os.startfile(program_path)
            
            # Risk modülü gibi 7 saniye + 5 saniye bekleme
            QTimer.singleShot(7000, self.on_mikro_finished)

        except Exception as e:
            self.status_label.setText(f"❌ Program çalıştırma hatası: {str(e)}")
            self._reset_mikro_state_if_needed()

    def _reload_data_after_update(self):
        self.mikro_calisiyor = False
        self.micro_button.setEnabled(True)
        self.micro_button.setText("Mikro")

        self.load_sanalpos_data()
        
        self.status_label.setText("Tablo başarıyla güncellendi.")

    def on_mikro_finished(self):
        """Mikro program bittikten sonra (Risk modülü gibi)"""
        self.status_label.setText("✅ SanalPos.exe tamamlandı, Google Sheets güncelleme bekleniyor...")
        
        # Google Sheets'e kaydedilmesi için ek bekleme (5 saniye)
        QTimer.singleShot(5000, self.delayed_data_refresh)
    
    def delayed_data_refresh(self):
        """Gecikmeli veri yenileme (Risk modülü gibi)"""
        self.status_label.setText("🔄 Google Sheets'ten güncel veriler alınıyor...")
        from PyQt5.QtWidgets import QApplication
        QApplication.processEvents()
        self.load_sanalpos_data()
    
    def _reset_mikro_state_if_needed(self):
        if self.mikro_calisiyor:
            self.progress_bar.setVisible(False)
            self.mikro_calisiyor = False
            self.micro_button.setEnabled(True)
            self.refresh_button.setEnabled(True)
            self.micro_button.setText("Mikro")
            self.status_label.setText("Zaman aşımı: Program tamamlandı. Veriler yeniden yükleniyor...")
            # Auto reload data after mikro update
            QTimer.singleShot(2000, self.load_sanalpos_data)
    
    def load_sanalpos_data(self, force_reload=False):
        """
        SanalPos verilerini yükle (cache-aware)

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
            if not force_reload and cache and cache.has("SanalPos"):
                sanal_pos_df = cache.get("SanalPos")

                # Tarih sütunlarını işle
                if "Belge tarihi" in sanal_pos_df.columns:
                    sanal_pos_df["Belge tarihi"] = pd.to_datetime(sanal_pos_df["Belge tarihi"], errors='coerce')
                    sanal_pos_df["Belge tarihi"] = sanal_pos_df["Belge tarihi"].apply(
                        lambda x: x.date() if pd.notna(x) else None
                    )
                if "Tarih" in sanal_pos_df.columns:
                    sanal_pos_df["Tarih"] = pd.to_datetime(sanal_pos_df["Tarih"], errors='coerce')
                    sanal_pos_df["Tarih"] = sanal_pos_df["Tarih"].apply(
                        lambda x: x.date() if pd.notna(x) else None
                    )

                self.show_table(sanal_pos_df, "Birleşmiş Veriler")
                self.status_label.setText(f"✅ SanalPos verileri yüklendi (Cache'den - anında) - {len(sanal_pos_df)} kayıt")
                return

            # Cache yoksa veya force_reload ise: Google Sheets'ten çek
            # Progress bar'ı göster (Risk modülü gibi)
            if not self.mikro_calisiyor:
                self.progress_bar.setVisible(True)
                self.progress_bar.setRange(0, 100)  # Yuzde bazli
                self.progress_bar.setValue(0)  # 0%
                self.micro_button.setEnabled(False)
                self.refresh_button.setEnabled(False)

            self.status_label.setText("📊 Google Sheets'ten SanalPos verileri yükleniyor...")
            from PyQt5.QtWidgets import QApplication
            QApplication.processEvents()

            # Progress: 10%
            self.progress_bar.setValue(10)
            QApplication.processEvents()

            # Load processed data from SanalPos sheet
            sanal_pos_df = self.load_data_from_sheets(self.sanal_pos_sheet_name)

            # Progress: 30%
            self.progress_bar.setValue(30)
            QApplication.processEvents()

            if sanal_pos_df.empty:
                self.status_label.setText("❌ SanalPos verileri yüklenemedi")
                if hasattr(self, "table"): self.table.clearContents(); self.table.setRowCount(0)
                return

            # Progress: 50% - Tarih sutunlari isleniyor
            self.progress_bar.setValue(50)
            QApplication.processEvents()

            # Process date columns if they exist
            if "Belge tarihi" in sanal_pos_df.columns:
                sanal_pos_df["Belge tarihi"] = pd.to_datetime(sanal_pos_df["Belge tarihi"], errors='coerce')
                # NaT değerlerini temizle
                sanal_pos_df["Belge tarihi"] = sanal_pos_df["Belge tarihi"].apply(
                    lambda x: x.date() if pd.notna(x) else None
                )
            if "Tarih" in sanal_pos_df.columns:
                sanal_pos_df["Tarih"] = pd.to_datetime(sanal_pos_df["Tarih"], errors='coerce')
                # NaT değerlerini temizle
                sanal_pos_df["Tarih"] = sanal_pos_df["Tarih"].apply(
                    lambda x: x.date() if pd.notna(x) else None
                )

            # Progress: 70% - Tablo gosteriliyor
            self.progress_bar.setValue(70)
            QApplication.processEvents()

            # Show the processed table
            self.show_table(sanal_pos_df, "Birleşmiş Veriler")

            # Progress: 90% - Cache'e kaydediliyor
            self.progress_bar.setValue(90)
            QApplication.processEvents()

            # Cache'e kaydet
            if cache:
                cache.set("SanalPos", sanal_pos_df)

            # Progress: 100% - Tamamlandi
            self.progress_bar.setValue(100)
            QApplication.processEvents()

            self.status_label.setText(f"✅ SanalPos verileri başarıyla yüklendi - {len(sanal_pos_df)} kayıt")

            # Progress bar'i 1 saniye sonra gizle
            QTimer.singleShot(1000, lambda: self.progress_bar.setVisible(False))

        except Exception as e:
            self.status_label.setText(f"❌ Veri yüklenirken bir hata oluştu: {e}")
            if hasattr(self, "table"): self.table.clearContents(); self.table.setRowCount(0)
            self.progress_bar.setVisible(False)  # Hata durumunda hemen gizle

        finally:
            # Butonları aktif et
            if not self.mikro_calisiyor:
                self.micro_button.setEnabled(True)
                self.refresh_button.setEnabled(True)
            else:
                # Mikro çalışıyorsa sadece state'i sıfırla
                self.mikro_calisiyor = False
                self.micro_button.setEnabled(True)
                self.refresh_button.setEnabled(True)
                self.micro_button.setText("Mikro")
        
    def show_table(self, dataframe, title):
        # Use existing table widget
        table = self.table
        table.clearContents()
        table.setRowCount(0)
        
        # Configure table
        table.setRowCount(dataframe.shape[0])
        table.setColumnCount(dataframe.shape[1])
        table.setHorizontalHeaderLabels(dataframe.columns)

        # Apply table styling - Light theme (risk_module.py gibi)
        table.setStyleSheet("""
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
        """)

        # Set table properties for better appearance
        table.setAlternatingRowColors(True)
        table.setShowGrid(True)
        table.setSortingEnabled(False)
        table.setSelectionBehavior(QAbstractItemView.SelectItems)  # Select individual cells, not rows
        table.setSelectionMode(QAbstractItemView.SingleSelection)  # Single cell selection
        table.setFocusPolicy(Qt.NoFocus)  # Remove focus policy to eliminate dotted borders

        # Fill table with data and apply enhanced formatting
        for i in range(dataframe.shape[0]):
            for j in range(dataframe.shape[1]):
                value = dataframe.iat[i, j]
                
                # Format date columns
                if isinstance(value, pd.Timestamp):
                    if pd.notna(value):
                        display_value = value.strftime('%Y-%m-%d')
                    else:
                        display_value = ""
                elif isinstance(value, date):
                    display_value = value.strftime('%Y-%m-%d')
                elif pd.isna(value) or value is None or str(value).lower() in ['nan', 'nat']:
                    display_value = ""
                else:
                    # Check if it's a float that's actually an integer (like 521041.0)
                    if isinstance(value, float) and value.is_integer():
                        display_value = str(int(value))
                    else:
                        display_value = str(value)
                    
                item = QTableWidgetItem(display_value)
                item.setFlags(item.flags() ^ Qt.ItemIsEditable)  # Make non-editable
                
                # Set font properties for better readability
                font = QFont('Segoe UI', 12)
                font.setBold(True)
                item.setFont(font)
                
                # Color coding for specific columns
                column_name = dataframe.columns[j]
                if 'tutar' in column_name.lower() or 'miktar' in column_name.lower():
                    try:
                        numeric_value = float(str(value).replace(',', ''))
                        if numeric_value > 0:
                            item.setForeground(QColor("#4CAF50"))  # Green for positive
                        elif numeric_value < 0:
                            item.setForeground(QColor("#f44336"))  # Red for negative
                        else:
                            item.setForeground(QColor("#ffffff"))  # White for zero
                    except:
                        item.setForeground(QColor("#ffffff"))
                else:
                    item.setForeground(QColor("#ffffff"))
                
                table.setItem(i, j, item)

        # Enhanced header styling
        header = table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setStretchLastSection(False)
        
        # Set minimum column widths
        for i in range(table.columnCount()):
            table.setColumnWidth(i, max(150, table.columnWidth(i)))

        # Resize columns to content but with minimum width
        table.resizeColumnsToContents()
        
        # Set row height for better readability
        for i in range(table.rowCount()):
            table.setRowHeight(i, 35)

        # Add context menu
        table.setContextMenuPolicy(Qt.CustomContextMenu)
        table.customContextMenuRequested.connect(self.show_context_menu)
        
        # Table is already added to layout in setup_ui

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