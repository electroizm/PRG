
import sys
import os
import pandas as pd
import requests
from io import BytesIO

# Üst dizini Python path'e ekle (central_config için)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Central config import
from central_config import CentralConfigManager

from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QStatusBar, QInputDialog, QMessageBox, QLineEdit, QLabel)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt
import time

# Embedded Resources
from embedded_resources import get_app_icon


# ============================================================================
# GLOBAL VERİ CACHE SİSTEMİ
# ============================================================================
class GlobalDataCache:
    """
    Tüm modüller için paylaşımlı veri cache'i
    - İlk açılışta: Google Sheets'ten çek + cache'e kaydet
    - Sonraki açılışlar: Cache'den anında yükle
    - Manuel refresh: Cache'i temizle + yeniden çek
    """
    _instance = None
    _cache = {}  # {sheet_name: {"data": df, "timestamp": time}}
    CACHE_DURATION = 300  # 5 dakika (opsiyonel - şimdilik sınırsız)

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(GlobalDataCache, cls).__new__(cls)
        return cls._instance

    def get(self, sheet_name):
        """Cache'den veri al"""
        if sheet_name in self._cache:
            return self._cache[sheet_name]["data"].copy()
        return None

    def set(self, sheet_name, data):
        """Cache'e veri kaydet"""
        self._cache[sheet_name] = {
            "data": data.copy(),
            "timestamp": time.time()
        }

    def clear(self, sheet_name=None):
        """Cache'i temizle (tümü veya belirli sayfa)"""
        if sheet_name:
            self._cache.pop(sheet_name, None)
        else:
            self._cache.clear()

    def has(self, sheet_name):
        """Cache'de var mı kontrol et"""
        return sheet_name in self._cache

# Core Architecture
from core_architecture import (
    EventBus, AppState, ThemeManager, CommandInvoker, ModuleRegistry,
    ModuleConfig, ModuleType, NavigateCommand
)

# UI Components
from ui_components import AdvancedNavigationBar, AdvancedPageManager

# Modules
from stok_module import StokApp
from sozlesme_module import SozlesmeApp
from ssh_module import SshModule
from okc_module import OKCYazarKasaApp
from risk_module import RiskApp
from kasa_module import KasaApp
from virman_module import VirmanModule
from sanalpos_module import SanalPosApp
from irsaliye_module import IrsaliyeWindow
from sevkiyat_module import SevkiyatModule
from fiyat_module import FiyatModule
from ayar_module import AyarlarApp

class ModernMainApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self._initialize_core_systems()
        self._register_modules()
        self._setup_window()
        self._setup_ui()
        self._connect_events()
        self._show_initial_module()
        self.ayarlar_password = None  # Cache for password
        self.virman_password = None  # Cache for virman password
    
    def _initialize_core_systems(self):
        self.event_bus = EventBus()
        self.app_state = AppState(self.event_bus)
        self.theme_manager = ThemeManager(self.event_bus)
        self.command_invoker = CommandInvoker()
        self.module_registry = ModuleRegistry(self.event_bus)

        # TÜM ayarları Google Sheets'ten yükle ve cache'e al (startup optimization)
        self._load_all_settings()

    def _load_all_settings(self):
        """PRGsheet/Ayar sayfasından TÜM ayarları yükle ve cache'e al (startup'ta)

        Bu method:
        - TÜM Ayar sayfasını cache'e alır (Global + App-specific)
        - SQL ayarlarını environment variable'lara set eder
        - Diğer modüller için ayarları hazır hale getirir
        """
        try:
            config_manager = CentralConfigManager()
            settings = config_manager.get_settings()  # Şifreli cache kullanır - TÜM ayarları yükler

            # SQL ayarlarını environment variable olarak set et (backward compatibility)
            os.environ['SQL_SERVER'] = settings.get('SQL_SERVER', '')
            os.environ['SQL_DATABASE'] = settings.get('SQL_DATABASE', '')
            os.environ['SQL_USERNAME'] = settings.get('SQL_USERNAME', '')
            os.environ['SQL_PASSWORD'] = settings.get('SQL_PASSWORD', '')

            print(f"[OK] {len(settings)} settings loaded and cached from PRGsheet/Ayar")
            print(f"   SQL Server: {settings.get('SQL_SERVER', 'N/A')}")
        except Exception as e:
            print(f"[WARNING] Could not load settings from PRGsheet/Ayar: {e}")
            # Hata olsa bile program çalışmaya devam etsin
    
    def _register_modules(self):
        modules_config = [
            ModuleConfig("stok", "Stok", ModuleType.STOK, StokApp),
            ModuleConfig("sevkiyat", "Sevkiyat", ModuleType.SEVKIYAT, SevkiyatModule),
            ModuleConfig("sozlesme", "Sözleşme", ModuleType.SOZLESME, SozlesmeApp),
            ModuleConfig("okc", "ÖKC YazarKasa", ModuleType.OKC, OKCYazarKasaApp),
            ModuleConfig("risk", "Risk", ModuleType.RISK, RiskApp),
            ModuleConfig("ssh", "SSH", ModuleType.SSH, SshModule),
            ModuleConfig("kasa", "Kasa", ModuleType.KASA, KasaApp),
            ModuleConfig("virman", "Virman", ModuleType.VIRMAN, VirmanModule),
            ModuleConfig("sanalpos", "Sanal Pos", ModuleType.SANALPOS, SanalPosApp),
            ModuleConfig("irsaliye", "İrsaliye", ModuleType.IRSALIYE, IrsaliyeWindow),
            ModuleConfig("fiyat", "Fiyat", ModuleType.FIYAT, FiyatModule),
            ModuleConfig("ayarlar", "Ayarlar", ModuleType.AYARLAR, AyarlarApp),
        ]

        for config in modules_config:
            self.module_registry.register_module(config)
    
    def _setup_window(self):
        self.setWindowTitle("PRG")
        # Use embedded icon (works everywhere including taskbar)
        self.setWindowIcon(get_app_icon())
        self.setMinimumSize(1200, 600)  # Minimum window boyutu (geometry hatası için)
        self.resize(1280, 800)

    def _get_ayarlar_password_from_sheets(self):
        """PRGsheet dosyasının Pass sayfasından AyarlarModul şifresini oku"""
        try:
            # Service Account ile PRGsheet'e erişim
            config_manager = CentralConfigManager()
            spreadsheet_id = config_manager.MASTER_SPREADSHEET_ID

            url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"
            response = requests.get(url, timeout=10)

            if response.status_code != 200:
                return None

            # Pass sayfasını oku
            df = pd.read_excel(BytesIO(response.content), sheet_name="Pass")

            # AyarlarModul değerini bul
            # Pass sayfasının yapısına göre değiştirmeniz gerekebilir
            # Varsayım: İlk sütunda isim, ikinci sütunda değer var
            for index, row in df.iterrows():
                if str(row.iloc[0]).strip() == "AyarlarModul":
                    return str(row.iloc[1]).strip()

            return None

        except Exception as e:
            print(f"Şifre okuma hatası: {e}")
            return None

    def _get_virman_password_from_sheets(self):
        """PRGsheet dosyasının Pass sayfasından VirmanModul şifresini oku"""
        try:
            # Service Account ile PRGsheet'e erişim
            config_manager = CentralConfigManager()
            spreadsheet_id = config_manager.MASTER_SPREADSHEET_ID

            url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"
            response = requests.get(url, timeout=10)

            if response.status_code != 200:
                return None

            # Pass sayfasını oku
            df = pd.read_excel(BytesIO(response.content), sheet_name="Pass")

            # VirmanModul değerini bul
            # Modul sütununda "VirmanModul" olan satırın Password sütunundaki değeri al
            for index, row in df.iterrows():
                # 'Modul' sütununu kontrol et
                modul_value = str(row.get('Modul', '')).strip() if 'Modul' in df.columns else str(row.iloc[0]).strip()
                if modul_value == "VirmanModul":
                    # 'Password' sütunundan değeri al
                    password_value = str(row.get('Password', '')).strip() if 'Password' in df.columns else str(row.iloc[1]).strip()
                    return password_value

            return None

        except Exception as e:
            print(f"Virman şifre okuma hatası: {e}")
            return None
    
    def _setup_ui(self):
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        main_layout = QVBoxLayout(self.central_widget)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        self.navigation_bar = AdvancedNavigationBar(
            self.module_registry, 
            self.theme_manager, 
            self.event_bus
        )
        self.page_manager = AdvancedPageManager(
            self.module_registry, 
            self.event_bus
        )
        
        main_layout.addWidget(self.navigation_bar)
        main_layout.addWidget(self.page_manager, 1)
        
        self._setup_status_bar()
    
    def _setup_status_bar(self):
        # Status bar oluştur
        self.status_bar = self.statusBar()
        self.status_bar.setStyleSheet("""
            QStatusBar {
                background-color: #f8f9fa;
                color: #343a40;
                border-top: 1px solid #dee2e6;
                font-size: 12px;
                padding: 5px;
            }
        """)

        # Sağ tarafa kalıcı copyright widget'ı ekle
        copyright_label = QLabel("© by İsmail Güneş   ")
        copyright_label.setStyleSheet("""
            background-color: transparent;
            color: black;
            padding-right: 17px;
        """)
        self.status_bar.addPermanentWidget(copyright_label)
    
    def _connect_events(self):
        from core_architecture import EventType
        
        self.navigation_bar.module_requested.connect(self._on_module_requested)
        self.event_bus.subscribe(EventType.PAGE_CHANGED, self._on_page_changed)
        self.event_bus.subscribe(EventType.THEME_CHANGED, self._on_theme_changed)
        
        # Dark tema stilini uygula
        self.setStyleSheet(self.theme_manager.get_main_style())
    
    def _on_module_requested(self, module_type: ModuleType):
        # Ayarlar modülüne geçmeden önce şifre kontrolü
        if module_type == ModuleType.AYARLAR:
            if not self._verify_ayarlar_password():
                # Şifre yanlışsa geri dön, modül değiştirme
                # Navigation bar'ı eski modüle geri al
                if self.app_state.current_module:
                    self.navigation_bar.set_active_module(self.app_state.current_module)
                return

        # Virman modülüne geçmeden önce şifre kontrolü
        if module_type == ModuleType.VIRMAN:
            if not self._verify_virman_password():
                # Şifre yanlışsa geri dön, modül değiştirme
                # Navigation bar'ı eski modüle geri al
                if self.app_state.current_module:
                    self.navigation_bar.set_active_module(self.app_state.current_module)
                return

        command = NavigateCommand(self.app_state, module_type)
        self.command_invoker.execute_command(command)
        self.page_manager.show_module(module_type)

        # SSH modülü seçildiyse status message'ları bağla
        if module_type == ModuleType.SSH:
            current_widget = self.page_manager.stacked_widget.currentWidget()
            if hasattr(current_widget, 'status_message'):
                current_widget.status_message.connect(self._update_status_bar)

    def _verify_ayarlar_password(self):
        """Ayarlar modülü şifre doğrulama"""
        # Şifre cache'de yoksa Google Sheets'ten al
        if self.ayarlar_password is None:
            self.ayarlar_password = self._get_ayarlar_password_from_sheets()

        # Şifre alınamadıysa
        if self.ayarlar_password is None:
            QMessageBox.warning(
                self,
                "Şifre Hatası",
                "Ayarlar modülü şifresi Google Sheets'ten okunamadı.\n"
                "Pass sayfasında 'AyarlarModul' kaydını kontrol edin."
            )
            return False

        # Kullanıcıdan şifre iste
        password, ok = QInputDialog.getText(
            self,
            "Ayarlar Modülü",
            "Şifre:",
            echo=QLineEdit.Password
        )

        if not ok:
            return False

        # Şifreyi doğrula
        if password.strip() == self.ayarlar_password:
            return True
        else:
            QMessageBox.warning(
                self,
                "Hatalı Şifre",
                "Girdiğiniz şifre yanlış!"
            )
            return False

    def _verify_virman_password(self):
        """Virman modülü şifre doğrulama"""
        # Şifre cache'de yoksa Google Sheets'ten al
        if self.virman_password is None:
            self.virman_password = self._get_virman_password_from_sheets()

        # Şifre alınamadıysa
        if self.virman_password is None:
            QMessageBox.warning(
                self,
                "Şifre Hatası",
                "Virman modülü şifresi Google Sheets'ten okunamadı.\n"
                "Pass sayfasında 'VirmanModul' kaydını kontrol edin."
            )
            return False

        # Kullanıcıdan şifre iste
        password, ok = QInputDialog.getText(
            self,
            "Virman Modülü",
            "Şifre:",
            echo=QLineEdit.Password
        )

        if not ok:
            return False

        # Şifreyi doğrula
        if password.strip() == self.virman_password:
            return True
        else:
            QMessageBox.warning(
                self,
                "Hatalı Şifre",
                "Girdiğiniz şifre yanlış!"
            )
            return False

    def _update_status_bar(self, message: str):
        """Status bar mesajını güncelle"""
        self.status_bar.showMessage(message)
    
    def _on_page_changed(self, data):
        current = data['current']
        self.navigation_bar.set_active_module(current)
    
    def _on_theme_changed(self, theme):
        self.setStyleSheet(self.theme_manager.get_main_style())
    
    def _show_initial_module(self):
        initial_module = ModuleType.SOZLESME  # Sözleşme modülü ile başla
        command = NavigateCommand(self.app_state, initial_module)
        self.command_invoker.execute_command(command)
        self.page_manager.show_module(initial_module)


def main():
    try:
        app = QApplication(sys.argv)

        # Set application icon for taskbar (Windows)
        app.setWindowIcon(get_app_icon())

        window = ModernMainApp()

        # Ekran boyutunu al ve ona göre maximize et (geometry hatası önlenmesi için)
        screen = app.primaryScreen()
        screen_geometry = screen.availableGeometry()
        window.setGeometry(screen_geometry)
        window.showMaximized()

        sys.exit(app.exec_())
    except Exception as e:
        import traceback
        with open("error_log.txt", "w", encoding="utf-8") as f:
            f.write(f"Hata: {str(e)}\n")
            f.write(f"Detaylı hata:\n{traceback.format_exc()}")
        raise


if __name__ == "__main__":
    main()