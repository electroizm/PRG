"""
UI Components - Navigasyon ve sayfa yönetimi bileşenleri
"""

from PyQt5.QtCore import pyqtSignal, Qt
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QStackedWidget, QLabel, QProgressBar, QSizePolicy, QLayout)

from core_architecture import (ModuleType, EventType, ModuleRegistry,
                                ThemeManager, EventBus)


class AdvancedNavigationBar(QWidget):
    module_requested = pyqtSignal(ModuleType)
    
    def __init__(self, module_registry, theme_manager, event_bus):
        super().__init__()
        self.module_registry = module_registry
        self.theme_manager = theme_manager
        self.event_bus = event_bus
        self.buttons = {}
        self.current_module = None
        
        self._setup_ui()
        self._connect_events()
    
    def _setup_ui(self):
        layout = QHBoxLayout(self)
        layout.setSpacing(5)
        layout.setContentsMargins(10, 5, 10, 5)
        layout.setSizeConstraint(QLayout.SetDefaultConstraint)  # Minimum size zorlamasını devre dışı bırak
        
        for module_type, config in self.module_registry.get_enabled_modules().items():
            button = QPushButton(config.title)
            button.setMinimumHeight(24)
            button.setMinimumWidth(80)   # Minimum genişlik
            button.setMaximumWidth(150)  # Maksimum genişlik sınırı (geometry hatası için)
            button.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)  # Yatayda esnek, dikeyde sabit
            button.setCheckable(True)
            button.clicked.connect(lambda checked, mt=module_type: self._on_button_clicked(mt))

            self.buttons[module_type] = button
            layout.addWidget(button)
        
        layout.addStretch()
        self.setMaximumHeight(50)  # Navigation bar yüksekliğini sınırla
        self._apply_theme()
    
    def _connect_events(self):
        self.event_bus.subscribe(EventType.THEME_CHANGED, self._apply_theme)
    
    def _on_button_clicked(self, module_type: ModuleType):
        self.set_active_module(module_type)
        self.module_requested.emit(module_type)
    
    def set_active_module(self, module_type: ModuleType):
        if self.current_module == module_type:
            return
        
        for mt, button in self.buttons.items():
            button.setChecked(mt == module_type)
        
        self.current_module = module_type
    
    def _apply_theme(self, theme=None):
        style = self.theme_manager.get_button_style()
        for button in self.buttons.values():
            button.setStyleSheet(style)


class AdvancedPageManager(QWidget):
    def __init__(self, module_registry, event_bus):
        super().__init__()
        self.module_registry = module_registry
        self.event_bus = event_bus
        self.stacked_widget = QStackedWidget()
        self.loading_widget = self._create_loading_widget()
        
        self._setup_ui()
    
    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.stacked_widget)
        
        self.stacked_widget.addWidget(self.loading_widget)
    
    def _create_loading_widget(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        label = QLabel("Yükleniyor...")
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 24px; color: #ffffff;")

        progress = QProgressBar()
        progress.setRange(0, 0)
        progress.setMaximumWidth(300)
        progress.setStyleSheet("""
            QProgressBar {
                border: 2px solid #2d2d2d;
                border-radius: 5px;
                background-color: #1a1a1a;
                color: #ffffff;
            }
            QProgressBar::chunk {
                background-color: #007acc;
                border-radius: 3px;
            }
        """)

        layout.addStretch()
        layout.addWidget(label)
        layout.addSpacing(20)
        layout.addWidget(progress, 0, Qt.AlignCenter)
        layout.addStretch()

        return widget
    
    def show_module(self, module_type: ModuleType):
        from PyQt5.QtWidgets import QApplication

        # Modül zaten yüklenmişse direkt göster (loading ekranı atla)
        instance = self.module_registry._instances.get(module_type)

        if instance and instance in [self.stacked_widget.widget(i) for i in range(self.stacked_widget.count())]:
            # Zaten cache'lenmiş modül - anında göster
            self.stacked_widget.setCurrentWidget(instance)
            return

        # Yeni modül yüklenecek - loading göster
        self.stacked_widget.setCurrentWidget(self.loading_widget)
        QApplication.processEvents()

        # Modülü oluştur
        instance = self.module_registry.create_module_instance(module_type)
        if instance:
            if instance not in [self.stacked_widget.widget(i) for i in range(self.stacked_widget.count())]:
                self.stacked_widget.addWidget(instance)
            self.stacked_widget.setCurrentWidget(instance)