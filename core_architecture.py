"""
Core Architecture - Modern mimarinin temel sınıfları
"""

import weakref
from collections import defaultdict
from enum import Enum, auto
from typing import Type, Dict, Optional, Any, Callable, List
from dataclasses import dataclass, field
from abc import ABC, abstractmethod

from PyQt5.QtCore import pyqtSignal
from PyQt5.QtWidgets import QWidget


class EventType(Enum):
    PAGE_CHANGED = "page_changed"
    THEME_CHANGED = "theme_changed"
    MODULE_LOADED = "module_loaded"
    ACTION_EXECUTED = "action_executed"


class ModuleType(Enum):
    STOK = "stok"
    SEVKIYAT = "sevkiyat"
    SOZLESME = "sozlesme"
    OKC = "okc"
    RISK = "risk"
    SSH = "ssh"
    AYARLAR = "ayarlar"
    KASA = "kasa"
    VIRMAN = "virman"
    SANALPOS = "sanalpos"
    IRSALIYE = "irsaliye"
    FIYAT = "fiyat"


@dataclass
class Theme:
    name: str
    primary_color: str = "#2c3e50"
    secondary_color: str = "#34495e"
    accent_color: str = "#3498db"
    background_color: str = "#ecf0f1"
    text_color: str = "#2c3e50"
    button_hover_color: str = "#34495e"
    is_dark: bool = False


@dataclass
class ModuleConfig:
    name: str
    title: str
    module_type: ModuleType
    widget_class: Type
    icon: Optional[str] = None
    description: str = ""
    dependencies: List[str] = field(default_factory=list)
    enabled: bool = True


class ICommand(ABC):
    @abstractmethod
    def execute(self) -> Any:
        pass
    
    @abstractmethod
    def undo(self) -> Any:
        pass
    
    @abstractmethod
    def description(self) -> str:
        pass


class NavigateCommand(ICommand):
    def __init__(self, app_state, target_module: ModuleType, previous_module: Optional[ModuleType] = None):
        self.app_state = app_state
        self.target_module = target_module
        self.previous_module = previous_module or app_state.current_module
    
    def execute(self) -> Any:
        self.app_state.set_current_module(self.target_module)
        return self.target_module
    
    def undo(self) -> Any:
        if self.previous_module:
            self.app_state.set_current_module(self.previous_module)
        return self.previous_module
    
    def description(self) -> str:
        return f"Navigate to {self.target_module.value}"


class ChangeThemeCommand(ICommand):
    def __init__(self, theme_manager, new_theme: Theme, previous_theme: Optional[Theme] = None):
        self.theme_manager = theme_manager
        self.new_theme = new_theme
        self.previous_theme = previous_theme or theme_manager.current_theme
    
    def execute(self) -> Any:
        self.theme_manager.set_theme(self.new_theme)
        return self.new_theme
    
    def undo(self) -> Any:
        if self.previous_theme:
            self.theme_manager.set_theme(self.previous_theme)
        return self.previous_theme
    
    def description(self) -> str:
        return f"Change theme to {self.new_theme.name}"


class EventBus:
    def __init__(self):
        self._observers = defaultdict(list)
        self._weak_observers = defaultdict(list)
    
    def subscribe(self, event_type: EventType, callback: Callable, weak_ref: bool = False):
        if weak_ref:
            self._weak_observers[event_type].append(weakref.ref(callback))
        else:
            self._observers[event_type].append(callback)
    
    def unsubscribe(self, event_type: EventType, callback: Callable):
        if callback in self._observers[event_type]:
            self._observers[event_type].remove(callback)
    
    def emit(self, event_type: EventType, data: Any = None):
        for callback in self._observers[event_type]:
            try:
                callback(data)
            except Exception as e:
                print(f"Error in event callback: {e}")
        
        for weak_callback in self._weak_observers[event_type][:]:
            callback = weak_callback()
            if callback is None:
                self._weak_observers[event_type].remove(weak_callback)
            else:
                try:
                    callback(data)
                except Exception as e:
                    print(f"Error in weak event callback: {e}")


class AppState:
    def __init__(self, event_bus):
        self.event_bus = event_bus
        self._current_module: Optional[ModuleType] = None
        self._module_history: List[ModuleType] = []
        self._user_preferences = {}
    
    @property
    def current_module(self) -> Optional[ModuleType]:
        return self._current_module
    
    def set_current_module(self, module: ModuleType):
        if self._current_module != module:
            if self._current_module:
                self._module_history.append(self._current_module)
            self._current_module = module
            self.event_bus.emit(EventType.PAGE_CHANGED, {
                'current': module,
                'previous': self._module_history[-1] if self._module_history else None
            })
    
    def get_previous_module(self) -> Optional[ModuleType]:
        return self._module_history[-1] if self._module_history else None
    
    def set_preference(self, key: str, value: Any):
        self._user_preferences[key] = value
    
    def get_preference(self, key: str, default: Any = None) -> Any:
        return self._user_preferences.get(key, default)


class ThemeManager:
    def __init__(self, event_bus):
        self.event_bus = event_bus
        self.themes = self._load_default_themes()
        self.current_theme = self.themes["dark"]
    
    def _load_default_themes(self) -> Dict[str, Theme]:
        return {
            "dark": Theme(
                name="Dark",
                primary_color="#1a1a1a",
                secondary_color="#2d2d2d",
                accent_color="#007acc",
                background_color="#0d1117",
                text_color="#ffffff",
                button_hover_color="#404040",
                is_dark=True
            )
        }
    
    def set_theme(self, theme: Theme):
        self.current_theme = theme
        self.event_bus.emit(EventType.THEME_CHANGED, theme)
    
    def get_button_style(self) -> str:
        theme = self.current_theme
        return f"""
            QPushButton {{
                background-color: {theme.primary_color};
                color: {theme.text_color if not theme.is_dark else '#ffffff'};
                padding: 6px 16px;
                font-size: 14px;
                font-weight: bold;
                border-radius: 6px;
                min-width: 120px;
                border: 2px solid {theme.secondary_color};
                font-family: 'Segoe UI', Arial, sans-serif;
            }}
            QPushButton:hover {{
                background-color: {theme.button_hover_color};
                border-color: {theme.accent_color};
            }}
            QPushButton:checked {{
                background-color: {theme.accent_color};
                border-color: {theme.accent_color};
            }}
        """
    
    def get_main_style(self) -> str:
        theme = self.current_theme
        return f"""
            QMainWindow {{
                background-color: {theme.background_color};
                color: {theme.text_color};
            }}
            QWidget {{
                background-color: {theme.background_color};
                color: {theme.text_color};
            }}
            QLabel {{
                color: {theme.text_color};
            }}
            QStatusBar {{
                background-color: {theme.secondary_color};
                color: {theme.text_color};
                border-top: 1px solid {theme.accent_color};
            }}
        """


class CommandInvoker:
    def __init__(self):
        self._history: List[ICommand] = []
        self._current_index = -1
    
    def execute_command(self, command: ICommand) -> Any:
        if self._current_index < len(self._history) - 1:
            self._history = self._history[:self._current_index + 1]
        
        result = command.execute()
        self._history.append(command)
        self._current_index += 1
        return result
    
    def undo(self) -> Optional[Any]:
        if self._current_index >= 0:
            command = self._history[self._current_index]
            result = command.undo()
            self._current_index -= 1
            return result
        return None
    
    def redo(self) -> Optional[Any]:
        if self._current_index < len(self._history) - 1:
            self._current_index += 1
            command = self._history[self._current_index]
            return command.execute()
        return None
    
    def can_undo(self) -> bool:
        return self._current_index >= 0
    
    def can_redo(self) -> bool:
        return self._current_index < len(self._history) - 1


class ModuleRegistry:
    def __init__(self, event_bus):
        self.event_bus = event_bus
        self._modules: Dict[ModuleType, ModuleConfig] = {}
        self._instances: Dict[ModuleType, Any] = {}
    
    def register_module(self, config: ModuleConfig):
        self._modules[config.module_type] = config
        self.event_bus.emit(EventType.MODULE_LOADED, config)
    
    def get_module_config(self, module_type: ModuleType) -> Optional[ModuleConfig]:
        return self._modules.get(module_type)
    
    def get_all_modules(self) -> Dict[ModuleType, ModuleConfig]:
        return self._modules.copy()
    
    def get_enabled_modules(self) -> Dict[ModuleType, ModuleConfig]:
        return {k: v for k, v in self._modules.items() if v.enabled}
    
    def create_module_instance(self, module_type: ModuleType) -> Optional[Any]:
        if module_type not in self._instances:
            config = self._modules.get(module_type)
            if config and config.enabled:
                try:
                    instance = config.widget_class()
                    self._instances[module_type] = instance
                    return instance
                except Exception as e:
                    print(f"Error creating module {module_type}: {e}")
                    return None
        return self._instances.get(module_type)