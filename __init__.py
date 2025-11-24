"""
PRG - Modüler PyQt5 Uygulaması

Bu paket, dosyadaki tüm modülleri
ayrı dosyalara bölerek organize edilmiş halidir.
"""

from main import main

__version__ = "2025.11.11"
__author__ = "İsmail Güneş"

# Public API - only export main function
# ModernMainApp is an internal implementation detail
__all__ = [
    'main'
]