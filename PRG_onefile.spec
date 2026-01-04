# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller Spec Dosyası - PRG
Tek exe dosyası - Tüm modüller ve bağımlılıklar dahil
"""

import sys
import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# Ana dizinler - mutlak yollar kullan
spec_dir = r'D:\GoogleDrive\PRG\OAuth2'
prg_dir = r'D:\GoogleDrive\PRG\OAuth2\PRG'

# Giriş noktası (entry point)
entry_script = os.path.join(prg_dir, 'run.py')

# Tüm PRG modüllerini hidden import olarak ekle
prg_hidden_imports = [
    'PRG',
    'PRG.main',
    'PRG.core_architecture',
    'PRG.ui_components',
    'PRG.embedded_resources',
    'PRG.Sozleme',
    'PRG.ayar_module',
    'PRG.fiyat_module',
    'PRG.irsaliye_module',
    'PRG.kasa_module',
    'PRG.okc_module',
    'PRG.risk_module',
    'PRG.sanalpos_module',
    'PRG.sevkiyat_module',
    'PRG.sozlesme_module',
    'PRG.ssh_module',
    'PRG.stok_module',
    'PRG.virman_module',
]

# Harici bağımlılıklar
external_imports = [
    'central_config',

    # PyQt5 - Tam dahil etme
    'PyQt5',
    'PyQt5.QtCore',
    'PyQt5.QtGui',
    'PyQt5.QtWidgets',
    'PyQt5.QtPrintSupport',
    'PyQt5.sip',

    # Veri işleme kütüphaneleri
    'pandas',
    'pandas.io.formats.style',
    'pandas.io.clipboard',
    'pandas.plotting',
    'numpy',
    'numpy.core',
    'openpyxl',

    # Google servisleri
    'gspread',
    'google',
    'google.auth',
    'google.auth.transport',
    'google.auth.transport.requests',
    'google.oauth2',
    'google.oauth2.service_account',
    'google.oauth2.credentials',

    # Ağ ve veritabanı
    'requests',
    'pyodbc',
    'urllib3',

    # Güvenlik/Şifreleme
    'cryptography',
    'cryptography.fernet',
    'cryptography.hazmat',
    'cryptography.hazmat.primitives',
    'cryptography.hazmat.backends',

    # Diğer
    'certifi',
    'charset_normalizer',
    'dateutil',
    'pytz',
    'lxml',
    'lxml.etree',
]

# Tüm hidden importları birleştir
hidden_imports = prg_hidden_imports + external_imports

# Dahil edilecek veri dosyaları
datas = [
    # İkonlar
    (os.path.join(prg_dir, 'icon.ico'), 'PRG'),
    (os.path.join(prg_dir, 'icon.jpg'), 'PRG'),

    # Ana dizindeki yapılandırma dosyaları
    (os.path.join(spec_dir, 'service_account.json'), '.'),
    (os.path.join(spec_dir, 'central_config.py'), '.'),

    # PRG klasöründeki tüm Python dosyalarını veri olarak dahil et (dinamik import için)
    (os.path.join(prg_dir, 'Sozleme.py'), 'PRG'),
    (os.path.join(prg_dir, '__init__.py'), 'PRG'),
    (os.path.join(prg_dir, 'main.py'), 'PRG'),
    (os.path.join(prg_dir, 'core_architecture.py'), 'PRG'),
    (os.path.join(prg_dir, 'ui_components.py'), 'PRG'),
    (os.path.join(prg_dir, 'embedded_resources.py'), 'PRG'),
    (os.path.join(prg_dir, 'ayar_module.py'), 'PRG'),
    (os.path.join(prg_dir, 'fiyat_module.py'), 'PRG'),
    (os.path.join(prg_dir, 'irsaliye_module.py'), 'PRG'),
    (os.path.join(prg_dir, 'kasa_module.py'), 'PRG'),
    (os.path.join(prg_dir, 'okc_module.py'), 'PRG'),
    (os.path.join(prg_dir, 'risk_module.py'), 'PRG'),
    (os.path.join(prg_dir, 'sanalpos_module.py'), 'PRG'),
    (os.path.join(prg_dir, 'sevkiyat_module.py'), 'PRG'),
    (os.path.join(prg_dir, 'sozlesme_module.py'), 'PRG'),
    (os.path.join(prg_dir, 'ssh_module.py'), 'PRG'),
    (os.path.join(prg_dir, 'stok_module.py'), 'PRG'),
    (os.path.join(prg_dir, 'virman_module.py'), 'PRG'),
]

# Certifi SSL sertifikalarını ekle
datas += collect_data_files('certifi')

a = Analysis(
    [entry_script],
    pathex=[spec_dir, prg_dir],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',      # Kullanılmayan GUI framework
        'matplotlib',   # Kullanılmıyor
        'PIL',          # Kullanılmıyor
        'IPython',      # Geliştirme aracı
        'jupyter',      # Geliştirme aracı
        'notebook',     # Geliştirme aracı
        'pytest',       # Test aracı
        'sphinx',       # Dokümantasyon aracı
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='PRG',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,               # UPX sıkıştırma (dosya boyutunu küçültür)
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # Konsol penceresi AÇILMASIN
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join(prg_dir, 'icon.ico'),  # Uygulama ikonu
    version_file=None,
)
