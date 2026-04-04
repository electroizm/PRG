"""
Barkod Module - Mikro SQL Server'dan Satış Faturalarını Supabase'e Senkronize Etme
"""

import os
import sys
import logging
import math
import pyodbc
from datetime import datetime, timedelta

import requests
import pandas as pd

# Üst dizini Python path'e ekle (central_config için)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from central_config import CentralConfigManager  # type: ignore

from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QProgressBar, QLabel, QTableWidget, QTableWidgetItem,
                             QHeaderView, QTextEdit, QLineEdit, QTabWidget,
                             QApplication, QMenu, QAction, QShortcut, QMessageBox,
                             QStyledItemDelegate, QStyle, QListWidget, QListWidgetItem)
from PyQt5.QtGui import QFont, QKeySequence, QColor, QBrush

# Logger
logger = logging.getLogger(__name__)


class NoFocusDelegate(QStyledItemDelegate):
    """Hücre seçildiğinde noktalı çerçeveyi (focus rectangle) kaldırır."""
    def paint(self, painter, option, index):
        option.state &= ~QStyle.State_HasFocus
        super().paint(painter, option, index)


_TR_CHAR_MAP = str.maketrans(
    u'\u015f\u015e\u00fc\u00dc\u00f6\u00d6\u00e7\u00c7\u0131\u0130\u011f\u011e',
    'sSuUoOcCiIgG'
)


def _normalize_turkish(text: str) -> str:
    """Turkce karakterleri ASCII karsiliklarina donusturur ve kucuk harfe cevirir."""
    return str(text).translate(_TR_CHAR_MAP).lower()


def _fuzzy_match(search_text: str, data_text: str) -> bool:
    """Sirayla bagli olmadan tum kelimelerin data icinde olup olmadigini kontrol eder.
    Turkce karakter normalize edilmis haliyle arar."""
    normalized_data = _normalize_turkish(data_text)
    words = _normalize_turkish(search_text).split()
    return all(w in normalized_data for w in words)


def _normalize_stok_kod(stok_kod: str) -> str:
    """stok_kod'u eslestirme icin normalize et - ilk '-' oncesini al.
    satis_faturasi: '3200418840-0' veya '3120013399-4495040-1223'
    satis_faturasi_okumalari: her zaman 10 haneli base kod (orn: '3200418840')"""
    s = str(stok_kod or '')
    dash_idx = s.find('-')
    if dash_idx > 0:
        return s[:dash_idx]
    return s


def _build_okuma_durumu_widget(miktar, paket, paket_readings, depo_no=None):
    """Okuma durumu widget'i olusturur - renkli P1, P2, P3... etiketleri.

    Args:
        miktar: Kac adet urun (grup sayisi)
        paket: Her urunun kac paketi var (birim paket)
        paket_readings: {paket_sira: [{'type','user','time'}, ...]}
        depo_no: Depo numarasi (manuel renk farki icin, None ise hep turuncu)
    """
    widget = QWidget()
    h_layout = QHBoxLayout(widget)
    h_layout.setContentsMargins(4, 2, 4, 2)
    h_layout.setSpacing(2)

    for unit_idx in range(miktar):
        if unit_idx > 0:
            sep = QLabel("|")
            sep.setAlignment(Qt.AlignCenter)
            sep.setFixedWidth(12)
            sep.setStyleSheet(
                "color: #9ca3af; font-size: 13px; font-weight: bold;"
            )
            h_layout.addWidget(sep)

        for p in range(1, paket + 1):
            reads = paket_readings.get(p, [])
            info = reads[unit_idx] if unit_idx < len(reads) else None
            reading_type = info['type'] if info else None

            label = QLabel(f"P{p}")
            label.setAlignment(Qt.AlignCenter)
            label.setFixedSize(30, 22)

            if reading_type == 'scanner':
                label.setStyleSheet(
                    "background-color: #86efac; color: #1f2937; "
                    "border-radius: 3px; font-size: 11px; font-weight: bold;"
                )
            elif reading_type == 'manual':
                if depo_no == '100':
                    manual_bg = '#ef4444'
                else:
                    manual_bg = '#f97316'
                label.setStyleSheet(
                    f"background-color: {manual_bg}; color: #1f2937; "
                    "border-radius: 3px; font-size: 11px; font-weight: bold;"
                )
            else:
                label.setStyleSheet(
                    "background-color: #e5e7eb; color: #9ca3af; "
                    "border-radius: 3px; font-size: 11px; font-weight: bold;"
                )

            if info:
                label.setToolTip(f"{info['user']}\n{info['time']}")
                label.setCursor(Qt.PointingHandCursor)

            h_layout.addWidget(label)

    h_layout.addStretch()
    return widget


# ================== CONFIG CONSTANTS ==================
FONT_FAMILY = "Segoe UI"
FONT_SIZE = 12
ROW_HEIGHT = 35
REQUEST_TIMEOUT_SEC = 30
BATCH_SIZE = 50
DOGTAS_TOKEN_CACHE_MINUTES = 55
DOGTAS_API_BASE_URL = 'https://connectapi.doganlarmobilyagrubu.com'

TABLE_COLUMNS = [
    ('evrakno_sira', 'Evrak Sıra'),
    ('tarih', 'Tarih'),
    ('stok_kod', 'Stok Kod'),
    ('miktar', 'Miktar'),
    ('cikis_depo_no', 'Depo'),
    ('paket_sayisi', 'Paket'),
    ('cari_adi', 'Cari Adı'),
    ('product_desc', 'Ürün Açıklama'),
    ('plasiyer_kodu', 'Satıcı'),
    ('satinalma_kalem_id', 'BagKodu'),
    ('okuma_durumu', 'Okuma Durumu'),
]

NAKLIYE_TABLE_COLUMNS = [
    ('oturum_id', 'Oturum'),
    ('nakliye_no', 'Nakliye No'),
    ('plaka', 'Plaka'),
    ('belge_tarihi', 'Tarih'),
    ('fatura_numarasi', 'Fatura No'),
    ('malzeme_no', 'Malzeme No'),
    ('satinalma_kalem_id', u'Sat\u0131nalma Kalem No'),
    ('miktar', 'Miktar'),
    ('depo_yeri', 'Depo'),
    ('paket_sayisi', 'Paket'),
    ('malzeme_adi', u'Malzeme Ad\u0131'),
    ('okuma_durumu', 'Okuma Durumu'),
]

CIKIS_TABLE_COLUMNS = [
    ('evrakno_sira', u'Evrak S\u0131ra'),
    ('tarih', 'Tarih'),
    ('stok_kod', 'Stok Kod'),
    ('miktar', 'Miktar'),
    ('depo', 'Depo'),
    ('paket_sayisi', 'Paket'),
    ('malzeme_adi', u'Malzeme Ad\u0131'),
    ('satinalma_kalem_id', 'BagKodu'),
    ('okuma_durumu', 'Okuma Durumu'),
]

GIRIS_TABLE_COLUMNS = [
    ('evrakno_sira', u'Evrak S\u0131ra'),
    ('tarih', 'Tarih'),
    ('stok_kod', 'Stok Kod'),
    ('miktar', 'Miktar'),
    ('depo', 'Depo'),
    ('paket_sayisi', 'Paket'),
    ('malzeme_adi', u'Malzeme Ad\u0131'),
    ('satinalma_kalem_id', 'BagKodu'),
    ('okuma_durumu', 'Okuma Durumu'),
]

SEVK_TABLE_COLUMNS = [
    ('evrakno_sira', u'Evrak S\u0131ra'),
    ('tarih', 'Tarih'),
    ('stok_kod', 'Stok Kod'),
    ('miktar', 'Miktar'),
    ('cikis_depo', u'\u00c7\u0131k\u0131\u015f Depo'),
    ('giris_depo', u'Giri\u015f Depo'),
    ('paket_sayisi', 'Paket'),
    ('malzeme_adi', u'Malzeme Ad\u0131'),
    ('satinalma_kalem_id', 'BagKodu'),
    ('okuma_durumu', 'Okuma Durumu'),
]

QR_LOG_TABLE_COLUMNS = [
    ('tarih', 'Tarih'),
    ('yon', u'Y\u00f6n'),
    ('kaynak', 'Kaynak'),
    ('evrak_no', 'Evrak No'),
    ('cari_adi', u'Cari Ad\u0131'),
    ('paket_sira', 'Paket'),
    ('qr_kod_kisa', 'QR Kod'),
    ('kullanici', u'Kullan\u0131c\u0131'),
]

# Giris (+) ve Cikis (-) tablo siniflandirmasi
_QR_GIRIS_KAYNAKLARI = {'Nakliye', u'Giri\u015f', 'Sevk', u'Say\u0131m'}
_QR_CIKIS_KAYNAKLARI = {u'Sat\u0131\u015f', u'\u00c7\u0131k\u0131\u015f'}

SAYIM_TABLE_COLUMNS = [
    ('sayim_kodu', u'Say\u0131m Kodu'),
    ('malzeme_kodu', 'Malzeme Kodu'),
    ('malzeme_adi', u'Malzeme Ad\u0131'),
    ('miktar', u'Say\u0131lan'),
    ('beklenen', 'Beklenen'),
    ('fark', 'Fark'),
    ('paket_sayisi', 'Paket'),
    ('okuma_durumu', 'Okuma Durumu'),
]

MIKRO_SQL_QUERY = """
    SELECT
        sth.sth_evrakno_seri,
        sth.sth_evrakno_sira,
        sth.sth_satirno,
        CONVERT(DATE, sth.sth_belge_tarih) AS tarih,
        sth.sth_stok_kod,
        sth.sth_miktar,
        sth.sth_cikis_depo_no,
        dbo.fn_StokHarEvrTip(sth.sth_evraktip) AS evrak_adi,
        cha.cha_kod AS cari_kodu,
        dbo.fn_CarininIsminiBul(cha.cha_cari_cins, cha.cha_kod) AS cari_adi,
        bar.bar_serino_veya_bagkodu AS bag_kodu,
        sto.sto_isim AS malzeme_adi,
        sth.sth_plasiyer_kodu
    FROM dbo.STOK_HAREKETLERI sth WITH (NOLOCK)
    LEFT JOIN dbo.CARI_HESAP_HAREKETLERI cha WITH (NOLOCK)
        ON sth.sth_evrakno_seri = cha.cha_evrakno_seri
        AND sth.sth_evrakno_sira = cha.cha_evrakno_sira
        AND cha.cha_evrak_tip = 63
    LEFT JOIN dbo.BARKOD_TANIMLARI bar WITH (NOLOCK)
        ON sth.sth_stok_kod = bar.bar_stokkodu
    LEFT JOIN dbo.STOKLAR sto WITH (NOLOCK)
        ON sto.sto_kod = sth.sth_stok_kod
        AND (sto.sto_pasif_fl IS NULL OR sto.sto_pasif_fl = 0)
    WHERE sth.sth_evraktip = 4
        AND sth.sth_belge_tarih >= ?
    ORDER BY sth.sth_evrakno_sira DESC
"""

MIKRO_CARI_ADRES_QUERY = """
    SELECT DISTINCT
        car.cari_kod,
        car.cari_unvan1,
        car.cari_vdaire_adi,
        car.cari_CepTel,
        adr.adr_cadde,
        adr.adr_sokak,
        adr.adr_posta_kodu,
        adr.adr_ilce,
        adr.adr_il
    FROM dbo.CARI_HESAPLAR car WITH (NOLOCK)
    LEFT JOIN dbo.CARI_HESAP_ADRESLERI adr WITH (NOLOCK)
        ON car.cari_kod = adr.adr_cari_kod
        AND adr.adr_adres_no = (
            SELECT MIN(a2.adr_adres_no)
            FROM dbo.CARI_HESAP_ADRESLERI a2 WITH (NOLOCK)
            WHERE a2.adr_cari_kod = car.cari_kod
        )
    WHERE car.cari_kod IN ({placeholders})
"""

MIKRO_CIKIS_SQL_QUERY = """
    SELECT
        sth.sth_evrakno_seri,
        sth.sth_evrakno_sira,
        CONVERT(DATE, sth.sth_belge_tarih) AS tarih,
        sth.sth_stok_kod,
        sth.sth_miktar,
        sth.sth_cikis_depo_no,
        dbo.fn_StokHarEvrTip(sth.sth_evraktip) AS evrak_adi,
        bar.bar_serino_veya_bagkodu AS bag_kodu,
        sto.sto_isim AS malzeme_adi
    FROM dbo.STOK_HAREKETLERI sth WITH (NOLOCK)
    LEFT JOIN dbo.BARKOD_TANIMLARI bar WITH (NOLOCK)
        ON bar.bar_stokkodu = sth.sth_stok_kod
    LEFT JOIN dbo.STOKLAR sto WITH (NOLOCK)
        ON sto.sto_kod = sth.sth_stok_kod
        AND (sto.sto_pasif_fl IS NULL OR sto.sto_pasif_fl = 0)
    WHERE sth.sth_evraktip = 0
        AND sth.sth_belge_tarih >= ?
    ORDER BY sth.sth_evrakno_sira DESC
"""

MIKRO_GIRIS_SQL_QUERY = """
    SELECT
        sth.sth_evrakno_seri,
        sth.sth_evrakno_sira,
        CONVERT(DATE, sth.sth_belge_tarih) AS tarih,
        sth.sth_stok_kod,
        sth.sth_miktar,
        dbo.fn_StokHarEvrTip(sth.sth_evraktip) AS evrak_adi,
        dbo.fn_StokHarDepoIsmi(sth.sth_giris_depo_no, sth.sth_cikis_depo_no, sth.sth_tip) AS depo,
        bar.bar_serino_veya_bagkodu AS bag_kodu,
        sto.sto_isim AS malzeme_adi
    FROM dbo.STOK_HAREKETLERI sth WITH (NOLOCK)
    LEFT JOIN dbo.BARKOD_TANIMLARI bar WITH (NOLOCK)
        ON bar.bar_stokkodu = sth.sth_stok_kod
    LEFT JOIN dbo.STOKLAR sto WITH (NOLOCK)
        ON sto.sto_kod = sth.sth_stok_kod
        AND (sto.sto_pasif_fl IS NULL OR sto.sto_pasif_fl = 0)
    WHERE sth.sth_evraktip = 12
        AND sth.sth_belge_tarih >= ?
    ORDER BY sth.sth_evrakno_sira DESC
"""

MIKRO_SEVK_SQL_QUERY = """
    SELECT
        sth.sth_evrakno_seri,
        sth.sth_evrakno_sira,
        CONVERT(DATE, sth.sth_belge_tarih) AS tarih,
        sth.sth_stok_kod,
        sth.sth_miktar,
        dbo.fn_StokHarEvrTip(sth.sth_evraktip) AS evrak_adi,
        dbo.fn_StokHarDepoIsmi(sth.sth_giris_depo_no, sth.sth_cikis_depo_no, 1) AS cikis_depo,
        dbo.fn_StokHarDepoIsmi(sth.sth_giris_depo_no, sth.sth_cikis_depo_no, 0) AS giris_depo,
        bar.bar_serino_veya_bagkodu AS bag_kodu,
        sto.sto_isim AS malzeme_adi
    FROM dbo.STOK_HAREKETLERI sth WITH (NOLOCK)
    LEFT JOIN dbo.BARKOD_TANIMLARI bar WITH (NOLOCK)
        ON bar.bar_stokkodu = sth.sth_stok_kod
    LEFT JOIN dbo.STOKLAR sto WITH (NOLOCK)
        ON sto.sto_kod = sth.sth_stok_kod
        AND (sto.sto_pasif_fl IS NULL OR sto.sto_pasif_fl = 0)
    WHERE sth.sth_evraktip = 2
        AND sth.sth_belge_tarih >= ?
    ORDER BY sth.sth_evrakno_sira DESC
"""


# ================== STYLESHEET CONSTANTS ==================
BUTTON_STYLE = """
    QPushButton {
        background-color: #dfdfdf;
        color: black;
        border: 1px solid #444;
        padding: 8px 16px;
        border-radius: 5px;
        font-size: 14px;
        font-weight: bold;
        min-width: 80px;
    }
    QPushButton:hover {
        background-color: #a0a5a2;
    }
    QPushButton:disabled {
        background-color: #f0f0f0;
        color: #999999;
    }
"""

SYNC_BUTTON_STYLE = """
    QPushButton {
        background-color: #2563eb;
        color: white;
        border: 1px solid #1d4ed8;
        padding: 8px 20px;
        border-radius: 5px;
        font-size: 14px;
        font-weight: bold;
        min-width: 160px;
    }
    QPushButton:hover {
        background-color: #1d4ed8;
    }
    QPushButton:disabled {
        background-color: #93c5fd;
        color: #dbeafe;
    }
"""

INFO_LABEL_STYLE = """
    QLabel {
        color: #374151;
        font-size: 13px;
        font-weight: bold;
        padding: 6px 10px;
        background-color: #f3f4f6;
        border: 1px solid #d1d5db;
        border-radius: 4px;
    }
"""

TABLE_STYLE = """
    QTableWidget {
        font-size: 15px;
        font-weight: bold;
        background-color: #ffffff;
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
        border: none;
        outline: none;
    }
    QTableWidget::item:focus {
        outline: none;
        border: none;
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
"""

SAYIM_TABLE_STYLE = """
    QTableWidget {
        font-size: 15px;
        font-weight: bold;
        background-color: #ffffff;
        gridline-color: #d0d0d0;
        border: 1px solid #d0d0d0;
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
"""

LOG_STYLE = """
    QTextEdit {
        background-color: #1a1a2e;
        color: #00ff00;
        font-family: Consolas, monospace;
        font-size: 12px;
        border: 1px solid #d0d0d0;
        padding: 4px;
    }
"""

FILTER_INPUT_STYLE = """
    QLineEdit {
        padding: 6px 10px;
        border: 1px solid #d1d5db;
        border-radius: 4px;
        font-size: 13px;
        background-color: #ffffff;
        color: #000000;
    }
    QLineEdit:focus {
        border-color: #2563eb;
    }
"""


# Oturum boyunca şifre önbelleği
_barkod_delete_verified = False


def _verify_barkod_delete_password(parent) -> bool:
    """BarkodApp silme şifresini doğrula. Oturumda bir kez doğrulanınca tekrar sorulmaz."""
    global _barkod_delete_verified
    if _barkod_delete_verified:
        return True

    # PRGsheet/Pass sayfasından şifreyi yükle
    try:
        import io as _io
        config_manager = CentralConfigManager()
        spreadsheet_id = config_manager.MASTER_SPREADSHEET_ID
        url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        pass_df = pd.read_excel(_io.BytesIO(response.content), sheet_name="Pass")
        row = pass_df[pass_df['Modul'] == 'BarkodApp']
        if row.empty:
            _show_message(parent, "Hata", "Pass sayfasında 'BarkodApp' kaydı bulunamadı.")
            return False
        correct_password = str(row.iloc[0]['Password']).strip()
    except Exception as e:
        _show_message(parent, "Hata", f"Şifre yüklenemedi:\n{str(e)}")
        return False

    # Şifre dialogu (özel stilize)
    from PyQt5.QtWidgets import QDialog, QLabel, QHBoxLayout, QVBoxLayout
    _btn_style = """
        QPushButton {
            background-color: #dfdfdf; color: #000000;
            border: 1px solid #444; padding: 6px 20px;
            font-size: 12px; font-weight: bold; border-radius: 4px; min-width: 70px;
        }
        QPushButton:hover { background-color: #a0a5a2; }
        QPushButton:pressed { background-color: #909090; }
    """
    dlg = QDialog(parent)
    dlg.setWindowTitle("Şifre Gerekli")
    dlg.setModal(True)
    dlg.setFixedWidth(320)
    dlg.setStyleSheet("background-color: #ffffff;")

    lbl = QLabel("Bu işlemi gerçekleştirmek için şifre giriniz:")
    lbl.setStyleSheet("color: #000000; font-size: 12px;")
    lbl.setWordWrap(True)

    pwd_input = QLineEdit()
    pwd_input.setEchoMode(QLineEdit.Password)
    pwd_input.setStyleSheet("""
        QLineEdit {
            background-color: #ffffff; color: #000000;
            border: 1px solid #888; border-radius: 3px;
            padding: 5px; font-size: 12px;
        }
    """)

    btn_ok = QPushButton("Tamam")
    btn_ok.setStyleSheet(_btn_style)
    btn_cancel = QPushButton("İptal")
    btn_cancel.setStyleSheet(_btn_style)
    btn_ok.clicked.connect(dlg.accept)
    btn_cancel.clicked.connect(dlg.reject)

    btn_row = QHBoxLayout()
    btn_row.addStretch()
    btn_row.addWidget(btn_ok)
    btn_row.addWidget(btn_cancel)

    layout = QVBoxLayout(dlg)
    layout.addWidget(lbl)
    layout.addWidget(pwd_input)
    layout.addLayout(btn_row)

    pwd_input.returnPressed.connect(dlg.accept)

    if dlg.exec_() != QDialog.Accepted:
        return False
    if pwd_input.text().strip() != correct_password:
        _show_message(parent, "Hata", "Yanlış şifre. İşlem iptal edildi.")
        return False

    _barkod_delete_verified = True
    return True


def _show_message(parent, title: str, message: str):
    """Stilize bilgi/hata mesaj dialogu."""
    from PyQt5.QtWidgets import QDialog, QLabel, QVBoxLayout, QHBoxLayout
    dlg = QDialog(parent)
    dlg.setWindowTitle(title)
    dlg.setModal(True)
    dlg.setFixedWidth(360)
    dlg.setStyleSheet("background-color: #ffffff;")

    lbl = QLabel(message)
    lbl.setStyleSheet("color: #000000; font-size: 12px;")
    lbl.setWordWrap(True)

    btn_ok = QPushButton("Tamam")
    btn_ok.setStyleSheet("""
        QPushButton {
            background-color: #dfdfdf; color: #000000;
            border: 1px solid #444; padding: 6px 20px;
            font-size: 12px; font-weight: bold; border-radius: 4px; min-width: 70px;
        }
        QPushButton:hover { background-color: #a0a5a2; }
        QPushButton:pressed { background-color: #909090; }
    """)
    btn_ok.clicked.connect(dlg.accept)

    btn_row = QHBoxLayout()
    btn_row.addStretch()
    btn_row.addWidget(btn_ok)

    layout = QVBoxLayout(dlg)
    layout.addWidget(lbl)
    layout.addLayout(btn_row)
    dlg.exec_()


def _confirm_delete(parent, message: str) -> bool:
    """Stilize silme onay dialogu. True döndürürse kullanıcı onayladı demektir."""
    from PyQt5.QtWidgets import QDialog, QLabel, QHBoxLayout, QVBoxLayout
    dlg = QDialog(parent)
    dlg.setWindowTitle("Onay")
    dlg.setModal(True)
    dlg.setStyleSheet("background-color: #ffffff;")

    lbl = QLabel(message)
    lbl.setStyleSheet("color: #000000; font-size: 13px; padding: 10px;")
    lbl.setWordWrap(True)

    _btn_style = """
        QPushButton {
            background-color: #dfdfdf; color: #000000;
            border: 1px solid #444; padding: 6px 20px;
            font-size: 12px; font-weight: bold; border-radius: 4px; min-width: 70px;
        }
        QPushButton:hover { background-color: #a0a5a2; }
        QPushButton:pressed { background-color: #909090; }
    """
    btn_yes = QPushButton("Evet")
    btn_yes.setStyleSheet(_btn_style)
    btn_no = QPushButton("Hayır")
    btn_no.setStyleSheet(_btn_style)
    btn_yes.clicked.connect(dlg.accept)
    btn_no.clicked.connect(dlg.reject)

    btn_layout = QHBoxLayout()
    btn_layout.addStretch()
    btn_layout.addWidget(btn_yes)
    btn_layout.addWidget(btn_no)

    layout = QVBoxLayout(dlg)
    layout.addWidget(lbl)
    layout.addLayout(btn_layout)

    return dlg.exec_() == QDialog.Accepted


def _toggle_btn_style(color_hex: str, active: bool) -> str:
    """Renkli toggle buton icin aktif/pasif stil olusturur."""
    if active:
        return f"""
            QPushButton {{
                background-color: {color_hex};
                color: white;
                border: 2px solid {color_hex};
                padding: 4px 10px;
                border-radius: 4px;
                font-size: 11px;
                font-weight: bold;
            }}
            QPushButton:hover {{ background-color: {color_hex}; }}
        """
    else:
        return f"""
            QPushButton {{
                background-color: #ffffff;
                color: #6b7280;
                border: 2px solid #d1d5db;
                padding: 4px 10px;
                border-radius: 4px;
                font-size: 11px;
                font-weight: bold;
            }}
            QPushButton:hover {{ border-color: {color_hex}; color: {color_hex}; }}
        """


TAB_STYLE = """
    QTabWidget {
        background-color: #ffffff;
    }
    QTabWidget::pane {
        border: 1px solid #d1d5db;
        background-color: #ffffff;
        top: -1px;
    }
    QTabWidget::tab-bar {
        alignment: left;
    }
    QTabBar {
        background-color: #ffffff;
    }
    QTabBar::tab {
        background-color: #e5e7eb;
        color: #1f2937;
        padding: 12px 28px;
        margin-right: 4px;
        border: 1px solid #d1d5db;
        border-bottom: none;
        border-top-left-radius: 8px;
        border-top-right-radius: 8px;
        font-size: 14px;
        font-weight: bold;
        min-width: 150px;
        min-height: 20px;
    }
    QTabBar::tab:selected {
        background-color: #ffffff;
        color: #2563eb;
        border-bottom: 2px solid #ffffff;
    }
    QTabBar::tab:!selected {
        margin-top: 3px;
    }
    QTabBar::tab:hover:!selected {
        background-color: #d1d5db;
        color: #111827;
    }
"""

PLACEHOLDER_STYLE = """
    QLabel {
        color: #9ca3af;
        font-size: 18px;
        font-weight: bold;
    }
"""


# ================== SUPABASE CLIENT ==================
def _request_with_retry(method, url, max_retries=3, **kwargs):
    """HTTP istegini retry ile gonder. 5xx ve timeout hatalarinda tekrar dene."""
    import time as _time
    last_error = None
    for attempt in range(max_retries):
        try:
            response = method(url, **kwargs)
            if response.status_code >= 500 and attempt < max_retries - 1:
                _time.sleep(2 ** attempt)
                continue
            return response
        except (requests.exceptions.ConnectionError,
                requests.exceptions.Timeout) as e:
            last_error = e
            if attempt < max_retries - 1:
                _time.sleep(2 ** attempt)
    raise last_error


class SupabaseClient:
    """Supabase REST API wrapper (requests ile, ek paket gerektirmez)"""

    def __init__(self, url: str, anon_key: str):
        self.base_url = url.rstrip('/')
        self.rest_url = f"{self.base_url}/rest/v1"
        self.headers = {
            'apikey': anon_key,
            'Authorization': f'Bearer {anon_key}',
            'Content-Type': 'application/json',
        }

    def get_max_evrakno_sira(self) -> int:
        url = f"{self.rest_url}/satis_faturasi"
        params = {
            'select': 'evrakno_sira',
            'order': 'evrakno_sira.desc',
            'limit': '1'
        }
        response = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=15)
        response.raise_for_status()
        data = response.json()
        return data[0]['evrakno_sira'] if data else 0

    def upsert_batch(self, records: list):
        url = f"{self.rest_url}/satis_faturasi"
        params = {'on_conflict': 'evrakno_seri,evrakno_sira,stok_kod,satirno'}
        headers = {**self.headers, 'Prefer': 'resolution=merge-duplicates'}
        response = requests.post(url, headers=headers, params=params,
                                 json=records, timeout=30)
        response.raise_for_status()

    def upsert_adres_batch(self, records: list):
        url = f"{self.rest_url}/satis_faturasi_adres"
        params = {'on_conflict': 'cari_kod'}
        headers = {**self.headers, 'Prefer': 'resolution=merge-duplicates'}
        response = requests.post(url, headers=headers, params=params,
                                 json=records, timeout=30)
        response.raise_for_status()

    def get_adres_by_cari_kod(self, cari_kod: str) -> dict:
        url = f"{self.rest_url}/satis_faturasi_adres"
        params = {'cari_kod': f'eq.{cari_kod}', 'limit': '1'}
        response = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=15)
        response.raise_for_status()
        data = response.json()
        return data[0] if data else None

    def get_all_invoices(self, limit=2000, min_tarih=None) -> list:
        """min_tarih: YYYY-MM-DD formatinda (orn: '2026-02-25')"""
        url = f"{self.rest_url}/satis_faturasi"
        all_data = []
        page_size = 500
        offset = 0
        select_cols = 'id,evrakno_seri,evrakno_sira,satirno,tarih,stok_kod,miktar,cikis_depo_no,paket_sayisi,cari_kodu,cari_adi,product_desc,plasiyer_kodu,malzeme_adi,evrak_adi,satinalma_kalem_id'
        while offset < limit:
            current_limit = min(page_size, limit - offset)
            params = {
                'select': select_cols,
                'order': 'evrakno_sira.desc',
                'limit': str(current_limit),
                'offset': str(offset)
            }
            if min_tarih:
                params['tarih'] = f'gte.{min_tarih}'
            response = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=60)
            response.raise_for_status()
            batch = response.json()
            all_data.extend(batch)
            if len(batch) < current_limit:
                break
            offset += current_limit
        return all_data

    def get_readings_by_invoices(self, fatura_no_list: list) -> list:
        """satis_faturasi_okumalari tablosundan okuma kayitlarini al"""
        if not fatura_no_list:
            return []
        all_readings = []
        batch_size = 50
        for i in range(0, len(fatura_no_list), batch_size):
            batch = fatura_no_list[i:i + batch_size]
            values = ','.join(str(v) for v in batch)
            url = f"{self.rest_url}/satis_faturasi_okumalari"
            params = {
                'select': 'fatura_no,kalem_id,stok_kod,paket_sira,qr_kod,kullanici,created_at',
                'fatura_no': f'in.({values})',
                'limit': '10000'
            }
            response = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=60)
            response.raise_for_status()
            all_readings.extend(response.json())
        return all_readings

    def delete_by_evrakno_sira_list(self, sira_list: list):
        """Belirtilen evrakno_sira degerleri icin kayitlari Supabase'den sil"""
        if not sira_list:
            return
        batch_size = 50
        for i in range(0, len(sira_list), batch_size):
            batch = sira_list[i:i + batch_size]
            values = ','.join(str(v) for v in batch)
            url = f"{self.rest_url}/satis_faturasi"
            params = {'evrakno_sira': f'in.({values})'}
            response = requests.delete(url, headers=self.headers, params=params, timeout=30)
            response.raise_for_status()

    def get_all_evrakno_sira(self, min_tarih: str = None) -> set:
        """Supabase'deki benzersiz evrakno_sira degerlerini al.
        min_tarih: 'YYYY-MM-DD' formatinda (verilirse sadece o tarihten sonrakileri getirir)"""
        url = f"{self.rest_url}/satis_faturasi"
        params = {
            'select': 'evrakno_sira',
            'limit': '10000'
        }
        if min_tarih:
            params['tarih'] = f'gte.{min_tarih}'
        response = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=30)
        response.raise_for_status()
        return set(row['evrakno_sira'] for row in response.json() if row.get('evrakno_sira'))

    def get_fatura_no_with_readings(self) -> set:
        """satis_faturasi_okumalari tablosundaki benzersiz fatura_no degerlerini al (okumasi olan faturalar)"""
        url = f"{self.rest_url}/satis_faturasi_okumalari"
        params = {
            'select': 'fatura_no',
            'limit': '10000'
        }
        response = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=30)
        response.raise_for_status()
        return set(row['fatura_no'] for row in response.json() if row.get('fatura_no'))

    # ---------- Cikis Fisi ----------
    def upsert_cikis_batch(self, records: list):
        url = f"{self.rest_url}/cikis_fisi"
        params = {'on_conflict': 'evrakno_seri,evrakno_sira,stok_kod'}
        headers = {**self.headers, 'Prefer': 'resolution=merge-duplicates'}
        response = requests.post(url, headers=headers, params=params,
                                 json=records, timeout=30)
        response.raise_for_status()

    def get_all_cikis_fisleri(self, limit=2000, min_tarih=None) -> list:
        """cikis_fisi tablosundan kayitlari al. min_tarih: YYYY-MM-DD formatinda"""
        url = f"{self.rest_url}/cikis_fisi"
        all_data = []
        page_size = 500
        offset = 0
        select_cols = 'id,evrakno_seri,evrakno_sira,tarih,stok_kod,miktar,depo,paket_sayisi,malzeme_adi,evrak_adi,satinalma_kalem_id'
        while offset < limit:
            current_limit = min(page_size, limit - offset)
            params = {
                'select': select_cols,
                'order': 'evrakno_sira.desc',
                'limit': str(current_limit),
                'offset': str(offset)
            }
            if min_tarih:
                params['tarih'] = f'gte.{min_tarih}'
            response = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=60)
            response.raise_for_status()
            batch = response.json()
            all_data.extend(batch)
            if len(batch) < current_limit:
                break
            offset += current_limit
        return all_data

    def get_cikis_readings_by_fis_no(self, fis_no_list: list) -> list:
        """cikis_fisi_okumalari tablosundan okuma kayitlarini al"""
        if not fis_no_list:
            return []
        all_readings = []
        batch_size = 50
        for i in range(0, len(fis_no_list), batch_size):
            batch = fis_no_list[i:i + batch_size]
            values = ','.join(str(v) for v in batch)
            url = f"{self.rest_url}/cikis_fisi_okumalari"
            params = {
                'select': 'fis_no,kalem_id,stok_kod,paket_sira,qr_kod,kullanici,created_at',
                'fis_no': f'in.({values})',
                'limit': '10000'
            }
            response = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=60)
            response.raise_for_status()
            all_readings.extend(response.json())
        return all_readings

    def get_cikis_fis_no_with_readings(self) -> set:
        """cikis_fisi_okumalari tablosundaki benzersiz fis_no degerlerini al"""
        url = f"{self.rest_url}/cikis_fisi_okumalari"
        params = {
            'select': 'fis_no',
            'limit': '10000'
        }
        response = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=30)
        response.raise_for_status()
        return set(row['fis_no'] for row in response.json() if row.get('fis_no'))

    def get_cikis_all_evrakno_sira(self, min_tarih: str = None) -> set:
        """cikis_fisi'ndeki benzersiz evrakno_sira degerlerini al"""
        url = f"{self.rest_url}/cikis_fisi"
        params = {
            'select': 'evrakno_sira',
            'limit': '10000'
        }
        if min_tarih:
            params['tarih'] = f'gte.{min_tarih}'
        response = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=30)
        response.raise_for_status()
        return set(row['evrakno_sira'] for row in response.json() if row.get('evrakno_sira'))

    def delete_cikis_by_evrakno_sira_list(self, sira_list: list):
        """cikis_fisi'nden evrakno_sira'ya gore sil — id'leri cekip cascade ile siler"""
        if not sira_list:
            return
        batch_size = 50
        all_ids = []
        for i in range(0, len(sira_list), batch_size):
            batch = sira_list[i:i + batch_size]
            values = ','.join(str(v) for v in batch)
            resp = _request_with_retry(requests.get, f"{self.rest_url}/cikis_fisi",
                                       headers=self.headers,
                                       params={'select': 'id', 'evrakno_sira': f'in.({values})'}, timeout=30)
            resp.raise_for_status()
            all_ids.extend(row['id'] for row in resp.json() if row.get('id') is not None)
        self.delete_cikis_by_id_list(all_ids)

    def delete_cikis_by_id_list(self, id_list: list):
        """cikis_fisi'nden belirtilen id degerleri icin kayitlari sil (okumalari ile birlikte)"""
        if not id_list:
            return
        batch_size = 50
        ids_int = [int(float(v)) for v in id_list]
        for i in range(0, len(ids_int), batch_size):
            batch = ids_int[i:i + batch_size]
            values = ','.join(str(v) for v in batch)
            # Önce okumalari sil (FK: fis_no -> cikis_fisi.id)
            url_okuma = f"{self.rest_url}/cikis_fisi_okumalari"
            resp = requests.delete(url_okuma, headers=self.headers,
                                   params={'fis_no': f'in.({values})'}, timeout=30)
            if resp.status_code not in (200, 204):
                resp.raise_for_status()
            # Sonra ana tabloyu sil
            url = f"{self.rest_url}/cikis_fisi"
            response = requests.delete(url, headers=self.headers,
                                       params={'id': f'in.({values})'}, timeout=30)
            response.raise_for_status()

    # ---------- Giris Fisi ----------
    def upsert_giris_batch(self, records: list):
        url = f"{self.rest_url}/giris_fisi"
        params = {'on_conflict': 'evrakno_seri,evrakno_sira,stok_kod'}
        headers = {**self.headers, 'Prefer': 'resolution=merge-duplicates'}
        response = requests.post(url, headers=headers, params=params,
                                 json=records, timeout=30)
        response.raise_for_status()

    def get_all_giris_fisleri(self, limit=2000, min_tarih=None) -> list:
        """giris_fisi tablosundan kayitlari al. min_tarih: YYYY-MM-DD formatinda"""
        url = f"{self.rest_url}/giris_fisi"
        all_data = []
        page_size = 500
        offset = 0
        select_cols = 'id,evrakno_seri,evrakno_sira,tarih,stok_kod,miktar,depo,paket_sayisi,malzeme_adi,evrak_adi,satinalma_kalem_id'
        while offset < limit:
            current_limit = min(page_size, limit - offset)
            params = {
                'select': select_cols,
                'order': 'evrakno_sira.desc',
                'limit': str(current_limit),
                'offset': str(offset)
            }
            if min_tarih:
                params['tarih'] = f'gte.{min_tarih}'
            response = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=60)
            response.raise_for_status()
            batch = response.json()
            all_data.extend(batch)
            if len(batch) < current_limit:
                break
            offset += current_limit
        return all_data

    def get_giris_readings_by_fis_no(self, fis_no_list: list) -> list:
        """giris_fisi_okumalari tablosundan okuma kayitlarini al"""
        if not fis_no_list:
            return []
        all_readings = []
        batch_size = 50
        for i in range(0, len(fis_no_list), batch_size):
            batch = fis_no_list[i:i + batch_size]
            values = ','.join(str(v) for v in batch)
            url = f"{self.rest_url}/giris_fisi_okumalari"
            params = {
                'select': 'fis_no,kalem_id,stok_kod,paket_sira,qr_kod,kullanici,created_at',
                'fis_no': f'in.({values})',
                'limit': '10000'
            }
            response = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=60)
            response.raise_for_status()
            all_readings.extend(response.json())
        return all_readings

    def get_giris_fis_no_with_readings(self) -> set:
        """giris_fisi_okumalari tablosundaki benzersiz fis_no degerlerini al"""
        url = f"{self.rest_url}/giris_fisi_okumalari"
        params = {
            'select': 'fis_no',
            'limit': '10000'
        }
        response = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=30)
        response.raise_for_status()
        return set(row['fis_no'] for row in response.json() if row.get('fis_no'))

    def get_giris_all_evrakno_sira(self, min_tarih: str = None) -> set:
        """giris_fisi'ndeki benzersiz evrakno_sira degerlerini al"""
        url = f"{self.rest_url}/giris_fisi"
        params = {
            'select': 'evrakno_sira',
            'limit': '10000'
        }
        if min_tarih:
            params['tarih'] = f'gte.{min_tarih}'
        response = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=30)
        response.raise_for_status()
        return set(row['evrakno_sira'] for row in response.json() if row.get('evrakno_sira'))

    def delete_giris_by_evrakno_sira_list(self, sira_list: list):
        """giris_fisi'nden evrakno_sira'ya gore sil — id'leri cekip cascade ile siler"""
        if not sira_list:
            return
        batch_size = 50
        all_ids = []
        for i in range(0, len(sira_list), batch_size):
            batch = sira_list[i:i + batch_size]
            values = ','.join(str(v) for v in batch)
            resp = _request_with_retry(requests.get, f"{self.rest_url}/giris_fisi",
                                       headers=self.headers,
                                       params={'select': 'id', 'evrakno_sira': f'in.({values})'}, timeout=30)
            resp.raise_for_status()
            all_ids.extend(row['id'] for row in resp.json() if row.get('id') is not None)
        self.delete_giris_by_id_list(all_ids)

    def delete_giris_by_id_list(self, id_list: list):
        """giris_fisi'nden belirtilen id degerleri icin kayitlari sil (okumalari ile birlikte)"""
        if not id_list:
            return
        batch_size = 50
        ids_int = [int(float(v)) for v in id_list]
        for i in range(0, len(ids_int), batch_size):
            batch = ids_int[i:i + batch_size]
            values = ','.join(str(v) for v in batch)
            # Önce okumalari sil (FK: fis_no -> giris_fisi.id)
            url_okuma = f"{self.rest_url}/giris_fisi_okumalari"
            resp = requests.delete(url_okuma, headers=self.headers,
                                   params={'fis_no': f'in.({values})'}, timeout=30)
            if resp.status_code not in (200, 204):
                resp.raise_for_status()
            # Sonra ana tabloyu sil
            url = f"{self.rest_url}/giris_fisi"
            response = requests.delete(url, headers=self.headers,
                                       params={'id': f'in.({values})'}, timeout=30)
            response.raise_for_status()

    # ---------- Sevk Fisi ----------
    def upsert_sevk_batch(self, records: list):
        url = f"{self.rest_url}/sevk_fisi"
        params = {'on_conflict': 'evrakno_seri,evrakno_sira,stok_kod'}
        headers = {**self.headers, 'Prefer': 'resolution=merge-duplicates'}
        response = requests.post(url, headers=headers, params=params,
                                 json=records, timeout=30)
        response.raise_for_status()

    def get_all_sevk_fisleri(self, limit=2000, min_tarih=None) -> list:
        """sevk_fisi tablosundan kayitlari al. min_tarih: YYYY-MM-DD formatinda"""
        url = f"{self.rest_url}/sevk_fisi"
        all_data = []
        page_size = 500
        offset = 0
        select_cols = 'id,evrakno_seri,evrakno_sira,tarih,stok_kod,miktar,cikis_depo,giris_depo,paket_sayisi,malzeme_adi,evrak_adi,satinalma_kalem_id'
        while offset < limit:
            current_limit = min(page_size, limit - offset)
            params = {
                'select': select_cols,
                'order': 'evrakno_sira.desc',
                'limit': str(current_limit),
                'offset': str(offset)
            }
            if min_tarih:
                params['tarih'] = f'gte.{min_tarih}'
            response = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=60)
            response.raise_for_status()
            batch = response.json()
            all_data.extend(batch)
            if len(batch) < current_limit:
                break
            offset += current_limit
        return all_data

    def get_sevk_readings_by_fis_no(self, fis_no_list: list) -> list:
        """sevk_fisi_okumalari tablosundan okuma kayitlarini al"""
        if not fis_no_list:
            return []
        all_readings = []
        batch_size = 50
        for i in range(0, len(fis_no_list), batch_size):
            batch = fis_no_list[i:i + batch_size]
            values = ','.join(str(v) for v in batch)
            url = f"{self.rest_url}/sevk_fisi_okumalari"
            params = {
                'select': 'fis_no,kalem_id,stok_kod,paket_sira,qr_kod,kullanici,created_at',
                'fis_no': f'in.({values})',
                'limit': '10000'
            }
            response = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=60)
            response.raise_for_status()
            all_readings.extend(response.json())
        return all_readings

    def get_sevk_fis_no_with_readings(self) -> set:
        """sevk_fisi_okumalari tablosundaki benzersiz fis_no degerlerini al"""
        url = f"{self.rest_url}/sevk_fisi_okumalari"
        params = {
            'select': 'fis_no',
            'limit': '10000'
        }
        response = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=30)
        response.raise_for_status()
        return set(row['fis_no'] for row in response.json() if row.get('fis_no'))

    def get_sevk_all_evrakno_sira(self, min_tarih: str = None) -> set:
        """sevk_fisi'ndeki benzersiz evrakno_sira degerlerini al"""
        url = f"{self.rest_url}/sevk_fisi"
        params = {
            'select': 'evrakno_sira',
            'limit': '10000'
        }
        if min_tarih:
            params['tarih'] = f'gte.{min_tarih}'
        response = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=30)
        response.raise_for_status()
        return set(row['evrakno_sira'] for row in response.json() if row.get('evrakno_sira'))

    def delete_sevk_by_evrakno_sira_list(self, sira_list: list):
        """sevk_fisi'nden evrakno_sira'ya gore sil — id'leri cekip cascade ile siler"""
        if not sira_list:
            return
        batch_size = 50
        all_ids = []
        for i in range(0, len(sira_list), batch_size):
            batch = sira_list[i:i + batch_size]
            values = ','.join(str(v) for v in batch)
            resp = _request_with_retry(requests.get, f"{self.rest_url}/sevk_fisi",
                                       headers=self.headers,
                                       params={'select': 'id', 'evrakno_sira': f'in.({values})'}, timeout=30)
            resp.raise_for_status()
            all_ids.extend(row['id'] for row in resp.json() if row.get('id') is not None)
        self.delete_sevk_by_id_list(all_ids)

    def delete_sevk_by_id_list(self, id_list: list):
        """sevk_fisi'nden belirtilen id degerleri icin kayitlari sil (okumalari ile birlikte)"""
        if not id_list:
            return
        batch_size = 50
        ids_int = [int(float(v)) for v in id_list]
        for i in range(0, len(ids_int), batch_size):
            batch = ids_int[i:i + batch_size]
            values = ','.join(str(v) for v in batch)
            # Önce okumalari sil (FK: fis_no -> sevk_fisi.id)
            url_okuma = f"{self.rest_url}/sevk_fisi_okumalari"
            resp = requests.delete(url_okuma, headers=self.headers,
                                   params={'fis_no': f'in.({values})'}, timeout=30)
            if resp.status_code not in (200, 204):
                resp.raise_for_status()
            # Sonra ana tabloyu sil
            url = f"{self.rest_url}/sevk_fisi"
            response = requests.delete(url, headers=self.headers,
                                       params={'id': f'in.({values})'}, timeout=30)
            response.raise_for_status()

    # ---------- Nakliye ----------
    def get_all_nakliye_fisleri(self, limit=2000, min_tarih=None) -> list:
        """nakliye_fisleri tablosundan kayitlari al. min_tarih: YYYYMMDD formatinda (orn: '20260226')"""
        url = f"{self.rest_url}/nakliye_fisleri"
        all_data = []
        page_size = 500
        offset = 0
        select_cols = 'id,oturum_id,nakliye_no,plaka,belge_tarihi,fatura_numarasi,malzeme_no,satinalma_kalem_id,miktar,depo_yeri,paket_sayisi,malzeme_adi'
        while offset < limit:
            current_limit = min(page_size, limit - offset)
            params = {
                'select': select_cols,
                'order': 'created_at.desc',
                'limit': str(current_limit),
                'offset': str(offset)
            }
            if min_tarih:
                params['belge_tarihi'] = f'gte.{min_tarih}'
            response = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=60)
            response.raise_for_status()
            batch = response.json()
            all_data.extend(batch)
            if len(batch) < current_limit:
                break
            offset += current_limit
        return all_data

    def get_nakliye_readings_by_kalem_ids(self, kalem_id_list: list) -> list:
        """nakliye_fisleri_okumalari tablosundan okuma kayitlarini al"""
        if not kalem_id_list:
            return []
        all_readings = []
        batch_size = 50
        for i in range(0, len(kalem_id_list), batch_size):
            batch = kalem_id_list[i:i + batch_size]
            values = ','.join(str(v) for v in batch)
            url = f"{self.rest_url}/nakliye_fisleri_okumalari"
            params = {
                'select': 'nakliye_kalem_id,paket_sira,qr_kod,okuyan_kullanici,okuma_zamani',
                'nakliye_kalem_id': f'in.({values})',
                'limit': '10000'
            }
            response = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=60)
            response.raise_for_status()
            all_readings.extend(response.json())
        return all_readings

    def delete_nakliye_by_id_list(self, id_list: list):
        """nakliye_fisleri'nden belirtilen id degerleri icin kayitlari sil (okumalari ile birlikte)"""
        if not id_list:
            return
        batch_size = 50
        ids_int = [int(float(v)) for v in id_list]
        for i in range(0, len(ids_int), batch_size):
            batch = ids_int[i:i + batch_size]
            values = ','.join(str(v) for v in batch)
            # Önce okumalari sil (FK: nakliye_kalem_id -> nakliye_fisleri.id)
            url_okuma = f"{self.rest_url}/nakliye_fisleri_okumalari"
            resp = requests.delete(url_okuma, headers=self.headers,
                                   params={'nakliye_kalem_id': f'in.({values})'}, timeout=30)
            if resp.status_code not in (200, 204):
                resp.raise_for_status()
            # Sonra ana tabloyu sil
            url = f"{self.rest_url}/nakliye_fisleri"
            response = requests.delete(url, headers=self.headers,
                                       params={'id': f'in.({values})'}, timeout=30)
            response.raise_for_status()

    def delete_sayim_oturum_by_id_list(self, oturum_id_list: list):
        """sayim_oturumlari'ndan belirtilen id degerleri icin kayitlari sil"""
        if not oturum_id_list:
            return
        batch_size = 50
        for i in range(0, len(oturum_id_list), batch_size):
            batch = oturum_id_list[i:i + batch_size]
            values = ','.join(str(v) for v in batch)
            url = f"{self.rest_url}/sayim_oturumlari"
            params = {'id': f'in.({values})'}
            response = requests.delete(url, headers=self.headers, params=params, timeout=30)
            response.raise_for_status()

    # ---------- Sayim ----------
    def get_all_sayim_oturumlari(self, limit=2000, min_tarih=None, lokasyon=None) -> list:
        """sayim_oturumlari tablosundan oturumlari al. lokasyon: DEPO/EXC/SUBE"""
        url = f"{self.rest_url}/sayim_oturumlari"
        all_data = []
        page_size = 500
        offset = 0
        select_cols = 'id,lokasyon,lokasyon_kodu,kullanici,durum,sayim_kodu,baslangic,bitis,toplam_cesit,toplam_adet'
        while offset < limit:
            current_limit = min(page_size, limit - offset)
            params = {
                'select': select_cols,
                'order': 'baslangic.desc',
                'limit': str(current_limit),
                'offset': str(offset)
            }
            if min_tarih:
                params['baslangic'] = f'gte.{min_tarih}'
            if lokasyon:
                params['lokasyon'] = f'eq.{lokasyon}'
            response = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=60)
            if not response.ok:
                try:
                    detail = response.json()
                except Exception:
                    detail = response.text
                raise Exception(f"Supabase {response.status_code} hatası: {detail}")
            batch = response.json()
            all_data.extend(batch)
            if len(batch) < current_limit:
                break
            offset += current_limit
        return all_data

    def get_sayim_okumalari_by_oturum_ids(self, oturum_id_list: list) -> list:
        """sayim_okumalari tablosundan okuma kayitlarini al"""
        if not oturum_id_list:
            return []
        all_readings = []
        batch_size = 50
        for i in range(0, len(oturum_id_list), batch_size):
            batch = oturum_id_list[i:i + batch_size]
            values = ','.join(str(v) for v in batch)
            url = f"{self.rest_url}/sayim_okumalari"
            params = {
                'select': 'oturum_id,stok_kod,malzeme_adi,paket_sira,paket_toplam,manuel,adet,kullanici,created_at',
                'oturum_id': f'in.({values})',
                'limit': '10000'
            }
            response = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=60)
            response.raise_for_status()
            all_readings.extend(response.json())
        return all_readings

    def get_qr_log_by_stok_kod(self, stok_kod):
        """Bir urunun tum okuma gecmisini 6 tablodan toplar"""
        results = []
        # Okuma tablolarinda stok_kod 10 haneli base kod (orn: '3200398839')
        base_kod = _normalize_stok_kod(stok_kod)
        stok_kod_18 = base_kod.zfill(18)  # nakliye icin 18 hane

        # satis_faturasi'ndan kalem_id -> cari_adi eslesmesi
        cari_map = {}  # kalem_id -> cari_adi
        try:
            url = f"{self.rest_url}/satis_faturasi"
            params = {
                'select': 'id,cari_adi',
                'stok_kod': f'like.{base_kod}*',
                'limit': '5000'
            }
            resp = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=30)
            resp.raise_for_status()
            for row in resp.json():
                cari_map[str(row.get('id', ''))] = row.get('cari_adi', '') or ''
        except Exception as e:
            logger.warning(f"QR log cari_adi hatasi: {e}")

        # 5 standart tablo (stok_kod alani var)
        # yon: + = giris, - = cikis
        tables = [
            ('satis_faturasi_okumalari', u'Sat\u0131\u015f', 'fatura_no', '-',
             'stok_kod,fatura_no,kalem_id,paket_sira,qr_kod,kullanici,created_at'),
            ('cikis_fisi_okumalari', u'\u00c7\u0131k\u0131\u015f', 'fis_no', '-',
             'stok_kod,fis_no,kalem_id,paket_sira,qr_kod,kullanici,created_at'),
            ('giris_fisi_okumalari', u'Giri\u015f', 'fis_no', '+',
             'stok_kod,fis_no,kalem_id,paket_sira,qr_kod,kullanici,created_at'),
            ('sevk_fisi_okumalari', 'Sevk', 'fis_no', '+',
             'stok_kod,fis_no,kalem_id,paket_sira,qr_kod,kullanici,created_at'),
            ('sayim_okumalari', u'Say\u0131m', 'oturum_id', '+',
             'stok_kod,oturum_id,paket_sira,qr_kod,kullanici,created_at'),
        ]

        for table_name, kaynak, evrak_field, yon, select_cols in tables:
            try:
                url = f"{self.rest_url}/{table_name}"
                params = {
                    'select': select_cols,
                    'stok_kod': f'like.{base_kod}*',
                    'order': 'created_at.asc',
                    'limit': '5000'
                }
                resp = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=60)
                resp.raise_for_status()
                for row in resp.json():
                    qr = str(row.get('qr_kod', '') or '')
                    kalem_id = str(row.get('kalem_id', '') or '')
                    results.append({
                        'tarih': row.get('created_at', ''),
                        'yon': yon,
                        'kaynak': kaynak,
                        'evrak_no': str(row.get(evrak_field, '') or ''),
                        'cari_adi': cari_map.get(kalem_id, ''),
                        'paket_sira': row.get('paket_sira', ''),
                        'qr_kod': qr,
                        'qr_kod_kisa': (qr[:20] + '...' + qr[-20:]) if len(qr) > 45 else qr,
                        'kullanici': row.get('kullanici', ''),
                    })
            except Exception as e:
                logger.warning(f"QR log {table_name} hatasi: {e}")

        # Nakliye (farkli alan adlari, giris +)
        try:
            url = f"{self.rest_url}/nakliye_fisleri_okumalari"
            params = {
                'select': 'malzeme_no_qr,nakliye_kalem_id,paket_sira,qr_kod,okuyan_kullanici,okuma_zamani',
                'malzeme_no_qr': f'eq.{stok_kod_18}',
                'order': 'okuma_zamani.asc',
                'limit': '5000'
            }
            resp = _request_with_retry(requests.get, url, headers=self.headers, params=params, timeout=60)
            resp.raise_for_status()
            for row in resp.json():
                qr = str(row.get('qr_kod', '') or '')
                results.append({
                    'tarih': row.get('okuma_zamani', ''),
                    'yon': '+',
                    'kaynak': 'Nakliye',
                    'evrak_no': str(row.get('nakliye_kalem_id', '') or ''),
                    'cari_adi': '',
                    'paket_sira': row.get('paket_sira', ''),
                    'qr_kod': qr,
                    'qr_kod_kisa': (qr[:20] + '...' + qr[-20:]) if len(qr) > 45 else qr,
                    'kullanici': row.get('okuyan_kullanici', ''),
                })
        except Exception as e:
            logger.warning(f"QR log nakliye hatasi: {e}")

        # Tarihe gore sirala (kronolojik)
        results.sort(key=lambda r: r.get('tarih', '') or '')
        return results


# ================== DOGTAS API CLIENT ==================
class DogtasApiClient:
    """Dogtas API client - token cache + paket bilgisi"""

    def __init__(self, config: dict):
        self.base_url = config.get('base_url', '').rstrip('/')
        self.user_name = config.get('userName', '')
        self.password = config.get('password', '')
        self.client_id = config.get('clientId', '')
        self.client_secret = config.get('clientSecret', '')
        self.application_code = config.get('applicationCode', '')
        self.customer_no = config.get('CustomerNo', '')
        self.nakliye_endpoint = config.get('nakliye', '')

        self._token = None
        self._token_expires_at = None

    def _get_token(self) -> str:
        now = datetime.now()
        if self._token and self._token_expires_at and now < self._token_expires_at:
            return self._token

        url = f"{self.base_url}/Authorization/GetAccessToken"
        payload = {
            'userName': self.user_name,
            'password': self.password,
            'clientId': self.client_id,
            'clientSecret': self.client_secret,
            'applicationCode': self.application_code
        }
        response = requests.post(url, json=payload, timeout=15)
        response.raise_for_status()
        data = response.json()

        if data.get('isSuccess') and data.get('data', {}).get('accessToken'):
            self._token = data['data']['accessToken']
            self._token_expires_at = now + timedelta(minutes=DOGTAS_TOKEN_CACHE_MINUTES)
            return self._token

        raise Exception(f"Dogtas token alinamadi: {data}")

    def get_product_packages(self, product_codes: list) -> dict:
        """
        Urun paket bilgilerini al.
        Returns: {product_code: {'productDesc': str, 'paketSayisi': int}}
        """
        if not product_codes:
            return {}

        token = self._get_token()
        url = f"{DOGTAS_API_BASE_URL}/api/SapDealer/GetProductPackages"
        payload = {
            'dealerCode': self.customer_no,
            'productCodes': product_codes
        }
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        response = requests.post(url, json=payload, headers=headers, timeout=30)
        response.raise_for_status()
        data = response.json()

        result = {}
        if data.get('isSuccess') and isinstance(data.get('data'), list):
            # productCode'a gore grupla, benzersiz materialCode ile paket sayisi hesapla
            urun_gruplari = {}
            for item in data['data']:
                pc = item.get('productCode', '')
                if pc not in urun_gruplari:
                    urun_gruplari[pc] = {
                        'productDesc': item.get('productDesc', ''),
                        'paketSayisi': 0,
                        '_material_codes': set()
                    }
                mc = item.get('materialCode', '')
                if mc and mc not in urun_gruplari[pc]['_material_codes']:
                    urun_gruplari[pc]['_material_codes'].add(mc)
                    urun_gruplari[pc]['paketSayisi'] += 1

            for pc, info in urun_gruplari.items():
                result[pc] = {
                    'productDesc': info['productDesc'],
                    'paketSayisi': info['paketSayisi']
                }

        return result

    def get_shipments(self, date_start: str, date_end: str, nakliye_no: str = '') -> list:
        """
        Nakliye fisleri cek.
        date_start, date_end: DD.MM.YYYY formatinda
        Returns: API'den gelen ham item listesi (EAN bos olanlar filtrelenir)
        """
        if not self.nakliye_endpoint:
            raise Exception("Dogtas nakliye endpoint ayarlanmamis (PRGsheet 'nakliye' ayarini kontrol edin)")

        token = self._get_token()
        url = f"{self.base_url}{self.nakliye_endpoint}"
        payload = {
            'deliveryDocument': '',
            'orderer': self.customer_no,
            'transportationNumber': nakliye_no,
            'documentDateStart': date_start,
            'documentDateEnd': date_end
        }
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        response = requests.post(url, json=payload, headers=headers, timeout=60)
        response.raise_for_status()
        data = response.json()

        if data.get('isSuccess') and isinstance(data.get('data'), list):
            # Yari mamulleri filtrele - EAN bos olanlar yari mamuldre
            return [item for item in data['data']
                    if str(item.get('ean', '') or '').strip()]
        raise Exception(
            f"Dogtas GetShipments hatasi: {data.get('message', 'Bilinmeyen hata')}"
        )


# ================== LOAD DATA THREAD ==================
class LoadDataThread(QThread):
    """Supabase'den fatura + okuma verilerini arka planda yukler"""

    data_loaded = pyqtSignal(list, dict)  # (all_data, readings_map)
    error_occurred = pyqtSignal(str)

    def __init__(self, supabase_client: SupabaseClient, min_tarih=None):
        super().__init__()
        self.supabase_client = supabase_client
        self.min_tarih = min_tarih

    def run(self):
        try:
            all_data = self.supabase_client.get_all_invoices(limit=5000, min_tarih=self.min_tarih)

            readings_map = {}
            try:
                fatura_no_list = list(set(
                    row['evrakno_sira'] for row in all_data
                    if row.get('evrakno_sira')
                ))
                readings = self.supabase_client.get_readings_by_invoices(fatura_no_list)

                from datetime import datetime, timedelta, timezone
                for r in readings:
                    kalem_id = r.get('kalem_id')
                    if not kalem_id:
                        continue
                    ps = int(r.get('paket_sira', 0) or 0)
                    qr = str(r.get('qr_kod', ''))

                    if kalem_id not in readings_map:
                        readings_map[kalem_id] = {}

                    read_type = 'manual' if qr.startswith('MANUEL_TOPLU_') else 'scanner'
                    raw_time = r.get('created_at', '')
                    tr_time = ''
                    if raw_time:
                        try:
                            dt = datetime.fromisoformat(raw_time.replace('Z', '+00:00'))
                            dt_tr = dt.astimezone(timezone(timedelta(hours=3)))
                            tr_time = dt_tr.strftime('%d.%m.%Y %H:%M')
                        except Exception:
                            tr_time = str(raw_time)[:16]

                    info = {
                        'type': read_type,
                        'user': r.get('kullanici', '') or '',
                        'time': tr_time,
                    }
                    if ps not in readings_map[kalem_id]:
                        readings_map[kalem_id][ps] = []
                    readings_map[kalem_id][ps].append(info)

                logger.info(f"{len(readings)} okuma kaydi yuklendi")
            except Exception as e:
                logger.warning(f"Okuma verileri yuklenemedi: {e}")

            self.data_loaded.emit(all_data, readings_map)
        except Exception as e:
            self.error_occurred.emit(str(e))


# ================== FABRIKA NAKLIYE PLAN LOAD THREAD ==================
class FabrikaNakliyeLoadDataThread(QThread):
    """PRGsheet Bekleyenler + Dogtas API nakliye verilerini arka planda yukler"""

    data_loaded = pyqtSignal(list, list, dict)  # (rows, column_names, nakliye_map)
    error_occurred = pyqtSignal(str)
    status_update = pyqtSignal(str)

    def __init__(self, dogtas_client, date_start: str, date_end: str):
        super().__init__()
        self.dogtas_client = dogtas_client
        self.date_start = date_start  # DD.MM.YYYY
        self.date_end = date_end      # DD.MM.YYYY

    def run(self):
        try:
            from io import BytesIO
            config_manager = CentralConfigManager()
            sid = config_manager.MASTER_SPREADSHEET_ID

            # 1. PRGsheet Bekleyenler sayfasini cek
            self.status_update.emit("PRGsheet 'Bekleyenler' sayfasi yukleniyor...")
            gsheets_url = f"https://docs.google.com/spreadsheets/d/{sid}/export?format=xlsx"
            resp = requests.get(gsheets_url, timeout=30)
            resp.raise_for_status()
            df = pd.read_excel(BytesIO(resp.content), sheet_name="Bekleyenler")

            # 2. "Durum" = "Sevke Hazir" filtrele
            durum_col = next((c for c in df.columns if str(c).strip() == 'Durum'), None)
            if durum_col is None:
                raise Exception("'Bekleyenler' sayfasinda 'Durum' sutunu bulunamadi")
            df = df[df[durum_col].astype(str).str.strip() == u'Sevke Haz\u0131r']

            column_names = [str(c) for c in df.columns]
            rows = []
            for _, row in df.iterrows():
                row_dict = {}
                for col in df.columns:
                    val = row[col]
                    try:
                        is_na = pd.isna(val)
                    except (TypeError, ValueError):
                        is_na = False
                    row_dict[str(col)] = '' if is_na else str(val)
                rows.append(row_dict)

            logger.info(f"FabrikaNakliye: {len(rows)} 'Sevke Hazir' kaydi yuklendi")
            self.status_update.emit(
                f"{len(rows)} 'Sevke Hazir' kayit bulundu, Dogtas nakliye verileri cekiliyor..."
            )

            # 3. Dogtas API'den nakliye kalemlerini cek
            # nakliye_map: {satinalma_kalem_id: [{'nakliye_no', 'plaka', 'tarih', ...}]}
            nakliye_map = {}
            if self.dogtas_client:
                try:
                    items = self.dogtas_client.get_shipments(
                        self.date_start, self.date_end
                    )
                    for item in items:
                        # satinalma_kalem_id = referenceDocumentNumber + referenceItemNumber
                        ref_no = str(item.get('referenceDocumentNumber', '') or '').strip()
                        ref_item = str(item.get('referenceItemNumber', '') or '').strip()
                        kid = ref_no + ref_item
                        if kid:
                            if kid not in nakliye_map:
                                nakliye_map[kid] = []
                            nakliye_map[kid].append({
                                'nakliye_no': str(item.get('distributionDocumentNumber', '') or ''),
                                'plaka': str(item.get('shipmentVehicleLicensePlate', '') or ''),
                                'sofor': str(item.get('shipmentVehicleDriverName', '') or ''),
                                'tarih': str(item.get('documanetDate', '') or ''),
                                'malzeme_no': str(item.get('materialNumber', '') or ''),
                                'malzeme_adi': str(item.get('materialName', '') or ''),
                                'depo': str(item.get('storageLocation', '') or ''),
                            })
                    logger.info(f"FabrikaNakliye: {len(items)} kalem, {len(nakliye_map)} benzersiz kalem id yuklendi")
                except Exception as e:
                    logger.warning(f"FabrikaNakliye Dogtas verileri yuklenemedi: {e}")
                    self.status_update.emit(f"Dogtas hatasi: {e}")

            self.data_loaded.emit(rows, column_names, nakliye_map)
        except Exception as e:
            self.error_occurred.emit(str(e))


class NakliyeLoadDataThread(QThread):
    """Supabase'den nakliye fisleri + okuma verilerini arka planda yukler"""

    data_loaded = pyqtSignal(list, dict)
    error_occurred = pyqtSignal(str)

    def __init__(self, supabase_client: SupabaseClient, min_tarih=None):
        super().__init__()
        self.supabase_client = supabase_client
        self.min_tarih = min_tarih

    def run(self):
        try:
            all_data = self.supabase_client.get_all_nakliye_fisleri(limit=5000, min_tarih=self.min_tarih)

            readings_map = {}
            try:
                kalem_id_list = list(set(
                    row['id'] for row in all_data
                    if row.get('id')
                ))
                readings = self.supabase_client.get_nakliye_readings_by_kalem_ids(kalem_id_list)

                from datetime import datetime, timedelta, timezone
                for r in readings:
                    kalem_id = r.get('nakliye_kalem_id')
                    ps = int(r.get('paket_sira', 0) or 0)
                    qr = str(r.get('qr_kod', ''))

                    if kalem_id not in readings_map:
                        readings_map[kalem_id] = {}

                    read_type = 'manual' if qr.startswith('MANUEL_TOPLU_') else 'scanner'
                    raw_time = r.get('okuma_zamani', '') or ''
                    tr_time = ''
                    if raw_time:
                        try:
                            dt = datetime.fromisoformat(raw_time.replace('Z', '+00:00'))
                            dt_tr = dt.astimezone(timezone(timedelta(hours=3)))
                            tr_time = dt_tr.strftime('%d.%m.%Y %H:%M')
                        except Exception:
                            tr_time = str(raw_time)[:16]

                    info = {
                        'type': read_type,
                        'user': r.get('okuyan_kullanici', '') or '',
                        'time': tr_time,
                    }
                    if ps not in readings_map[kalem_id]:
                        readings_map[kalem_id][ps] = []
                    readings_map[kalem_id][ps].append(info)

                logger.info(f"Nakliye: {len(readings)} okuma kaydi yuklendi")
            except Exception as e:
                logger.warning(f"Nakliye okuma verileri yuklenemedi: {e}")

            self.data_loaded.emit(all_data, readings_map)
        except Exception as e:
            self.error_occurred.emit(str(e))


# ================== CIKIS LOAD DATA THREAD ==================
class CikisLoadDataThread(QThread):
    """Supabase'den cikis fisi + okuma verilerini arka planda yukler"""

    data_loaded = pyqtSignal(list, dict)  # (all_data, readings_map)
    error_occurred = pyqtSignal(str)

    def __init__(self, supabase_client: SupabaseClient, min_tarih=None):
        super().__init__()
        self.supabase_client = supabase_client
        self.min_tarih = min_tarih

    def run(self):
        try:
            all_data = self.supabase_client.get_all_cikis_fisleri(limit=5000, min_tarih=self.min_tarih)

            readings_map = {}
            try:
                fis_no_list = list(set(
                    row['evrakno_sira'] for row in all_data
                    if row.get('evrakno_sira')
                ))
                readings = self.supabase_client.get_cikis_readings_by_fis_no(fis_no_list)

                from datetime import datetime, timedelta, timezone
                for r in readings:
                    kalem_id = r.get('kalem_id')
                    if not kalem_id:
                        continue
                    ps = int(r.get('paket_sira', 0) or 0)
                    qr = str(r.get('qr_kod', ''))

                    if kalem_id not in readings_map:
                        readings_map[kalem_id] = {}

                    read_type = 'manual' if qr.startswith('MANUEL_TOPLU_') else 'scanner'
                    raw_time = r.get('created_at', '')
                    tr_time = ''
                    if raw_time:
                        try:
                            dt = datetime.fromisoformat(raw_time.replace('Z', '+00:00'))
                            dt_tr = dt.astimezone(timezone(timedelta(hours=3)))
                            tr_time = dt_tr.strftime('%d.%m.%Y %H:%M')
                        except Exception:
                            tr_time = str(raw_time)[:16]

                    info = {
                        'type': read_type,
                        'user': r.get('kullanici', '') or '',
                        'time': tr_time,
                    }
                    if ps not in readings_map[kalem_id]:
                        readings_map[kalem_id][ps] = []
                    readings_map[kalem_id][ps].append(info)

                logger.info(f"Cikis: {len(readings)} okuma kaydi yuklendi")
            except Exception as e:
                logger.warning(f"Cikis okuma verileri yuklenemedi: {e}")

            self.data_loaded.emit(all_data, readings_map)
        except Exception as e:
            self.error_occurred.emit(str(e))


# ================== GIRIS LOAD DATA THREAD ==================
class GirisLoadDataThread(QThread):
    """Supabase'den giris fisi + okuma verilerini arka planda yukler"""

    data_loaded = pyqtSignal(list, dict)  # (all_data, readings_map)
    error_occurred = pyqtSignal(str)

    def __init__(self, supabase_client: SupabaseClient, min_tarih=None):
        super().__init__()
        self.supabase_client = supabase_client
        self.min_tarih = min_tarih

    def run(self):
        try:
            all_data = self.supabase_client.get_all_giris_fisleri(limit=5000, min_tarih=self.min_tarih)

            readings_map = {}
            try:
                fis_no_list = list(set(
                    row['evrakno_sira'] for row in all_data
                    if row.get('evrakno_sira')
                ))
                readings = self.supabase_client.get_giris_readings_by_fis_no(fis_no_list)

                from datetime import datetime, timedelta, timezone
                for r in readings:
                    kalem_id = r.get('kalem_id')
                    if not kalem_id:
                        continue
                    ps = int(r.get('paket_sira', 0) or 0)
                    qr = str(r.get('qr_kod', ''))

                    if kalem_id not in readings_map:
                        readings_map[kalem_id] = {}

                    read_type = 'manual' if qr.startswith('MANUEL_TOPLU_') else 'scanner'
                    raw_time = r.get('created_at', '')
                    tr_time = ''
                    if raw_time:
                        try:
                            dt = datetime.fromisoformat(raw_time.replace('Z', '+00:00'))
                            dt_tr = dt.astimezone(timezone(timedelta(hours=3)))
                            tr_time = dt_tr.strftime('%d.%m.%Y %H:%M')
                        except Exception:
                            tr_time = str(raw_time)[:16]

                    info = {
                        'type': read_type,
                        'user': r.get('kullanici', '') or '',
                        'time': tr_time,
                    }
                    if ps not in readings_map[kalem_id]:
                        readings_map[kalem_id][ps] = []
                    readings_map[kalem_id][ps].append(info)

                logger.info(f"Giris: {len(readings)} okuma kaydi yuklendi")
            except Exception as e:
                logger.warning(f"Giris okuma verileri yuklenemedi: {e}")

            self.data_loaded.emit(all_data, readings_map)
        except Exception as e:
            self.error_occurred.emit(str(e))


# ================== SEVK LOAD DATA THREAD ==================
class SevkLoadDataThread(QThread):
    """Supabase'den sevk fisi + okuma verilerini arka planda yukler"""

    data_loaded = pyqtSignal(list, dict)  # (all_data, readings_map)
    error_occurred = pyqtSignal(str)

    def __init__(self, supabase_client: SupabaseClient, min_tarih=None):
        super().__init__()
        self.supabase_client = supabase_client
        self.min_tarih = min_tarih

    def run(self):
        try:
            all_data = self.supabase_client.get_all_sevk_fisleri(limit=5000, min_tarih=self.min_tarih)

            readings_map = {}
            try:
                fis_no_list = list(set(
                    row['evrakno_sira'] for row in all_data
                    if row.get('evrakno_sira')
                ))
                readings = self.supabase_client.get_sevk_readings_by_fis_no(fis_no_list)

                from datetime import datetime, timedelta, timezone
                for r in readings:
                    kalem_id = r.get('kalem_id')
                    if not kalem_id:
                        continue
                    ps = int(r.get('paket_sira', 0) or 0)
                    qr = str(r.get('qr_kod', ''))

                    if kalem_id not in readings_map:
                        readings_map[kalem_id] = {}

                    read_type = 'manual' if qr.startswith('MANUEL_TOPLU_') else 'scanner'
                    raw_time = r.get('created_at', '')
                    tr_time = ''
                    if raw_time:
                        try:
                            dt = datetime.fromisoformat(raw_time.replace('Z', '+00:00'))
                            dt_tr = dt.astimezone(timezone(timedelta(hours=3)))
                            tr_time = dt_tr.strftime('%d.%m.%Y %H:%M')
                        except Exception:
                            tr_time = str(raw_time)[:16]

                    info = {
                        'type': read_type,
                        'user': r.get('kullanici', '') or '',
                        'time': tr_time,
                    }
                    if ps not in readings_map[kalem_id]:
                        readings_map[kalem_id][ps] = []
                    readings_map[kalem_id][ps].append(info)

                logger.info(f"Sevk: {len(readings)} okuma kaydi yuklendi")
            except Exception as e:
                logger.warning(f"Sevk okuma verileri yuklenemedi: {e}")

            self.data_loaded.emit(all_data, readings_map)
        except Exception as e:
            self.error_occurred.emit(str(e))


# ================== SAYIM LOAD DATA THREAD ==================
class SayimLoadDataThread(QThread):
    """Supabase'den sayim oturumlari + okuma verilerini arka planda yukler.
    Okumalari (oturum_id, stok_kod) bazinda gruplayarak sanal satirlar olusturur."""

    data_loaded = pyqtSignal(list, dict)  # (all_data, readings_map)
    error_occurred = pyqtSignal(str)

    def __init__(self, supabase_client: SupabaseClient, min_tarih=None, lokasyon=None):
        super().__init__()
        self.supabase_client = supabase_client
        self.min_tarih = min_tarih
        self.lokasyon = lokasyon

    def run(self):
        try:
            oturumlar = self.supabase_client.get_all_sayim_oturumlari(
                limit=5000, min_tarih=self.min_tarih, lokasyon=self.lokasyon
            )
            oturum_map = {o['id']: o for o in oturumlar}

            all_data = []
            readings_map = {}

            try:
                oturum_ids = list(oturum_map.keys())
                okumalar = self.supabase_client.get_sayim_okumalari_by_oturum_ids(oturum_ids)

                from datetime import datetime, timedelta, timezone

                # Okumalari (oturum_id, stok_kod) bazinda grupla
                grouped = {}  # key: (oturum_id, stok_kod) -> list of readings
                for r in okumalar:
                    oid = r.get('oturum_id')
                    sk = r.get('stok_kod', '')
                    if not oid or not sk:
                        continue
                    key = (oid, sk)
                    if key not in grouped:
                        grouped[key] = []
                    grouped[key].append(r)

                for (oid, sk), readings in grouped.items():
                    oturum = oturum_map.get(oid, {})
                    composite_key = f"{oid}::{sk}"

                    # paket_toplam ve malzeme_adi belirle
                    paket_toplam = 1
                    malzeme_adi = sk
                    for r in readings:
                        pt = int(r.get('paket_toplam', 0) or 0)
                        if pt > paket_toplam:
                            paket_toplam = pt
                        if r.get('malzeme_adi'):
                            malzeme_adi = r['malzeme_adi']

                    # readings_map olustur
                    paket_readings = {}
                    manuel_infos = []  # paket_sira olmayan manuel okumalar
                    for r in readings:
                        ps = int(r.get('paket_sira', 0) or 0)
                        is_manuel = r.get('manuel', False)
                        read_type = 'manual' if is_manuel else 'scanner'
                        raw_time = r.get('created_at', '')
                        tr_time = ''
                        if raw_time:
                            try:
                                dt = datetime.fromisoformat(raw_time.replace('Z', '+00:00'))
                                dt_tr = dt.astimezone(timezone(timedelta(hours=3)))
                                tr_time = dt_tr.strftime('%d.%m.%Y %H:%M')
                            except Exception:
                                tr_time = str(raw_time)[:16]
                        info = {
                            'type': read_type,
                            'user': r.get('kullanici', '') or '',
                            'time': tr_time,
                        }
                        if ps > 0:
                            if ps not in paket_readings:
                                paket_readings[ps] = []
                            paket_readings[ps].append(info)
                        elif is_manuel:
                            # paket_sira olmayan manuel okuma: adet kadar birim
                            adet = int(r.get('adet', 1) or 1)
                            for _ in range(adet):
                                manuel_infos.append(info)

                    # Manuel okumalari paket_readings'e ekle
                    # Her manuel birim icin tum paket pozisyonlarini doldur
                    if manuel_infos:
                        if paket_toplam < 1:
                            paket_toplam = 1
                        qr_miktar = max((len(v) for v in paket_readings.values()), default=0)
                        for mi in manuel_infos:
                            for ps in range(1, paket_toplam + 1):
                                if ps not in paket_readings:
                                    paket_readings[ps] = []
                                # QR miktar ile hizala (bos pozisyonlari None ile doldur)
                                while len(paket_readings[ps]) < qr_miktar:
                                    paket_readings[ps].append(None)
                                paket_readings[ps].append(mi)

                    # miktar = en cok okunan paket_sira'nin sayisi
                    miktar = max((len(v) for v in paket_readings.values()), default=0)

                    # Tarih formatla
                    raw_tarih = oturum.get('baslangic', '')
                    tarih_str = ''
                    if raw_tarih:
                        try:
                            dt = datetime.fromisoformat(raw_tarih.replace('Z', '+00:00'))
                            dt_tr = dt.astimezone(timezone(timedelta(hours=3)))
                            tarih_str = dt_tr.strftime('%d.%m.%Y %H:%M')
                        except Exception:
                            tarih_str = str(raw_tarih)[:16]

                    row = {
                        'composite_key': composite_key,
                        'sayim_kodu': oturum.get('sayim_kodu', ''),
                        'lokasyon': oturum.get('lokasyon', ''),
                        'lokasyon_kodu': str(oturum.get('lokasyon_kodu', '') or ''),
                        'durum': oturum.get('durum', ''),
                        'tarih': tarih_str,
                        'stok_kod': sk,
                        'malzeme_adi': malzeme_adi,
                        'miktar': miktar,
                        'paket_sayisi': paket_toplam,
                        'kullanici': oturum.get('kullanici', ''),
                    }
                    all_data.append(row)
                    readings_map[composite_key] = paket_readings

                logger.info(f"Sayim ({self.lokasyon}): {len(okumalar)} okuma, {len(all_data)} urun satiri yuklendi")
            except Exception as e:
                logger.warning(f"Sayim okuma verileri yuklenemedi: {e}")

            # PRGsheet Stok verisinden beklenen adetleri cek
            beklenen_map = {}  # stok_kod -> {'malzeme_adi': str, 'beklenen': float}
            try:
                from io import BytesIO
                config_manager = CentralConfigManager()
                sid = config_manager.MASTER_SPREADSHEET_ID
                gsheets_url = f"https://docs.google.com/spreadsheets/d/{sid}/export?format=xlsx"
                resp = requests.get(gsheets_url, timeout=30)
                resp.raise_for_status()
                stok_df = pd.read_excel(BytesIO(resp.content), sheet_name="Stok")

                kod_col = stok_df.columns[0]
                lokasyon_col = self.lokasyon  # DEPO, SUBE, EXC
                malzeme_adi_col = u'Malzeme Ad\u0131'
                malzeme_kodu_col = 'Malzeme Kodu'

                if lokasyon_col in stok_df.columns:
                    for _, srow in stok_df.iterrows():
                        kod = str(srow[kod_col]).strip() if pd.notna(srow[kod_col]) else ''
                        if not kod:
                            continue
                        bek = float(srow[lokasyon_col]) if pd.notna(srow[lokasyon_col]) else 0
                        mal_adi = str(srow[malzeme_adi_col]) if malzeme_adi_col in stok_df.columns and pd.notna(srow.get(malzeme_adi_col)) else kod
                        mal_kodu = str(srow[malzeme_kodu_col]).strip() if malzeme_kodu_col in stok_df.columns and pd.notna(srow.get(malzeme_kodu_col)) else ''
                        beklenen_map[kod] = {'malzeme_adi': mal_adi, 'beklenen': bek, 'malzeme_kodu': mal_kodu}
                    logger.info(f"PRGsheet Stok: {len(beklenen_map)} urun beklenen verisi yuklendi ({self.lokasyon})")
                else:
                    logger.warning(f"PRGsheet Stok: '{lokasyon_col}' kolonu bulunamadi")
            except Exception as e:
                logger.warning(f"PRGsheet stok verisi yuklenemedi: {e}")

            # Toplam sayilan hesapla (stok_kod bazinda tum oturumlar dahil)
            sayilan_toplam = {}
            for rd in all_data:
                sk = rd.get('stok_kod', '')
                m = int(float(rd.get('miktar', 0) or 0))
                sayilan_toplam[sk] = sayilan_toplam.get(sk, 0) + m

            # Beklenen, fark ve malzeme_kodu bilgisini her satira ekle
            for rd in all_data:
                sk = rd.get('stok_kod', '')
                bek_info = beklenen_map.get(sk, {})
                beklenen = bek_info.get('beklenen', 0)
                toplam = sayilan_toplam.get(sk, 0)
                rd['beklenen'] = beklenen
                rd['fark'] = toplam - beklenen
                rd['malzeme_kodu'] = bek_info.get('malzeme_kodu', '') or sk

            # Sayilmamis urunleri ekle (beklenen > 0 ama hic okuma yok)
            lok_kodu_map = {'DEPO': '100', 'SUBE': '200', 'EXC': '300'}
            counted_kodlar = set(sayilan_toplam.keys())
            for sk, info in beklenen_map.items():
                if sk in counted_kodlar or info['beklenen'] <= 0:
                    continue
                row = {
                    'composite_key': f"beklenen::{sk}",
                    'sayim_kodu': '',
                    'lokasyon': self.lokasyon or '',
                    'lokasyon_kodu': lok_kodu_map.get(self.lokasyon, ''),
                    'durum': '',
                    'tarih': '',
                    'stok_kod': sk,
                    'malzeme_kodu': info.get('malzeme_kodu', '') or sk,
                    'malzeme_adi': info['malzeme_adi'],
                    'miktar': 0,
                    'paket_sayisi': 0,
                    'kullanici': '',
                    'beklenen': info['beklenen'],
                    'fark': -info['beklenen'],
                }
                all_data.append(row)

            self.data_loaded.emit(all_data, readings_map)
        except Exception as e:
            self.error_occurred.emit(str(e))


# ================== SYNC THREAD ==================
class SyncThread(QThread):
    """Mikro -> Supabase senkronizasyon thread'i"""

    progress_updated = pyqtSignal(int, str)
    sync_finished = pyqtSignal(dict)
    error_occurred = pyqtSignal(str)

    def __init__(self, supabase_client: SupabaseClient, dogtas_client: DogtasApiClient = None):
        super().__init__()
        self.supabase_client = supabase_client
        self.dogtas_client = dogtas_client
        self.is_cancelled = False

    def run(self):
        try:
            # 1. Supabase'den okumasi olan fatura numaralarini al
            self.progress_updated.emit(5, "Okumasi olan faturalar kontrol ediliyor...")
            okunan_fatura_nolar = self.supabase_client.get_fatura_no_with_readings()
            logger.info(f"Okumasi olan {len(okunan_fatura_nolar)} fatura bulundu")

            if self.is_cancelled:
                return

            # 2. Mikro SQL Server'dan son 7 gunun faturalarini cek
            self.progress_updated.emit(15, "Mikro SQL Server'a baglaniliyor...")
            from datetime import timedelta
            min_tarih = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
            faturalar = self._fetch_from_mikro(min_tarih)

            if self.is_cancelled:
                return

            if not faturalar:
                # Yeni fatura yok ama adres sync'i yine de calistir
                self.progress_updated.emit(80, "Cari adres bilgileri kontrol ediliyor...")
                try:
                    self._sync_all_cari_addresses()
                except Exception as e:
                    logger.error(f"Cari adres sync hatasi: {e}", exc_info=True)
                    self.progress_updated.emit(0, f"Adres sync hatasi: {e}")

                self.sync_finished.emit({
                    'eklenen': 0,
                    'atlanan': 0,
                    'evrak_sayisi': 0,
                    'toplam': 0,
                    'mesaj': "Mikro'dan son 7 gunde Satis Faturasi bulunamadi."
                })
                return

            # 3. Okumasi olan faturalari filtrele (dokunma)
            faturalar_filtreli = [
                f for f in faturalar
                if f['sth_evrakno_sira'] not in okunan_fatura_nolar
            ]
            atlanan_okumali = len(faturalar) - len(faturalar_filtreli)
            if atlanan_okumali > 0:
                logger.info(f"{atlanan_okumali} fatura okumasi oldugu icin atlanacak")

            self.progress_updated.emit(40, f"{len(faturalar_filtreli)} kayit islenecek ({atlanan_okumali} okumali atlandi)")

            if not faturalar_filtreli:
                # Tum faturalarin okumasi var, upsert gerek yok ama cleanup calistir
                logger.info("Tum faturalarin okumasi var, upsert atlanacak")
                if not self.is_cancelled:
                    self.progress_updated.emit(93, "Cari adres bilgileri aliniyor...")
                    try:
                        self._sync_all_cari_addresses()
                    except Exception as e:
                        logger.warning(f"Cari adres sync hatasi (devam ediliyor): {e}", exc_info=True)

                silinen = 0
                if not self.is_cancelled:
                    self.progress_updated.emit(96, "Iptal edilen faturalar kontrol ediliyor...")
                    try:
                        silinen = self._cleanup_cancelled_invoices(okunan_fatura_nolar)
                    except Exception as e:
                        logger.warning(f"Fatura temizleme hatasi (devam ediliyor): {e}")

                self.progress_updated.emit(100, "Senkronizasyon tamamlandi")
                mesaj = f"Son 7 gunde {len(faturalar)} fatura bulundu, tumunun okumasi var."
                if silinen > 0:
                    mesaj += f" {silinen} iptal edilmis fatura silindi."
                self.sync_finished.emit({
                    'eklenen': 0,
                    'atlanan': atlanan_okumali,
                    'evrak_sayisi': 0,
                    'toplam': len(faturalar),
                    'silinen': silinen,
                    'mesaj': mesaj
                })
                return

            # 4. Benzersiz product code'lari topla
            product_codes = list(set(
                row['sth_stok_kod'][:10]
                for row in faturalar_filtreli
                if row.get('sth_stok_kod')
            ))

            # 5. Dogtas API'den paket bilgisi al
            paket_bilgileri = {}
            if self.dogtas_client and product_codes:
                self.progress_updated.emit(50, f"{len(product_codes)} urun icin paket bilgisi aliniyor...")
                try:
                    paket_bilgileri = self.dogtas_client.get_product_packages(product_codes)
                    logger.info(f"{len(paket_bilgileri)} urun icin paket bilgisi alindi")
                except Exception as e:
                    logger.warning(f"Dogtas API hatasi (devam ediliyor): {e}")

            if self.is_cancelled:
                return

            # 6. Veriyi donustur
            self.progress_updated.emit(60, "Veriler donusturuluyor...")
            records = []
            for fatura in faturalar_filtreli:
                product_code = fatura['sth_stok_kod'][:10] if fatura.get('sth_stok_kod') else None
                paket_info = paket_bilgileri.get(product_code) if product_code else None

                tarih_val = fatura.get('tarih')
                if tarih_val is not None:
                    tarih_str = str(tarih_val)
                else:
                    tarih_str = None

                record = {
                    'evrakno_seri': fatura.get('sth_evrakno_seri', '') or '',
                    'evrakno_sira': fatura['sth_evrakno_sira'],
                    'satirno': fatura.get('sth_satirno', 0),
                    'tarih': tarih_str,
                    'stok_kod': fatura.get('sth_stok_kod'),
                    'miktar': float(fatura.get('sth_miktar', 0) or 0),
                    'cikis_depo_no': fatura.get('sth_cikis_depo_no'),
                    'evrak_adi': fatura.get('evrak_adi') or 'Satis Faturasi',
                    'cari_kodu': fatura.get('cari_kodu') or '',
                    'cari_adi': fatura.get('cari_adi') or '',
                    'product_code': product_code,
                    'product_desc': paket_info['productDesc'] if paket_info else None,
                    'paket_sayisi': paket_info['paketSayisi'] if paket_info else 1,
                    'satinalma_kalem_id': fatura.get('bag_kodu'),
                    'malzeme_adi': fatura.get('malzeme_adi'),
                    'plasiyer_kodu': (fatura.get('sth_plasiyer_kodu') or '')[2:] if len(fatura.get('sth_plasiyer_kodu') or '') > 2 else fatura.get('sth_plasiyer_kodu') or '',
                }
                records.append(record)

            # 6. Duplicate kayitlari temizle (JOIN'lerden dolayi ayni urun birden fazla gelebilir)
            unique_records = {}
            for rec in records:
                key = (rec['evrakno_seri'], rec['evrakno_sira'], rec['stok_kod'], rec['satirno'])
                unique_records[key] = rec  # ayni key varsa son kayit kalir
            records = list(unique_records.values())
            logger.info(f"Duplicate temizleme sonrasi: {len(records)} benzersiz kayit")

            # 7. Supabase'e batch upsert
            self.progress_updated.emit(65, "Supabase'e kaydediliyor...")
            eklenen = 0
            atlanan = 0

            for i in range(0, len(records), BATCH_SIZE):
                batch = records[i:i + BATCH_SIZE]
                try:
                    self.supabase_client.upsert_batch(batch)
                    eklenen += len(batch)
                except requests.exceptions.HTTPError as e:
                    hata_detay = e.response.text if e.response is not None else str(e)
                    logger.error(f"Batch upsert hatasi (index {i}): {hata_detay}")
                    self.progress_updated.emit(0, f"Batch hatasi: {hata_detay[:200]}")
                    atlanan += len(batch)
                except Exception as e:
                    logger.error(f"Batch upsert hatasi (index {i}): {e}")
                    atlanan += len(batch)

                progress = 65 + int(((i + len(batch)) / len(records)) * 30)
                self.progress_updated.emit(
                    min(progress, 95),
                    f"{eklenen}/{len(records)} kayit kaydedildi"
                )
                if self.is_cancelled:
                    return

            # Benzersiz evrak sayisi
            evrak_set = set()
            for f in faturalar_filtreli:
                key = f"{f.get('sth_evrakno_seri', '')}_{f['sth_evrakno_sira']}"
                evrak_set.add(key)

            # 8. Cari adres bilgilerini sync et
            if not self.is_cancelled:
                self.progress_updated.emit(93, "Cari adres bilgileri aliniyor...")
                try:
                    self._sync_all_cari_addresses()
                except Exception as e:
                    logger.warning(f"Cari adres sync hatasi (devam ediliyor): {e}", exc_info=True)
                    self.progress_updated.emit(97, f"Adres sync hatasi: {e}")

            # 9. Iptal edilen faturalari temizle (okumasi olanlar korunur)
            silinen = 0
            if not self.is_cancelled:
                self.progress_updated.emit(96, "Iptal edilen faturalar kontrol ediliyor...")
                try:
                    silinen = self._cleanup_cancelled_invoices(okunan_fatura_nolar)
                    if silinen > 0:
                        logger.info(f"{silinen} iptal edilmis fatura Supabase'den silindi")
                except Exception as e:
                    logger.warning(f"Fatura temizleme hatasi (devam ediliyor): {e}")

            self.progress_updated.emit(100, "Senkronizasyon tamamlandi")
            mesaj = f"{len(evrak_set)} Satis Faturasi ({eklenen} Satir) Kaydedildi."
            if atlanan_okumali > 0:
                mesaj += f" {atlanan_okumali} okumali fatura atlandi."
            if silinen > 0:
                mesaj += f" {silinen} iptal edilmis fatura silindi."
            self.sync_finished.emit({
                'eklenen': eklenen,
                'atlanan': atlanan,
                'evrak_sayisi': len(evrak_set),
                'toplam': len(records),
                'silinen': silinen,
                'mesaj': mesaj
            })

        except Exception as e:
            logger.exception("Sync hatasi")
            self.error_occurred.emit(str(e))

    def _fetch_from_mikro(self, min_tarih: str) -> list:
        """Mikro'dan son 7 gunun satis faturalarini cek.
        min_tarih: 'YYYY-MM-DD' formatinda tarih"""
        server = os.getenv('SQL_SERVER', '')
        database = os.getenv('SQL_DATABASE', '')
        username = os.getenv('SQL_USERNAME', '')
        password = os.getenv('SQL_PASSWORD', '')

        if not all([server, database, username]):
            raise Exception("SQL baglanti bilgileri eksik! PRGsheet/Ayar sayfasini kontrol edin.")

        conn_str = (
            f'DRIVER={{ODBC Driver 17 for SQL Server}};'
            f'SERVER={server};'
            f'DATABASE={database};'
            f'UID={username};'
            f'PWD={password}'
        )

        logger.info(f"Mikro SQL Server'a baglaniliyor: {server}/{database} (tarih >= {min_tarih})")

        with pyodbc.connect(conn_str, timeout=30) as conn:
            cursor = conn.cursor()
            cursor.execute(MIKRO_SQL_QUERY, min_tarih)
            columns = [col[0] for col in cursor.description]
            rows = cursor.fetchall()

        logger.info(f"Mikro'dan {len(rows)} kayit alindi (tarih >= {min_tarih})")
        return [dict(zip(columns, row)) for row in rows]

    def _cleanup_cancelled_invoices(self, okunan_fatura_nolar: set = None) -> int:
        """Mikro'da silinmis (iptal edilmis) faturalari Supabase'den temizle.
        Sadece son 7 gunun faturalarini kontrol eder.
        Okumasi olan faturalar asla silinmez."""
        if okunan_fatura_nolar is None:
            okunan_fatura_nolar = set()

        # Mikro'daki son 7 gunun aktif evrakno_sira'lari
        active_sira = self._fetch_active_evrakno_sira()

        # Supabase'deki son 7 gunun evrakno_sira'lari
        from datetime import timedelta
        min_tarih = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
        supabase_sira = self.supabase_client.get_all_evrakno_sira(min_tarih=min_tarih)

        # Mikro'da olmayan Supabase kayitlarini bul
        to_delete_candidates = supabase_sira - active_sira

        # Okumasi olan faturalari koru - asla silme
        to_delete = [s for s in to_delete_candidates if s not in okunan_fatura_nolar]

        korunan = len(to_delete_candidates) - len(to_delete)
        if korunan > 0:
            logger.info(f"{korunan} iptal edilmis fatura okumasi oldugu icin korunuyor")

        if to_delete:
            logger.info(f"Iptal edilmis {len(to_delete)} fatura silinecek: {to_delete}")
            self.supabase_client.delete_by_evrakno_sira_list(to_delete)

        return len(to_delete)

    def _fetch_active_evrakno_sira(self) -> set:
        """Mikro'daki son 7 gunun aktif satis faturasi evrakno_sira degerlerini al"""
        server = os.getenv('SQL_SERVER', '')
        database = os.getenv('SQL_DATABASE', '')
        username = os.getenv('SQL_USERNAME', '')
        password = os.getenv('SQL_PASSWORD', '')

        conn_str = (
            f'DRIVER={{ODBC Driver 17 for SQL Server}};'
            f'SERVER={server};'
            f'DATABASE={database};'
            f'UID={username};'
            f'PWD={password}'
        )

        from datetime import timedelta
        min_tarih = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')

        query = """
            SELECT DISTINCT sth_evrakno_sira
            FROM dbo.STOK_HAREKETLERI WITH (NOLOCK)
            WHERE sth_evraktip = 4 AND sth_belge_tarih >= ?
        """

        with pyodbc.connect(conn_str, timeout=30) as conn:
            cursor = conn.cursor()
            cursor.execute(query, min_tarih)
            rows = cursor.fetchall()

        logger.info(f"Mikro'da son 7 gunde {len(rows)} aktif evrakno_sira bulundu (tarih >= {min_tarih})")
        return set(row[0] for row in rows)

    def _sync_all_cari_addresses(self):
        """Supabase'deki tum faturalardaki cari kodlarin adres bilgilerini sync et"""
        self.progress_updated.emit(93, "Supabase'den cari kodlar aliniyor...")

        # Supabase'deki tum faturalardan benzersiz cari kodlari al
        all_invoices = self.supabase_client.get_all_invoices(limit=5000)
        self.progress_updated.emit(94, f"Supabase'den {len(all_invoices)} fatura alindi")
        logger.info(f"Adres sync: Supabase'den {len(all_invoices)} fatura alindi")

        cari_kodlar = list(set(
            r.get('cari_kodu') for r in all_invoices
            if r.get('cari_kodu')
        ))

        if not cari_kodlar:
            logger.info("Adres sync: Hicbir cari kodu bulunamadi")
            return

        logger.info(f"Adres sync: {len(cari_kodlar)} benzersiz cari kodu bulundu")

        self.progress_updated.emit(95, f"{len(cari_kodlar)} cari kod icin adres bilgisi aliniyor...")
        adres_records = self._fetch_cari_addresses(cari_kodlar)
        if adres_records:
            self.progress_updated.emit(96, f"{len(adres_records)} adres Supabase'e kaydediliyor...")
            try:
                for i in range(0, len(adres_records), BATCH_SIZE):
                    batch = adres_records[i:i + BATCH_SIZE]
                    self.supabase_client.upsert_adres_batch(batch)
                logger.info(f"{len(adres_records)} cari adres Supabase'e kaydedildi")
                self.progress_updated.emit(97, f"{len(adres_records)} cari adres kaydedildi")
            except Exception as e:
                logger.error(f"Adres upsert hatasi: {e}", exc_info=True)
                self.progress_updated.emit(97, f"ADRES KAYIT HATASI: {e}")
        else:
            logger.info("Mikro'dan hicbir adres bilgisi alinamadi")
            self.progress_updated.emit(97, "Mikro'dan adres bilgisi alinamadi")

    def _fetch_cari_addresses(self, cari_kod_list: list) -> list:
        """Mikro'dan cari hesap + adres bilgilerini cek"""
        if not cari_kod_list:
            return []

        server = os.getenv('SQL_SERVER', '')
        database = os.getenv('SQL_DATABASE', '')
        username = os.getenv('SQL_USERNAME', '')
        password = os.getenv('SQL_PASSWORD', '')

        conn_str = (
            f'DRIVER={{ODBC Driver 17 for SQL Server}};'
            f'SERVER={server};'
            f'DATABASE={database};'
            f'UID={username};'
            f'PWD={password}'
        )

        all_results = []
        batch_size = 50
        for i in range(0, len(cari_kod_list), batch_size):
            batch = cari_kod_list[i:i + batch_size]
            placeholders = ','.join([f"N'{kod}'" for kod in batch])
            query = MIKRO_CARI_ADRES_QUERY.format(placeholders=placeholders)

            logger.info(f"Adres batch {i//batch_size + 1}: {len(batch)} cari kod sorgulanıyor")

            with pyodbc.connect(conn_str, timeout=30) as conn:
                cursor = conn.cursor()
                cursor.execute(query)
                columns = [col[0] for col in cursor.description]
                rows = cursor.fetchall()

            logger.info(f"Adres batch {i//batch_size + 1}: {len(rows)} sonuc geldi")

            for row in rows:
                record = dict(zip(columns, row))
                all_results.append({
                    k.lower(): (str(v).strip() if v is not None else None)
                    for k, v in record.items()
                })

        logger.info(f"Mikro'dan {len(all_results)} cari adres bilgisi alindi")
        return all_results

    def cancel(self):
        self.is_cancelled = True


# ================== CIKIS FISI SYNC THREAD ==================
class CikisSyncThread(QThread):
    """Mikro -> Supabase cikis fisi senkronizasyon thread'i"""

    progress_updated = pyqtSignal(int, str)
    sync_finished = pyqtSignal(dict)
    error_occurred = pyqtSignal(str)

    def __init__(self, supabase_client: SupabaseClient, dogtas_client: DogtasApiClient = None):
        super().__init__()
        self.supabase_client = supabase_client
        self.dogtas_client = dogtas_client
        self.is_cancelled = False

    def run(self):
        try:
            # 1. Supabase'den okumasi olan fis numaralarini al
            self.progress_updated.emit(5, "Okumasi olan fisler kontrol ediliyor...")
            okunan_fis_nolar = self.supabase_client.get_cikis_fis_no_with_readings()
            logger.info(f"Cikis: Okumasi olan {len(okunan_fis_nolar)} fis bulundu")

            if self.is_cancelled:
                return

            # 2. Mikro SQL Server'dan son 7 gunun cikis fislerini cek
            self.progress_updated.emit(15, "Mikro SQL Server'a baglaniliyor...")
            from datetime import timedelta
            min_tarih = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
            fisler = self._fetch_from_mikro(min_tarih)

            if self.is_cancelled:
                return

            if not fisler:
                self.sync_finished.emit({
                    'eklenen': 0,
                    'atlanan': 0,
                    'evrak_sayisi': 0,
                    'toplam': 0,
                    'mesaj': u"Mikro'dan son 7 g\u00fcnde \u00c7\u0131k\u0131\u015f Fi\u015fi bulunamad\u0131."
                })
                return

            # 3. Okumasi olan fisleri filtrele (dokunma)
            fisler_filtreli = [
                f for f in fisler
                if f['sth_evrakno_sira'] not in okunan_fis_nolar
            ]
            atlanan_okumali = len(fisler) - len(fisler_filtreli)
            if atlanan_okumali > 0:
                logger.info(f"Cikis: {atlanan_okumali} fis okumasi oldugu icin atlanacak")

            self.progress_updated.emit(40, f"{len(fisler_filtreli)} kayit islenecek ({atlanan_okumali} okumali atlandi)")

            if not fisler_filtreli:
                # Tum fislerin okumasi var, cleanup calistir
                silinen = 0
                if not self.is_cancelled:
                    self.progress_updated.emit(96, "Iptal edilen fisler kontrol ediliyor...")
                    try:
                        silinen = self._cleanup_cancelled_fis(okunan_fis_nolar)
                    except Exception as e:
                        logger.warning(f"Cikis fis temizleme hatasi: {e}")

                self.progress_updated.emit(100, "Senkronizasyon tamamlandi")
                mesaj = f"Son 7 gunde {len(fisler)} fis bulundu, tumunun okumasi var."
                if silinen > 0:
                    mesaj += f" {silinen} iptal edilmis fis silindi."
                self.sync_finished.emit({
                    'eklenen': 0,
                    'atlanan': atlanan_okumali,
                    'evrak_sayisi': 0,
                    'toplam': len(fisler),
                    'silinen': silinen,
                    'mesaj': mesaj
                })
                return

            # 4. Benzersiz product code'lari topla
            product_codes = list(set(
                row['sth_stok_kod'][:10]
                for row in fisler_filtreli
                if row.get('sth_stok_kod')
            ))

            # 5. Dogtas API'den paket bilgisi al
            paket_bilgileri = {}
            if self.dogtas_client and product_codes:
                self.progress_updated.emit(50, f"{len(product_codes)} urun icin paket bilgisi aliniyor...")
                try:
                    paket_bilgileri = self.dogtas_client.get_product_packages(product_codes)
                    logger.info(f"Cikis: {len(paket_bilgileri)} urun icin paket bilgisi alindi")
                except Exception as e:
                    logger.warning(f"Dogtas API hatasi (devam ediliyor): {e}")

            if self.is_cancelled:
                return

            # 6. Veriyi donustur
            self.progress_updated.emit(60, "Veriler donusturuluyor...")
            records = []
            for fis in fisler_filtreli:
                product_code = fis['sth_stok_kod'][:10] if fis.get('sth_stok_kod') else None
                paket_info = paket_bilgileri.get(product_code) if product_code else None

                tarih_val = fis.get('tarih')
                tarih_str = str(tarih_val) if tarih_val is not None else None

                record = {
                    'evrakno_seri': fis.get('sth_evrakno_seri', '') or '',
                    'evrakno_sira': fis['sth_evrakno_sira'],
                    'tarih': tarih_str,
                    'stok_kod': fis.get('sth_stok_kod'),
                    'miktar': float(fis.get('sth_miktar', 0) or 0),
                    'depo': fis.get('sth_cikis_depo_no'),
                    'evrak_adi': fis.get('evrak_adi') or u'\u00c7\u0131k\u0131\u015f Fi\u015fi',
                    'paket_sayisi': paket_info['paketSayisi'] if paket_info else 1,
                    'malzeme_adi': fis.get('malzeme_adi'),
                    'satinalma_kalem_id': fis.get('bag_kodu'),
                }
                records.append(record)

            # Duplicate kayitlari temizle
            unique_records = {}
            for rec in records:
                key = (rec['evrakno_seri'], rec['evrakno_sira'], rec['stok_kod'])
                unique_records[key] = rec
            records = list(unique_records.values())
            logger.info(f"Cikis: Duplicate temizleme sonrasi: {len(records)} benzersiz kayit")

            # 7. Supabase'e batch upsert
            self.progress_updated.emit(65, "Supabase'e kaydediliyor...")
            eklenen = 0
            atlanan = 0

            for i in range(0, len(records), BATCH_SIZE):
                batch = records[i:i + BATCH_SIZE]
                try:
                    self.supabase_client.upsert_cikis_batch(batch)
                    eklenen += len(batch)
                except requests.exceptions.HTTPError as e:
                    hata_detay = e.response.text if e.response is not None else str(e)
                    logger.error(f"Cikis batch upsert hatasi (index {i}): {hata_detay}")
                    self.progress_updated.emit(0, f"Batch hatasi: {hata_detay[:200]}")
                    atlanan += len(batch)
                except Exception as e:
                    logger.error(f"Cikis batch upsert hatasi (index {i}): {e}")
                    atlanan += len(batch)

                progress = 65 + int(((i + len(batch)) / len(records)) * 30)
                self.progress_updated.emit(
                    min(progress, 95),
                    f"{eklenen}/{len(records)} kayit kaydedildi"
                )
                if self.is_cancelled:
                    return

            # Benzersiz evrak sayisi
            evrak_set = set()
            for f in fisler_filtreli:
                key = f"{f.get('sth_evrakno_seri', '')}_{f['sth_evrakno_sira']}"
                evrak_set.add(key)

            # 8. Iptal edilen fisleri temizle
            silinen = 0
            if not self.is_cancelled:
                self.progress_updated.emit(96, "Iptal edilen fisler kontrol ediliyor...")
                try:
                    silinen = self._cleanup_cancelled_fis(okunan_fis_nolar)
                    if silinen > 0:
                        logger.info(f"Cikis: {silinen} iptal edilmis fis silindi")
                except Exception as e:
                    logger.warning(f"Cikis fis temizleme hatasi: {e}")

            self.progress_updated.emit(100, "Senkronizasyon tamamlandi")
            mesaj = u"{} \u00c7\u0131k\u0131\u015f Fi\u015fi ({} Sat\u0131r) Kaydedildi.".format(len(evrak_set), eklenen)
            if atlanan_okumali > 0:
                mesaj += f" {atlanan_okumali} okumali fis atlandi."
            if silinen > 0:
                mesaj += f" {silinen} iptal edilmis fis silindi."
            self.sync_finished.emit({
                'eklenen': eklenen,
                'atlanan': atlanan,
                'evrak_sayisi': len(evrak_set),
                'toplam': len(records),
                'silinen': silinen,
                'mesaj': mesaj
            })

        except Exception as e:
            logger.exception("Cikis sync hatasi")
            self.error_occurred.emit(str(e))

    def _fetch_from_mikro(self, min_tarih: str) -> list:
        """Mikro'dan son 7 gunun cikis fislerini cek."""
        server = os.getenv('SQL_SERVER', '')
        database = os.getenv('SQL_DATABASE', '')
        username = os.getenv('SQL_USERNAME', '')
        password = os.getenv('SQL_PASSWORD', '')

        if not all([server, database, username]):
            raise Exception("SQL baglanti bilgileri eksik!")

        conn_str = (
            f'DRIVER={{ODBC Driver 17 for SQL Server}};'
            f'SERVER={server};'
            f'DATABASE={database};'
            f'UID={username};'
            f'PWD={password}'
        )

        logger.info(f"Cikis: Mikro SQL Server'a baglaniliyor: {server}/{database} (tarih >= {min_tarih})")

        with pyodbc.connect(conn_str, timeout=30) as conn:
            cursor = conn.cursor()
            cursor.execute(MIKRO_CIKIS_SQL_QUERY, min_tarih)
            columns = [col[0] for col in cursor.description]
            rows = cursor.fetchall()

        logger.info(f"Cikis: Mikro'dan {len(rows)} kayit alindi (tarih >= {min_tarih})")
        return [dict(zip(columns, row)) for row in rows]

    def _cleanup_cancelled_fis(self, okunan_fis_nolar: set = None) -> int:
        """Mikro'da silinmis cikis fislerini Supabase'den temizle."""
        if okunan_fis_nolar is None:
            okunan_fis_nolar = set()

        # Mikro'daki son 7 gunun aktif evrakno_sira'lari
        active_sira = self._fetch_active_cikis_evrakno_sira()

        # Supabase'deki son 7 gunun evrakno_sira'lari
        from datetime import timedelta
        min_tarih = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
        supabase_sira = self.supabase_client.get_cikis_all_evrakno_sira(min_tarih=min_tarih)

        # Mikro'da olmayan Supabase kayitlarini bul
        to_delete_candidates = supabase_sira - active_sira

        # Okumasi olan fisleri koru
        to_delete = [s for s in to_delete_candidates if s not in okunan_fis_nolar]

        korunan = len(to_delete_candidates) - len(to_delete)
        if korunan > 0:
            logger.info(f"Cikis: {korunan} iptal edilmis fis okumasi oldugu icin korunuyor")

        if to_delete:
            logger.info(f"Cikis: Iptal edilmis {len(to_delete)} fis silinecek: {to_delete}")
            self.supabase_client.delete_cikis_by_evrakno_sira_list(to_delete)

        return len(to_delete)

    def _fetch_active_cikis_evrakno_sira(self) -> set:
        """Mikro'daki son 7 gunun aktif cikis fisi evrakno_sira degerlerini al"""
        server = os.getenv('SQL_SERVER', '')
        database = os.getenv('SQL_DATABASE', '')
        username = os.getenv('SQL_USERNAME', '')
        password = os.getenv('SQL_PASSWORD', '')

        conn_str = (
            f'DRIVER={{ODBC Driver 17 for SQL Server}};'
            f'SERVER={server};'
            f'DATABASE={database};'
            f'UID={username};'
            f'PWD={password}'
        )

        from datetime import timedelta
        min_tarih = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')

        query = """
            SELECT DISTINCT sth_evrakno_sira
            FROM dbo.STOK_HAREKETLERI WITH (NOLOCK)
            WHERE sth_evraktip = 0 AND sth_belge_tarih >= ?
        """

        with pyodbc.connect(conn_str, timeout=30) as conn:
            cursor = conn.cursor()
            cursor.execute(query, min_tarih)
            rows = cursor.fetchall()

        logger.info(f"Cikis: Mikro'da son 7 gunde {len(rows)} aktif cikis evrakno_sira bulundu")
        return set(row[0] for row in rows)

    def cancel(self):
        self.is_cancelled = True


# ================== GIRIS FISI SYNC THREAD ==================
class GirisSyncThread(QThread):
    """Mikro -> Supabase giris fisi senkronizasyon thread'i"""

    progress_updated = pyqtSignal(int, str)
    sync_finished = pyqtSignal(dict)
    error_occurred = pyqtSignal(str)

    def __init__(self, supabase_client: SupabaseClient, dogtas_client: DogtasApiClient = None):
        super().__init__()
        self.supabase_client = supabase_client
        self.dogtas_client = dogtas_client
        self.is_cancelled = False

    def run(self):
        try:
            # 1. Supabase'den okumasi olan fis numaralarini al
            self.progress_updated.emit(5, "Okumasi olan fisler kontrol ediliyor...")
            okunan_fis_nolar = self.supabase_client.get_giris_fis_no_with_readings()
            logger.info(f"Giris: Okumasi olan {len(okunan_fis_nolar)} fis bulundu")

            if self.is_cancelled:
                return

            # 2. Mikro SQL Server'dan son 7 gunun giris fislerini cek
            self.progress_updated.emit(15, "Mikro SQL Server'a baglaniliyor...")
            from datetime import timedelta
            min_tarih = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
            fisler = self._fetch_from_mikro(min_tarih)

            if self.is_cancelled:
                return

            if not fisler:
                self.sync_finished.emit({
                    'eklenen': 0,
                    'atlanan': 0,
                    'evrak_sayisi': 0,
                    'toplam': 0,
                    'mesaj': u"Mikro'dan son 7 g\u00fcnde Giri\u015f Fi\u015fi bulunamad\u0131."
                })
                return

            # 3. Okumasi olan fisleri filtrele (dokunma)
            fisler_filtreli = [
                f for f in fisler
                if f['sth_evrakno_sira'] not in okunan_fis_nolar
            ]
            atlanan_okumali = len(fisler) - len(fisler_filtreli)
            if atlanan_okumali > 0:
                logger.info(f"Giris: {atlanan_okumali} fis okumasi oldugu icin atlanacak")

            self.progress_updated.emit(40, f"{len(fisler_filtreli)} kayit islenecek ({atlanan_okumali} okumali atlandi)")

            if not fisler_filtreli:
                # Tum fislerin okumasi var, cleanup calistir
                silinen = 0
                if not self.is_cancelled:
                    self.progress_updated.emit(96, "Iptal edilen fisler kontrol ediliyor...")
                    try:
                        silinen = self._cleanup_cancelled_fis(okunan_fis_nolar)
                    except Exception as e:
                        logger.warning(f"Giris fis temizleme hatasi: {e}")

                self.progress_updated.emit(100, "Senkronizasyon tamamlandi")
                mesaj = f"Son 7 gunde {len(fisler)} fis bulundu, tumunun okumasi var."
                if silinen > 0:
                    mesaj += f" {silinen} iptal edilmis fis silindi."
                self.sync_finished.emit({
                    'eklenen': 0,
                    'atlanan': atlanan_okumali,
                    'evrak_sayisi': 0,
                    'toplam': len(fisler),
                    'silinen': silinen,
                    'mesaj': mesaj
                })
                return

            # 4. Benzersiz product code'lari topla
            product_codes = list(set(
                row['sth_stok_kod'][:10]
                for row in fisler_filtreli
                if row.get('sth_stok_kod')
            ))

            # 5. Dogtas API'den paket bilgisi al
            paket_bilgileri = {}
            if self.dogtas_client and product_codes:
                self.progress_updated.emit(50, f"{len(product_codes)} urun icin paket bilgisi aliniyor...")
                try:
                    paket_bilgileri = self.dogtas_client.get_product_packages(product_codes)
                    logger.info(f"Giris: {len(paket_bilgileri)} urun icin paket bilgisi alindi")
                except Exception as e:
                    logger.warning(f"Dogtas API hatasi (devam ediliyor): {e}")

            if self.is_cancelled:
                return

            # 6. Veriyi donustur
            self.progress_updated.emit(60, "Veriler donusturuluyor...")
            records = []
            for fis in fisler_filtreli:
                product_code = fis['sth_stok_kod'][:10] if fis.get('sth_stok_kod') else None
                paket_info = paket_bilgileri.get(product_code) if product_code else None

                tarih_val = fis.get('tarih')
                tarih_str = str(tarih_val) if tarih_val is not None else None

                record = {
                    'evrakno_seri': fis.get('sth_evrakno_seri', '') or '',
                    'evrakno_sira': fis['sth_evrakno_sira'],
                    'tarih': tarih_str,
                    'stok_kod': fis.get('sth_stok_kod'),
                    'miktar': float(fis.get('sth_miktar', 0) or 0),
                    'depo': fis.get('depo'),
                    'evrak_adi': fis.get('evrak_adi') or u'Giri\u015f Fi\u015fi',
                    'paket_sayisi': paket_info['paketSayisi'] if paket_info else 1,
                    'malzeme_adi': fis.get('malzeme_adi'),
                    'satinalma_kalem_id': fis.get('bag_kodu'),
                }
                records.append(record)

            # Duplicate kayitlari temizle
            unique_records = {}
            for rec in records:
                key = (rec['evrakno_seri'], rec['evrakno_sira'], rec['stok_kod'])
                unique_records[key] = rec
            records = list(unique_records.values())
            logger.info(f"Giris: Duplicate temizleme sonrasi: {len(records)} benzersiz kayit")

            # 7. Supabase'e batch upsert
            self.progress_updated.emit(65, "Supabase'e kaydediliyor...")
            eklenen = 0
            atlanan = 0

            for i in range(0, len(records), BATCH_SIZE):
                batch = records[i:i + BATCH_SIZE]
                try:
                    self.supabase_client.upsert_giris_batch(batch)
                    eklenen += len(batch)
                except requests.exceptions.HTTPError as e:
                    hata_detay = e.response.text if e.response is not None else str(e)
                    logger.error(f"Giris batch upsert hatasi (index {i}): {hata_detay}")
                    self.progress_updated.emit(0, f"Batch hatasi: {hata_detay[:200]}")
                    atlanan += len(batch)
                except Exception as e:
                    logger.error(f"Giris batch upsert hatasi (index {i}): {e}")
                    atlanan += len(batch)

                progress = 65 + int(((i + len(batch)) / len(records)) * 30)
                self.progress_updated.emit(
                    min(progress, 95),
                    f"{eklenen}/{len(records)} kayit kaydedildi"
                )
                if self.is_cancelled:
                    return

            # Benzersiz evrak sayisi
            evrak_set = set()
            for f in fisler_filtreli:
                key = f"{f.get('sth_evrakno_seri', '')}_{f['sth_evrakno_sira']}"
                evrak_set.add(key)

            # 8. Iptal edilen fisleri temizle
            silinen = 0
            if not self.is_cancelled:
                self.progress_updated.emit(96, "Iptal edilen fisler kontrol ediliyor...")
                try:
                    silinen = self._cleanup_cancelled_fis(okunan_fis_nolar)
                    if silinen > 0:
                        logger.info(f"Giris: {silinen} iptal edilmis fis silindi")
                except Exception as e:
                    logger.warning(f"Giris fis temizleme hatasi: {e}")

            self.progress_updated.emit(100, "Senkronizasyon tamamlandi")
            mesaj = u"{} Giri\u015f Fi\u015fi ({} Sat\u0131r) Kaydedildi.".format(len(evrak_set), eklenen)
            if atlanan_okumali > 0:
                mesaj += f" {atlanan_okumali} okumali fis atlandi."
            if silinen > 0:
                mesaj += f" {silinen} iptal edilmis fis silindi."
            self.sync_finished.emit({
                'eklenen': eklenen,
                'atlanan': atlanan,
                'evrak_sayisi': len(evrak_set),
                'toplam': len(records),
                'silinen': silinen,
                'mesaj': mesaj
            })

        except Exception as e:
            logger.exception("Giris sync hatasi")
            self.error_occurred.emit(str(e))

    def _fetch_from_mikro(self, min_tarih: str) -> list:
        """Mikro'dan son 7 gunun giris fislerini cek."""
        server = os.getenv('SQL_SERVER', '')
        database = os.getenv('SQL_DATABASE', '')
        username = os.getenv('SQL_USERNAME', '')
        password = os.getenv('SQL_PASSWORD', '')

        if not all([server, database, username]):
            raise Exception("SQL baglanti bilgileri eksik!")

        conn_str = (
            f'DRIVER={{ODBC Driver 17 for SQL Server}};'
            f'SERVER={server};'
            f'DATABASE={database};'
            f'UID={username};'
            f'PWD={password}'
        )

        logger.info(f"Giris: Mikro SQL Server'a baglaniliyor: {server}/{database} (tarih >= {min_tarih})")

        with pyodbc.connect(conn_str, timeout=30) as conn:
            cursor = conn.cursor()
            cursor.execute(MIKRO_GIRIS_SQL_QUERY, min_tarih)
            columns = [col[0] for col in cursor.description]
            rows = cursor.fetchall()

        logger.info(f"Giris: Mikro'dan {len(rows)} kayit alindi (tarih >= {min_tarih})")
        return [dict(zip(columns, row)) for row in rows]

    def _cleanup_cancelled_fis(self, okunan_fis_nolar: set = None) -> int:
        """Mikro'da silinmis giris fislerini Supabase'den temizle."""
        if okunan_fis_nolar is None:
            okunan_fis_nolar = set()

        # Mikro'daki son 7 gunun aktif evrakno_sira'lari
        active_sira = self._fetch_active_giris_evrakno_sira()

        # Supabase'deki son 7 gunun evrakno_sira'lari
        from datetime import timedelta
        min_tarih = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
        supabase_sira = self.supabase_client.get_giris_all_evrakno_sira(min_tarih=min_tarih)

        # Mikro'da olmayan Supabase kayitlarini bul
        to_delete_candidates = supabase_sira - active_sira

        # Okumasi olan fisleri koru
        to_delete = [s for s in to_delete_candidates if s not in okunan_fis_nolar]

        korunan = len(to_delete_candidates) - len(to_delete)
        if korunan > 0:
            logger.info(f"Giris: {korunan} iptal edilmis fis okumasi oldugu icin korunuyor")

        if to_delete:
            logger.info(f"Giris: Iptal edilmis {len(to_delete)} fis silinecek: {to_delete}")
            self.supabase_client.delete_giris_by_evrakno_sira_list(to_delete)

        return len(to_delete)

    def _fetch_active_giris_evrakno_sira(self) -> set:
        """Mikro'daki son 7 gunun aktif giris fisi evrakno_sira degerlerini al"""
        server = os.getenv('SQL_SERVER', '')
        database = os.getenv('SQL_DATABASE', '')
        username = os.getenv('SQL_USERNAME', '')
        password = os.getenv('SQL_PASSWORD', '')

        conn_str = (
            f'DRIVER={{ODBC Driver 17 for SQL Server}};'
            f'SERVER={server};'
            f'DATABASE={database};'
            f'UID={username};'
            f'PWD={password}'
        )

        from datetime import timedelta
        min_tarih = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')

        query = """
            SELECT DISTINCT sth_evrakno_sira
            FROM dbo.STOK_HAREKETLERI WITH (NOLOCK)
            WHERE sth_evraktip = 12 AND sth_belge_tarih >= ?
        """

        with pyodbc.connect(conn_str, timeout=30) as conn:
            cursor = conn.cursor()
            cursor.execute(query, min_tarih)
            rows = cursor.fetchall()

        logger.info(f"Giris: Mikro'da son 7 gunde {len(rows)} aktif giris evrakno_sira bulundu")
        return set(row[0] for row in rows)

    def cancel(self):
        self.is_cancelled = True


# ================== SEVK FISI SYNC THREAD ==================
class SevkSyncThread(QThread):
    """Mikro -> Supabase sevk fisi senkronizasyon thread'i"""

    progress_updated = pyqtSignal(int, str)
    sync_finished = pyqtSignal(dict)
    error_occurred = pyqtSignal(str)

    def __init__(self, supabase_client: SupabaseClient, dogtas_client: DogtasApiClient = None):
        super().__init__()
        self.supabase_client = supabase_client
        self.dogtas_client = dogtas_client
        self.is_cancelled = False

    def run(self):
        try:
            # 1. Supabase'den okumasi olan fis numaralarini al
            self.progress_updated.emit(5, "Okumasi olan fisler kontrol ediliyor...")
            okunan_fis_nolar = self.supabase_client.get_sevk_fis_no_with_readings()
            logger.info(f"Sevk: Okumasi olan {len(okunan_fis_nolar)} fis bulundu")

            if self.is_cancelled:
                return

            # 2. Mikro SQL Server'dan son 7 gunun sevk fislerini cek
            self.progress_updated.emit(15, "Mikro SQL Server'a baglaniliyor...")
            from datetime import timedelta
            min_tarih = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
            fisler = self._fetch_from_mikro(min_tarih)

            if self.is_cancelled:
                return

            if not fisler:
                self.sync_finished.emit({
                    'eklenen': 0,
                    'atlanan': 0,
                    'evrak_sayisi': 0,
                    'toplam': 0,
                    'mesaj': u"Mikro'dan son 7 g\u00fcnde Sevk Fi\u015fi bulunamad\u0131."
                })
                return

            # 3. Okumasi olan fisleri filtrele (dokunma)
            fisler_filtreli = [
                f for f in fisler
                if f['sth_evrakno_sira'] not in okunan_fis_nolar
            ]
            atlanan_okumali = len(fisler) - len(fisler_filtreli)
            if atlanan_okumali > 0:
                logger.info(f"Sevk: {atlanan_okumali} fis okumasi oldugu icin atlanacak")

            self.progress_updated.emit(40, f"{len(fisler_filtreli)} kayit islenecek ({atlanan_okumali} okumali atlandi)")

            if not fisler_filtreli:
                silinen = 0
                if not self.is_cancelled:
                    self.progress_updated.emit(96, "Iptal edilen fisler kontrol ediliyor...")
                    try:
                        silinen = self._cleanup_cancelled_fis(okunan_fis_nolar)
                    except Exception as e:
                        logger.warning(f"Sevk fis temizleme hatasi: {e}")

                self.progress_updated.emit(100, "Senkronizasyon tamamlandi")
                mesaj = f"Son 7 gunde {len(fisler)} fis bulundu, tumunun okumasi var."
                if silinen > 0:
                    mesaj += f" {silinen} iptal edilmis fis silindi."
                self.sync_finished.emit({
                    'eklenen': 0,
                    'atlanan': atlanan_okumali,
                    'evrak_sayisi': 0,
                    'toplam': len(fisler),
                    'silinen': silinen,
                    'mesaj': mesaj
                })
                return

            # 4. Benzersiz product code'lari topla
            product_codes = list(set(
                row['sth_stok_kod'][:10]
                for row in fisler_filtreli
                if row.get('sth_stok_kod')
            ))

            # 5. Dogtas API'den paket bilgisi al
            paket_bilgileri = {}
            if self.dogtas_client and product_codes:
                self.progress_updated.emit(50, f"{len(product_codes)} urun icin paket bilgisi aliniyor...")
                try:
                    paket_bilgileri = self.dogtas_client.get_product_packages(product_codes)
                    logger.info(f"Sevk: {len(paket_bilgileri)} urun icin paket bilgisi alindi")
                except Exception as e:
                    logger.warning(f"Dogtas API hatasi (devam ediliyor): {e}")

            if self.is_cancelled:
                return

            # 6. Veriyi donustur
            self.progress_updated.emit(60, "Veriler donusturuluyor...")
            records = []
            for fis in fisler_filtreli:
                product_code = fis['sth_stok_kod'][:10] if fis.get('sth_stok_kod') else None
                paket_info = paket_bilgileri.get(product_code) if product_code else None

                tarih_val = fis.get('tarih')
                tarih_str = str(tarih_val) if tarih_val is not None else None

                record = {
                    'evrakno_seri': fis.get('sth_evrakno_seri', '') or '',
                    'evrakno_sira': fis['sth_evrakno_sira'],
                    'tarih': tarih_str,
                    'stok_kod': fis.get('sth_stok_kod'),
                    'miktar': float(fis.get('sth_miktar', 0) or 0),
                    'cikis_depo': fis.get('cikis_depo'),
                    'giris_depo': fis.get('giris_depo'),
                    'evrak_adi': fis.get('evrak_adi') or u'Sevk Fi\u015fi',
                    'paket_sayisi': paket_info['paketSayisi'] if paket_info else 1,
                    'malzeme_adi': fis.get('malzeme_adi'),
                    'satinalma_kalem_id': fis.get('bag_kodu'),
                }
                records.append(record)

            # Duplicate kayitlari temizle
            unique_records = {}
            for rec in records:
                key = (rec['evrakno_seri'], rec['evrakno_sira'], rec['stok_kod'])
                unique_records[key] = rec
            records = list(unique_records.values())
            logger.info(f"Sevk: Duplicate temizleme sonrasi: {len(records)} benzersiz kayit")

            # 7. Supabase'e batch upsert
            self.progress_updated.emit(65, "Supabase'e kaydediliyor...")
            eklenen = 0
            atlanan = 0

            for i in range(0, len(records), BATCH_SIZE):
                batch = records[i:i + BATCH_SIZE]
                try:
                    self.supabase_client.upsert_sevk_batch(batch)
                    eklenen += len(batch)
                except requests.exceptions.HTTPError as e:
                    hata_detay = e.response.text if e.response is not None else str(e)
                    logger.error(f"Sevk batch upsert hatasi (index {i}): {hata_detay}")
                    self.progress_updated.emit(0, f"Batch hatasi: {hata_detay[:200]}")
                    atlanan += len(batch)
                except Exception as e:
                    logger.error(f"Sevk batch upsert hatasi (index {i}): {e}")
                    atlanan += len(batch)

                progress = 65 + int(((i + len(batch)) / len(records)) * 30)
                self.progress_updated.emit(
                    min(progress, 95),
                    f"{eklenen}/{len(records)} kayit kaydedildi"
                )
                if self.is_cancelled:
                    return

            # Benzersiz evrak sayisi
            evrak_set = set()
            for f in fisler_filtreli:
                key = f"{f.get('sth_evrakno_seri', '')}_{f['sth_evrakno_sira']}"
                evrak_set.add(key)

            # 8. Iptal edilen fisleri temizle
            silinen = 0
            if not self.is_cancelled:
                self.progress_updated.emit(96, "Iptal edilen fisler kontrol ediliyor...")
                try:
                    silinen = self._cleanup_cancelled_fis(okunan_fis_nolar)
                    if silinen > 0:
                        logger.info(f"Sevk: {silinen} iptal edilmis fis silindi")
                except Exception as e:
                    logger.warning(f"Sevk fis temizleme hatasi: {e}")

            self.progress_updated.emit(100, "Senkronizasyon tamamlandi")
            mesaj = u"{} Sevk Fi\u015fi ({} Sat\u0131r) Kaydedildi.".format(len(evrak_set), eklenen)
            if atlanan_okumali > 0:
                mesaj += f" {atlanan_okumali} okumali fis atlandi."
            if silinen > 0:
                mesaj += f" {silinen} iptal edilmis fis silindi."
            self.sync_finished.emit({
                'eklenen': eklenen,
                'atlanan': atlanan,
                'evrak_sayisi': len(evrak_set),
                'toplam': len(records),
                'silinen': silinen,
                'mesaj': mesaj
            })

        except Exception as e:
            logger.exception("Sevk sync hatasi")
            self.error_occurred.emit(str(e))

    def _fetch_from_mikro(self, min_tarih: str) -> list:
        """Mikro'dan son 7 gunun sevk fislerini cek."""
        server = os.getenv('SQL_SERVER', '')
        database = os.getenv('SQL_DATABASE', '')
        username = os.getenv('SQL_USERNAME', '')
        password = os.getenv('SQL_PASSWORD', '')

        if not all([server, database, username]):
            raise Exception("SQL baglanti bilgileri eksik!")

        conn_str = (
            f'DRIVER={{ODBC Driver 17 for SQL Server}};'
            f'SERVER={server};'
            f'DATABASE={database};'
            f'UID={username};'
            f'PWD={password}'
        )

        logger.info(f"Sevk: Mikro SQL Server'a baglaniliyor: {server}/{database} (tarih >= {min_tarih})")

        with pyodbc.connect(conn_str, timeout=30) as conn:
            cursor = conn.cursor()
            cursor.execute(MIKRO_SEVK_SQL_QUERY, min_tarih)
            columns = [col[0] for col in cursor.description]
            rows = cursor.fetchall()

        logger.info(f"Sevk: Mikro'dan {len(rows)} kayit alindi (tarih >= {min_tarih})")
        return [dict(zip(columns, row)) for row in rows]

    def _cleanup_cancelled_fis(self, okunan_fis_nolar: set = None) -> int:
        """Mikro'da silinmis sevk fislerini Supabase'den temizle."""
        if okunan_fis_nolar is None:
            okunan_fis_nolar = set()

        active_sira = self._fetch_active_sevk_evrakno_sira()

        from datetime import timedelta
        min_tarih = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
        supabase_sira = self.supabase_client.get_sevk_all_evrakno_sira(min_tarih=min_tarih)

        to_delete_candidates = supabase_sira - active_sira
        to_delete = [s for s in to_delete_candidates if s not in okunan_fis_nolar]

        korunan = len(to_delete_candidates) - len(to_delete)
        if korunan > 0:
            logger.info(f"Sevk: {korunan} iptal edilmis fis okumasi oldugu icin korunuyor")

        if to_delete:
            logger.info(f"Sevk: Iptal edilmis {len(to_delete)} fis silinecek: {to_delete}")
            self.supabase_client.delete_sevk_by_evrakno_sira_list(to_delete)

        return len(to_delete)

    def _fetch_active_sevk_evrakno_sira(self) -> set:
        """Mikro'daki son 7 gunun aktif sevk fisi evrakno_sira degerlerini al"""
        server = os.getenv('SQL_SERVER', '')
        database = os.getenv('SQL_DATABASE', '')
        username = os.getenv('SQL_USERNAME', '')
        password = os.getenv('SQL_PASSWORD', '')

        conn_str = (
            f'DRIVER={{ODBC Driver 17 for SQL Server}};'
            f'SERVER={server};'
            f'DATABASE={database};'
            f'UID={username};'
            f'PWD={password}'
        )

        from datetime import timedelta
        min_tarih = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')

        query = """
            SELECT DISTINCT sth_evrakno_sira
            FROM dbo.STOK_HAREKETLERI WITH (NOLOCK)
            WHERE sth_evraktip = 2 AND sth_belge_tarih >= ?
        """

        with pyodbc.connect(conn_str, timeout=30) as conn:
            cursor = conn.cursor()
            cursor.execute(query, min_tarih)
            rows = cursor.fetchall()

        logger.info(f"Sevk: Mikro'da son 7 gunde {len(rows)} aktif sevk evrakno_sira bulundu")
        return set(row[0] for row in rows)

    def cancel(self):
        self.is_cancelled = True


# ================== SATIS / TESLIMAT FISI WIDGET ==================
class SatisTeslimatWidget(QWidget):
    """Satis / Teslimat Fisi sekmesi - Mikro satis faturalarini Supabase'e senkronize etme"""

    def __init__(self):
        super().__init__()
        self._data_loaded = False
        self.sync_thread = None
        self.last_sync_time = None

        # Clients
        self.supabase_client = None
        self.dogtas_client = None
        self._init_clients()

        # Data
        self.all_data = []
        self.filtered_data = []
        # readings_map: kalem_id -> {paket_sira: [{'type','user','time'}, ...]}
        # Her paket_sira icin okuma listesi, uzunluk = okunan birim sayisi
        self.readings_map = {}

        # UI
        self.setup_ui()
        self.setup_connections()

    def showEvent(self, event):
        super().showEvent(event)
        if not self._data_loaded:
            self._data_loaded = True
            QTimer.singleShot(100, self.load_invoice_table)

    def _init_clients(self):
        try:
            config_manager = CentralConfigManager()
            settings = config_manager.get_settings()

            # Supabase credentials (Barkod_SUPABASE_URL veya SUPABASE_URL)
            supabase_url = (settings.get('Barkod_SUPABASE_URL') or
                            settings.get('SUPABASE_URL', ''))
            supabase_key = (settings.get('Barkod_SUPABASE_ANON_KEY') or
                            settings.get('SUPABASE_ANON_KEY', ''))

            # Cache'de yoksa Google Sheets'ten taze cek
            if not supabase_url or not supabase_key:
                logger.info("Supabase ayarlari cache'de yok, Google Sheets'ten taze cekilecek...")
                config_manager.settings_cache = {}
                settings = config_manager.get_settings(use_cache=False)
                supabase_url = (settings.get('Barkod_SUPABASE_URL') or
                                settings.get('SUPABASE_URL', ''))
                supabase_key = (settings.get('Barkod_SUPABASE_ANON_KEY') or
                                settings.get('SUPABASE_ANON_KEY', ''))

            if supabase_url and supabase_key:
                self.supabase_client = SupabaseClient(supabase_url, supabase_key)
                logger.info("Supabase client basariyla olusturuldu")
            else:
                logger.warning("Supabase ayarlari eksik (SUPABASE_URL / SUPABASE_ANON_KEY)")

            # Dogtas API config (Global scope keys)
            dogtas_config = {
                'base_url': settings.get('base_url', ''),
                'userName': settings.get('userName', ''),
                'password': settings.get('password', ''),
                'clientId': settings.get('clientId', ''),
                'clientSecret': settings.get('clientSecret', ''),
                'applicationCode': settings.get('applicationCode', ''),
                'CustomerNo': settings.get('CustomerNo', ''),
            }

            if dogtas_config['base_url']:
                self.dogtas_client = DogtasApiClient(dogtas_config)
                logger.info("Dogtas API client basariyla olusturuldu")
            else:
                logger.warning("Dogtas API ayarlari eksik (base_url)")

        except Exception as e:
            logger.error(f"Client initialization hatasi: {e}")

    # ==================== UI SETUP ====================
    def setup_ui(self):
        self.setStyleSheet("QWidget { background-color: #ffffff; }")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        # Header
        layout.addWidget(self._create_header())

        # Filter bar
        layout.addWidget(self._create_filter_bar())

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setMaximumHeight(20)
        layout.addWidget(self.progress_bar)

        # Table
        self.table = QTableWidget()
        self.table.setStyleSheet(TABLE_STYLE)
        self.table.setItemDelegate(NoFocusDelegate(self.table))
        self.table.setSelectionBehavior(QTableWidget.SelectItems)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.setFocusPolicy(Qt.StrongFocus)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        # Ctrl+C kısayolu - self üzerine bağlanır ki focus nerede olursa olsun çalışsın
        self.copy_shortcut = QShortcut(QKeySequence("Ctrl+C"), self)
        self.copy_shortcut.setContext(Qt.WidgetWithChildrenShortcut)
        self.copy_shortcut.activated.connect(self.handle_ctrl_c)
        layout.addWidget(self.table, 3)

        # Log area
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(120)
        self.log_text.setStyleSheet(LOG_STYLE)
        layout.addWidget(self.log_text, 1)

        # Status bar
        self.status_label = QLabel("Hazir")
        self.status_label.setStyleSheet("QLabel { color: #6b7280; font-size: 12px; padding: 4px; }")
        layout.addWidget(self.status_label)

    def _create_header(self) -> QWidget:
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 0)

        # Sync button
        self.sync_button = QPushButton(u"Sat\u0131\u015f / Teslimat Fi\u015fi Aktar")
        self.sync_button.setStyleSheet(SYNC_BUTTON_STYLE)

        # Refresh button
        self.refresh_button = QPushButton("Yenile")
        self.refresh_button.setStyleSheet(BUTTON_STYLE)

        # All button
        self.all_button = QPushButton("Hepsi")
        self.all_button.setStyleSheet(BUTTON_STYLE)

        # Excel button
        self.export_button = QPushButton("Excel")
        self.export_button.setStyleSheet(BUTTON_STYLE)

        # Checkbox buttons
        self.btn_tumunu = QPushButton(u"T\u00fcm\u00fc")
        self.btn_tumunu.setStyleSheet(BUTTON_STYLE)

        self.btn_sil = QPushButton(u"Se\u00e7ilenleri Sil")
        self.btn_sil.setStyleSheet(BUTTON_STYLE)

        # Info labels
        self.last_sync_label = QLabel("Son Sync: -")
        self.last_sync_label.setStyleSheet(INFO_LABEL_STYLE)

        header_layout.addWidget(self.sync_button)
        header_layout.addWidget(self.refresh_button)
        header_layout.addWidget(self.all_button)
        header_layout.addWidget(self.export_button)
        header_layout.addStretch()
        header_layout.addWidget(self.last_sync_label)

        header_widget = QWidget()
        header_widget.setLayout(header_layout)
        return header_widget

    def _create_filter_bar(self) -> QWidget:
        filter_layout = QHBoxLayout()
        filter_layout.setContentsMargins(0, 0, 0, 0)

        self.filter_evrak = QLineEdit()
        self.filter_evrak.setPlaceholderText("Evrak No")
        self.filter_evrak.setStyleSheet(FILTER_INPUT_STYLE)

        self.filter_stok = QLineEdit()
        self.filter_stok.setPlaceholderText("Stok Kodu")
        self.filter_stok.setStyleSheet(FILTER_INPUT_STYLE)

        self.filter_cari = QLineEdit()
        self.filter_cari.setPlaceholderText("Cari Adi")
        self.filter_cari.setStyleSheet(FILTER_INPUT_STYLE)

        self.filter_malzeme = QLineEdit()
        self.filter_malzeme.setPlaceholderText("Urun Aciklama")
        self.filter_malzeme.setStyleSheet(FILTER_INPUT_STYLE)

        self.filter_tarih = QLineEdit()
        self.filter_tarih.setPlaceholderText("Tarihten itibaren (YYYY-MM-DD)")
        self.filter_tarih.setStyleSheet(FILTER_INPUT_STYLE)

        depo_color = '#3b82f6'
        self.btn_depo_100 = QPushButton("DEPO")
        self.btn_depo_100.setCheckable(True)
        self.btn_depo_100.setStyleSheet(_toggle_btn_style(depo_color, False))

        self.btn_depo_200 = QPushButton(u"\u015eUBE")
        self.btn_depo_200.setCheckable(True)
        self.btn_depo_200.setStyleSheet(_toggle_btn_style(depo_color, False))

        self.btn_depo_300 = QPushButton("EXC")
        self.btn_depo_300.setCheckable(True)
        self.btn_depo_300.setStyleSheet(_toggle_btn_style(depo_color, False))

        self.btn_barkod = QPushButton("Barkod Okunan")
        self.btn_barkod.setCheckable(True)
        self.btn_barkod.setStyleSheet(_toggle_btn_style('#22c55e', False))

        self.btn_manuel_100 = QPushButton("Manuel (100)")
        self.btn_manuel_100.setCheckable(True)
        self.btn_manuel_100.setStyleSheet(_toggle_btn_style('#ef4444', False))

        self.btn_manuel_diger = QPushButton(u"Manuel (Di\u011fer)")
        self.btn_manuel_diger.setCheckable(True)
        self.btn_manuel_diger.setStyleSheet(_toggle_btn_style('#f97316', False))

        self.filter_clear_btn = QPushButton("Temizle")
        self.filter_clear_btn.setStyleSheet(BUTTON_STYLE)

        filter_layout.addWidget(self.btn_tumunu)
        filter_layout.addWidget(self.btn_sil)
        filter_layout.addWidget(self.filter_evrak)
        filter_layout.addWidget(self.filter_stok)
        filter_layout.addWidget(self.filter_cari)
        filter_layout.addWidget(self.filter_malzeme)
        filter_layout.addWidget(self.filter_tarih)
        filter_layout.addWidget(self.btn_depo_100)
        filter_layout.addWidget(self.btn_depo_200)
        filter_layout.addWidget(self.btn_depo_300)
        filter_layout.addWidget(self.btn_barkod)
        filter_layout.addWidget(self.btn_manuel_100)
        filter_layout.addWidget(self.btn_manuel_diger)
        filter_layout.addWidget(self.filter_clear_btn)

        filter_widget = QWidget()
        filter_widget.setLayout(filter_layout)
        return filter_widget

    # ==================== CONNECTIONS ====================
    def setup_connections(self):
        self.sync_button.clicked.connect(self.start_sync)
        self.refresh_button.clicked.connect(self.load_invoice_table)
        self.all_button.clicked.connect(self.load_all_invoice_table)
        self.export_button.clicked.connect(self.export_to_excel)
        self.btn_tumunu.clicked.connect(self._toggle_select_all_rows)
        self.btn_sil.clicked.connect(self._delete_selected_rows)
        self.filter_clear_btn.clicked.connect(self._clear_filters)

        # Debounced filter
        self.filter_timer = QTimer()
        self.filter_timer.setSingleShot(True)
        self.filter_timer.timeout.connect(self.apply_filters)

        for f in [self.filter_evrak, self.filter_cari, self.filter_stok,
                   self.filter_malzeme, self.filter_tarih]:
            f.textChanged.connect(self._schedule_filter)

        self.btn_depo_100.clicked.connect(self._on_toggle_btn)
        self.btn_depo_200.clicked.connect(self._on_toggle_btn)
        self.btn_depo_300.clicked.connect(self._on_toggle_btn)
        self.btn_barkod.clicked.connect(self._on_toggle_btn)
        self.btn_manuel_100.clicked.connect(self._on_toggle_btn)
        self.btn_manuel_diger.clicked.connect(self._on_toggle_btn)

    def _schedule_filter(self):
        self.filter_timer.start(300)

    def _on_toggle_btn(self):
        """Toggle buton tiklandiginda stilini guncelle ve filtrele."""
        btn = self.sender()
        color_map = {
            self.btn_depo_100: '#3b82f6',
            self.btn_depo_200: '#3b82f6',
            self.btn_depo_300: '#3b82f6',
            self.btn_barkod: '#22c55e',
            self.btn_manuel_100: '#ef4444',
            self.btn_manuel_diger: '#f97316',
        }
        color = color_map.get(btn, '#666')
        btn.setStyleSheet(_toggle_btn_style(color, btn.isChecked()))
        self._schedule_filter()

    def _clear_filters(self):
        for f in [self.filter_evrak, self.filter_cari, self.filter_stok,
                   self.filter_malzeme, self.filter_tarih]:
            f.clear()
        for btn, color in [(self.btn_depo_100, '#3b82f6'),
                           (self.btn_depo_200, '#3b82f6'),
                           (self.btn_depo_300, '#3b82f6'),
                           (self.btn_barkod, '#22c55e'),
                           (self.btn_manuel_100, '#ef4444'),
                           (self.btn_manuel_diger, '#f97316')]:
            btn.setChecked(False)
            btn.setStyleSheet(_toggle_btn_style(color, False))
        self.apply_filters()

    # ==================== SYNC ====================
    def start_sync(self):
        if not self.supabase_client:
            self.log("HATA: Supabase baglantisi yapilandirilmamis!")
            self.log("PRGsheet/Ayar sayfasina SUPABASE_URL ve SUPABASE_ANON_KEY ekleyin.")
            self.status_label.setText("Supabase ayarlari eksik")
            return

        if self.sync_thread and self.sync_thread.isRunning():
            self.log("Sync zaten calisiyor...")
            return

        self.log("Senkronizasyon basladi...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self._set_buttons_enabled(False)

        self.sync_thread = SyncThread(self.supabase_client, self.dogtas_client)
        self.sync_thread.progress_updated.connect(self._on_sync_progress)
        self.sync_thread.sync_finished.connect(self._on_sync_finished)
        self.sync_thread.error_occurred.connect(self._on_sync_error)
        self.sync_thread.finished.connect(self._on_thread_finished)
        self.sync_thread.start()

    def _on_sync_progress(self, progress, message):
        self.progress_bar.setValue(progress)
        self.status_label.setText(message)
        self.log(message)

    def _on_sync_finished(self, result):
        self.last_sync_time = datetime.now()
        self.last_sync_label.setText(
            f"Son Sync: {self.last_sync_time.strftime('%Y-%m-%d %H:%M:%S')}"
        )
        mesaj = result['mesaj']
        if result.get('atlanan', 0) > 0:
            mesaj += f" ({result['atlanan']} atlandi)"
        self.log(f"Tamamlandi: {mesaj}")
        self.status_label.setText(mesaj)

        # Tabloyu yenile
        QTimer.singleShot(500, self.load_invoice_table)

    def _on_sync_error(self, error_message):
        self.log(f"HATA: {error_message}")
        self.status_label.setText(f"Sync hatasi: {error_message}")

    def _on_thread_finished(self):
        self._set_buttons_enabled(True)
        QTimer.singleShot(1000, lambda: self.progress_bar.setVisible(False))

    def _set_buttons_enabled(self, enabled: bool):
        self.sync_button.setEnabled(enabled)
        self.refresh_button.setEnabled(enabled)
        self.all_button.setEnabled(enabled)
        self.export_button.setEnabled(enabled)
        self.btn_tumunu.setEnabled(enabled)
        self.btn_sil.setEnabled(enabled)

    # ==================== TABLE ====================
    def load_invoice_table(self):
        if not self.supabase_client:
            self.status_label.setText("Supabase ayarlari eksik")
            return
        from datetime import timedelta
        one_week_ago = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
        self._start_invoice_load(min_tarih=one_week_ago)

    def load_all_invoice_table(self):
        if not self.supabase_client:
            self.status_label.setText("Supabase ayarlari eksik")
            return
        self._start_invoice_load(min_tarih=None)

    def _start_invoice_load(self, min_tarih=None):
        label = "Son 1 hafta" if min_tarih else "Tumu"
        self.status_label.setText(f"Faturalar yukleniyor ({label})...")
        self.refresh_button.setEnabled(False)
        self.all_button.setEnabled(False)

        self._load_thread = LoadDataThread(self.supabase_client, min_tarih=min_tarih)
        self._load_thread.data_loaded.connect(self._on_data_loaded)
        self._load_thread.error_occurred.connect(self._on_load_error)
        self._load_thread.start()

    def _on_data_loaded(self, all_data, readings_map):
        self.all_data = all_data
        self.readings_map = readings_map
        self.apply_filters()
        self.status_label.setText(f"{len(self.all_data)} kayit yuklendi")
        self.refresh_button.setEnabled(True)
        self.all_button.setEnabled(True)

    def _on_load_error(self, error_msg):
        self.status_label.setText(f"Yukleme hatasi: {error_msg}")
        self.log(f"HATA: Tablo yukleme hatasi: {error_msg}")
        self.refresh_button.setEnabled(True)
        self.all_button.setEnabled(True)

    def _get_row_colors(self, row_data):
        """Satirin P etiketlerinde bulunan renkleri dondurur: {'green','red','orange'}"""
        colors = set()
        kalem_id = row_data.get('id')
        depo_no = str(row_data.get('cikis_depo_no', '') or '')
        paket_readings = self.readings_map.get(kalem_id, {})

        for ps, reads in paket_readings.items():
            for info in reads:
                if info['type'] == 'scanner':
                    colors.add('green')
                elif info['type'] == 'manual':
                    if depo_no == '100':
                        colors.add('red')
                    else:
                        colors.add('orange')
        return colors

    def apply_filters(self):
        filtered = self.all_data[:]

        evrak_text = self.filter_evrak.text().strip()
        cari_text = self.filter_cari.text().strip()
        stok_text = self.filter_stok.text().strip()
        malzeme_text = self.filter_malzeme.text().strip()
        tarih_text = self.filter_tarih.text().strip()

        if evrak_text:
            filtered = [r for r in filtered
                        if evrak_text in str(r.get('evrakno_sira', ''))]
        if stok_text:
            filtered = [r for r in filtered
                        if _fuzzy_match(stok_text, str(r.get('stok_kod', '')))]
        if cari_text:
            filtered = [r for r in filtered
                        if _fuzzy_match(cari_text, str(r.get('cari_adi', '')))]
        if malzeme_text:
            filtered = [r for r in filtered
                        if _fuzzy_match(malzeme_text, str(r.get('product_desc', '')))
                        or _fuzzy_match(malzeme_text, str(r.get('malzeme_adi', '')))]
        if tarih_text:
            # tarih_text'ten itibaren (>=) filtrele, format: YYYY-MM-DD
            filtered = [r for r in filtered
                        if str(r.get('tarih', '') or '') >= tarih_text]

        # Depo filtresi (toggle butonlar - birden fazla secim mumkun)
        want_depots = set()
        if self.btn_depo_100.isChecked():
            want_depots.add('100')
        if self.btn_depo_200.isChecked():
            want_depots.add('200')
        if self.btn_depo_300.isChecked():
            want_depots.add('300')
        if want_depots:
            filtered = [r for r in filtered
                        if str(r.get('cikis_depo_no', '') or '') in want_depots]

        # Renk filtresi (toggle butonlar)
        want_colors = set()
        if self.btn_barkod.isChecked():
            want_colors.add('green')
        if self.btn_manuel_100.isChecked():
            want_colors.add('red')
        if self.btn_manuel_diger.isChecked():
            want_colors.add('orange')

        if want_colors:
            filtered = [r for r in filtered
                        if self._get_row_colors(r) & want_colors]

        self.filtered_data = filtered
        self.populate_table()

    def populate_table(self):
        self.table.setUpdatesEnabled(False)
        self.table.setSortingEnabled(False)

        try:
            self.table.setRowCount(len(self.filtered_data))
            self.table.setColumnCount(len(TABLE_COLUMNS) + 1)
            self.table.setHorizontalHeaderLabels([u'\u2611'] + [c[1] for c in TABLE_COLUMNS])

            # Okuma Durumu sutun index'i (+1: checkbox kolonu)
            okuma_col_idx = next(
                (j + 1 for j, (k, _) in enumerate(TABLE_COLUMNS) if k == 'okuma_durumu'),
                None
            )

            for i, row_data in enumerate(self.filtered_data):
                # Checkbox sutunu (index 0)
                chk_item = QTableWidgetItem()
                chk_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
                chk_item.setCheckState(Qt.Unchecked)
                chk_item.setData(Qt.UserRole, row_data.get('evrakno_sira'))
                self.table.setItem(i, 0, chk_item)

                for j, (key, _) in enumerate(TABLE_COLUMNS):
                    if key == 'okuma_durumu':
                        continue  # Ayri handle edilecek
                    value = row_data.get(key, '')
                    if value is None:
                        text = ''
                    elif key == 'miktar':
                        num = float(value)
                        text = str(int(num)) if num == int(num) else str(num)
                    else:
                        text = str(value)
                    item = QTableWidgetItem(text)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    font = QFont(FONT_FAMILY, FONT_SIZE)
                    font.setBold(True)
                    item.setFont(font)
                    self.table.setItem(i, j + 1, item)

                # Okuma Durumu sutunu - renkli P1, P2, P3... etiketleri
                if okuma_col_idx is not None:
                    miktar = int(float(row_data.get('miktar', 0) or 0))
                    paket = int(row_data.get('paket_sayisi', 1) or 1)
                    kalem_id = row_data.get('id')
                    paket_readings = self.readings_map.get(kalem_id, {})
                    depo_no = str(row_data.get('cikis_depo_no', '') or '')
                    widget = _build_okuma_durumu_widget(miktar, paket, paket_readings, depo_no)
                    self.table.setCellWidget(i, okuma_col_idx, widget)

            # Header
            header = self.table.horizontalHeader()
            header.setMinimumSectionSize(0)
            header.setSectionResizeMode(QHeaderView.ResizeToContents)
            header.setStretchLastSection(False)
            header.setSectionResizeMode(0, QHeaderView.Fixed)
            self.table.setColumnWidth(0, 30)

            # Urun Aciklama sutunu ResizeToContents
            product_desc_idx = next(
                (j + 1 for j, (k, _) in enumerate(TABLE_COLUMNS) if k == 'product_desc'),
                len(TABLE_COLUMNS) - 1
            )
            header.setSectionResizeMode(product_desc_idx, QHeaderView.ResizeToContents)

            # Okuma Durumu sutunu minimum genislik
            if okuma_col_idx is not None:
                if self.table.columnWidth(okuma_col_idx) < 200:
                    self.table.setColumnWidth(okuma_col_idx, 200)

            for i in range(self.table.rowCount()):
                self.table.setRowHeight(i, ROW_HEIGHT)

        finally:
            self.table.setUpdatesEnabled(True)
            self.table.setSortingEnabled(True)

    # ==================== CHECKBOX ISLEMLERI ====================
    def _toggle_select_all_rows(self):
        any_checked = any(
            self.table.item(i, 0) and self.table.item(i, 0).checkState() == Qt.Checked
            for i in range(self.table.rowCount())
        )
        new_state = Qt.Unchecked if any_checked else Qt.Checked
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 0)
            if item:
                item.setCheckState(new_state)

    def _delete_selected_rows(self):
        if not self.supabase_client:
            _show_message(self, "Hata", "Supabase bağlantısı yok.")
            return
        ids = []
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 0)
            if item and item.checkState() == Qt.Checked:
                id_val = item.data(Qt.UserRole)
                if id_val is not None:
                    ids.append(id_val)
        if not ids:
            _show_message(self, "Bilgi", "Silinecek satır seçilmedi.")
            return
        if not _verify_barkod_delete_password(self):
            return
        if not _confirm_delete(self, f"{len(ids)} satır Supabase'den silinecek. Emin misiniz?"):
            return
        try:
            self.supabase_client.delete_by_evrakno_sira_list(ids)
            self.status_label.setText(f"{len(ids)} sat\u0131r silindi.")
            self.load_invoice_table()
        except Exception as e:
            _show_message(self, "Hata", f"Silme hatası: {e}")

    # ==================== EXPORT ====================
    def export_to_excel(self):
        if not self.filtered_data:
            self.status_label.setText("Disari aktarilacak veri yok")
            return

        try:
            df = pd.DataFrame(self.filtered_data)
            output_path = "D:/GoogleDrive/~ Barkod_Export.xlsx"
            df.to_excel(output_path, index=False, engine='openpyxl')
            self.status_label.setText(f"Excel export: {output_path}")
            self.log(f"Veriler disari aktarildi: {output_path}")
        except Exception as e:
            self.status_label.setText(f"Export hatasi: {e}")
            self.log(f"HATA: Export hatasi: {e}")

    # ==================== KOPYALAMA ====================
    def show_context_menu(self, pos):
        menu = QMenu(self.table)
        menu.setStyleSheet("""
            QMenu {
                background-color: #ffffff;
                border: 1px solid #a0a0a0;
                padding: 4px 0px;
            }
            QMenu::item {
                padding: 6px 24px;
                color: #000000;
                font-size: 14px;
            }
            QMenu::item:selected {
                background-color: #3399ff;
                color: #ffffff;
            }
        """)
        hucre_action = QAction("Kopyala", self)
        hucre_action.triggered.connect(lambda: self.copy_cell(pos))
        menu.addAction(hucre_action)
        menu.exec_(self.table.viewport().mapToGlobal(pos))

    def copy_cell(self, pos):
        item = self.table.itemAt(pos)
        if item:
            QApplication.clipboard().setText(item.text())
            old_text = self.status_label.text()
            self.status_label.setText("✅ Kopyalandı")
            QTimer.singleShot(1500, lambda t=old_text: self.status_label.setText(t))

    def copy_selection(self):
        selected = self.table.selectedItems()
        if not selected:
            return
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
        old_text = self.status_label.text()
        self.status_label.setText("✅ Kopyalandı")
        QTimer.singleShot(1500, lambda t=old_text: self.status_label.setText(t))

    def handle_ctrl_c(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            return
        if len(selected_items) == 1:
            QApplication.clipboard().setText(selected_items[0].text())
        else:
            rows = sorted({item.row() for item in selected_items})
            cols = sorted({item.column() for item in selected_items})
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
        old_text = self.status_label.text()
        self.status_label.setText("✅ Kopyalandı")
        QTimer.singleShot(1500, lambda t=old_text: self.status_label.setText(t))

    # ==================== LOG ====================
    def log(self, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        # Max 100 satir
        doc = self.log_text.document()
        if doc.blockCount() > 100:
            cursor = self.log_text.textCursor()
            cursor.movePosition(cursor.MoveOperation.Start)
            cursor.movePosition(cursor.MoveOperation.Down, cursor.MoveMode.KeepAnchor,
                                doc.blockCount() - 100)
            cursor.removeSelectedText()


# ================== FABRIKA NAKLIYE PLAN WIDGET ==================
class FabrikaNakliyePlanWidget(QWidget):
    u"""Fabrika Nakliye Plan\u0131 sekmesi - PRGsheet Bekleyenler + Supabase nakliye eslestirmesi"""

    def __init__(self):
        super().__init__()
        self._data_loaded = False
        self.dogtas_client = None
        self._init_clients()
        self.all_data = []
        self.filtered_data = []
        self.column_names = []
        self.nakliye_map = {}
        self.kalem_no_col = None
        self.setup_ui()
        self.setup_connections()

    def showEvent(self, event):
        super().showEvent(event)
        if not self._data_loaded:
            self._data_loaded = True
            QTimer.singleShot(100, self.load_data)

    def _init_clients(self):
        try:
            config_manager = CentralConfigManager()
            settings = config_manager.get_settings()
            dogtas_config = {
                'base_url': settings.get('base_url', ''),
                'nakliye': settings.get('nakliye', ''),
                'userName': settings.get('userName', ''),
                'password': settings.get('password', ''),
                'clientId': settings.get('clientId', ''),
                'clientSecret': settings.get('clientSecret', ''),
                'applicationCode': settings.get('applicationCode', ''),
                'CustomerNo': settings.get('CustomerNo', ''),
            }
            if dogtas_config['base_url'] and dogtas_config['nakliye']:
                self.dogtas_client = DogtasApiClient(dogtas_config)
            else:
                logger.warning("FabrikaNakliye: Dogtas API ayarlari eksik (base_url veya nakliye endpoint)")
        except Exception as e:
            logger.error(f"FabrikaNakliye client init hatasi: {e}")

    def setup_ui(self):
        self.setStyleSheet("QWidget { background-color: #ffffff; }")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        layout.addWidget(self._create_header())
        layout.addWidget(self._create_filter_bar())

        self.table = QTableWidget()
        self.table.setStyleSheet(TABLE_STYLE)
        self.table.setItemDelegate(NoFocusDelegate(self.table))
        self.table.setSelectionBehavior(QTableWidget.SelectItems)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.setFocusPolicy(Qt.StrongFocus)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        self.copy_shortcut = QShortcut(QKeySequence("Ctrl+C"), self)
        self.copy_shortcut.setContext(Qt.WidgetWithChildrenShortcut)
        self.copy_shortcut.activated.connect(self.handle_ctrl_c)
        layout.addWidget(self.table, 3)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(100)
        self.log_text.setStyleSheet(LOG_STYLE)
        layout.addWidget(self.log_text, 1)

        self.status_label = QLabel("Hazir")
        self.status_label.setStyleSheet("QLabel { color: #6b7280; font-size: 12px; padding: 4px; }")
        layout.addWidget(self.status_label)

    def _create_header(self) -> QWidget:
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 0)

        _header_btn_style = """
            QPushButton {
                background-color: #dfdfdf; color: black;
                border: 1px solid #444; padding: 8px 16px;
                border-radius: 5px; font-size: 14px; font-weight: bold;
                min-width: 110px; max-width: 110px;
            }
            QPushButton:hover { background-color: #a0a5a2; }
            QPushButton:disabled { background-color: #f0f0f0; color: #999; }
        """
        self.refresh_button = QPushButton("Yenile")
        self.refresh_button.setStyleSheet(_header_btn_style)

        self.export_button = QPushButton("Excel")
        self.export_button.setStyleSheet(_header_btn_style)

        self.mail_btn = QPushButton("Mail Gönder")
        self.mail_btn.setStyleSheet(_header_btn_style)

        self.last_load_label = QLabel("Son Yukleme: -")
        self.last_load_label.setStyleSheet(INFO_LABEL_STYLE)

        header_layout.addWidget(self.refresh_button)
        header_layout.addWidget(self.export_button)
        header_layout.addWidget(self.mail_btn)
        header_layout.addStretch()
        header_layout.addWidget(self.last_load_label)

        header_widget = QWidget()
        header_widget.setLayout(header_layout)
        return header_widget

    def _create_filter_bar(self) -> QWidget:
        filter_layout = QHBoxLayout()
        filter_layout.setContentsMargins(0, 0, 0, 0)

        self.btn_select_all = QPushButton(u"T\u00fcm\u00fc")
        self.btn_select_all.setCheckable(True)
        self.btn_select_all.setStyleSheet(_toggle_btn_style('#6366f1', False))

        self.filter_search = QLineEdit()
        self.filter_search.setPlaceholderText(u"Ara (t\u00fcm s\u00fctunlarda)")
        self.filter_search.setStyleSheet(FILTER_INPUT_STYLE)

        self.filter_kalem = QLineEdit()
        self.filter_kalem.setPlaceholderText("Kalem No")
        self.filter_kalem.setStyleSheet(FILTER_INPUT_STYLE)

        # Dogtas API tarih araligi - varsayilan: 7 gun once - bugun
        default_start = (datetime.now() - timedelta(days=7)).strftime('%d.%m.%Y')
        default_end = datetime.now().strftime('%d.%m.%Y')

        self.filter_tarih_start = QLineEdit()
        self.filter_tarih_start.setPlaceholderText("Bas. Tarih (GG.AA.YYYY)")
        self.filter_tarih_start.setText(default_start)
        self.filter_tarih_start.setStyleSheet(FILTER_INPUT_STYLE)
        self.filter_tarih_start.setMaximumWidth(150)

        self.filter_tarih_end = QLineEdit()
        self.filter_tarih_end.setPlaceholderText("Bit. Tarih (GG.AA.YYYY)")
        self.filter_tarih_end.setText(default_end)
        self.filter_tarih_end.setStyleSheet(FILTER_INPUT_STYLE)
        self.filter_tarih_end.setMaximumWidth(150)

        # Depo Yeri Plaka filtresi: 002/2=Biga, 200=Inegol
        self.btn_biga = QPushButton(u"B\u0130GA")
        self.btn_biga.setCheckable(True)
        self.btn_biga.setStyleSheet(_toggle_btn_style('#f97316', False))

        self.btn_inegol = QPushButton(u"\u0130NEGO\u0308L")
        self.btn_inegol.setCheckable(True)
        self.btn_inegol.setStyleSheet(_toggle_btn_style('#8b5cf6', False))

        self.filter_clear_btn = QPushButton("Temizle")
        self.filter_clear_btn.setStyleSheet(BUTTON_STYLE)

        filter_layout.addWidget(self.btn_select_all)
        filter_layout.addWidget(self.filter_search)
        filter_layout.addWidget(self.filter_kalem)
        filter_layout.addWidget(self.filter_tarih_start)
        filter_layout.addWidget(self.filter_tarih_end)
        filter_layout.addWidget(self.btn_biga)
        filter_layout.addWidget(self.btn_inegol)
        filter_layout.addWidget(self.filter_clear_btn)

        filter_widget = QWidget()
        filter_widget.setLayout(filter_layout)
        return filter_widget

    def setup_connections(self):
        self.refresh_button.clicked.connect(self.load_data)
        self.export_button.clicked.connect(self.export_to_excel)
        self.mail_btn.clicked.connect(self._send_mail)
        self.btn_select_all.clicked.connect(self._select_all_rows)
        self.filter_clear_btn.clicked.connect(self._clear_filters)

        self.filter_timer = QTimer()
        self.filter_timer.setSingleShot(True)
        self.filter_timer.timeout.connect(self.apply_filters)

        for f in [self.filter_search, self.filter_kalem]:
            f.textChanged.connect(self._schedule_filter)

        self.btn_biga.clicked.connect(self._on_toggle_btn)
        self.btn_inegol.clicked.connect(self._on_toggle_btn)

    def _schedule_filter(self):
        self.filter_timer.start(300)

    def _on_toggle_btn(self):
        btn = self.sender()
        color_map = {
            self.btn_biga: '#f97316',
            self.btn_inegol: '#8b5cf6',
        }
        color = color_map.get(btn, '#666')
        btn.setStyleSheet(_toggle_btn_style(color, btn.isChecked()))
        # BİGA / İNEGÖL birbirini devre dışı bırakır
        if btn is self.btn_biga and btn.isChecked():
            self.btn_inegol.setChecked(False)
            self.btn_inegol.setStyleSheet(_toggle_btn_style('#8b5cf6', False))
        elif btn is self.btn_inegol and btn.isChecked():
            self.btn_biga.setChecked(False)
            self.btn_biga.setStyleSheet(_toggle_btn_style('#f97316', False))
        self._schedule_filter()

    def _select_all_rows(self):
        """Toggle: aktifse tüm checkable satırları seç, değilse seçimi kaldır."""
        checked = self.btn_select_all.isChecked()
        self.btn_select_all.setStyleSheet(_toggle_btn_style('#6366f1', checked))
        state = Qt.Checked if checked else Qt.Unchecked
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 0)
            if item and (item.flags() & Qt.ItemIsUserCheckable):
                item.setCheckState(state)

    def _clear_filters(self):
        for f in [self.filter_search, self.filter_kalem]:
            f.clear()
        for btn, color in [(self.btn_biga, '#f97316'),
                           (self.btn_inegol, '#8b5cf6'),
                           (self.btn_select_all, '#6366f1')]:
            btn.setChecked(False)
            btn.setStyleSheet(_toggle_btn_style(color, False))
        self.apply_filters()

    def load_data(self):
        if not self.dogtas_client:
            self.status_label.setText(
                u"Dogtas API ayarlari eksik (base_url ve nakliye endpoint gerekli)"
            )
            return
        date_start = self.filter_tarih_start.text().strip() or \
            (datetime.now() - timedelta(days=7)).strftime('%d.%m.%Y')
        date_end = self.filter_tarih_end.text().strip() or \
            datetime.now().strftime('%d.%m.%Y')
        self.status_label.setText(
            u"PRGsheet 'Bekleyenler' ve Dogtas nakliye verileri y\u00fckleniyor..."
        )
        self.refresh_button.setEnabled(False)
        self._load_thread = FabrikaNakliyeLoadDataThread(
            self.dogtas_client, date_start, date_end
        )
        self._load_thread.data_loaded.connect(self._on_data_loaded)
        self._load_thread.error_occurred.connect(self._on_load_error)
        self._load_thread.status_update.connect(self.status_label.setText)
        self._load_thread.start()

    def _on_data_loaded(self, rows, column_names, nakliye_map):
        self.all_data = rows
        self.column_names = column_names
        self.nakliye_map = nakliye_map
        # "Kalem No" sutununu bul
        self.kalem_no_col = next(
            (c for c in column_names if 'kalem' in c.lower() and 'no' in c.lower()), None
        )
        if self.kalem_no_col is None:
            self.kalem_no_col = next(
                (c for c in column_names if 'kalem' in c.lower()), None
            )
        self.apply_filters()
        matched = sum(1 for r in rows if self._get_kalem_no(r) in nakliye_map)
        self.status_label.setText(
            f"{len(rows)} 'Sevke Hazir' kayit \u2014 {matched} nakliyede eslesti"
        )
        self.refresh_button.setEnabled(True)
        self.last_load_label.setText(
            f"Son Yukleme: {datetime.now().strftime('%H:%M:%S')}"
        )

    def _on_load_error(self, error_msg):
        self.status_label.setText(f"Yukleme hatasi: {error_msg}")
        self.log(f"HATA: {error_msg}")
        self.refresh_button.setEnabled(True)

    def _get_kalem_no(self, row):
        if self.kalem_no_col:
            return str(row.get(self.kalem_no_col, '') or '').strip()
        return ''

    @staticmethod
    def _depo_norm(val: str) -> str:
        """Depo Yeri Plaka degerini karsilastirma icin normalize et (bas sifirlari kaldir)."""
        return str(val or '').strip().lstrip('0') or '0'

    def apply_filters(self):
        filtered = self.all_data[:]

        search_text = self.filter_search.text().strip()
        kalem_text = self.filter_kalem.text().strip()

        if search_text:
            filtered = [r for r in filtered
                        if any(_fuzzy_match(search_text, str(v)) for v in r.values())]
        if kalem_text:
            filtered = [r for r in filtered
                        if kalem_text.lower() in self._get_kalem_no(r).lower()]

        # Depo Yeri Plaka filtresi: Biga=002/2, Inegol=200
        want_biga = self.btn_biga.isChecked()
        want_inegol = self.btn_inegol.isChecked()
        if want_biga or want_inegol:
            result = []
            for r in filtered:
                norm = self._depo_norm(r.get(u'Depo Yeri Plaka', ''))
                if want_biga and norm == '2':
                    result.append(r)
                elif want_inegol and norm == '200':
                    result.append(r)
            filtered = result

        # Siralama: Eklenmedi (nakliyede olmayan) once, sonra eski tarih once
        def _sort_key(r):
            has_nakliye = self._get_kalem_no(r) in self.nakliye_map
            tarih = str(r.get(u'Sipari\u015f Tarihi', '') or '').strip()
            for sep in (' ', 'T'):
                if sep in tarih:
                    tarih = tarih.split(sep)[0]
                    break
            return (1 if has_nakliye else 0, tarih)

        filtered.sort(key=_sort_key)

        self.filtered_data = filtered
        self.populate_table()

    # Once gosterilecek sabit sutunlar (siraya gore)
    _FIXED_ORDER = [
        u'Sipari\u015f Tarihi',
        u'Prosap S\u00f6zle\u015fme Ad Soyad',
        u'\u00dcr\u00fcn Ad\u0131',
        u'Bekleyen Adet',
    ]
    # Hic gosterilmeyecek sutunlar
    _HIDDEN_COLS = {u'Teslimat Tarihi', u'Spec Ad\u0131', u'KDV(%)'}

    @staticmethod
    def _format_date(val: str) -> str:
        """Tarih degerinden saat kismini kaldir."""
        s = str(val or '').strip()
        for sep in (' ', 'T'):
            if sep in s:
                s = s.split(sep)[0]
                break
        return s

    def _build_active_cols(self):
        """
        Gosterilecek sutun listesini olustur:
        1. Siparis Tarihi (varsa)
        2. Nakliye Durumu (hesaplanan, her zaman)
        3. Diger sabit sutunlar (varsa)
        4. Kalan sheet sutunlari (_FIXED_ORDER ve _HIDDEN_COLS disindakiler)
        Her eleman: (sheet_col_or_None, header_str)
        """
        fixed_set = set(self._FIXED_ORDER)
        result = []

        # 1. Siparis Tarihi
        if u'Sipari\u015f Tarihi' in self.column_names:
            result.append((u'Sipari\u015f Tarihi', u'Sipari\u015f Tarihi'))

        # 2. Nakliye Durumu — hesaplanan
        result.append((None, u'Nakliye Durumu'))

        # 3. Diger sabit sutunlar (Siparis Tarihi haric)
        for col in self._FIXED_ORDER[1:]:
            if col in self.column_names:
                result.append((col, col))

        # 4. Kalan sheet sutunlari (sabit veya gizli olmayanlar)
        for col in self.column_names:
            if col not in fixed_set and col not in self._HIDDEN_COLS:
                result.append((col, col))

        return result

    def populate_table(self):
        if not self.column_names:
            return

        active_cols = self._build_active_cols()
        today = datetime.now().date()
        depo_filtre_aktif = self.btn_biga.isChecked() or self.btn_inegol.isChecked()

        self.table.setUpdatesEnabled(False)
        self.table.setSortingEnabled(False)
        try:
            self.table.setRowCount(len(self.filtered_data))
            # +1: ilk sutun checkbox
            self.table.setColumnCount(len(active_cols) + 1)
            self.table.setHorizontalHeaderLabels([u'☑'] + [h for (_, h) in active_cols])

            for i, row_data in enumerate(self.filtered_data):
                kalem_no = self._get_kalem_no(row_data)
                nakliye_records = self.nakliye_map.get(kalem_no, [])
                has_nakliye = bool(nakliye_records)

                # Sutun 0: "Eklenmedi" satirlara checkbox; nakliyesi olan satirlara koyma
                chk_item = QTableWidgetItem()
                if not has_nakliye:
                    chk_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
                    # Depo filtresi aktifse ve tarih 27 gundan eskiyse otomatik isaretli
                    auto_check = False
                    if depo_filtre_aktif:
                        tarih_str = self._format_date(str(row_data.get('Sipariş Tarihi', '') or ''))
                        if tarih_str:
                            try:
                                row_date = datetime.strptime(tarih_str, '%Y-%m-%d').date()
                                auto_check = (today - row_date).days > 27
                            except Exception:
                                pass
                    chk_item.setCheckState(Qt.Checked if auto_check else Qt.Unchecked)
                else:
                    chk_item.setFlags(Qt.ItemIsEnabled)  # nakliyeli satir, checkbox yok
                self.table.setItem(i, 0, chk_item)

                for j, (sheet_col, _) in enumerate(active_cols):
                    font = QFont(FONT_FAMILY, FONT_SIZE)
                    font.setBold(True)

                    if sheet_col is None:
                        # Nakliye Durumu
                        if has_nakliye:
                            nakliye_nos = list(dict.fromkeys(
                                str(r.get('nakliye_no', '')) for r in nakliye_records
                                if r.get('nakliye_no')
                            ))
                            nos_str = ', '.join(nakliye_nos[:3])
                            if len(nakliye_nos) > 3:
                                nos_str += f' +{len(nakliye_nos) - 3}'
                            text = u'\u2705 ' + nos_str
                            cell_color = QColor('#dcfce7')
                        else:
                            text = u'Eklenmedi'
                            cell_color = QColor('#fef9c3')
                        item = QTableWidgetItem(text)
                        item.setBackground(QBrush(cell_color))
                    else:
                        raw = str(row_data.get(sheet_col, '') or '')
                        if sheet_col == u'Sipari\u015f Tarihi':
                            raw = self._format_date(raw)
                        item = QTableWidgetItem(raw)

                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    item.setFont(font)
                    self.table.setItem(i, j + 1, item)  # +1 checkbox offset

                self.table.setRowHeight(i, ROW_HEIGHT)

            header = self.table.horizontalHeader()
            header.setMinimumSectionSize(0)
            header.setSectionResizeMode(QHeaderView.ResizeToContents)
            header.setSectionResizeMode(0, QHeaderView.Fixed)
            self.table.setColumnWidth(0, 30)

        finally:
            self.table.setUpdatesEnabled(True)
            # setSortingEnabled(True) kasitli kullanilmiyor:
            # header tiklamayla apply_filters siralamasi ezilmesin

    def export_to_excel(self):
        if not self.filtered_data:
            self.status_label.setText(u"D\u0131\u015far\u0131 aktar\u0131lacak veri yok")
            return
        try:
            active_cols = self._build_active_cols()
            rows_for_export = []
            for row_data in self.filtered_data:
                kalem_no = self._get_kalem_no(row_data)
                nakliye_records = self.nakliye_map.get(kalem_no, [])
                export_row = {}
                for sheet_col, header in active_cols:
                    if sheet_col is None:
                        export_row[header] = ', '.join(
                            str(r.get('nakliye_no', '')) for r in nakliye_records[:3]
                            if r.get('nakliye_no')
                        ) if nakliye_records else 'Eklenmedi'
                    elif sheet_col == u'Sipari\u015f Tarihi':
                        export_row[header] = self._format_date(row_data.get(sheet_col, ''))
                    else:
                        export_row[header] = str(row_data.get(sheet_col, '') or '')
                rows_for_export.append(export_row)
            col_order = [h for (_, h) in active_cols]
            df = pd.DataFrame(rows_for_export, columns=col_order)
            output_path = "D:/GoogleDrive/~ FabrikaNakliye_Export.xlsx"
            df.to_excel(output_path, index=False, engine='openpyxl')
            self.status_label.setText(f"Excel export: {output_path}")
            self.log(f"Veriler disari aktarildi: {output_path}")
        except Exception as e:
            self.status_label.setText(f"Export hatasi: {e}")
            self.log(f"HATA: Export hatasi: {e}")

    def show_context_menu(self, pos):
        menu = QMenu(self.table)
        menu.setStyleSheet("""
            QMenu {
                background-color: #ffffff;
                border: 1px solid #a0a0a0;
                padding: 4px 0px;
            }
            QMenu::item {
                padding: 6px 24px;
                color: #000000;
                font-size: 14px;
            }
            QMenu::item:selected {
                background-color: #3399ff;
                color: #ffffff;
            }
        """)
        hucre_action = QAction("Kopyala", self)
        hucre_action.triggered.connect(lambda: self.copy_cell(pos))
        menu.addAction(hucre_action)
        menu.exec_(self.table.viewport().mapToGlobal(pos))

    def copy_cell(self, pos):
        item = self.table.itemAt(pos)
        if item:
            QApplication.clipboard().setText(item.text())
            old_text = self.status_label.text()
            self.status_label.setText(u'\u2705 Kopyaland\u0131')
            QTimer.singleShot(1500, lambda t=old_text: self.status_label.setText(t))

    def handle_ctrl_c(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            return
        if len(selected_items) == 1:
            QApplication.clipboard().setText(selected_items[0].text())
        else:
            rows = sorted({item.row() for item in selected_items})
            cols = sorted({item.column() for item in selected_items})
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
        old_text = self.status_label.text()
        self.status_label.setText(u'\u2705 Kopyaland\u0131')
        QTimer.singleShot(1500, lambda t=old_text: self.status_label.setText(t))

    def _msgbox(self, kind, title, text, buttons=None):
        """Okunabilir (beyaz zemin, siyah yazi) QMessageBox gosterir."""
        box = QMessageBox(self)
        box.setStyleSheet(
            "QMessageBox { background-color: #ffffff; }"
            "QMessageBox QLabel { color: #000000; font-size: 13px; }"
            "QMessageBox QPushButton { color: #000000; background-color: #f3f4f6;"
            " border: 1px solid #d1d5db; border-radius: 4px;"
            " padding: 4px 12px; min-width: 60px; }"
            "QMessageBox QPushButton:hover { background-color: #e5e7eb; }"
        )
        box.setWindowTitle(title)
        box.setText(text)
        icon_map = {
            'info': QMessageBox.Information,
            'warning': QMessageBox.Warning,
            'critical': QMessageBox.Critical,
            'question': QMessageBox.Question,
        }
        box.setIcon(icon_map.get(kind, QMessageBox.NoIcon))
        if buttons:
            box.setStandardButtons(buttons)
            box.setDefaultButton(QMessageBox.No)
        return box.exec_()

    @staticmethod
    def _format_kalem_no(kalem: str) -> str:
        """'1102678987000410' -> '1102678987-410' (ilk 10 karakter + '-' + kalan bas sifirlar silinmis)"""
        s = str(kalem or '').strip()
        if 'E+' in s or 'e+' in s:
            try:
                s = str(int(float(s)))
            except Exception:
                pass
        if len(s) > 10:
            return s[:10] + '-' + s[10:].lstrip('0')
        return s

    def _send_mail(self):
        """Filtrelenmiş 'Eklenmedi' satırlarını BİGA veya İNEGÖL için mail olarak gönderir."""
        import smtplib
        from email.mime.text import MIMEText
        from email.mime.multipart import MIMEMultipart
        from email.header import Header
        from io import BytesIO

        # Hangi depo secili?
        if self.btn_biga.isChecked():
            depo = "BİGA"
        elif self.btn_inegol.isChecked():
            depo = "İNEGÖL"
        else:
            self._msgbox('warning', "Depo Seçilmedi",
                         "Lütfen önce BİGA veya İNEGÖL butonuna tıklayın.")
            return

        # Secili (checkbox isaretli) satirlari al
        selected_rows = []
        for i in range(self.table.rowCount()):
            chk = self.table.item(i, 0)
            if chk and chk.checkState() == Qt.Checked and i < len(self.filtered_data):
                selected_rows.append(self.filtered_data[i])
        eklenmedi_rows = selected_rows
        if not eklenmedi_rows:
            self._msgbox('info', "Satır Seçilmedi",
                         "Lütfen göndermek istediğiniz satırları işaretleyin.")
            return

        # Mail ayarlarini PRGsheet "Mail" sayfasindan yukle
        try:
            config_manager = CentralConfigManager()
            sid = config_manager.MASTER_SPREADSHEET_ID
            url = f"https://docs.google.com/spreadsheets/d/{sid}/export?format=xlsx"
            resp = requests.get(url, timeout=30)
            resp.raise_for_status()
            mail_df = pd.read_excel(BytesIO(resp.content), sheet_name="Mail")
            mail_sevk_df = mail_df[mail_df['fonksiyon'] == 'mail_sevk_gonder']
            if mail_sevk_df.empty:
                self._msgbox('warning', "Mail Ayarı Bulunamadı",
                             "PRGsheet 'Mail' sayfasında 'mail_sevk_gonder' satırı bulunamadı.")
                return
            mail_info = mail_sevk_df.iloc[0]
            sender_email = str(mail_info["sender_email"])
            receiver_email = str(mail_info["receiver_email"])
            cc_email = str(mail_info["cc_email"]) if pd.notna(mail_info.get("cc_email")) else ""
            bcc_email = str(mail_info["bcc_email"]) if pd.notna(mail_info.get("bcc_email")) else ""
            password = str(mail_info["password"])
            smtp_server = str(mail_info["smtp_server"])
        except Exception as e:
            self._msgbox('critical', "Mail Ayarı Hatası", f"Mail ayarları yüklenemedi: {e}")
            return

        # Mail kolonlari
        mail_cols = [
            'Sipariş Tarihi', 'Kalem No', 'Ürün Adı',
            'Spec Adı', 'Bekleyen Adet', 'Durum',
            'Teslimat Tarihi', 'Depo Yeri Plaka', 'Teslim Deposu',
        ]

        # Tablo satirlarini olustur
        table_rows = []
        for r in eklenmedi_rows:
            row = {}
            for col in mail_cols:
                val = str(r.get(col, '') or '')
                if col == 'Sipariş Tarihi':
                    val = self._format_date(val)
                elif col == 'Kalem No':
                    val = self._format_kalem_no(val)
                row[col] = val
            table_rows.append(row)

        df_mail = pd.DataFrame(table_rows, columns=mail_cols)
        html_table = df_mail.to_html(index=False, border=1)

        subject = f"{depo} BAYİ SEVK"
        body = f"""
        <p>Merhaba,</p>
        <p>Ekteki ürünler sevkiyat planına dahil edilmemiştir. Plana alınması için destek rica ederim.</p>
        {html_table}
        <p>İyi çalışmalar diliyorum.</p>
        """

        msg = MIMEMultipart()
        msg["From"] = str(Header(sender_email, "utf-8"))
        msg["To"] = str(Header(receiver_email, "utf-8"))
        if cc_email:
            msg["Cc"] = str(Header(cc_email, "utf-8"))
        msg["Subject"] = str(Header(subject, "utf-8"))
        msg.attach(MIMEText(body, "html", "utf-8"))

        to_addrs = [receiver_email]
        if cc_email:
            to_addrs.append(cc_email)
        if bcc_email:
            to_addrs.append(bcc_email)

        reply = self._msgbox(
            'question', "E-posta Gönderimi",
            f"{depo} deposu için {len(eklenmedi_rows)} satır gönderilecek.\n"
            f"Alıcı: {receiver_email}\n\nEmin misiniz?",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        try:
            with smtplib.SMTP(smtp_server, 587) as server:
                server.starttls()
                server.login(sender_email, password)
                server.sendmail(sender_email, to_addrs, msg.as_string())
            self._msgbox('info', "E-posta Gönderildi",
                         f"E-posta başarıyla gönderildi.\nKime: {receiver_email}")
            self.status_label.setText("✅ Sevk e-postası başarıyla gönderildi")
            self.log(f"Mail gonderildi: {subject} -> {receiver_email}")
        except Exception as e:
            self._msgbox('critical', "E-posta Gönderme Hatası",
                         f"E-posta gönderilemedi: {e}")
            self.status_label.setText(f"❌ Mail gönderme hatası: {e}")
            self.log(f"HATA: Mail gonderilemedi: {e}")

    def log(self, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        doc = self.log_text.document()
        if doc.blockCount() > 100:
            cursor = self.log_text.textCursor()
            cursor.movePosition(cursor.MoveOperation.Start)
            cursor.movePosition(cursor.MoveOperation.Down, cursor.MoveMode.KeepAnchor,
                                doc.blockCount() - 100)
            cursor.removeSelectedText()


# ================== NAKLIYE YUKLEME WIDGET ==================
class NakliyeYuklemeWidget(QWidget):
    """Nakliye Yukleme sekmesi - Supabase nakliye fislerini ve okuma durumunu gosterir"""

    def __init__(self):
        super().__init__()
        self._data_loaded = False
        self.supabase_client = None
        self._init_clients()
        self.all_data = []
        self.filtered_data = []
        self.readings_map = {}
        self.setup_ui()
        self.setup_connections()

    def showEvent(self, event):
        super().showEvent(event)
        if not self._data_loaded:
            self._data_loaded = True
            QTimer.singleShot(100, self.load_data)

    def _init_clients(self):
        try:
            config_manager = CentralConfigManager()
            settings = config_manager.get_settings()
            supabase_url = (settings.get('Barkod_SUPABASE_URL') or
                            settings.get('SUPABASE_URL', ''))
            supabase_key = (settings.get('Barkod_SUPABASE_ANON_KEY') or
                            settings.get('SUPABASE_ANON_KEY', ''))
            if not supabase_url or not supabase_key:
                config_manager.settings_cache = {}
                settings = config_manager.get_settings(use_cache=False)
                supabase_url = (settings.get('Barkod_SUPABASE_URL') or
                                settings.get('SUPABASE_URL', ''))
                supabase_key = (settings.get('Barkod_SUPABASE_ANON_KEY') or
                                settings.get('SUPABASE_ANON_KEY', ''))
            if supabase_url and supabase_key:
                self.supabase_client = SupabaseClient(supabase_url, supabase_key)
        except Exception as e:
            logger.error(f"Nakliye client init hatasi: {e}")

    def setup_ui(self):
        self.setStyleSheet("QWidget { background-color: #ffffff; }")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        layout.addWidget(self._create_header())
        layout.addWidget(self._create_filter_bar())

        self.table = QTableWidget()
        self.table.setStyleSheet(TABLE_STYLE)
        self.table.setItemDelegate(NoFocusDelegate(self.table))
        self.table.setSelectionBehavior(QTableWidget.SelectItems)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.setFocusPolicy(Qt.StrongFocus)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        # Ctrl+C kısayolu - self üzerine bağlanır ki focus nerede olursa olsun çalışsın
        self.copy_shortcut = QShortcut(QKeySequence("Ctrl+C"), self)
        self.copy_shortcut.setContext(Qt.WidgetWithChildrenShortcut)
        self.copy_shortcut.activated.connect(self.handle_ctrl_c)
        layout.addWidget(self.table, 3)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(120)
        self.log_text.setStyleSheet(LOG_STYLE)
        layout.addWidget(self.log_text, 1)

        self.status_label = QLabel("Hazir")
        self.status_label.setStyleSheet("QLabel { color: #6b7280; font-size: 12px; padding: 4px; }")
        layout.addWidget(self.status_label)

    def _create_header(self) -> QWidget:
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 0)

        self.refresh_button = QPushButton("Yenile")
        self.refresh_button.setStyleSheet(SYNC_BUTTON_STYLE)

        self.all_button = QPushButton("Hepsi")
        self.all_button.setStyleSheet(BUTTON_STYLE)

        self.export_button = QPushButton("Excel")
        self.export_button.setStyleSheet(BUTTON_STYLE)

        self.btn_tumunu = QPushButton(u"T\u00fcm\u00fc")
        self.btn_tumunu.setStyleSheet(BUTTON_STYLE)

        self.btn_sil = QPushButton(u"Se\u00e7ilenleri Sil")
        self.btn_sil.setStyleSheet(BUTTON_STYLE)

        self.last_load_label = QLabel("Son Yukleme: -")
        self.last_load_label.setStyleSheet(INFO_LABEL_STYLE)

        header_layout.addWidget(self.refresh_button)
        header_layout.addWidget(self.all_button)
        header_layout.addWidget(self.export_button)
        header_layout.addStretch()
        header_layout.addWidget(self.last_load_label)

        header_widget = QWidget()
        header_widget.setLayout(header_layout)
        return header_widget

    def _create_filter_bar(self) -> QWidget:
        filter_layout = QHBoxLayout()
        filter_layout.setContentsMargins(0, 0, 0, 0)

        self.filter_oturum = QLineEdit()
        self.filter_oturum.setPlaceholderText("Oturum")
        self.filter_oturum.setStyleSheet(FILTER_INPUT_STYLE)

        self.filter_nakliye = QLineEdit()
        self.filter_nakliye.setPlaceholderText("Nakliye No")
        self.filter_nakliye.setStyleSheet(FILTER_INPUT_STYLE)

        self.filter_malzeme_no = QLineEdit()
        self.filter_malzeme_no.setPlaceholderText("Malzeme No")
        self.filter_malzeme_no.setStyleSheet(FILTER_INPUT_STYLE)

        self.filter_malzeme = QLineEdit()
        self.filter_malzeme.setPlaceholderText("Malzeme")
        self.filter_malzeme.setStyleSheet(FILTER_INPUT_STYLE)

        self.filter_tarih = QLineEdit()
        self.filter_tarih.setPlaceholderText("Tarihten itibaren (YYYY-MM-DD)")
        self.filter_tarih.setStyleSheet(FILTER_INPUT_STYLE)

        self.btn_barkod = QPushButton("Barkod Okunan")
        self.btn_barkod.setCheckable(True)
        self.btn_barkod.setStyleSheet(_toggle_btn_style('#22c55e', False))

        self.btn_manuel = QPushButton("Manuel")
        self.btn_manuel.setCheckable(True)
        self.btn_manuel.setStyleSheet(_toggle_btn_style('#ef4444', False))

        self.filter_clear_btn = QPushButton("Temizle")
        self.filter_clear_btn.setStyleSheet(BUTTON_STYLE)

        filter_layout.addWidget(self.btn_tumunu)
        filter_layout.addWidget(self.btn_sil)
        filter_layout.addWidget(self.filter_oturum)
        filter_layout.addWidget(self.filter_nakliye)
        filter_layout.addWidget(self.filter_malzeme_no)
        filter_layout.addWidget(self.filter_malzeme)
        filter_layout.addWidget(self.filter_tarih)
        filter_layout.addWidget(self.btn_barkod)
        filter_layout.addWidget(self.btn_manuel)
        filter_layout.addWidget(self.filter_clear_btn)

        filter_widget = QWidget()
        filter_widget.setLayout(filter_layout)
        return filter_widget

    def setup_connections(self):
        self.refresh_button.clicked.connect(self.load_data)
        self.all_button.clicked.connect(self.load_all_data)
        self.export_button.clicked.connect(self.export_to_excel)
        self.btn_tumunu.clicked.connect(self._toggle_select_all_rows)
        self.btn_sil.clicked.connect(self._delete_selected_rows)
        self.filter_clear_btn.clicked.connect(self._clear_filters)

        self.filter_timer = QTimer()
        self.filter_timer.setSingleShot(True)
        self.filter_timer.timeout.connect(self.apply_filters)

        for f in [self.filter_oturum, self.filter_nakliye, self.filter_malzeme_no,
                   self.filter_malzeme, self.filter_tarih]:
            f.textChanged.connect(self._schedule_filter)

        self.btn_barkod.clicked.connect(self._on_toggle_btn)
        self.btn_manuel.clicked.connect(self._on_toggle_btn)

    def _schedule_filter(self):
        self.filter_timer.start(300)

    def _on_toggle_btn(self):
        btn = self.sender()
        color_map = {
            self.btn_barkod: '#22c55e',
            self.btn_manuel: '#ef4444',
        }
        color = color_map.get(btn, '#666')
        btn.setStyleSheet(_toggle_btn_style(color, btn.isChecked()))
        self._schedule_filter()

    def _clear_filters(self):
        for f in [self.filter_oturum, self.filter_nakliye, self.filter_malzeme_no,
                   self.filter_malzeme, self.filter_tarih]:
            f.clear()
        for btn, color in [(self.btn_barkod, '#22c55e'),
                           (self.btn_manuel, '#ef4444')]:
            btn.setChecked(False)
            btn.setStyleSheet(_toggle_btn_style(color, False))
        self.apply_filters()

    def load_data(self):
        if not self.supabase_client:
            self.status_label.setText("Supabase ayarlari eksik")
            return
        from datetime import timedelta
        one_week_ago = (datetime.now() - timedelta(days=7)).strftime('%Y%m%d')
        self._start_load(min_tarih=one_week_ago)

    def load_all_data(self):
        if not self.supabase_client:
            self.status_label.setText("Supabase ayarlari eksik")
            return
        self._start_load(min_tarih=None)

    def _start_load(self, min_tarih=None):
        label = "Son 1 hafta" if min_tarih else "Tumu"
        self.status_label.setText(f"Nakliye fisleri yukleniyor ({label})...")
        self.refresh_button.setEnabled(False)
        self.all_button.setEnabled(False)
        self._load_thread = NakliyeLoadDataThread(self.supabase_client, min_tarih=min_tarih)
        self._load_thread.data_loaded.connect(self._on_data_loaded)
        self._load_thread.error_occurred.connect(self._on_load_error)
        self._load_thread.start()

    def _on_data_loaded(self, all_data, readings_map):
        self.all_data = all_data
        self.readings_map = readings_map
        self.apply_filters()
        self.status_label.setText(f"{len(self.all_data)} nakliye kaydi yuklendi")
        self.refresh_button.setEnabled(True)
        self.all_button.setEnabled(True)
        self.last_load_label.setText(
            f"Son Yukleme: {datetime.now().strftime('%H:%M:%S')}"
        )

    def _on_load_error(self, error_msg):
        self.status_label.setText(f"Yukleme hatasi: {error_msg}")
        self.log(f"HATA: {error_msg}")
        self.refresh_button.setEnabled(True)
        self.all_button.setEnabled(True)

    def _get_row_colors(self, row_data):
        colors = set()
        kalem_id = row_data.get('id')
        paket_readings = self.readings_map.get(kalem_id, {})
        for ps, reads in paket_readings.items():
            for info in reads:
                if info['type'] == 'scanner':
                    colors.add('green')
                elif info['type'] == 'manual':
                    colors.add('red')
        return colors

    def apply_filters(self):
        filtered = self.all_data[:]

        oturum_text = self.filter_oturum.text().strip()
        nakliye_text = self.filter_nakliye.text().strip()
        malzeme_no_text = self.filter_malzeme_no.text().strip()
        malzeme_text = self.filter_malzeme.text().strip()
        tarih_text = self.filter_tarih.text().strip()

        if oturum_text:
            filtered = [r for r in filtered
                        if oturum_text in str(r.get('oturum_id', ''))]
        if nakliye_text:
            filtered = [r for r in filtered
                        if nakliye_text in str(r.get('nakliye_no', ''))]
        if malzeme_no_text:
            filtered = [r for r in filtered
                        if malzeme_no_text in str(r.get('malzeme_no', '')).lstrip('0')]
        if malzeme_text:
            filtered = [r for r in filtered
                        if _fuzzy_match(malzeme_text, str(r.get('malzeme_adi', '')))]
        if tarih_text:
            # YYYY-MM-DD -> YYYYMMDD'ye çevir ve >= karşılaştır
            tarih_threshold = tarih_text.replace('-', '')
            filtered = [r for r in filtered
                        if str(r.get('belge_tarihi', '') or '') >= tarih_threshold]

        want_colors = set()
        if self.btn_barkod.isChecked():
            want_colors.add('green')
        if self.btn_manuel.isChecked():
            want_colors.add('red')
        if want_colors:
            filtered = [r for r in filtered
                        if self._get_row_colors(r) & want_colors]

        self.filtered_data = filtered
        self.populate_table()

    def populate_table(self):
        self.table.setUpdatesEnabled(False)
        self.table.setSortingEnabled(False)

        try:
            self.table.setRowCount(len(self.filtered_data))
            self.table.setColumnCount(len(NAKLIYE_TABLE_COLUMNS) + 1)
            self.table.setHorizontalHeaderLabels([u'\u2611'] + [c[1] for c in NAKLIYE_TABLE_COLUMNS])

            okuma_col_idx = next(
                (j + 1 for j, (k, _) in enumerate(NAKLIYE_TABLE_COLUMNS) if k == 'okuma_durumu'),
                None
            )

            for i, row_data in enumerate(self.filtered_data):
                # Checkbox sutunu (index 0)
                chk_item = QTableWidgetItem()
                chk_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
                chk_item.setCheckState(Qt.Unchecked)
                chk_item.setData(Qt.UserRole, row_data.get('id'))
                self.table.setItem(i, 0, chk_item)

                for j, (key, _) in enumerate(NAKLIYE_TABLE_COLUMNS):
                    if key == 'okuma_durumu':
                        continue
                    value = row_data.get(key, '')
                    if value is None:
                        text = ''
                    elif key == 'miktar':
                        num = float(str(value).replace(',', '.'))
                        text = str(int(num)) if num == int(num) else str(num)
                    elif key == 'belge_tarihi':
                        # "20260302" -> "2026-03-02"
                        s = str(value).strip()
                        if len(s) == 8 and s.isdigit():
                            text = f"{s[0:4]}-{s[4:6]}-{s[6:8]}"
                        else:
                            text = s
                    elif key == 'malzeme_no':
                        text = str(value).lstrip('0') or '0'
                    elif key == 'satinalma_kalem_id':
                        text = str(value) if value else ''
                    else:
                        text = str(value)
                    item = QTableWidgetItem(text)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    font = QFont(FONT_FAMILY, FONT_SIZE)
                    font.setBold(True)
                    item.setFont(font)
                    self.table.setItem(i, j + 1, item)

                if okuma_col_idx is not None:
                    miktar_raw = str(row_data.get('miktar', '0') or '0').replace(',', '.')
                    miktar = max(1, int(float(miktar_raw)))
                    paket = int(row_data.get('paket_sayisi', 1) or 1)
                    kalem_id = row_data.get('id')
                    paket_readings = self.readings_map.get(kalem_id, {})
                    widget = _build_okuma_durumu_widget(miktar, paket, paket_readings, depo_no='100')
                    self.table.setCellWidget(i, okuma_col_idx, widget)

            header = self.table.horizontalHeader()
            header.setMinimumSectionSize(0)
            header.setSectionResizeMode(QHeaderView.ResizeToContents)
            header.setStretchLastSection(False)
            header.setSectionResizeMode(0, QHeaderView.Fixed)
            self.table.setColumnWidth(0, 30)

            malzeme_adi_idx = next(
                (j + 1 for j, (k, _) in enumerate(NAKLIYE_TABLE_COLUMNS) if k == 'malzeme_adi'),
                len(NAKLIYE_TABLE_COLUMNS) - 1
            )
            header.setSectionResizeMode(malzeme_adi_idx, QHeaderView.ResizeToContents)

            if okuma_col_idx is not None:
                if self.table.columnWidth(okuma_col_idx) < 200:
                    self.table.setColumnWidth(okuma_col_idx, 200)

            for i in range(self.table.rowCount()):
                self.table.setRowHeight(i, ROW_HEIGHT)

        finally:
            self.table.setUpdatesEnabled(True)
            self.table.setSortingEnabled(True)

    # ==================== CHECKBOX ISLEMLERI ====================
    def _toggle_select_all_rows(self):
        any_checked = any(
            self.table.item(i, 0) and self.table.item(i, 0).checkState() == Qt.Checked
            for i in range(self.table.rowCount())
        )
        new_state = Qt.Unchecked if any_checked else Qt.Checked
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 0)
            if item:
                item.setCheckState(new_state)

    def _delete_selected_rows(self):
        if not self.supabase_client:
            _show_message(self, "Hata", "Supabase bağlantısı yok.")
            return
        ids = []
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 0)
            if item and item.checkState() == Qt.Checked:
                id_val = item.data(Qt.UserRole)
                if id_val is not None:
                    ids.append(id_val)
        if not ids:
            _show_message(self, "Bilgi", "Silinecek satır seçilmedi.")
            return
        if not _verify_barkod_delete_password(self):
            return
        if not _confirm_delete(self, f"{len(ids)} satır Supabase'den silinecek. Emin misiniz?"):
            return
        try:
            self.supabase_client.delete_nakliye_by_id_list(ids)
            self.status_label.setText(f"{len(ids)} sat\u0131r silindi.")
            self.load_data()
        except Exception as e:
            _show_message(self, "Hata", f"Silme hatası: {e}")

    def export_to_excel(self):
        if not self.filtered_data:
            self.status_label.setText("Disari aktarilacak veri yok")
            return
        try:
            df = pd.DataFrame(self.filtered_data)
            output_path = "D:/GoogleDrive/~ Nakliye_Export.xlsx"
            df.to_excel(output_path, index=False, engine='openpyxl')
            self.status_label.setText(f"Excel export: {output_path}")
            self.log(f"Veriler disari aktarildi: {output_path}")
        except Exception as e:
            self.status_label.setText(f"Export hatasi: {e}")
            self.log(f"HATA: Export hatasi: {e}")

    def show_context_menu(self, pos):
        menu = QMenu(self.table)
        menu.setStyleSheet("""
            QMenu {
                background-color: #ffffff;
                border: 1px solid #a0a0a0;
                padding: 4px 0px;
            }
            QMenu::item {
                padding: 6px 24px;
                color: #000000;
                font-size: 14px;
            }
            QMenu::item:selected {
                background-color: #3399ff;
                color: #ffffff;
            }
        """)
        hucre_action = QAction("Kopyala", self)
        hucre_action.triggered.connect(lambda: self.copy_cell(pos))
        menu.addAction(hucre_action)
        menu.exec_(self.table.viewport().mapToGlobal(pos))

    def copy_cell(self, pos):
        item = self.table.itemAt(pos)
        if item:
            QApplication.clipboard().setText(item.text())
            old_text = self.status_label.text()
            self.status_label.setText("✅ Kopyalandı")
            QTimer.singleShot(1500, lambda t=old_text: self.status_label.setText(t))

    def copy_selection(self):
        selected = self.table.selectedItems()
        if not selected:
            return
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
        old_text = self.status_label.text()
        self.status_label.setText("✅ Kopyalandı")
        QTimer.singleShot(1500, lambda t=old_text: self.status_label.setText(t))

    def handle_ctrl_c(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            return
        if len(selected_items) == 1:
            QApplication.clipboard().setText(selected_items[0].text())
        else:
            rows = sorted({item.row() for item in selected_items})
            cols = sorted({item.column() for item in selected_items})
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
        old_text = self.status_label.text()
        self.status_label.setText("✅ Kopyalandı")
        QTimer.singleShot(1500, lambda t=old_text: self.status_label.setText(t))

    def log(self, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        doc = self.log_text.document()
        if doc.blockCount() > 100:
            cursor = self.log_text.textCursor()
            cursor.movePosition(cursor.MoveOperation.Start)
            cursor.movePosition(cursor.MoveOperation.Down, cursor.MoveMode.KeepAnchor,
                                doc.blockCount() - 100)
            cursor.removeSelectedText()


# ================== CIKIS FISI WIDGET ==================
class CikisFisiWidget(QWidget):
    u"""Di\u011fer \u00c7\u0131k\u0131\u015flar sekmesi - Mikro cikis fislerini Supabase'e senkronize etme"""

    def __init__(self):
        super().__init__()
        self._data_loaded = False
        self.sync_thread = None
        self.last_sync_time = None

        # Clients
        self.supabase_client = None
        self.dogtas_client = None
        self._init_clients()

        # Data
        self.all_data = []
        self.filtered_data = []
        self.readings_map = {}

        # UI
        self.setup_ui()
        self.setup_connections()

    def showEvent(self, event):
        super().showEvent(event)
        if not self._data_loaded:
            self._data_loaded = True
            QTimer.singleShot(100, self.load_data)

    def _init_clients(self):
        try:
            config_manager = CentralConfigManager()
            settings = config_manager.get_settings()

            supabase_url = (settings.get('Barkod_SUPABASE_URL') or
                            settings.get('SUPABASE_URL', ''))
            supabase_key = (settings.get('Barkod_SUPABASE_ANON_KEY') or
                            settings.get('SUPABASE_ANON_KEY', ''))

            if not supabase_url or not supabase_key:
                config_manager.settings_cache = {}
                settings = config_manager.get_settings(use_cache=False)
                supabase_url = (settings.get('Barkod_SUPABASE_URL') or
                                settings.get('SUPABASE_URL', ''))
                supabase_key = (settings.get('Barkod_SUPABASE_ANON_KEY') or
                                settings.get('SUPABASE_ANON_KEY', ''))

            if supabase_url and supabase_key:
                self.supabase_client = SupabaseClient(supabase_url, supabase_key)
                logger.info("Cikis: Supabase client basariyla olusturuldu")
            else:
                logger.warning("Cikis: Supabase ayarlari eksik")

            # Dogtas API config
            dogtas_config = {
                'base_url': settings.get('base_url', ''),
                'userName': settings.get('userName', ''),
                'password': settings.get('password', ''),
                'clientId': settings.get('clientId', ''),
                'clientSecret': settings.get('clientSecret', ''),
                'applicationCode': settings.get('applicationCode', ''),
                'CustomerNo': settings.get('CustomerNo', ''),
            }

            if dogtas_config['base_url']:
                self.dogtas_client = DogtasApiClient(dogtas_config)
            else:
                logger.warning("Cikis: Dogtas API ayarlari eksik")

        except Exception as e:
            logger.error(f"Cikis client init hatasi: {e}")

    # ==================== UI SETUP ====================
    def setup_ui(self):
        self.setStyleSheet("QWidget { background-color: #ffffff; }")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        # Header
        layout.addWidget(self._create_header())

        # Filter bar
        layout.addWidget(self._create_filter_bar())

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setMaximumHeight(20)
        layout.addWidget(self.progress_bar)

        # Table
        self.table = QTableWidget()
        self.table.setStyleSheet(TABLE_STYLE)
        self.table.setItemDelegate(NoFocusDelegate(self.table))
        self.table.setSelectionBehavior(QTableWidget.SelectItems)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.setFocusPolicy(Qt.StrongFocus)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        # Ctrl+C kısayolu - self üzerine bağlanır ki focus nerede olursa olsun çalışsın
        self.copy_shortcut = QShortcut(QKeySequence("Ctrl+C"), self)
        self.copy_shortcut.setContext(Qt.WidgetWithChildrenShortcut)
        self.copy_shortcut.activated.connect(self.handle_ctrl_c)
        layout.addWidget(self.table, 3)

        # Log area
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(120)
        self.log_text.setStyleSheet(LOG_STYLE)
        layout.addWidget(self.log_text, 1)

        # Status bar
        self.status_label = QLabel("Hazir")
        self.status_label.setStyleSheet("QLabel { color: #6b7280; font-size: 12px; padding: 4px; }")
        layout.addWidget(self.status_label)

    def _create_header(self) -> QWidget:
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 0)

        self.sync_button = QPushButton(u"\u00c7\u0131k\u0131\u015f Fi\u015fi Aktar")
        self.sync_button.setStyleSheet(SYNC_BUTTON_STYLE)

        self.refresh_button = QPushButton("Yenile")
        self.refresh_button.setStyleSheet(BUTTON_STYLE)

        self.all_button = QPushButton("Hepsi")
        self.all_button.setStyleSheet(BUTTON_STYLE)

        self.export_button = QPushButton("Excel")
        self.export_button.setStyleSheet(BUTTON_STYLE)

        self.btn_tumunu = QPushButton(u"T\u00fcm\u00fc")
        self.btn_tumunu.setStyleSheet(BUTTON_STYLE)

        self.btn_sil = QPushButton(u"Se\u00e7ilenleri Sil")
        self.btn_sil.setStyleSheet(BUTTON_STYLE)

        self.last_sync_label = QLabel("Son Sync: -")
        self.last_sync_label.setStyleSheet(INFO_LABEL_STYLE)

        header_layout.addWidget(self.sync_button)
        header_layout.addWidget(self.refresh_button)
        header_layout.addWidget(self.all_button)
        header_layout.addWidget(self.export_button)
        header_layout.addStretch()
        header_layout.addWidget(self.last_sync_label)

        header_widget = QWidget()
        header_widget.setLayout(header_layout)
        return header_widget

    def _create_filter_bar(self) -> QWidget:
        filter_layout = QHBoxLayout()
        filter_layout.setContentsMargins(0, 0, 0, 0)

        self.filter_evrak = QLineEdit()
        self.filter_evrak.setPlaceholderText("Evrak No")
        self.filter_evrak.setStyleSheet(FILTER_INPUT_STYLE)

        self.filter_stok = QLineEdit()
        self.filter_stok.setPlaceholderText("Stok Kodu")
        self.filter_stok.setStyleSheet(FILTER_INPUT_STYLE)

        self.filter_malzeme = QLineEdit()
        self.filter_malzeme.setPlaceholderText("Malzeme")
        self.filter_malzeme.setStyleSheet(FILTER_INPUT_STYLE)

        self.filter_tarih = QLineEdit()
        self.filter_tarih.setPlaceholderText("Tarihten itibaren (YYYY-MM-DD)")
        self.filter_tarih.setStyleSheet(FILTER_INPUT_STYLE)

        self.btn_barkod = QPushButton("Barkod Okunan")
        self.btn_barkod.setCheckable(True)
        self.btn_barkod.setStyleSheet(_toggle_btn_style('#22c55e', False))

        self.btn_manuel = QPushButton("Manuel")
        self.btn_manuel.setCheckable(True)
        self.btn_manuel.setStyleSheet(_toggle_btn_style('#ef4444', False))

        self.filter_clear_btn = QPushButton("Temizle")
        self.filter_clear_btn.setStyleSheet(BUTTON_STYLE)

        filter_layout.addWidget(self.btn_tumunu)
        filter_layout.addWidget(self.btn_sil)
        filter_layout.addWidget(self.filter_evrak)
        filter_layout.addWidget(self.filter_stok)
        filter_layout.addWidget(self.filter_malzeme)
        filter_layout.addWidget(self.filter_tarih)
        filter_layout.addWidget(self.btn_barkod)
        filter_layout.addWidget(self.btn_manuel)
        filter_layout.addWidget(self.filter_clear_btn)

        filter_widget = QWidget()
        filter_widget.setLayout(filter_layout)
        return filter_widget

    # ==================== CONNECTIONS ====================
    def setup_connections(self):
        self.sync_button.clicked.connect(self.start_sync)
        self.refresh_button.clicked.connect(self.load_data)
        self.all_button.clicked.connect(self.load_all_data)
        self.export_button.clicked.connect(self.export_to_excel)
        self.btn_tumunu.clicked.connect(self._toggle_select_all_rows)
        self.btn_sil.clicked.connect(self._delete_selected_rows)
        self.filter_clear_btn.clicked.connect(self._clear_filters)

        self.filter_timer = QTimer()
        self.filter_timer.setSingleShot(True)
        self.filter_timer.timeout.connect(self.apply_filters)

        for f in [self.filter_evrak, self.filter_stok,
                   self.filter_malzeme, self.filter_tarih]:
            f.textChanged.connect(self._schedule_filter)

        self.btn_barkod.clicked.connect(self._on_toggle_btn)
        self.btn_manuel.clicked.connect(self._on_toggle_btn)

    def _schedule_filter(self):
        self.filter_timer.start(300)

    def _on_toggle_btn(self):
        btn = self.sender()
        color_map = {
            self.btn_barkod: '#22c55e',
            self.btn_manuel: '#ef4444',
        }
        color = color_map.get(btn, '#666')
        btn.setStyleSheet(_toggle_btn_style(color, btn.isChecked()))
        self._schedule_filter()

    def _clear_filters(self):
        for f in [self.filter_evrak, self.filter_stok,
                   self.filter_malzeme, self.filter_tarih]:
            f.clear()
        for btn, color in [(self.btn_barkod, '#22c55e'),
                           (self.btn_manuel, '#ef4444')]:
            btn.setChecked(False)
            btn.setStyleSheet(_toggle_btn_style(color, False))
        self.apply_filters()

    # ==================== SYNC ====================
    def start_sync(self):
        if not self.supabase_client:
            self.log("HATA: Supabase baglantisi yapilandirilmamis!")
            self.status_label.setText("Supabase ayarlari eksik")
            return

        if self.sync_thread and self.sync_thread.isRunning():
            self.log("Sync zaten calisiyor...")
            return

        self.log("Senkronizasyon basladi...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self._set_buttons_enabled(False)

        self.sync_thread = CikisSyncThread(self.supabase_client, self.dogtas_client)
        self.sync_thread.progress_updated.connect(self._on_sync_progress)
        self.sync_thread.sync_finished.connect(self._on_sync_finished)
        self.sync_thread.error_occurred.connect(self._on_sync_error)
        self.sync_thread.finished.connect(self._on_thread_finished)
        self.sync_thread.start()

    def _on_sync_progress(self, progress, message):
        self.progress_bar.setValue(progress)
        self.status_label.setText(message)
        self.log(message)

    def _on_sync_finished(self, result):
        self.last_sync_time = datetime.now()
        self.last_sync_label.setText(
            f"Son Sync: {self.last_sync_time.strftime('%Y-%m-%d %H:%M:%S')}"
        )
        mesaj = result['mesaj']
        if result.get('atlanan', 0) > 0:
            mesaj += f" ({result['atlanan']} atlandi)"
        self.log(f"Tamamlandi: {mesaj}")
        self.status_label.setText(mesaj)

        QTimer.singleShot(500, self.load_data)

    def _on_sync_error(self, error_message):
        self.log(f"HATA: {error_message}")
        self.status_label.setText(f"Sync hatasi: {error_message}")

    def _on_thread_finished(self):
        self._set_buttons_enabled(True)
        QTimer.singleShot(1000, lambda: self.progress_bar.setVisible(False))

    def _set_buttons_enabled(self, enabled: bool):
        self.sync_button.setEnabled(enabled)
        self.refresh_button.setEnabled(enabled)
        self.all_button.setEnabled(enabled)
        self.export_button.setEnabled(enabled)
        self.btn_tumunu.setEnabled(enabled)
        self.btn_sil.setEnabled(enabled)

    # ==================== TABLE ====================
    def load_data(self):
        if not self.supabase_client:
            self.status_label.setText("Supabase ayarlari eksik")
            return
        from datetime import timedelta
        one_week_ago = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
        self._start_load(min_tarih=one_week_ago)

    def load_all_data(self):
        if not self.supabase_client:
            self.status_label.setText("Supabase ayarlari eksik")
            return
        self._start_load(min_tarih=None)

    def _start_load(self, min_tarih=None):
        label = "Son 1 hafta" if min_tarih else "Tumu"
        self.status_label.setText(f"Cikis fisleri yukleniyor ({label})...")
        self.refresh_button.setEnabled(False)
        self.all_button.setEnabled(False)

        self._load_thread = CikisLoadDataThread(self.supabase_client, min_tarih=min_tarih)
        self._load_thread.data_loaded.connect(self._on_data_loaded)
        self._load_thread.error_occurred.connect(self._on_load_error)
        self._load_thread.start()

    def _on_data_loaded(self, all_data, readings_map):
        self.all_data = all_data
        self.readings_map = readings_map
        self.apply_filters()
        self.status_label.setText(f"{len(self.all_data)} kayit yuklendi")
        self.refresh_button.setEnabled(True)
        self.all_button.setEnabled(True)

    def _on_load_error(self, error_msg):
        self.status_label.setText(f"Yukleme hatasi: {error_msg}")
        self.log(f"HATA: Tablo yukleme hatasi: {error_msg}")
        self.refresh_button.setEnabled(True)
        self.all_button.setEnabled(True)

    def _get_row_colors(self, row_data):
        colors = set()
        kalem_id = row_data.get('id')
        paket_readings = self.readings_map.get(kalem_id, {})
        for ps, reads in paket_readings.items():
            for info in reads:
                if info['type'] == 'scanner':
                    colors.add('green')
                elif info['type'] == 'manual':
                    colors.add('red')
        return colors

    def apply_filters(self):
        filtered = self.all_data[:]

        evrak_text = self.filter_evrak.text().strip()
        stok_text = self.filter_stok.text().strip()
        malzeme_text = self.filter_malzeme.text().strip()
        tarih_text = self.filter_tarih.text().strip()

        if evrak_text:
            filtered = [r for r in filtered
                        if evrak_text in str(r.get('evrakno_sira', ''))]
        if stok_text:
            filtered = [r for r in filtered
                        if _fuzzy_match(stok_text, str(r.get('stok_kod', '')))]
        if malzeme_text:
            filtered = [r for r in filtered
                        if _fuzzy_match(malzeme_text, str(r.get('malzeme_adi', '')))]
        if tarih_text:
            filtered = [r for r in filtered
                        if str(r.get('tarih', '') or '') >= tarih_text]

        # Renk filtresi
        want_colors = set()
        if self.btn_barkod.isChecked():
            want_colors.add('green')
        if self.btn_manuel.isChecked():
            want_colors.add('red')
        if want_colors:
            filtered = [r for r in filtered
                        if self._get_row_colors(r) & want_colors]

        self.filtered_data = filtered
        self.populate_table()

    def populate_table(self):
        self.table.setUpdatesEnabled(False)
        self.table.setSortingEnabled(False)

        try:
            self.table.setRowCount(len(self.filtered_data))
            self.table.setColumnCount(len(CIKIS_TABLE_COLUMNS) + 1)
            self.table.setHorizontalHeaderLabels([u'\u2611'] + [c[1] for c in CIKIS_TABLE_COLUMNS])

            okuma_col_idx = next(
                (j + 1 for j, (k, _) in enumerate(CIKIS_TABLE_COLUMNS) if k == 'okuma_durumu'),
                None
            )

            for i, row_data in enumerate(self.filtered_data):
                # Checkbox sutunu (index 0)
                chk_item = QTableWidgetItem()
                chk_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
                chk_item.setCheckState(Qt.Unchecked)
                chk_item.setData(Qt.UserRole, row_data.get('id'))
                self.table.setItem(i, 0, chk_item)

                for j, (key, _) in enumerate(CIKIS_TABLE_COLUMNS):
                    if key == 'okuma_durumu':
                        continue
                    value = row_data.get(key, '')
                    if value is None:
                        text = ''
                    elif key == 'miktar':
                        num = float(value)
                        text = str(int(num)) if num == int(num) else str(num)
                    elif key == 'tarih':
                        text = str(value)[:10] if value else ''
                    else:
                        text = str(value)
                    item = QTableWidgetItem(text)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    font = QFont(FONT_FAMILY, FONT_SIZE)
                    font.setBold(True)
                    item.setFont(font)
                    self.table.setItem(i, j + 1, item)

                # Okuma Durumu sutunu
                if okuma_col_idx is not None:
                    miktar = int(float(row_data.get('miktar', 0) or 0))
                    paket = int(row_data.get('paket_sayisi', 1) or 1)
                    kalem_id = row_data.get('id')
                    paket_readings = self.readings_map.get(kalem_id, {})
                    depo_no = str(row_data.get('depo', '') or '')
                    widget = _build_okuma_durumu_widget(miktar, paket, paket_readings, depo_no)
                    self.table.setCellWidget(i, okuma_col_idx, widget)

            header = self.table.horizontalHeader()
            header.setMinimumSectionSize(0)
            header.setSectionResizeMode(QHeaderView.ResizeToContents)
            header.setStretchLastSection(False)
            header.setSectionResizeMode(0, QHeaderView.Fixed)
            self.table.setColumnWidth(0, 30)

            # Malzeme Adi sutunu ResizeToContents
            malzeme_adi_idx = next(
                (j + 1 for j, (k, _) in enumerate(CIKIS_TABLE_COLUMNS) if k == 'malzeme_adi'),
                len(CIKIS_TABLE_COLUMNS) - 1
            )
            header.setSectionResizeMode(malzeme_adi_idx, QHeaderView.ResizeToContents)

            if okuma_col_idx is not None:
                if self.table.columnWidth(okuma_col_idx) < 200:
                    self.table.setColumnWidth(okuma_col_idx, 200)

            for i in range(self.table.rowCount()):
                self.table.setRowHeight(i, ROW_HEIGHT)

        finally:
            self.table.setUpdatesEnabled(True)
            self.table.setSortingEnabled(True)

    # ==================== CHECKBOX ISLEMLERI ====================
    def _toggle_select_all_rows(self):
        any_checked = any(
            self.table.item(i, 0) and self.table.item(i, 0).checkState() == Qt.Checked
            for i in range(self.table.rowCount())
        )
        new_state = Qt.Unchecked if any_checked else Qt.Checked
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 0)
            if item:
                item.setCheckState(new_state)

    def _delete_selected_rows(self):
        if not self.supabase_client:
            _show_message(self, "Hata", "Supabase bağlantısı yok.")
            return
        ids = []
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 0)
            if item and item.checkState() == Qt.Checked:
                id_val = item.data(Qt.UserRole)
                if id_val is not None:
                    ids.append(id_val)
        if not ids:
            _show_message(self, "Bilgi", "Silinecek satır seçilmedi.")
            return
        if not _verify_barkod_delete_password(self):
            return
        if not _confirm_delete(self, f"{len(ids)} satır Supabase'den silinecek. Emin misiniz?"):
            return
        try:
            self.supabase_client.delete_cikis_by_id_list(ids)
            self.status_label.setText(f"{len(ids)} satır silindi.")
            self.load_data()
        except Exception as e:
            _show_message(self, "Hata", f"Silme hatası: {e}")

    # ==================== EXPORT ====================
    def export_to_excel(self):
        if not self.filtered_data:
            self.status_label.setText("Disari aktarilacak veri yok")
            return

        try:
            df = pd.DataFrame(self.filtered_data)
            output_path = "D:/GoogleDrive/~ CikisFisi_Export.xlsx"
            df.to_excel(output_path, index=False, engine='openpyxl')
            self.status_label.setText(f"Excel export: {output_path}")
            self.log(f"Veriler disari aktarildi: {output_path}")
        except Exception as e:
            self.status_label.setText(f"Export hatasi: {e}")
            self.log(f"HATA: Export hatasi: {e}")

    # ==================== KOPYALAMA ====================
    def show_context_menu(self, pos):
        menu = QMenu(self.table)
        menu.setStyleSheet("""
            QMenu {
                background-color: #ffffff;
                border: 1px solid #a0a0a0;
                padding: 4px 0px;
            }
            QMenu::item {
                padding: 6px 24px;
                color: #000000;
                font-size: 14px;
            }
            QMenu::item:selected {
                background-color: #3399ff;
                color: #ffffff;
            }
        """)
        hucre_action = QAction("Kopyala", self)
        hucre_action.triggered.connect(lambda: self.copy_cell(pos))
        menu.addAction(hucre_action)
        menu.exec_(self.table.viewport().mapToGlobal(pos))

    def copy_cell(self, pos):
        item = self.table.itemAt(pos)
        if item:
            QApplication.clipboard().setText(item.text())
            old_text = self.status_label.text()
            self.status_label.setText("✅ Kopyalandı")
            QTimer.singleShot(1500, lambda t=old_text: self.status_label.setText(t))

    def handle_ctrl_c(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            return
        if len(selected_items) == 1:
            QApplication.clipboard().setText(selected_items[0].text())
        else:
            rows = sorted({item.row() for item in selected_items})
            cols = sorted({item.column() for item in selected_items})
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
        old_text = self.status_label.text()
        self.status_label.setText("✅ Kopyalandı")
        QTimer.singleShot(1500, lambda t=old_text: self.status_label.setText(t))

    # ==================== LOG ====================
    def log(self, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        doc = self.log_text.document()
        if doc.blockCount() > 100:
            cursor = self.log_text.textCursor()
            cursor.movePosition(cursor.MoveOperation.Start)
            cursor.movePosition(cursor.MoveOperation.Down, cursor.MoveMode.KeepAnchor,
                                doc.blockCount() - 100)
            cursor.removeSelectedText()


# ================== GIRIS FISI WIDGET ==================
class GirisFisiWidget(QWidget):
    u"""Di\u011fer Giri\u015fler sekmesi - Mikro giris fislerini Supabase'e senkronize etme"""

    def __init__(self):
        super().__init__()
        self._data_loaded = False
        self.sync_thread = None
        self.last_sync_time = None

        # Clients
        self.supabase_client = None
        self.dogtas_client = None
        self._init_clients()

        # Data
        self.all_data = []
        self.filtered_data = []
        self.readings_map = {}

        # UI
        self.setup_ui()
        self.setup_connections()

    def showEvent(self, event):
        super().showEvent(event)
        if not self._data_loaded:
            self._data_loaded = True
            QTimer.singleShot(100, self.load_data)

    def _init_clients(self):
        try:
            config_manager = CentralConfigManager()
            settings = config_manager.get_settings()

            supabase_url = (settings.get('Barkod_SUPABASE_URL') or
                            settings.get('SUPABASE_URL', ''))
            supabase_key = (settings.get('Barkod_SUPABASE_ANON_KEY') or
                            settings.get('SUPABASE_ANON_KEY', ''))

            if not supabase_url or not supabase_key:
                config_manager.settings_cache = {}
                settings = config_manager.get_settings(use_cache=False)
                supabase_url = (settings.get('Barkod_SUPABASE_URL') or
                                settings.get('SUPABASE_URL', ''))
                supabase_key = (settings.get('Barkod_SUPABASE_ANON_KEY') or
                                settings.get('SUPABASE_ANON_KEY', ''))

            if supabase_url and supabase_key:
                self.supabase_client = SupabaseClient(supabase_url, supabase_key)
                logger.info("Giris: Supabase client basariyla olusturuldu")
            else:
                logger.warning("Giris: Supabase ayarlari eksik")

            # Dogtas API config
            dogtas_config = {
                'base_url': settings.get('base_url', ''),
                'userName': settings.get('userName', ''),
                'password': settings.get('password', ''),
                'clientId': settings.get('clientId', ''),
                'clientSecret': settings.get('clientSecret', ''),
                'applicationCode': settings.get('applicationCode', ''),
                'CustomerNo': settings.get('CustomerNo', ''),
            }

            if dogtas_config['base_url']:
                self.dogtas_client = DogtasApiClient(dogtas_config)
            else:
                logger.warning("Giris: Dogtas API ayarlari eksik")

        except Exception as e:
            logger.error(f"Giris client init hatasi: {e}")

    # ==================== UI SETUP ====================
    def setup_ui(self):
        self.setStyleSheet("QWidget { background-color: #ffffff; }")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        # Header
        layout.addWidget(self._create_header())

        # Filter bar
        layout.addWidget(self._create_filter_bar())

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setMaximumHeight(20)
        layout.addWidget(self.progress_bar)

        # Table
        self.table = QTableWidget()
        self.table.setStyleSheet(TABLE_STYLE)
        self.table.setItemDelegate(NoFocusDelegate(self.table))
        self.table.setSelectionBehavior(QTableWidget.SelectItems)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.setFocusPolicy(Qt.StrongFocus)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        # Ctrl+C kısayolu - self üzerine bağlanır ki focus nerede olursa olsun çalışsın
        self.copy_shortcut = QShortcut(QKeySequence("Ctrl+C"), self)
        self.copy_shortcut.setContext(Qt.WidgetWithChildrenShortcut)
        self.copy_shortcut.activated.connect(self.handle_ctrl_c)
        layout.addWidget(self.table, 3)

        # Log area
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(120)
        self.log_text.setStyleSheet(LOG_STYLE)
        layout.addWidget(self.log_text, 1)

        # Status bar
        self.status_label = QLabel("Hazir")
        self.status_label.setStyleSheet("QLabel { color: #6b7280; font-size: 12px; padding: 4px; }")
        layout.addWidget(self.status_label)

    def _create_header(self) -> QWidget:
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 0)

        self.sync_button = QPushButton(u"Giri\u015f Fi\u015fi Aktar")
        self.sync_button.setStyleSheet(SYNC_BUTTON_STYLE)

        self.refresh_button = QPushButton("Yenile")
        self.refresh_button.setStyleSheet(BUTTON_STYLE)

        self.all_button = QPushButton("Hepsi")
        self.all_button.setStyleSheet(BUTTON_STYLE)

        self.export_button = QPushButton("Excel")
        self.export_button.setStyleSheet(BUTTON_STYLE)

        self.btn_tumunu = QPushButton(u"T\u00fcm\u00fc")
        self.btn_tumunu.setStyleSheet(BUTTON_STYLE)

        self.btn_sil = QPushButton(u"Se\u00e7ilenleri Sil")
        self.btn_sil.setStyleSheet(BUTTON_STYLE)

        self.last_sync_label = QLabel("Son Sync: -")
        self.last_sync_label.setStyleSheet(INFO_LABEL_STYLE)

        header_layout.addWidget(self.sync_button)
        header_layout.addWidget(self.refresh_button)
        header_layout.addWidget(self.all_button)
        header_layout.addWidget(self.export_button)
        header_layout.addStretch()
        header_layout.addWidget(self.last_sync_label)

        header_widget = QWidget()
        header_widget.setLayout(header_layout)
        return header_widget

    def _create_filter_bar(self) -> QWidget:
        filter_layout = QHBoxLayout()
        filter_layout.setContentsMargins(0, 0, 0, 0)

        self.filter_evrak = QLineEdit()
        self.filter_evrak.setPlaceholderText("Evrak No")
        self.filter_evrak.setStyleSheet(FILTER_INPUT_STYLE)

        self.filter_stok = QLineEdit()
        self.filter_stok.setPlaceholderText("Stok Kodu")
        self.filter_stok.setStyleSheet(FILTER_INPUT_STYLE)

        self.filter_malzeme = QLineEdit()
        self.filter_malzeme.setPlaceholderText("Malzeme")
        self.filter_malzeme.setStyleSheet(FILTER_INPUT_STYLE)

        self.filter_tarih = QLineEdit()
        self.filter_tarih.setPlaceholderText("Tarihten itibaren (YYYY-MM-DD)")
        self.filter_tarih.setStyleSheet(FILTER_INPUT_STYLE)

        self.btn_barkod = QPushButton("Barkod Okunan")
        self.btn_barkod.setCheckable(True)
        self.btn_barkod.setStyleSheet(_toggle_btn_style('#22c55e', False))

        self.btn_manuel = QPushButton("Manuel")
        self.btn_manuel.setCheckable(True)
        self.btn_manuel.setStyleSheet(_toggle_btn_style('#ef4444', False))

        self.filter_clear_btn = QPushButton("Temizle")
        self.filter_clear_btn.setStyleSheet(BUTTON_STYLE)

        filter_layout.addWidget(self.btn_tumunu)
        filter_layout.addWidget(self.btn_sil)
        filter_layout.addWidget(self.filter_evrak)
        filter_layout.addWidget(self.filter_stok)
        filter_layout.addWidget(self.filter_malzeme)
        filter_layout.addWidget(self.filter_tarih)
        filter_layout.addWidget(self.btn_barkod)
        filter_layout.addWidget(self.btn_manuel)
        filter_layout.addWidget(self.filter_clear_btn)

        filter_widget = QWidget()
        filter_widget.setLayout(filter_layout)
        return filter_widget

    # ==================== CONNECTIONS ====================
    def setup_connections(self):
        self.sync_button.clicked.connect(self.start_sync)
        self.refresh_button.clicked.connect(self.load_data)
        self.all_button.clicked.connect(self.load_all_data)
        self.export_button.clicked.connect(self.export_to_excel)
        self.btn_tumunu.clicked.connect(self._toggle_select_all_rows)
        self.btn_sil.clicked.connect(self._delete_selected_rows)
        self.filter_clear_btn.clicked.connect(self._clear_filters)

        self.filter_timer = QTimer()
        self.filter_timer.setSingleShot(True)
        self.filter_timer.timeout.connect(self.apply_filters)

        for f in [self.filter_evrak, self.filter_stok,
                   self.filter_malzeme, self.filter_tarih]:
            f.textChanged.connect(self._schedule_filter)

        self.btn_barkod.clicked.connect(self._on_toggle_btn)
        self.btn_manuel.clicked.connect(self._on_toggle_btn)

    def _schedule_filter(self):
        self.filter_timer.start(300)

    def _on_toggle_btn(self):
        btn = self.sender()
        color_map = {
            self.btn_barkod: '#22c55e',
            self.btn_manuel: '#ef4444',
        }
        color = color_map.get(btn, '#666')
        btn.setStyleSheet(_toggle_btn_style(color, btn.isChecked()))
        self._schedule_filter()

    def _clear_filters(self):
        for f in [self.filter_evrak, self.filter_stok,
                   self.filter_malzeme, self.filter_tarih]:
            f.clear()
        for btn, color in [(self.btn_barkod, '#22c55e'),
                           (self.btn_manuel, '#ef4444')]:
            btn.setChecked(False)
            btn.setStyleSheet(_toggle_btn_style(color, False))
        self.apply_filters()

    # ==================== SYNC ====================
    def start_sync(self):
        if not self.supabase_client:
            self.log("HATA: Supabase baglantisi yapilandirilmamis!")
            self.status_label.setText("Supabase ayarlari eksik")
            return

        if self.sync_thread and self.sync_thread.isRunning():
            self.log("Sync zaten calisiyor...")
            return

        self.log("Senkronizasyon basladi...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self._set_buttons_enabled(False)

        self.sync_thread = GirisSyncThread(self.supabase_client, self.dogtas_client)
        self.sync_thread.progress_updated.connect(self._on_sync_progress)
        self.sync_thread.sync_finished.connect(self._on_sync_finished)
        self.sync_thread.error_occurred.connect(self._on_sync_error)
        self.sync_thread.finished.connect(self._on_thread_finished)
        self.sync_thread.start()

    def _on_sync_progress(self, progress, message):
        self.progress_bar.setValue(progress)
        self.status_label.setText(message)
        self.log(message)

    def _on_sync_finished(self, result):
        self.last_sync_time = datetime.now()
        self.last_sync_label.setText(
            f"Son Sync: {self.last_sync_time.strftime('%Y-%m-%d %H:%M:%S')}"
        )
        mesaj = result['mesaj']
        if result.get('atlanan', 0) > 0:
            mesaj += f" ({result['atlanan']} atlandi)"
        self.log(f"Tamamlandi: {mesaj}")
        self.status_label.setText(mesaj)

        QTimer.singleShot(500, self.load_data)

    def _on_sync_error(self, error_message):
        self.log(f"HATA: {error_message}")
        self.status_label.setText(f"Sync hatasi: {error_message}")

    def _on_thread_finished(self):
        self._set_buttons_enabled(True)
        QTimer.singleShot(1000, lambda: self.progress_bar.setVisible(False))

    def _set_buttons_enabled(self, enabled: bool):
        self.sync_button.setEnabled(enabled)
        self.refresh_button.setEnabled(enabled)
        self.all_button.setEnabled(enabled)
        self.export_button.setEnabled(enabled)
        self.btn_tumunu.setEnabled(enabled)
        self.btn_sil.setEnabled(enabled)

    # ==================== TABLE ====================
    def load_data(self):
        if not self.supabase_client:
            self.status_label.setText("Supabase ayarlari eksik")
            return
        from datetime import timedelta
        one_week_ago = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
        self._start_load(min_tarih=one_week_ago)

    def load_all_data(self):
        if not self.supabase_client:
            self.status_label.setText("Supabase ayarlari eksik")
            return
        self._start_load(min_tarih=None)

    def _start_load(self, min_tarih=None):
        label = "Son 1 hafta" if min_tarih else "Tumu"
        self.status_label.setText(f"Giris fisleri yukleniyor ({label})...")
        self.refresh_button.setEnabled(False)
        self.all_button.setEnabled(False)

        self._load_thread = GirisLoadDataThread(self.supabase_client, min_tarih=min_tarih)
        self._load_thread.data_loaded.connect(self._on_data_loaded)
        self._load_thread.error_occurred.connect(self._on_load_error)
        self._load_thread.start()

    def _on_data_loaded(self, all_data, readings_map):
        self.all_data = all_data
        self.readings_map = readings_map
        self.apply_filters()
        self.status_label.setText(f"{len(self.all_data)} kayit yuklendi")
        self.refresh_button.setEnabled(True)
        self.all_button.setEnabled(True)

    def _on_load_error(self, error_msg):
        self.status_label.setText(f"Yukleme hatasi: {error_msg}")
        self.log(f"HATA: Tablo yukleme hatasi: {error_msg}")
        self.refresh_button.setEnabled(True)
        self.all_button.setEnabled(True)

    def _get_row_colors(self, row_data):
        colors = set()
        kalem_id = row_data.get('id')
        paket_readings = self.readings_map.get(kalem_id, {})
        for ps, reads in paket_readings.items():
            for info in reads:
                if info['type'] == 'scanner':
                    colors.add('green')
                elif info['type'] == 'manual':
                    colors.add('red')
        return colors

    def apply_filters(self):
        filtered = self.all_data[:]

        evrak_text = self.filter_evrak.text().strip()
        stok_text = self.filter_stok.text().strip()
        malzeme_text = self.filter_malzeme.text().strip()
        tarih_text = self.filter_tarih.text().strip()

        if evrak_text:
            filtered = [r for r in filtered
                        if evrak_text in str(r.get('evrakno_sira', ''))]
        if stok_text:
            filtered = [r for r in filtered
                        if _fuzzy_match(stok_text, str(r.get('stok_kod', '')))]
        if malzeme_text:
            filtered = [r for r in filtered
                        if _fuzzy_match(malzeme_text, str(r.get('malzeme_adi', '')))]
        if tarih_text:
            filtered = [r for r in filtered
                        if str(r.get('tarih', '') or '') >= tarih_text]

        # Renk filtresi
        want_colors = set()
        if self.btn_barkod.isChecked():
            want_colors.add('green')
        if self.btn_manuel.isChecked():
            want_colors.add('red')
        if want_colors:
            filtered = [r for r in filtered
                        if self._get_row_colors(r) & want_colors]

        self.filtered_data = filtered
        self.populate_table()

    def populate_table(self):
        self.table.setUpdatesEnabled(False)
        self.table.setSortingEnabled(False)

        try:
            self.table.setRowCount(len(self.filtered_data))
            self.table.setColumnCount(len(GIRIS_TABLE_COLUMNS) + 1)
            self.table.setHorizontalHeaderLabels([u'\u2611'] + [c[1] for c in GIRIS_TABLE_COLUMNS])

            okuma_col_idx = next(
                (j + 1 for j, (k, _) in enumerate(GIRIS_TABLE_COLUMNS) if k == 'okuma_durumu'),
                None
            )

            for i, row_data in enumerate(self.filtered_data):
                # Checkbox sutunu (index 0)
                chk_item = QTableWidgetItem()
                chk_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
                chk_item.setCheckState(Qt.Unchecked)
                chk_item.setData(Qt.UserRole, row_data.get('id'))
                self.table.setItem(i, 0, chk_item)

                for j, (key, _) in enumerate(GIRIS_TABLE_COLUMNS):
                    if key == 'okuma_durumu':
                        continue
                    value = row_data.get(key, '')
                    if value is None:
                        text = ''
                    elif key == 'miktar':
                        num = float(value)
                        text = str(int(num)) if num == int(num) else str(num)
                    elif key == 'tarih':
                        text = str(value)[:10] if value else ''
                    else:
                        text = str(value)
                    item = QTableWidgetItem(text)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    font = QFont(FONT_FAMILY, FONT_SIZE)
                    font.setBold(True)
                    item.setFont(font)
                    self.table.setItem(i, j + 1, item)

                # Okuma Durumu sutunu
                if okuma_col_idx is not None:
                    miktar = int(float(row_data.get('miktar', 0) or 0))
                    paket = int(row_data.get('paket_sayisi', 1) or 1)
                    kalem_id = row_data.get('id')
                    paket_readings = self.readings_map.get(kalem_id, {})
                    depo_no = str(row_data.get('depo', '') or '')
                    widget = _build_okuma_durumu_widget(miktar, paket, paket_readings, depo_no)
                    self.table.setCellWidget(i, okuma_col_idx, widget)

            header = self.table.horizontalHeader()
            header.setMinimumSectionSize(0)
            header.setSectionResizeMode(QHeaderView.ResizeToContents)
            header.setStretchLastSection(False)
            header.setSectionResizeMode(0, QHeaderView.Fixed)
            self.table.setColumnWidth(0, 30)

            # Malzeme Adi sutunu ResizeToContents
            malzeme_adi_idx = next(
                (j + 1 for j, (k, _) in enumerate(GIRIS_TABLE_COLUMNS) if k == 'malzeme_adi'),
                len(GIRIS_TABLE_COLUMNS) - 1
            )
            header.setSectionResizeMode(malzeme_adi_idx, QHeaderView.ResizeToContents)

            if okuma_col_idx is not None:
                if self.table.columnWidth(okuma_col_idx) < 200:
                    self.table.setColumnWidth(okuma_col_idx, 200)

            for i in range(self.table.rowCount()):
                self.table.setRowHeight(i, ROW_HEIGHT)

        finally:
            self.table.setUpdatesEnabled(True)
            self.table.setSortingEnabled(True)

    # ==================== CHECKBOX ISLEMLERI ====================
    def _toggle_select_all_rows(self):
        any_checked = any(
            self.table.item(i, 0) and self.table.item(i, 0).checkState() == Qt.Checked
            for i in range(self.table.rowCount())
        )
        new_state = Qt.Unchecked if any_checked else Qt.Checked
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 0)
            if item:
                item.setCheckState(new_state)

    def _delete_selected_rows(self):
        if not self.supabase_client:
            _show_message(self, "Hata", "Supabase bağlantısı yok.")
            return
        ids = []
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 0)
            if item and item.checkState() == Qt.Checked:
                id_val = item.data(Qt.UserRole)
                if id_val is not None:
                    ids.append(id_val)
        if not ids:
            _show_message(self, "Bilgi", "Silinecek satır seçilmedi.")
            return
        if not _verify_barkod_delete_password(self):
            return
        if not _confirm_delete(self, f"{len(ids)} satır Supabase'den silinecek. Emin misiniz?"):
            return
        try:
            self.supabase_client.delete_giris_by_id_list(ids)
            self.status_label.setText(f"{len(ids)} satır silindi.")
            self.load_data()
        except Exception as e:
            _show_message(self, "Hata", f"Silme hatası: {e}")

    # ==================== EXPORT ====================
    def export_to_excel(self):
        if not self.filtered_data:
            self.status_label.setText("Disari aktarilacak veri yok")
            return

        try:
            df = pd.DataFrame(self.filtered_data)
            output_path = "D:/GoogleDrive/~ GirisFisi_Export.xlsx"
            df.to_excel(output_path, index=False, engine='openpyxl')
            self.status_label.setText(f"Excel export: {output_path}")
            self.log(f"Veriler disari aktarildi: {output_path}")
        except Exception as e:
            self.status_label.setText(f"Export hatasi: {e}")
            self.log(f"HATA: Export hatasi: {e}")

    # ==================== KOPYALAMA ====================
    def show_context_menu(self, pos):
        menu = QMenu(self.table)
        menu.setStyleSheet("""
            QMenu {
                background-color: #ffffff;
                border: 1px solid #a0a0a0;
                padding: 4px 0px;
            }
            QMenu::item {
                padding: 6px 24px;
                color: #000000;
                font-size: 14px;
            }
            QMenu::item:selected {
                background-color: #3399ff;
                color: #ffffff;
            }
        """)
        hucre_action = QAction("Kopyala", self)
        hucre_action.triggered.connect(lambda: self.copy_cell(pos))
        menu.addAction(hucre_action)
        menu.exec_(self.table.viewport().mapToGlobal(pos))

    def copy_cell(self, pos):
        item = self.table.itemAt(pos)
        if item:
            QApplication.clipboard().setText(item.text())
            old_text = self.status_label.text()
            self.status_label.setText("✅ Kopyalandı")
            QTimer.singleShot(1500, lambda t=old_text: self.status_label.setText(t))

    def handle_ctrl_c(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            return
        if len(selected_items) == 1:
            QApplication.clipboard().setText(selected_items[0].text())
        else:
            rows = sorted({item.row() for item in selected_items})
            cols = sorted({item.column() for item in selected_items})
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
        old_text = self.status_label.text()
        self.status_label.setText("✅ Kopyalandı")
        QTimer.singleShot(1500, lambda t=old_text: self.status_label.setText(t))

    # ==================== LOG ====================
    def log(self, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        doc = self.log_text.document()
        if doc.blockCount() > 100:
            cursor = self.log_text.textCursor()
            cursor.movePosition(cursor.MoveOperation.Start)
            cursor.movePosition(cursor.MoveOperation.Down, cursor.MoveMode.KeepAnchor,
                                doc.blockCount() - 100)
            cursor.removeSelectedText()


# ================== SEVK FISI WIDGET ==================
class SevkFisiWidget(QWidget):
    u"""Depolar Aras\u0131 Sevk sekmesi - Mikro sevk fislerini Supabase'e senkronize etme"""

    def __init__(self):
        super().__init__()
        self._data_loaded = False
        self.sync_thread = None
        self.last_sync_time = None

        # Clients
        self.supabase_client = None
        self.dogtas_client = None
        self._init_clients()

        # Data
        self.all_data = []
        self.filtered_data = []
        self.readings_map = {}

        # UI
        self.setup_ui()
        self.setup_connections()

    def showEvent(self, event):
        super().showEvent(event)
        if not self._data_loaded:
            self._data_loaded = True
            QTimer.singleShot(100, self.load_data)

    def _init_clients(self):
        try:
            config_manager = CentralConfigManager()
            settings = config_manager.get_settings()

            supabase_url = (settings.get('Barkod_SUPABASE_URL') or
                            settings.get('SUPABASE_URL', ''))
            supabase_key = (settings.get('Barkod_SUPABASE_ANON_KEY') or
                            settings.get('SUPABASE_ANON_KEY', ''))

            if not supabase_url or not supabase_key:
                config_manager.settings_cache = {}
                settings = config_manager.get_settings(use_cache=False)
                supabase_url = (settings.get('Barkod_SUPABASE_URL') or
                                settings.get('SUPABASE_URL', ''))
                supabase_key = (settings.get('Barkod_SUPABASE_ANON_KEY') or
                                settings.get('SUPABASE_ANON_KEY', ''))

            if supabase_url and supabase_key:
                self.supabase_client = SupabaseClient(supabase_url, supabase_key)
                logger.info("Sevk: Supabase client basariyla olusturuldu")
            else:
                logger.warning("Sevk: Supabase ayarlari eksik")

            # Dogtas API config
            dogtas_config = {
                'base_url': settings.get('base_url', ''),
                'userName': settings.get('userName', ''),
                'password': settings.get('password', ''),
                'clientId': settings.get('clientId', ''),
                'clientSecret': settings.get('clientSecret', ''),
                'applicationCode': settings.get('applicationCode', ''),
                'CustomerNo': settings.get('CustomerNo', ''),
            }

            if dogtas_config['base_url']:
                self.dogtas_client = DogtasApiClient(dogtas_config)
            else:
                logger.warning("Sevk: Dogtas API ayarlari eksik")

        except Exception as e:
            logger.error(f"Sevk client init hatasi: {e}")

    # ==================== UI SETUP ====================
    def setup_ui(self):
        self.setStyleSheet("QWidget { background-color: #ffffff; }")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        # Header
        layout.addWidget(self._create_header())

        # Filter bar
        layout.addWidget(self._create_filter_bar())

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setMaximumHeight(20)
        layout.addWidget(self.progress_bar)

        # Table
        self.table = QTableWidget()
        self.table.setStyleSheet(TABLE_STYLE)
        self.table.setItemDelegate(NoFocusDelegate(self.table))
        self.table.setSelectionBehavior(QTableWidget.SelectItems)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.setFocusPolicy(Qt.StrongFocus)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        # Ctrl+C kısayolu - self üzerine bağlanır ki focus nerede olursa olsun çalışsın
        self.copy_shortcut = QShortcut(QKeySequence("Ctrl+C"), self)
        self.copy_shortcut.setContext(Qt.WidgetWithChildrenShortcut)
        self.copy_shortcut.activated.connect(self.handle_ctrl_c)
        layout.addWidget(self.table, 3)

        # Log area
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(120)
        self.log_text.setStyleSheet(LOG_STYLE)
        layout.addWidget(self.log_text, 1)

        # Status bar
        self.status_label = QLabel("Hazir")
        self.status_label.setStyleSheet("QLabel { color: #6b7280; font-size: 12px; padding: 4px; }")
        layout.addWidget(self.status_label)

    def _create_header(self) -> QWidget:
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 0)

        self.sync_button = QPushButton(u"Sevk Fi\u015fi Aktar")
        self.sync_button.setStyleSheet(SYNC_BUTTON_STYLE)

        self.refresh_button = QPushButton("Yenile")
        self.refresh_button.setStyleSheet(BUTTON_STYLE)

        self.all_button = QPushButton("Hepsi")
        self.all_button.setStyleSheet(BUTTON_STYLE)

        self.export_button = QPushButton("Excel")
        self.export_button.setStyleSheet(BUTTON_STYLE)

        self.btn_tumunu = QPushButton(u"T\u00fcm\u00fc")
        self.btn_tumunu.setStyleSheet(BUTTON_STYLE)

        self.btn_sil = QPushButton(u"Se\u00e7ilenleri Sil")
        self.btn_sil.setStyleSheet(BUTTON_STYLE)

        self.last_sync_label = QLabel("Son Sync: -")
        self.last_sync_label.setStyleSheet(INFO_LABEL_STYLE)

        header_layout.addWidget(self.sync_button)
        header_layout.addWidget(self.refresh_button)
        header_layout.addWidget(self.all_button)
        header_layout.addWidget(self.export_button)
        header_layout.addStretch()
        header_layout.addWidget(self.last_sync_label)

        header_widget = QWidget()
        header_widget.setLayout(header_layout)
        return header_widget

    def _create_filter_bar(self) -> QWidget:
        filter_layout = QHBoxLayout()
        filter_layout.setContentsMargins(0, 0, 0, 0)

        self.filter_evrak = QLineEdit()
        self.filter_evrak.setPlaceholderText("Evrak No")
        self.filter_evrak.setStyleSheet(FILTER_INPUT_STYLE)

        self.filter_stok = QLineEdit()
        self.filter_stok.setPlaceholderText("Stok Kodu")
        self.filter_stok.setStyleSheet(FILTER_INPUT_STYLE)

        self.filter_malzeme = QLineEdit()
        self.filter_malzeme.setPlaceholderText("Malzeme")
        self.filter_malzeme.setStyleSheet(FILTER_INPUT_STYLE)

        self.filter_tarih = QLineEdit()
        self.filter_tarih.setPlaceholderText("Tarihten itibaren (YYYY-MM-DD)")
        self.filter_tarih.setStyleSheet(FILTER_INPUT_STYLE)

        self.btn_barkod = QPushButton("Barkod Okunan")
        self.btn_barkod.setCheckable(True)
        self.btn_barkod.setStyleSheet(_toggle_btn_style('#22c55e', False))

        self.btn_manuel = QPushButton("Manuel")
        self.btn_manuel.setCheckable(True)
        self.btn_manuel.setStyleSheet(_toggle_btn_style('#ef4444', False))

        self.filter_clear_btn = QPushButton("Temizle")
        self.filter_clear_btn.setStyleSheet(BUTTON_STYLE)

        filter_layout.addWidget(self.btn_tumunu)
        filter_layout.addWidget(self.btn_sil)
        filter_layout.addWidget(self.filter_evrak)
        filter_layout.addWidget(self.filter_stok)
        filter_layout.addWidget(self.filter_malzeme)
        filter_layout.addWidget(self.filter_tarih)
        filter_layout.addWidget(self.btn_barkod)
        filter_layout.addWidget(self.btn_manuel)
        filter_layout.addWidget(self.filter_clear_btn)

        filter_widget = QWidget()
        filter_widget.setLayout(filter_layout)
        return filter_widget

    # ==================== CONNECTIONS ====================
    def setup_connections(self):
        self.sync_button.clicked.connect(self.start_sync)
        self.refresh_button.clicked.connect(self.load_data)
        self.all_button.clicked.connect(self.load_all_data)
        self.export_button.clicked.connect(self.export_to_excel)
        self.btn_tumunu.clicked.connect(self._toggle_select_all_rows)
        self.btn_sil.clicked.connect(self._delete_selected_rows)
        self.filter_clear_btn.clicked.connect(self._clear_filters)

        self.filter_timer = QTimer()
        self.filter_timer.setSingleShot(True)
        self.filter_timer.timeout.connect(self.apply_filters)

        for f in [self.filter_evrak, self.filter_stok,
                   self.filter_malzeme, self.filter_tarih]:
            f.textChanged.connect(self._schedule_filter)

        self.btn_barkod.clicked.connect(self._on_toggle_btn)
        self.btn_manuel.clicked.connect(self._on_toggle_btn)

    def _schedule_filter(self):
        self.filter_timer.start(300)

    def _on_toggle_btn(self):
        btn = self.sender()
        color_map = {
            self.btn_barkod: '#22c55e',
            self.btn_manuel: '#ef4444',
        }
        color = color_map.get(btn, '#666')
        btn.setStyleSheet(_toggle_btn_style(color, btn.isChecked()))
        self._schedule_filter()

    def _clear_filters(self):
        for f in [self.filter_evrak, self.filter_stok,
                   self.filter_malzeme, self.filter_tarih]:
            f.clear()
        for btn, color in [(self.btn_barkod, '#22c55e'),
                           (self.btn_manuel, '#ef4444')]:
            btn.setChecked(False)
            btn.setStyleSheet(_toggle_btn_style(color, False))
        self.apply_filters()

    # ==================== SYNC ====================
    def start_sync(self):
        if not self.supabase_client:
            self.log("HATA: Supabase baglantisi yapilandirilmamis!")
            self.status_label.setText("Supabase ayarlari eksik")
            return

        if self.sync_thread and self.sync_thread.isRunning():
            self.log("Sync zaten calisiyor...")
            return

        self.log("Senkronizasyon basladi...")
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self._set_buttons_enabled(False)

        self.sync_thread = SevkSyncThread(self.supabase_client, self.dogtas_client)
        self.sync_thread.progress_updated.connect(self._on_sync_progress)
        self.sync_thread.sync_finished.connect(self._on_sync_finished)
        self.sync_thread.error_occurred.connect(self._on_sync_error)
        self.sync_thread.finished.connect(self._on_thread_finished)
        self.sync_thread.start()

    def _on_sync_progress(self, progress, message):
        self.progress_bar.setValue(progress)
        self.status_label.setText(message)
        self.log(message)

    def _on_sync_finished(self, result):
        self.last_sync_time = datetime.now()
        self.last_sync_label.setText(
            f"Son Sync: {self.last_sync_time.strftime('%Y-%m-%d %H:%M:%S')}"
        )
        mesaj = result['mesaj']
        if result.get('atlanan', 0) > 0:
            mesaj += f" ({result['atlanan']} atlandi)"
        self.log(f"Tamamlandi: {mesaj}")
        self.status_label.setText(mesaj)

        QTimer.singleShot(500, self.load_data)

    def _on_sync_error(self, error_message):
        self.log(f"HATA: {error_message}")
        self.status_label.setText(f"Sync hatasi: {error_message}")

    def _on_thread_finished(self):
        self._set_buttons_enabled(True)
        QTimer.singleShot(1000, lambda: self.progress_bar.setVisible(False))

    def _set_buttons_enabled(self, enabled: bool):
        self.sync_button.setEnabled(enabled)
        self.refresh_button.setEnabled(enabled)
        self.all_button.setEnabled(enabled)
        self.export_button.setEnabled(enabled)
        self.btn_tumunu.setEnabled(enabled)
        self.btn_sil.setEnabled(enabled)

    # ==================== TABLE ====================
    def load_data(self):
        if not self.supabase_client:
            self.status_label.setText("Supabase ayarlari eksik")
            return
        from datetime import timedelta
        one_week_ago = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
        self._start_load(min_tarih=one_week_ago)

    def load_all_data(self):
        if not self.supabase_client:
            self.status_label.setText("Supabase ayarlari eksik")
            return
        self._start_load(min_tarih=None)

    def _start_load(self, min_tarih=None):
        label = "Son 1 hafta" if min_tarih else "Tumu"
        self.status_label.setText(f"Sevk fisleri yukleniyor ({label})...")
        self.refresh_button.setEnabled(False)
        self.all_button.setEnabled(False)

        self._load_thread = SevkLoadDataThread(self.supabase_client, min_tarih=min_tarih)
        self._load_thread.data_loaded.connect(self._on_data_loaded)
        self._load_thread.error_occurred.connect(self._on_load_error)
        self._load_thread.start()

    def _on_data_loaded(self, all_data, readings_map):
        self.all_data = all_data
        self.readings_map = readings_map
        self.apply_filters()
        self.status_label.setText(f"{len(self.all_data)} kayit yuklendi")
        self.refresh_button.setEnabled(True)
        self.all_button.setEnabled(True)

    def _on_load_error(self, error_msg):
        self.status_label.setText(f"Yukleme hatasi: {error_msg}")
        self.log(f"HATA: Tablo yukleme hatasi: {error_msg}")
        self.refresh_button.setEnabled(True)
        self.all_button.setEnabled(True)

    def _get_row_colors(self, row_data):
        colors = set()
        kalem_id = row_data.get('id')
        paket_readings = self.readings_map.get(kalem_id, {})
        for ps, reads in paket_readings.items():
            for info in reads:
                if info['type'] == 'scanner':
                    colors.add('green')
                elif info['type'] == 'manual':
                    colors.add('red')
        return colors

    def apply_filters(self):
        filtered = self.all_data[:]

        evrak_text = self.filter_evrak.text().strip()
        stok_text = self.filter_stok.text().strip()
        malzeme_text = self.filter_malzeme.text().strip()
        tarih_text = self.filter_tarih.text().strip()

        if evrak_text:
            filtered = [r for r in filtered
                        if evrak_text in str(r.get('evrakno_sira', ''))]
        if stok_text:
            filtered = [r for r in filtered
                        if _fuzzy_match(stok_text, str(r.get('stok_kod', '')))]
        if malzeme_text:
            filtered = [r for r in filtered
                        if _fuzzy_match(malzeme_text, str(r.get('malzeme_adi', '')))]
        if tarih_text:
            filtered = [r for r in filtered
                        if str(r.get('tarih', '') or '') >= tarih_text]

        # Renk filtresi
        want_colors = set()
        if self.btn_barkod.isChecked():
            want_colors.add('green')
        if self.btn_manuel.isChecked():
            want_colors.add('red')
        if want_colors:
            filtered = [r for r in filtered
                        if self._get_row_colors(r) & want_colors]

        self.filtered_data = filtered
        self.populate_table()

    def populate_table(self):
        self.table.setUpdatesEnabled(False)
        self.table.setSortingEnabled(False)

        try:
            self.table.setRowCount(len(self.filtered_data))
            self.table.setColumnCount(len(SEVK_TABLE_COLUMNS) + 1)
            self.table.setHorizontalHeaderLabels([u'\u2611'] + [c[1] for c in SEVK_TABLE_COLUMNS])

            okuma_col_idx = next(
                (j + 1 for j, (k, _) in enumerate(SEVK_TABLE_COLUMNS) if k == 'okuma_durumu'),
                None
            )

            for i, row_data in enumerate(self.filtered_data):
                # Checkbox sutunu (index 0)
                chk_item = QTableWidgetItem()
                chk_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
                chk_item.setCheckState(Qt.Unchecked)
                chk_item.setData(Qt.UserRole, row_data.get('id'))
                self.table.setItem(i, 0, chk_item)

                for j, (key, _) in enumerate(SEVK_TABLE_COLUMNS):
                    if key == 'okuma_durumu':
                        continue
                    value = row_data.get(key, '')
                    if value is None:
                        text = ''
                    elif key == 'miktar':
                        num = float(value)
                        text = str(int(num)) if num == int(num) else str(num)
                    elif key == 'tarih':
                        text = str(value)[:10] if value else ''
                    else:
                        text = str(value)
                    item = QTableWidgetItem(text)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    font = QFont(FONT_FAMILY, FONT_SIZE)
                    font.setBold(True)
                    item.setFont(font)
                    self.table.setItem(i, j + 1, item)

                # Okuma Durumu sutunu
                if okuma_col_idx is not None:
                    miktar = int(float(row_data.get('miktar', 0) or 0))
                    paket = int(row_data.get('paket_sayisi', 1) or 1)
                    kalem_id = row_data.get('id')
                    paket_readings = self.readings_map.get(kalem_id, {})
                    depo_no = str(row_data.get('cikis_depo', '') or '')
                    widget = _build_okuma_durumu_widget(miktar, paket, paket_readings, depo_no)
                    self.table.setCellWidget(i, okuma_col_idx, widget)

            header = self.table.horizontalHeader()
            header.setMinimumSectionSize(0)
            header.setSectionResizeMode(QHeaderView.ResizeToContents)
            header.setStretchLastSection(False)
            header.setSectionResizeMode(0, QHeaderView.Fixed)
            self.table.setColumnWidth(0, 30)

            # Malzeme Adi sutunu ResizeToContents
            malzeme_adi_idx = next(
                (j + 1 for j, (k, _) in enumerate(SEVK_TABLE_COLUMNS) if k == 'malzeme_adi'),
                len(SEVK_TABLE_COLUMNS) - 1
            )
            header.setSectionResizeMode(malzeme_adi_idx, QHeaderView.ResizeToContents)

            if okuma_col_idx is not None:
                if self.table.columnWidth(okuma_col_idx) < 200:
                    self.table.setColumnWidth(okuma_col_idx, 200)

            for i in range(self.table.rowCount()):
                self.table.setRowHeight(i, ROW_HEIGHT)

        finally:
            self.table.setUpdatesEnabled(True)
            self.table.setSortingEnabled(True)

    # ==================== CHECKBOX ISLEMLERI ====================
    def _toggle_select_all_rows(self):
        any_checked = any(
            self.table.item(i, 0) and self.table.item(i, 0).checkState() == Qt.Checked
            for i in range(self.table.rowCount())
        )
        new_state = Qt.Unchecked if any_checked else Qt.Checked
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 0)
            if item:
                item.setCheckState(new_state)

    def _delete_selected_rows(self):
        if not self.supabase_client:
            _show_message(self, "Hata", "Supabase bağlantısı yok.")
            return
        ids = []
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 0)
            if item and item.checkState() == Qt.Checked:
                id_val = item.data(Qt.UserRole)
                if id_val is not None:
                    ids.append(id_val)
        if not ids:
            _show_message(self, "Bilgi", "Silinecek satır seçilmedi.")
            return
        if not _verify_barkod_delete_password(self):
            return
        if not _confirm_delete(self, f"{len(ids)} satır Supabase'den silinecek. Emin misiniz?"):
            return
        try:
            self.supabase_client.delete_sevk_by_id_list(ids)
            self.status_label.setText(f"{len(ids)} satır silindi.")
            self.load_data()
        except Exception as e:
            _show_message(self, "Hata", f"Silme hatası: {e}")

    # ==================== EXPORT ====================
    def export_to_excel(self):
        if not self.filtered_data:
            self.status_label.setText("Disari aktarilacak veri yok")
            return

        try:
            df = pd.DataFrame(self.filtered_data)
            output_path = "D:/GoogleDrive/~ SevkFisi_Export.xlsx"
            df.to_excel(output_path, index=False, engine='openpyxl')
            self.status_label.setText(f"Excel export: {output_path}")
            self.log(f"Veriler disari aktarildi: {output_path}")
        except Exception as e:
            self.status_label.setText(f"Export hatasi: {e}")
            self.log(f"HATA: Export hatasi: {e}")

    # ==================== KOPYALAMA ====================
    def show_context_menu(self, pos):
        menu = QMenu(self.table)
        menu.setStyleSheet("""
            QMenu {
                background-color: #ffffff;
                border: 1px solid #a0a0a0;
                padding: 4px 0px;
            }
            QMenu::item {
                padding: 6px 24px;
                color: #000000;
                font-size: 14px;
            }
            QMenu::item:selected {
                background-color: #3399ff;
                color: #ffffff;
            }
        """)
        hucre_action = QAction("Kopyala", self)
        hucre_action.triggered.connect(lambda: self.copy_cell(pos))
        menu.addAction(hucre_action)
        menu.exec_(self.table.viewport().mapToGlobal(pos))

    def copy_cell(self, pos):
        item = self.table.itemAt(pos)
        if item:
            QApplication.clipboard().setText(item.text())
            old_text = self.status_label.text()
            self.status_label.setText("✅ Kopyalandı")
            QTimer.singleShot(1500, lambda t=old_text: self.status_label.setText(t))

    def handle_ctrl_c(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            return
        if len(selected_items) == 1:
            QApplication.clipboard().setText(selected_items[0].text())
        else:
            rows = sorted({item.row() for item in selected_items})
            cols = sorted({item.column() for item in selected_items})
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
        old_text = self.status_label.text()
        self.status_label.setText("✅ Kopyalandı")
        QTimer.singleShot(1500, lambda t=old_text: self.status_label.setText(t))

    # ==================== LOG ====================
    def log(self, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.append(f"[{timestamp}] {message}")
        doc = self.log_text.document()
        if doc.blockCount() > 100:
            cursor = self.log_text.textCursor()
            cursor.movePosition(cursor.MoveOperation.Start)
            cursor.movePosition(cursor.MoveOperation.Down, cursor.MoveMode.KeepAnchor,
                                doc.blockCount() - 100)
            cursor.removeSelectedText()


# ================== SAYIM LOKASYON WIDGET ==================
class SayimLokasyonWidget(QWidget):
    u"""Say\u0131m lokasyon sekmesi - Belirli bir lokasyonun sayim verilerini gosterir"""

    def __init__(self, lokasyon, lokasyon_kodu):
        super().__init__()
        self.lokasyon = lokasyon          # DEPO / EXC / SUBE
        self.lokasyon_kodu = lokasyon_kodu  # 100 / 300 / 200
        self._data_loaded = False

        # Clients
        self.supabase_client = None
        self._init_clients()

        # Data
        self.all_data = []
        self.filtered_data = []
        self.readings_map = {}

        # UI
        self.setup_ui()
        self.setup_connections()

    def showEvent(self, event):
        super().showEvent(event)
        if not self._data_loaded:
            self._data_loaded = True
            QTimer.singleShot(100, self.load_data)

    def _init_clients(self):
        try:
            config_manager = CentralConfigManager()
            settings = config_manager.get_settings()

            supabase_url = (settings.get('Barkod_SUPABASE_URL') or
                            settings.get('SUPABASE_URL', ''))
            supabase_key = (settings.get('Barkod_SUPABASE_ANON_KEY') or
                            settings.get('SUPABASE_ANON_KEY', ''))

            if not supabase_url or not supabase_key:
                config_manager.settings_cache = {}
                settings = config_manager.get_settings(use_cache=False)
                supabase_url = (settings.get('Barkod_SUPABASE_URL') or
                                settings.get('SUPABASE_URL', ''))
                supabase_key = (settings.get('Barkod_SUPABASE_ANON_KEY') or
                                settings.get('SUPABASE_ANON_KEY', ''))

            if supabase_url and supabase_key:
                self.supabase_client = SupabaseClient(supabase_url, supabase_key)
            else:
                logger.warning(f"Sayim {self.lokasyon}: Supabase ayarlari eksik")

        except Exception as e:
            logger.error(f"Sayim {self.lokasyon} client init hatasi: {e}")

    # ==================== UI SETUP ====================
    def setup_ui(self):
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setObjectName("sayimLokasyonWidget")
        self.setStyleSheet("#sayimLokasyonWidget { background-color: #ffffff; }")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        layout.addWidget(self._create_header())
        layout.addWidget(self._create_filter_bar())

        # Table (SAYIM_TABLE_STYLE: ::item yok, setBackground icin)
        self.table = QTableWidget()
        self.table.setStyleSheet(SAYIM_TABLE_STYLE)
        self.table.setItemDelegate(NoFocusDelegate(self.table))
        self.table.setSelectionBehavior(QTableWidget.SelectItems)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.verticalHeader().setVisible(False)
        self.table.setFocusPolicy(Qt.StrongFocus)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        # Ctrl+C kısayolu - self üzerine bağlanır ki focus nerede olursa olsun çalışsın
        self.copy_shortcut = QShortcut(QKeySequence("Ctrl+C"), self)
        self.copy_shortcut.setContext(Qt.WidgetWithChildrenShortcut)
        self.copy_shortcut.activated.connect(self.handle_ctrl_c)
        layout.addWidget(self.table, 1)

        # Status bar
        self.status_label = QLabel("Hazir")
        self.status_label.setStyleSheet("QLabel { color: #6b7280; font-size: 12px; padding: 4px; }")
        layout.addWidget(self.status_label)

    def _create_header(self) -> QWidget:
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 0)

        self.refresh_button = QPushButton("Yenile")
        self.refresh_button.setStyleSheet(SYNC_BUTTON_STYLE)

        self.all_button = QPushButton("Hepsi")
        self.all_button.setStyleSheet(BUTTON_STYLE)

        self.export_button = QPushButton("Excel")
        self.export_button.setStyleSheet(BUTTON_STYLE)

        self.csv_button = QPushButton(".csv")
        self.csv_button.setStyleSheet(BUTTON_STYLE)

        self.btn_tumunu = QPushButton(u"T\u00fcm\u00fc")
        self.btn_tumunu.setStyleSheet(BUTTON_STYLE)

        self.btn_sil = QPushButton(u"Se\u00e7ilenleri Sil")
        self.btn_sil.setStyleSheet(BUTTON_STYLE)

        self.last_load_label = QLabel(u"Son Y\u00fckleme: -")
        self.last_load_label.setStyleSheet(INFO_LABEL_STYLE)

        header_layout.addWidget(self.refresh_button)
        header_layout.addWidget(self.all_button)
        header_layout.addWidget(self.export_button)
        header_layout.addWidget(self.csv_button)
        header_layout.addStretch()
        header_layout.addWidget(self.last_load_label)

        header_widget = QWidget()
        header_widget.setLayout(header_layout)
        return header_widget

    def _create_filter_bar(self) -> QWidget:
        filter_layout = QHBoxLayout()
        filter_layout.setContentsMargins(0, 0, 0, 0)

        self.filter_stok = QLineEdit()
        self.filter_stok.setPlaceholderText("Malzeme Kodu")
        self.filter_stok.setStyleSheet(FILTER_INPUT_STYLE)

        self.filter_malzeme = QLineEdit()
        self.filter_malzeme.setPlaceholderText("Malzeme")
        self.filter_malzeme.setStyleSheet(FILTER_INPUT_STYLE)

        # Fark filtre butonlari (radio mantigi: sadece biri secili olabilir)
        self.btn_tumu = QPushButton(u"T\u00fcm\u00fc")
        self.btn_tumu.setCheckable(True)
        self.btn_tumu.setStyleSheet(_toggle_btn_style('#3b82f6', False))

        self.btn_esit = QPushButton(u"E\u015fit")
        self.btn_esit.setCheckable(True)
        self.btn_esit.setStyleSheet(_toggle_btn_style('#22c55e', False))

        self.btn_eksik = QPushButton("Eksik")
        self.btn_eksik.setCheckable(True)
        self.btn_eksik.setStyleSheet(_toggle_btn_style('#ef4444', False))

        self.btn_fazla = QPushButton("Fazla")
        self.btn_fazla.setCheckable(True)
        self.btn_fazla.setStyleSheet(_toggle_btn_style('#ca8a04', False))

        self._fark_buttons = [self.btn_tumu, self.btn_esit, self.btn_eksik, self.btn_fazla]
        self._fark_btn_colors = {
            self.btn_tumu: '#3b82f6',
            self.btn_esit: '#22c55e',
            self.btn_eksik: '#ef4444',
            self.btn_fazla: '#ca8a04',
        }

        self.btn_barkod = QPushButton("Barkod Okunan")
        self.btn_barkod.setCheckable(True)
        self.btn_barkod.setStyleSheet(_toggle_btn_style('#22c55e', False))

        self.btn_manuel = QPushButton("Manuel")
        self.btn_manuel.setCheckable(True)
        self.btn_manuel.setStyleSheet(_toggle_btn_style('#f97316', False))

        self.filter_clear_btn = QPushButton("Temizle")
        self.filter_clear_btn.setStyleSheet(BUTTON_STYLE)

        filter_layout.addWidget(self.btn_tumunu)
        filter_layout.addWidget(self.btn_sil)
        filter_layout.addWidget(self.btn_tumu)
        filter_layout.addWidget(self.btn_esit)
        filter_layout.addWidget(self.btn_eksik)
        filter_layout.addWidget(self.btn_fazla)
        filter_layout.addWidget(self.filter_stok)
        filter_layout.addWidget(self.filter_malzeme)
        filter_layout.addWidget(self.btn_barkod)
        filter_layout.addWidget(self.btn_manuel)
        filter_layout.addWidget(self.filter_clear_btn)

        filter_widget = QWidget()
        filter_widget.setLayout(filter_layout)
        return filter_widget

    # ==================== CONNECTIONS ====================
    def setup_connections(self):
        self.refresh_button.clicked.connect(self.load_data)
        self.all_button.clicked.connect(self.load_all_data)
        self.export_button.clicked.connect(self.export_to_excel)
        self.csv_button.clicked.connect(self.export_to_csv)
        self.btn_tumunu.clicked.connect(self._toggle_select_all_rows)
        self.btn_sil.clicked.connect(self._delete_selected_rows)
        self.filter_clear_btn.clicked.connect(self._clear_filters)

        self.filter_timer = QTimer()
        self.filter_timer.setSingleShot(True)
        self.filter_timer.timeout.connect(self.apply_filters)

        for f in [self.filter_stok, self.filter_malzeme]:
            f.textChanged.connect(self._schedule_filter)

        for btn in self._fark_buttons:
            btn.clicked.connect(self._on_fark_btn)
        self.btn_barkod.clicked.connect(self._on_toggle_btn)
        self.btn_manuel.clicked.connect(self._on_toggle_btn)

    def _schedule_filter(self):
        self.filter_timer.start(300)

    def _on_fark_btn(self):
        """Fark butonlari radio mantigi: sadece biri secili olabilir"""
        clicked = self.sender()
        for btn in self._fark_buttons:
            if btn is clicked:
                btn.setChecked(True)
                btn.setStyleSheet(_toggle_btn_style(self._fark_btn_colors[btn], True))
            else:
                btn.setChecked(False)
                btn.setStyleSheet(_toggle_btn_style(self._fark_btn_colors[btn], False))
        self._schedule_filter()

    def _on_toggle_btn(self):
        btn = self.sender()
        color_map = {
            self.btn_barkod: '#22c55e',
            self.btn_manuel: '#f97316',
        }
        color = color_map.get(btn, '#666')
        btn.setStyleSheet(_toggle_btn_style(color, btn.isChecked()))
        self._schedule_filter()

    def _clear_filters(self):
        for f in [self.filter_stok, self.filter_malzeme]:
            f.clear()
        for btn in self._fark_buttons:
            btn.setChecked(False)
            btn.setStyleSheet(_toggle_btn_style(self._fark_btn_colors[btn], False))
        for btn, color in [(self.btn_barkod, '#22c55e'),
                           (self.btn_manuel, '#f97316')]:
            btn.setChecked(False)
            btn.setStyleSheet(_toggle_btn_style(color, False))
        self.apply_filters()

    # ==================== TABLE ====================
    def load_data(self):
        if not self.supabase_client:
            self.status_label.setText("Supabase ayarlari eksik")
            return
        from datetime import timedelta
        one_week_ago = (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
        self._start_load(min_tarih=one_week_ago)

    def load_all_data(self):
        if not self.supabase_client:
            self.status_label.setText("Supabase ayarlari eksik")
            return
        self._start_load(min_tarih=None)

    def _start_load(self, min_tarih=None):
        label = "Son 1 hafta" if min_tarih else "Tumu"
        self.status_label.setText(f"{self.lokasyon} sayim verileri yukleniyor ({label})...")
        self.refresh_button.setEnabled(False)
        self.all_button.setEnabled(False)

        self._load_thread = SayimLoadDataThread(
            self.supabase_client, min_tarih=min_tarih, lokasyon=self.lokasyon
        )
        self._load_thread.data_loaded.connect(self._on_data_loaded)
        self._load_thread.error_occurred.connect(self._on_load_error)
        self._load_thread.start()

    def _on_data_loaded(self, all_data, readings_map):
        self.all_data = all_data
        self.readings_map = readings_map
        self.last_load_label.setText(
            f"Son Y\u00fckleme: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        )
        self.apply_filters()
        self.status_label.setText(f"{self.lokasyon}: {len(self.all_data)} urun satiri yuklendi")
        self.refresh_button.setEnabled(True)
        self.all_button.setEnabled(True)

    def _on_load_error(self, error_msg):
        self.status_label.setText(f"Yukleme hatasi: {error_msg}")
        self.refresh_button.setEnabled(True)
        self.all_button.setEnabled(True)

    def _get_row_colors(self, row_data):
        colors = set()
        composite_key = row_data.get('composite_key')
        depo_no = str(row_data.get('lokasyon_kodu', '') or '')
        paket_readings = self.readings_map.get(composite_key, {})
        for ps, reads in paket_readings.items():
            for info in reads:
                if info is None:
                    continue
                if info['type'] == 'scanner':
                    colors.add('green')
                elif info['type'] == 'manual':
                    if depo_no == '100':
                        colors.add('red')
                    else:
                        colors.add('orange')
        return colors

    def apply_filters(self):
        filtered = self.all_data[:]

        stok_text = self.filter_stok.text().strip()
        malzeme_text = self.filter_malzeme.text().strip()

        if stok_text:
            filtered = [r for r in filtered
                        if _fuzzy_match(stok_text, str(r.get('malzeme_kodu', '')))
                        or _fuzzy_match(stok_text, str(r.get('stok_kod', '')))]
        if malzeme_text:
            filtered = [r for r in filtered
                        if _fuzzy_match(malzeme_text, str(r.get('malzeme_adi', '')))]

        # Fark filtresi (radio: Tumu / Esit / Eksik / Fazla)
        if self.btn_tumu.isChecked():
            # Sadece sayilan satirlari goster (sayim_kodu dolu)
            filtered = [r for r in filtered if r.get('sayim_kodu')]
        elif self.btn_esit.isChecked():
            filtered = [r for r in filtered if float(r.get('fark', 0) or 0) == 0 and r.get('sayim_kodu')]
        elif self.btn_eksik.isChecked():
            filtered = [r for r in filtered if float(r.get('fark', 0) or 0) < 0]
        elif self.btn_fazla.isChecked():
            filtered = [r for r in filtered if float(r.get('fark', 0) or 0) > 0]

        # Okuma renk filtresi
        want_colors = set()
        if self.btn_barkod.isChecked():
            want_colors.add('green')
        if self.btn_manuel.isChecked():
            want_colors.add('red')
            want_colors.add('orange')
        if want_colors:
            filtered = [r for r in filtered
                        if self._get_row_colors(r) & want_colors]

        # Okunan satirlari once goster (sayim_kodu bos olmayanlar)
        filtered.sort(key=lambda r: (0 if r.get('sayim_kodu') else 1))

        self.filtered_data = filtered
        self.populate_table()

    def populate_table(self):
        self.table.setUpdatesEnabled(False)
        self.table.setSortingEnabled(False)

        try:
            self.table.setRowCount(len(self.filtered_data))
            self.table.setColumnCount(len(SAYIM_TABLE_COLUMNS) + 1)
            self.table.setHorizontalHeaderLabels([u'\u2611'] + [c[1] for c in SAYIM_TABLE_COLUMNS])

            okuma_col_idx = next(
                (j + 1 for j, (k, _) in enumerate(SAYIM_TABLE_COLUMNS) if k == 'okuma_durumu'),
                None
            )

            for i, row_data in enumerate(self.filtered_data):
                # Sayim satiri icin oturum_id (composite_key'den al: "oturum_id::stok_kod")
                ck = row_data.get('composite_key', '')
                oturum_id = ck.split('::')[0] if '::' in ck and not ck.startswith('beklenen::') else None

                # Checkbox sutunu (index 0)
                chk_item = QTableWidgetItem()
                chk_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
                chk_item.setCheckState(Qt.Unchecked)
                chk_item.setData(Qt.UserRole, oturum_id)
                self.table.setItem(i, 0, chk_item)

                # Satir arka plan rengi belirle (fark bazli)
                fark_val = float(row_data.get('fark', 0) or 0)
                beklenen_val = float(row_data.get('beklenen', 0) or 0)
                if beklenen_val != 0 or fark_val != 0:
                    if fark_val < 0:
                        row_bg = QBrush(QColor('#fef2f2'))   # Acik kirmizi: eksik
                    elif fark_val > 0:
                        row_bg = QBrush(QColor('#fefce8'))   # Acik sari: fazla
                    else:
                        row_bg = QBrush(QColor('#f0fdf4'))   # Acik yesil: esit
                else:
                    row_bg = None  # Beklenen bilgisi yok

                for j, (key, _) in enumerate(SAYIM_TABLE_COLUMNS):
                    if key == 'okuma_durumu':
                        continue
                    value = row_data.get(key, '')
                    if value is None:
                        text = ''
                    elif key in ('miktar', 'beklenen'):
                        num = float(value) if value else 0
                        text = str(int(num)) if num == int(num) else str(num)
                    elif key == 'fark':
                        num = float(value) if value else 0
                        text = str(int(num)) if num == int(num) else str(num)
                        if num > 0:
                            text = f'+{text}'
                    else:
                        text = str(value)
                    item = QTableWidgetItem(text)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    font = QFont(FONT_FAMILY, FONT_SIZE)
                    font.setBold(True)
                    item.setFont(font)
                    # Fark sutunu yazi rengi
                    if key == 'fark' and (beklenen_val != 0 or fark_val != 0):
                        if fark_val < 0:
                            item.setForeground(QColor('#dc2626'))
                        elif fark_val > 0:
                            item.setForeground(QColor('#ca8a04'))
                        else:
                            item.setForeground(QColor('#16a34a'))
                    # Satir arka plan (Okuma Durumu haric)
                    if row_bg is not None:
                        item.setBackground(row_bg)
                    self.table.setItem(i, j + 1, item)

                # Okuma Durumu sutunu
                if okuma_col_idx is not None:
                    miktar = int(float(row_data.get('miktar', 0) or 0))
                    paket = int(row_data.get('paket_sayisi', 1) or 1)
                    composite_key = row_data.get('composite_key')
                    paket_readings = self.readings_map.get(composite_key, {})
                    depo_no = str(row_data.get('lokasyon_kodu', '') or '')
                    widget = _build_okuma_durumu_widget(miktar, paket, paket_readings, depo_no)
                    self.table.setCellWidget(i, okuma_col_idx, widget)

            header = self.table.horizontalHeader()
            header.setMinimumSectionSize(0)
            header.setSectionResizeMode(QHeaderView.ResizeToContents)
            header.setStretchLastSection(False)
            header.setSectionResizeMode(0, QHeaderView.Fixed)
            self.table.setColumnWidth(0, 30)

            malzeme_adi_idx = next(
                (j + 1 for j, (k, _) in enumerate(SAYIM_TABLE_COLUMNS) if k == 'malzeme_adi'),
                len(SAYIM_TABLE_COLUMNS) - 1
            )
            header.setSectionResizeMode(malzeme_adi_idx, QHeaderView.ResizeToContents)

            if okuma_col_idx is not None:
                if self.table.columnWidth(okuma_col_idx) < 200:
                    self.table.setColumnWidth(okuma_col_idx, 200)

            for i in range(self.table.rowCount()):
                self.table.setRowHeight(i, ROW_HEIGHT)

        finally:
            self.table.setUpdatesEnabled(True)
            self.table.setSortingEnabled(True)

    # ==================== CHECKBOX ISLEMLERI ====================
    def _toggle_select_all_rows(self):
        any_checked = any(
            self.table.item(i, 0) and self.table.item(i, 0).checkState() == Qt.Checked
            for i in range(self.table.rowCount())
        )
        new_state = Qt.Unchecked if any_checked else Qt.Checked
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 0)
            if item:
                item.setCheckState(new_state)

    def _delete_selected_rows(self):
        if not self.supabase_client:
            _show_message(self, "Hata", "Supabase bağlantısı yok.")
            return
        oturum_ids = []
        seen = set()
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 0)
            if item and item.checkState() == Qt.Checked:
                id_val = item.data(Qt.UserRole)
                if id_val is not None and id_val not in seen:
                    seen.add(id_val)
                    oturum_ids.append(id_val)
        if not oturum_ids:
            _show_message(self, "Bilgi", "Silinecek satır seçilmedi.")
            return
        if not _verify_barkod_delete_password(self):
            return
        if not _confirm_delete(self, f"{len(oturum_ids)} sayım oturumu Supabase'den silinecek. Emin misiniz?"):
            return
        try:
            self.supabase_client.delete_sayim_oturum_by_id_list(oturum_ids)
            self.status_label.setText(f"{len(oturum_ids)} oturum silindi.")
            self.load_data()
        except Exception as e:
            _show_message(self, "Hata", f"Silme hatası: {e}")

    # ==================== EXPORT ====================
    def export_to_excel(self):
        if not self.filtered_data:
            self.status_label.setText("Disari aktarilacak veri yok")
            return
        try:
            df = pd.DataFrame(self.filtered_data)
            output_path = f"D:/GoogleDrive/~ Sayim_{self.lokasyon}_Export.xlsx"
            df.to_excel(output_path, index=False, engine='openpyxl')
            self.status_label.setText(f"Excel export: {output_path}")
        except Exception as e:
            self.status_label.setText(f"Export hatasi: {e}")

    def export_to_csv(self):
        if not self.btn_eksik.isChecked() and not self.btn_fazla.isChecked():
            from PyQt5.QtWidgets import QMessageBox
            msg = QMessageBox(QMessageBox.Warning, "CSV Export",
                u"\u00d6nce 'Eksik' veya 'Fazla' butonuna basarak filtreleme yap\u0131n\u0131z.",
                QMessageBox.Ok)
            msg.setStyleSheet("QMessageBox { background-color: #f0f0f0; } "
                "QLabel { color: #000000; } "
                "QPushButton { background-color: #e0e0e0; color: #000000; padding: 6px 20px; }")
            msg.exec_()
            return
        # Sadece okunan satirlari al (sayim_kodu bos olmayanlar)
        csv_data = [r for r in self.filtered_data if r.get('sayim_kodu')]
        if not csv_data:
            self.status_label.setText("Disari aktarilacak veri yok")
            return
        try:
            output_path = f"D:/GoogleDrive/~ Sayim_{self.lokasyon}_Rapor.csv"
            with open(output_path, 'w', encoding='utf-8-sig') as f:
                f.write('Malzeme Kodu;Fark;Malzeme Adi\n')
                for row in csv_data:
                    mk = str(row.get('malzeme_kodu', '')).replace(';', ',')
                    ma = str(row.get('malzeme_adi', '')).replace(';', ',')
                    fark = float(row.get('fark', 0) or 0)
                    fark_abs = int(abs(fark))
                    f.write(f'{mk};{fark_abs};{ma}\n')
            self.status_label.setText(f"CSV export: {output_path} ({len(csv_data)} satir)")
        except Exception as e:
            self.status_label.setText(f"CSV export hatasi: {e}")

    # ==================== KOPYALAMA ====================
    def show_context_menu(self, pos):
        menu = QMenu(self.table)
        menu.setStyleSheet("""
            QMenu {
                background-color: #ffffff;
                border: 1px solid #a0a0a0;
                padding: 4px 0px;
            }
            QMenu::item {
                padding: 6px 24px;
                color: #000000;
                font-size: 14px;
            }
            QMenu::item:selected {
                background-color: #3399ff;
                color: #ffffff;
            }
        """)
        hucre_action = QAction("Kopyala", self)
        hucre_action.triggered.connect(lambda: self.copy_cell(pos))
        menu.addAction(hucre_action)
        menu.exec_(self.table.viewport().mapToGlobal(pos))

    def copy_cell(self, pos):
        item = self.table.itemAt(pos)
        if item:
            QApplication.clipboard().setText(item.text())
            old_text = self.status_label.text()
            self.status_label.setText("✅ Kopyalandı")
            QTimer.singleShot(1500, lambda t=old_text: self.status_label.setText(t))

    def handle_ctrl_c(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            return
        if len(selected_items) == 1:
            QApplication.clipboard().setText(selected_items[0].text())
        else:
            rows = sorted({item.row() for item in selected_items})
            cols = sorted({item.column() for item in selected_items})
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
        old_text = self.status_label.text()
        self.status_label.setText("✅ Kopyalandı")
        QTimer.singleShot(1500, lambda t=old_text: self.status_label.setText(t))



# ================== SAYIM WIDGET (CONTAINER) ==================
class SayimWidget(QWidget):
    u"""Say\u0131m sekmesi - Lokasyonlara gore alt sekmeler icerir"""

    def __init__(self):
        super().__init__()
        self.setStyleSheet("QWidget { background-color: #ffffff; }")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.tab_widget = QTabWidget()
        self.tab_widget.setTabPosition(QTabWidget.North)
        self.tab_widget.setStyleSheet(TAB_STYLE)

        self.depo_tab = SayimLokasyonWidget('DEPO', '100')
        self.tab_widget.addTab(self.depo_tab, "DEPO")

        self.exc_tab = SayimLokasyonWidget('EXC', '300')
        self.tab_widget.addTab(self.exc_tab, "EXC")

        self.sube_tab = SayimLokasyonWidget('SUBE', '200')
        self.tab_widget.addTab(self.sube_tab, "SUBE")

        layout.addWidget(self.tab_widget)


# ================== PLACEHOLDER WIDGET ==================
# ================== QR LOGLAMA THREAD + WIDGET ==================

class QrLogStokLoadThread(QThread):
    """PRGsheet'ten stok listesini yukler (Malzeme Adi, Malzeme Kodu, stok_kod)"""
    data_loaded = pyqtSignal(list)  # [{stok_kod, malzeme_adi, malzeme_kodu}, ...]
    error_occurred = pyqtSignal(str)

    def run(self):
        try:
            from io import BytesIO
            config_manager = CentralConfigManager()
            sid = config_manager.MASTER_SPREADSHEET_ID
            gsheets_url = f"https://docs.google.com/spreadsheets/d/{sid}/export?format=xlsx"
            resp = requests.get(gsheets_url, timeout=30)
            resp.raise_for_status()
            stok_df = pd.read_excel(BytesIO(resp.content), sheet_name="Stok")

            kod_col = stok_df.columns[0]
            malzeme_adi_col = u'Malzeme Ad\u0131'
            malzeme_kodu_col = 'Malzeme Kodu'

            stok_list = []
            for _, srow in stok_df.iterrows():
                kod = str(srow[kod_col]).strip() if pd.notna(srow[kod_col]) else ''
                if not kod:
                    continue
                mal_adi = str(srow[malzeme_adi_col]) if malzeme_adi_col in stok_df.columns and pd.notna(srow.get(malzeme_adi_col)) else ''
                mal_kodu = str(srow[malzeme_kodu_col]).strip() if malzeme_kodu_col in stok_df.columns and pd.notna(srow.get(malzeme_kodu_col)) else ''
                stok_list.append({
                    'stok_kod': kod,
                    'malzeme_adi': mal_adi,
                    'malzeme_kodu': mal_kodu,
                })
            logger.info(f"QR Log: {len(stok_list)} urun yuklendi (PRGsheet)")
            self.data_loaded.emit(stok_list)
        except Exception as e:
            logger.error(f"QR Log stok yukleme hatasi: {e}")
            self.error_occurred.emit(str(e))


class QrLogSearchThread(QThread):
    """Bir urunun 6 tablodan okuma gecmisini arar"""
    results_loaded = pyqtSignal(list)
    error_occurred = pyqtSignal(str)

    def __init__(self, supabase_client, stok_kod):
        super().__init__()
        self.supabase_client = supabase_client
        self.stok_kod = stok_kod

    def run(self):
        try:
            results = self.supabase_client.get_qr_log_by_stok_kod(self.stok_kod)
            self.results_loaded.emit(results)
        except Exception as e:
            logger.error(f"QR Log arama hatasi: {e}")
            self.error_occurred.emit(str(e))


class QrLogWidget(QWidget):
    u"""QR Loglama sekmesi - Bir urunun tum barkod gecmisini gosterir"""

    KAYNAK_COLORS = {
        'Nakliye': '#3b82f6',    # mavi
        u'Giri\u015f': '#22c55e',  # yesil
        'Sevk': '#f97316',       # turuncu
        u'Sat\u0131\u015f': '#ef4444',  # kirmizi
        u'\u00c7\u0131k\u0131\u015f': '#a855f7',  # mor
        u'Say\u0131m': '#6b7280',  # gri
    }

    def __init__(self):
        super().__init__()
        self._stok_loaded = False
        self.stok_list = []       # [{stok_kod, malzeme_adi, malzeme_kodu}, ...]
        self.supabase_client = None
        self._init_clients()
        self.setup_ui()
        self.setup_connections()

    def showEvent(self, event):
        super().showEvent(event)
        if not self._stok_loaded:
            self._stok_loaded = True
            QTimer.singleShot(100, self._load_stok_list)

    def _init_clients(self):
        try:
            config_manager = CentralConfigManager()
            settings = config_manager.get_settings()
            supabase_url = (settings.get('Barkod_SUPABASE_URL') or
                            settings.get('SUPABASE_URL', ''))
            supabase_key = (settings.get('Barkod_SUPABASE_ANON_KEY') or
                            settings.get('SUPABASE_ANON_KEY', ''))
            if not supabase_url or not supabase_key:
                config_manager.settings_cache = {}
                settings = config_manager.get_settings(use_cache=False)
                supabase_url = (settings.get('Barkod_SUPABASE_URL') or
                                settings.get('SUPABASE_URL', ''))
                supabase_key = (settings.get('Barkod_SUPABASE_ANON_KEY') or
                                settings.get('SUPABASE_ANON_KEY', ''))
            if supabase_url and supabase_key:
                self.supabase_client = SupabaseClient(supabase_url, supabase_key)
        except Exception as e:
            logger.error(f"QR Log client init hatasi: {e}")

    def setup_ui(self):
        self.setObjectName("qrLogWidget")
        self.setStyleSheet("#qrLogWidget { background-color: #ffffff; }")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)

        # --- Arama bölümü ---
        search_layout = QHBoxLayout()

        lbl = QLabel("Malzeme:")
        lbl.setStyleSheet("font-size: 14px; font-weight: bold; color: #000;")
        search_layout.addWidget(lbl)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText(u"Malzeme kodu veya ad\u0131 yazarak aray\u0131n...")
        self.search_input.setStyleSheet(FILTER_INPUT_STYLE)
        search_layout.addWidget(self.search_input, 1)

        self.search_btn = QPushButton("Ara")
        self.search_btn.setStyleSheet("""
            QPushButton {
                background-color: #3b82f6; color: white; font-weight: bold;
                padding: 8px 20px; border-radius: 4px; font-size: 14px;
            }
            QPushButton:hover { background-color: #2563eb; }
        """)
        search_layout.addWidget(self.search_btn)

        self.search_clear_btn = QPushButton("Temizle")
        self.search_clear_btn.setStyleSheet(BUTTON_STYLE)
        search_layout.addWidget(self.search_clear_btn)

        self.btn_qr_esle = QPushButton(u"QR E\u015fle")
        self.btn_qr_esle.setCheckable(True)
        self.btn_qr_esle.setStyleSheet("""
            QPushButton {
                background-color: #dcfce7; color: black; border: 1px solid #444;
                padding: 8px 16px; border-radius: 5px; font-size: 14px; font-weight: bold;
                min-width: 80px;
            }
            QPushButton:hover { background-color: #bbf7d0; }
            QPushButton:checked { background-color: #86efac; border: 1px solid #16a34a;
                padding: 8px 16px; min-width: 80px; }
        """)
        search_layout.addWidget(self.btn_qr_esle)

        self.export_button = QPushButton("Excel")
        self.export_button.setStyleSheet(BUTTON_STYLE)
        search_layout.addWidget(self.export_button)

        search_widget = QWidget()
        search_widget.setLayout(search_layout)
        layout.addWidget(search_widget)

        # --- Ürün listesi (dropdown) ---
        self.product_list = QListWidget()
        self.product_list.setMaximumHeight(200)
        self.product_list.setStyleSheet("""
            QListWidget {
                font-size: 14px; background-color: #ffffff; color: #000000;
                border: 1px solid #d0d0d0; padding: 4px;
            }
            QListWidget::item { padding: 6px; }
            QListWidget::item:selected { background-color: #3b82f6; color: white; }
            QListWidget::item:hover { background-color: #e0e7ff; }
        """)
        self.product_list.setVisible(False)
        layout.addWidget(self.product_list)

        # --- Seçilen ürün bilgisi ---
        self.selected_label = QLabel("")
        self.selected_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #1e40af; padding: 4px;")
        layout.addWidget(self.selected_label)

        # --- Sonuç tablosu ---
        self.table = QTableWidget()
        self.table.setStyleSheet(SAYIM_TABLE_STYLE)
        self.table.setItemDelegate(NoFocusDelegate(self.table))
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        self.table.setColumnCount(len(QR_LOG_TABLE_COLUMNS))
        self.table.setHorizontalHeaderLabels([c[1] for c in QR_LOG_TABLE_COLUMNS])
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._context_menu)
        layout.addWidget(self.table, 1)

        # --- Alt bilgilendirme çubuğu ---
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("font-size: 13px; color: #333; padding: 4px;")
        layout.addWidget(self.status_label)

    def setup_connections(self):
        self.search_input.textChanged.connect(self._on_search_text_changed)
        self.search_btn.clicked.connect(self._on_search_btn)
        self.search_clear_btn.clicked.connect(self._on_search_clear)
        self.product_list.itemClicked.connect(self._on_product_selected)
        self.search_input.returnPressed.connect(self._on_enter_pressed)
        self.btn_qr_esle.clicked.connect(self._on_qr_esle)
        self.export_button.clicked.connect(self.export_to_excel)

    # ==================== STOK YÜKLEME ====================
    def _load_stok_list(self):
        self.status_label.setText(u"\u00dcr\u00fcn listesi y\u00fckl\u00fceniyor...")
        self._stok_thread = QrLogStokLoadThread()
        self._stok_thread.data_loaded.connect(self._on_stok_loaded)
        self._stok_thread.error_occurred.connect(self._on_stok_error)
        self._stok_thread.start()

    def _on_stok_loaded(self, stok_list):
        self.stok_list = stok_list
        self.status_label.setText(f"{len(stok_list)} urun yuklendi")

    def _on_stok_error(self, err):
        self.status_label.setText(f"Stok yukleme hatasi: {err}")

    # ==================== ARAMA ====================
    def _on_search_text_changed(self, text=None):
        text = self.search_input.text().strip()
        if len(text) < 2:
            self.product_list.setVisible(False)
            return
        # Fuzzy match — malzeme kodu veya adi
        matches = []
        for item in self.stok_list:
            if _fuzzy_match(text, item.get('malzeme_kodu', '')) or _fuzzy_match(text, item.get('malzeme_adi', '')):
                display = f"{item.get('malzeme_kodu', '')}  -  {item['malzeme_adi']}"
                matches.append((display, item['stok_kod']))
            if len(matches) >= 50:
                break
        self.product_list.clear()
        if matches:
            for display, sk in matches:
                li = QListWidgetItem(display)
                li.setData(Qt.UserRole, sk)
                self.product_list.addItem(li)
            self.product_list.setVisible(True)
        else:
            self.product_list.setVisible(False)

    def _on_enter_pressed(self):
        if self.product_list.isVisible() and self.product_list.count() > 0:
            self.product_list.setCurrentRow(0)
            self._on_product_selected(self.product_list.item(0))

    def _on_search_btn(self):
        self._on_enter_pressed()

    def _on_search_clear(self):
        self.search_input.clear()
        self.product_list.setVisible(False)
        self.btn_qr_esle.setChecked(False)
        self.selected_label.setText("")
        self.table.setRowCount(0)
        self._all_results = []

    def _on_product_selected(self, item):
        stok_kod = item.data(Qt.UserRole)
        display = item.text()
        self.product_list.setVisible(False)
        self.search_input.blockSignals(True)
        self.search_input.setText(display)
        self.search_input.blockSignals(False)
        self.selected_label.setText(f"Araniyor: {display} (stok_kod: {stok_kod})")
        self._search_qr_log(stok_kod)

    def _search_qr_log(self, stok_kod):
        if not self.supabase_client:
            self.status_label.setText("Supabase ayarlari eksik")
            return
        self.status_label.setText("Arama yapiliyor...")
        self.table.setRowCount(0)
        self._search_thread = QrLogSearchThread(self.supabase_client, stok_kod)
        self._search_thread.results_loaded.connect(self._on_results_loaded)
        self._search_thread.error_occurred.connect(self._on_search_error)
        self._search_thread.start()

    def _on_results_loaded(self, results):
        self._all_results = results  # tum sonuclari sakla (filtre icin)
        self.btn_qr_esle.setChecked(False)
        self.status_label.setText(f"{len(results)} okuma kaydi bulundu")
        self._populate_table(results)

    def _on_qr_esle(self):
        """QR Esle: eslesen +/- satirlari gizle, eslesmeyenler kalsin"""
        if not hasattr(self, '_all_results'):
            return
        checked = self.btn_qr_esle.isChecked()
        if checked:
            # + ve - tablolarindaki qr_kodlari bul
            qr_giris = set()
            qr_cikis = set()
            for r in self._all_results:
                qr = r.get('qr_kod', '')
                if not qr or qr.startswith('MANUEL_TOPLU_'):
                    continue
                if r.get('yon') == '+':
                    qr_giris.add(qr)
                elif r.get('yon') == '-':
                    qr_cikis.add(qr)
            qr_tamamlanan = qr_giris & qr_cikis
            # Eslesmeyenleri goster
            filtered = [r for r in self._all_results
                        if r.get('qr_kod', '') not in qr_tamamlanan]
            self._populate_table(filtered)
        else:
            self._populate_table(self._all_results)

    def export_to_excel(self):
        if not hasattr(self, '_all_results') or not self._all_results:
            self.status_label.setText("Disari aktarilacak veri yok")
            return
        try:
            # Tabloda gorunen veriyi export et
            if self.btn_qr_esle.isChecked():
                qr_giris = set()
                qr_cikis = set()
                for r in self._all_results:
                    qr = r.get('qr_kod', '')
                    if not qr or qr.startswith('MANUEL_TOPLU_'):
                        continue
                    if r.get('yon') == '+':
                        qr_giris.add(qr)
                    elif r.get('yon') == '-':
                        qr_cikis.add(qr)
                qr_tamamlanan = qr_giris & qr_cikis
                export_data = [r for r in self._all_results
                               if r.get('qr_kod', '') not in qr_tamamlanan]
            else:
                export_data = self._all_results

            df = pd.DataFrame(export_data)
            output_path = "D:/GoogleDrive/~ QR_Loglama_Export.xlsx"
            df.to_excel(output_path, index=False, engine='openpyxl')
            self.status_label.setText(f"Excel export: {output_path} ({len(export_data)} satir)")
        except Exception as e:
            self.status_label.setText(f"Export hatasi: {e}")

    def _on_search_error(self, err):
        self.status_label.setText(f"Arama hatasi: {err}")

    # ==================== TABLO ====================
    def _populate_table(self, results):
        self.table.setUpdatesEnabled(False)
        self.table.setSortingEnabled(False)
        try:
            # QR kodlarin + ve - durumlarini belirle
            qr_giris = set()   # + tablolarinda gecen qr_kodlar
            qr_cikis = set()   # - tablolarinda gecen qr_kodlar
            for r in results:
                qr = r.get('qr_kod', '')
                if not qr or qr.startswith('MANUEL_TOPLU_'):
                    continue
                if r.get('yon') == '+':
                    qr_giris.add(qr)
                elif r.get('yon') == '-':
                    qr_cikis.add(qr)
            # Hem girisi hem cikisi olan = tamamlanmis
            qr_tamamlanan = qr_giris & qr_cikis

            # LIFO siralama: tamamlanan QR kodlar uste,
            # son cikis tarihi en yeni olan en uste
            qr_last_cikis = {}  # qr_kod -> en son cikis tarihi
            for r in results:
                qr = r.get('qr_kod', '')
                if qr in qr_tamamlanan and r.get('yon') == '-':
                    tarih = r.get('tarih', '')
                    if tarih > qr_last_cikis.get(qr, ''):
                        qr_last_cikis[qr] = tarih

            def lifo_sort_key(r):
                qr = r.get('qr_kod', '')
                is_tam = qr in qr_tamamlanan
                # 0 = tamamlanan (uste), 1 = diger
                # Tamamlananlar icinde: son cikis tarihi buyuk olan uste (desc)
                last_cikis = qr_last_cikis.get(qr, '')
                return (0 if is_tam else 1, '' if not last_cikis else chr(255) - last_cikis[0] if last_cikis else '', r.get('tarih', '') or '')

            # Basit LIFO: tamamlananlar uste, kendi icinde qr_kod gruplari
            tamamlanan_rows = []
            diger_rows = []
            for r in results:
                qr = r.get('qr_kod', '')
                if qr in qr_tamamlanan:
                    tamamlanan_rows.append(r)
                else:
                    diger_rows.append(r)
            # Tamamlanmamislar uste: paket kucukten buyuge, ayni pakette tarih yeniden eskiye
            diger_rows.sort(key=lambda r: (
                int(r.get('paket_sira', 0) or 0),
                ''.join(chr(255 - ord(c)) for c in (r.get('tarih', '') or ''))
            ))
            # Tamamlananlari son cikis tarihine gore sirala (en yeni uste)
            tamamlanan_rows.sort(key=lambda r: qr_last_cikis.get(r.get('qr_kod', ''), ''), reverse=True)
            results = diger_rows + tamamlanan_rows

            self.table.setRowCount(len(results))
            for row_idx, row in enumerate(results):
                qr = row.get('qr_kod', '')
                is_tamamlanan = qr in qr_tamamlanan

                for col_idx, (key, _) in enumerate(QR_LOG_TABLE_COLUMNS):
                    val = str(row.get(key, '') or '')

                    # Tarih formatla
                    if key == 'tarih' and val:
                        try:
                            from datetime import datetime as dt, timedelta, timezone
                            dt_obj = dt.fromisoformat(val.replace('Z', '+00:00'))
                            dt_tr = dt_obj.astimezone(timezone(timedelta(hours=3)))
                            val = dt_tr.strftime('%d.%m.%Y %H:%M')
                        except Exception:
                            val = val[:16]

                    item = QTableWidgetItem(val)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    item.setForeground(QBrush(QColor('#000000')))

                    # Tamamlanmis satirlari acik yesil
                    if is_tamamlanan and key != 'kaynak':
                        item.setBackground(QBrush(QColor('#dcfce7')))

                    # Yon sutunu renklendir
                    if key == 'yon':
                        item.setTextAlignment(Qt.AlignCenter)
                        if val == '+':
                            item.setForeground(QBrush(QColor('#16a34a')))
                        elif val == '-':
                            item.setForeground(QBrush(QColor('#dc2626')))

                    # Kaynak renkli arka plan
                    if key == 'kaynak':
                        color = self.KAYNAK_COLORS.get(val, '#666666')
                        item.setForeground(QBrush(QColor('#ffffff')))
                        item.setBackground(QBrush(QColor(color)))
                        item.setTextAlignment(Qt.AlignCenter)

                    if key == 'paket_sira':
                        ps = row.get('paket_sira', '')
                        if ps:
                            item.setText(f"P{ps}")
                        item.setTextAlignment(Qt.AlignCenter)

                    # QR kod tam halini data olarak sakla
                    if key == 'qr_kod_kisa':
                        item.setData(Qt.UserRole, row.get('qr_kod', ''))
                        item.setToolTip(row.get('qr_kod', ''))

                    self.table.setItem(row_idx, col_idx, item)

            self.table.resizeColumnsToContents()

            # Istatistik
            tamamlanan_count = len(qr_tamamlanan)
            toplam_qr = len(qr_giris | qr_cikis)
            if toplam_qr > 0:
                self.status_label.setText(
                    f"{len(results)} okuma kaydi | {tamamlanan_count}/{toplam_qr} QR tamamlandi")
            else:
                self.status_label.setText(f"{len(results)} okuma kaydi bulundu")
        finally:
            self.table.setSortingEnabled(True)
            self.table.setUpdatesEnabled(True)

    # ==================== CONTEXT MENU ====================
    def _context_menu(self, pos):
        menu = QMenu(self.table)
        menu.setStyleSheet("""
            QMenu { background-color: #ffffff; border: 1px solid #a0a0a0; color: #000000; }
            QMenu::item:selected { background-color: #b3d9ff; color: #000000; }
        """)
        copy_action = menu.addAction("Kopyala")
        copy_qr_action = menu.addAction(u"QR Kodu Kopyala (tam)")

        action = menu.exec_(self.table.viewport().mapToGlobal(pos))
        if action == copy_action:
            self._copy_selected()
        elif action == copy_qr_action:
            self._copy_qr_code()

    def _copy_selected(self):
        from PyQt5.QtWidgets import QApplication
        items = self.table.selectedItems()
        if items:
            texts = [item.text() for item in items]
            QApplication.clipboard().setText('\t'.join(texts))

    def _copy_qr_code(self):
        from PyQt5.QtWidgets import QApplication
        row = self.table.currentRow()
        if row >= 0:
            qr_col = next(i for i, (k, _) in enumerate(QR_LOG_TABLE_COLUMNS) if k == 'qr_kod_kisa')
            item = self.table.item(row, qr_col)
            if item:
                qr_full = item.data(Qt.UserRole) or item.text()
                QApplication.clipboard().setText(qr_full)


class PlaceholderWidget(QWidget):
    """Henuz icerigi olmayan sekmeler icin placeholder"""

    def __init__(self, sekme_adi: str):
        super().__init__()
        self.setStyleSheet("QWidget { background-color: #ffffff; }")
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)

        label = QLabel(f"{sekme_adi}\n\nYakim zamanda eklenecek...")
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet(PLACEHOLDER_STYLE)
        layout.addWidget(label)


# ================== MAIN MODULE WIDGET (TABBED) ==================
class BarkodApp(QWidget):
    """Barkod modulu - Sekmeli ana konteyner"""

    def __init__(self):
        super().__init__()
        self.setStyleSheet("QWidget { background-color: #ffffff; }")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Tab widget
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabPosition(QTabWidget.North)
        self.tab_widget.setStyleSheet(TAB_STYLE)

        # Sekmeler
        self.satis_tab = SatisTeslimatWidget()
        self.tab_widget.addTab(self.satis_tab, u"Sat\u0131\u015f / Teslimat Fi\u015fi")
        self.fabrika_nakliye_tab = FabrikaNakliyePlanWidget()
        self.tab_widget.addTab(self.fabrika_nakliye_tab, u"Fabrika Nakliye Plan\u0131")
        self.nakliye_tab = NakliyeYuklemeWidget()
        self.tab_widget.addTab(self.nakliye_tab, u"Nakliye Y\u00fckleme")
        self.sevk_tab = SevkFisiWidget()
        self.tab_widget.addTab(self.sevk_tab, u"Depolar Aras\u0131 Sevk")
        self.qr_log_tab = QrLogWidget()
        self.tab_widget.addTab(self.qr_log_tab, "QR Loglama")
        self.sayim_tab = SayimWidget()
        self.tab_widget.addTab(self.sayim_tab, u"Say\u0131m")
        self.giris_tab = GirisFisiWidget()
        self.tab_widget.addTab(self.giris_tab, u"Di\u011fer Giri\u015fler")
        self.cikis_tab = CikisFisiWidget()
        self.tab_widget.addTab(self.cikis_tab, u"Di\u011fer \u00c7\u0131k\u0131\u015flar")

        layout.addWidget(self.tab_widget)
