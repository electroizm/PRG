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
                             QApplication, QMenu, QAction, QShortcut,
                             QStyledItemDelegate, QStyle)
from PyQt5.QtGui import QFont, QKeySequence

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

MIKRO_SQL_QUERY = """
    SELECT
        sth.sth_evrakno_seri,
        sth.sth_evrakno_sira,
        sth.sth_satirno,
        CONVERT(DATE, sth.sth_tarih) AS tarih,
        sth.sth_stok_kod,
        sth.sth_miktar,
        sth.sth_cikis_depo_no,
        dbo.fn_StokHarEvrTip(sth.sth_evraktip) AS evrak_adi,
        cha.cha_kod AS cari_kodu,
        dbo.fn_CarininIsminiBul(cha.cha_cari_cins, cha.cha_kod) AS cari_adi,
        bar.bar_serino_veya_bagkodu AS bag_kodu,
        sto.sto_isim AS malzeme_adi
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
        AND sth.sth_tarih >= ?
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
        select_cols = 'id,evrakno_seri,evrakno_sira,satirno,tarih,stok_kod,miktar,cikis_depo_no,paket_sayisi,cari_kodu,cari_adi,product_desc,malzeme_adi,evrak_adi,bag_kodu'
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
                    'bag_kodu': fatura.get('bag_kodu'),
                    'malzeme_adi': fatura.get('malzeme_adi'),
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
            WHERE sth_evraktip = 4 AND sth_tarih >= ?
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
        self.copy_shortcut = QShortcut(QKeySequence("Ctrl+C"), self.table)
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
            self.table.setColumnCount(len(TABLE_COLUMNS))
            self.table.setHorizontalHeaderLabels([c[1] for c in TABLE_COLUMNS])

            # Okuma Durumu sutun index'i
            okuma_col_idx = next(
                (j for j, (k, _) in enumerate(TABLE_COLUMNS) if k == 'okuma_durumu'),
                None
            )

            for i, row_data in enumerate(self.filtered_data):
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
                    self.table.setItem(i, j, item)

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
            header.setSectionResizeMode(QHeaderView.Interactive)
            header.setStretchLastSection(False)
            self.table.resizeColumnsToContents()

            # Urun Aciklama sutunu stretch
            product_desc_idx = next(
                (j for j, (k, _) in enumerate(TABLE_COLUMNS) if k == 'product_desc'),
                len(TABLE_COLUMNS) - 2
            )
            header.setSectionResizeMode(product_desc_idx, QHeaderView.Stretch)

            # Okuma Durumu sutunu minimum genislik
            if okuma_col_idx is not None:
                if self.table.columnWidth(okuma_col_idx) < 200:
                    self.table.setColumnWidth(okuma_col_idx, 200)

            for i in range(self.table.rowCount()):
                self.table.setRowHeight(i, ROW_HEIGHT)

        finally:
            self.table.setUpdatesEnabled(True)
            self.table.setSortingEnabled(True)

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
        self.copy_shortcut = QShortcut(QKeySequence("Ctrl+C"), self.table)
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
            self.table.setColumnCount(len(NAKLIYE_TABLE_COLUMNS))
            self.table.setHorizontalHeaderLabels([c[1] for c in NAKLIYE_TABLE_COLUMNS])

            okuma_col_idx = next(
                (j for j, (k, _) in enumerate(NAKLIYE_TABLE_COLUMNS) if k == 'okuma_durumu'),
                None
            )

            for i, row_data in enumerate(self.filtered_data):
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
                    self.table.setItem(i, j, item)

                if okuma_col_idx is not None:
                    miktar_raw = str(row_data.get('miktar', '0') or '0').replace(',', '.')
                    miktar = max(1, int(float(miktar_raw)))
                    paket = int(row_data.get('paket_sayisi', 1) or 1)
                    kalem_id = row_data.get('id')
                    paket_readings = self.readings_map.get(kalem_id, {})
                    widget = _build_okuma_durumu_widget(miktar, paket, paket_readings, depo_no='100')
                    self.table.setCellWidget(i, okuma_col_idx, widget)

            header = self.table.horizontalHeader()
            header.setSectionResizeMode(QHeaderView.Interactive)
            header.setStretchLastSection(False)
            self.table.resizeColumnsToContents()

            malzeme_adi_idx = next(
                (j for j, (k, _) in enumerate(NAKLIYE_TABLE_COLUMNS) if k == 'malzeme_adi'),
                len(NAKLIYE_TABLE_COLUMNS) - 2
            )
            header.setSectionResizeMode(malzeme_adi_idx, QHeaderView.Stretch)

            if okuma_col_idx is not None:
                if self.table.columnWidth(okuma_col_idx) < 200:
                    self.table.setColumnWidth(okuma_col_idx, 200)

            for i in range(self.table.rowCount()):
                self.table.setRowHeight(i, ROW_HEIGHT)

        finally:
            self.table.setUpdatesEnabled(True)
            self.table.setSortingEnabled(True)

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


# ================== PLACEHOLDER WIDGET ==================
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
        self.nakliye_tab = NakliyeYuklemeWidget()
        self.tab_widget.addTab(self.nakliye_tab, u"Nakliye Y\u00fckleme")
        self.tab_widget.addTab(PlaceholderWidget(u"Depolar Aras\u0131 Sevk"), u"Depolar Aras\u0131 Sevk")
        self.tab_widget.addTab(PlaceholderWidget(u"Say\u0131m"), u"Say\u0131m")
        self.tab_widget.addTab(PlaceholderWidget(u"Di\u011fer Giri\u015fler"), u"Di\u011fer Giri\u015fler")
        self.tab_widget.addTab(PlaceholderWidget(u"Di\u011fer \u00c7\u0131k\u0131\u015flar"), u"Di\u011fer \u00c7\u0131k\u0131\u015flar")

        layout.addWidget(self.tab_widget)
