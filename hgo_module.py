"""
HGO Modulu - Dogtas Bayi Ticari Prim Hesaplama
================================================

Kullanici tarafindan girilen aylik hedef verileri ile Dogtas API'sinden
cekilen gerceklesen siparis ve fatura verilerini karsilastirarak,
Dogtas 2026 Q1 Ticari Politikasi'na uygun prim hakedislerini hesaplar.
"""

import sys
import os
import requests
import logging
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path

# Ust dizini Python path'e ekle (central_config icin)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from central_config import CentralConfigManager

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QComboBox, QLineEdit, QPushButton, QTableWidget,
    QTableWidgetItem, QProgressBar, QGroupBox, QHeaderView,
    QMessageBox, QTextEdit, QTabWidget
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QColor

# ============================================================================
# LOGGING
# ============================================================================

if getattr(sys, 'frozen', False):
    _base_dir = Path(sys.executable).parent
else:
    _base_dir = Path(__file__).parent

_log_dir = _base_dir / 'logs'
_log_dir.mkdir(exist_ok=True)
_log_file = _log_dir / 'prim_hesaplama.log'

logging.basicConfig(
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(_log_file, encoding='utf-8')
    ]
)
_logger = logging.getLogger(__name__)

# ============================================================================
# SABITLER
# ============================================================================

TURKISH_MONTHS = {
    1: 'Ocak', 2: '\u015eubat', 3: 'Mart', 4: 'Nisan',
    5: 'May\u0131s', 6: 'Haziran', 7: 'Temmuz', 8: 'A\u011fustos',
    9: 'Eyl\u00fcl', 10: 'Ekim', 11: 'Kas\u0131m', 12: 'Aral\u0131k'
}


# ============================================================================
# PRIM API CLIENT
# ============================================================================

class PrimApiClient:
    """Dogtas API uzerinden siparis ve fatura verisi ceken istemci."""

    def __init__(self, config_manager: CentralConfigManager):
        self.config_manager = config_manager
        self.token = None
        self._load_config()

    def _load_config(self):
        try:
            sheet = self.config_manager.gc.open("PRGsheet").worksheet('Ayar')
            all_values = sheet.get_all_values()

            if not all_values:
                raise ValueError("Ayar sayfasi bos")

            headers = all_values[0]
            key_index = headers.index('Key')
            value_index = headers.index('Value')

            config = {}
            for row in all_values[1:]:
                if len(row) > max(key_index, value_index):
                    key = row[key_index].strip() if row[key_index] else ''
                    value = row[value_index].strip() if row[value_index] else ''
                    if key:
                        config[key] = value

            self.base_url = config.get('base_url', '')
            self.endpoint = config.get('bekleyenler', '')
            self.customer_no = config.get('CustomerNo', '')

            self.auth_data = {
                "userName": config.get('userName', ''),
                "password": config.get('password', ''),
                "clientId": config.get('clientId', ''),
                "clientSecret": config.get('clientSecret', ''),
                "applicationCode": config.get('applicationCode', '')
            }

        except Exception as e:
            _logger.error(f"Config yukleme hatasi: {e}")
            self.base_url = ''
            self.endpoint = ''
            self.customer_no = ''
            self.auth_data = {}

    def _get_token(self) -> bool:
        try:
            response = requests.post(
                f"{self.base_url}/Authorization/GetAccessToken",
                json=self.auth_data,
                timeout=10
            )

            if response.status_code == 200:
                data = response.json()
                if data.get('isSuccess') and 'data' in data:
                    self.token = data['data']['accessToken']
                    return True
            return False
        except Exception as e:
            _logger.error(f"Token alma hatasi: {e}")
            return False

    def fetch_data(self, start_date: str, end_date: str) -> list:
        """
        Belirtilen tarih araligindaki TUM siparisleri ceker.
        Iptal edilmis siparisler haric tutulur.
        """
        if not self.token and not self._get_token():
            _logger.error("Token alinamadi")
            return []

        try:
            payload = {
                "orderId": "",
                "CustomerNo": self.customer_no,
                "RegistrationDateStart": start_date,
                "RegistrationDateEnd": end_date,
                "referenceDocumentNo": "",
                "SalesDocumentType": ""
            }

            response = requests.post(
                f"{self.base_url}{self.endpoint}",
                json=payload,
                headers={
                    'Authorization': f'Bearer {self.token}',
                    'Content-Type': 'application/json'
                },
                timeout=30
            )

            if response.status_code != 200:
                _logger.error(f"API yanit kodu: {response.status_code}")
                return []

            result = response.json()
            if not result.get('isSuccess') or not isinstance(result.get('data'), list):
                _logger.error("API basarisiz yanit")
                return []

            data = result['data']

            filtered = [
                record for record in data
                if 'iptal' not in str(record.get('orderStatus', '')).lower()
                and str(record.get('odemeKosulu', '')).strip() != 'Z347'
            ]

            seen = set()
            unique = []
            for record in filtered:
                key = f"{record.get('orderId', '')}-{record.get('orderLineId', '')}"
                if key not in seen:
                    seen.add(key)
                    unique.append(record)

            return unique

        except Exception as e:
            _logger.error(f"API cagrisi hatasi: {e}")
            return []


# ============================================================================
# PRIM CALCULATOR
# ============================================================================

class PrimCalculator:
    """Dogtas ticari politikasina gore prim hesaplama motoru."""

    @staticmethod
    def calculate_monthly_premium(
        realized_order: Decimal,
        target: Decimal,
        realized_invoice: Decimal
    ) -> dict:
        if target <= 0:
            return {
                'hgo': Decimal('0'),
                'rate': Decimal('0'),
                'premium_amount': Decimal('0')
            }

        hgo = (realized_order / target) * 100

        if hgo >= 120:
            rate = Decimal('3.0')
        elif hgo >= 110:
            rate = Decimal('2.5')
        elif hgo >= 100:
            rate = Decimal('2.0')
        elif hgo >= 90:
            rate = Decimal('1')
        else:
            rate = Decimal('0')

        premium_amount = (rate / 100) * realized_invoice

        return {
            'hgo': hgo,
            'rate': rate,
            'premium_amount': premium_amount
        }

    @staticmethod
    def calculate_quarterly_extra_premium(
        total_order: Decimal,
        total_target: Decimal,
        total_invoice: Decimal,
        ek_prim_tiers: list = None
    ) -> dict:
        if total_target <= 0:
            return {
                'eligible': False,
                'hgo': Decimal('0'),
                'rate': Decimal('0'),
                'premium_amount': Decimal('0'),
                'reason': 'Hedef tan\u0131mlanmam\u0131\u015f'
            }

        hgo = (total_order / total_target) * 100

        if hgo < 100:
            return {
                'eligible': False,
                'hgo': hgo,
                'rate': Decimal('0'),
                'premium_amount': Decimal('0'),
                'reason': f'\u00c7eyrek HGO %100 alt\u0131 (%{hgo:.2f})'
            }

        if ek_prim_tiers is None:
            ek_prim_tiers = [
                {'alt_sinir': Decimal('50000000'), 'oran': Decimal('5')},
                {'alt_sinir': Decimal('35000000'), 'oran': Decimal('4')},
                {'alt_sinir': Decimal('20000000'), 'oran': Decimal('3')},
                {'alt_sinir': Decimal('10000000'), 'oran': Decimal('2')},
                {'alt_sinir': Decimal('7000000'), 'oran': Decimal('1')},
            ]

        sorted_tiers = sorted(ek_prim_tiers, key=lambda t: t['alt_sinir'], reverse=True)

        rate = Decimal('0')
        for tier in sorted_tiers:
            if total_order >= tier['alt_sinir']:
                rate = tier['oran']
                break

        if rate == 0:
            return {
                'eligible': False,
                'hgo': hgo,
                'rate': Decimal('0'),
                'premium_amount': Decimal('0'),
                'reason': f'Ciro baraj\u0131 kar\u015f\u0131lanmad\u0131 ({_format_currency(total_order)})'
            }

        premium_amount = (rate / 100) * total_invoice

        return {
            'eligible': True,
            'hgo': hgo,
            'rate': rate,
            'premium_amount': premium_amount,
            'reason': ''
        }

    @staticmethod
    def generate_forecast(monthly_data: dict, months: list, ek_prim_tiers: list = None) -> list:
        lines = []
        current_month = datetime.now().month

        active_months = [m for m in months if monthly_data[m]['realized_order'] > 0]
        remaining_months = [m for m in months if monthly_data[m]['realized_order'] == 0]

        if not active_months:
            lines.append("Hen\u00fcz veri yok. Hesapla butonuna bas\u0131n.")
            return lines

        total_order = sum(monthly_data[m]['realized_order'] for m in months)
        total_invoice = sum(monthly_data[m]['realized_invoice'] for m in months)
        total_target = sum(monthly_data[m]['target'] for m in months)

        if total_target <= 0:
            return ["Hedef tan\u0131mlanmam\u0131\u015f."]

        current_hgo = (total_order / total_target) * 100

        if remaining_months:
            avg_monthly_order = total_order / len(active_months)
            avg_monthly_invoice = total_invoice / len(active_months)
            projected_order = total_order + avg_monthly_order * len(remaining_months)
            projected_invoice = total_invoice + avg_monthly_invoice * len(remaining_months)
            projected_hgo = (projected_order / total_target) * 100

            lines.append(f"-- \u00c7EYREK SONU TAHM\u0130N\u0130 ({len(active_months)} ay verisi ile) --")
            lines.append(f"Ayl\u0131k ortalama sipari\u015f: {_format_currency(avg_monthly_order)}")
            lines.append(f"Tahmini \u00e7eyrek sonu sipari\u015f: {_format_currency(projected_order)}")
            lines.append(f"Tahmini \u00e7eyrek sonu HGO: %{projected_hgo:.1f}")
            lines.append("")
        else:
            projected_order = total_order
            projected_invoice = total_invoice
            projected_hgo = current_hgo
            lines.append("-- \u00c7EYREK TAMAMLANDI --")
            lines.append(f"Toplam sipari\u015f: {_format_currency(total_order)}")
            lines.append(f"\u00c7eyrek HGO: %{current_hgo:.1f}")
            lines.append("")

        hgo_tiers = [
            (Decimal('120'), Decimal('3'), 'Alt\u0131n'),
            (Decimal('110'), Decimal('2.5'), 'G\u00fcm\u00fc\u015f'),
            (Decimal('100'), Decimal('2'), 'Bronz'),
            (Decimal('90'), Decimal('1'), 'Ba\u015flang\u0131\u00e7'),
        ]

        if current_month in months and current_month in monthly_data:
            m_data = monthly_data[current_month]
            m_target = m_data['target']
            m_order = m_data['realized_order']
            m_name = _get_turkish_month_name(current_month)

            if m_target > 0:
                m_hgo = (m_order / m_target) * 100
                lines.append(f"-- {m_name.upper()} AYI BIREYSEL HEDEF DURUMU --")
                lines.append(f"Hedef: {_format_currency(m_target)} | Sipari\u015f: {_format_currency(m_order)} | HGO: %{m_hgo:.1f}")

                for threshold, rate, name in reversed(hgo_tiers):
                    needed = (threshold / 100) * m_target
                    if m_order >= needed:
                        lines.append(f"  %{threshold} ({name} - %{rate} prim): ULA\u015eILDI")
                    else:
                        gap = needed - m_order
                        lines.append(f"  %{threshold} ({name} - %{rate} prim): {_format_currency(gap)} daha sipari\u015f gerekli")

                lines.append("")

        third_month = months[-1]
        if current_month == third_month and ek_prim_tiers:
            lines.append("-- EK PR\u0130M DURUMU --")

            reached = None
            next_tier = None
            for tier in sorted(ek_prim_tiers, key=lambda t: t['alt_sinir'], reverse=True):
                if projected_order >= tier['alt_sinir']:
                    if reached is None:
                        reached = tier
                else:
                    next_tier = tier

            if reached:
                lines.append(
                    f"  Mevcut dilim: {_format_currency(reached['alt_sinir'])} (%{reached['oran']} ek prim) - ULA\u015eILDI"
                )
            else:
                lines.append("  Hen\u00fcz hi\u00e7bir dilime ula\u015f\u0131lamad\u0131.")

            if next_tier:
                gap = next_tier['alt_sinir'] - projected_order
                lines.append(
                    f"  Sonraki dilim: {_format_currency(next_tier['alt_sinir'])} (%{next_tier['oran']} ek prim)"
                    f" - {_format_currency(gap)} daha sipari\u015f gerekli"
                )

            lines.append("")

        lines.append("-- STRATEJI \u00d6NER\u0130S\u0130 --")

        needed_for_100 = total_target
        if projected_order >= needed_for_100:
            lines.append(f"  \u00c7eyrek hedefi (%100): ULA\u015eILDI")
        else:
            gap = needed_for_100 - total_order
            if remaining_months:
                monthly_needed = gap / len(remaining_months)
                lines.append(
                    f"  \u00c7eyrek hedefi (%100 - Bronz): {_format_currency(gap)} daha sipari\u015f gerekli"
                )
                lines.append(
                    f"  Kalan {len(remaining_months)} ayda aylik {_format_currency(monthly_needed)} sipari\u015f ile ula\u015f\u0131labilir."
                )
            else:
                lines.append(
                    f"  \u00c7eyrek hedefi (%100 - Bronz): {_format_currency(gap)} eksik"
                )

        if projected_hgo >= 90:
            proj_result = PrimCalculator.calculate_monthly_premium(
                projected_order, total_target, projected_invoice
            )
            lines.append(f"  Tahmini toplam ayl\u0131k prim: {_format_currency(proj_result['premium_amount'])}")

        return lines


# ============================================================================
# STORAGE MANAGER
# ============================================================================

class _StorageManager:
    """Google Sheets uzerinde Hedef verilerini yoneten sinif."""

    WORKSHEET_NAME = 'Hedef'
    HEADERS = ['Y\u0131l', '\u00c7eyrek', 'Ay', 'Hedef Tutar']

    def __init__(self, config_manager: CentralConfigManager):
        self.config_manager = config_manager

    def _get_or_create_worksheet(self):
        spreadsheet = self.config_manager.gc.open_by_key(
            self.config_manager.MASTER_SPREADSHEET_ID
        )
        try:
            worksheet = spreadsheet.worksheet(self.WORKSHEET_NAME)
        except Exception:
            worksheet = spreadsheet.add_worksheet(
                title=self.WORKSHEET_NAME,
                rows=100,
                cols=len(self.HEADERS)
            )
            worksheet.update(values=[self.HEADERS], range_name='A1')
        return worksheet

    def load_targets(self, year: int, quarter: int) -> list:
        try:
            worksheet = self._get_or_create_worksheet()
            all_values = worksheet.get_all_values()

            if len(all_values) <= 1:
                return []

            headers = all_values[0]
            try:
                yil_idx = headers.index('Y\u0131l')
                ceyrek_idx = headers.index('\u00c7eyrek')
                ay_idx = headers.index('Ay')
                tutar_idx = headers.index('Hedef Tutar')
            except ValueError:
                return []

            quarter_str = f"Q{quarter}"
            targets = []

            for row in all_values[1:]:
                if len(row) > max(yil_idx, ceyrek_idx, ay_idx, tutar_idx):
                    if row[yil_idx].strip() == str(year) and row[ceyrek_idx].strip() == quarter_str:
                        try:
                            targets.append({
                                'yil': year,
                                'ceyrek': quarter_str,
                                'ay': int(row[ay_idx].strip()),
                                'hedef_tutar': Decimal(row[tutar_idx].strip())
                            })
                        except (ValueError, InvalidOperation):
                            continue

            return targets

        except Exception as e:
            _logger.error(f"Hedef yukleme hatasi: {e}")
            return []

    def save_targets(self, year: int, quarter: int, targets: list):
        try:
            worksheet = self._get_or_create_worksheet()
            all_values = worksheet.get_all_values()

            quarter_str = f"Q{quarter}"

            kept_rows = []
            if len(all_values) > 1:
                headers = all_values[0]
                try:
                    yil_idx = headers.index('Y\u0131l')
                    ceyrek_idx = headers.index('\u00c7eyrek')
                except ValueError:
                    yil_idx, ceyrek_idx = 0, 1

                for row in all_values[1:]:
                    if len(row) > max(yil_idx, ceyrek_idx):
                        if not (row[yil_idx].strip() == str(year) and row[ceyrek_idx].strip() == quarter_str):
                            kept_rows.append(row)

            for t in targets:
                kept_rows.append([
                    str(year),
                    quarter_str,
                    str(t['ay']),
                    str(t['hedef_tutar'])
                ])

            worksheet.clear()
            all_data = [self.HEADERS] + kept_rows
            worksheet.update(values=all_data, range_name='A1')

        except Exception as e:
            _logger.error(f"Hedef kaydetme hatasi: {e}")

    # --- EkPrim Worksheet ---

    EK_PRIM_WORKSHEET = 'EkPrim'
    EK_PRIM_HEADERS = ['Y\u0131l', '\u00c7eyrek', 'AltS\u0131n\u0131r', 'PrimOran\u0131']

    EK_PRIM_DEFAULTS = [
        {'alt_sinir': Decimal('50000000'), 'oran': Decimal('5')},
        {'alt_sinir': Decimal('35000000'), 'oran': Decimal('4')},
        {'alt_sinir': Decimal('20000000'), 'oran': Decimal('3')},
        {'alt_sinir': Decimal('10000000'), 'oran': Decimal('2')},
        {'alt_sinir': Decimal('7000000'), 'oran': Decimal('1')},
    ]

    def _get_or_create_ek_prim_worksheet(self):
        spreadsheet = self.config_manager.gc.open_by_key(
            self.config_manager.MASTER_SPREADSHEET_ID
        )
        try:
            worksheet = spreadsheet.worksheet(self.EK_PRIM_WORKSHEET)
        except Exception:
            worksheet = spreadsheet.add_worksheet(
                title=self.EK_PRIM_WORKSHEET,
                rows=100,
                cols=len(self.EK_PRIM_HEADERS)
            )
            worksheet.update(values=[self.EK_PRIM_HEADERS], range_name='A1')
        return worksheet

    def load_ek_prim_tiers(self, year: int, quarter: int) -> list:
        try:
            worksheet = self._get_or_create_ek_prim_worksheet()
            all_values = worksheet.get_all_values()

            if len(all_values) <= 1:
                return []

            headers = all_values[0]
            try:
                yil_idx = headers.index('Y\u0131l')
                ceyrek_idx = headers.index('\u00c7eyrek')
                sinir_idx = headers.index('AltS\u0131n\u0131r')
                oran_idx = headers.index('PrimOran\u0131')
            except ValueError:
                return []

            quarter_str = f"Q{quarter}"
            tiers = []

            for row in all_values[1:]:
                if len(row) > max(yil_idx, ceyrek_idx, sinir_idx, oran_idx):
                    if row[yil_idx].strip() == str(year) and row[ceyrek_idx].strip() == quarter_str:
                        try:
                            tiers.append({
                                'alt_sinir': Decimal(row[sinir_idx].strip()),
                                'oran': Decimal(row[oran_idx].strip())
                            })
                        except (ValueError, InvalidOperation):
                            continue

            return sorted(tiers, key=lambda t: t['alt_sinir'], reverse=True)

        except Exception as e:
            _logger.error(f"EkPrim y\u00fckleme hatas\u0131: {e}")
            return []

    def save_ek_prim_tiers(self, year: int, quarter: int, tiers: list):
        try:
            worksheet = self._get_or_create_ek_prim_worksheet()
            all_values = worksheet.get_all_values()

            quarter_str = f"Q{quarter}"

            kept_rows = []
            if len(all_values) > 1:
                headers = all_values[0]
                try:
                    yil_idx = headers.index('Y\u0131l')
                    ceyrek_idx = headers.index('\u00c7eyrek')
                except ValueError:
                    yil_idx, ceyrek_idx = 0, 1

                for row in all_values[1:]:
                    if len(row) > max(yil_idx, ceyrek_idx):
                        if not (row[yil_idx].strip() == str(year) and row[ceyrek_idx].strip() == quarter_str):
                            kept_rows.append(row)

            for t in tiers:
                kept_rows.append([
                    str(year),
                    quarter_str,
                    str(t['alt_sinir']),
                    str(t['oran'])
                ])

            worksheet.clear()
            all_data = [self.EK_PRIM_HEADERS] + kept_rows
            worksheet.update(values=all_data, range_name='A1')

        except Exception as e:
            _logger.error(f"EkPrim kaydetme hatas\u0131: {e}")


# ============================================================================
# YARDIMCI FONKSIYONLAR
# ============================================================================

def _get_quarter_dates(year: int, quarter: int):
    mapping = {
        1: (date(year, 1, 1), date(year, 3, 31)),
        2: (date(year, 4, 1), date(year, 6, 30)),
        3: (date(year, 7, 1), date(year, 9, 30)),
        4: (date(year, 10, 1), date(year, 12, 31)),
    }
    return mapping.get(quarter, (None, None))


def _get_quarter_months(quarter: int) -> list:
    mapping = {1: [1, 2, 3], 2: [4, 5, 6], 3: [7, 8, 9], 4: [10, 11, 12]}
    return mapping.get(quarter, [])


def _get_turkish_month_name(month_num: int) -> str:
    return TURKISH_MONTHS.get(month_num, str(month_num))


def _format_currency(amount: Decimal) -> str:
    return f"{amount:,.0f} TL"


def _parse_date(date_str: str):
    if not date_str:
        return None
    try:
        if 'T' in date_str:
            return datetime.strptime(date_str.split('T')[0], "%Y-%m-%d")
        elif '-' in date_str:
            return datetime.strptime(date_str, "%Y-%m-%d")
        elif '.' in date_str:
            return datetime.strptime(date_str, "%d.%m.%Y")
    except ValueError:
        pass
    return None


def _parse_invoice_date(date_str: str):
    if not date_str or date_str == '00000000':
        return None
    try:
        if len(date_str) == 8 and date_str.isdigit():
            return datetime.strptime(date_str, "%Y%m%d")
        elif '.' in date_str:
            return datetime.strptime(date_str, "%d.%m.%Y")
        elif '-' in date_str:
            return datetime.strptime(date_str, "%Y-%m-%d")
    except ValueError:
        pass
    return None


def _safe_decimal(value) -> Decimal:
    try:
        if value is None or value == '':
            return Decimal('0')
        s = str(value).strip()
        if ',' in s:
            s = s.replace('.', '').replace(',', '.')
        return Decimal(s)
    except (InvalidOperation, ValueError):
        return Decimal('0')


def _process_raw_data(raw_data: list, start_date: date, end_date: date, months: list, target_map: dict) -> dict:
    monthly_data = {}
    for m in months:
        monthly_data[m] = {
            'target': target_map.get(m, Decimal('0')),
            'realized_order': Decimal('0'),
            'realized_invoice': Decimal('0'),
        }

    for item in raw_data:
        qty = _safe_decimal(item.get('orderLineQuantity', '0'))
        net_price = _safe_decimal(item.get('netPrice', '0'))
        original_price = _safe_decimal(item.get('originalPrice', '0'))

        price = net_price if net_price != 0 else original_price
        line_total = qty * price

        o_date = _parse_date(item.get('orderDate1', ''))
        if o_date and o_date.month in monthly_data:
            if start_date <= o_date.date() <= end_date:
                monthly_data[o_date.month]['realized_order'] += line_total

        inv_date_str = item.get('purchaseInvoiceDate', '')
        if inv_date_str and inv_date_str != '00000000':
            inv_date = _parse_invoice_date(inv_date_str)
            if inv_date and inv_date.month in monthly_data:
                if start_date <= inv_date.date() <= end_date:
                    monthly_data[inv_date.month]['realized_invoice'] += line_total

    return monthly_data


# ============================================================================
# DATA FETCH WORKER (QThread)
# ============================================================================

class _DataFetchWorker(QThread):
    progress = pyqtSignal(str)
    finished_ok = pyqtSignal(list)
    finished_err = pyqtSignal(str)

    def __init__(self, api_client: PrimApiClient, start_date: str, end_date: str):
        super().__init__()
        self.api_client = api_client
        self.start_date = start_date
        self.end_date = end_date

    def run(self):
        try:
            self.progress.emit("API'ye ba\u011flan\u0131l\u0131yor...")
            data = self.api_client.fetch_data(self.start_date, self.end_date)
            if data:
                self.progress.emit(f"{len(data)} kay\u0131t \u00e7ekildi. Hesaplan\u0131yor...")
                self.finished_ok.emit(data)
            else:
                self.finished_err.emit("API'den veri \u00e7ekilemedi veya kay\u0131t bulunamad\u0131.")
        except Exception as e:
            self.finished_err.emit(f"Hata: {e}")


# ============================================================================
# STYLESHEET
# ============================================================================

_STYLESHEET = """
QWidget#hgo_root {
    font-family: 'Segoe UI', Arial;
    font-size: 12px;
    background-color: #e8e8e8;
}
QWidget#hgo_root QWidget {
    background-color: #e8e8e8;
    color: #212529;
}
QWidget#hgo_root QGroupBox {
    font-weight: bold;
    font-size: 12px;
    border: 1px solid #ced4da;
    border-radius: 4px;
    margin-top: 6px;
    padding-top: 10px;
    background-color: #ffffff;
    color: #495057;
}
QWidget#hgo_root QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 8px;
    padding: 0 4px;
    color: #495057;
    background-color: #ffffff;
}
QWidget#hgo_root QGroupBox QWidget {
    background-color: #ffffff;
}
QWidget#hgo_root QLineEdit {
    padding: 6px 10px;
    border: 1px solid #ced4da;
    border-radius: 4px;
    background-color: #ffffff;
    font-size: 13px;
    color: #212529;
}
QWidget#hgo_root QLineEdit:focus {
    border-color: #007acc;
}
QWidget#hgo_root QComboBox {
    padding: 6px 10px;
    border: 1px solid #ced4da;
    border-radius: 4px;
    background-color: #ffffff;
    font-size: 13px;
    min-width: 80px;
    color: #212529;
}
QWidget#hgo_root QComboBox QAbstractItemView {
    background-color: #ffffff;
    color: #212529;
    selection-background-color: #007acc;
    selection-color: #ffffff;
}
QWidget#hgo_root QPushButton {
    padding: 8px 20px;
    border: none;
    border-radius: 4px;
    font-size: 13px;
    font-weight: bold;
    color: #ffffff;
    background-color: #007acc;
}
QWidget#hgo_root QPushButton:hover {
    background-color: #005fa3;
}
QWidget#hgo_root QPushButton:disabled {
    background-color: #adb5bd;
}
QWidget#hgo_root QPushButton#btn_save {
    background-color: #28a745;
}
QWidget#hgo_root QPushButton#btn_save:hover {
    background-color: #218838;
}
QWidget#hgo_root QTableWidget {
    border: 1px solid #dee2e6;
    border-radius: 4px;
    background-color: #ffffff;
    gridline-color: #e9ecef;
    font-size: 12px;
    color: #212529;
}
QWidget#hgo_root QTableWidget::item {
    padding: 6px;
    background-color: #ffffff;
    color: #212529;
}
QWidget#hgo_root QTableWidget::item:selected {
    background-color: #e8f4fd;
    color: #212529;
    border: none;
}
QWidget#hgo_root QTableWidget:focus {
    outline: none;
    border: 1px solid #dee2e6;
}
QWidget#hgo_root QTableWidget {
    selection-background-color: #e8f4fd;
    selection-color: #212529;
}
QWidget#hgo_root QHeaderView::section {
    background-color: #343a40;
    color: #ffffff;
    padding: 8px;
    border: none;
    font-weight: bold;
    font-size: 12px;
}
QWidget#hgo_root QProgressBar {
    border: 1px solid #dee2e6;
    border-radius: 4px;
    text-align: center;
    background-color: #e9ecef;
    height: 22px;
}
QWidget#hgo_root QProgressBar::chunk {
    background-color: #007acc;
    border-radius: 3px;
}
QWidget#hgo_root QLabel {
    color: #212529;
    background: transparent;
}
QWidget#hgo_root QTextEdit {
    background-color: #f8f9fa;
    color: #212529;
}
QWidget#hgo_root QTabWidget::pane {
    border: 1px solid #dee2e6;
    border-radius: 4px;
    background: #ffffff;
}
QWidget#hgo_root QTabBar::tab {
    padding: 6px 30px;
    font-size: 13px;
    font-weight: bold;
    min-width: 140px;
}
QWidget#hgo_root QTabBar::tab:selected {
    background: #007acc;
    color: #ffffff;
    border-radius: 4px 4px 0 0;
}
QWidget#hgo_root QTabBar::tab:!selected {
    background: #e9ecef;
    color: #495057;
}
"""


# ============================================================================
# ANA MODUL SINIFI
# ============================================================================

class HgoModule(QWidget):
    """Dogtas HGO Prim Hesaplama Modulu"""

    def __init__(self):
        super().__init__()
        self.setObjectName("hgo_root")
        self.setStyleSheet(_STYLESHEET)

        # Backend
        self.config_manager = CentralConfigManager()
        self.storage = _StorageManager(self.config_manager)
        self.api_client = PrimApiClient(self.config_manager)
        self.worker = None

        # State
        self.target_inputs = {}
        self.ek_prim_tiers = None

        self._setup_ui()
        self._load_targets()
        self._load_ek_prim_tiers()

    # ---------------------------------------------------------------
    # UI SETUP
    # ---------------------------------------------------------------

    def _setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(4)
        main_layout.setContentsMargins(8, 4, 8, 4)

        # -- Tab Widget --
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget, 1)

        # === TAB 1: PRIM HESAPLAMA ===
        tab1 = QWidget()
        tab1_layout = QVBoxLayout(tab1)
        tab1_layout.setSpacing(4)
        tab1_layout.setContentsMargins(6, 6, 6, 6)

        # -- Donem + Hedef (tek satir) --
        top_row = QHBoxLayout()
        top_row.setSpacing(6)

        period_group = QGroupBox("D\u00f6nem")
        period_group.setFixedHeight(60)
        period_layout = QHBoxLayout(period_group)
        period_layout.setContentsMargins(8, 2, 8, 2)

        period_layout.addWidget(QLabel("Y\u0131l:"))
        self.year_combo = QComboBox()
        current_year = datetime.now().year
        for y in range(current_year - 2, current_year + 3):
            self.year_combo.addItem(str(y), y)
        self.year_combo.setCurrentText(str(current_year))
        self.year_combo.currentIndexChanged.connect(self._on_period_changed)
        period_layout.addWidget(self.year_combo)

        period_layout.addSpacing(10)
        period_layout.addWidget(QLabel("\u00c7eyrek:"))
        self.quarter_combo = QComboBox()
        for q in range(1, 5):
            self.quarter_combo.addItem(f"Q{q}", q)
        current_quarter = (datetime.now().month - 1) // 3 + 1
        self.quarter_combo.setCurrentIndex(current_quarter - 1)
        self.quarter_combo.currentIndexChanged.connect(self._on_period_changed)
        period_layout.addWidget(self.quarter_combo)

        top_row.addWidget(period_group)

        self.btn_calculate = QPushButton("HESAPLA")
        self.btn_calculate.setMinimumWidth(100)
        self.btn_calculate.setFixedHeight(36)
        self.btn_calculate.clicked.connect(self._on_calculate)
        top_row.addWidget(self.btn_calculate)

        # -- Hedefler (ayni satir) --
        self.target_group = QGroupBox("Ayl\u0131k Hedefler")
        self.target_group.setFixedHeight(60)
        self.target_layout = QGridLayout(self.target_group)
        self.target_layout.setSpacing(4)
        self.target_layout.setContentsMargins(8, 2, 8, 2)

        self.btn_save_targets = QPushButton("Kaydet")
        self.btn_save_targets.setObjectName("btn_save")
        self.btn_save_targets.clicked.connect(self._save_targets)

        self._rebuild_target_inputs()
        top_row.addWidget(self.target_group, 1)

        tab1_layout.addLayout(top_row)

        # -- Sonuc Tablosu --
        table_group = QGroupBox("Ayl\u0131k Performans Raporu")
        table_layout = QVBoxLayout(table_group)
        table_layout.setContentsMargins(6, 12, 6, 4)

        self.result_table = QTableWidget(3, 7)
        self.result_table.setHorizontalHeaderLabels(
            ["Ay", "Hedef", "Sipari\u015f", "HGO %", "Fatura", "Prim Oran\u0131", "Prim Tutar\u0131"]
        )
        self.result_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.result_table.verticalHeader().setVisible(False)
        self.result_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.result_table.setSelectionMode(QTableWidget.SingleSelection)
        self.result_table.setAlternatingRowColors(True)
        self.result_table.setStyleSheet(
            "QTableWidget { alternate-background-color: #f5f5f5; }"
        )
        self.result_table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.result_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.result_table.setFixedHeight(160)
        table_layout.addWidget(self.result_table)

        tab1_layout.addWidget(table_group)

        # -- Ceyrek Ozeti + Toplam Prim (yatay) --
        summary_row = QHBoxLayout()
        summary_row.setSpacing(6)

        summary_group = QGroupBox("\u00c7eyrek Sonu \u00d6zeti")
        summary_layout = QGridLayout(summary_group)
        summary_layout.setSpacing(4)
        summary_layout.setContentsMargins(8, 12, 8, 4)

        self.lbl_total_target = QLabel("-")
        self.lbl_total_order = QLabel("-")
        self.lbl_total_invoice = QLabel("-")
        self.lbl_quarter_hgo = QLabel("-")
        self.lbl_extra_premium = QLabel("-")

        labels_data = [
            ("Toplam Hedef:", self.lbl_total_target, 0, 0),
            ("Toplam Sipari\u015f:", self.lbl_total_order, 0, 2),
            ("Toplam Fatura:", self.lbl_total_invoice, 0, 4),
            ("\u00c7eyrek HGO:", self.lbl_quarter_hgo, 1, 0),
            ("Ek Hacim Primi:", self.lbl_extra_premium, 1, 2),
        ]
        for text, widget, row, col in labels_data:
            lbl = QLabel(text)
            lbl.setFont(QFont("Segoe UI", 10, QFont.Bold))
            lbl.setStyleSheet("background: transparent;")
            summary_layout.addWidget(lbl, row, col)
            widget.setFont(QFont("Segoe UI", 10))
            widget.setStyleSheet("background: transparent;")
            summary_layout.addWidget(widget, row, col + 1)

        summary_row.addWidget(summary_group, 2)

        # -- Toplam Prim (sag taraf) --
        self.lbl_total_premium = QLabel("TOPLAM\n-")
        self.lbl_total_premium.setAlignment(Qt.AlignCenter)
        self.lbl_total_premium.setFont(QFont("Segoe UI", 14, QFont.Bold))
        self.lbl_total_premium.setStyleSheet(
            "color: #28a745; padding: 8px; background-color: #ffffff; "
            "border: 2px solid #28a745; border-radius: 8px;"
        )
        self.lbl_total_premium.setMinimumWidth(280)
        summary_row.addWidget(self.lbl_total_premium)

        tab1_layout.addLayout(summary_row)

        # -- Tahmin ve Oneri --
        forecast_group = QGroupBox("Tahmin ve \u00d6neri")
        forecast_layout = QHBoxLayout(forecast_group)
        forecast_layout.setSpacing(6)
        forecast_layout.setContentsMargins(6, 12, 6, 4)

        forecast_text_style = (
            "QTextEdit { background-color: #f8f9fa; border: 1px solid #dee2e6; "
            "border-radius: 4px; font-family: 'Consolas', 'Courier New', monospace; "
            "font-size: 14px; padding: 6px; color: #212529; }"
        )

        self.forecast_left = QTextEdit()
        self.forecast_left.setReadOnly(True)
        self.forecast_left.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.forecast_left.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.forecast_left.setStyleSheet(forecast_text_style)
        self.forecast_left.setPlaceholderText("\u00c7eyrek sonu tahmini...")
        forecast_layout.addWidget(self.forecast_left)

        self.forecast_right = QTextEdit()
        self.forecast_right.setReadOnly(True)
        self.forecast_right.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.forecast_right.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.forecast_right.setStyleSheet(forecast_text_style)
        self.forecast_right.setPlaceholderText("Ayl\u0131k bireysel hedef durumu...")
        forecast_layout.addWidget(self.forecast_right)

        tab1_layout.addWidget(forecast_group, 1)

        self.tab_widget.addTab(tab1, "Prim Hesaplama")

        # === TAB 2: EK PRIM AYARLARI ===
        tab2 = QWidget()
        tab2_layout = QVBoxLayout(tab2)
        tab2_layout.setSpacing(6)
        tab2_layout.setContentsMargins(6, 6, 6, 6)

        ep_period_group = QGroupBox("D\u00f6nem Se\u00e7imi")
        ep_period_group.setFixedHeight(60)
        ep_period_layout = QHBoxLayout(ep_period_group)
        ep_period_layout.setContentsMargins(8, 2, 8, 2)

        ep_period_layout.addWidget(QLabel("Y\u0131l:"))
        self.ep_year_combo = QComboBox()
        for y in range(current_year - 2, current_year + 3):
            self.ep_year_combo.addItem(str(y), y)
        self.ep_year_combo.setCurrentText(str(current_year))
        ep_period_layout.addWidget(self.ep_year_combo)

        ep_period_layout.addSpacing(20)
        ep_period_layout.addWidget(QLabel("\u00c7eyrek:"))
        self.ep_quarter_combo = QComboBox()
        for q in range(1, 5):
            self.ep_quarter_combo.addItem(f"Q{q}", q)
        self.ep_quarter_combo.setCurrentIndex(current_quarter - 1)
        ep_period_layout.addWidget(self.ep_quarter_combo)

        ep_period_layout.addSpacing(20)
        btn_ep_load = QPushButton("Y\u00fckle")
        btn_ep_load.clicked.connect(self._load_ek_prim_tiers)
        ep_period_layout.addWidget(btn_ep_load)

        ep_period_layout.addStretch()
        tab2_layout.addWidget(ep_period_group)

        # -- Dilim tablosu --
        tiers_group = QGroupBox("Toplam Sipari\u015f Hacmi Prim Dilimleri")
        tiers_layout = QVBoxLayout(tiers_group)

        self.ep_table = QTableWidget(0, 2)
        self.ep_table.setHorizontalHeaderLabels(["Alt S\u0131n\u0131r (TL)", "Prim Oran\u0131 (%)"])
        self.ep_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.ep_table.verticalHeader().setVisible(False)
        self.ep_table.setStyleSheet(
            "QTableWidget { alternate-background-color: #f5f5f5; }"
        )
        self.ep_table.setAlternatingRowColors(True)
        tiers_layout.addWidget(self.ep_table)

        ep_btn_layout = QHBoxLayout()
        btn_add_row = QPushButton("Sat\u0131r Ekle")
        btn_add_row.setMinimumWidth(100)
        btn_add_row.clicked.connect(self._ep_add_row)
        ep_btn_layout.addWidget(btn_add_row)

        btn_del_row = QPushButton("Sat\u0131r Sil")
        btn_del_row.setMinimumWidth(100)
        btn_del_row.setStyleSheet(
            "QPushButton { background-color: #dc3545; }"
            "QPushButton:hover { background-color: #c82333; }"
        )
        btn_del_row.clicked.connect(self._ep_del_row)
        ep_btn_layout.addWidget(btn_del_row)

        btn_defaults = QPushButton("Varsay\u0131lanlar\u0131 Y\u00fckle")
        btn_defaults.setMinimumWidth(140)
        btn_defaults.setStyleSheet(
            "QPushButton { background-color: #6c757d; }"
            "QPushButton:hover { background-color: #5a6268; }"
        )
        btn_defaults.clicked.connect(self._ep_load_defaults)
        ep_btn_layout.addWidget(btn_defaults)

        ep_btn_layout.addStretch()

        btn_ep_save = QPushButton("Kaydet")
        btn_ep_save.setObjectName("btn_save")
        btn_ep_save.setMinimumWidth(120)
        btn_ep_save.clicked.connect(self._save_ek_prim_tiers)
        ep_btn_layout.addWidget(btn_ep_save)

        tiers_layout.addLayout(ep_btn_layout)
        tab2_layout.addWidget(tiers_group, 1)

        ep_info = QLabel(
            "Not: Ek hacim primi, ceyrekteki 3 aylik toplam siparis hedefini %100 ve uzeri "
            "gerceklestiren bayiler icin gecerlidir. Prim orani toplam siparis tutarina gore "
            "belirlenir, prim tutari ise toplam faturalanan net tutar uzerinden hesaplanir."
        )
        ep_info.setWordWrap(True)
        ep_info.setStyleSheet(
            "color: #6c757d; padding: 8px; background: #f8f9fa; "
            "border: 1px solid #dee2e6; border-radius: 4px;"
        )
        tab2_layout.addWidget(ep_info)

        self.tab_widget.addTab(tab2, "EkPrim Ayarlari")

        # -- Progress + Status --
        status_layout = QHBoxLayout()
        status_layout.setContentsMargins(0, 0, 0, 0)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setVisible(False)
        self.progress_bar.setMaximumHeight(16)
        status_layout.addWidget(self.progress_bar)

        self.status_label = QLabel("Haz\u0131r")
        self.status_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.status_label.setStyleSheet("color: #6c757d; background: transparent; font-size: 11px;")
        status_layout.addWidget(self.status_label)
        main_layout.addLayout(status_layout)

    # ---------------------------------------------------------------
    # DONEM DEGISIMI
    # ---------------------------------------------------------------

    def _rebuild_target_inputs(self):
        self.target_layout.removeWidget(self.btn_save_targets)

        while self.target_layout.count():
            child = self.target_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

        self.target_inputs = {}
        self.target_labels = {}

        quarter = self.quarter_combo.currentData()
        if quarter is None:
            return

        months = _get_quarter_months(quarter)
        col = 0
        for m in months:
            lbl = QLabel(f"{_get_turkish_month_name(m)}:")
            self.target_layout.addWidget(lbl, 0, col)
            self.target_labels[m] = lbl
            col += 1

            inp = QLineEdit()
            inp.setPlaceholderText("0")
            inp.setAlignment(Qt.AlignRight)
            inp.setMinimumWidth(100)
            self.target_layout.addWidget(inp, 0, col)
            self.target_inputs[m] = inp
            col += 1

        self.target_layout.addWidget(self.btn_save_targets, 0, col)

    def _on_period_changed(self):
        self._rebuild_target_inputs()
        self._load_targets()

    # ---------------------------------------------------------------
    # EK PRIM AYARLARI (Tab 2)
    # ---------------------------------------------------------------

    def _load_ek_prim_tiers(self):
        year = self.ep_year_combo.currentData()
        quarter = self.ep_quarter_combo.currentData()
        if year is None or quarter is None:
            return

        tiers = self.storage.load_ek_prim_tiers(year, quarter)

        if not tiers:
            self._ep_load_defaults()
            self.status_label.setText(f"EkPrim: {year} Q{quarter} icin kay\u0131t yok, varsay\u0131lanlar y\u00fcklendi")
            return

        self.ek_prim_tiers = tiers
        self._ep_populate_table(tiers)
        self.status_label.setText(f"EkPrim: {year} Q{quarter} dilimleri y\u00fcklendi")

    def _ep_populate_table(self, tiers: list):
        self.ep_table.setRowCount(len(tiers))
        for row, tier in enumerate(tiers):
            sinir_item = QTableWidgetItem(f"{tier['alt_sinir']:,.0f}")
            sinir_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.ep_table.setItem(row, 0, sinir_item)

            oran_item = QTableWidgetItem(f"{tier['oran']}")
            oran_item.setTextAlignment(Qt.AlignCenter)
            self.ep_table.setItem(row, 1, oran_item)

    def _ep_add_row(self):
        row = self.ep_table.rowCount()
        self.ep_table.insertRow(row)
        sinir_item = QTableWidgetItem("0")
        sinir_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.ep_table.setItem(row, 0, sinir_item)
        oran_item = QTableWidgetItem("0")
        oran_item.setTextAlignment(Qt.AlignCenter)
        self.ep_table.setItem(row, 1, oran_item)

    def _ep_del_row(self):
        row = self.ep_table.currentRow()
        if row >= 0:
            self.ep_table.removeRow(row)

    def _ep_load_defaults(self):
        defaults = _StorageManager.EK_PRIM_DEFAULTS
        self._ep_populate_table(defaults)

    def _save_ek_prim_tiers(self):
        year = self.ep_year_combo.currentData()
        quarter = self.ep_quarter_combo.currentData()

        tiers = []
        for row in range(self.ep_table.rowCount()):
            sinir_text = self.ep_table.item(row, 0)
            oran_text = self.ep_table.item(row, 1)
            if not sinir_text or not oran_text:
                continue
            try:
                sinir_val = sinir_text.text().strip().replace(',', '').replace(' ', '')
                oran_val = oran_text.text().strip().replace(',', '.')
                alt_sinir = Decimal(sinir_val)
                oran = Decimal(oran_val)
                if alt_sinir < 0 or oran < 0:
                    QMessageBox.warning(self, "Uyar\u0131", "Negatif de\u011fer girilemez.")
                    return
                tiers.append({'alt_sinir': alt_sinir, 'oran': oran})
            except (InvalidOperation, ValueError):
                QMessageBox.warning(self, "Uyar\u0131", f"Sat\u0131r {row + 1}: Ge\u00e7erli say\u0131 giriniz.")
                return

        if not tiers:
            QMessageBox.warning(self, "Uyar\u0131", "En az bir dilim tan\u0131mlanmal\u0131d\u0131r.")
            return

        try:
            self.storage.save_ek_prim_tiers(year, quarter, tiers)
            self.ek_prim_tiers = sorted(tiers, key=lambda t: t['alt_sinir'], reverse=True)
            self.status_label.setText(f"EkPrim: {year} Q{quarter} dilimleri kaydedildi")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Kaydetme hatas\u0131: {e}")

    # ---------------------------------------------------------------
    # HEDEF YUKLEME / KAYDETME
    # ---------------------------------------------------------------

    def _load_targets(self):
        year = self.year_combo.currentData()
        quarter = self.quarter_combo.currentData()
        if year is None or quarter is None:
            return

        try:
            targets = self.storage.load_targets(year, quarter)
            months = _get_quarter_months(quarter)
            target_map = {t['ay']: t['hedef_tutar'] for t in targets}

            for m in months:
                if m in self.target_inputs:
                    val = target_map.get(m)
                    if val is not None:
                        self.target_inputs[m].setText(f"{val:,.0f}".replace(",", "."))
                    else:
                        self.target_inputs[m].clear()

            if targets:
                self.status_label.setText("Hedefler y\u00fcklendi")
        except Exception as e:
            self.status_label.setText(f"Hedef y\u00fckleme hatas\u0131: {e}")

    def _save_targets(self):
        year = self.year_combo.currentData()
        quarter = self.quarter_combo.currentData()
        months = _get_quarter_months(quarter)

        targets = []
        for m in months:
            inp = self.target_inputs.get(m)
            if not inp:
                continue
            val = inp.text().strip()
            if not val:
                QMessageBox.warning(self, "Uyari",
                                    f"{_get_turkish_month_name(m)} i\u00e7in hedef giriniz.")
                return
            try:
                val = val.replace('.', '').replace(',', '.')
                amount = Decimal(val)
                if amount < 0:
                    QMessageBox.warning(self, "Uyar\u0131", "Negatif de\u011fer girilemez.")
                    return
                targets.append({'ay': m, 'hedef_tutar': amount})
            except (InvalidOperation, ValueError):
                QMessageBox.warning(self, "Uyari",
                                    f"{_get_turkish_month_name(m)}: Ge\u00e7erli bir say\u0131 giriniz.")
                return

        try:
            self.storage.save_targets(year, quarter, targets)
            self.status_label.setText("Hedefler kaydedildi")
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Kaydetme hatas\u0131: {e}")

    # ---------------------------------------------------------------
    # HESAPLAMA
    # ---------------------------------------------------------------

    def _on_calculate(self):
        year = self.year_combo.currentData()
        quarter = self.quarter_combo.currentData()

        months = _get_quarter_months(quarter)
        for m in months:
            inp = self.target_inputs.get(m)
            if not inp or not inp.text().strip():
                QMessageBox.warning(self, "Uyar\u0131",
                                    f"{_get_turkish_month_name(m)} i\u00e7in hedef giriniz.")
                return

        start_date, end_date = _get_quarter_dates(year, quarter)
        extended_start = start_date - relativedelta(years=1)
        api_start = extended_start.strftime("%d.%m.%Y")
        api_end = end_date.strftime("%d.%m.%Y")

        self.btn_calculate.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.status_label.setText("API'den veriler \u00e7ekiliyor...")

        self.worker = _DataFetchWorker(self.api_client, api_start, api_end)
        self.worker.progress.connect(self._on_progress)
        self.worker.finished_ok.connect(self._on_data_received)
        self.worker.finished_err.connect(self._on_error)
        self.worker.finished.connect(self._on_worker_done)
        self.worker.start()

    def _on_progress(self, msg: str):
        self.status_label.setText(msg)

    def _on_error(self, msg: str):
        self.status_label.setText(msg)
        QMessageBox.warning(self, "Uyar\u0131", msg)

    def _on_worker_done(self):
        self.btn_calculate.setEnabled(True)
        self.progress_bar.setVisible(False)

    def _on_data_received(self, raw_data: list):
        year = self.year_combo.currentData()
        quarter = self.quarter_combo.currentData()
        months = _get_quarter_months(quarter)
        start_date, end_date = _get_quarter_dates(year, quarter)

        target_map = {}
        for m in months:
            inp = self.target_inputs.get(m)
            if inp and inp.text().strip():
                val = inp.text().strip().replace('.', '').replace(',', '.')
                target_map[m] = Decimal(val)
            else:
                target_map[m] = Decimal('0')

        monthly_data = _process_raw_data(raw_data, start_date, end_date, months, target_map)

        self._populate_results_table(monthly_data)

        self.status_label.setText(f"{len(raw_data)} kay\u0131t i\u015flendi")

    # ---------------------------------------------------------------
    # TABLO VE OZET
    # ---------------------------------------------------------------

    def _populate_results_table(self, monthly_data: dict):
        sorted_months = sorted(monthly_data.keys())
        self.result_table.setRowCount(len(sorted_months))
        self.result_table.setUpdatesEnabled(False)

        total_premium = Decimal('0')
        quarter_total_target = Decimal('0')
        quarter_total_order = Decimal('0')
        quarter_total_invoice = Decimal('0')

        for row, m in enumerate(sorted_months):
            data = monthly_data[m]
            result = PrimCalculator.calculate_monthly_premium(
                data['realized_order'], data['target'], data['realized_invoice']
            )

            quarter_total_target += data['target']
            quarter_total_order += data['realized_order']
            quarter_total_invoice += data['realized_invoice']
            total_premium += result['premium_amount']

            hgo_val = float(result['hgo'])

            if hgo_val >= 100:
                hgo_color = QColor("#28a745")
            elif hgo_val >= 90:
                hgo_color = QColor("#fd7e14")
            else:
                hgo_color = QColor("#dc3545")

            row_data = [
                _get_turkish_month_name(m),
                _format_currency(data['target']),
                _format_currency(data['realized_order']),
                f"%{hgo_val:.1f}",
                _format_currency(data['realized_invoice']),
                f"%{result['rate']}",
                _format_currency(result['premium_amount']),
            ]

            for col, text in enumerate(row_data):
                item = QTableWidgetItem(text)
                item.setTextAlignment(Qt.AlignCenter)
                if col == 0:
                    item.setFont(QFont("Segoe UI", 11, QFont.Bold))
                elif col == 3:
                    item.setForeground(hgo_color)
                    item.setFont(QFont("Segoe UI", 11, QFont.Bold))
                elif col == 6 and result['premium_amount'] > 0:
                    item.setForeground(QColor("#28a745"))
                    item.setFont(QFont("Segoe UI", 11, QFont.Bold))
                self.result_table.setItem(row, col, item)

            self.result_table.setRowHeight(row, 40)

        self.result_table.setUpdatesEnabled(True)

        extra = PrimCalculator.calculate_quarterly_extra_premium(
            quarter_total_order, quarter_total_target, quarter_total_invoice,
            ek_prim_tiers=self.ek_prim_tiers
        )

        self.lbl_total_target.setText(_format_currency(quarter_total_target))
        self.lbl_total_order.setText(_format_currency(quarter_total_order))
        self.lbl_total_invoice.setText(_format_currency(quarter_total_invoice))
        self.lbl_quarter_hgo.setText(f"%{float(extra['hgo']):.2f}")

        if extra['eligible']:
            self.lbl_extra_premium.setText(
                f"KAZANILDI (%{extra['rate']}) = {_format_currency(extra['premium_amount'])}"
            )
            self.lbl_extra_premium.setStyleSheet("color: #28a745; font-weight: bold; background: transparent;")
            total_premium += extra['premium_amount']
        else:
            self.lbl_extra_premium.setText(f"KAZANILMADI ({extra.get('reason', '')})")
            self.lbl_extra_premium.setStyleSheet("color: #dc3545; font-weight: bold; background: transparent;")

        self.lbl_total_premium.setText(f"TOPLAM\n{_format_currency(total_premium)}")

        forecast_lines = PrimCalculator.generate_forecast(
            monthly_data, sorted_months, ek_prim_tiers=self.ek_prim_tiers
        )
        self.forecast_left.clear()
        self.forecast_right.clear()

        right_sections = ['BIREYSEL HEDEF DURUMU', 'EK HACIM PRIMI DURUMU']
        current_target = self.forecast_left
        for line in forecast_lines:
            if line.startswith("--") and line.endswith("--"):
                if any(s in line for s in right_sections):
                    current_target = self.forecast_right
                elif 'STRATEJI' in line:
                    current_target = self.forecast_left

            self._append_forecast_line(current_target, line)

    def _append_forecast_line(self, widget: QTextEdit, line: str):
        if line.startswith("--") and line.endswith("--"):
            widget.append(f'<b style="color: #007acc;">{line}</b>')
        elif "ULA\u015eILDI" in line:
            widget.append(f'<span style="color: #28a745;">{line}</span>')
        elif "eksik" in line or "gerekli" in line:
            widget.append(f'<span style="color: #dc3545;">{line}</span>')
        elif line == "":
            widget.append("")
        else:
            widget.append(line)
