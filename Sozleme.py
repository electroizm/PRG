import sys
import os

# Parent directory'yi Python path'e ekle (central_config için)
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

from zeep import Client, Settings
from zeep.transports import Transport
from requests import Session
from requests.auth import HTTPBasicAuth
import xml.etree.ElementTree as ET
from zeep.plugins import HistoryPlugin
from lxml import etree
from datetime import datetime, timedelta
from central_config import CentralConfigManager

def load_config_from_sheets():
    """PRGsheet Google Sheets'ten Ayar sayfasından yapılandırma bilgilerini yükler - Service Account ile"""
    try:
        # Service Account kullanan merkezi config manager'ı kullan
        config_manager = CentralConfigManager()

        # CentralConfigManager'ın get_settings metodunu kullan (daha güvenilir)
        settings = config_manager.get_settings()

        if not settings:
            print("UYARI: PRGsheet/Ayar sayfasından ayarlar yüklenemedi!")
            return {}

        return settings

    except Exception as e:
        print(f"Google Sheets'ten yapılandırma yükleme hatası: {e}")
        print("Lütfen Service Account credentials'ınızın geçerli olduğundan ve Ayar sayfasının mevcut olduğundan emin olun.")
        import traceback
        traceback.print_exc()
        # Güvenlik için varsayılan değerler kaldırıldı
        return {}

# Config cache (her sorguda yeniden yüklenebilir)
_config_cache = None

def get_config(force_refresh=False):
    """Config'i yükle veya cache'den döndür"""
    global _config_cache
    if _config_cache is None or force_refresh:
        _config_cache = load_config_from_sheets()
    return _config_cache

def get_setting(key, default=None):
    """Belirli bir ayarı config'den al"""
    config = get_config()
    return config.get(key, default)

# İlk config yükleme (uyumluluk için)
config = get_config()
SERVICE_URL = config.get('SERVICE_URL')
SERVICE_USERNAME = config.get('SERVICE_USERNAME')
SERVICE_PASSWORD = config.get('SERVICE_PASSWORD')
BAYI_USERNAME = config.get('BAYI_USERNAME')
BAYI_PASSWORD = config.get('BAYI_PASSWORD')
BAYI_KODU = config.get('BAYI_KODU')

def print_xml(xml_string):
    """XML verisini okunabilir şekilde yazdır"""
    root = ET.fromstring(xml_string)
    ET.indent(root, space="  ", level=0)
    print(ET.tostring(root, encoding='unicode'))

def safe_get(obj, attr, default='N/A'):
    """Güvenli veri çekme fonksiyonu"""
    if not obj:
        return default
    return getattr(obj, attr, default) if hasattr(obj, attr) else default

def format_contract_report(contract_info):
    """Sözleşme bilgilerini güzel formatlı rapor olarak hazırlar"""
    if not contract_info:
        return "Sözleşme bilgisi alınamadı"
    
    report = []
    report.append("=" * 80)
    report.append("                        SÖZLEŞME BİLGİLERİ")
    report.append("=" * 80)
    
    # Müşteri bilgileri
    report.append("\n[MÜŞTERİ] MÜŞTERİ BİLGİLERİ:")
    report.append("-" * 40)
    customer_name = f"{safe_get(contract_info, 'CUSTOMER_NAMEFIRST')} {safe_get(contract_info, 'CUSTOMER_NAMELAST')}".strip()
    report.append(f"Ad Soyad      : {customer_name}")
    report.append(f"Telefon 1     : {safe_get(contract_info, 'CUSTOMER_PHONE1')}")
    report.append(f"Telefon 2     : {safe_get(contract_info, 'CUSTOMER_PHONE2')}")
    report.append(f"E-mail        : {safe_get(contract_info, 'CUSTOMER_MAIL')}")
    report.append(f"Vergi No      : {safe_get(contract_info, 'CUSTOMER_TAXNR')}")
    report.append(f"Vergi Dairesi : {safe_get(contract_info, 'CUSTOMER_TAXOFFICE')}")
    report.append(f"Şehir/İlçe    : {safe_get(contract_info, 'CUSTOMER_CITY')}/{safe_get(contract_info, 'CUSTOMER_DISTRICT')}")
    report.append(f"Adres         : {safe_get(contract_info, 'CUSTOMER_ADDRESS')}")
    report.append(f"Posta Kodu    : {safe_get(contract_info, 'CUSTOMER_POSTCODE')}")
    
    # Sipariş bilgileri
    header_text = safe_get(contract_info, 'HEADER_TEXT', '0')
    report.append("\n[SİPARİŞ] SİPARİŞ BİLGİLERİ:")
    report.append("-" * 40)
    report.append(f"Sipariş Tarihi: {safe_get(contract_info, 'ORD_DATE')}")
    report.append(f"Durum         : {safe_get(contract_info, 'STATUS_TEXT')} ({safe_get(contract_info, 'STATUS')})")
    report.append(f"Fiyat Listesi : {safe_get(contract_info, 'PRICE_LIST_TEXT')}")
    report.append(f"Toplam Tutar  : {header_text} TL")
    
    # Satış temsilcisi
    report.append("\n[SATIŞ] SATIŞ TEMSİLCİSİ:")
    report.append("-" * 40)
    salesman_name = f"{safe_get(contract_info, 'SALESMAN_NAMEFIRST')} {safe_get(contract_info, 'SALESMAN_NAMELAST')}".strip()
    report.append(f"Ad Soyad      : {salesman_name}")
    report.append(f"Satış Ofisi   : {safe_get(contract_info, 'SALES_OFFICE')}")
    
    # Teslimat bilgileri
    report.append("\n[TESLİMAT] TESLİMAT BİLGİLERİ:")
    report.append("-" * 40)
    del_customer_name = f"{safe_get(contract_info, 'DEL_CUSTOMER_NAMEFIRST')} {safe_get(contract_info, 'DEL_CUSTOMER_NAMELAST')}".strip()
    report.append(f"Teslim Alan   : {del_customer_name}")
    report.append(f"Telefon       : {safe_get(contract_info, 'DEL_CUSTOMER_PHONE1')}")
    report.append(f"Adres         : {safe_get(contract_info, 'DEL_CUSTOMER_ADDRESS')}")
    report.append(f"Şehir/İlçe    : {safe_get(contract_info, 'DEL_CUSTOMER_CITY')}")
    report.append(f"Posta Kodu    : {safe_get(contract_info, 'DEL_CUSTOMER_POSTCODE')}")
    
    # Ürünler
    if hasattr(contract_info, 'ITEMS') and hasattr(contract_info.ITEMS, 'item'):
        report.append("\n[ÜRÜNLER] ÜRÜNLER:")
        report.append("-" * 40)
        
        total_net = 0
        total_tax = 0
        
        for i, item in enumerate(contract_info.ITEMS.item, 1):
            product_code = safe_get(item, 'PRODUCT_CODE')
            description = safe_get(item, 'DESCRIPTION')
            quantity = safe_get(item, 'QUANTITY')
            unit_price = float(safe_get(item, 'UNIT_PRICE', '0'))
            total_price = float(safe_get(item, 'TOTAL_PRICE', '0'))
            net_amount = float(safe_get(item, 'NET_AMOUNT', '0'))
            tax_amount = float(safe_get(item, 'TAX_AMOUNT', '0'))
            tax_rate = safe_get(item, 'TAX_RATE', '0')
            discount = float(safe_get(item, 'TOTAL_DISCOUNT', '0'))
            
            total_net += net_amount
            total_tax += tax_amount
            
            siparis = safe_get(item, 'SIPARIS')
            sip_kalem_no = safe_get(item, 'SIP_KALEM_NO')
            kalem_no = safe_get(item, 'KALEM_NO')
            
            report.append(f"\n{i}. ÜRÜN:")
            report.append(f"   Kod           : {product_code}")
            report.append(f"   Açıklama      : {description}")
            report.append(f"   Miktar        : {quantity} adet")
            report.append(f"   Birim Fiyat   : {unit_price:,.2f} TL")
            report.append(f"   Toplam Fiyat  : {total_price:,.2f} TL")
            report.append(f"   Net Tutar     : {net_amount:,.2f} TL")
            report.append(f"   KDV ({tax_rate}%)     : {tax_amount:,.2f} TL")
            if discount != 0:
                report.append(f"   İndirim       : {discount:,.2f} TL")
            report.append(f"   Sipariş No    : {siparis}")
            report.append(f"   Sip Kalem No  : {sip_kalem_no}")
            report.append(f"   Kalem No      : {kalem_no}")
            
            # Ürün özellikleri varsa
            if hasattr(item, 'SPEC') and hasattr(item.SPEC, 'item'):
                report.append("   Özellikler    :")
                for spec_item in item.SPEC.item:
                    charc = safe_get(spec_item, 'CHARC')
                    value = safe_get(spec_item, 'VALUE')
                    report.append(f"     - {charc}: {value}")
        
        # Genel toplam
        report.append("\n[TOPLAM] GENEL TOPLAM:")
        report.append("-" * 40)
        report.append(f"Net Toplam    : {total_net:,.2f} TL")
        report.append(f"KDV Toplam    : {total_tax:,.2f} TL")
        report.append(f"Genel Toplam  : {(total_net + total_tax):,.2f} TL")
    
    report.append("\n" + "=" * 80)
    
    return "\n".join(report)

def process_contract_response(response, history):
    """Sözleşme yanıtını işler ve güzel formatlı rapor yazdırır"""
    
    if hasattr(response, 'ES_CONTRACT_INFO'):
        contract_info = response.ES_CONTRACT_INFO
        
        # Güzel formatlı rapor - konsol çıktısı kaldırıldı
        # print(format_contract_report(contract_info))
        
    else:
        # Sözleşme bilgisi alınamadı - konsol çıktısı kaldırıldı
        pass
    
    # Sistem mesajlarını göster
    # Sistem mesajları - konsol çıktısı kaldırıldı
    # if hasattr(response, 'ET_RETURN') and hasattr(response.ET_RETURN, 'item'):
    #     for msg in response.ET_RETURN.item:
    #         msg_type = {
    #             'I': 'Bilgi',
    #             'W': 'Uyarı',
    #             'E': 'Hata',
    #             'S': 'Başarı'
    #         }.get(safe_get(msg, 'MESSAGE_TYPE'), 'Bilinmeyen')

    return response


def get_all_contract_info(contract_id):
    try:
        # Her sorguda config'i yeniden yükle (güncel ayarları almak için)
        config = get_config(force_refresh=True)
        service_url = config.get('SERVICE_URL')
        service_username = config.get('SERVICE_USERNAME')
        service_password = config.get('SERVICE_PASSWORD')
        bayi_username = config.get('BAYI_USERNAME')
        bayi_password = config.get('BAYI_PASSWORD')

        # Debug için history plugin ekle
        history = HistoryPlugin()

        # Zeep ayarları
        settings = Settings(
            strict=False,
            xml_huge_tree=True,
            extra_http_headers={'Content-Type': 'text/xml; charset=utf-8'}
        )

        # Oturum oluştur
        session = Session()
        session.auth = HTTPBasicAuth(service_username, service_password)
        transport = Transport(session=session, timeout=30)

        # SOAP istemcisi oluştur
        client = Client(
            service_url,
            transport=transport,
            settings=settings,
            plugins=[history]
        )
        

        # Servis çağrısı yap
        response = client.service.ZCRM_CONTRACT_INFO_GET_RFC(
            IV_CONTRACT_ID=contract_id,
            IV_USERNAME=bayi_username,
            IV_PASSWORD=bayi_password
        )
        
        return process_contract_response(response, history)

    except Exception as e:
        print(f"\n[HATA] SOAP Servis Hatasi: {str(e)}")
        print(f"       Hata Tipi: {type(e).__name__}")

        if hasattr(e, 'detail'):
            print(f"       Hata Detayi: {e.detail}")

        # Detaylı hata bilgisi
        import traceback
        print("\n--- Detayli Hata Bilgisi ---")
        traceback.print_exc()
        print("----------------------------\n")

        return None

# Ana program
if __name__ == "__main__":
    print("\n" + "="*80)
    print("         DOĞTAŞ KELEBEK CRM - SÖZLEŞME BİLGİ SORGULAMA")
    print("="*80)

    # Ayarlar kontrolü
    if not all([SERVICE_URL, SERVICE_USERNAME, SERVICE_PASSWORD, BAYI_USERNAME, BAYI_PASSWORD]):
        print("\n[HATA] Gerekli servis ayarlari eksik!")
        print("Lutfen PRGsheet/Ayar sayfasini kontrol edin.")
        import sys
        sys.exit(1)

    print(f"\n[OK] Servis ayarlari yuklendi")
    print(f"[OK] Bayi: {BAYI_USERNAME}")
    print(f"[OK] Bayi Kodu: {BAYI_KODU}")

    # Tek sözleşme sorgulama
    contract_id = input("\nSorgulanacak sozlesme numarasini girin: ")
    if contract_id.strip():
        print(f"\n[SORGU] Sozlesme {contract_id.strip()} sorgulanıyor...")
        result = get_all_contract_info(contract_id.strip())

        if result:
            print(f"\n[BASARILI] Sozlesme {contract_id.strip()} basariyla sorgulandı!")
            print("\n" + format_contract_report(result.ES_CONTRACT_INFO if hasattr(result, 'ES_CONTRACT_INFO') else None))
        else:
            print(f"\n[HATA] Sozlesme {contract_id.strip()} bulunamadı veya hata olustu.")
            print("Yukaridaki hata detaylarini kontrol edin.")
    else:
        print("\n[HATA] Gecersiz sozlesme numarasi!")