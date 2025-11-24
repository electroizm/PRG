#!/usr/bin/env python3
"""
PRG Modüler Uygulama Başlatıcı - Service Account Versiyonu

Kullanım:
    # OAuth2 klasöründen çalıştırın:
    cd "D:/GoogleDrive/PRG/OAuth2"
    python PRG/run.py

Not:
    - Tüm ayarlar PRGsheet/Ayar sayfasından çekilir
    - .env dosyası kullanılmaz
    - Service Account credentials gereklidir (service_account.json)
"""

import sys
import os

# Current directory'yi Python path'e ekle
# Bu sayede central_config modülü import edilebilir
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
sys.path.insert(0, parent_dir)

if __name__ == "__main__":
    try:
        from main import main  # type: ignore
        main()
    except ImportError as e:
        print("ERROR: Module import error:", str(e))
        print("\nPlease check:")
        print("   1. Are all required libraries installed?")
        print("      pip install -r requirements.txt")
        print("   2. Is central_config.py present in parent directory?")
        print("      D:/GoogleDrive/PRG/OAuth2/central_config.py")
        print("   3. Is Python path correct?")
        print("   4. Run from OAuth2 directory:")
        print("      cd D:/GoogleDrive/PRG/OAuth2 && python PRG/run.py")
        print("   5. Service Account credentials present?")
        print("      D:/GoogleDrive/PRG/OAuth2/service_account.json")
        print("\nNote: .env file is no longer used - all settings come from PRGsheet/Ayar")
        sys.exit(1)
    except Exception as e:
        print("ERROR: Application startup error:", str(e))
        import traceback
        traceback.print_exc()
        sys.exit(1)