# PRG - Enterprise Management System

[![Python](https://img.shields.io/badge/Python-3.13-blue.svg)](https://www.python.org/)
[![PyQt5](https://img.shields.io/badge/PyQt5-5.15+-green.svg)](https://www.riverbankcomputing.com/software/pyqt/)
[![License](https://img.shields.io/badge/license-Private-red.svg)]()

**PRG** is a comprehensive enterprise management system built with PyQt5, designed for managing various business operations including inventory, contracts, shipping, financial transactions, and more.

## 📋 Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Architecture](#architecture)
- [Modules](#modules)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [Development](#development)
- [Recent Updates](#recent-updates)
- [Technical Stack](#technical-stack)

## 🌟 Overview

PRG is a modular enterprise management application that integrates with:
- **Google Sheets** (via Service Account) for data storage and synchronization
- **Microsoft SQL Server** (Mikro ERP) for financial and inventory data
- **Email services** for automated notifications
- **WhatsApp** for customer communication

**Statistics:**
- **22 Python files**
- **22,456+ lines of code**
- **12 functional modules**
- **Modern PyQt5 UI with dark/light themes**

## ✨ Features

### Core Features
- 🔐 **Centralized Configuration** - Service Account based authentication
- 🎨 **Modern UI** - Clean, responsive PyQt5 interface
- 💾 **Global Data Cache** - Efficient data management with caching
- 📊 **Real-time Data Sync** - Bidirectional sync with Google Sheets
- 🔄 **Lazy Loading** - Optimized performance with on-demand data loading
- 🎯 **Focus Border Free** - Improved UX with clean table selection
- 📱 **Multi-platform Support** - Windows-optimized with EXE packaging

### Business Features
- 📦 **Inventory Management** - Complete stock tracking and management
- 📝 **Contract Management** - Full contract lifecycle management
- 🚚 **Shipping Operations** - Comprehensive shipping and logistics
- 💰 **Financial Tracking** - Cash register, transfers, and POS operations
- ⚠️ **Risk Management** - Customer risk analysis and monitoring
- 🔐 **SSH Management** - Secure shell access and management
- 💳 **Payment Processing** - Virtual POS and payment tracking
- 📄 **Document Management** - Waybills and invoices

## 🏗️ Architecture

### Core Components

#### `run.py` - Application Entry Point
Entry point for the PRG application. Handles:
- Python path configuration
- Module initialization
- Error handling and diagnostics
- Service Account setup verification

#### `main.py` - Main Application Logic
Contains the main application window and core logic:
- **GlobalDataCache** - Centralized data caching system
- **PRGMainWindow** - Main window with tabbed interface
- Module integration and lifecycle management
- Global data refresh mechanism

#### `core_architecture.py` - Architectural Foundation
Modern architecture patterns:
- **EventType & ModuleType** - Event-driven architecture
- **Theme** - UI theming system
- **EventBus** - Inter-module communication
- **ModuleRegistry** - Dynamic module loading

#### `ui_components.py` - UI Components
Reusable UI components and widgets

#### `embedded_resources.py` - Resource Management
Application icons and embedded resources

## 📦 Modules

### 1. **Stok Module** (`stok_module.py`)
**Inventory & Stock Management**

Comprehensive stock management system with:
- Real-time stock levels (DEPO, EXCLUSIVE, SUBE)
- Shopping cart (Sepet) management
- Advanced filtering and search
- SQL Server integration for Mikro ERP data
- Price calculations with KDV and margins
- Excel import/export functionality
- Context menu for quick actions

**Key Features:**
- Multi-warehouse support
- Automated price calculations
- Real-time stock updates
- Smart search with fuzzy matching
- Editable shopping cart
- Focus border removed for clean UX

---

### 2. **Sevkiyat Module** (`sevkiyat_module.py`)
**Shipping & Logistics Management**

Complete shipping operations management:
- Customer search with autocomplete
- Multi-tab shipping data (Sevkiyat, Bekleyenler, Araç, Malzeme)
- WhatsApp integration for notifications
- Email notifications
- Excel export for all tabs
- Risk analysis integration
- Mikro ERP integration

**Key Features:**
- Fuzzy customer name matching
- Contract product lookup
- Vehicle and material tracking
- Automated email/WhatsApp messaging
- Multi-view data filtering
- Custom date range filtering
- Focus border removed from customer list

---

### 3. **Sozlesme Module** (`sozlesme_module.py`)
**Contract Management**

Advanced contract lifecycle management:
- Contract details viewing
- Product line items management
- Customer and order information
- Mikro ERP integration (Cari, Stok, Sipariş)
- IPT status tracking
- Header information management
- Multi-table data display

**Key Features:**
- Contract search and filtering
- Customer selection dialog
- Product table editing
- SAP/ERP transfer operations
- Stock card creation
- Order transfer
- Focus border removed from 3 tables

---

### 4. **Risk Module** (`risk_module.py`)
**Customer Risk Analysis**

Customer credit and risk management:
- Risk level monitoring
- Credit limit tracking
- Payment history analysis
- Mikro ERP data integration
- Excel export capabilities
- Automated risk updates

**Key Features:**
- Real-time risk calculations
- Color-coded risk indicators
- Threshold-based alerts
- Historical risk tracking
- Focus border removed for clean tables

---

### 5. **OKC Module** (`okc_module.py`)
**OKC YazarKasa Management**

Cash register and payment management:
- Invoice tracking
- Payment amount filtering
- Date formatting (removed 00:00 time display)
- Excel export
- Mikro ERP integration
- Quick navigation

**Key Features:**
- Amount-based filtering (1000 TL multiplier)
- Invoice date management
- Payment tracking
- Color-coded status indicators
- Clean date display (DD.MM.YYYY)

---

### 6. **SSH Module** (`ssh_module.py`)
**Secure Shell Management**

SSH connection and management system:
- Connection management
- Two-table interface for different SSH data views
- Status monitoring
- Quick actions
- Print support

**Key Features:**
- Multi-table SSH data display
- Connection status tracking
- Print functionality
- Focus border removed from 2 tables
- Real-time updates

---

### 7. **Kasa Module** (`kasa_module.py`)
**Cash Register Operations**

Financial transaction management:
- Monthly cash register data
- Year/month filtering
- Transaction categorization
- Excel export
- Balance calculations

**Key Features:**
- Monthly view with current date default
- Color-coded transaction types
- Balance tracking
- Quick navigation
- Export capabilities

---

### 8. **Sanalpos Module** (`sanalpos_module.py`)
**Virtual POS Management**

Online payment processing and tracking:
- POS transaction monitoring
- Payment status tracking
- Date-based filtering
- Excel export
- Integration with Kasa data

**Key Features:**
- Real-time POS data
- Transaction history
- Status indicators
- QApplication import fix applied
- Export functionality

---

### 9. **Irsaliye Module** (`irsaliye_module.py`)
**Waybill Management**

Shipping document management:
- Waybill creation and tracking
- Multi-tab interface
- Document export
- Customer assignment
- Date tracking

**Key Features:**
- Tab-based organization
- Document search
- Export to Excel
- Context menu with copy function
- Focus border removed
- Bold font styling

---

### 10. **Fiyat Module** (`fiyat_module.py`)
**Price & Label Management**

Product pricing and labeling:
- SAP code generation
- Price list management
- Stock data integration
- Label printing preparation
- Excel export/import

**Key Features:**
- Automated SAP code creation
- Multi-source data integration (DEPO, EXC, SUBE)
- Price calculation
- Batch processing
- Threading for performance

---

### 11. **Virman Module** (`virman_module.py`)
**Transfer Management**

Inter-account transfer operations:
- Account transfer tracking
- Monthly data view
- Balance verification
- SQL Server integration
- Transaction history

**Key Features:**
- Month-based filtering
- Transfer verification
- Balance checking
- Transaction logging
- Real-time updates

---

### 12. **Ayar Module** (`ayar_module.py`)
**Settings & Configuration**

System configuration management:
- Multi-tab settings (Ayar, Mail, NoRisk)
- Google Sheets integration
- Configuration editing
- Settings persistence
- Lazy loading optimization

**Key Features:**
- Tab-based organization
- Direct Google Sheets editing
- Configuration validation
- Save/Reload functionality
- Real-time updates

## 🚀 Installation

### Prerequisites

```bash
# Python 3.13+
python --version

# Required packages
pip install -r requirements.txt
```

### Required Dependencies

```
PyQt5>=5.15.0
pandas>=2.0.0
numpy>=1.24.0
requests>=2.31.0
gspread>=5.0.0
google-auth>=2.0.0
openpyxl>=3.1.0
pyodbc>=4.0.0
python-dotenv>=1.0.0
fuzzywuzzy>=0.18.0
python-levenshtein>=0.21.0
pyperclip>=1.8.0
cryptography>=41.0.0
```

### Service Account Setup

1. Create a Google Cloud project
2. Enable Google Sheets API
3. Create a Service Account
4. Download `service_account.json`
5. Place in parent directory (`D:/GoogleDrive/PRG/OAuth2/`)
6. Share Google Sheets with service account email

### Configuration

Create `central_config.py` in parent directory:

```python
class CentralConfigManager:
    MASTER_SPREADSHEET_ID = "your_spreadsheet_id_here"
    # ... other configuration
```

## 💻 Usage

### Running the Application

```bash
# From OAuth2 directory
cd D:/GoogleDrive/PRG/OAuth2
python PRG/run.py
```

### Building Executable

```bash
# Using PyInstaller
pyinstaller PRG_onefile.spec --clean --noconfirm
```

The executable will be created in `dist/PRG.exe` (~76MB).

## 🛠️ Development

### Project Structure

```
PRG/
├── run.py                  # Entry point
├── main.py                 # Main application
├── core_architecture.py    # Architecture patterns
├── ui_components.py        # UI components
├── embedded_resources.py   # Resources
├── ayar_module.py          # Settings
├── stok_module.py          # Inventory
├── sevkiyat_module.py      # Shipping
├── sozlesme_module.py      # Contracts
├── risk_module.py          # Risk management
├── okc_module.py           # Cash register
├── ssh_module.py           # SSH management
├── kasa_module.py          # Cash operations
├── sanalpos_module.py      # Virtual POS
├── irsaliye_module.py      # Waybills
├── fiyat_module.py         # Pricing
├── virman_module.py        # Transfers
├── icon.ico                # Application icon
└── icon.jpg                # Icon source
```

### Code Style

- **PEP 8** compliance
- **Type hints** where applicable
- **Docstrings** for all modules and classes
- **Constants** for configuration values
- **Centralized styling** via stylesheet constants

### Architecture Patterns

- **Lazy Loading** - Data loaded only when needed
- **Global Cache** - Shared data cache across modules
- **Event Bus** - Inter-module communication
- **Module Registry** - Dynamic module loading
- **Service Account** - Centralized authentication

## 🔄 Recent Updates

### UI/UX Improvements
- ✅ **Focus Border Removal** - Clean table selection across all modules
  - stok_module.py - Table widgets
  - sevkiyat_module.py - Customer list
  - sozlesme_module.py - 3 tables (products_table, dialog table, main table)
  - risk_module.py - Risk table
  - okc_module.py - OKC table
  - ssh_module.py - 2 SSH tables
  - irsaliye_module.py - Document tables
  - CSS: `QTableWidget::item:focus { outline: none; border: none; }`
  - Policy: `setFocusPolicy(Qt.NoFocus)`

### Bug Fixes
- ✅ **Date Format Fix** - okc_module.py
  - Changed from `strftime('%d.%m.%Y %H:%M')` to `strftime('%d.%m.%Y')`
  - Removed "00:00" from date displays
  - Cleaner date presentation

- ✅ **Import Fix** - sanalpos_module.py
  - Added QApplication import
  - Fixed NameError in clipboard operations

### Style Improvements
- ✅ **Constants Architecture** - irsaliye_module.py
  - Added CONFIG CONSTANTS section
  - Added STYLESHEET CONSTANTS section
  - Bold font implementation
  - Context menu with copy function

## 🔧 Technical Stack

### Frontend
- **PyQt5** - GUI framework
- **QTableWidget** - Data display
- **QTabWidget** - Multi-view interface
- **Custom Stylesheets** - Modern styling

### Backend
- **pandas** - Data manipulation
- **numpy** - Numerical operations
- **requests** - HTTP requests
- **pyodbc** - SQL Server connectivity

### Integration
- **gspread** - Google Sheets API
- **google-auth** - Service Account authentication
- **cryptography** - Secure data handling

### Tools
- **PyInstaller** - Executable packaging
- **openpyxl** - Excel file handling
- **fuzzywuzzy** - Fuzzy string matching

## 📝 License

This is proprietary software. All rights reserved.

## 👥 Authors

**PRG Development Team**

## 🤝 Contributing

This is a private project. Contributions are managed internally.

## 📞 Support

For internal support, contact the development team.

---

**Generated with ❤️ by PRG Development Team**

Last Updated: November 24, 2025
