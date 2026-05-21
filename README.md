# BHFL DSA Automation

**One-Click Desktop Application for DSA Mail Distribution**

![Status](https://img.shields.io/badge/Status-Production-green)
![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![OS](https://img.shields.io/badge/OS-Windows-lightblue)
![License](https://img.shields.io/badge/License-BHFL-red)

---

## Overview

BHFL DSA Automation is a fully automated desktop application that eliminates manual steps in DSA mail distribution. With a single click, the application:

- ✅ Reads your master data file
- ✅ Automatically detects unique DSAs
- ✅ Creates DSA-wise Excel files
- ✅ Looks up email addresses
- ✅ Creates professional Outlook emails
- ✅ Attaches files automatically
- ✅ Sends all emails without confirmation
- ✅ Archives processed files
- ✅ Generates comprehensive reports
- ✅ Shows final statistics dashboard

**Zero manual steps. Zero user interaction. Just click and go.**

---

## Key Features

### 🚀 One-Click Operation
- Single "RUN AUTOMATION" button
- No file selection dialogs
- No confirmation popups
- Fully automated workflow

### 📊 Auto-Detection
- Automatically finds Master.xlsx in Input folder
- Automatically loads email map from config folder
- Instantly ready to run

### 📈 Real-Time Monitoring
- Live progress bar (0-100%)
- Live console output
- Step-by-step execution logs
- Current operation display

### 🔄 Intelligent Retry Logic
- Automatic 3-attempt retry for failed mails
- Exponential backoff
- Detailed failure logs

### 📁 Automatic File Management
- Generates DSA-wise Excel files
- Automatically archives processed files
- Keeps Output folder clean
- Maintains Archive for records

### 📋 Comprehensive Reporting
- Status report Excel file (RUN_STATUS.xlsx)
- Detailed logs with timestamps
- Final dashboard with statistics:
  - Total DSAs
  - Mails Sent
  - Failed Attempts
  - Success Rate

### 🎨 Professional GUI
- Clean, intuitive interface
- BHFL branding
- Professional styling
- Multiple action buttons
- Scrollable console output

---

## Quick Start

### Requirements
- Windows 10 or later
- Python 3.8+ (if running from source)
- Microsoft Outlook (for sending mails)
- 4GB RAM minimum

### Installation (5 minutes)

1. **Clone Repository**
   ```bash
   git clone https://github.com/tanishrane24-droid/bhfl-dsa-automation.git
   cd bhfl-dsa-automation
   ```

2. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   pip install PyInstaller
   ```

3. **Build Executable**
   ```bash
   python build_exe.py
   ```

4. **Run**
   ```bash
   dist/BHFL_DSA_AUTOMATION.exe
   ```

### Setup Files (2 minutes)

1. **Create Input/Master.xlsx**
   - Column: "DSA"
   - Add your data with DSA names

2. **Create config/DSA_MAIL_MAP.xlsx**
   - Columns: "DSA", "Email"
   - Map each DSA to email address

3. **Done!** Application auto-detects both files

---

## Usage

```
┌─────────────────────────────────────────┐
│  BHFL DSA AUTOMATION                    │
│                                         │
│  Status:                                │
│  ✓ Master File: Master.xlsx             │
│  ✓ Mail Map: DSA_MAIL_MAP.xlsx          │
│                                         │
│  Progress: [████████░░░░░░░░] 50%      │
│  Creating DSA-wise files...             │
│                                         │
│  Console Output:                        │
│  ✓ Read 500 records                     │
│  ✓ Detected 5 unique DSAs               │
│  ✓ Creating files...                    │
│                                         │
│  [▶ RUN] [OPEN OUTPUT] [LOGS] [EXIT]   │
└─────────────────────────────────────────┘
```

### Just Click: "▶ RUN AUTOMATION"

Everything else happens automatically:
1. Reads master file
2. Identifies DSAs
3. Creates files
4. Sends mails (with retry)
5. Archives files
6. Generates reports
7. Shows final dashboard

---

## Folder Structure

```
bhfl-dsa-automation/
├── main.py                      # Main application
├── build_exe.py                 # EXE builder
├── requirements.txt             # Dependencies
├── setup.py                     # Setup configuration
├── Input/
│   └── Master.xlsx             # Your input data (CREATE)
├── config/
│   └── DSA_MAIL_MAP.xlsx       # Email mapping (CREATE)
├── Output/                      # Generated DSA files (temp)
├── Archive/                     # Processed files
├── Logs/                        # Automation logs
├── assets/
│   └── icon.ico                # Application icon
├── INSTALLATION_GUIDE.md        # Setup instructions
├── USAGE_GUIDE.md              # User manual
└── README.md                    # This file
```

---

## File Specifications

### Input/Master.xlsx
```excel
DSA (Required)    | Field1        | Field2        | ...
ABC Finance       | Data1         | Data2         | ...
XYZ Associates    | Data1         | Data2         | ...
QRS Brokers       | Data1         | Data2         | ...
```

### config/DSA_MAIL_MAP.xlsx
```excel
DSA (Required)      | Email (Required)
ABC Finance         | abc@company.com
XYZ Associates      | xyz@company.com
QRS Brokers         | qrs@company.com
```

---

## Output

### Generated Files

**Archive Folder** (after completion):
- `ABC_Finance.xlsx`
- `XYZ_Associates.xlsx`
- `QRS_Brokers.xlsx`
- etc.

**Logs Folder**:
- `RUN_STATUS_YYYYMMDD_HHMMSS.xlsx` - Status report
- `automation_YYYYMMDD_HHMMSS.log` - Detailed log

### Status Report (RUN_STATUS.xlsx)

| Date | DSA | Mail | Status | Remarks |
|------|-----|------|--------|----------|
| 2024-05-21 14:30 | ABC Finance | abc@company.com | SUCCESS | Mail sent successfully |
| 2024-05-21 14:31 | XYZ Associates | xyz@company.com | SUCCESS | Mail sent successfully |
| 2024-05-21 14:32 | QRS Brokers | qrs@company.com | FAILED | Invalid email |

### Final Dashboard

```
============================================================
AUTOMATION COMPLETED
============================================================
Total DSAs:      5
Mails Sent:      4
Failed:          1
Success Rate:    80.0%
============================================================
```

---

## Features in Detail

### 1. Auto-Detection
✅ No file browsing  
✅ No manual selection  
✅ Just works  

```python
# On startup, automatically looks for:
Input/Master.xlsx
config/DSA_MAIL_MAP.xlsx
```

### 2. Intelligent Processing
✅ Reads master file  
✅ Extracts unique DSA names  
✅ Creates DSA-wise files  
✅ Maps to email addresses  
✅ No manual configuration  

### 3. Automatic Mail Sending
✅ Creates professional emails  
✅ Attaches files automatically  
✅ Sends via Outlook  
✅ No confirmation dialogs  
✅ No manual steps  

### 4. Retry Logic
✅ Failed mail? Retries automatically  
✅ 3 attempts maximum  
✅ Exponential backoff  
✅ Logs all attempts  

### 5. Archival System
✅ Moves processed files  
✅ Keeps records  
✅ Cleans output folder  
✅ Maintains history  

### 6. Comprehensive Logging
✅ Detailed execution logs  
✅ Status report (Excel)  
✅ Timestamp on everything  
✅ Error tracking  

---

## Performance

- **Startup**: < 2 seconds
- **File Detection**: Instant
- **Processing**: ~30 seconds per 100 records
- **Mail Sending**: ~10 seconds per email
- **Complete Run**: 2-5 minutes (depending on data size)

### Tested With
- 500 records, 5 DSAs ✅
- 1000 records, 10 DSAs ✅
- 5000 records, 20 DSAs ✅

---

## Troubleshooting

### "Master File: Not detected"
→ Place Master.xlsx in Input folder and restart

### "Mail Map: Not found"
→ Place DSA_MAIL_MAP.xlsx in config folder and restart

### "No email found for DSA: ABC Finance"
→ Check DSA names match between files (case-sensitive)

### "Outlook is not installed"
→ Install Microsoft Outlook or Office 365

### Mails not sending
→ Check internet, Outlook running, email addresses valid

**Full troubleshooting guide**: See [INSTALLATION_GUIDE.md](INSTALLATION_GUIDE.md)

---

## Development

### Requirements
- Python 3.8+
- pandas
- openpyxl
- pywin32
- tkinter (included with Python)

### Setup Dev Environment
```bash
git clone https://github.com/tanishrane24-droid/bhfl-dsa-automation.git
cd bhfl-dsa-automation
pip install -r requirements.txt
python main.py  # Run from source
```

### Build Executable
```bash
pip install PyInstaller
python build_exe.py
```

### Project Structure
```
main.py           - Main application (600+ lines)
build_exe.py      - PyInstaller builder
setup.py          - Python setup configuration
requirements.txt  - Python dependencies
```

---

## Documentation

- **[INSTALLATION_GUIDE.md](INSTALLATION_GUIDE.md)** - Complete setup instructions
- **[USAGE_GUIDE.md](USAGE_GUIDE.md)** - Detailed user manual
- **[README.md](README.md)** - This file

---

## Changelog

### v1.0.0 (2024)
- Initial release
- Single-click automation
- Auto file detection
- Live progress tracking
- Retry logic (3 attempts)
- Automatic archival
- Comprehensive logging
- Status reporting
- Professional GUI

---

## Support

For issues or questions:
1. Check [INSTALLATION_GUIDE.md](INSTALLATION_GUIDE.md) troubleshooting
2. Review logs in Logs/ folder
3. Contact BHFL support team

---

## License

Proprietary © BHFL 2024

---

## Author

**BHFL Team**

---

## Acknowledgments

Built with:
- Python 3.10+
- pandas (data processing)
- openpyxl (Excel handling)
- pywin32 (Outlook integration)
- tkinter (GUI)

---

## Version

**Current**: 1.0.0  
**Release Date**: 2024

---

**Status**: ✅ Production Ready

**Latest Update**: 2024
