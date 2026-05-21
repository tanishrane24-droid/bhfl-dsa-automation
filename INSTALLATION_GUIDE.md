# BHFL DSA Automation - Installation Guide

## System Requirements
- **OS**: Windows 10 or later
- **Python**: 3.8 or later
- **RAM**: 4GB minimum
- **Disk Space**: 500MB
- **Software**: Microsoft Outlook (for mail sending)

## Installation Steps

### Option 1: Run from Source Code (Recommended for Setup)

#### Step 1: Install Python
1. Download Python 3.10+ from [python.org](https://www.python.org/downloads/)
2. During installation, **CHECK** "Add Python to PATH"
3. Click Install Now

#### Step 2: Clone Repository
```bash
git clone https://github.com/tanishrane24-droid/bhfl-dsa-automation.git
cd bhfl-dsa-automation
```

#### Step 3: Install Dependencies
```bash
pip install -r requirements.txt
pip install PyInstaller
```

#### Step 4: Install pywin32 COM Support
```bash
python -m pip install pywin32
python -m pip install pywin32==305
python Scripts/pywin32_postinstall.py -install
```

#### Step 5: Run Application
```bash
python main.py
```

### Option 2: Build Standalone Executable

#### Step 1-3: Same as above

#### Step 4: Build EXE
```bash
pip install PyInstaller
python build_exe.py
```

Executable will be created at: `dist/BHFL_DSA_AUTOMATION.exe`

#### Step 5: Create Desktop Shortcut
1. Right-click on `BHFL_DSA_AUTOMATION.exe`
2. Send to → Desktop (create shortcut)
3. Double-click to run

### Option 3: Create Installer

#### Using NSIS (Advanced)
```bash
# Install NSIS first from https://nsis.sourceforge.io/
pip install pynsist
# Edit installer_config.cfg
pynsist installer_config.cfg
```

## Setup Configuration Files

### 1. Create Master Excel File
**Location**: `Input/Master.xlsx`

**Required Columns**:
- `DSA` - DSA Name (e.g., "ABC Finance")
- Any other data columns you want to distribute

**Example**:
```
DSA             | Field1  | Field2    | Field3
ABC Finance     | Value1  | Value2    | Value3
XYZ Associates  | Value1  | Value2    | Value3
QRS Brokers     | Value1  | Value2    | Value3
```

### 2. Create Mail Map
**Location**: `config/DSA_MAIL_MAP.xlsx`

**Required Columns**:
- `DSA` - Must match names in Master.xlsx
- `Email` - Recipient email address

**Example**:
```
DSA             | Email
ABC Finance     | abc.finance@company.com
XYZ Associates  | xyz.associates@company.com
QRS Brokers     | qrs.brokers@company.com
```

## Folder Structure

```
bhfl-dsa-automation/
├── main.py                      # Main application
├── requirements.txt             # Python dependencies
├── build_exe.py                 # EXE builder script
├── Input/
│   └── Master.xlsx             # Input data file (CREATE THIS)
├── config/
│   └── DSA_MAIL_MAP.xlsx       # Email mapping (CREATE THIS)
├── Output/                      # Generated DSA files
├── Archive/                     # Processed files backup
├── Logs/                        # Automation logs
└── assets/
    └── icon.ico                 # Application icon (optional)
```

## Troubleshooting

### Issue: "No module named 'win32com'"
**Solution**:
```bash
python -m pip install pywin32
python -m pip install pywin32==305
python Scripts/pywin32_postinstall.py -install
```

### Issue: "Outlook is not installed"
**Solution**: Install Microsoft Outlook or Office 365

### Issue: "File not found: Master.xlsx"
**Solution**: 
1. Create `Input` folder
2. Place `Master.xlsx` inside it
3. Restart application

### Issue: "No email found for DSA"
**Solution**: 
1. Check that DSA names in Master.xlsx match DSA_MAIL_MAP.xlsx
2. Names are case-sensitive

### Issue: "Permission denied when sending mail"
**Solution**:
1. Ensure Outlook is properly configured
2. Enable "Allow apps to send emails" in Outlook security settings
3. Run application as Administrator

## Usage

1. **Open Application**
   - Double-click `BHFL_DSA_AUTOMATION.exe` (or run `python main.py`)

2. **Verify Files**
   - Application auto-detects `Input/Master.xlsx`
   - Application auto-detects `config/DSA_MAIL_MAP.xlsx`
   - Status shown in GUI

3. **Click "RUN AUTOMATION"**
   - No other steps required
   - Process runs fully automatically

4. **Monitor Progress**
   - Live console shows each step
   - Progress bar indicates completion

5. **View Results**
   - Final dashboard shows statistics
   - Click "OPEN OUTPUT" to view generated files
   - Click "VIEW LOGS" to see detailed logs

## Output Files

After automation completes:

1. **Generated Files**: `Output/` folder
   - DSA-wise Excel files (e.g., `ABC_Finance.xlsx`)
   - Later moved to `Archive/` folder

2. **Status Report**: `Logs/RUN_STATUS_YYYYMMDD_HHMMSS.xlsx`
   - Columns: Date, DSA, Mail, Status, Remarks

3. **Log Files**: `Logs/automation_YYYYMMDD_HHMMSS.log`
   - Detailed execution logs

## Advanced Configuration

### Change Retry Attempts
Edit `main.py`, line ~195:
```python
max_retries = 3  # Change this number
```

### Change Mail Subject/Body
Edit `main.py`, lines ~208-220:
```python
mail.Subject = f"Your custom subject"
mail.Body = f"""Your custom message"""
```

### Change Folder Paths
Edit `main.py`, lines ~22-26:
```python
INPUT_DIR = BASE_DIR / "Your_Input_Folder"
OUTPUT_DIR = BASE_DIR / "Your_Output_Folder"
# etc.
```

## Performance Tips

1. Keep Master.xlsx file size under 10MB
2. Limit to 100 DSAs per run
3. Run during off-peak hours
4. Ensure stable internet connection
5. Keep Outlook window minimized during execution

## Support

For issues or questions:
1. Check "VIEW LOGS" for error details
2. Review Troubleshooting section above
3. Verify file formats match examples
4. Contact BHFL support team

## Version History

- **v1.0.0** (2024) - Initial release
  - Single-click automation
  - Auto file detection
  - Live progress tracking
  - Retry logic
  - Status reporting

---

**Created**: 2024
**Last Updated**: 2024
**License**: Proprietary - BHFL
