# SubhoSpellRupees

> Indian Currency Speller for Microsoft Excel — converts numbers to words using the Indian numbering system (Crores, Lakhs, Thousands, Rupees, and Paise).

---

## What it does

`=SpellRupees(A1)` turns a number in a cell into spelled-out Indian currency:

| Input | Output |
|-------|--------|
| `1500000` | `Fifteen Lakhs Rupees Only` |
| `10000000` | `One Crore Rupees Only` |
| `75250.50` | `Seventy Five Thousand Two Hundred Fifty Rupees and Fifty Paise Only` |
| `0` | `Zero Rupees Only` |

The function works across **all** Excel workbooks on your machine once installed — no need to copy anything into individual files.

---

## Installation

### Option 1 — 1-Click Installer (Recommended)

1. **Download** [`install_addin.bat`](https://github.com/visiontech-com-ai/subhospellrupees/releases/latest/download/install_addin.bat) from the latest release.
2. **Double-click** `install_addin.bat` to run it.
3. **Open (or restart) Excel** — `=SpellRupees()` is ready to use.

The installer:
- Downloads `SubhoSpellRupees.xlam` directly from this repository
- Places it in `%APPDATA%\Microsoft\AddIns\` (the standard Excel add-ins folder)
- Registers it in the Windows Registry so Excel loads it automatically on every startup
- Requires no admin rights

### Option 2 — Manual Installation

1. Download [`SubhoSpellRupees.xlam`](https://github.com/visiontech-com-ai/subhospellrupees/releases/latest/download/SubhoSpellRupees.xlam) from the latest release.
2. Copy it to `%APPDATA%\Microsoft\AddIns\`
   - Press `Win + R`, type `%APPDATA%\Microsoft\AddIns`, press Enter.
3. In Excel, go to **File → Options → Add-ins**.
4. At the bottom, set **Manage** to **Excel Add-ins** and click **Go**.
5. Click **Browse**, navigate to the file, and click **OK**.
6. Tick the **SubhoSpellRupees** checkbox and click **OK**.

---

## Usage

```excel
=SpellRupees(number_or_cell)
```

### Examples

```excel
=SpellRupees(1000)
→ One Thousand Rupees Only

=SpellRupees(A2)
→ (spells out whatever number is in A2)

=SpellRupees(SUM(B2:B10))
→ (spells out the sum of a range)
```

### Notes

- Input can be a literal number, a cell reference, or any formula that returns a number.
- Amounts up to **99,99,99,99,999** (99 Arab 99 Crore 99 Lakh 99 Thousand 999) are supported.
- Paise (decimal) values are rounded to 2 decimal places.
- Negative numbers return an error — negate the input before spelling if needed.

---

## Requirements

- Microsoft Excel 2010 or later (Windows)
- Windows 7 or later
- Internet connection (installer only, for the initial download)

---

## Uninstalling

### Using the installer machine

1. In Excel, go to **File → Options → Add-ins → Manage: Excel Add-ins → Go**.
2. Uncheck **SubhoSpellRupees** and click **OK**.
3. Delete `%APPDATA%\Microsoft\AddIns\SubhoSpellRupees.xlam`.

### Fully clean (registry)

Open PowerShell and run:

```powershell
$key = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Excel\Options'
Get-ItemProperty $key | Select-Object -Property OPEN* |
  Get-Member -MemberType NoteProperty |
  Where-Object { (Get-ItemProperty $key).$($_.Name) -like '*SubhoSpellRupees*' } |
  ForEach-Object { Remove-ItemProperty -Path $key -Name $_.Name }
```

Replace `16.0` with your Office version (`15.0` for 2013, `14.0` for 2010).

---

## Security & Privacy

- The VBA source code inside the add-in is **password-protected** to preserve function integrity.
- The installer downloads files only from this public GitHub repository over HTTPS (TLS 1.2+).
- No data leaves your machine when using `=SpellRupees()` — all computation is local.
- No admin rights are required at any point.

---

## Troubleshooting

**"Failed to download file"** — Check your internet connection. If you are behind a corporate proxy, download the `.xlam` manually (Option 2 above).

**Function not found after install** — Close Excel completely and reopen it. If the issue persists, check that `SubhoSpellRupees.xlam` exists in `%APPDATA%\Microsoft\AddIns\`.

**Macro security warning** — Go to **File → Options → Trust Center → Trust Center Settings → Trusted Locations** and add `%APPDATA%\Microsoft\AddIns\`. This folder is the standard trusted location for Excel add-ins.

**Running on older Office (2010/2013)** — The installer tries Office versions 16.0, 15.0, and 14.0 in order. If your registry path differs, use the manual installation method.

---

## License

This project is released for free personal and commercial use. The VBA source is protected; redistribution of modified versions is not permitted.

---

*Developed by Subho · Distributed by DwiSha Ventures Sdn Bhd, Distributed by VisionTech*
