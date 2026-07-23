# SubhoSpellRupees 🇮🇳

> Automatic Indian Currency Speller for Microsoft Excel — System-wide `.xlam` Add-In with an automated 1-click installation script.

**SubhoSpellRupees** is a lightweight Microsoft Excel Add-In (`.xlam`) that converts numerical amounts into words using the **Indian Numbering System** (*Crores, Lakhs, Thousands, Rupees, and Paise*). 

It includes a single-click Batch/PowerShell installer script that automatically downloads, installs, and registers the Add-In into Excel so it works seamlessly across every workbook on your machine.

---

## 🌟 Features

* **Indian Numbering Format:** Formats numbers into *Crores*, *Lakhs*, *Thousands*, *Rupees*, and *Paise*.
* **System-Wide Availability:** Once installed, `=SpellRupees()` works across **all** current and future Excel workbooks.
* **1-Click Automated Deployment:** Downloads the latest release from GitHub and registers the Add-In via COM automation without requiring admin rights.
* **Protected Core:** The VBA source code is password-locked to preserve function integrity and prevent accidental code changes.

---

## 🚀 Quick Installation

1. Download or clone this repository to your computer.
2. Double-click **`install_addin.bat`** to run the installer.
3. Open Microsoft Excel and start using `=SpellRupees()` right away.

---

## 🧪 Usage & Syntax

In any cell of any Excel sheet, use the formula:

```excel
=SpellRupees(Cell_Reference_or_Number)
