<div align="center">
  
  # 🚀 Monthly Rollover Automation Suite
  
  **A comprehensive Python toolkit to automate the monthly transition of daily operational and financial Excel spreadsheets.**

  ![Python Version](https://img.shields.io/badge/Python-3.8+-blue.svg)
  ![OpenPyXL](https://img.shields.io/badge/Dependency-OpenPyXL-brightgreen)
  ![CustomTkinter](https://img.shields.io/badge/GUI-CustomTkinter-darkblue)
  
  <br />
</div>

## 📖 Overview

The **Monthly Rollover Automation Suite** takes the pain out of end-of-month administrative tasks. Built specifically for retail/service station environments, this suite of scripts automatically generates the next month's Excel workbooks based on the latest available data. 

It handles everything from copying closing balances to opening balances, dynamically updating date columns, clearing daily entry ranges, and even managing financial tax year folder structures (March to February).

## ✨ Key Features

- 🎨 **Sleek GUI Launcher:** A modern, dark-mode ready graphical interface built with `customtkinter` to run multiple rollovers in a single click.
- 🔄 **Intelligent Carry-Overs:** Automatically pulls closing balances (e.g., stock levels, cash amounts) from the previous month and inserts them as the new month's opening balances.
- 📅 **Dynamic Date Handling:** Automatically generates correct daily dates for the new month, properly handling leap years and varying month lengths.
- 📁 **Tax Year Detection:** Automatically detects when a rollover crosses from February to March and builds the correct new `Tax Year` directory structure.
- 🛡️ **Dry-Run Mode:** Built-in `--dry-run` functionality allows you to test scripts and view output logs without making any actual changes to the filesystem.
- 🛠️ **Standalone Flexibility:** Scripts can be run all together via the GUI or individually via the Command Line.

## 📦 Supported Rollovers

This suite currently automates the rollover process for the following daily procedure documents:

| Module | Description |
| :--- | :--- |
| 💵 **Cash Ups** | Rolls over the daily cash control sheets and updates batch labels. |
| 📉 **Cashiers Shorts & Overs** | Prepares the cashier tracking sheet with fresh dates and cleared entries. |
| 🧾 **Daily Bill Payment Recon** | Appends a new monthly sheet and links the previous month's balances. |
| 💧 **Daily Drinking Water** | Rolls over the water stock recon, carrying over ending totals. |
| 🧊 **Daily Ice Stock** | Rolls over the ice recon, moving the final stock count to the new month. |
| ⛽ **Daily Sales** | Manages fuel sales (ULP 95, Diesel ULS), updating date matrices and total sales. |
| 🛢️ **Gas Bottles** | Calculates final gas bottle counts and carries them forward. |
| 💳 **Speed Point Control** | Updates daily card machine tracking and sequence numbers. |
| 📊 **FMS Recon** | Appends and clears the new month's FMS reconciliation sheet. |

## 🚀 Getting Started

### Prerequisites

You will need Python 3.8 or newer. Install the required dependencies using `pip`:

```bash
pip install openpyxl customtkinter
```

### Folder Structure

The tool expects to run in or be pointed to a root directory containing a `DAILY PROCEDURE DOCUMENTS` folder (and an `FMS Recon` folder) with the respective subdirectories for each task.

### Running the GUI Application

To launch the interactive GUI, simply run:

```bash
python monthly_rollover_launcher.py
```
*From the GUI, you can select which specific modules to run, toggle "Dry run", and view real-time log outputs.*

### Running via Command Line (CLI)

You can run any script individually. For example, to roll over the Gas Bottles workbook:

```bash
python gas_bottles_rollover.py --root /path/to/your/files
```

**CLI Arguments:**
* `--root`: The root directory containing the procedure documents (defaults to current directory).
* `--dry-run`: Shows what files would be copied/created without actually writing them.

## 🏗️ Built With

* [Python](https://www.python.org/) - The core language
* [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/) - For reading, copying, and modifying `.xlsx` files
* [CustomTkinter](https://customtkinter.tomschimansky.com/) - For the modern Python UI

---

<div align="center">
	<b><a href="https://devon-gs.github.io/" target="_blank">Coded by Devon - Redfire IT</a></b>
</div>

---