# 💰 ActiveContracts – Banking GUI for Collateral Loans

**ActiveContracts** is a Georgian-language GUI banking application built in Python using **PyQt5** and **SQLite**. It is designed to manage **collateral loan contracts** — where clients pawn items in exchange for money and repay interest (percent) over time.

This system is ideal for small banks, pawn shops, or private lenders who need to track **active and closed contracts**, **interest payments**, and **money flow** efficiently through a visual, database-driven interface.

---

## 🛠 Technologies Used

- **Python 3**
- **PyQt5** – for the graphical user interface
- **SQLite** – lightweight SQL database for storage
- **pandas** – for Excel export
- **QSqlTableModel / QTableView** – for data display and filtering

---

## 🌐 Language

- **Interface:** Georgian 🇬🇪
- **Codebase:** English-based (Python)
  
This allows non-technical Georgian users to use the app, while developers can still understand and maintain the code easily.

---

## 🔐 Authorization System

There are two types of users:

- **Admin**
  - Full access to all functions
  - Can edit, delete, and manage all records
- **User**
  - Restricted access
  - Can open/close contracts and manage interest and principal additions
  - Cannot perform full administrative actions

---

## 🧭 Main Structure

### 🏠 Main Page
- Entry point for navigating between contract management, registry, and money flow.

### 📑 Registry Window
- Add, edit, and remove new clients or contracts.
- Collects information like item description, given money, day quantity, and percent.

### 🔓 Active Contracts Window
- View and manage currently open contracts.
- Add additional money or interest payments.
- Extend or modify contracts as needed.

### 🔒 Closed Contracts Window
- View contracts that have been fully paid or closed.
- Archive for legal or historical records.

### 💸 Money Control Window
- Track money inflow/outflow.
- Monitor paid interests and added principal amounts.
- Export data to Excel for financial records.

### 📊 Percent Management Window
- Add, edit, and track daily interest payments.
- Visualize remaining interest or upcoming due dates.

---

## 📁 Database Structure

Multiple **SQLite** databases are used to organize data:
- `active_contracts.db` – stores all currently active loan contracts
- `contracts.db` – registry for all contracts and clients
- `adding_percent_amount.db` – tracks percent-based interest payments
- `inflow_order_both.db` and `outflow_order.db` – track money movements
- `paid_principle_registry.db` – tracks added principal amounts
- and etc.

---

## 🔄 Features

- 📅 **Date Filtering**: View records within a specific range
- ✍️ **Edit Support**: Modify any record through modal dialogs
- 📤 **Export to Excel**: Save tables and filtered data
- 🔍 **Search and Filter**: Quick lookup using text fields and dates
- 🧮 **Percent Calculations**: Auto-calculates expected interest based on given money, percent, and days
- 🛠 **Admin Tools**: Full control over all aspects of the system
- and etc.

---

## 🚀 How to Run

- ✅ You must manually populate the credentials database before running

- 🗃️ All other database files are auto-generated

---

## ⚠️ Additional Note

- 🖨️ A few features (like printing) are implemented only for Windows
