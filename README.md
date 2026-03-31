# 🚀 Stock Portfolio Management (Java + Excel)

A lightweight desktop application for managing stock transactions, built with **Java Swing** and **Apache POI**. The app helps track buy/sell activities, calculate portfolio metrics, and store everything in Excel.

---

## ✨ Features

* Manage multiple customers and their stock portfolios
* Record **buy (MUA)** and **sell (BÁN)** transactions
* Automatically calculate:

  * Average price (Giá TB)
  * Remaining quantity (SL còn lại)
  * Profit / Loss
  * Total transaction fees
* Excel-based storage (no database required)
* Highlight the **latest active stock row** (only one row per stock when quantity > 0)

---

## 🛠 Tech Stack

* **Java (Swing)** – UI
* **Apache POI** – Excel processing
* **Maven** – Dependency management

---

## 📂 Project Structure

```
src/
 └── main/java/org/example/
      ├── StartFrame.java
      ├── QuanLyKhachHangApp.java
      └── ...
template_giaodich.xlsx
QuanLyChungKhoan.xlsx
```

---

## ⚙️ Setup & Run

### 1. Clone project

```bash
git clone https://github.com/your-username/stock-portfolio-management.git
cd stock-portfolio-management
```

### 2. Open with IDE

* IntelliJ IDEA / Eclipse
* Make sure Maven dependencies are loaded

### 3. Run application

Run:

```
StartFrame.java
```

---

## 📌 Notes

* Excel file will be auto-created if not found
* Each customer has a separate transaction sheet
* Only the latest row of each stock is highlighted when still holding

---

## 💡 Future Improvements

* MySQL database integration
* REST API + Web frontend
* Export reports (PDF, charts)
* Real-time stock price integration

---

## 👤 Author

* Thương

---
