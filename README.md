# HEXALEVEL — Financial Management & Revenue Distribution Dashboard

A full-featured, automated Financial Management & Revenue Distribution Dashboard for **HEXALEVEL** built with FastAPI, HTML5, Chart.js, and OpenPyXL.

---

## 🏛️ System Features

1. **Exact Excel Reconciliation & Automation**:
   - Replicates and auto-syncs historical 50/50 model (pre-Sep 2026) and new 30/70 model (from Sep 2026 onwards).
   - Generates fully calculated Excel workbooks with COM formula evaluation (`Hexalevel_Financial_Management_Dashboard.xlsx`).

2. **Scenario C Skip Rule & Company Allocation**:
   - Handles multi-member participation rules where non-working members are skipped.
   - Automatically re-routes the unallocated 23.33% share to the **Company** (raising company share to 53.33%).
   - Direct payouts for skipped members are set to `SKIP (PKR 0.00)`.
   - The company surplus is distributed 50/50 **strictly to active working members**; skipped members receive PKR 0.00 surplus.

3. **Company Net Profit Waterfall**:
   - Comprehensive reconciliation tracking Gross Inflows (Base 50% + Base 30% + Skipped Shares).
   - Deductions tracking for Purchases & Operating Expenses and Send Money transfers.
   - Real-time Company Net Profit and profit margin calculation.

4. **Multi-User Role Authentication & Audit Trail**:
   - Authentication system with email and password sign-in for team members.
   - Individual member session tracking.
   - Comprehensive Activity & Audit Log recording user actions, timestamps, and target tabs.
   - **Admin Super-User**: Can view and filter activity logs for any team member or inspect the entire company audit trail. Regular members default to their own activity.

5. **Full Dynamic CRUD Operations**:
   - Add, Edit, and Delete client projects with automatic multi-table and statement synchronization.
   - Add and track Purchases / Operating Expenses.
   - Add and track Send Money Outflows.
   - One-click Excel `.xlsx` download.

---

## 🚀 Quick Start Guide

### 1. Requirements
- Python 3.10+
- Dependencies:
  ```bash
  pip install fastapi uvicorn openpyxl
  ```

### 2. Run the Dashboard
```bash
python app.py
```
Open your browser and navigate to:
```
http://localhost:8000
```

---

## 👥 User Accounts & Credentials

| User | Email | Default Password | Role |
| :--- | :--- | :--- | :--- |
| **Amara** | `amara@hexalevel.com` | `123456` | Member |
| **Shaista** | `shaista@hexalevel.com` | `123456` | Member |
| **Urooj** | `urooj@hexalevel.com` | `123456` | Member |
| **Admin** | `admin@hexalevel.com` | `123456` | Super Administrator |

---

## 📁 Project Structure

```
├── app.py                                        # FastAPI Backend Server & Calculation Engine
├── index.html                                    # Single Page Dashboard Application
├── dashboard_data.json                           # Persistent Financial Database
├── Hexalevel_Financial_Management_Dashboard.xlsx # Automated Excel Workbook
├── generate_matching_excel.py                    # Excel Generator Script
├── .gitignore                                    # Git Ignore Rules
└── README.md                                     # Documentation
```
