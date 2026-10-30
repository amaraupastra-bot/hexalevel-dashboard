# -*- coding: utf-8 -*-
"""
HEXALEVEL Enterprise Financial Management Web Application Backend
Features:
- Authentication system with email & password verification
- Clean naming: removed '(Co-Founder)' and 'Partner'
- User-specific Activity Log filtering
- Full automatic synchronization: adding/editing/deleting in Projects automatically syncs into Revenue Split and Scenario C.
- Partner Financial Statements tracking skips and deductions.
"""

import os
import json
from datetime import datetime
from typing import Optional
from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse, FileResponse
from pydantic import BaseModel

app = FastAPI(title="HEXALEVEL Financial Management")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "dashboard_data.json")
EXCEL_PATH = os.path.join(BASE_DIR, "Hexalevel_Financial_Management_Dashboard.xlsx")

# Registered User Accounts
USERS_DB = {
    "amara@hexalevel.com": {"name": "Amara", "passwords": ["amara123", "123456", "admin"]},
    "shaista@hexalevel.com": {"name": "Shaista", "passwords": ["shaista123", "123456", "admin"]},
    "urooj@hexalevel.com": {"name": "Urooj", "passwords": ["urooj123", "123456", "admin"]},
    "admin@hexalevel.com": {"name": "Admin", "passwords": ["admin123", "123456", "admin"]}
}

def load_data():
    if not os.path.exists(DATA_FILE):
        raise HTTPException(status_code=500, detail="Data store not found")
    with open(DATA_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
    if "activity_logs" not in data:
        data["activity_logs"] = [
            {
                "id": "ACT-001",
                "timestamp": "2026-10-30 16:15:20",
                "user": "System",
                "action": "SYSTEM_INITIALIZE",
                "tab": "Dashboard",
                "details": "Initialized HEXALEVEL Financial Management system and verified Excel reconciliation."
            }
        ]
    return data

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

def log_activity(data, user: str, action: str, tab: str, details: str):
    clean_user = user.replace("(Co-Founder)", "").replace("(Partner)", "").strip() if user else "Admin"
    log_entry = {
        "id": f"ACT-{len(data.get('activity_logs', []))+1:03d}",
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "user": clean_user or "Admin",
        "action": action,
        "tab": tab,
        "details": details
    }
    if "activity_logs" not in data:
        data["activity_logs"] = []
    data["activity_logs"].insert(0, log_entry)  # Newest first

def parse_scenario_shares(scenario_str: str, amt: float):
    if "Scenario A" in scenario_str:
        return {
            "shaista": amt * 0.25,
            "amara": amt * 0.25,
            "urooj": 0.0,
            "company": amt * 0.50,
            "skip_member": None,
            "skip_co": 0.0
        }
    elif "Scenario B" in scenario_str:
        return {
            "shaista": amt * (7/30),
            "amara": amt * (7/30),
            "urooj": amt * (7/30),
            "company": amt * 0.30,
            "skip_member": None,
            "skip_co": 0.0
        }
    else:  # Scenario C
        if "Urooj Skipped" in scenario_str or ("Shaista & Amara" in scenario_str):
            skip_member = "Urooj"
            sh = amt * (7/30)
            am = amt * (7/30)
            ur = 0.0
        elif "Amara Skipped" in scenario_str or ("Shaista & Urooj" in scenario_str):
            skip_member = "Amara"
            sh = amt * (7/30)
            am = 0.0
            ur = amt * (7/30)
        elif "Shaista Skipped" in scenario_str or ("Amara & Urooj" in scenario_str):
            skip_member = "Shaista"
            sh = 0.0
            am = amt * (7/30)
            ur = amt * (7/30)
        else:
            skip_member = "Urooj"
            sh = amt * (7/30)
            am = amt * (7/30)
            ur = 0.0

        skip_co = amt * (7/30)
        tot_co = (amt * 0.30) + skip_co
        return {
            "shaista": sh,
            "amara": am,
            "urooj": ur,
            "company": tot_co,
            "skip_member": skip_member,
            "skip_co": skip_co
        }

class LoginCredentials(BaseModel):
    email: str
    password: str

class ProjectModel(BaseModel):
    id: Optional[str] = None
    start: str
    date: str
    client: str
    project: str
    service: str
    amount: float
    currency: Optional[str] = "PKR"
    status: str
    period: str
    scenario: str
    bank: Optional[str] = ""
    notes: Optional[str] = ""
    user: Optional[str] = "Admin"

class ExpenseModel(BaseModel):
    id: Optional[str] = None
    period: str
    item: str
    link: Optional[str] = ""
    price: float
    user: Optional[str] = "Admin"

class SendMoneyModel(BaseModel):
    id: Optional[str] = None
    period: str
    recipient: str
    amount: float
    user: Optional[str] = "Admin"

# ----------------- AUTHENTICATION ENDPOINT -----------------
@app.post("/api/auth/login")
def login(creds: LoginCredentials):
    email = creds.email.strip().lower()
    user_info = USERS_DB.get(email)
    if not user_info:
        # Check by name matching (e.g. 'amara', 'shaista')
        for em, u in USERS_DB.items():
            if u["name"].lower() == email:
                user_info = u
                email = em
                break

    if not user_info or creds.password not in user_info["passwords"]:
        raise HTTPException(status_code=401, detail="Invalid email or password.")

    user_name = user_info["name"]
    data = load_data()
    log_activity(data, user_name, "USER_LOGIN", "Authentication", f"{user_name} logged into the dashboard session ({email}).")
    save_data(data)

    return {
        "status": "success",
        "user": user_name,
        "email": email
    }

@app.get("/api/data")
def get_dashboard_data(filter_user: Optional[str] = None):
    data = load_data()

    projects = data.get("projects", [])
    split_old = data.get("split_old", [])
    split_new = data.get("split_new", [])
    expenses_old = data.get("expenses_old", [])
    expenses_new = data.get("expenses_new", [])
    send_money_old = data.get("send_money_old", [])
    send_money_new = data.get("send_money_new", [])
    scenario_c = data.get("scenario_c", [])
    all_logs = data.get("activity_logs", [])

    # Filter logs if user requested specific person
    if filter_user and filter_user.strip().lower() not in ["all", ""]:
        filtered_logs = [log for log in all_logs if log.get("user", "").lower() == filter_user.strip().lower()]
    else:
        filtered_logs = all_logs

    # Old Split Totals
    old_rev = sum(x["amount"] for x in split_old if x["status"] == "Received")
    old_co = old_rev * 0.50
    old_team = old_rev * 0.50
    old_purchases = sum(e["price"] for e in expenses_old)
    old_send = sum(s["amount"] for s in send_money_old)
    old_net = old_co - old_purchases - old_send
    old_projects_count = len(split_old)
    old_received_count = sum(1 for x in split_old if x["status"] == "Received")

    # New Split Totals
    new_rev = sum(x["amount"] for x in split_new if x["status"] == "Received")
    new_co = sum(x["company"] for x in split_new if x["status"] == "Received")
    new_team = sum(x["shaista"] + x["amara"] + x["urooj"] for x in split_new if x["status"] == "Received")
    new_purchases = sum(e["price"] for e in expenses_new)
    new_send = sum(s["amount"] for s in send_money_new)
    new_net = new_co - new_purchases - new_send
    new_projects_count = len(split_new)
    new_received_count = sum(1 for x in split_new if x["status"] == "Received")

    all_rev = old_rev + new_rev
    all_net = old_net + new_net
    all_projects_count = old_projects_count + new_projects_count
    all_received_count = old_received_count + new_received_count

    # Scenario C Aggregates
    sc_total_amt = sum(c["amount"] for c in scenario_c if c["status"] == "Received")
    sc_shaista_dir = sum(c["shaista"] for c in scenario_c if c["status"] == "Received")
    sc_amara_dir = sum(c["amara"] for c in scenario_c if c["status"] == "Received")
    sc_urooj_dir = sum(c["urooj"] for c in scenario_c if c["status"] == "Received")
    sc_skip_co = sum(c["skip_to_co"] for c in scenario_c if c["status"] == "Received")
    sc_total_co = sum(c["total_co"] for c in scenario_c if c["status"] == "Received")

    # Accurate Member Calculations (No double counting between split_new and scenario_c)
    # 1. Shaista
    sh_old_direct = sum(x["team"] / 2 for x in split_old if x["status"] == "Received")
    sh_new_direct = sum(x.get("shaista", 0) for x in split_new if x["status"] == "Received")
    sh_total_direct = sh_old_direct + sh_new_direct
    # Surplus from company skips is credited ONLY to members who actually worked on that project
    sh_sc_surplus = sum(c["skip_to_co"] / 2 for c in scenario_c if c["status"] == "Received" and c.get("shaista", 0) > 0)
    sh_net = sh_total_direct + sh_sc_surplus
    sh_worked_count = len(split_old) + sum(1 for x in split_new if x.get("shaista", 0) > 0)
    sh_skipped_count = sum(1 for c in scenario_c if c.get("skip_member", "").lower() == "shaista")
    sh_skipped_deduction = sum(c["amount"] * (7/30) for c in scenario_c if c.get("skip_member", "").lower() == "shaista")

    # 2. Amara
    am_old_direct = sum(x["team"] / 2 for x in split_old if x["status"] == "Received")
    am_new_direct = sum(x.get("amara", 0) for x in split_new if x["status"] == "Received")
    am_total_direct = am_old_direct + am_new_direct
    am_sc_surplus = sum(c["skip_to_co"] / 2 for c in scenario_c if c["status"] == "Received" and c.get("amara", 0) > 0)
    am_net = am_total_direct + am_sc_surplus
    am_worked_count = len(split_old) + sum(1 for x in split_new if x.get("amara", 0) > 0)
    am_skipped_count = sum(1 for c in scenario_c if c.get("skip_member", "").lower() == "amara")
    am_skipped_deduction = sum(c["amount"] * (7/30) for c in scenario_c if c.get("skip_member", "").lower() == "amara")

    # 3. Urooj
    ur_new_direct = sum(x.get("urooj", 0) for x in split_new if x["status"] == "Received")
    ur_total_direct = ur_new_direct
    ur_sc_surplus = sum(c["skip_to_co"] / 2 for c in scenario_c if c["status"] == "Received" and c.get("urooj", 0) > 0)
    ur_net = ur_total_direct + ur_sc_surplus
    ur_worked_count = sum(1 for x in split_new if x.get("urooj", 0) > 0)
    ur_skipped_count = sum(1 for c in scenario_c if c.get("skip_member", "").lower() == "urooj")
    ur_skipped_deduction = sum(c["amount"] * (7/30) for c in scenario_c if c.get("skip_member", "").lower() == "urooj")

    partners = {
        "shaista": {
            "name": "Shaista",
            "direct": sh_total_direct,
            "surplus": sh_sc_surplus,
            "total_net": sh_net,
            "projects_worked": sh_worked_count,
            "projects_skipped": sh_skipped_count,
            "skipped_deduction": sh_skipped_deduction
        },
        "amara": {
            "name": "Amara",
            "direct": am_total_direct,
            "surplus": am_sc_surplus,
            "total_net": am_net,
            "projects_worked": am_worked_count,
            "projects_skipped": am_skipped_count,
            "skipped_deduction": am_skipped_deduction
        },
        "urooj": {
            "name": "Urooj",
            "direct": ur_total_direct,
            "surplus": ur_sc_surplus,
            "total_net": ur_net,
            "projects_worked": ur_worked_count,
            "projects_skipped": ur_skipped_count,
            "skipped_deduction": ur_skipped_deduction
        }
    }

    # Dedicated Company Financial Statement & Waterfall
    new_co_base = sum(x["amount"] * 0.30 for x in split_new if x["status"] == "Received")
    total_gross_co = old_co + new_co
    total_purchases = old_purchases + new_purchases
    total_send = old_send + new_send
    total_deductions = total_purchases + total_send
    company_net = total_gross_co - total_deductions

    company_statement = {
        "old_co_share": old_co,
        "new_co_base_share": new_co_base,
        "sc_skip_additions": sc_skip_co,
        "total_gross_inflow": total_gross_co,
        "purchases_deducted": total_purchases,
        "send_money_deducted": total_send,
        "total_deductions": total_deductions,
        "overall_net_profit": company_net,
        "total_revenue": all_rev,
        "net_profit_margin_pct": round((company_net / all_rev * 100) if all_rev > 0 else 0, 2)
    }

    chart_data = {
        "scenario_comparison": {
            "labels": ["Scenario A (2 Working)", "Scenario B (3 Working)", "Scenario C (2 Working, 1 Skipped)"],
            "company_pct": [50.0, 30.0, 53.33],
            "shaista_pct": [25.0, 23.33, 23.33],
            "amara_pct": [25.0, 23.33, 23.33],
            "urooj_pct": [0.0, 23.33, 0.0]
        },
        "partner_earnings": {
            "labels": ["Shaista", "Amara", "Urooj"],
            "direct": [sh_total_direct, am_total_direct, ur_total_direct],
            "surplus": [sh_sc_surplus, am_sc_surplus, ur_sc_surplus],
            "net": [sh_net, am_net, ur_net],
            "deductions": [sh_skipped_deduction, am_skipped_deduction, ur_skipped_deduction]
        },
        "revenue_breakdown": {
            "labels": ["Company Net Profit", "Purchases", "Send Money", "Team Payouts"],
            "values": [all_net, old_purchases + new_purchases, old_send + new_send, old_team + new_team]
        }
    }

    return {
        "dashboard": {
            "old_rev": old_rev,
            "old_co": old_co,
            "old_team": old_team,
            "old_purchases": old_purchases,
            "old_net": old_net,
            "old_projects": old_projects_count,
            "old_received": old_received_count,
            "old_send": old_send,
            "new_rev": new_rev,
            "new_co": new_co,
            "new_team": new_team,
            "new_purchases": new_purchases,
            "new_net": new_net,
            "new_projects": new_projects_count,
            "new_received": new_received_count,
            "new_send": new_send,
            "all_rev": all_rev,
            "all_net": all_net,
            "all_projects": all_projects_count,
            "all_received": all_received_count
        },
        "projects": projects,
        "split_old": split_old,
        "split_new": split_new,
        "expenses_old": expenses_old,
        "expenses_new": expenses_new,
        "send_money_old": send_money_old,
        "send_money_new": send_money_new,
        "scenario_c": scenario_c,
        "scenario_c_summary": {
            "total_amount": sc_total_amt,
            "shaista": sc_shaista_dir,
            "amara": sc_amara_dir,
            "urooj": sc_urooj_dir,
            "skip_to_co": sc_skip_co,
            "total_co": sc_total_co
        },
        "partners": partners,
        "company_statement": company_statement,
        "chart_data": chart_data,
        "activity_logs": filtered_logs,
        "total_logs_count": len(all_logs)
    }

# ----------------- PROJECTS CRUD WITH AUTO-SYNC -----------------
@app.post("/api/project")
def add_project(p: ProjectModel):
    data = load_data()
    pid = f"PRJ-{len(data['projects'])+1:03d}"
    user_name = p.user.replace("(Co-Founder)", "").replace("(Partner)", "").strip() if p.user else "Admin"

    new_proj = {
        "id": pid,
        "start": p.start,
        "client": p.client,
        "project": p.project,
        "service": p.service,
        "amount": p.amount,
        "currency": p.currency or "PKR",
        "status": p.status,
        "date": p.date,
        "notes": p.notes or "",
        "bank": p.bank or "",
        "period": p.period,
        "scenario": p.scenario
    }
    data["projects"].append(new_proj)

    amt = p.amount
    month_name = p.start or p.date[:7]
    shares = parse_scenario_shares(p.scenario, amt)

    if p.period == "old":
        data["split_old"].append({
            "id": pid,
            "month": month_name,
            "amount": amt,
            "company": shares["company"],
            "team": shares["shaista"] + shares["amara"],
            "status": p.status,
            "date": p.date,
            "shaista": p.bank or "Nayapay",
            "amara": "Askari",
            "tax": ""
        })
        log_activity(data, user_name, "ADD_PROJECT", "Projects & Revenue Split", 
                     f"{user_name} added project {pid} '{p.client}' (PKR {amt:,.2f}) into Upper 50/50 Table.")
    else:
        if "Scenario C" in p.scenario:
            skip_name = shares["skip_member"]
            sh_share = shares["shaista"]
            am_share = shares["amara"]
            ur_share = shares["urooj"]
            skip_co = shares["skip_co"]
            tot_co = shares["company"]

            data["scenario_c"].append({
                "id": pid,
                "client": f"{p.client} ({p.project})",
                "amount": amt,
                "status": p.status,
                "scenario": p.scenario,
                "shaista": sh_share,
                "amara": am_share,
                "urooj": ur_share,
                "skip_member": skip_name,
                "skip_to_co": skip_co,
                "total_co": tot_co,
                "notes": f"{skip_name} skipped -> unallocated 23.33% (PKR {skip_co:,.2f}) to Company"
            })

            data["split_new"].append({
                "id": pid,
                "month": month_name,
                "amount": amt,
                "company": tot_co,
                "shaista": sh_share,
                "amara": am_share,
                "urooj": ur_share,
                "status": p.status,
                "date": p.date,
                "shaista_acc": "Nayapay" if sh_share > 0 else "SKIP (0.00)",
                "amara_acc": "Askari" if am_share > 0 else "SKIP (0.00)",
                "urooj_acc": "UBL" if ur_share > 0 else "SKIP (0.00)",
                "tax": "",
                "notes": f"Scenario C: {skip_name} Skipped -> Co 53.33%",
                "scenario": p.scenario
            })

            log_activity(data, user_name, "ADD_PROJECT", "Projects, Revenue Split & Scenario C", 
                         f"{user_name} added Scenario C project {pid} '{p.client}' (PKR {amt:,.2f}) where {skip_name} was SKIPPED.")
        else:
            data["split_new"].append({
                "id": pid,
                "month": month_name,
                "amount": amt,
                "company": shares["company"],
                "shaista": shares["shaista"],
                "amara": shares["amara"],
                "urooj": shares["urooj"],
                "status": p.status,
                "date": p.date,
                "shaista_acc": "Nayapay",
                "amara_acc": "Askari",
                "urooj_acc": "UBL",
                "tax": "",
                "notes": p.notes or "",
                "scenario": p.scenario
            })
            log_activity(data, user_name, "ADD_PROJECT", "Projects & Revenue Split", 
                         f"{user_name} added project {pid} '{p.client}' (PKR {amt:,.2f}) into Lower 30/70 Table.")

    save_data(data)
    return {"status": "success", "id": pid}

@app.put("/api/project/{pid}")
def edit_project(pid: str, p: ProjectModel):
    data = load_data()
    idx = next((i for i, item in enumerate(data["projects"]) if item["id"] == pid), None)
    if idx is None:
        raise HTTPException(status_code=404, detail="Project not found")

    user_name = p.user.replace("(Co-Founder)", "").replace("(Partner)", "").strip() if p.user else "Admin"

    data["projects"][idx] = {
        "id": pid,
        "start": p.start,
        "client": p.client,
        "project": p.project,
        "service": p.service,
        "amount": p.amount,
        "currency": p.currency or "PKR",
        "status": p.status,
        "date": p.date,
        "notes": p.notes or "",
        "bank": p.bank or "",
        "period": p.period,
        "scenario": p.scenario
    }

    # Automatically sync update in Revenue Split & Scenario C
    data["split_old"] = [x for x in data["split_old"] if x.get("id") != pid]
    data["split_new"] = [x for x in data["split_new"] if x.get("id") != pid]
    data["scenario_c"] = [x for x in data["scenario_c"] if x.get("id") != pid]

    amt = p.amount
    month_name = p.start or p.date[:7]
    shares = parse_scenario_shares(p.scenario, amt)

    if p.period == "old":
        data["split_old"].append({
            "id": pid,
            "month": month_name,
            "amount": amt,
            "company": shares["company"],
            "team": shares["shaista"] + shares["amara"],
            "status": p.status,
            "date": p.date,
            "shaista": p.bank or "Nayapay",
            "amara": "Askari",
            "tax": ""
        })
    else:
        if "Scenario C" in p.scenario:
            skip_name = shares["skip_member"]
            sh_share = shares["shaista"]
            am_share = shares["amara"]
            ur_share = shares["urooj"]
            skip_co = shares["skip_co"]
            tot_co = shares["company"]

            data["scenario_c"].append({
                "id": pid,
                "client": f"{p.client} ({p.project})",
                "amount": amt,
                "status": p.status,
                "scenario": p.scenario,
                "shaista": sh_share,
                "amara": am_share,
                "urooj": ur_share,
                "skip_member": skip_name,
                "skip_to_co": skip_co,
                "total_co": tot_co,
                "notes": f"{skip_name} skipped -> unallocated 23.33% to Co"
            })
            data["split_new"].append({
                "id": pid,
                "month": month_name,
                "amount": amt,
                "company": tot_co,
                "shaista": sh_share,
                "amara": am_share,
                "urooj": ur_share,
                "status": p.status,
                "date": p.date,
                "shaista_acc": "Nayapay" if sh_share > 0 else "SKIP (0.00)",
                "amara_acc": "Askari" if am_share > 0 else "SKIP (0.00)",
                "urooj_acc": "UBL" if ur_share > 0 else "SKIP (0.00)",
                "tax": "",
                "notes": f"Scenario C: {skip_name} Skipped -> Co 53.33%",
                "scenario": p.scenario
            })
        else:
            data["split_new"].append({
                "id": pid,
                "month": month_name,
                "amount": amt,
                "company": shares["company"],
                "shaista": shares["shaista"],
                "amara": shares["amara"],
                "urooj": shares["urooj"],
                "status": p.status,
                "date": p.date,
                "shaista_acc": "Nayapay",
                "amara_acc": "Askari",
                "urooj_acc": "UBL",
                "tax": "",
                "notes": p.notes or "",
                "scenario": p.scenario
            })

    log_activity(data, user_name, "EDIT_PROJECT", "Projects & Revenue Split", 
                 f"{user_name} edited project {pid} '{p.client}' (PKR {amt:,.2f}).")
    save_data(data)
    return {"status": "success"}

@app.delete("/api/project/{pid}")
def delete_project(pid: str, user: Optional[str] = "Admin"):
    data = load_data()
    user_name = user.replace("(Co-Founder)", "").replace("(Partner)", "").strip() if user else "Admin"
    proj_to_delete = next((p for p in data["projects"] if p["id"] == pid), None)
    client_name = proj_to_delete["client"] if proj_to_delete else pid

    data["projects"] = [p for p in data["projects"] if p["id"] != pid]
    data["split_old"] = [x for x in data["split_old"] if x.get("id") != pid]
    data["split_new"] = [x for x in data["split_new"] if x.get("id") != pid]
    data["scenario_c"] = [c for c in data["scenario_c"] if c.get("id") != pid]

    log_activity(data, user_name, "DELETE_PROJECT", "Projects & Revenue Split", 
                 f"{user_name} deleted project {pid} '{client_name}' and synchronized records.")
    save_data(data)
    return {"status": "success"}

# ----------------- PURCHASES / EXPENSES CRUD -----------------
@app.post("/api/expense")
def add_expense(e: ExpenseModel):
    data = load_data()
    user_name = e.user.replace("(Co-Founder)", "").replace("(Partner)", "").strip() if e.user else "Admin"
    if e.period == "old":
        eid = f"EXP-OLD-{len(data['expenses_old'])+1:03d}"
        data["expenses_old"].append({"id": eid, "item": e.item, "link": e.link or "", "price": e.price})
        tab_name = "Purchases (Upper)"
    else:
        eid = f"EXP-NEW-{len(data['expenses_new'])+1:03d}"
        data["expenses_new"].append({"id": eid, "item": e.item, "link": e.link or "", "price": e.price})
        tab_name = "Purchases (Lower)"

    log_activity(data, user_name, "ADD_PURCHASE", tab_name, 
                 f"{user_name} added purchase '{e.item}' (PKR {e.price:,.2f}).")
    save_data(data)
    return {"status": "success", "id": eid}

@app.put("/api/expense/{eid}")
def edit_expense(eid: str, e: ExpenseModel):
    data = load_data()
    user_name = e.user.replace("(Co-Founder)", "").replace("(Partner)", "").strip() if e.user else "Admin"
    for idx, item in enumerate(data["expenses_old"]):
        if item["id"] == eid:
            data["expenses_old"][idx] = {"id": eid, "item": e.item, "link": e.link or "", "price": e.price}
            log_activity(data, user_name, "EDIT_PURCHASE", "Purchases (Upper)", f"{user_name} edited purchase {eid} '{e.item}'.")
            save_data(data)
            return {"status": "success"}
    for idx, item in enumerate(data["expenses_new"]):
        if item["id"] == eid:
            data["expenses_new"][idx] = {"id": eid, "item": e.item, "link": e.link or "", "price": e.price}
            log_activity(data, user_name, "EDIT_PURCHASE", "Purchases (Lower)", f"{user_name} edited purchase {eid} '{e.item}'.")
            save_data(data)
            return {"status": "success"}
    raise HTTPException(status_code=404, detail="Expense not found")

@app.delete("/api/expense/{eid}")
def delete_expense(eid: str, user: Optional[str] = "Admin"):
    data = load_data()
    user_name = user.replace("(Co-Founder)", "").replace("(Partner)", "").strip() if user else "Admin"
    data["expenses_old"] = [x for x in data["expenses_old"] if x["id"] != eid]
    data["expenses_new"] = [x for x in data["expenses_new"] if x["id"] != eid]
    log_activity(data, user_name, "DELETE_PURCHASE", "Purchases", f"{user_name} deleted purchase {eid}.")
    save_data(data)
    return {"status": "success"}

# ----------------- SEND MONEY CRUD -----------------
@app.post("/api/sendmoney")
def add_send_money(s: SendMoneyModel):
    data = load_data()
    user_name = s.user.replace("(Co-Founder)", "").replace("(Partner)", "").strip() if s.user else "Admin"
    if s.period == "old":
        sid = f"SEND-OLD-{len(data['send_money_old'])+1:03d}"
        data["send_money_old"].append({"id": sid, "recipient": s.recipient, "amount": s.amount})
        tab_name = "Send Money (Upper)"
    else:
        sid = f"SEND-NEW-{len(data['send_money_new'])+1:03d}"
        data["send_money_new"].append({"id": sid, "recipient": s.recipient, "amount": s.amount})
        tab_name = "Send Money (Lower)"

    log_activity(data, user_name, "ADD_SEND_MONEY", tab_name, 
                 f"{user_name} sent money to '{s.recipient}' (PKR {s.amount:,.2f}).")
    save_data(data)
    return {"status": "success", "id": sid}

@app.put("/api/sendmoney/{sid}")
def edit_send_money(sid: str, s: SendMoneyModel):
    data = load_data()
    user_name = s.user.replace("(Co-Founder)", "").replace("(Partner)", "").strip() if s.user else "Admin"
    for idx, item in enumerate(data["send_money_old"]):
        if item["id"] == sid:
            data["send_money_old"][idx] = {"id": sid, "recipient": s.recipient, "amount": s.amount}
            log_activity(data, user_name, "EDIT_SEND_MONEY", "Send Money", f"{user_name} edited payment to '{s.recipient}'.")
            save_data(data)
            return {"status": "success"}
    for idx, item in enumerate(data["send_money_new"]):
        if item["id"] == sid:
            data["send_money_new"][idx] = {"id": sid, "recipient": s.recipient, "amount": s.amount}
            log_activity(data, user_name, "EDIT_SEND_MONEY", "Send Money", f"{user_name} edited payment to '{s.recipient}'.")
            save_data(data)
            return {"status": "success"}
    raise HTTPException(status_code=404, detail="Record not found")

@app.delete("/api/sendmoney/{sid}")
def delete_send_money(sid: str, user: Optional[str] = "Admin"):
    data = load_data()
    user_name = user.replace("(Co-Founder)", "").replace("(Partner)", "").strip() if user else "Admin"
    data["send_money_old"] = [x for x in data["send_money_old"] if x["id"] != sid]
    data["send_money_new"] = [x for x in data["send_money_new"] if x["id"] != sid]
    log_activity(data, user_name, "DELETE_SEND_MONEY", "Send Money", f"{user_name} deleted payment {sid}.")
    save_data(data)
    return {"status": "success"}

@app.get("/download-excel")
def download_excel():
    if not os.path.exists(EXCEL_PATH):
        raise HTTPException(status_code=404, detail="Excel file not found")
    return FileResponse(
        EXCEL_PATH,
        filename="Hexalevel_Financial_Management_Dashboard.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.get("/", response_class=HTMLResponse)
def index_page():
    index_path = os.path.join(BASE_DIR, "index.html")
    with open(index_path, "r", encoding="utf-8") as f:
        return f.read()

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
