import streamlit as st
import pandas as pd
import joblib
import os
import re
import random
import requests
from datetime import datetime, timedelta
import openpyxl

# Page Configuration
st.set_page_config(page_title="KFH Dashboard", layout="wide")

# --- MULTI-TENANT ACCOUNTS SYSTEM ---
if "user_accounts" not in st.session_state:
    st.session_state["user_accounts"] = {
        "admin@kfh.com": {"password": "Admin@KFH2026", "role": "SUPER_ADMIN", "gym_name": "KFH", "mobile": "9999999999", "data_file": "gym_data.xlsx.xlsx", "daily_acc_file": "Daily Accounts 2024-2026 (1).xlsx"},
        "kfhgym@kfh.com": {"password": "gym@2026Secure", "role": "GYM_OWNER", "gym_name": "KFH", "mobile": "9876543210", "data_file": "gym_data.xlsx.xlsx", "daily_acc_file": "Daily Accounts 2024-2026 (1).xlsx"},
        "gym2@kfh.com": {"password": "Gym2Password", "role": "GYM_OWNER", "gym_name": "Gym 2", "mobile": "9123456789", "data_file": "data_gym2_at_kfh_com.xlsx", "daily_acc_file": "daily_acc_gym2.xlsx"}
    }

if "otp_store" not in st.session_state:
    st.session_state["otp_store"] = {}

if "attendance_log" not in st.session_state:
    st.session_state["attendance_log"] = []

if "leads_data" not in st.session_state:
    st.session_state["leads_data"] = []

# --- HELPER TO GET OR INITIALIZE DAILY ACCOUNTS EXCEL FILE ---
def get_or_create_daily_acc_excel(user_email, default_file="Daily Accounts 2024-2026 (1).xlsx"):
    account = st.session_state["user_accounts"].get(user_email)
    acc_filename = account.get("daily_acc_file") if account else f"daily_acc_{user_email.replace('@','_at_').replace('.','_')}.xlsx"
    if not os.path.exists(acc_filename):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Oct 2026"
        headers = ['Date', 'Particular', 'Collection', 'Cash', 'Axis A/C', 'Pers A/C', 'Check', 'Bank Locker', 'Payal Cash', 'Kartikey Sir', 'Ajit Sir', 'Expense', 'Super Bazaar']
        ws.append([datetime.now().strftime('%Y-%m-%d')] + ['']*(len(headers)-1))
        ws.append(headers)
        wb.save(acc_filename)
    return acc_filename

def get_or_create_gym_excel(user_email, default_file="gym_data.xlsx.xlsx"):
    account = st.session_state["user_accounts"].get(user_email)
    if account and account["role"] == "SUPER_ADMIN":
        target_file = st.session_state.get("admin_selected_file", default_file)
        if os.path.exists(target_file):
            return target_file
        return default_file

    data_filename = account.get("data_file") if account else f"data_{user_email.replace('@','_at_').replace('.','_')}.xlsx"
    if not os.path.exists(data_filename):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Oct 2026"
        ws.append(['SR.NO', 'Date', 'Name', 'Contact no.', 'Facility', 'Package', 'Fees', 'Paid', 'Exp. Date', 'PT trainer', 'New/Renew', 'Mode of payment', 'Referene'])
        wb.save(data_filename)
    return data_filename

# --- MULTI-USER LOGIN SYSTEM ---
def multi_user_login():
    if "user_authenticated" not in st.session_state:
        st.session_state["user_authenticated"] = False
    if "view_mode" not in st.session_state:
        st.session_state["view_mode"] = "LOGIN"

    if not st.session_state["user_authenticated"]:
        st.title("🔒 KFH Portal Login")
        if st.session_state["view_mode"] == "LOGIN":
            st.markdown("### 🔑 Welcome! Login to KFH")
            email = st.text_input("Username / Email ID / Mobile:")
            pwd = st.text_input("Password:", type="password")
            col_l1, col_l2 = st.columns([1, 1])
            with col_l1:
                if st.button("🚀 Login"):
                    accounts = st.session_state["user_accounts"]
                    matched_acc, matched_email = None, None
                    for k, v in accounts.items():
                        if (k == email or v.get("mobile") == email) and v["password"] == pwd:
                            matched_acc, matched_email = v, k
                            break
                    if matched_acc:
                        st.session_state["user_authenticated"] = True
                        st.session_state["current_user"] = matched_acc
                        st.session_state["user_email"] = matched_email
                        st.rerun()
                    else:
                        st.error("❌ Invalid Username/Mobile or Password!")
            with col_l2:
                if st.button("🔑 Forgot Password / Reset via Mobile OTP"):
                    st.session_state["view_mode"] = "FORGOT_PWD"
                    st.rerun()
        return False
    return True

if not multi_user_login():
    st.stop()

# Current User Information
current_user = st.session_state["current_user"]
user_email = st.session_state["user_email"]
active_file_path = get_or_create_gym_excel(user_email)
daily_acc_file_path = get_or_create_daily_acc_excel(user_email)

# --- CHRONOLOGICAL MULTI-SHEET EXCEL LOADER ---
@st.cache_data(ttl=1)
def load_all_gym_data(file_path):
    if not os.path.exists(file_path):
        empty_df = pd.DataFrame(columns=['SR.NO', 'Date', 'Name', 'Contact no.', 'Facility', 'Package', 'Fees', 'Paid', 'Exp. Date', 'PT trainer', 'New/Renew', 'Mode of payment', 'Referene', 'Month_Sheet'])
        empty_df['Parsed_Exp_Date'] = pd.to_datetime([])
        empty_df['Parsed_Admission_Date'] = pd.to_datetime([])
        return empty_df, ["Oct 2026"]

    xls = pd.ExcelFile(file_path)
    month_order = {'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6, 'jul': 7, 'july': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12}
    
    def get_sheet_sort_key(sheet_name):
        s_lower = sheet_name.lower()
        m_val = 99
        for m, val in month_order.items():
            if m in s_lower:
                m_val = val
                break
        year_match = re.search(r'\d{4}', sheet_name)
        y_val = int(year_match.group(0)) if year_match else 2026
        return (y_val, m_val)

    sorted_sheets = sorted(xls.sheet_names, key=get_sheet_sort_key)
    all_sheets = []
    
    for sheet in sorted_sheets:
        sheet_df = pd.read_excel(file_path, sheet_name=sheet)
        if not sheet_df.empty:
            if len(sheet_df) > 0 and 'Name' not in sheet_df.columns:
                headers = sheet_df.iloc[0].values
                data_df = sheet_df.iloc[1:].copy()
                data_df.columns = [str(h).strip() for h in headers]
            else:
                data_df = sheet_df.copy()
            data_df = data_df[data_df['Name'].notna() | data_df['SR.NO'].notna()]
            data_df['Month_Sheet'] = sheet
            all_sheets.append(data_df)
            
    combined = pd.concat(all_sheets, ignore_index=True) if all_sheets else pd.DataFrame(columns=['SR.NO', 'Date', 'Name', 'Contact no.', 'Facility', 'Package', 'Fees', 'Paid', 'Exp. Date', 'PT trainer', 'New/Renew', 'Mode of payment', 'Referene', 'Month_Sheet'])
    combined['Parsed_Exp_Date'] = pd.to_datetime(combined['Exp. Date'], errors='coerce') if 'Exp. Date' in combined.columns else pd.to_datetime([])
    combined['Parsed_Admission_Date'] = pd.to_datetime(combined['Date'], errors='coerce') if 'Date' in combined.columns else pd.to_datetime([])
    return combined, sorted_sheets

# --- LOADER FOR DAILY ACCOUNTS FILE ---
@st.cache_data(ttl=1)
def load_daily_accounts(file_path):
    if not os.path.exists(file_path):
        return pd.DataFrame(), ["Oct 2026"]
    xls = pd.ExcelFile(file_path)
    all_acc_sheets = []
    for sheet in xls.sheet_names:
        sheet_df = pd.read_excel(file_path, sheet_name=sheet)
        if not sheet_df.empty:
            if 'Particular' not in sheet_df.columns and len(sheet_df) > 0:
                sheet_df.columns = [str(c).strip() for c in sheet_df.iloc[0].values]
                sheet_df = sheet_df.iloc[1:].copy()
            sheet_df['Sheet'] = sheet
            all_acc_sheets.append(sheet_df)
    combined_acc = pd.concat(all_acc_sheets, ignore_index=True) if all_acc_sheets else pd.DataFrame()
    return combined_acc, xls.sheet_names

# --- APPEND MEMBER TO EXCEL REGISTER ---
def append_member_to_excel(file_path, member_dict, target_sheet="Oct 2026"):
    try:
        wb = openpyxl.load_workbook(file_path) if os.path.exists(file_path) else openpyxl.Workbook()
        ws = wb[target_sheet] if target_sheet in wb.sheetnames else wb.create_sheet(title=target_sheet)
        headers = ['SR.NO', 'Date', 'Name', 'Contact no.', 'Facility', 'Package', 'Fees', 'Paid', 'Exp. Date', 'PT trainer', 'New/Renew', 'Mode of payment', 'Referene']
        if len(list(ws.rows)) == 0:
            ws.append(headers)
        new_row = [
            member_dict.get('SR.NO', ''), member_dict.get('Date', ''), member_dict.get('Name', ''),
            member_dict.get('Contact no.', ''), member_dict.get('Facility', ''), member_dict.get('Package', ''),
            member_dict.get('Fees', 0), member_dict.get('Paid', 0), member_dict.get('Exp. Date', ''),
            member_dict.get('PT trainer', ''), member_dict.get('New/Renew', 'New'),
            member_dict.get('Mode of payment', 'cash'), member_dict.get('Referene', 'Direct Entry')
        ]
        ws.append(new_row)
        wb.save(file_path)
        return True, "Data Saved Successfully!"
    except Exception as e:
        return False, f"Save Error: {e}"

# --- APPEND DAILY ACCOUNT ENTRY TO EXCEL ---
def append_daily_account_entry(file_path, entry_dict, target_sheet="Oct 2026"):
    try:
        wb = openpyxl.load_workbook(file_path) if os.path.exists(file_path) else openpyxl.Workbook()
        ws = wb[target_sheet] if target_sheet in wb.sheetnames else wb.create_sheet(title=target_sheet)
        headers = ['Date', 'Particular', 'Collection', 'Cash', 'Axis A/C', 'Pers A/C', 'Check', 'Bank Locker', 'Payal Cash', 'Kartikey Sir', 'Ajit Sir', 'Expense', 'Super Bazaar']
        if len(list(ws.rows)) == 0:
            ws.append([datetime.now().strftime('%Y-%m-%d')] + ['']*(len(headers)-1))
            ws.append(headers)
        new_row = [entry_dict.get(h, '') for h in headers]
        ws.append(new_row)
        wb.save(file_path)
        return True, "Daily Accounts Entry Saved Successfully!"
    except Exception as e:
        return False, f"Save Error: {e}"

# --- SIDEBAR USER INFO & LOGOUT ---
st.sidebar.markdown(f"### 👤 KFH")
st.sidebar.info(f"**Gym:** {current_user['gym_name']}\n\n**Email:** {user_email}\n\n**Role:** `{current_user['role']}`")

if st.sidebar.button("🚪 Logout"):
    st.session_state["user_authenticated"] = False
    st.rerun()

# --- LANGUAGE SELECTION ---
st.sidebar.title("🌍 Language / भाषा")
lang = st.sidebar.selectbox("Select Language:", ["English", "हिंदी", "मराठी"])

t = {
    "English": {
        "title": f"🏋️‍♂️ KFH",
        "subtitle": "DASHBOARD",
        "records": "📊 Member Admission Records",
        "add_member": "➕ New Member Admission Form",
        "attendance": "⏱️ Member Check-in & Attendance",
        "leads": "🎯 Lead & Inquiry Manager",
        "accounts": "💰 Accounts & Financial Ledger",
        "daily_ledger": "📑 Daily Accounting Register",
        "security": "🔒 Account Security",
    },
    "हिंदी": {
        "title": f"🏋️‍♂️ KFH",
        "subtitle": "डैशबोर्ड",
        "records": "📊 सदस्य एडमिशन रिकॉर्ड्स",
        "add_member": "➕ नया सदस्य एडमिशन फॉर्म",
        "attendance": "⏱️ सदस्य उपस्थिति और चेक-इन",
        "leads": "🎯 लीड और पूछताछ मैनेजर",
        "accounts": "💰 अकाउंट्स और फाइनेंस",
        "daily_ledger": "📑 डेली अकाउंटिंग रजिस्टर",
        "security": "🔒 अकाउंट सिक्योरिटी",
    },
    "मराठी": {
        "title": f"🏋️‍♂️ KFH",
        "subtitle": "डॅशबोर्ड",
        "records": "📊 सदस्य ॲडमिशन रेकॉर्ड्स",
        "add_member": "➕ नवीन सदस्य नोंदणी फॉर्म",
        "attendance": "⏱️ सदस्य हजेरी आणि चेक-इन",
        "leads": "🎯 लीड आणि चौकशी मॅनेजर",
        "accounts": "💰 अकाउंट्स आणि फायनान्स",
        "daily_ledger": "📑 डेली अकाउंटिंग रजिस्टर",
        "security": "🔒 खाते सुरक्षा",
    }
}

st.title(t[lang]["title"])
st.markdown(t[lang]["subtitle"])

df, loaded_sheets = load_all_gym_data(active_file_path)
df_str = df.astype(str)
df_daily_acc, daily_sheets = load_daily_accounts(daily_acc_file_path)

cash_sr = [int(re.search(r'\d+', str(s)).group(0)) for s in df['SR.NO'].dropna() if str(s).strip().startswith('C') and re.search(r'\d+', str(s))] if not df.empty and 'SR.NO' in df.columns else []
online_sr = [int(re.search(r'\d+', str(s)).group(0)) for s in df['SR.NO'].dropna() if str(s).strip().startswith('O') and re.search(r'\d+', str(s))] if not df.empty and 'SR.NO' in df.columns else []

next_cash_sr = f"C {max(cash_sr) + 1}" if cash_sr else "C 101"
next_online_sr = f"O{max(online_sr) + 1}" if online_sr else "O101"

# --- TABS ---
tabs_list = [
    "📅 Expiry & Renewals Overview", 
    "📊 Master Records (Jan-Dec)",
    f"{t[lang]['add_member']}", 
    f"{t[lang]['attendance']}", 
    f"{t[lang]['leads']}",
    f"{t[lang]['accounts']}",
    f"{t[lang]['daily_ledger']}",
    f"{t[lang]['security']}"
]

if current_user["role"] == "SUPER_ADMIN":
    tabs_list.append("👑 Super Admin Panel")

tab_objects = st.tabs(tabs_list)

tab_expiry = tab_objects[0]
tab_master = tab_objects[1]
tab_add = tab_objects[2]
tab_att = tab_objects[3]
tab_leads = tab_objects[4]
tab_accounts = tab_objects[5]
tab_daily_ledger = tab_objects[6]
tab_security = tab_objects[7]

# MODULE 0: DYNAMIC EXPIRY & RENEWALS OVERVIEW
with tab_expiry:
    st.markdown("### 📅 Expiry & Renewals Overview (2026 - 2027)")
    
    col_e1, col_e2, col_e3 = st.columns(3)
    now = datetime.now()
    
    with col_e1:
        sel_exp_month = st.selectbox("Select Month:", list(range(1, 13)), index=now.month-1, format_func=lambda x: datetime(2026, x, 1).strftime('%B'), key="exp_month_sel")
    with col_e2:
        sel_exp_year = st.selectbox("Select Year:", [2026, 2027], index=0, key="exp_year_sel")
    with col_e3:
        exp_search = st.text_input("🔍 Search Expiry/Renewals:", placeholder="Name or Phone...", key="exp_search_bar")

    if not df.empty and 'Parsed_Exp_Date' in df.columns:
        exp_filtered = df[(df['Parsed_Exp_Date'].dt.month == sel_exp_month) & (df['Parsed_Exp_Date'].dt.year == sel_exp_year)].copy()
        
        if 'Parsed_Admission_Date' in df.columns:
            adm_filtered = df[(df['Parsed_Admission_Date'].dt.month == sel_exp_month) & (df['Parsed_Admission_Date'].dt.year == sel_exp_year)].copy()
        else:
            adm_filtered = pd.DataFrame()

        if exp_search:
            exp_filtered = exp_filtered[exp_filtered.astype(str).apply(lambda row: row.str.contains(exp_search, case=False).any(), axis=1)]
            adm_filtered = adm_filtered[adm_filtered.astype(str).apply(lambda row: row.str.contains(exp_search, case=False).any(), axis=1)]

        st.markdown(f"#### 🔴 Members Expiring in {datetime(2026, sel_exp_month, 1).strftime('%B')} {sel_exp_year} (Total: {len(exp_filtered)})")
        disp_cols = [c for c in ['SR.NO', 'Date', 'Name', 'Contact no.', 'Facility', 'Package', 'Fees', 'Paid', 'Exp. Date', 'New/Renew', 'Month_Sheet'] if c in exp_filtered.columns]
        if not exp_filtered.empty:
            st.dataframe(exp_filtered[disp_cols], use_container_width=True)
        else:
            st.info("Is mahine koi membership expire nahi ho rahi hai.")

        st.markdown("---")
        renewed_cnt = len(adm_filtered[adm_filtered['New/Renew'].astype(str).str.contains('Renew', case=False, na=False)]) if not adm_filtered.empty else 0
        new_cnt = len(adm_filtered[adm_filtered['New/Renew'].astype(str).str.contains('New', case=False, na=False)]) if not adm_filtered.empty else 0
        
        st.markdown(f"#### 🔄 Renewals & ✨ New Admissions ({datetime(2026, sel_exp_month, 1).strftime('%B')} {sel_exp_year})")
        m_met1, m_met2 = st.columns(2)
        with m_met1:
            st.metric(label="🔄 Renewed Members", value=renewed_cnt)
        with m_met2:
            st.metric(label="✨ New Admissions", value=new_cnt)

        if not adm_filtered.empty:
            st.dataframe(adm_filtered[disp_cols], use_container_width=True)
        else:
            st.info("Is mahine koi naya admission ya renewal record nahi hai.")
    else:
        st.info("No expiry records found in database.")

# MODULE 1: MASTER RECORDS (JAN TO DEC FULL DATA)
with tab_master:
    st.markdown("### 📊 Master Admission Records (Jan 2026 - Dec 2027)")
    global_search = st.text_input("🔍 Global Search across all months (Name, Phone, ID):", placeholder="Type to search...")
    if not df.empty:
        master_df = df[df.astype(str).apply(lambda row: row.str.contains(global_search, case=False).any(), axis=1)] if global_search else df
        st.write(f"Total Records Displayed: **{len(master_df)}**")
        st.dataframe(master_df, use_container_width=True)
    else:
        st.info("No master records available.")

# MODULE 2: MEMBER ADMISSION FORM
with tab_add:
    st.markdown("### 📝 New Member Admission Form")
    pay_mode_input = st.selectbox("Mode of payment:*", ["cash", "online", "gpay", "phonepe"], key="pay_mode_select_form")
    suggested_sr = next_cash_sr if pay_mode_input == "cash" else next_online_sr

    with st.form("add_member_form_working"):
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            sr_no = st.text_input("Receipt No. / SR.NO:*", value=suggested_sr)
            admission_date = st.date_input("Admission Date (Date):*", value=datetime.now())
            m_name = st.text_input("Member Name:*")
            m_contact = st.text_input("Contact no. (Mobile Number):*")
        with col_f2:
            facility = st.selectbox("Facility:*", ["GT", "PT", "Couple GT", "Steam Bath"])
            pkg_choice = st.selectbox("Package:*", ["1 Month", "2 Month", "3 Month", "6 Month", "1 Year", "1 Day", "15 Days"])
            reg_type = st.selectbox("New / Renew:*", ["New", "Renew", "Couple"])
            pt_trainer = st.text_input("PT trainer / Remarks:", value="")
        with col_f3:
            fees = st.number_input("Fees (₹):*", value=3000, step=500)
            paid = st.number_input("Paid (₹):*", value=3000, step=500)
            target_sheet_choice = st.selectbox("Register Sheet Month:*", loaded_sheets if loaded_sheets else ["Oct 2026"], index=len(loaded_sheets)-1 if loaded_sheets else 0)
            reference = st.text_input("Referene / Notes:", value="KFH Direct Admission")

        exp_date_calc = admission_date + timedelta(days=365 if "1 Year" in pkg_choice else (180 if "6 Month" in pkg_choice else (90 if "3 Month" in pkg_choice else (60 if "2 Month" in pkg_choice else (15 if "15 Days" in pkg_choice else 30)))))
        st.info(f"💡 **Auto Exp. Date:** `{exp_date_calc.strftime('%Y-%m-%d')}` | **Pending Balance:** `₹{fees - paid}`")
        submit_btn = st.form_submit_button("✅ Save Member Entry Permanently")

    if submit_btn:
        if not m_name or not m_contact:
            st.error("❌ Name aur Contact Number enter karna compulsory hai!")
        else:
            member_record = {
                'SR.NO': sr_no, 'Date': admission_date.strftime('%Y-%m-%d'), 'Name': m_name,
                'Contact no.': m_contact, 'Facility': facility, 'Package': pkg_choice,
                'Fees': fees, 'Paid': paid, 'Exp. Date': exp_date_calc.strftime('%Y-%m-%d'),
                'PT trainer': pt_trainer, 'New/Renew': reg_type, 'Mode of payment': pay_mode_input, 'Referene': reference
            }
            success, msg = append_member_to_excel(active_file_path, member_record, target_sheet=target_sheet_choice)
            if success:
                st.success(f"🎉 **Admission Successful!** Member **{m_name}** ({reg_type}) registered.")
                st.cache_data.clear()
                st.rerun()

# MODULE 3: ATTENDANCE
with tab_att:
    st.markdown("### ⏱️ Quick Member Check-in")
    col_a1, col_a2 = st.columns([2, 1])
    with col_a1:
        att_search = st.text_input("Enter Member Name / Contact no. / SR.NO:")
        if st.button("🔴 Register Check-in"):
            if att_search:
                log_entry = {"Time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Member": att_search, "Status": "Checked-In"}
                st.session_state["attendance_log"].append(log_entry)
                st.success(f"✅ Check-in recorded for **{att_search}**")
    with col_a2:
        st.markdown("#### 📋 Attendance Log")
        if st.session_state["attendance_log"]:
            st.dataframe(pd.DataFrame(st.session_state["attendance_log"]), use_container_width=True)

# MODULE 4: LEADS
with tab_leads:
    st.markdown("### 🎯 Lead & Inquiry Manager")
    with st.form("add_lead_form_working"):
        col_l1, col_l2, col_l3 = st.columns(3)
        with col_l1:
            lead_name = st.text_input("Lead Name:")
            lead_phone = st.text_input("Contact no.:")
        with col_l2:
            lead_status = st.selectbox("Lead Status:", ["Hot", "Warm", "Cold"])
            lead_interest = st.selectbox("Facility / Package:", ["GT 1 Year", "GT 6 Month", "GT 3 Month", "PT", "Steam Bath"])
        with col_l3:
            follow_date = st.date_input("Follow-up Date:", value=datetime.now() + timedelta(days=2))
            lead_notes = st.text_input("Notes:")
        lead_submit = st.form_submit_button("📌 Save Lead")
    if lead_submit and lead_name:
        st.session_state["leads_data"].append({
            "Date": datetime.now().strftime("%Y-%m-%d"), "Name": lead_name, "Contact no.": lead_phone,
            "Status": lead_status, "Interest": lead_interest, "Follow-up": follow_date.strftime("%Y-%m-%d"), "Notes": lead_notes
        })
        st.success(f"🎯 Lead **{lead_name}** Saved!")
    if st.session_state["leads_data"]:
        st.dataframe(pd.DataFrame(st.session_state["leads_data"]), use_container_width=True)

# 💰 MODULE: FINANCIAL SUMMARY
with tab_accounts:
    st.markdown("### 💰 Financial Summary & Pending Dues")
    col_acc1, col_acc2 = st.columns(2)
    with col_acc1:
        st.markdown("#### 📅 Date-Wise Collection Summary")
        selected_rep_date = st.date_input("Select Report Date:", value=datetime.now())
        if not df.empty and 'Parsed_Admission_Date' in df.columns:
            day_df = df[df['Parsed_Admission_Date'].dt.date == selected_rep_date].copy()
            total_day_collection = day_df['Paid'].sum() if not day_df.empty and 'Paid' in day_df.columns else 0
            st.metric(label=f"Total Collection on {selected_rep_date.strftime('%d %b %Y')}", value=f"₹{total_day_collection:,.2f}")
            if not day_df.empty and 'Mode of payment' in day_df.columns:
                st.dataframe(day_df.groupby('Mode of payment')['Paid'].sum().reset_index(), use_container_width=True)
    with col_acc2:
        st.markdown("#### ⚠️ Pending Dues / Remaining Balance")
        if not df.empty and 'Fees' in df.columns and 'Paid' in df.columns:
            df_dues = df.copy()
            df_dues['Pending'] = pd.to_numeric(df_dues['Fees'], errors='coerce').fillna(0) - pd.to_numeric(df_dues['Paid'], errors='coerce').fillna(0)
            pending_df = df_dues[df_dues['Pending'] > 0]
            st.metric(label="Total Outstanding Pending Dues", value=f"₹{pending_df['Pending'].sum():,.2f}")
            if not pending_df.empty:
                st.dataframe(pending_df[['SR.NO', 'Name', 'Contact no.', 'Fees', 'Paid', 'Pending']], use_container_width=True)

# 📑 MODULE: DAILY ACCOUNTING REGISTER
with tab_daily_ledger:
    st.markdown("### 📑 Daily Accounting Register & Ledger Entry")
    with st.form("add_daily_ledger_entry"):
        st.markdown("#### ➕ Add New Daily Transaction Entry")
        col_dl1, col_dl2, col_dl3 = st.columns(3)
        with col_dl1:
            dl_date = st.date_input("Transaction Date:*", value=datetime.now())
            dl_particular = st.text_input("Particular / Description:*", placeholder="e.g. Collection, Nashta, Salary")
            dl_target_sheet = st.selectbox("Target Month Sheet:*", daily_sheets if daily_sheets else ["Oct 2026"], index=0)
        with col_dl2:
            dl_collection = st.number_input("Collection (₹):", value=0.0, step=100.0)
            dl_cash = st.number_input("Cash Amount (₹):", value=0.0, step=100.0)
            dl_axis = st.number_input("Axis A/C (₹):", value=0.0, step=100.0)
            dl_pers = st.number_input("Pers A/C (₹):", value=0.0, step=100.0)
            dl_check = st.number_input("Check Amount (₹):", value=0.0, step=100.0)
        with col_dl3:
            dl_locker = st.number_input("Bank Locker (₹):", value=0.0, step=100.0)
            dl_payal = st.number_input("Payal Cash (₹):", value=0.0, step=100.0)
            dl_kartikey = st.number_input("Kartikey Sir (₹):", value=0.0, step=100.0)
            dl_expense = st.number_input("Expense (₹):", value=0.0, step=100.0)
            dl_superbazaar = st.number_input("Super Bazaar (₹):", value=0.0, step=100.0)
        dl_submit = st.form_submit_button("✅ Save Daily Transaction Entry")
    if dl_submit and dl_particular:
        entry_dict = {
            'Date': dl_date.strftime('%Y-%m-%d'), 'Particular': dl_particular,
            'Collection': dl_collection if dl_collection != 0 else '', 'Cash': dl_cash if dl_cash != 0 else '',
            'Axis A/C': dl_axis if dl_axis != 0 else '', 'Pers A/C': dl_pers if dl_pers != 0 else '',
            'Check': dl_check if dl_check != 0 else '', 'Bank Locker': dl_locker if dl_locker != 0 else '',
            'Payal Cash': dl_payal if dl_payal != 0 else '', 'Kartikey Sir': dl_kartikey if dl_kartikey != 0 else '',
            'Expense': dl_expense if dl_expense != 0 else '', 'Super Bazaar': dl_superbazaar if dl_superbazaar != 0 else ''
        }
        success, msg = append_daily_account_entry(daily_acc_file_path, entry_dict, target_sheet=dl_target_sheet)
        if success:
            st.success(f"🎉 Entry Saved! '{dl_particular}' successfully added!")
            st.cache_data.clear()
            st.rerun()
    st.markdown("---")
    st.markdown("#### 📋 Existing Daily Accounting Records")
    if not df_daily_acc.empty:
        selected_sheet_filter = st.selectbox("Filter Records by Sheet:", ["ALL"] + daily_sheets)
        view_acc_df = df_daily_acc[df_daily_acc['Sheet'] == selected_sheet_filter] if selected_sheet_filter != "ALL" else df_daily_acc
        st.dataframe(view_acc_df, use_container_width=True)

# SECURITY MODULE
with tab_security:
    st.markdown("### 🔑 Change Your Password")
    with st.form("change_password_form"):
        old_pwd = st.text_input("Current Password:*", type="password")
        new_pwd = st.text_input("New Password:*", type="password")
        confirm_pwd = st.text_input("Confirm New Password:*", type="password")
        pwd_submit = st.form_submit_button("🔄 Update Password")
    if pwd_submit and old_pwd == current_user["password"] and new_pwd == confirm_pwd and new_pwd != "":
        st.session_state["user_accounts"][user_email]["password"] = new_pwd
        st.success("🎉 Password Successfully Updated!")

# SUPER ADMIN PANEL
if current_user["role"] == "SUPER_ADMIN":
    tab_admin = tab_objects[8]
    with tab_admin:
        st.markdown("### 👑 KFH Master Admin Panel")
        acc_df = pd.DataFrame([
            {"Username": k, "Gym Name": v["gym_name"], "Mobile": v.get("mobile", "N/A"), "Role": v["role"]}
            for k, v in st.session_state["user_accounts"].items()
        ])
        st.dataframe(acc_df, use_container_width=True)