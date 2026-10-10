import datetime
import io
import pandas as pd
import streamlit as st
import hashlib

# นำเข้าเครื่องมือสำหรับตกแต่ง Excel
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

# นำเข้าเครื่องมือสำหรับคุยกับ Supabase (PostgreSQL)
import psycopg2
from psycopg2 import IntegrityError
from sqlalchemy import create_engine

# ---------------------------------------------------------
# SETUP & CONFIG
# ---------------------------------------------------------
st.set_page_config(page_title="ระบบวันลา เวชระเบียน", page_icon="🏥", layout="wide")

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.emp_id = ""
    st.session_state.emp_name = ""
    st.session_state.role = ""

LEAVE_QUOTAS = {
    "ลากิจ": 3,
    "ลาป่วย": 15,
    "ลาพักผ่อน": 10
}

def make_hash(password):
    return hashlib.sha256(password.encode()).hexdigest()

# ---------------------------------------------------------
# DATABASE SETUP (SUPABASE POSTGRESQL)
# ---------------------------------------------------------
# ลิงก์เชื่อมต่อฐานข้อมูลของคุณ
DB_URL = "postgresql://postgres:GNKCxEPuD11@db.koslwumkjvvlhboutpcy.supabase.co:5432/postgres"

# Engine สำหรับให้ Pandas ใช้อ่านข้อมูลตารางได้ง่ายๆ
engine = create_engine(DB_URL.replace("postgresql://", "postgresql+psycopg2://"))

def get_connection():
    # ฟังก์ชันเชื่อมต่อฐานข้อมูล
    return psycopg2.connect(DB_URL)

def init_db():
    conn = get_connection()
    c = conn.cursor()
    # สร้างตารางพนักงาน (เปลี่ยนเป็น SERIAL สำหรับ Postgres)
    c.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY,
            emp_id TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            phone TEXT NOT NULL,
            email TEXT NOT NULL,
            password TEXT NOT NULL,
            role TEXT NOT NULL
        )
    ''')
    # สร้างตารางการลา
    c.execute('''
        CREATE TABLE IF NOT EXISTS leave_requests (
            id SERIAL PRIMARY KEY,
            emp_id TEXT NOT NULL,
            user_name TEXT NOT NULL,
            position TEXT NOT NULL,
            approver TEXT NOT NULL,
            leave_type TEXT NOT NULL,
            start_date TEXT NOT NULL,
            start_day_type TEXT NOT NULL,
            end_date TEXT NOT NULL,
            end_day_type TEXT NOT NULL,
            total_days REAL NOT NULL,
            reason TEXT
        )
    ''')
    
    # ตรวจสอบและสร้างบัญชี Admin ถ้ายังไม่มี
    c.execute("SELECT * FROM users WHERE emp_id='admin'")
    if not c.fetchone():
        c.execute("INSERT INTO users (emp_id, name, phone, email, password, role) VALUES (%s, %s, %s, %s, %s, %s)", 
                  ('admin', 'Administrator', '-', 'admin@company.com', make_hash('1234'), 'admin'))
    conn.commit()
    conn.close()

# เริ่มต้นสร้างตาราง
init_db()

# --- ฟังก์ชันจัดการผู้ใช้และลา ---
def add_user(emp_id, name, phone, email, password, role='user'):
    conn = get_connection()
    c = conn.cursor()
    try:
        c.execute("INSERT INTO users (emp_id, name, phone, email, password, role) VALUES (%s, %s, %s, %s, %s, %s)", 
                  (emp_id.strip(), name.strip(), phone.strip(), email.strip(), make_hash(password), role))
        conn.commit()
        success = True
    except IntegrityError:
        success = False 
    conn.close()
    return success

def delete_user(emp_id):
    conn = get_connection()
    c = conn.cursor()
    c.execute("DELETE FROM users WHERE emp_id=%s", (emp_id,))
    conn.commit()
    conn.close()

def get_all_users():
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT emp_id, name FROM users WHERE role='user'")
    users = c.fetchall()
    conn.close()
    return users

def login_user(emp_id, password):
    conn = get_connection()
    c = conn.cursor()
    c.execute("SELECT role, name FROM users WHERE emp_id=%s AND password=%s", (emp_id.strip(), make_hash(password)))
    result = c.fetchone()
    conn.close()
    return result

def calculate_leave_days(start_date, start_type, end_date, end_type):
    delta_days = (end_date - start_date).days + 1
    total_days = float(delta_days)
    
    if start_date == end_date:
        if start_type in ["ครึ่งวันแรก (เช้า)", "ครึ่งวันหลัง (บ่าย)"]:
            return 0.5
        return 1.0
    else:
        if start_type in ["ครึ่งวันแรก (เช้า)", "ครึ่งวันหลัง (บ่าย)"]:
            total_days -= 0.5
        if end_type in ["ครึ่งวันแรก (เช้า)", "ครึ่งวันหลัง (บ่าย)"]:
            total_days -= 0.5
        return total_days

def add_leave_request(emp_id, user_name, position, approver, leave_type, start_date, start_day_type, end_date, end_day_type, reason):
    total_days = calculate_leave_days(start_date, start_day_type, end_date, end_day_type)
    conn = get_connection()
    c = conn.cursor()
    c.execute('''
        INSERT INTO leave_requests (emp_id, user_name, position, approver, leave_type, start_date, start_day_type, end_date, end_day_type, total_days, reason)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
    ''', (emp_id, user_name, position, approver, leave_type, str(start_date), start_day_type, str(end_date), end_day_type, total_days, reason))
    conn.commit()
    conn.close()

def get_all_leaves(emp_id=None):
    if emp_id:
        df = pd.read_sql_query(f"SELECT * FROM leave_requests WHERE emp_id='{emp_id}' ORDER BY start_date DESC", engine)
    else:
        df = pd.read_sql_query("SELECT * FROM leave_requests ORDER BY start_date DESC", engine)
    return df

def get_leave_balances(emp_id=None):
    users = get_all_users()
    df_leaves = get_all_leaves()
    current_year = datetime.datetime.now().year
    
    if not df_leaves.empty:
        df_leaves['year'] = pd.to_datetime(df_leaves['start_date']).dt.year
        df_leaves_current_year = df_leaves[df_leaves['year'] == current_year]
    else:
        df_leaves_current_year = pd.DataFrame()
    
    balances = []
    for u in users:
        u_id, u_name = u
        if emp_id and u_id != emp_id:
            continue
            
        if not df_leaves_current_year.empty:
            user_leaves = df_leaves_current_year[df_leaves_current_year['emp_id'] == u_id]
            used_personal = user_leaves[user_leaves['leave_type'] == 'ลากิจ']['total_days'].sum()
            used_sick = user_leaves[user_leaves['leave_type'] == 'ลาป่วย']['total_days'].sum()
            used_annual = user_leaves[user_leaves['leave_type'] == 'ลาพักผ่อน']['total_days'].sum()
        else:
            used_personal = used_sick = used_annual = 0.0
            
        balances.append({
            "รหัสพนักงาน": u_id,
            "ชื่อ-นามสกุล": u_name,
            "ลากิจ (ใช้ไป/โควต้า)": f"{used_personal} / {LEAVE_QUOTAS['ลากิจ']}",
            "ลากิจ (คงเหลือ)": LEAVE_QUOTAS['ลากิจ'] - used_personal,
            "ลาป่วย (ใช้ไป/โควต้า)": f"{used_sick} / {LEAVE_QUOTAS['ลาป่วย']}",
            "ลาป่วย (คงเหลือ)": LEAVE_QUOTAS['ลาป่วย'] - used_sick,
            "ลาพักผ่อน (ใช้ไป/โควต้า)": f"{used_annual} / {LEAVE_QUOTAS['ลาพักผ่อน']}",
            "ลาพักผ่อน (คงเหลือ)": LEAVE_QUOTAS['ลาพักผ่อน'] - used_annual,
        })
    return pd.DataFrame(balances)

# ---------------------------------------------------------
# FRONTEND
# ---------------------------------------------------------
if not st.session_state.logged_in:
    st.title("🔐 เข้าสู่ระบบ (หน่วยเวชระเบียน)")
    with st.container():
        log_empid = st.text_input("รหัสพนักงาน (หรือ 'admin')")
        log_password = st.text_input("รหัสผ่าน", type="password")
        if st.button("เข้าสู่ระบบ", type="primary"):
            user_data = login_user(log_empid, log_password)
            if user_data:
                st.session_state.logged_in = True
                st.session_state.emp_id = log_empid
                st.session_state.emp_name = user_data[1]
                st.session_state.role = user_data[0]
                st.rerun()
            else:
                st.error("รหัสพนักงานหรือรหัสผ่านไม่ถูกต้อง!")
else:
    st.sidebar.title(f"👤 สวัสดี, {st.session_state.emp_name}")
    st.sidebar.caption(f"สถานะ: {'ผู้ดูแลระบบ (Admin)' if st.session_state.role == 'admin' else 'พนักงาน'}")
    if st.sidebar.button("🚪 ออกจากระบบ"):
        st.session_state.logged_in = False
        st.session_state.emp_id = ""
        st.session_state.emp_name = ""
        st.session_state.role = ""
        st.rerun()

    if st.session_state.role == "admin":
        st.title("🛠️ หน้าต่างผู้ดูแลระบบ (Admin)")
        tab1, tab2, tab3, tab4 = st.tabs(["👥 จัดการพนักงาน", "📋 ตรวจสอบประวัติการลา", "📊 สรุปโควต้าวันลา", "📤 Export ข้อมูล"])

        with tab1:
            st.subheader("➕ เพิ่มบัญชีพนักงานใหม่")
            col1, col2 = st.columns(2)
            with col1:
                new_empid = st.text_input("รหัสพนักงาน (ตัวเลข 8 หลัก) *", max_chars=8)
                new_name = st.text_input("ชื่อ-นามสกุล *")
                new_phone = st.text_input("เบอร์โทรศัพท์ *")
            with col2:
                new_email = st.text_input("อีเมล (Email) *")
                new_pass = st.text_input("ตั้งรหัสผ่าน *", type="password")
                new_conf = st.text_input("ยืนยันรหัสผ่าน *", type="password")
            
            if st.button("บันทึกข้อมูลพนักงาน", type="primary"):
                if not (new_empid and new_name and new_phone and new_email and new_pass and new_conf):
                    st.warning("⚠️ กรุณากรอกข้อมูลให้ครบทุกช่อง")
                elif len(new_empid) != 8 or not new_empid.isdigit():
                    st.error("⚠️ รหัสพนักงานต้องเป็นตัวเลข 8 หลักเท่านั้น")
                elif new_pass != new_conf:
                    st.error("⚠️ รหัสผ่าน และ ยืนยันรหัสผ่าน ไม่ตรงกัน!")
                else:
                    if add_user(new_empid, new_name, new_phone, new_email, new_pass, 'user'):
                        st.success(f"✅ เพิ่มพนักงาน '{new_name}' สำเร็จ!")
                    else:
                        st.error("❌ รหัสพนักงานนี้มีอยู่ในระบบแล้ว!")
            
            st.divider()
            st.subheader("🗑️ ลบบัญชีพนักงาน")
            user_list = get_all_users()
            if user_list:
                options = {f"{u[0]} - {u[1]}": u[0] for u in user_list}
                del_choice = st.selectbox("เลือกรายชื่อที่ต้องการลบ", options.keys())
                if st.button("ลบพนักงานรายนี้", type="secondary"):
                    delete_user(options[del_choice])
                    st.success(f"ลบข้อมูลเรียบร้อยแล้ว")
                    st.rerun()

        with tab2:
            st.subheader("📋 ตรวจสอบประวัติการลาของพนักงาน")
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                filter_start = st.date_input("ค้นหาตั้งแต่วันที่:", datetime.date.today() - datetime.timedelta(days=30))
            with col_f2:
                filter_end = st.date_input("ถึงวันที่:", datetime.date.today() + datetime.timedelta(days=30))
                
            all_leaves = get_all_leaves()
            if not all_leaves.empty:
                all_leaves['start_dt'] = pd.to_datetime(all_leaves['start_date']).dt.date
                all_leaves['end_dt'] = pd.to_datetime(all_leaves['end_date']).dt.date
                mask = (all_leaves['start_dt'] <= filter_end) & (all_leaves['end_dt'] >= filter_start)
                filtered_df = all_leaves.loc[mask].copy()
                
                if not filtered_df.empty:
                    display_df = filtered_df[['emp_id', 'user_name', 'leave_type', 'start_date', 'start_day_type', 'end_date', 'end_day_type', 'total_days', 'reason']]
                    display_df.columns = ["รหัส", "ชื่อ-นามสกุล", "ประเภท", "วันเริ่ม", "รูปแบบ(เริ่ม)", "วันสิ้นสุด", "รูปแบบ(สิ้นสุด)", "รวม(วัน)", "เหตุผล"]
                    st.dataframe(display_df, use_container_width=True)
                else:
                    st.info("ไม่พบข้อมูลการลาในช่วงเวลาดังกล่าว")
            else:
                st.info("ยังไม่มีข้อมูลการลาในระบบ")

        with tab3:
            st.subheader(f"📊 สรุปสิทธิวันลาคงเหลือของพนักงานทุกคน (ประจำปี {datetime.datetime.now().year})")
            df_balances = get_leave_balances()
            if not df_balances.empty:
                st.dataframe(df_balances, use_container_width=True)
            else:
                st.info("ยังไม่มีพนักงานในระบบ")

        with tab4:
            st.subheader("📤 ดาวน์โหลดข้อมูลการลา (Excel)")
            col_ex1, col_ex2 = st.columns(2)
            with col_ex1:
                export_start = st.date_input("ออกรายงานตั้งแต่วันที่:", datetime.date.today().replace(day=1), key="exp_start")
            with col_ex2:
                export_end = st.date_input("ถึงวันที่:", datetime.date.today(), key="exp_end")
                
            all_leaves_export = get_all_leaves()
            
            if not all_leaves_export.empty:
                all_leaves_export['start_dt'] = pd.to_datetime(all_leaves_export['start_date']).dt.date
                all_leaves_export['end_dt'] = pd.to_datetime(all_leaves_export['end_date']).dt.date
                mask_exp = (all_leaves_export['start_dt'] <= export_end) & (all_leaves_export['end_dt'] >= export_start)
                df_export = all_leaves_export.loc[mask_exp].copy()
                
                if not df_export.empty:
                    df_export = df_export[['emp_id', 'user_name', 'position', 'approver', 'leave_type', 'start_date', 'start_day_type', 'end_date', 'end_day_type', 'total_days', 'reason']]
                    df_export.columns = ["รหัสพนักงาน", "ชื่อ-นามสกุล", "ตำแหน่ง", "ผู้พิจารณา", "ประเภทการลา", "วันเริ่ม", "รูปแบบ(เริ่ม)", "วันสิ้นสุด", "รูปแบบ(สิ้นสุด)", "จำนวน(วัน)", "เหตุผล"]
                    st.write(f"พบข้อมูลทั้งหมด **{len(df_export)}** รายการ สำหรับช่วงเวลาที่เลือก")
                    
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                        df_export.to_excel(writer, index=False, sheet_name="LeaveReport")
                        workbook = writer.book
                        worksheet = writer.sheets["LeaveReport"]
                        yellow_fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")
                        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                        bold_font = Font(bold=True)
                        center_align = Alignment(horizontal="center", vertical="center")
                        left_align = Alignment(horizontal="left", vertical="center")
                        
                        for cell in worksheet[1]:
                            cell.fill = yellow_fill
                            cell.font = bold_font
                            cell.border = thin_border
                            cell.alignment = center_align
                            
                        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
                            for cell in row:
                                cell.border = thin_border
                                if cell.column in [1, 5, 6, 7, 8, 9, 10]:
                                    cell.alignment = center_align
                                else:
                                    cell.alignment = left_align
                        
                        for col in worksheet.columns:
                            max_length = 0
                            column_letter = col[0].column_letter
                            for cell in col:
                                try:
                                    if len(str(cell.value)) > max_length:
                                        max_length = len(str(cell.value))
                                except:
                                    pass
                            worksheet.column_dimensions[column_letter].width = max_length + 6
                    
                    st.download_button(
                        label="📥 ดาวน์โหลดไฟล์ Excel (.xlsx)",
                        data=buffer.getvalue(),
                        file_name=f"Leave_Report_{export_start}_to_{export_end}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                else:
                    st.warning("⚠️ ไม่มีข้อมูลการลาในช่วงเวลาที่เลือก")
            else:
                st.info("ยังไม่มีข้อมูลการลาในระบบ")

    elif st.session_state.role == "user":
        st.title("📝 ระบบวันลา (หน่วยเวชระเบียน)")
        tab1, tab2 = st.tabs(["📤 ยื่นใบลา", "ประวัติและสิทธิวันลาของฉัน"])
        
        with tab1:
            df_leaves = get_all_leaves()
            with st.container():
                st.markdown("### 1. ประเภทการลา")
                leave_type = st.radio("เลือกประเภทการลาที่ต้องการ", ["ลากิจ", "ลาป่วย", "ลาพักผ่อน"], horizontal=True)
                st.divider()

                st.markdown("### 2. ข้อมูลผู้ขอลา")
                col2_1, col2_2 = st.columns(2)
                with col2_1:
                    st.text_input("ชื่อ-นามสกุล", value=st.session_state.emp_name, disabled=True)
                    position = st.text_input("ตำแหน่ง / ภาระงาน", value="เจ้าหน้าที่ หน่วยเวชระเบียน", disabled=True)
                with col2_2:
                    st.text_input("รหัสพนักงาน", value=st.session_state.emp_id, disabled=True)
                    approver = st.text_input("ผู้พิจารณา (ธุรการการลา)", value="นาย ชัชวาลย์ คงขำ, น.ส. นิตยา เจริญสุข, น.ส. ปรีญาณัช บุญนาวงษ์", disabled=True)
                st.divider()

                st.markdown("### 3. ระยะเวลาที่ต้องการลา")
                col3_1, col3_2 = st.columns(2)
                with col3_1:
                    start_date = st.date_input("ขอลาตั้งแต่วันที่", datetime.date.today())
                    start_day_type = st.selectbox("รูปแบบวันเริ่มลา", ["เต็มวัน", "ครึ่งวันแรก (เช้า)", "ครึ่งวันหลัง (บ่าย)"])
                with col3_2:
                    end_date = st.date_input("ถึงวันที่", datetime.date.today())
                    end_day_type = st.selectbox("รูปแบบวันสิ้นสุด", ["เต็มวัน", "ครึ่งวันแรก (เช้า)", "ครึ่งวันหลัง (บ่าย)"])

                st.divider()
                reason = st.text_area("4. วัตถุประสงค์ของการลา (โปรดระบุให้ชัดเจน)")
                
                can_leave = True
                if start_date > end_date:
                    st.error("⚠️ วันเริ่มต้นต้องไม่เกินวันสิ้นสุด")
                    can_leave = False
                else:
                    req_days = calculate_leave_days(start_date, start_day_type, end_date, end_day_type)
                    my_balances = get_leave_balances(st.session_state.emp_id)
                    if not my_balances.empty:
                        rem_days = my_balances.iloc[0][f"{leave_type} (คงเหลือ)"]
                        if req_days > rem_days:
                            st.error(f"❌ สิทธิการ**{leave_type}**ของท่านในปีนี้ คงเหลือเพียง {rem_days} วัน (ต้องการใช้ {req_days} วัน)")
                            can_leave = False

                    if can_leave and not df_leaves.empty:
                        delta = end_date - start_date
                        requested_dates = [start_date + datetime.timedelta(days=i) for i in range(delta.days + 1)]
                        blocked_dates = []
                        warning_dates = []

                        for d in requested_dates:
                            count_on_day = 0
                            for _, row in df_leaves.iterrows():
                                e_start = datetime.datetime.strptime(row['start_date'], "%Y-%m-%d").date()
                                e_end = datetime.datetime.strptime(row['end_date'], "%Y-%m-%d").date()
                                if e_start <= d <= e_end:
                                    count_on_day += 1
                            if count_on_day >= 2:
                                can_leave = False
                                blocked_dates.append(f"- วันที่ {d.strftime('%d/%m/%Y')} (โควต้าเต็ม 2 คนแล้ว)")
                            elif count_on_day == 1:
                                warning_dates.append(f"- วันที่ {d.strftime('%d/%m/%Y')} (มีเจ้าหน้าที่ลาแล้ว 1 คน)")
                        
                        if not can_leave:
                            st.error("❌ **ไม่สามารถยื่นใบลาได้ เนื่องจากมีผู้ลาเต็มโควต้าในวันที่ดังต่อไปนี้:**\n" + "\n".join(blocked_dates))
                        else:
                            if warning_dates:
                                st.warning("⚠️ **ข้อควรทราบ:** ในช่วงวันที่ท่านเลือก มีผู้ลาไปแล้ว 1 คน (หากท่านกดยืนยัน โควตาวันดังกล่าวจะเต็ม 2 คนพอดี)\n" + "\n".join(warning_dates))
                            st.success(f"✅ วันที่ท่านเลือก สามารถยื่นลาได้ (คำนวณวันลา: {req_days} วัน)")

                if st.button("📤 ส่งใบลา", type="primary", use_container_width=True, disabled=not can_leave):
                    if not reason.strip():
                        st.warning("⚠️ กรุณาระบุวัตถุประสงค์ของการลา")
                    else:
                        add_leave_request(st.session_state.emp_id, st.session_state.emp_name, position, approver, 
                                          leave_type, start_date, start_day_type, end_date, end_day_type, reason)
                        st.success("🎉 บันทึกข้อมูลการลาเข้าสู่ระบบเรียบร้อยแล้ว!")
                        st.balloons()
                        
        with tab2:
            st.subheader(f"📊 สิทธิวันลาคงเหลือของฉัน (ประจำปี {datetime.datetime.now().year})")
            my_balances = get_leave_balances(st.session_state.emp_id)
            if not my_balances.empty:
                col_b1, col_b2, col_b3 = st.columns(3)
                data = my_balances.iloc[0]
                with col_b1:
                    st.metric("ลากิจ (เหลือ)", f"{data['ลากิจ (คงเหลือ)']} วัน", f"จาก {LEAVE_QUOTAS['ลากิจ']} วัน", delta_color="off")
                with col_b2:
                    st.metric("ลาป่วย (เหลือ)", f"{data['ลาป่วย (คงเหลือ)']} วัน", f"จาก {LEAVE_QUOTAS['ลาป่วย']} วัน", delta_color="off")
                with col_b3:
                    st.metric("ลาพักผ่อน (เหลือ)", f"{data['ลาพักผ่อน (คงเหลือ)']} วัน", f"จาก {LEAVE_QUOTAS['ลาพักผ่อน']} วัน", delta_color="off")
            
            st.divider()
            st.subheader("📋 ประวัติการลาของฉัน")
            my_leaves = get_all_leaves(st.session_state.emp_id)
            if not my_leaves.empty:
                display_my = my_leaves[['leave_type', 'start_date', 'start_day_type', 'end_date', 'end_day_type', 'total_days', 'reason']]
                display_my.columns = ["ประเภท", "วันเริ่ม", "รูปแบบ(เริ่ม)", "วันสิ้นสุด", "รูปแบบ(สิ้นสุด)", "รวม(วัน)", "เหตุผล"]
                st.dataframe(display_my, use_container_width=True)
            else:
                st.info("ยังไม่มีประวัติการลา")