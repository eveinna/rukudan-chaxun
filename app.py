# -*- coding: utf-8 -*-
"""
入库单查询系统 - Web版 (Streamlit)
支持多用户登录、中心数据库、云端部署
含报损数据板块（内置模块，便于单文件部署）
"""

import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io
import os
import json

# ===================== 配置 =====================
USERS = {
    "admin": {"password": "admin123", "role": "admin"},
    "user1": {"password": "123456", "role": "user"},
    "user2": {"password": "123456", "role": "user"},
}

DATA_FILE = "入库单数据.xlsx"
USERS_FILE = "users.json"

COL_KEYS = {
    "日期": ["日期", "时间", "date"],
    "供应商": ["供应商", "供货商"],
    "商品名称": ["商品名称", "品名", "货品"],
    "数量": ["数量", "件数"],
    "重量": ["重量", "斤", "kg"],
    "单价": ["单价", "价格"],
    "箱费": ["箱费"],
    "金额": ["金额", "总价", "合计"],
    "采购员": ["采购员", "业务员"],
    "发货地": ["发货地", "产地"],
}

DISPLAY_COLS = ["日期", "供应商", "商品名称", "数量", "重量", "单价", "箱费", "金额", "采购员", "发货地"]

# ===================== 报损数据 - 板块配置 =====================
# 在此处增删板块名称和对应文件名
DAMAGE_SECTIONS = {
    "MMT板块": "报损-MMT板块.xlsx",
    "W板块": "报损-W板块.xlsx",
    "CH板块": "报损-CH板块.xlsx",
    "XG板块": "报损-XG板块.xlsx",
}

DISPLAY_COLS_DAMAGE = [
    "板块名称", "序号", "售后日期", "供应商名称", "配送日期", "门店名称",
    "商品名称", "下单件数", "破损重量", "金额报损",
    "当日到货报损", "榴莲三天报损", "报损原因", "处理方式", "凭证图片", "备注"
]

# ===================== 通用工具函数 =====================
def clean_cols(df):
    """清理列名中的换行符"""
    df = df.copy()
    df.columns = [str(c).replace('\n', '').strip() for c in df.columns]
    return df

def convert_excel_date(val):
    """Excel日期转换"""
    if pd.isna(val):
        return ""
    try:
        if isinstance(val, (datetime, date)):
            return val.strftime("%Y-%m-%d")
        if isinstance(val, (int, float)):
            if 1 < val < 100000:
                d = datetime(1899, 12, 30) + timedelta(days=int(val))
                return d.strftime("%Y-%m-%d")
        return str(val)
    except:
        return str(val)

def get_date_range(quick_mode, min_date=None, max_date=None):
    """根据快捷选项计算日期范围"""
    today = date.today()
    if quick_mode == "今天":
        s = e = today
    elif quick_mode == "昨天":
        s = e = today - timedelta(days=1)
    elif quick_mode == "本周":
        s = today - timedelta(days=today.weekday())
        e = today
    elif quick_mode == "本月":
        s = date(today.year, today.month, 1)
        e = today
    elif quick_mode == "本季":
        q = (today.month - 1) // 3
        s = date(today.year, q * 3 + 1, 1)
        e = today
    elif quick_mode == "本年":
        s = date(today.year, 1, 1)
        e = today
    elif quick_mode == "全部":
        s = min_date or today
        e = max_date or today
    else:
        s = None
        e = None
    return s, e

# ===================== 用户管理 =====================
def load_users():
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return USERS.copy()
    return USERS.copy()

def save_users(users):
    with open(USERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(users, f, ensure_ascii=False, indent=2)

def check_login(username, password):
    users = load_users()
    user = users.get(username)
    if user and user.get("password") == password:
        return True, user.get("role", "user")
    return False, None

def login_page():
    st.title("📦 入库单查询系统")
    st.markdown("---")
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.subheader("请登录")
        username = st.text_input("用户名", key="login_username")
        password = st.text_input("密码", type="password", key="login_password")
        if st.button("登录", use_container_width=True):
            success, role = check_login(username, password)
            if success:
                st.session_state.logged_in = True
                st.session_state.username = username
                st.session_state.role = role
                st.rerun()
            else:
                st.error("用户名或密码错误！")

# ===================== 数据加载 =====================
@st.cache_data(ttl=300)
def load_data():
    try:
        if os.path.exists(DATA_FILE):
            df = pd.read_excel(DATA_FILE, header=1)
        else:
            st.warning(f"数据文件 {DATA_FILE} 不存在，请先上传数据文件。")
            return pd.DataFrame()
        col_map = {}
        for standard_name, variants in COL_KEYS.items():
            for col in df.columns:
                if str(col).strip() in variants:
                    col_map[standard_name] = col
                    break
        return df, col_map
    except Exception as e:
        st.error(f"加载数据失败：{e}")
        return pd.DataFrame(), {}

# ===================== 管理员功能 =====================
def admin_page():
    st.title("⚙️ 系统管理")
    tab1, tab2 = st.tabs(["📤 数据管理", "👥 用户管理"])

    with tab1:
        st.markdown("### 上传/更新数据文件")
        st.info("上传新的 Excel 文件将替换现有数据。文件第一行应为标题，第二行为列名。")
        uploaded_file = st.file_uploader("选择 Excel 文件", type=['xlsx', 'xls'])
        if uploaded_file is not None:
            if st.button("确认上传", type="primary"):
                try:
                    with open(DATA_FILE, 'wb') as f:
                        f.write(uploaded_file.getvalue())
                    load_data.clear()
                    st.success("✅ 数据文件上传成功！")
                    st.info("数据已更新，返回查询页面查看。")
                except Exception as e:
                    st.error(f"上传失败：{e}")
        st.markdown("---")
        st.markdown("### 当前数据信息")
        if os.path.exists(DATA_FILE):
            file_size = os.path.getsize(DATA_FILE) / 1024
            st.write(f"- 文件名：{DATA_FILE}")
            st.write(f"- 文件大小：{file_size:.1f} KB")
            df, col_map = load_data()
            if not df.empty:
                st.write(f"- 数据行数：{len(df)} 条")
                st.write(f"- 数据列数：{len(df.columns)} 列")
                st.write(f"- 识别到的列：{', '.join(col_map.keys())}")
        else:
            st.warning("暂无数据文件")

    with tab2:
        st.markdown("### 添加新用户")
        col1, col2, col3 = st.columns(3)
        with col1:
            new_username = st.text_input("用户名", key="new_username")
        with col2:
            new_password = st.text_input("密码", type="password", key="new_password")
        with col3:
            new_role = st.selectbox("角色", ["user", "admin"], key="new_role")
        if st.button("添加用户", type="primary"):
            if new_username and new_password:
                users = load_users()
                if new_username in users:
                    st.error(f"用户 '{new_username}' 已存在！")
                else:
                    users[new_username] = {"password": new_password, "role": new_role}
                    save_users(users)
                    st.success(f"✅ 用户 '{new_username}' 添加成功！")
            else:
                st.warning("请填写用户名和密码")
        st.markdown("---")
        st.markdown("### 现有用户列表")
        users = load_users()
        for username, info in users.items():
            col1, col2, col3, col4 = st.columns([2, 1, 1, 1])
            with col1:
                st.write(f"**{username}**")
            with col2:
                role_badge = "🔴 管理员" if info.get("role") == "admin" else "🔵 普通用户"
                st.write(role_badge)
            with col3:
                if username != st.session_state.username:
                    if st.button("删除", key=f"del_{username}"):
                        del users[username]
                        save_users(users)
                        st.rerun()
            with col4:
                if username != "admin":
                    if st.button("重置密码", key=f"reset_{username}"):
                        users[username]["password"] = "123456"
                        save_users(users)
                        st.success(f"{username} 密码已重置为 123456")

# ===================== 报损数据页面 =====================
def damage_report_page():
    """报损数据管理页面"""
    st.title("📉 报损数据管理")

    # ===== 文件上传区域 =====
    st.markdown("### 📂 上传板块数据文件")
    st.info("请依次上传各板块的 Excel 文件（表头格式需一致），系统将自动合并并标注来源板块。")

    uploaded = {}
    cols = st.columns(4)
    section_names = list(DAMAGE_SECTIONS.keys())

    for i, section in enumerate(section_names):
        with cols[i]:
            st.markdown(f"**{section}**")
            f = st.file_uploader(f"选择 {section} 文件", type=['xlsx', 'xls'], key=f"upload_{section}")
            if f:
                st.success(f"✅ 已选择：{f.name}")
                uploaded[section] = f
            else:
                st.write("⏳ 未上传")

    st.markdown("---")
    if st.button("🔄 合并数据", type="primary", use_container_width=True):
        if len(uploaded) == 0:
            st.warning("请至少上传一个板块的数据文件！")
        else:
            all_dfs = []
            errors = []
            for section, file in uploaded.items():
                try:
                    df = pd.read_excel(file, header=1)
                    df = clean_cols(df)
                    df['板块名称'] = section
                    all_dfs.append(df)
                    st.success(f"✅ {section}：{len(df)} 条记录")
                except Exception as e:
                    errors.append(f"{section}：{str(e)}")

            if errors:
                for err in errors:
                    st.error(f"❌ {err}")

            if all_dfs:
                combined = pd.concat(all_dfs, ignore_index=True)
                st.session_state['damage_df'] = combined
                st.session_state['damage_loaded'] = True
                st.success(f"🎉 合并完成！共 {len(combined)} 条记录，来源：{', '.join(uploaded.keys())}")
                st.rerun()

    # ===== 显示数据 =====
    if st.session_state.get('damage_loaded', False):
        _show_damage_data()
    else:
        st.info("👆 请先上传板块数据文件，然后点击「合并数据」按钮")


def _show_damage_data():
    """显示报损数据、筛选、统计、导出"""
    df = st.session_state['damage_df']

    st.markdown("---")
    st.markdown("### 📊 统计概览")

    total_records = len(df)
    total_amount = pd.to_numeric(df['金额报损'], errors='coerce').sum()
    total_weight = pd.to_numeric(df['破损重量'], errors='coerce').sum()

    stat_cols = st.columns(4)
    with stat_cols[0]:
        st.metric("总记录数", f"{total_records} 条")
    with stat_cols[1]:
        st.metric("总报损金额", f"¥{total_amount:,.2f}")
    with stat_cols[2]:
        st.metric("总破损重量", f"{total_weight:,.1f} kg")
    with stat_cols[3]:
        st.metric("板块数量", f"{df['板块名称'].nunique()} 个")

    st.markdown("#### 各板块明细")
    # 先转换数值列，处理非数字值
    df_agg = df.copy()
    df_agg['金额报损_num'] = pd.to_numeric(df_agg['金额报损'], errors='coerce').fillna(0)
    df_agg['破损重量_num'] = pd.to_numeric(df_agg['破损重量'], errors='coerce').fillna(0)
    
    section_stats = df_agg.groupby('板块名称').agg(
        记录数=('序号', 'count'),
        总金额=('金额报损_num', 'sum'),
        总重量=('破损重量_num', 'sum')
    ).reset_index()
    section_stats['总金额'] = section_stats['总金额'].apply(lambda x: f"¥{x:,.2f}")
    section_stats['总重量'] = section_stats['总重量'].apply(lambda x: f"{x:,.1f} kg")
    section_stats.columns = ['板块名称', '记录数', '总报损金额', '总破损重量']
    st.dataframe(section_stats, use_container_width=True, hide_index=True)

    # ===== 筛选区域 =====
    st.markdown("---")
    st.markdown("### 🔍 筛选条件")

    filter_cols = st.columns(4)

    with filter_cols[0]:
        st.markdown("**板块**")
        sections = ["全部"] + list(DAMAGE_SECTIONS.keys())
        section_filter = st.selectbox("选择板块", sections, key="filter_section")

    with filter_cols[1]:
        st.markdown("**日期范围**")
        try:
            df['售后日期_dt'] = pd.to_datetime(df['售后日期'], errors='coerce')
            min_date = df['售后日期_dt'].min().date() if df['售后日期_dt'].notna().any() else date.today()
            max_date = df['售后日期_dt'].max().date() if df['售后日期_dt'].notna().any() else date.today()

            quick = st.selectbox(
                "快捷选择",
                ["自定义", "今天", "昨天", "本周", "本月", "本季", "本年", "全部"],
                key="filter_quick"
            )

            if quick == "自定义":
                s_date = st.date_input("开始日期", min_date, key="filter_sdate")
                e_date = st.date_input("结束日期", max_date, key="filter_edate")
            else:
                s_date, e_date = get_date_range(quick, min_date, max_date)
                if s_date and e_date:
                    st.caption(f"📅 {s_date} ~ {e_date}")
        except Exception as ex:
            s_date = e_date = None
            st.error(f"日期处理异常：{ex}")

    with filter_cols[2]:
        st.markdown("**门店名称**")
        stores = ["全部"] + sorted(df['门店名称'].dropna().astype(str).unique().tolist())
        store_filter = st.selectbox("选择门店", stores, key="filter_store")

    with filter_cols[3]:
        st.markdown("**商品名称**")
        products = ["全部"] + sorted(df['商品名称'].dropna().astype(str).unique().tolist())
        product_filter = st.selectbox("选择商品", products, key="filter_product")

    filter_cols2 = st.columns([2, 2, 1])
    with filter_cols2[0]:
        st.markdown("**报损原因**")
        reasons = ["全部"] + sorted(df['报损原因'].dropna().astype(str).unique().tolist())
        reason_filter = st.selectbox("选择原因", reasons, key="filter_reason")
    with filter_cols2[1]:
        st.markdown("**处理方式**")
        methods = ["全部"] + sorted(df['处理方式'].dropna().astype(str).unique().tolist())
        method_filter = st.selectbox("选择方式", methods, key="filter_method")

    # 应用筛选
    filtered = df.copy()
    if section_filter != "全部":
        filtered = filtered[filtered['板块名称'] == section_filter]
    if store_filter != "全部":
        filtered = filtered[filtered['门店名称'].astype(str) == store_filter]
    if product_filter != "全部":
        filtered = filtered[filtered['商品名称'].astype(str) == product_filter]
    if reason_filter != "全部":
        filtered = filtered[filtered['报损原因'].astype(str) == reason_filter]
    if method_filter != "全部":
        filtered = filtered[filtered['处理方式'].astype(str) == method_filter]
    if s_date and e_date and '售后日期_dt' in filtered.columns:
        filtered = filtered[
            (filtered['售后日期_dt'].dt.date >= s_date) &
            (filtered['售后日期_dt'].dt.date <= e_date)
        ]

    # ===== 数据明细 =====
    st.markdown("---")
    st.markdown(f"### 📋 数据明细（筛选后 {len(filtered)} 条）")

    display_cols = [c for c in filtered.columns if c not in ['售后日期_dt']]
    
    # 转换数据为可显示格式（避免类型不兼容错误）
    display_df = filtered[display_cols].copy()
    for col in display_df.columns:
        if col == '售后日期':
            display_df[col] = display_df[col].apply(convert_excel_date)
        else:
            # 统一转为字符串，避免混合类型导致渲染失败
            display_df[col] = display_df[col].astype(str).replace('nan', '').replace('None', '')
    
    st.dataframe(
        display_df,
        use_container_width=True,
        height=400,
        hide_index=True,
    )

    # ===== 导出 =====
    st.markdown("---")
    export_col1, export_col2 = st.columns([1, 3])
    with export_col1:
        fmt = st.radio("导出格式", ["Excel (.xlsx)", "CSV (.csv)"])
    with export_col2:
        if st.button("📥 导出筛选结果", use_container_width=True):
            exp_df = filtered[display_cols].copy()
            if '售后日期' in exp_df.columns:
                exp_df['售后日期'] = exp_df['售后日期'].apply(convert_excel_date)

            if fmt == "Excel (.xlsx)":
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                    exp_df.to_excel(writer, index=False, sheet_name='报损数据')
                buf.seek(0)
                st.download_button(
                    label="⬇️ 下载 Excel",
                    data=buf,
                    file_name=f"报损数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                csv = exp_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="⬇️ 下载 CSV",
                    data=csv,
                    file_name=f"报损数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                )

# ===================== 主界面 =====================
def main_app():
    st.title("📦 入库单查询系统")

    with st.sidebar:
        st.markdown(f"### 👤 当前用户：{st.session_state.username}")
        role = st.session_state.get("role", "user")
        role_display = "🔴 管理员" if role == "admin" else "🔵 普通用户"
        st.markdown(f"**{role_display}**")
        st.markdown("---")

        if role == "admin":
            page = st.radio("导航", ["📦 数据查询", "📉 报损数据", "⚙️ 系统管理"])
        else:
            page = st.radio("导航", ["📦 数据查询", "📉 报损数据"])

        st.markdown("---")
        if st.button("退出登录"):
            st.session_state.logged_in = False
            st.session_state.username = ""
            st.session_state.role = ""
            st.rerun()

        st.markdown("---")
        st.markdown("### 📋 操作说明")
        st.markdown("""
        1. 选择日期范围筛选
        2. 使用下拉框筛选供应商等
        3. 点击列头可排序
        4. 可导出筛选结果
        """)

    if page == "⚙️ 系统管理":
        admin_page()
        return
    elif page == "📉 报损数据":
        damage_report_page()
        return

    # ===== 数据查询页面 =====
    df, col_map = load_data()
    if df.empty:
        st.info("请上传数据文件或联系管理员添加数据。")
        return

    st.markdown(f"**数据总量**：{len(df)} 条记录")

    st.markdown("### 🔍 筛选条件")
    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("**日期范围**")
        date_col = col_map.get("日期")
        if date_col and date_col in df.columns:
            dates = df[date_col].apply(convert_excel_date)
            dates = pd.to_datetime(dates, errors='coerce')
            min_date = dates.min().date() if not dates.isna().all() else date.today()
            max_date = dates.max().date() if not dates.isna().all() else date.today()
            quick = st.selectbox("快捷选择", ["自定义", "今天", "昨天", "本周", "本月", "本季", "本年", "全部"])
            if quick == "自定义":
                start_date = st.date_input("开始日期", min_date)
                end_date = st.date_input("结束日期", max_date)
            elif quick == "今天":
                start_date = end_date = date.today()
            elif quick == "昨天":
                start_date = end_date = date.today() - timedelta(days=1)
            elif quick == "本周":
                start_date = date.today() - timedelta(days=date.today().weekday())
                end_date = date.today()
            elif quick == "本月":
                start_date = date(date.today().year, date.today().month, 1)
                end_date = date.today()
            elif quick == "本季":
                q = (date.today().month - 1) // 3
                start_date = date(date.today().year, q * 3 + 1, 1)
                end_date = date.today()
            elif quick == "本年":
                start_date = date(date.today().year, 1, 1)
                end_date = date.today()
            else:
                start_date = min_date
                end_date = max_date
        else:
            start_date = end_date = None
            st.warning("未找到日期列")

    with col2:
        st.markdown("**供应商**")
        supplier_col = col_map.get("供应商")
        if supplier_col and supplier_col in df.columns:
            suppliers = ["全部"] + sorted(df[supplier_col].dropna().astype(str).unique().tolist())
            supplier_filter = st.selectbox("选择供应商", suppliers)
        else:
            supplier_filter = "全部"

    with col3:
        st.markdown("**采购员**")
        buyer_col = col_map.get("采购员")
        if buyer_col and buyer_col in df.columns:
            buyers = ["全部"] + sorted(df[buyer_col].dropna().astype(str).unique().tolist())
            buyer_filter = st.selectbox("选择采购员", buyers)
        else:
            buyer_filter = "全部"

    col4, col5 = st.columns(2)
    with col4:
        st.markdown("**商品名称**")
        product_col = col_map.get("商品名称")
        if product_col and product_col in df.columns:
            products = ["全部"] + sorted(df[product_col].dropna().astype(str).unique().tolist())
            product_filter = st.selectbox("选择商品", products)
        else:
            product_filter = "全部"

    with col5:
        st.markdown("**发货地**")
        origin_col = col_map.get("发货地")
        if origin_col and origin_col in df.columns:
            origins = ["全部"] + sorted(df[origin_col].dropna().astype(str).unique().tolist())
            origin_filter = st.selectbox("选择发货地", origins)
        else:
            origin_filter = "全部"

    # 应用筛选
    filtered_df = df.copy()
    if start_date and end_date and date_col:
        dates = pd.to_datetime(filtered_df[date_col].apply(convert_excel_date), errors='coerce')
        filtered_df = filtered_df[
            (dates.dt.date >= start_date) & (dates.dt.date <= end_date)
        ]
    if supplier_filter != "全部" and supplier_col:
        filtered_df = filtered_df[filtered_df[supplier_col].astype(str) == supplier_filter]
    if buyer_filter != "全部" and buyer_col:
        filtered_df = filtered_df[filtered_df[buyer_col].astype(str) == buyer_filter]
    if product_filter != "全部" and product_col:
        filtered_df = filtered_df[filtered_df[product_col].astype(str) == product_filter]
    if origin_filter != "全部" and origin_col:
        filtered_df = filtered_df[filtered_df[origin_col].astype(str) == origin_filter]

    # 统计信息
    st.markdown("---")
    st.markdown("### 📊 统计信息")
    stat_col1, stat_col2, stat_col3, stat_col4, stat_col5 = st.columns(5)
    amount_col = col_map.get("金额")
    qty_col = col_map.get("数量")
    weight_col = col_map.get("重量")
    supplier_col = col_map.get("供应商")

    with stat_col1:
        st.metric("筛选记录数", f"{len(filtered_df)} 条")
    with stat_col2:
        if amount_col and amount_col in filtered_df.columns:
            st.metric("总金额", f"¥{pd.to_numeric(filtered_df[amount_col], errors='coerce').sum():,.2f}")
        else:
            st.metric("总金额", "N/A")
    with stat_col3:
        if qty_col and qty_col in filtered_df.columns:
            st.metric("总数量", f"{pd.to_numeric(filtered_df[qty_col], errors='coerce').sum():,.0f}")
        else:
            st.metric("总数量", "N/A")
    with stat_col4:
        if weight_col and weight_col in filtered_df.columns:
            st.metric("总重量", f"{pd.to_numeric(filtered_df[weight_col], errors='coerce').sum():,.0f}")
        else:
            st.metric("总重量", "N/A")
    with stat_col5:
        if supplier_col and supplier_col in filtered_df.columns:
            st.metric("供应商数", f"{filtered_df[supplier_col].nunique()} 家")
        else:
            st.metric("供应商数", "N/A")

    # 数据表格
    st.markdown("---")
    st.markdown("### 📋 数据明细")
    display_cols_actual = [col_map.get(c) for c in DISPLAY_COLS if col_map.get(c) in filtered_df.columns]
    if display_cols_actual:
        display_df = filtered_df[display_cols_actual].copy()
        if date_col and date_col in display_df.columns:
            display_df[date_col] = display_df[date_col].apply(convert_excel_date)
        rename_map = {v: k for k, v in col_map.items() if v in display_cols_actual}
        display_df = display_df.rename(columns=rename_map)
        st.dataframe(display_df, use_container_width=True, height=400, hide_index=True)
        if len(display_df) > 1000:
            st.info(f"显示前 1000 条，共 {len(display_df)} 条")
    else:
        st.warning("没有可显示的列")

    # 导出
    st.markdown("---")
    st.markdown("### 📥 导出数据")
    col_exp1, col_exp2 = st.columns([1, 3])
    with col_exp1:
        export_format = st.radio("导出格式", ["Excel (.xlsx)", "CSV (.csv)"])
    with col_exp2:
        if st.button("📥 导出筛选结果", use_container_width=True):
            if export_format == "Excel (.xlsx)":
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                    filtered_df.to_excel(writer, index=False, sheet_name='入库单')
                buf.seek(0)
                st.download_button(
                    label="⬇️ 下载 Excel 文件",
                    data=buf,
                    file_name=f"入库单筛选结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                csv = filtered_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="⬇️ 下载 CSV 文件",
                    data=csv,
                    file_name=f"入库单筛选结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                )

# ===================== 主程序 =====================
def main():
    st.set_page_config(
        page_title="入库单查询系统",
        page_icon="📦",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False
        st.session_state.username = ""
        st.session_state.role = ""
    if st.session_state.logged_in:
        main_app()
    else:
        login_page()

if __name__ == "__main__":
    main()
