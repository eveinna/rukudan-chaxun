# -*- coding: utf-8 -*-
"""
入库单查询系统 - Web版 (Streamlit)
支持多用户登录、中心数据库、云端部署
"""

import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io
import os

# ===================== 配置 =====================
# 用户账号信息（包含角色：admin/user）
# 部署时可改为从 secrets 读取
USERS = {
    "admin": {"password": "admin123", "role": "admin"},
    "user1": {"password": "123456", "role": "user"},
    "user2": {"password": "123456", "role": "user"},
}

# 数据文件路径
DATA_FILE = "入库单数据.xlsx"
USERS_FILE = "users.json"  # 存储用户信息的文件



# 列名映射（支持多种列名变体）
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

# 显示列
DISPLAY_COLS = ["日期", "供应商", "商品名称", "数量", "重量", "单价", "箱费", "金额", "采购员", "发货地"]

# ===================== 用户管理 =====================
import json

def load_users():
    """从文件加载用户列表"""
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return USERS.copy()
    return USERS.copy()

def save_users(users):
    """保存用户列表到文件"""
    with open(USERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(users, f, ensure_ascii=False, indent=2)

def check_login(username, password):
    """验证用户登录，返回 (成功, 角色)"""
    users = load_users()
    user = users.get(username)
    if user and user.get("password") == password:
        return True, user.get("role", "user")
    return False, None

def login_page():
    """登录页面"""
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
        
        st.markdown("---")
        st.caption("默认账号：admin / admin123")

# ===================== 数据加载 =====================
@st.cache_data(ttl=300)  # 缓存5分钟
def load_data():
    """加载数据"""
    try:
        # 尝试加载 Excel 文件
        # header=1 表示从第2行读取列名（跳过第1行的标题）
        if os.path.exists(DATA_FILE):
            df = pd.read_excel(DATA_FILE, header=1)
        else:
            # 如果没有数据文件，返回空 DataFrame
            st.warning(f"数据文件 {DATA_FILE} 不存在，请先上传数据文件。")
            return pd.DataFrame()
        
        # 智能列名映射
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

def convert_excel_date(val):
    """Excel 日期序列号转日期字符串"""
    if pd.isna(val):
        return ""
    try:
        # 如果已经是日期类型
        if isinstance(val, (datetime, date)):
            return val.strftime("%Y-%m-%d")
        # 如果是数字（Excel 序列号）
        if isinstance(val, (int, float)):
            if 1 < val < 100000:  # 合理的日期范围
                d = datetime(1899, 12, 30) + timedelta(days=int(val))
                return d.strftime("%Y-%m-%d")
        return str(val)
    except:
        return str(val)

# ===================== 管理员功能 =====================
def admin_page():
    """管理员页面 - 管理数据源和用户"""
    st.title("⚙️ 系统管理")
    
    tab1, tab2 = st.tabs(["📤 数据管理", "👥 用户管理"])
    
    # ===== 数据管理标签页 =====
    with tab1:
        st.markdown("### 上传/更新数据文件")
        st.info("上传新的 Excel 文件将替换现有数据。文件第一行应为标题，第二行为列名。")
        
        uploaded_file = st.file_uploader("选择 Excel 文件", type=['xlsx', 'xls'])
        
        if uploaded_file is not None:
            if st.button("确认上传", type="primary"):
                try:
                    # 保存上传的文件
                    with open(DATA_FILE, 'wb') as f:
                        f.write(uploaded_file.getvalue())
                    # 清除缓存，强制重新加载数据
                    load_data.clear()
                    st.success("✅ 数据文件上传成功！")
                    st.info("数据已更新，返回查询页面查看。")
                except Exception as e:
                    st.error(f"上传失败：{e}")
        
        st.markdown("---")
        
        # 显示当前数据信息
        st.markdown("### 当前数据信息")
        if os.path.exists(DATA_FILE):
            file_size = os.path.getsize(DATA_FILE) / 1024  # KB
            st.write(f"- 文件名：{DATA_FILE}")
            st.write(f"- 文件大小：{file_size:.1f} KB")
            
            # 加载数据查看基本信息
            df, col_map = load_data()
            if not df.empty:
                st.write(f"- 数据行数：{len(df)} 条")
                st.write(f"- 数据列数：{len(df.columns)} 列")
                st.write(f"- 识别到的列：{', '.join(col_map.keys())}")
        else:
            st.warning("暂无数据文件")
    
    # ===== 用户管理标签页 =====
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
        
        # 显示现有用户列表
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
                if username != st.session_state.username:  # 不能删除自己
                    if st.button("删除", key=f"del_{username}"):
                        del users[username]
                        save_users(users)
                        st.rerun()
            with col4:
                if username != "admin":  # 不能重置 admin 密码
                    if st.button("重置密码", key=f"reset_{username}"):
                        users[username]["password"] = "123456"
                        save_users(users)
                        st.success(f"{username} 密码已重置为 123456")

# ===================== 主界面 =====================
def main_app():
    """主应用界面"""
    # 侧边栏
    with st.sidebar:
        st.markdown(f"### 👤 当前用户：{st.session_state.username}")
        
        # 显示角色
        role = st.session_state.get("role", "user")
        role_display = "🔴 管理员" if role == "admin" else "🔵 普通用户"
        st.markdown(f"**{role_display}**")
        
        st.markdown("---")
        
        # 导航菜单
        if role == "admin":
            page = st.radio("导航", ["📦 数据查询", "⚙️ 系统管理"])
        else:
            page = "📦 数据查询"
        
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
    
    # 根据选择的页面显示不同内容
    if page == "⚙️ 系统管理":
        admin_page()
        return
    
    # 加载数据
    df, col_map = load_data()
    
    if df.empty:
        st.info("请上传数据文件或联系管理员添加数据。")
        return
    
    # 标题
    st.title("📦 入库单查询系统")
    st.markdown(f"**数据总量**：{len(df)} 条记录")
    
    # ===================== 筛选区域 =====================
    st.markdown("### 🔍 筛选条件")
    
    col1, col2, col3 = st.columns(3)
    
    # 日期筛选
    with col1:
        st.markdown("**日期范围**")
        date_col = col_map.get("日期")
        if date_col and date_col in df.columns:
            # 获取日期范围
            dates = df[date_col].apply(convert_excel_date)
            dates = pd.to_datetime(dates, errors='coerce')
            min_date = dates.min().date() if not dates.isna().all() else date.today()
            max_date = dates.max().date() if not dates.isna().all() else date.today()
            
            # 快捷选择
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
            else:  # 全部
                start_date = min_date
                end_date = max_date
        else:
            start_date = None
            end_date = None
            st.warning("未找到日期列")
    
    # 供应商筛选
    with col2:
        st.markdown("**供应商**")
        supplier_col = col_map.get("供应商")
        if supplier_col and supplier_col in df.columns:
            suppliers = ["全部"] + sorted(df[supplier_col].dropna().astype(str).unique().tolist())
            supplier_filter = st.selectbox("选择供应商", suppliers)
        else:
            supplier_filter = "全部"
    
    # 采购员筛选
    with col3:
        st.markdown("**采购员**")
        buyer_col = col_map.get("采购员")
        if buyer_col and buyer_col in df.columns:
            buyers = ["全部"] + sorted(df[buyer_col].dropna().astype(str).unique().tolist())
            buyer_filter = st.selectbox("选择采购员", buyers)
        else:
            buyer_filter = "全部"
    
    col4, col5 = st.columns(2)
    
    # 商品名称筛选
    with col4:
        st.markdown("**商品名称**")
        product_col = col_map.get("商品名称")
        if product_col and product_col in df.columns:
            products = ["全部"] + sorted(df[product_col].dropna().astype(str).unique().tolist())
            product_filter = st.selectbox("选择商品", products)
        else:
            product_filter = "全部"
    
    # 发货地筛选
    with col5:
        st.markdown("**发货地**")
        origin_col = col_map.get("发货地")
        if origin_col and origin_col in df.columns:
            origins = ["全部"] + sorted(df[origin_col].dropna().astype(str).unique().tolist())
            origin_filter = st.selectbox("选择发货地", origins)
        else:
            origin_filter = "全部"
    
    # ===================== 应用筛选 =====================
    filtered_df = df.copy()
    
    # 日期筛选
    if start_date and end_date and date_col:
        dates = pd.to_datetime(filtered_df[date_col].apply(convert_excel_date), errors='coerce')
        filtered_df = filtered_df[
            (dates.dt.date >= start_date) & (dates.dt.date <= end_date)
        ]
    
    # 供应商筛选
    if supplier_filter != "全部" and supplier_col:
        filtered_df = filtered_df[filtered_df[supplier_col].astype(str) == supplier_filter]
    
    # 采购员筛选
    if buyer_filter != "全部" and buyer_col:
        filtered_df = filtered_df[filtered_df[buyer_col].astype(str) == buyer_filter]
    
    # 商品筛选
    if product_filter != "全部" and product_col:
        filtered_df = filtered_df[filtered_df[product_col].astype(str) == product_filter]
    
    # 发货地筛选
    if origin_filter != "全部" and origin_col:
        filtered_df = filtered_df[filtered_df[origin_col].astype(str) == origin_filter]
    
    # ===================== 统计信息 =====================
    st.markdown("---")
    st.markdown("### 📊 统计信息")
    
    stat_col1, stat_col2, stat_col3, stat_col4, stat_col5 = st.columns(5)
    
    amount_col = col_map.get("金额")
    qty_col = col_map.get("数量")
    weight_col = col_map.get("重量")
    
    with stat_col1:
        st.metric("筛选记录数", f"{len(filtered_df)} 条")
    
    with stat_col2:
        if amount_col and amount_col in filtered_df.columns:
            total_amount = pd.to_numeric(filtered_df[amount_col], errors='coerce').sum()
            st.metric("总金额", f"¥{total_amount:,.2f}")
        else:
            st.metric("总金额", "N/A")
    
    with stat_col3:
        if qty_col and qty_col in filtered_df.columns:
            total_qty = pd.to_numeric(filtered_df[qty_col], errors='coerce').sum()
            st.metric("总数量", f"{total_qty:,.0f}")
        else:
            st.metric("总数量", "N/A")
    
    with stat_col4:
        if weight_col and weight_col in filtered_df.columns:
            total_weight = pd.to_numeric(filtered_df[weight_col], errors='coerce').sum()
            st.metric("总重量", f"{total_weight:,.0f}")
        else:
            st.metric("总重量", "N/A")
    
    with stat_col5:
        if supplier_col and supplier_col in filtered_df.columns:
            supplier_count = filtered_df[supplier_col].nunique()
            st.metric("供应商数", f"{supplier_count} 家")
        else:
            st.metric("供应商数", "N/A")
    
    # ===================== 数据表格 =====================
    st.markdown("---")
    st.markdown("### 📋 数据明细")
    
    # 只显示指定列
    display_cols_actual = [col_map.get(c) for c in DISPLAY_COLS if col_map.get(c) in filtered_df.columns]
    
    if display_cols_actual:
        display_df = filtered_df[display_cols_actual].copy()
        
        # 转换日期列显示
        if date_col and date_col in display_df.columns:
            display_df[date_col] = display_df[date_col].apply(convert_excel_date)
        
        # 重命名列为标准名称
        rename_map = {v: k for k, v in col_map.items() if v in display_cols_actual}
        display_df = display_df.rename(columns=rename_map)
        
        # 显示表格
        st.dataframe(
            display_df,
            use_container_width=True,
            height=400,
            hide_index=True,
        )
        
        # 分页提示
        if len(display_df) > 1000:
            st.info(f"显示前 1000 条，共 {len(display_df)} 条")
    else:
        st.warning("没有可显示的列")
    
    # ===================== 导出功能 =====================
    st.markdown("---")
    st.markdown("### 📥 导出数据")
    
    col_exp1, col_exp2 = st.columns([1, 3])
    
    with col_exp1:
        export_format = st.radio("导出格式", ["Excel (.xlsx)", "CSV (.csv)"])
    
    with col_exp2:
        if st.button("📥 导出筛选结果", use_container_width=True):
            if export_format == "Excel (.xlsx)":
                # 导出 Excel
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    filtered_df.to_excel(writer, index=False, sheet_name='入库单')
                buffer.seek(0)
                
                st.download_button(
                    label="⬇️ 下载 Excel 文件",
                    data=buffer,
                    file_name=f"入库单筛选结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                # 导出 CSV
                csv = filtered_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="⬇️ 下载 CSV 文件",
                    data=csv,
                    file_name=f"入库单筛选结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                )

# ===================== 主程序 =====================
def main():
    """主程序入口"""
    # 设置页面配置
    st.set_page_config(
        page_title="入库单查询系统",
        page_icon="📦",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    
    # 初始化 session state
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False
        st.session_state.username = ""
        st.session_state.role = ""
    
    # 根据登录状态显示不同页面
    if st.session_state.logged_in:
        main_app()
    else:
        login_page()

if __name__ == "__main__":
    main()
