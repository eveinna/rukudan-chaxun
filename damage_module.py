# -*- coding: utf-8 -*-
"""
报损数据模块 - 独立组件（可被 Streamlit 主应用导入）
维护此文件即可更新线上报损功能，无需动主程序

用法：from damage_module import damage_report_page, DAMAGE_SECTIONS
"""

import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io

# ===================== 板块配置 =====================
# 在此处增删板块名称和对应文件名
DAMAGE_SECTIONS = {
    "MMT板块": "报损-MMT板块.xlsx",
    "W板块": "报损-W板块.xlsx",
    "CH板块": "报损-CH板块.xlsx",
    "XG板块": "报损-XG板块.xlsx",
}

# 显示列配置
DISPLAY_COLS_DAMAGE = [
    "板块名称", "序号", "售后日期", "供应商名称", "配送日期", "门店名称",
    "商品名称", "下单件数", "破损重量", "金额报损",
    "当日到货报损", "榴莲三天报损", "报损原因", "处理方式", "凭证图片", "备注"
]


def clean_cols(df):
    """清理列名中的换行符"""
    df = df.copy()
    df.columns = [str(c).replace('\n', '').strip() for c in df.columns]
    return df


def convert_excel_date(val):
    """Excel日期转换（序列号/datetime → 字符串）"""
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
    """
    根据快捷选项计算日期范围（与桌面版 Tkinter 完全一致）
    
    参数:
        quick_mode: 快捷选择值
        min_date: 数据最小日期（用于"全部"模式）
        max_date: 数据最大日期（用于"全部"模式）
    
    返回:
        (start_date, end_date) 元组
    """
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
    else:  # 自定义
        s = None
        e = None

    return s, e


# ===================== 报损数据主页面 =====================
def damage_report_page():
    """Streamlit 报损数据管理页面"""
    st.title("📉 报损数据管理")

    # ===== 文件上传区域 =====
    st.markdown("### 📂 上传板块数据文件")
    st.info("请依次上传各板块的 Excel 文件（表头格式需一致），系统将自动合并并标注来源板块。")

    # 4个板块的上传框
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

    # 合并按钮
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

    # ===== 统计概览 =====
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
        section_counts = df['板块名称'].value_counts().to_dict()
        st.metric("板块数量", f"{len(section_counts)} 个")

    # 各板块统计
    st.markdown("#### 各板块明细")
    section_stats = df.groupby('板块名称').agg(
        记录数=('序号', 'count'),
        总金额=('金额报损', 'sum'),
        总重量=('破损重量', 'sum')
    ).reset_index()
    section_stats['总金额'] = section_stats['总金额'].apply(lambda x: f"¥{x:,.2f}")
    section_stats['总重量'] = section_stats['总重量'].apply(lambda x: f"{x:,.1f} kg")
    section_stats.columns = ['板块名称', '记录数', '总报损金额', '总破损重量']
    st.dataframe(section_stats, use_container_width=True, hide_index=True)

    # ===== 筛选区域 =====
    st.markdown("---")
    st.markdown("### 🔍 筛选条件")

    filter_cols = st.columns(4)

    # 板块筛选
    with filter_cols[0]:
        st.markdown("**板块**")
        sections = ["全部"] + list(DAMAGE_SECTIONS.keys())
        section_filter = st.selectbox("选择板块", sections, key="filter_section")

    # 日期筛选（与桌面版完全一致的快捷逻辑）
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
                # 显示当前范围（只读提示）
                if s_date and e_date:
                    st.caption(f"📅 {s_date} ~ {e_date}")
        except Exception as ex:
            s_date = None
            e_date = None
            st.error(f"日期处理异常：{ex}")

    # 门店筛选
    with filter_cols[2]:
        st.markdown("**门店名称**")
        stores = ["全部"] + sorted(df['门店名称'].dropna().astype(str).unique().tolist())
        store_filter = st.selectbox("选择门店", stores, key="filter_store")

    # 商品筛选
    with filter_cols[3]:
        st.markdown("**商品名称**")
        products = ["全部"] + sorted(df['商品名称'].dropna().astype(str).unique().tolist())
        product_filter = st.selectbox("选择商品", products, key="filter_product")

    # 第二行：报损原因 | 处理方式
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
    st.dataframe(
        filtered[display_cols],
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
