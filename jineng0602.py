import os
import time
from datetime import datetime
from typing import List, Tuple
import pandas as pd
import streamlit as st
from streamlit_autorefresh import st_autorefresh
from streamlit_echarts import st_echarts
import plotly.graph_objects as go

# ==================== 页面基础配置 ====================
st.set_page_config(page_title="技能覆盖分析大屏", layout="wide", initial_sidebar_state="expanded")

# 全局样式优化（适配Streamlit新版，无兼容性问题）
PAGE_CSS = """
<style>
body, [data-testid="stAppViewContainer"]{
    background-color: #e6f7ff !important;
    color: #003366 !important;
}
[data-testid="stSidebar"]{
    background-color: #d1e7f5 !important;
    color: #003366 !important;
}
div.stButton>button{
    background-color: #4cc9f0 !important;
    color: #000000 !important;
    border-radius:10px;
    height:40px;
    font-weight:700;
    margin:5px 0;
    width:100%;
}
div.stButton>button:hover{
    background-color:#4895ef !important;
    color:#ffffff !important;
}
.metric-card{
    background-color: #ffffff !important;
    padding:20px;
    border-radius:16px;
    text-align:center;
    box-shadow:0 0 15px rgba(0,0,0,0.08);
}
.metric-value{
    font-size:36px;
    font-weight:800;
    color: #0066cc !important;
}
.metric-label{
    font-size:14px;
    color: #336699 !important;
}
hr{
    border:none;
    border-top:1px solid #bbd9f7;
    margin:16px 0;
}
.heatmap-container {
    max-height: 700px;
    overflow-y: auto;
    overflow-x: auto;
    border-radius: 8px;
    background-color: #ffffff;
}
.heatmap-container::-webkit-scrollbar {
    width: 8px;
    height: 8px;
}
.heatmap-container::-webkit-scrollbar-thumb {
    background-color: #99c2ff;
    border-radius: 4px;
}
.heatmap-container::-webkit-scrollbar-track {
    background-color: #e6f7ff;
}
</style>
"""
st.markdown(PAGE_CSS, unsafe_allow_html=True)

# ==================== 核心配置与工具函数 ====================
# 全局文件路径（适配Streamlit Cloud容器环境）
DEFAULT_FILE_NAME = "jixiao.xlsx"
# 全局配色池
COLOR_POOL = [
    "#FF3333", "#33FF33", "#3333FF", "#FFAA00", "#9933FF",
    "#00FFFF", "#FF99CC", "#FFFF33", "#008080", "#FF00FF",
    "#8B4513", "#20B2AA", "#FF6347", "#9370DB", "#32CD32"
]

# Excel写入工具函数
def get_excel_writer(file_path: str, mode: str = "w") -> pd.ExcelWriter:
    if mode == "a" and os.path.exists(file_path):
        return pd.ExcelWriter(file_path, mode="a", if_sheet_exists="replace", engine="openpyxl")
    return pd.ExcelWriter(file_path, engine="openpyxl")

# 分数总和计算函数
def calc_score_sum(df: pd.DataFrame, score_col: str) -> pd.DataFrame:
    if score_col not in df.columns or "明细" not in df.columns:
        return df
    sum_col_name = f"{score_col}_数量总和"
    if sum_col_name in df.columns:
        df = df.drop(columns=[sum_col_name])
    sum_df = df.groupby("明细", as_index=False)[score_col].sum().rename(columns={score_col: sum_col_name})
    df = df.merge(sum_df, on="明细", how="left")
    return df

def calc_all_sum(df: pd.DataFrame) -> pd.DataFrame:
    df = calc_score_sum(df, "自评值")
    df = calc_score_sum(df, "互评值")
    return df

# ==================== 数据加载核心函数（适配Streamlit Cloud） ====================
@st.cache_data(ttl=300, show_spinner="正在加载数据...")
def load_sheets(file_path: str) -> Tuple[List[str], dict]:
    if not os.path.exists(file_path):
        return [], {}
    try:
        xpd = pd.ExcelFile(file_path)
        frames = {}
        required_cols = {"明细", "员工", "自评值", "互评值"}
        for s in xpd.sheet_names:
            try:
                df0 = pd.read_excel(xpd, sheet_name=s)
                if df0.empty:
                    continue
                df0 = df0.fillna("")
                if not required_cols.issubset(df0.columns):
                    st.sidebar.warning(f"表 {s} 缺少必要列，已跳过。")
                    continue
                # 处理分组表头格式
                if df0.iloc[0, 0] == "分组":
                    groups = df0.iloc[0, 1:].tolist()
                    df0 = df0.drop(0).reset_index(drop=True)
                    emp_cols = [c for c in df0.columns if c not in ["明细", "自评值_数量总和", "互评值_数量总和", "编号"]]
                    group_map = {emp: groups[i] if i < len(groups) else "默认分组" for i, emp in enumerate(emp_cols)}
                    df_long = df0.melt(
                        id_vars=["明细"],
                        value_vars=emp_cols,
                        var_name="员工",
                        value_name="临时值"
                    )
                    df_long["分组"] = df_long["员工"].map(group_map)
                    df_long["自评值"] = pd.to_numeric(df_long["临时值"], errors="coerce").fillna(0)
                    df_long["互评值"] = pd.to_numeric(df_long["临时值"], errors="coerce").fillna(0)
                    df_long = df_long.drop(columns=["临时值"], errors="ignore")
                    frames[s] = df_long
                else:
                    if "分组" not in df0.columns:
                        df0["分组"] = "默认分组"
                    df0["自评值"] = pd.to_numeric(df0["自评值"], errors="coerce").fillna(0)
                    df0["互评值"] = pd.to_numeric(df0["互评值"], errors="coerce").fillna(0)
                    frames[s] = df0
            except Exception as e:
                st.sidebar.error(f"读取 {s} 失败: {str(e)}")
        return xpd.sheet_names, frames
    except Exception as e:
        st.sidebar.error(f"文件读取失败: {str(e)}")
        return [], {}

# ==================== 页面主体逻辑 ====================
def main():
    # 1. 侧边栏：文件上传与路径配置
    st.sidebar.markdown("## 📁 数据文件配置")
    # 支持用户上传Excel文件（Streamlit Cloud核心适配）
    uploaded_file = st.sidebar.file_uploader("上传 jixiao.xlsx 文件", type="xlsx", key="file_upload")
    
    # 确定最终文件路径
    if uploaded_file is not None:
        # 保存上传的文件到临时目录
        SAVE_FILE = f"/tmp/{uploaded_file.name}"
        with open(SAVE_FILE, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.sidebar.success(f"✅ 已上传文件: {uploaded_file.name}")
    else:
        # 尝试读取默认路径的文件
        possible_paths = [
            f"./{DEFAULT_FILE_NAME}",
            f"/mount/src/{DEFAULT_FILE_NAME}",
            f"/mount/src/guibit/{DEFAULT_FILE_NAME}"
        ]
        SAVE_FILE = None
        for path in possible_paths:
            if os.path.exists(path):
                SAVE_FILE = path
                break
        if SAVE_FILE is None:
            SAVE_FILE = f"/tmp/{DEFAULT_FILE_NAME}"
            st.sidebar.warning("⚠️ 未找到默认文件，请上传Excel文件")

    # 2. 加载数据
    sheets, sheet_frames = load_sheets(SAVE_FILE)
    
    # 3. 自动修复数据总和列
    if sheet_frames:
        repaired_count = 0
        repaired_frames = {}
        for sheet_name, df0 in sheet_frames.items():
            df_new = calc_all_sum(df0)
            if not df0.equals(df_new):
                repaired_count += 1
                repaired_frames[sheet_name] = df_new
        if repaired_frames:
            with get_excel_writer(SAVE_FILE, mode="w") as writer:
                for sn, df0 in sheet_frames.items():
                    if sn in repaired_frames:
                        repaired_frames[sn].to_excel(writer, sheet_name=sn, index=False)
                        sheet_frames[sn] = repaired_frames[sn]
                    else:
                        df0.to_excel(writer, sheet_name=sn, index=False)
            st.cache_data.clear()
            st.sidebar.info(f"🔧 自动修复 {repaired_count} 张表的数量总和")

    # 4. 侧边栏：新增时间点
    st.sidebar.markdown("---")
    st.sidebar.markdown("### ➕ 新增数据时间点")
    current_year = datetime.now().year
    year = st.sidebar.selectbox("选择年份", list(range(current_year - 2, current_year + 2)), index=2)
    mode = st.sidebar.radio("时间类型", ["月份", "季度"], horizontal=True)
    if mode == "月份":
        month = st.sidebar.selectbox("选择月份", list(range(1, 13)))
        new_sheet_name = f"{year}_{month:02d}"
    else:
        quarter = st.sidebar.selectbox("选择季度", ["Q1", "Q2", "Q3", "Q4"])
        new_sheet_name = f"{year}_{quarter}"
    
    if st.sidebar.button("创建新的时间点", use_container_width=True):
        if new_sheet_name in sheets:
            st.sidebar.error(f"❌ 时间点 {new_sheet_name} 已存在！")
        else:
            try:
                base_df = pd.DataFrame(columns=["明细", "自评值_数量总和", "互评值_数量总和", "员工", "自评值", "互评值", "分组"])
                # 继承上期数据
                prev_sheets = sorted([s for s in sheets if s.split("_")[0] == str(year) and s < new_sheet_name])
                if not prev_sheets:
                    prev_years = sorted([int(s.split("_")[0]) for s in sheets if s.split("_")[0].isdigit()])
                    if prev_years:
                        latest_prev_year = max(y for y in prev_years if y < year) if any(y < year for y in prev_years) else None
                        if latest_prev_year:
                            prev_sheets = sorted([s for s in sheets if s.startswith(str(latest_prev_year))])
                if prev_sheets:
                    prev_name = prev_sheets[-1]
                    base_df = sheet_frames.get(prev_name, base_df).copy()
                    st.sidebar.info(f"📋 继承上期数据: {prev_name}")
                else:
                    st.sidebar.info("📝 无上期数据，创建空白模板")
                # 写入文件
                with get_excel_writer(SAVE_FILE, mode="a") as writer:
                    base_df.to_excel(writer, sheet_name=new_sheet_name, index=False)
                st.cache_data.clear()
                st.sidebar.success(f"✅ 创建成功: {new_sheet_name}")
                # 重新加载数据
                sheets, sheet_frames = load_sheets(SAVE_FILE)
            except Exception as e:
                st.sidebar.error(f"❌ 创建失败: {str(e)}")

    # 5. 侧边栏：数据修复工具
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔧 数据修复工具")
    if st.sidebar.button("一键更新所有表总和", use_container_width=True):
        try:
            if not os.path.exists(SAVE_FILE):
                st.sidebar.warning("⚠️ 未找到文件")
            else:
                xls = pd.ExcelFile(SAVE_FILE)
                updated_frames = {}
                for sheet_name in xls.sheet_names:
                    df0 = pd.read_excel(xls, sheet_name=sheet_name)
                    df0 = calc_all_sum(df0)
                    updated_frames[sheet_name] = df0
                with get_excel_writer(SAVE_FILE, mode="w") as writer:
                    for sn, df0 in updated_frames.items():
                        df0.to_excel(writer, sheet_name=sn, index=False)
                st.cache_data.clear()
                st.sidebar.success("✅ 所有工作表总和已更新！")
                # 重新加载数据
                sheets, sheet_frames = load_sheets(SAVE_FILE)
        except Exception as e:
            st.sidebar.error(f"❌ 更新失败: {str(e)}")

    # 6. 侧边栏：筛选器
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔍 数据筛选")
    all_time_list = sheets
    time_choice = st.sidebar.multiselect("选择月份/季度（支持跨年份）", all_time_list, default=all_time_list[:1] if all_time_list else [])
    
    # 分组筛选
    all_groups = []
    if sheet_frames:
        all_df_concat = pd.concat(sheet_frames.values())
        all_groups = all_df_concat["分组"].dropna().unique().tolist()
    selected_groups = st.sidebar.multiselect("选择分组", all_groups, default=all_groups)
    
    # 分数维度
    score_dimension = st.sidebar.radio(
        "分数维度",
        ["自评分数", "互评分数", "双维度对比"],
        horizontal=True,
        index=2
    )
    
    # 视图选择
    view = st.sidebar.radio(
        "切换视图",
        ["编辑数据", "大屏轮播", "单页模式", "能力分析", "基础子弹图", "高级子弹图"]
    )

    # 7. 数据合并
    def get_merged_df(keys: List[str], groups: List[str]) -> pd.DataFrame:
        dfs = []
        for k in keys:
            df0 = sheet_frames.get(k)
            if df0 is None:
                continue
            if groups and "分组" in df0.columns:
                df0 = df0[df0["分组"].isin(groups)]
            df0["自评值"] = pd.to_numeric(df0["自评值"], errors="coerce").fillna(0)
            df0["互评值"] = pd.to_numeric(df0["互评值"], errors="coerce").fillna(0)
            dfs.append(df0)
        if not dfs:
            return pd.DataFrame()
        merged_df = pd.concat(dfs, axis=0, ignore_index=True)
        merged_df = merged_df[merged_df["明细"].notna() & (merged_df["明细"] != "") & (merged_df["明细"] != "分数总和")]
        return merged_df

    df = get_merged_df(time_choice, selected_groups)

    # 8. 页面主体内容
    st.title("📊 技能覆盖分析大屏")
    st.markdown("---")

    # 空数据兜底
    if df.empty:
        st.warning("⚠️ 当前无可用数据，请检查：1. 已上传正确的Excel文件；2. 已选择正确的时间/分组筛选条件")
        return

    # 核心指标卡片
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        total_emp = df["员工"].nunique()
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{total_emp}</div>
            <div class="metric-label">参与人数</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        total_task = df["明细"].nunique()
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{total_task}</div>
            <div class="metric-label">技能项总数</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        total_self = df["自评值"].sum()
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{int(total_self)}</div>
            <div class="metric-label">自评总分数</div>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        total_peer = df["互评值"].sum()
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{int(total_peer)}</div>
            <div class="metric-label">互评总分数</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")

    # 图表公共函数
    def get_score_cols() -> Tuple[str, str]:
        if score_dimension == "自评分数":
            return "自评值", "自评分数"
        elif score_dimension == "互评分数":
            return "互评值", "互评分数"
        else:
            return "自评值", "互评值"

    # 1. 人员排名柱状图
    st.subheader("📈 人员分数排名")
    def chart_total(df0: pd.DataFrame):
        s1, s2 = get_score_cols()
        fig = go.Figure()
        if score_dimension == "双维度对比":
            emp_stats = df0.groupby("员工").agg({"自评值":"sum","互评值":"sum"}).reset_index()
            emp_stats = emp_stats.sort_values("自评值", ascending=False)
            fig.add_trace(go.Bar(x=emp_stats["员工"], y=emp_stats["自评值"], name="自评", marker_color="#4cc9f0"))
            fig.add_trace(go.Bar(x=emp_stats["员工"], y=emp_stats["互评值"], name="互评", marker_color="#f72585"))
            fig.update_layout(barmode="group", xaxis_title="员工", yaxis_title="总分")
        else:
            emp_stats = df0.groupby("员工")[s1].sum().sort_values(ascending=False).reset_index()
            fig.add_trace(go.Bar(x=emp_stats["员工"], y=emp_stats[s1], name=s2, marker_color="#4cc9f0"))
            fig.update_layout(xaxis_title="员工", yaxis_title=s2)
        fig.update_layout(
            template="plotly_white",
            legend=dict(orientation="h", y=-0.2),
            height=500,
            font=dict(color="#003366")
        )
        return fig

    st.plotly_chart(chart_total(df), use_container_width=True)
    st.markdown("---")

    # 2. 任务对比堆叠柱状图
    st.subheader("📋 各技能项分数分布")
    def chart_stack(df0: pd.DataFrame):
        fig = go.Figure()
        agg_df = df0.groupby(["明细", "员工"])[["自评值", "互评值"]].sum().reset_index()
        if score_dimension == "双维度对比":
            for emp in agg_df["员工"].unique():
                sub = agg_df[agg_df["员工"] == emp]
                fig.add_trace(go.Bar(x=sub["明细"], y=sub["互评值"], name=f"互评-{emp}", marker_color="#f72585", opacity=0.7))
                fig.add_trace(go.Bar(x=sub["明细"], y=sub["自评值"], name=f"自评-{emp}", marker_color="#4cc9f0", opacity=0.8))
        else:
            col, name_text = get_score_cols()
            for emp in agg_df["员工"].unique():
                sub = agg_df[agg_df["员工"] == emp]
                fig.add_trace(go.Bar(x=sub["明细"], y=sub[col], name=emp))
        fig.update_layout(
            barmode="stack",
            template="plotly_white",
            xaxis_title="技能项",
            yaxis_title="分数",
            legend=dict(orientation="h", y=-0.2),
            height=600,
            font=dict(color="#003366")
        )
        return fig

    st.plotly_chart(chart_stack(df), use_container_width=True)
    st.markdown("---")

    # 3. 热力图（人员-技能项分数）
    st.subheader("🔥 人员-技能项分数热力图")
    def chart_heat(df0: pd.DataFrame):
        task_list = df0["明细"].dropna().unique().tolist()
        user_list = df0["员工"].dropna().unique().tolist()
        if len(task_list) == 0 or len(user_list) == 0:
            return {
                "title": {"text": "暂无有效数据", "left": "center", "textStyle": {"color": "#333333"}},
                "backgroundColor": "#ffffff"
            }
        # 数据透视
        pivot_df = df0.pivot_table(index="员工", columns="明细", values="自评值" if score_dimension == "自评分数" else "互评值", aggfunc="sum", fill_value=0)
        # 转换为ECharts格式
        data = []
        for i, user in enumerate(pivot_df.index):
            for j, task in enumerate(pivot_df.columns):
                data.append([j, i, pivot_df.iloc[i, j]])
        # ECharts配置
        option = {
            "title": {
                "text": "人员-技能项分数热力图",
                "left": "center",
                "textStyle": {"color": "#003366", "fontSize": 20}
            },
            "tooltip": {
                "position": "top",
                "formatter": "{b}<br/>{a}: {c}分"
            },
            "grid": {
                "height": "70%",
                "top": "10%"
            },
            "xAxis": {
                "type": "category",
                "data": pivot_df.columns.tolist(),
                "splitArea": {"show": True},
                "axisLabel": {
                    "interval": 0,
                    "rotate": -45,
                    "color": "#003366"
                }
            },
            "yAxis": {
                "type": "category",
                "data": pivot_df.index.tolist(),
                "splitArea": {"show": True},
                "axisLabel": {"color": "#003366"}
            },
            "visualMap": {
                "min": 0,
                "max": df0["自评值" if score_dimension == "自评分数" else "互评值"].max(),
                "calculable": True,
                "orient": "horizontal",
                "left": "center",
                "bottom": "0%",
                "inRange": {
                    "color": ["#e6f7ff", "#99c2ff", "#4cc9f0", "#0066cc"]
                }
            },
            "series": [
                {
                    "name": "分数",
                    "type": "heatmap",
                    "data": data,
                    "label": {
                        "show": True,
                        "color": "#000000"
                    }
                }
            ],
            "backgroundColor": "#ffffff"
        }
        return option

    # 热力图容器
    with st.container():
        st.markdown('<div class="heatmap-container">', unsafe_allow_html=True)
        st_echarts(chart_heat(df), height=700)
        st.markdown('</div>', unsafe_allow_html=True)

# 运行主函数
if __name__ == "__main__":
    main()
