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
st.set_page_config(page_title="技能覆盖分析大屏", layout="wide")

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


# -------------------- GUIbit数据读取函数 --------------------
def load_data_from_gui():
    """从GUIbit目录读取jixiao.xlsx文件"""
    try:
        # 定义GUIbit目录路径 - 根据你的项目结构调整
        # 尝试多种可能的路径
        possible_paths = [
            "./guibit/jixiao.xlsx",  # 当前目录下的guibit文件夹
            "./jixiao.xlsx",  # 当前目录下
            "../guibit/jixiao.xlsx",  # 上级目录下的guibit文件夹
            "jixiao.xlsx",  # 当前目录下
        ]

        file_path = None
        for path in possible_paths:
            if os.path.exists(path):
                file_path = path
                break

        if not file_path:
            st.sidebar.error("❌ 未找到jixiao.xlsx文件")
            st.sidebar.info("请确保jixiao.xlsx文件在以下任一位置：")
            for path in possible_paths:
                st.sidebar.info(f"  • {path}")
            return [], {}, "文件不存在"

        st.sidebar.info(f"🔄 正在从 {file_path} 读取数据...")
        return file_path

    except Exception as e:
        st.sidebar.error(f"读取路径异常：{str(e)}")
        return None

# 全局变量：文件路径
SAVE_FILE = load_data_from_gui()
if SAVE_FILE is None:
    SAVE_FILE = "jixiao.xlsx"


# 全局配色池（多颜色，区分不同时间点）
COLOR_POOL = [
    "#FF3333", "#33FF33", "#3333FF", "#FFAA00", "#9933FF",
    "#00FFFF", "#FF99CC", "#FFFF33", "#008080", "#FF00FF",
    "#8B4513", "#20B2AA", "#FF6347", "#9370DB", "#32CD32"
]

# ==================== 工具函数 ====================
def get_excel_writer(file_path: str, mode: str = "w") -> pd.ExcelWriter:
    if mode == "a" and os.path.exists(file_path):
        return pd.ExcelWriter(file_path, mode="a", if_sheet_exists="replace", engine="openpyxl")
    return pd.ExcelWriter(file_path, engine="openpyxl")

def calc_score_sum(df: pd.DataFrame, score_col: str) -> pd.DataFrame:
    """统一计算单维度分数总和"""
    if score_col not in df.columns or "明细" not in df.columns:
        return df
    sum_col_name = f"{score_col}_数量总和"
    if sum_col_name in df.columns:
        df = df.drop(columns=[sum_col_name])
    sum_df = df.groupby("明细", as_index=False)[score_col].sum().rename(columns={score_col: sum_col_name})
    df = df.merge(sum_df, on="明细", how="left")
    return df

def calc_all_sum(df: pd.DataFrame) -> pd.DataFrame:
    """一次性计算自评+互评两个总和"""
    df = calc_score_sum(df, "自评值")
    df = calc_score_sum(df, "互评值")
    return df

# ==================== 数据加载 ====================
@st.cache_data(ttl=300)
def load_sheets(file: str, ts=None) -> Tuple[List[str], dict]:
    if not os.path.exists(file):
        return [], {}
    xpd = pd.ExcelFile(file)
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

# 初始化数据
sheets, sheet_frames = [], {}
try:
    mtime = os.path.getmtime(SAVE_FILE) if os.path.exists(SAVE_FILE) else None
    sheets, sheet_frames = load_sheets(SAVE_FILE, ts=mtime)
    st.sidebar.success(f"已加载文件: {SAVE_FILE}")

    # 自动修复总和列
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
        st.sidebar.info(f"自动修复 {repaired_count} 张表的数量总和")

except Exception as e:
    st.sidebar.warning(f"读取文件失败: {str(e)}")
    # 示例测试数据
    sheet_frames = {
        "2025_01": pd.DataFrame({
            "明细": ["任务A", "任务B", "任务C", "任务A", "任务B", "任务C"],
            "自评值_数量总和": [3, 2, 5, 3, 2, 5],
            "互评值_数量总和": [4, 2, 4, 4, 2, 4],
            "员工": ["张三", "李四", "王五", "李四", "王五", "张三"],
            "自评值": [1, 1, 1, 2, 1, 2],
            "互评值": [2, 1, 1, 2, 1, 1],
            "分组": ["A8", "B7", "VN", "A8", "B7", "VN"]
        }),
        "2026_02": pd.DataFrame({
            "明细": ["任务A", "任务B", "任务C"],
            "自评值_数量总和": [4, 3, 6],
            "互评值_数量总和": [3, 4, 5],
            "员工": ["张三", "李四", "王五"],
            "自评值": [4, 3, 6],
            "互评值": [3, 4, 5],
            "分组": ["A8", "B7", "VN"]
        })
    }
    sheets = ["2025_01", "2026_02"]

# ==================== 侧边栏 - 新增时间点 ====================
st.sidebar.markdown("### 新增数据时间点")
current_year = datetime.now().year
year = st.sidebar.selectbox("选择年份", list(range(current_year - 2, current_year + 2)), index=2)
mode = st.sidebar.radio("时间类型", ["月份", "季度"], horizontal=True)

if mode == "月份":
    month = st.sidebar.selectbox("选择月份", list(range(1, 13)))
    new_sheet_name = f"{year}_{month:02d}"
else:
    quarter = st.sidebar.selectbox("选择季度", ["Q1", "Q2", "Q3", "Q4"])
    new_sheet_name = f"{year}_{quarter}"

if st.sidebar.button("创建新的时间点"):
    if new_sheet_name in sheets:
        st.sidebar.error(f"时间点 {new_sheet_name} 已存在！")
    else:
        try:
            base_df = pd.DataFrame(columns=["明细", "自评值_数量总和", "互评值_数量总和", "员工", "自评值", "互评值", "分组"])
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
                st.sidebar.info(f"继承上期数据: {prev_name}")
            else:
                st.sidebar.info("无上期数据，创建空白模板")

            with get_excel_writer(SAVE_FILE, mode="a") as writer:
                base_df.to_excel(writer, sheet_name=new_sheet_name, index=False)
            st.cache_data.clear()
            st.sidebar.success(f"创建成功: {new_sheet_name}")
        except Exception as e:
            st.sidebar.error(f"创建失败: {str(e)}")

# ==================== 侧边栏 - 全局数据修复 ====================
st.sidebar.markdown("### 数据修复工具")
if st.sidebar.button("一键更新所有表总和"):
    try:
        if not os.path.exists(SAVE_FILE):
            st.sidebar.warning("未找到 jixiao.xlsx")
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
            st.sidebar.success("所有工作表总和已更新！")
    except Exception as e:
        st.sidebar.error(f"更新失败: {str(e)}")

# ==================== 侧边栏 - 筛选器 ====================
all_time_list = sheets
time_choice = st.sidebar.multiselect("选择月份/季度（支持跨年份）", all_time_list, default=all_time_list[:1])

all_groups = []
if sheet_frames:
    all_df_concat = pd.concat(sheet_frames.values())
    all_groups = all_df_concat["分组"].dropna().unique().tolist()
selected_groups = st.sidebar.multiselect("选择分组", all_groups, default=all_groups)

# 分数维度（全局变量，所有图表共用）
score_dimension = st.sidebar.radio(
    "分数维度",
    ["自评分数", "互评分数", "双维度对比"],
    horizontal=True,
    index=2
)

# 视图选择
view = st.sidebar.radio(
    "切换视图",
    ["编辑数据", "大屏轮播", "单页模式", "显示所有视图", "能力分析", "基础子弹图", "高级子弹图"]
)

# ==================== 数据合并函数 ====================
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
        st.warning("当前无可用数据，请重新选择时间/分组")
        return pd.DataFrame()
    merged_df = pd.concat(dfs, axis=0, ignore_index=True)
    merged_df = merged_df[merged_df["明细"].notna() & (merged_df["明细"] != "") & (merged_df["明细"] != "分数总和")]
    return merged_df

df = get_merged_df(time_choice, selected_groups)

# ==================== 图表公共函数 ====================
def get_score_cols() -> Tuple[str, str]:
    if score_dimension == "自评分数":
        return "自评值", "自评分数"
    elif score_dimension == "互评分数":
        return "互评值", "互评分数"
    else:
        return "自评值", "互评值"

# 1. 人员排名柱状图
def chart_total(df0: pd.DataFrame):
    if df0.empty:
        return go.Figure()
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
        fig.add_trace(go.Bar(x=emp_stats["员工"], y=emp_stats[s1], name=s2))
        fig.update_layout(xaxis_title="员工", yaxis_title=s2)
    fig.update_layout(template="plotly_dark", legend=dict(orientation="h", y=-0.2))
    return fig

# 2. 任务对比堆叠柱状图
def chart_stack(df0: pd.DataFrame):
    if df0.empty:
        return go.Figure()
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
        template="plotly_dark",
        xaxis_title="任务",
        yaxis_title="分数",
        legend=dict(orientation="h", y=-0.2)
    )
    return fig

# ===================== 热力图函数（已改为白底+深色文字） =====================
def chart_heat(df0: pd.DataFrame):
    # 全局空数据拦截
    if df0.empty:
        return {
            "title": {"text": "暂无有效数据", "left": "center", "textStyle": {"color": "#333333"}},
            "backgroundColor": "#ffffff"
        }

    # 提取维度并去重、清洗
    task_list = df0["明细"].dropna().unique().tolist()
    user_list = df0["员工"].dropna().unique().tolist()

    # 维度为空拦截
    if len(task_list) == 0 or len(user_list) == 0:
        return {
            "title": {"text": "任务/人员数据为空，无法生成热力图", "left": "center", "textStyle": {"color": "#333333"}},
            "backgroundColor": "#ffffff"
        }

    # 数据透视聚合，强制填充0，规避索引异常
    try:
        pivot_self = df0.groupby(["明细", "员工"])["自评值"].sum().unstack(fill_value=0)
        pivot_peer = df0.groupby(["明细", "员工"])["互评值"].sum().unstack(fill_value=0)
    except Exception:
        return {
            "title": {"text": "数据格式异常，生成失败", "left": "center", "textStyle": {"color": "#333333"}},
            "backgroundColor": "#ffffff"
        }

    data = []
    title_text = ""
    color_list = []
    min_val = 0
    max_val = 0

    # 分支1：自评分数
    if score_dimension == "自评分数":
        title_text = "自评分数 热力图"
        color_list = ["#e8f4f8", "#4cc9f0"]
        all_vals = []
        for y_idx, task in enumerate(task_list):
            for x_idx, user in enumerate(user_list):
                val = float(pivot_self.loc[task, user]) if task in pivot_self.index and user in pivot_self.columns else 0
                data.append([x_idx, y_idx, val])
                all_vals.append(val)
        min_val = min(all_vals) if all_vals else 0
        max_val = max(all_vals) if all_vals else 0

    # 分支2：互评分数
    elif score_dimension == "互评分数":
        title_text = "互评分数 热力图"
        color_list = ["#fff0f3", "#f72585"]
        all_vals = []
        for y_idx, task in enumerate(task_list):
            for x_idx, user in enumerate(user_list):
                val = float(pivot_peer.loc[task, user]) if task in pivot_peer.index and user in pivot_peer.columns else 0
                data.append([x_idx, y_idx, val])
                all_vals.append(val)
        min_val = min(all_vals) if all_vals else 0
        max_val = max(all_vals) if all_vals else 0

    # 分支3：双维度差值
    else:
        title_text = "自评-互评 分数差值热力图"
        color_list = ["#f72585", "#ffffff", "#4cc9f0"]
        all_diff = []
        for y_idx, task in enumerate(task_list):
            for x_idx, user in enumerate(user_list):
                s = float(pivot_self.loc[task, user]) if task in pivot_self.index and user in pivot_self.columns else 0
                p = float(pivot_peer.loc[task, user]) if task in pivot_peer.index and user in pivot_peer.columns else 0
                diff = round(s - p, 1)
                data.append([x_idx, y_idx, diff])
                all_diff.append(diff)
        min_val = min(all_diff) if all_diff else -1
        max_val = max(all_diff) if all_diff else 1

    # 兜底极值
    if min_val == max_val:
        min_val -= 1
        max_val += 1

    # ECharts 配置：白底 + 深色文字/坐标轴
    option = {
        "backgroundColor": "#ffffff",
        "title": {
            "text": title_text,
            "left": "center",
            "textStyle": {"color": "#333333", "fontSize": 16}
        },
        "tooltip": {
            "trigger": "item",
            "formatter": "人员：{b}<br/>任务：{a}<br/>数值：{c}"
        },
        "grid": {"left": "3%", "right": "3%", "top": "12%", "bottom": "18%", "containLabel": True},
        "xAxis": {
            "type": "category",
            "data": user_list,
            "axisLabel": {"color": "#333333", "rotate": 45, "fontSize": 11},
            "axisLine": {"lineStyle": {"color": "#999999"}}
        },
        "yAxis": {
            "type": "category",
            "data": task_list,
            "axisLabel": {"color": "#333333", "fontSize": 11},
            "axisLine": {"lineStyle": {"color": "#999999"}}
        },
        "visualMap": {
            "min": min_val,
            "max": max_val,
            "show": True,
            "orient": "horizontal",
            "left": "center",
            "bottom": "8%",
            "inRange": {"color": color_list},
            "textStyle": {"color": "#333333"}
        },
        "series": [{
            "name": "分数",
            "type": "heatmap",
            "data": data,
            "label": {"show": True, "color": "#000000", "fontSize": 10},
            "itemStyle": {"borderColor": "#eeeeee", "borderWidth": 1},
            "emphasis": {"itemStyle": {"shadowBlur": 8}}
        }]
    }
    return option

# ===================== 子弹图 =====================
def chart_bullet_base(df0: pd.DataFrame, dim: str = "员工"):
    if df0.empty:
        return go.Figure()
    if dim == "员工":
        agg = df0.groupby("员工").agg({"自评值":"sum","互评值":"sum"}).reset_index()
        cat_col = "员工"
        title = "员工自评/互评对比"
    else:
        agg = df0.groupby("明细").agg({"自评值":"sum","互评值":"sum"}).reset_index()
        cat_col = "明细"
        title = "任务自评/互评对比"

    fig = go.Figure()
    # 底层：自评（正常宽度）
    fig.add_trace(go.Bar(
        y=agg[cat_col],
        x=agg["自评值"],
        orientation="h",
        name="自评分数",
        marker_color="#ff7f0e",
        opacity=1.0,
        width=0.6
    ))
    # 上层：互评（宽度收窄、上浮、透明度60%）
    fig.add_trace(go.Bar(
        y=agg[cat_col],
        x=agg["互评值"],
        orientation="h",
        name="互评分数",
        marker_color="#4cc9f0",
        opacity=0.8,
        width=0.4
    ))

    fig.update_layout(
        title=title,
        template="plotly_dark",
        height=max(400, len(agg)*40),
        legend=dict(orientation="h", y=-0.15, x=0.5, xanchor="center"),
        barmode="overlay",
        xaxis=dict(title="分数", showgrid=True, gridcolor="#444"),
        yaxis=dict(title=cat_col, showgrid=False),
        margin=dict(l=10, r=10, t=40, b=60)
    )
    return fig

def chart_bullet_advanced(df0: pd.DataFrame, dim: str = "员工"):
    if df0.empty:
        return go.Figure()
    if dim == "员工":
        agg_df = df0.groupby("员工").agg({"自评值":"sum","互评值":"sum"}).reset_index()
        cat_col = "员工"
        title = "【高级版】员工自评&互评分数对比"
    else:
        agg_df = df0.groupby("明细").agg({"自评值":"sum","互评值":"sum"}).reset_index()
        cat_col = "明细"
        title = "【高级版】任务自评&互评分数对比"

    agg_df = agg_df.sort_values("互评值", ascending=True).reset_index(drop=True)
    all_max = max(agg_df["自评值"].max(), agg_df["互评值"].max()) * 1.2

    fig = go.Figure()
    # 底层：自评
    fig.add_trace(go.Bar(
        y=agg_df[cat_col],
        x=agg_df["自评值"],
        orientation="h",
        marker_color="#ff7f0e",
        name="自评分数",
        opacity=1.0,
        width=0.6
    ))
    # 上层：互评（窄宽度+透明度60%）
    fig.add_trace(go.Bar(
        y=agg_df[cat_col],
        x=agg_df["互评值"],
        orientation="h",
        marker_color="#4cc9f0",
        name="互评分数",
        opacity=0.8,
        width=0.4
    ))

    fig.update_layout(
        title=title,
        template="plotly_dark",
        height=max(450, len(agg_df)*42),
        xaxis=dict(range=[0, all_max], title="分数", gridcolor="#444"),
        yaxis=dict(title=cat_col),
        legend=dict(orientation="h", y=-0.18, xanchor="center", x=0.5),
        barmode="overlay",
        margin=dict(l=10, r=10, t=40, b=65)
    )
    return fig

# ===================== 能力分析 =====================
def chart_ability(df0: pd.DataFrame, selected_emps: List[str]):
    if df0.empty:
        return go.Figure(), go.Figure(), go.Figure()
    tasks = df0["明细"].unique().tolist()
    fig1, fig2, fig3 = go.Figure(), go.Figure(), go.Figure()

    for idx, sheet in enumerate(time_choice):
        color = COLOR_POOL[idx % len(COLOR_POOL)]
        df_sheet = get_merged_df([sheet], selected_groups)
        if df_sheet.empty:
            continue
        pivot_self = df_sheet.pivot_table(index="明细", columns="员工", values="自评值", aggfunc="sum", fill_value=0)
        pivot_peer = df_sheet.pivot_table(index="明细", columns="员工", values="互评值", aggfunc="sum", fill_value=0)

        # 根据分数维度动态渲染曲线
        for emp in selected_emps:
            # 仅自评
            if score_dimension == "自评分数":
                if emp in pivot_self.columns:
                    fig1.add_trace(go.Scatter(
                        x=tasks, y=pivot_self[emp].reindex(tasks, fill_value=0),
                        mode="lines+markers", name=f"{sheet}-{emp}",
                        line=dict(color=color, width=3), marker=dict(size=7)
                    ))
            # 仅互评
            elif score_dimension == "互评分数":
                if emp in pivot_peer.columns:
                    fig1.add_trace(go.Scatter(
                        x=tasks, y=pivot_peer[emp].reindex(tasks, fill_value=0),
                        mode="lines+markers", name=f"{sheet}-{emp}",
                        line=dict(color=color, width=3), marker=dict(size=7)
                    ))
            # 双维度
            else:
                if emp in pivot_self.columns:
                    fig1.add_trace(go.Scatter(
                        x=tasks, y=pivot_self[emp].reindex(tasks, fill_value=0),
                        mode="lines+markers", name=f"{sheet}-{emp}(自评)",
                        line=dict(color=color, width=3), marker=dict(size=7)
                    ))
                if emp in pivot_peer.columns:
                    fig1.add_trace(go.Scatter(
                        x=tasks, y=pivot_peer[emp].reindex(tasks, fill_value=0),
                        mode="lines+markers", name=f"{sheet}-{emp}(互评)",
                        line=dict(color=color, width=3, dash="dash"), marker=dict(size=7)
                    ))

        # 任务汇总曲线
        if score_dimension == "自评分数":
            sum_data = pivot_self.sum(axis=1).reindex(tasks, fill_value=0)
            fig2.add_trace(go.Scatter(
                x=tasks, y=sum_data, mode="lines+markers",
                name=f"{sheet}", line=dict(color=color, width=3)
            ))
        elif score_dimension == "互评分数":
            sum_data = pivot_peer.sum(axis=1).reindex(tasks, fill_value=0)
            fig2.add_trace(go.Scatter(
                x=tasks, y=sum_data, mode="lines+markers",
                name=f"{sheet}", line=dict(color=color, width=3)
            ))
        else:
            sum_self = pivot_self.sum(axis=1).reindex(tasks, fill_value=0)
            sum_peer = pivot_peer.sum(axis=1).reindex(tasks, fill_value=0)
            fig2.add_trace(go.Scatter(
                x=tasks, y=sum_self, mode="lines+markers",
                name=f"{sheet}(自评)", line=dict(color=color, width=3)
            ))
            fig2.add_trace(go.Scatter(
                x=tasks, y=sum_peer, mode="lines+markers",
                name=f"{sheet}(互评)", line=dict(color=color, width=3, dash="dash")
            ))

        # 员工总分曲线
        if score_dimension == "自评分数":
            emp_sum = pivot_self.sum(axis=0)
            fig3.add_trace(go.Scatter(
                x=emp_sum.index, y=emp_sum.values, mode="lines+markers",
                name=f"{sheet}", line=dict(color=color, width=3)
            ))
        elif score_dimension == "互评分数":
            emp_sum = pivot_peer.sum(axis=0)
            fig3.add_trace(go.Scatter(
                x=emp_sum.index, y=emp_sum.values, mode="lines+markers",
                name=f"{sheet}", line=dict(color=color, width=3)
            ))
        else:
            emp_sum_self = pivot_self.sum(axis=0)
            emp_sum_peer = pivot_peer.sum(axis=0)
            fig3.add_trace(go.Scatter(
                x=emp_sum_self.index, y=emp_sum_self.values, mode="lines+markers",
                name=f"{sheet}(自评)", line=dict(color=color, width=3)
            ))
            fig3.add_trace(go.Scatter(
                x=emp_sum_peer.index, y=emp_sum_peer.values, mode="lines+markers",
                name=f"{sheet}(互评)", line=dict(color=color, width=3, dash="dash")
            ))

    # 统一布局
    title_map = {
        "自评分数": "员工任务完成曲线（自评）",
        "互评分数": "员工任务完成曲线（互评）",
        "双维度对比": "员工任务完成曲线（双维度）"
    }
    fig1.update_layout(title=title_map[score_dimension], template="plotly_dark", legend=dict(orientation="h", y=-0.25))
    fig2.update_layout(title="任务整体趋势", template="plotly_dark", legend=dict(orientation="h", y=-0.25))
    fig3.update_layout(title="员工总分对比", template="plotly_dark", legend=dict(orientation="h", y=-0.25))
    return fig1, fig2, fig3

# 指标卡片
def show_cards(df0: pd.DataFrame):
    if df0.empty:
        return
    total_task = df0["明细"].nunique()
    total_emp = df0["员工"].nunique()
    s1, s2 = get_score_cols()

    if score_dimension == "双维度对比":
        g_self = df0.groupby("员工")["自评值"].sum()
        g_peer = df0.groupby("员工")["互评值"].sum()
        top_self = g_self.idxmax() if not g_self.empty else "-"
        top_peer = g_peer.idxmax() if not g_peer.empty else "-"
        avg_self = round(g_self.mean(),1) if not g_self.empty else 0
        avg_peer = round(g_peer.mean(),1) if not g_peer.empty else 0

        c1,c2,c3,c4,c5,c6 = st.columns(6)
        c1.markdown(f"""<div class='metric-card'><div class='metric-value'>{total_task}</div><div class='metric-label'>任务总数</div></div>""", unsafe_allow_html=True)
        c2.markdown(f"""<div class='metric-card'><div class='metric-value'>{total_emp}</div><div class='metric-label'>人员总数</div></div>""", unsafe_allow_html=True)
        c3.markdown(f"""<div class='metric-card'><div class='metric-value'>{top_self}</div><div class='metric-label'>自评最高人员</div></div>""", unsafe_allow_html=True)
        c4.markdown(f"""<div class='metric-card'><div class='metric-value'>{top_peer}</div><div class='metric-label'>互评最高人员</div></div>""", unsafe_allow_html=True)
        c5.markdown(f"""<div class='metric-card'><div class='metric-value'>{avg_self}</div><div class='metric-label'>自评平均分</div></div>""", unsafe_allow_html=True)
        c6.markdown(f"""<div class='metric-card'><div class='metric-value'>{avg_peer}</div><div class='metric-label'>互评平均分</div></div>""", unsafe_allow_html=True)
    else:
        g = df0.groupby("员工")[s1].sum()
        top_name = g.idxmax() if not g.empty else "-"
        avg_val = round(g.mean(),1) if not g.empty else 0
        c1,c2,c3,c4 = st.columns(4)
        c1.markdown(f"""<div class='metric-card'><div class='metric-value'>{total_task}</div><div class='metric-label'>任务总数</div></div>""", unsafe_allow_html=True)
        c2.markdown(f"""<div class='metric-card'><div class='metric-value'>{total_emp}</div><div class='metric-label'>人员总数</div></div>""", unsafe_allow_html=True)
        c3.markdown(f"""<div class='metric-card'><div class='metric-value'>{top_name}</div><div class='metric-label'>{s2}最高人员</div></div>""", unsafe_allow_html=True)
        c4.markdown(f"""<div class='metric-card'><div class='metric-value'>{avg_val}</div><div class='metric-label'>{s2}平均分</div></div>""", unsafe_allow_html=True)
    st.markdown("<hr/>", unsafe_allow_html=True)

# ==================== 主页面渲染 ====================
st.title("技能覆盖分析大屏")

if view == "编辑数据":
    if not time_choice:
        st.warning("请先选择时间点再编辑数据")
    else:
        show_cards(df)
        st.info("直接编辑表格，修改后可点击下方按钮保存或刷新总和")
        edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True)

        # 按钮顺序：保存在上，更新总和在下
        if st.button("💾 保存修改到Excel文件"):
            try:
                sheet_name = time_choice[0]
                edited_df = calc_all_sum(edited_df)
                with get_excel_writer(SAVE_FILE, mode="a") as writer:
                    edited_df.to_excel(writer, sheet_name=sheet_name, index=False)
                st.cache_data.clear()
                st.success(f"已保存至 {sheet_name}")
            except Exception as e:
                st.error(f"保存失败: {str(e)}")

        if st.button("🔄 一键更新 自评/互评 数量总和"):
            edited_df = calc_all_sum(edited_df)
            st.success("已重新计算数量总和！")
            st.rerun()

elif view == "大屏轮播":
    if not time_choice:
        st.warning("请选择时间点")
    else:
        st_autorefresh(interval=10000, key="auto_ref")
        show_cards(df)
        chart_list = [("人员排名", chart_total(df)), ("任务堆叠图", chart_stack(df)), ("热力图", chart_heat(df))]
        if "carousel_idx" not in st.session_state:
            st.session_state.carousel_idx = 0
        st.session_state.carousel_idx = (st.session_state.carousel_idx + 1) % len(chart_list)
        name, opt = chart_list[st.session_state.carousel_idx]
        st.subheader(name)
        if isinstance(opt, go.Figure):
            st.plotly_chart(opt, use_container_width=True)
        else:
            st.markdown('<div class="heatmap-container">', unsafe_allow_html=True)
            st_echarts(opt, height=f"{max(600, len(df['明细'].unique())*28)}px", theme="dark")
            st.markdown('</div>', unsafe_allow_html=True)

elif view == "单页模式":
    if not time_choice:
        st.warning("请选择时间点")
    else:
        show_cards(df)
        opt_name = st.sidebar.selectbox("选择图表", ["人员完成任务数量排名","任务对比（堆叠柱状图）","任务-人员热力图"])
        if opt_name == "人员完成任务数量排名":
            fig = chart_total(df)
            st.plotly_chart(fig, use_container_width=True)
        elif opt_name == "任务对比（堆叠柱状图）":
            fig = chart_stack(df)
            st.plotly_chart(fig, use_container_width=True)
        else:
            opt = chart_heat(df)
            st.markdown('<div class="heatmap-container">', unsafe_allow_html=True)
            st_echarts(opt, height=f"{max(600, len(df['明细'].unique())*28)}px", theme="dark")
            st.markdown('</div>', unsafe_allow_html=True)

elif view == "显示所有视图":
    if not time_choice:
        st.warning("请选择时间点")
    else:
        show_cards(df)
        st.subheader("人员完成任务数量排名")
        st.plotly_chart(chart_total(df), use_container_width=True)
        st.subheader("任务对比（堆叠柱状图）")
        st.plotly_chart(chart_stack(df), use_container_width=True)
        st.subheader("任务-人员热力图")
        opt = chart_heat(df)
        st.markdown('<div class="heatmap-container">', unsafe_allow_html=True)
        st_echarts(opt, height=f"{max(600, len(df['明细'].unique())*28)}px", theme="dark")
        st.markdown('</div>', unsafe_allow_html=True)

elif view == "能力分析":
    if not time_choice:
        st.warning("请选择时间点")
    else:
        st.subheader("能力分析图表")
        emp_list = df["员工"].unique().tolist()
        sel_emp = st.sidebar.multiselect("选择展示员工", emp_list, default=emp_list)
        f1,f2,f3 = chart_ability(df, sel_emp)
        st.plotly_chart(f1, use_container_width=True)
        st.plotly_chart(f2, use_container_width=True)
        st.plotly_chart(f3, use_container_width=True)

elif view == "基础子弹图":
    if not time_choice:
        st.warning("请选择时间点")
    else:
        st.subheader("基础自评-互评子弹图")
        dim = st.radio("对比维度", ["员工维度","任务维度"], horizontal=True)
        d = "员工" if dim == "员工维度" else "明细"
        fig = chart_bullet_base(df, d)
        st.plotly_chart(fig, use_container_width=True)

elif view == "高级子弹图":
    if not time_choice:
        st.warning("请选择时间点")
    else:
        st.subheader("高级自评-互评子弹图")
        dim = st.radio("对比维度", ["员工维度","任务维度"], horizontal=True)
        d = "员工" if dim == "员工维度" else "明细"
        fig = chart_bullet_advanced(df, d)
        st.plotly_chart(fig, use_container_width=True)