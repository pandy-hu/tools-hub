# -*- coding: utf-8 -*-
"""三合一小工具核心逻辑: PDF转Excel / Excel公式 / AI简历。
与 UI 解耦, 便于测试与在 Streamlit Cloud 上部署。"""
import io
import json
import pandas as pd
import pdfplumber
import requests
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ===================== PDF 转 Excel =====================
def pdf_extract_tables(pdf_bytes):
    """返回 [(页码, 表序号, 二维表)] —— 二维表为 list[list[str|None]]。"""
    tables = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for pi, page in enumerate(pdf.pages):
            found = page.extract_tables()
            for ti, tbl in enumerate(found):
                if tbl and len(tbl) > 0:
                    tables.append((pi + 1, ti + 1, tbl))
    return tables


def pdf_tables_to_excel(tables, merge_all=True, skip_empty=True):
    """把表格写成 Excel 字节流, 返回 (bytes, sheet_names)。"""
    output = io.BytesIO()
    if merge_all:
        all_rows = []
        for _pi, _ti, tbl in tables:
            for r in tbl:
                row = [str(c).strip() if c is not None else "" for c in r]
                if skip_empty and all(v == "" for v in row):
                    continue
                all_rows.append(row)
        maxc = max((len(r) for r in all_rows), default=1)
        df = pd.DataFrame([r + [""] * (maxc - len(r)) for r in all_rows])
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="合并表格", index=False, header=False)
            _style_sheet(writer.sheets["合并表格"])
        return output.getvalue(), ["合并表格"]
    else:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            sheet_names = []
            for _idx, (pi, ti, tbl) in enumerate(tables):
                rows = []
                for r in tbl:
                    row = [str(c).strip() if c is not None else "" for c in r]
                    if skip_empty and all(v == "" for v in row):
                        continue
                    rows.append(row)
                if not rows:
                    continue
                maxc = max((len(r) for r in rows), default=1)
                df = pd.DataFrame([r + [""] * (maxc - len(r)) for r in rows])
                name = f"P{pi}_T{ti}"[:31]
                df.to_excel(writer, sheet_name=name, index=False, header=False)
                _style_sheet(writer.sheets[name])
                sheet_names.append(name)
        return output.getvalue(), sheet_names


def _style_sheet(ws):
    thin = Side(style="thin", color="D0D7E5")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill("solid", fgColor="2D6CDF")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = hdr_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws.iter_rows():
        for c in row:
            c.border = border
            if c.row != 1:
                c.alignment = Alignment(vertical="center")
    for col in ws.columns:
        width = max((len(str(c.value)) for c in col if c.value is not None), default=8)
        ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 10), 40)


# ===================== 大模型调用 (公式 + 简历共用) =====================
PRESETS = {
    "DeepSeek (推荐·便宜)": {"base_url": "https://api.deepseek.com/v1", "model": "deepseek-chat"},
    "OpenAI": {"base_url": "https://api.openai.com/v1", "model": "gpt-4o-mini"},
    "通义千问": {"base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1", "model": "qwen-plus"},
    "自定义": {"base_url": "", "model": ""},
}

FORMULA_MOCK = {
    "formula": "=SUM(A1:A10)/COUNT(A1:A10)",
    "explain": "先求 A1 到 A10 的总和，再除以个数，得到平均值。",
    "steps": [
        "用 SUM(A1:A10) 计算这一列的总和",
        "用 COUNT(A1:A10) 计算有多少个数字",
        "两者相除得到平均值（等价于 AVERAGE）",
    ],
    "example": "A1:A10 = 10,20,30 → SUM=60, COUNT=3 → 结果 20",
}

RESUME_MOCK = {
    "summary": "拥有 5 年互联网运营经验，擅长用户增长与内容策划，曾主导多个从 0 到 1 的项目，数据驱动决策，具备强执行力与跨团队协同能力。",
    "experience": [
        "● 主导某 APP 用户增长项目，3 个月内新增注册用户 12 万，留存率提升 18%",
        "● 搭建内容矩阵，公众号粉丝从 2 万增至 15 万，单篇最高阅读 50 万+",
        "● 设计裂变活动 SOP，获客成本降低 40%",
    ],
    "skills": ["用户增长", "内容运营", "数据分析", "活动策划", "SQL", "私域运营"],
    "education_tip": "建议补充 PMP 或 Google Analytics 证书以增强竞争力（如暂无相关学历背书）。",
    "ats_tips": [
        "使用标准职位名称（如'运营经理'）便于系统匹配",
        "技能区放置与 JD 高度重合的关键词",
        "避免表格/图片排版，纯文本更利于 ATS 解析",
    ],
}

FORMULA_SYSTEM = """你是一个 Excel 公式专家。用户会用自然语言描述他想要的计算，
请只返回一个 JSON 对象，不要包含任何多余文字或 markdown 代码块标记。
JSON 结构：
{
  "formula": "写在这里，例如 =SUM(A1:A10)",
  "explain": "用中文一句话解释这个公式在做什么",
  "steps": ["步骤1", "步骤2", "步骤3"],
  "example": "给出一个具体单元格的例子，例如 A1=10, A2=20 → 结果 30"
}
要求：formula 必须是可在 Excel / WPS 中直接使用的合法公式，以 = 开头。"""

RESUME_SYSTEM = """你是一位资深的中文简历优化顾问，熟悉 ATS（简历筛选系统）规则。
用户会提供：目标岗位、工作年限、以及一段原始经历/背景文字。
请只返回一个 JSON 对象，不要包含任何多余文字或 markdown 代码块标记。
JSON 结构：
{
  "summary": "3-4 句专业个人总结，突出核心竞争力与岗位匹配度",
  "experience": ["每条用 ● 开头，量化成果、用动词开头，例如 ● 主导XX项目，使转化率提升30%"],
  "skills": ["技能关键词1","技能关键词2"],
  "education_tip": "针对其背景的学历/证书建议（若用户未提供则给通用建议）",
  "ats_tips": ["ATS 优化建议1","ATS 优化建议2"]
}
要求：全部使用简体中文，专业、具体、可量化，避免空话。"""


def _call_llm(system_prompt, user_msg, api_key, base_url, model):
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_msg},
        ],
        "temperature": 0.3,
        "response_format": {"type": "json_object"},
    }
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    try:
        r = requests.post(
            base_url.rstrip("/") + "/chat/completions",
            headers=headers, json=payload, timeout=30,
        )
        r.raise_for_status()
        return json.loads(r.json()["choices"][0]["message"]["content"])
    except requests.exceptions.HTTPError:
        # 部分模型不支持 response_format，降级重试一次
        payload.pop("response_format", None)
        r2 = requests.post(
            base_url.rstrip("/") + "/chat/completions",
            headers=headers, json=payload, timeout=30,
        )
        r2.raise_for_status()
        return json.loads(r2.json()["choices"][0]["message"]["content"])


def formula_generate(prompt, api_key, base_url, model, mock=False):
    """调用大模型生成 Excel 公式。mock=True 或缺少 key 时返回演示数据。"""
    if mock or not api_key or not base_url:
        return FORMULA_MOCK
    return _call_llm(FORMULA_SYSTEM, prompt, api_key, base_url, model)


def resume_generate(target_role, years, background, api_key, base_url, model, mock=False):
    """调用大模型生成优化简历。mock=True 或缺少 key 时返回演示数据。"""
    if mock or not api_key or not base_url:
        return RESUME_MOCK
    user_msg = f"目标岗位：{target_role}\n工作年限：{years}\n原始背景/经历：\n{background}"
    return _call_llm(RESUME_SYSTEM, user_msg, api_key, base_url, model)


# ===================== 签证要求聚合 (信息差版) =====================
import json as _json
import os as _os

def load_visa_data():
    """读取签证数据集, 返回 (meta, list[dict])。"""
    path = _os.path.join(_os.path.dirname(__file__), "visa_data.json")
    with open(path, "r", encoding="utf-8") as f:
        data = _json.load(f)
    return data.get("_meta", {}), data.get("countries", [])


def visa_search(keyword="", region="全部", policy="全部"):
    """按关键词(国家名) / 地区 / 签证类型筛选。返回 list[dict]。"""
    _meta, countries = load_visa_data()
    kw = keyword.strip().lower()
    result = []
    for c in countries:
        if kw and kw not in c["country"].lower() and kw not in c.get("region", "").lower():
            continue
        if region != "全部" and c.get("region") != region:
            continue
        if policy != "全部" and c.get("policy") != policy:
            continue
        result.append(c)
    return result


def visa_regions():
    _meta, countries = load_visa_data()
    rs = sorted({c.get("region", "其他") for c in countries})
    return ["全部"] + rs


def visa_policies():
    return ["全部", "免签", "免签/落地签", "落地签", "落地签/电子签", "电子签",
            "电子旅行授权", "电子签/持美签免签", "需提前办理"]


# ===================== 数据可视化 =====================
def viz_load_data(raw_bytes, filename):
    """从上传的 CSV/Excel 字节加载为 DataFrame。"""
    name = (filename or "").lower()
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(io.BytesIO(raw_bytes))
    last_err = None
    for enc in ("utf-8-sig", "utf-8", "gbk", "gb18030", "latin-1"):
        try:
            return pd.read_csv(io.BytesIO(raw_bytes), encoding=enc)
        except Exception as e:
            last_err = e
    raise ValueError(f"无法解析该文件（试试 UTF-8 或 GBK 编码的 CSV）：{last_err}")


def viz_is_numeric(series):
    """判断某列是否可当作数值（原生数值，或多数能转数值）。"""
    if pd.api.types.is_numeric_dtype(series):
        return True
    return pd.to_numeric(series, errors="coerce").notna().mean() > 0.5


def viz_make_figure(df, chart_type, x_col, y_cols, title="", agg="不聚合", top_n=0):
    """返回 plotly.graph_objects.Figure。

    chart_type: 柱状图 / 折线图 / 饼图 / 散点图
    agg: 不聚合 / 求和 / 平均 / 计数（x 为分类列、y 为数值列时生效）
    top_n: >0 时按 y 第一列取前 N 行（绘图前排序）
    """
    import plotly.graph_objects as go
    work = df.copy()
    if agg != "不聚合" and x_col and y_cols:
        try:
            if agg == "求和":
                work = work.groupby(x_col, as_index=False)[y_cols].sum(numeric_only=True)
            elif agg == "平均":
                work = work.groupby(x_col, as_index=False)[y_cols].mean(numeric_only=True)
            elif agg == "计数":
                work = work.groupby(x_col, as_index=False)[y_cols].count()
        except Exception:
            pass
    if top_n and top_n > 0 and y_cols:
        y0 = y_cols[0]
        try:
            work = work.sort_values(y0, ascending=False).head(top_n)
        except Exception:
            pass

    fig = go.Figure()
    if chart_type == "柱状图":
        for y in y_cols:
            fig.add_bar(name=str(y), x=work[x_col], y=work[y])
        fig.update_layout(barmode="group")
    elif chart_type == "折线图":
        for y in y_cols:
            fig.add_trace(go.Scatter(x=work[x_col], y=work[y], mode="lines+markers", name=str(y)))
    elif chart_type == "饼图":
        y = y_cols[0] if y_cols else None
        fig.add_trace(go.Pie(labels=work[x_col], values=work[y], textinfo="label+percent"))
    elif chart_type == "散点图":
        y = y_cols[0] if y_cols else None
        if y:
            fig.add_trace(go.Scatter(x=work[x_col], y=work[y], mode="markers", name=str(y)))
    fig.update_layout(
        title=title or "数据可视化",
        template="plotly_white",
        font=dict(family="Microsoft YaHei, PingFang SC, sans-serif", size=13),
        margin=dict(l=40, r=20, t=50, b=40),
        legend=dict(orientation="h", y=-0.2),
    )
    return fig
