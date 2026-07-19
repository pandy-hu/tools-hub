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
