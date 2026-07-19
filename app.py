# -*- coding: utf-8 -*-
"""三合一小工具站: PDF转Excel / Excel公式机器人 / AI简历生成器。
对标 StarterStory 成功案例, 可一键部署到 Streamlit Cloud 获得公开网址。"""
import pandas as pd
import streamlit as st
from core import (
    pdf_extract_tables, pdf_tables_to_excel,
    formula_generate, resume_generate, PRESETS,
)

st.set_page_config(page_title="小工具站 · 三个能赚钱的小工具", page_icon="🧰", layout="centered")

PRIMARY = "#2D6CDF"
st.markdown(f"""
<style>
.big-title {{ font-size: 26px; font-weight: 800; color:#1a2b4a; margin-bottom:2px; }}
.sub {{ color:#5b6b85; font-size:14px; margin-bottom:18px; }}
.card {{ background:#F5F8FF; border:1px solid #e3ebfb; border-radius:14px; padding:18px; margin:14px 0; }}
.formula-box {{ background:#0f1b12; color:#7CFC9B; font-family:Consolas,monospace; font-size:16px; padding:14px 16px; border-radius:10px; margin:8px 0; word-break:break-all; }}
.sec {{ font-weight:700; color:{PRIMARY}; margin:14px 0 4px; font-size:15px; }}
.box {{ background:#fff; border:1px solid #e3ebfb; border-radius:10px; padding:12px 14px; }}
.foot {{ color:#9aa7bd; font-size:12px; margin-top:24px; text-align:center; }}
div.stButton > button {{ background:{PRIMARY}; color:#fff; font-weight:700; border-radius:10px; padding:10px 22px; }}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="big-title">🧰 小工具站</div>', unsafe_allow_html=True)
st.markdown('<div class="sub">三个对标海外成功案例的小工具，免费可用。上方切换。</div>', unsafe_allow_html=True)

tab1, tab2, tab3, tab4 = st.tabs(["📄 PDF转Excel", "📊 Excel公式", "📋 AI简历", "💡 关于/升级"])

# ---------------- Tab 1: PDF 转 Excel ----------------
with tab1:
    st.markdown("### 📄 PDF 发票/账单 → Excel")
    st.caption("上传任意带表格的 PDF（发票、账单、银行流水、对账单），一键提取全部表格并导出 Excel。纯本地处理，文件不上传任何服务器。")
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        uploaded = st.file_uploader("拖入或点击上传 PDF 文件", type=["pdf"], key="pdf_up")
        c1, c2 = st.columns(2)
        with c1:
            merge_all = st.checkbox("合并为单个工作表", value=True, key="pdf_merge")
        with c2:
            skip_empty = st.checkbox("跳过空行", value=True, key="pdf_skip")
        st.markdown('</div>', unsafe_allow_html=True)
    if uploaded is not None:
        with st.spinner("正在解析 PDF 表格…"):
            raw = uploaded.read()
            try:
                tables = pdf_extract_tables(raw)
            except Exception as e:
                st.error(f"解析失败：{e}")
                tables = []
            if not tables:
                st.warning("未在 PDF 中检测到表格。可能是图片型 PDF（后续可加 OCR）。")
            else:
                st.success(f"✅ 检测到 {len(tables)} 张表格，来自 {len(set(t[0] for t in tables))} 页")
                preview_df = pd.DataFrame(tables[0][2])
                st.markdown("**第一张表预览：**")
                st.dataframe(preview_df, use_container_width=True, height=260)
                xlsx, sheets = pdf_tables_to_excel(tables, merge_all, skip_empty)
                st.download_button(
                    "⬇️ 下载 Excel",
                    data=xlsx,
                    file_name=uploaded.name.replace(".pdf", "") + "_提取.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.caption(f"已生成工作表：{', '.join(sheets)}")

# ---------------- Tab 2: Excel 公式机器人 ----------------
with tab2:
    st.markdown("### 📊 Excel 公式生成机器人")
    st.caption("用大白话描述你想算什么，一键生成可直接粘贴进 Excel/WPS 的合法公式。填 Key 联网，不填则演示。")
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        f_provider = st.selectbox("选择大模型", list(PRESETS.keys()), index=0, key="f_provider")
        f_key = st.text_input("API Key（留空则演示模式）", type="password", key="f_key")
        ca, cb = st.columns(2)
        with ca:
            f_url = st.text_input("API Base URL", value=PRESETS[f_provider]["base_url"], key="f_url")
        with cb:
            f_model = st.text_input("模型名", value=PRESETS[f_provider]["model"], key="f_model")
        st.markdown('</div>', unsafe_allow_html=True)
    st.markdown("**描述你的需求：**")
    examples = [
        "求 A 列中所有大于 100 的数字之和",
        "如果 B 列等于'已完成'则显示'是'否则'否'",
        "根据今天的日期算出还有多少天到月底",
        "在 C 列查找某个姓名对应的手机号",
    ]
    for ex in examples:
        if st.button(f"💡 {ex}", key="f_" + ex):
            st.session_state["formula_prompt"] = ex
    f_prompt = st.text_area("在这里输入（中文即可）", value=st.session_state.get("formula_prompt", ""),
                            height=90, key="f_prompt", placeholder="例如：统计 D 列中不重复的客户数量")
    if st.button("🚀 生成公式", use_container_width=True, key="f_go"):
        if not f_prompt.strip():
            st.warning("请先描述你的需求～")
        else:
            mock = (not f_key.strip()) or (not f_url.strip())
            with st.spinner("演示模式（未联网）…" if mock else "正在生成…"):
                try:
                    res = formula_generate(f_prompt, f_key, f_url, f_model, mock=mock)
                except Exception as e:
                    st.error(f"出错了：{e}")
                    res = None
            if res:
                st.markdown('<div class="formula-box">= ' + res.get("formula", "").lstrip("=") + '</div>', unsafe_allow_html=True)
                st.markdown(f"**说明：** {res.get('explain', '')}")
                if res.get("steps"):
                    st.markdown("**步骤：**")
                    for s in res["steps"]:
                        st.markdown(f"- {s}")
                if res.get("example"):
                    st.info(f"📌 例子：{res['example']}")
                st.caption("复制上面公式框内容，直接粘到 Excel/WPS 单元格即可。")

# ---------------- Tab 3: 中文 AI 简历生成器 ----------------
with tab3:
    st.markdown("### 📋 中文 AI 简历生成器")
    st.caption("填目标岗位+经历，AI 帮你重写得更专业、量化、过 ATS。填 Key 联网，不填则演示。")
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        r_provider = st.selectbox("选择大模型", list(PRESETS.keys()), index=0, key="r_provider")
        r_key = st.text_input("API Key（留空则演示模式）", type="password", key="r_key")
        ra, rb = st.columns(2)
        with ra:
            r_url = st.text_input("API Base URL", value=PRESETS[r_provider]["base_url"], key="r_url")
        with rb:
            r_model = st.text_input("模型名", value=PRESETS[r_provider]["model"], key="r_model")
        st.markdown('</div>', unsafe_allow_html=True)
    r_target = st.text_input("🎯 目标岗位", placeholder="例如：用户增长运营 / Python 后端工程师", key="r_target")
    r_years = st.text_input("⏳ 工作年限", placeholder="例如：5 年 / 应届", key="r_years")
    r_bg = st.text_area("📝 你的经历/背景（越具体越好）", height=160, key="r_bg",
                         placeholder="例如：在XX公司做运营，负责公众号，3个月涨粉10万；做过裂变活动…")
    if st.button("✨ 生成优化简历", use_container_width=True, key="r_go"):
        if not r_bg.strip():
            st.warning("至少填一下你的经历吧～")
        else:
            mock = (not r_key.strip()) or (not r_url.strip())
            with st.spinner("演示模式（未联网）…" if mock else "AI 正在重写你的简历…"):
                try:
                    res = resume_generate(r_target or "未指定", r_years or "未指定", r_bg, r_key, r_url, r_model, mock=mock)
                except Exception as e:
                    st.error(f"出错了：{e}")
                    res = None
            if res:
                st.markdown('<div class="sec">🧑‍💼 个人总结</div>', unsafe_allow_html=True)
                st.markdown(f'<div class="box">{res.get("summary", "")}</div>', unsafe_allow_html=True)
                st.markdown('<div class="sec">💼 工作经历（优化后）</div>', unsafe_allow_html=True)
                st.markdown('<div class="box">' + "<br>".join(res.get("experience", [])) + "</div>", unsafe_allow_html=True)
                st.markdown('<div class="sec">🛠 技能关键词</div>', unsafe_allow_html=True)
                st.markdown('<div class="box">' + "　".join(f"`{s}`" for s in res.get("skills", [])) + "</div>", unsafe_allow_html=True)
                if res.get("education_tip"):
                    st.markdown('<div class="sec">🎓 学历/证书建议</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="box">{res["education_tip"]}</div>', unsafe_allow_html=True)
                st.markdown('<div class="sec">✅ ATS 优化建议</div>', unsafe_allow_html=True)
                for t in res.get("ats_tips", []):
                    st.markdown(f"- {t}")
                full = (f"【个人总结】\n{res.get('summary','')}\n\n【工作经历】\n" + "\n".join(res.get("experience", []))
                         + f"\n\n【技能】\n" + "、".join(res.get("skills", [])) + f"\n\n【ATS建议】\n" + "\n".join(res.get("ats_tips", [])))
                st.download_button("⬇️ 下载简历文本", full, file_name="我的AI简历.txt", mime="text/plain")

# ---------------- Tab 4: 关于 / 升级 ----------------
with tab4:
    st.markdown("### 💡 关于这个小工具站")
    st.markdown("""
    这是一组**对标海外成功案例**的小工具（灵感来自 StarterStory）：

    - 📄 **PDF 转 Excel** —— 对标海外 $40K/月案例
    - 📊 **Excel 公式机器人** —— 对标海外 $20K/月案例
    - 📋 **中文 AI 简历生成器** —— 对标 Rezi $215K/月（中文差异化版）

    全部免费可用。PDF 工具本地运行、无需 Key；另外两个填 Key 即联网、不填也能看演示。
    """)
    upgrade_url = st.secrets.get("UPGRADE_URL", "")
    st.markdown("### 🔓 升级高级版")
    if upgrade_url:
        st.markdown(f"解锁 **批量处理 / API 接口 / 去演示限制 / 定制功能**，点这里 👉 [{upgrade_url}]({upgrade_url})")
    else:
        st.info("高级版（批量处理、API 接口、去演示限制、定制功能）即将推出。\n\n"
                 "想抢先体验或提需求：可在部署平台的 Secrets 里配置 `UPGRADE_URL` 指向你的购买/联系方式链接，"
                 "本页会自动显示升级按钮。")
    st.markdown('<div class="foot">MVP · 由 WorkBuddy 生成 · 三个工具合成一站</div>', unsafe_allow_html=True)
