# -*- coding: utf-8 -*-
"""小工具站: PDF转Excel / Excel公式机器人 / AI简历生成器 / 签证政策聚合 / 数据可视化 / Airtable导入 / AI配图生成器。
对标 StarterStory 成功案例, 可一键部署到 Streamlit Cloud 获得公开网址。"""
import pandas as pd
import streamlit as st
from core import (
    pdf_extract_tables, pdf_tables_to_excel,
    formula_generate, resume_generate, PRESETS,
    visa_search, visa_regions, visa_policies,
    viz_load_data, viz_is_numeric, viz_make_figure,
    airtable_list_tables, airtable_import,
    poster_themes, poster_sizes, generate_poster_html, poster_preview_html,
    poster_copy, PRESETS,
    image_process, image_tool_formats,
    qr_generate,
)

st.set_page_config(page_title="小工具站 · 一组能赚钱的小工具", page_icon="🧰", layout="centered")


def get_cloud_key(provider):
    """从 Streamlit Cloud Secrets 读取该服务商对应的密钥（云端已配置则免粘贴）。"""
    mapping = {
        "DeepSeek (推荐·便宜)": "DEEPSEEK_API_KEY",
        "OpenAI": "OPENAI_API_KEY",
        "通义千问": "DASHSCOPE_API_KEY",
        "自定义": "CUSTOM_API_KEY",
    }
    name = mapping.get(provider)
    try:
        return st.secrets.get(name, "") if name else ""
    except Exception:
        return ""


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
st.markdown('<div class="sub">一组对标海外成功案例的小工具（含中国护照信息差版），免费可用。上方切换。</div>', unsafe_allow_html=True)

# ---------------- 顶部导航：每行 6 个，超过自动换行 ----------------
TOOLS = [
    ("📄 PDF", "pdf"), ("📊 公式", "formula"), ("📋 简历", "resume"),
    ("🛂 签证", "visa"), ("📈 数据", "data"), ("🗄️ 导入", "airtable"),
    ("🎨 配图", "poster"), ("🖼️ 图片", "image"), ("🔳 二维码", "qr"), ("💡 关于", "about"),
]
PER_ROW = 6
if "active_tool" not in st.session_state:
    st.session_state["active_tool"] = "pdf"
for _i in range(0, len(TOOLS), PER_ROW):
    _cols = st.columns(PER_ROW)
    for _j, (_label, _key) in enumerate(TOOLS[_i:_i + PER_ROW]):
        with _cols[_j]:
            _active = st.session_state["active_tool"] == _key
            if st.button(_label, key="nav_" + _key, use_container_width=True,
                         type="primary" if _active else "secondary"):
                st.session_state["active_tool"] = _key
st.markdown('<hr style="margin:14px 0;border:none;border-top:1px solid #e3ebfb;">', unsafe_allow_html=True)

# ---------------- Tab 1: PDF 转 Excel ----------------
if st.session_state["active_tool"] == "pdf":
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
if st.session_state["active_tool"] == "formula":
    st.markdown("### 📊 Excel 公式生成机器人")
    st.caption("用大白话描述你想算什么，一键生成可直接粘贴进 Excel/WPS 的合法公式。填 Key 联网，不填则演示。")
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        f_provider = st.selectbox("选择大模型", list(PRESETS.keys()), index=0, key="f_provider")
        f_cloud_key = get_cloud_key(f_provider)
        f_key = st.text_input("API Key（留空则演示模式）", type="password", value=f_cloud_key, key="f_key")
        if f_cloud_key:
            st.success("✅ 已检测到云端密钥，直接联网生成（无需粘贴）")
        else:
            st.caption("没填 Key 则走演示模式。也可在部署平台 Secrets 配置密钥后免填。")
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
if st.session_state["active_tool"] == "resume":
    st.markdown("### 📋 中文 AI 简历生成器")
    st.caption("填目标岗位+经历，AI 帮你重写得更专业、量化、过 ATS。填 Key 联网，不填则演示。")
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        r_provider = st.selectbox("选择大模型", list(PRESETS.keys()), index=0, key="r_provider")
        r_cloud_key = get_cloud_key(r_provider)
        r_key = st.text_input("API Key（留空则演示模式）", type="password", value=r_cloud_key, key="r_key")
        if r_cloud_key:
            st.success("✅ 已检测到云端密钥，直接联网生成（无需粘贴）")
        else:
            st.caption("没填 Key 则走演示模式。也可在部署平台 Secrets 配置密钥后免填。")
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

# ---------------- Tab 4: 签证要求聚合 (信息差版) ----------------
if st.session_state["active_tool"] == "visa":
    st.markdown("### 🛂 中国护照签证政策聚合")
    st.caption("把散落在各使馆官网的签证政策，汇成一张可搜索的表 —— 对标海外 $20K/月案例的「中文信息差」版。")

    _meta, _ = (None, None)
    try:
        from core import load_visa_data
        _meta, _ = load_visa_data()
    except Exception:
        _meta = {}
    if _meta.get("disclaimer"):
        st.warning("⚠️ " + _meta["disclaimer"], icon="ℹ️")

    v_kw = st.text_input("🔍 搜国家 / 地区（如 泰国、欧洲）", key="v_kw", placeholder="输入国家名或地区")
    vc1, vc2 = st.columns(2)
    with vc1:
        v_region = st.selectbox("按地区筛选", visa_regions(), key="v_region")
    with vc2:
        v_policy = st.selectbox("按签证类型筛选", visa_policies(), key="v_policy")

    results = visa_search(v_kw, v_region, v_policy)
    st.markdown(f"**共匹配 {len(results)} 个目的地**")

    for c in results:
        pol = c.get("policy", "")
        if "免签" in pol and "需" not in pol:
            badge = f'<span style="background:#e6f7ec;color:#1a9e54;padding:2px 10px;border-radius:20px;font-size:13px;font-weight:700;">{pol}</span>'
        elif "需提前" in pol:
            badge = f'<span style="background:#fdeaea;color:#d4380d;padding:2px 10px;border-radius:20px;font-size:13px;font-weight:700;">{pol}</span>'
        else:
            badge = f'<span style="background:#fff4e0;color:#b8770a;padding:2px 10px;border-radius:20px;font-size:13px;font-weight:700;">{pol}</span>'
        fee = "免费" if not c.get("fee_cny") else f"约 ¥{c['fee_cny']}"
        st.markdown(f"""
        <div class="card">
          <div style="display:flex;justify-content:space-between;align-items:center;">
            <div style="font-size:18px;font-weight:800;color:#1a2b4a;">{c.get('flag','')} {c.get('country','')} <span style="font-size:13px;color:#5b6b85;font-weight:500;">· {c.get('region','')}</span></div>
            {badge}
          </div>
          <div style="margin-top:8px;color:#3a4a66;font-size:14px;">
            🕒 可停留：<b>{c.get('stay','-')}</b> ｜ ⏱ 办理：{c.get('processing','-')} ｜ 💰 费用：<b>{fee}</b>
          </div>
          <div style="margin-top:6px;color:#5b6b85;font-size:13px;">{c.get('note','')}</div>
        </div>
        """, unsafe_allow_html=True)
        with st.expander(f"📋 {c.get('country')} 所需材料 & 官方链接"):
            for r in c.get("requirements", []):
                st.markdown(f"- {r}")
            st.markdown(f"🔗 官方参考：[{c.get('official','')}]({c.get('official','')})")

# ---------------- Tab 5: 数据可视化 ----------------
if st.session_state["active_tool"] == "data":
    import io as _io
    st.markdown("### 📈 数据可视化小工具")
    st.caption("上传 CSV / Excel，自动识别字段，选个图表类型就出图。纯本地处理，数据不上传服务器。对标海外 $10K/月案例。")
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        v_file = st.file_uploader("拖入或点击上传 数据文件（.csv / .xlsx）", type=["csv", "xlsx", "xls"], key="viz_up")
        st.markdown('</div>', unsafe_allow_html=True)
    if v_file is not None:
        try:
            v_df = viz_load_data(v_file.read(), v_file.name)
        except Exception as e:
            st.error(f"读取失败：{e}")
            v_df = None
        if v_df is not None:
            cols = list(v_df.columns)
            num_cols = [c for c in cols if viz_is_numeric(v_df[c])]
            st.success(f"✅ 已加载 {len(v_df)} 行 × {len(cols)} 列")
            with st.expander("🔍 数据预览（前 20 行）"):
                st.dataframe(v_df.head(20), use_container_width=True, height=300)
            st.markdown('<div class="card">', unsafe_allow_html=True)
            v_type = st.selectbox("图表类型", ["柱状图", "折线图", "饼图", "散点图"], key="viz_type")
            vx, vy = st.columns(2)
            with vx:
                v_x = st.selectbox("X 轴 / 名称列", cols, key="viz_x")
            with vy:
                if v_type == "饼图" or v_type == "散点图":
                    v_y = st.selectbox("数值列（Y）", num_cols or cols, key="viz_y")
                    v_y_cols = [v_y]
                else:
                    v_y_cols = st.multiselect("数值列（Y，可多选）", num_cols or cols,
                                              default=(num_cols[:1] if num_cols else cols[:1]), key="viz_ys")
            v_title = st.text_input("图表标题（可选）", key="viz_title", placeholder="例如：2025 各月销售额")
            if v_type in ("柱状图", "折线图", "饼图"):
                v_agg = st.selectbox("聚合方式（X 为分类列时）", ["不聚合", "求和", "平均", "计数"], key="viz_agg")
            else:
                v_agg = "不聚合"
            v_top = st.number_input("只取前 N 行（0 = 全部）", min_value=0, max_value=100, value=0, step=1, key="viz_top")
            st.markdown('</div>', unsafe_allow_html=True)
            if st.button("🚀 生成图表", use_container_width=True, key="viz_go"):
                if not v_y_cols:
                    st.warning("请至少选一个数值列～")
                else:
                    with st.spinner("绘制中…"):
                        try:
                            fig = viz_make_figure(v_df, v_type, v_x, v_y_cols, v_title, v_agg, int(v_top))
                            st.plotly_chart(fig, use_container_width=True)
                        except Exception as e:
                            st.error(f"出图失败：{e}")
            # 下载：图表 PNG（best-effort）+ 数据 Excel
            try:
                import plotly.graph_objects as _go
                _fig = viz_make_figure(v_df, v_type, v_x, v_y_cols, v_title, v_agg, int(v_top))
                _buf = _io.BytesIO()
                _fig.write_image(_buf, format="png", width=900, height=520, scale=2)
                st.download_button("⬇️ 下载图表 PNG", data=_buf.getvalue(),
                                   file_name=v_file.name.rsplit(".", 1)[0] + "_图表.png", mime="image/png")
            except Exception:
                st.caption("（PNG 导出需 kaleido，未安装则跳过；可改下方式下载数据）")
            try:
                _xb = _io.BytesIO()
                v_df.to_excel(_xb, index=False)
                st.download_button("⬇️ 下载数据 (Excel)", data=_xb.getvalue(),
                                   file_name=v_file.name.rsplit(".", 1)[0] + "_数据.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception:
                pass

# ---------------- Tab 6: CSV → Airtable 导入 ----------------
if st.session_state["active_tool"] == "airtable":
    st.markdown("### 🗄️ CSV → Airtable 导入")
    st.caption("把 Excel/CSV 数据一键灌进你的 Airtable 表。需你自己的 Airtable Token（数据直连 Airtable，不经本站服务器）。对标海外 $20K/月案例。")
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        a_file = st.file_uploader("上传 CSV / Excel", type=["csv", "xlsx", "xls"], key="at_up")
        a_token = st.text_input("Airtable Token（以 pat 或 key 开头）", type="password", key="at_token")
        a_base = st.text_input("Base ID（形如 appXXXX）", key="at_base", placeholder="appXXXXXXXXXXXXXX")
        a_dry = st.checkbox("演练模式（只模拟、不真实写入）", value=True, key="at_dry")
        st.markdown('</div>', unsafe_allow_html=True)
    if a_file is not None:
        try:
            a_df = viz_load_data(a_file.read(), a_file.name)
        except Exception as e:
            st.error(f"读取失败：{e}")
            a_df = None
        if a_df is not None:
            st.success(f"✅ 已加载 {len(a_df)} 行 × {len(a_df.columns)} 列")
            if st.button("🔗 获取表结构", key="at_fetch"):
                if not (a_token and a_base):
                    st.warning("请先填 Token 和 Base ID")
                else:
                    with st.spinner("拉取表结构…"):
                        try:
                            tbls = airtable_list_tables(a_token, a_base)
                            st.session_state["at_tables"] = tbls
                            st.success(f"✅ 找到 {len(tbls)} 张表")
                        except Exception as e:
                            st.error(f"获取失败：{e}")
            tbls = st.session_state.get("at_tables", [])
            if tbls:
                tnames = [t["name"] for t in tbls]
                a_tname = st.selectbox("选择目标表", tnames, key="at_tname")
                sel = next((t for t in tbls if t["name"] == a_tname), None)
                if sel:
                    st.markdown("**字段映射**（CSV 列 → Airtable 字段）：")
                    mapping = {}
                    for f in sel["fields"]:
                        mapping[f["name"]] = st.selectbox(
                            f"🔑 {f['name']}（{f['type']}）", ["(忽略)"] + list(a_df.columns),
                            key="map_" + str(f["name"]))
                    if st.button("🚀 开始导入", use_container_width=True, key="at_go"):
                        records = []
                        for _, row in a_df.iterrows():
                            rec = {}
                            for fname, cname in mapping.items():
                                if cname != "(忽略)" and pd.notna(row[cname]):
                                    rec[fname] = str(row[cname])
                            if rec:
                                records.append(rec)
                        if not records:
                            st.warning("没有可导入的数据，请检查字段映射～")
                        elif a_dry:
                            st.info(f"🧪 演练模式：将向表「{a_tname}」写入 {len(records)} 条记录（未真实写入）。取消勾选「演练模式」后再次点击即可真实导入。")
                        else:
                            with st.spinner(f"导入中（{len(records)} 条）…"):
                                res = airtable_import(a_token, a_base, a_tname, records, dry_run=False)
                            if res["failed"] == 0:
                                st.success(f"✅ 导入成功 {res['success']} 条")
                            else:
                                st.warning(f"成功 {res['success']} 条，失败 {res['failed']} 条")
                                for e in res["errors"][:5]:
                                    st.caption(e)

# ---------------- Tab 8: AI 配图 / 海报生成器 ----------------
if st.session_state["active_tool"] == "poster":
    st.markdown("### 🎨 AI 配图 / 海报生成器")
    st.caption("输入主题，选模板和尺寸，一键生成社媒配图（小红书 / 公众号 / YouTube / 朋友圈）。纯本地生成，免 Key；填 Key 还能让 AI 帮你写文案。")
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        p_topic = st.text_input("🎯 内容主题（如：普通人如何用 AI 搞副业）", key="p_topic",
                                placeholder="输入你想做配图的主题")
        pa, pb = st.columns(2)
        with pa:
            p_theme = st.selectbox("配色模板", poster_themes(), key="p_theme")
        with pb:
            p_size = st.selectbox("尺寸 / 平台", poster_sizes(), key="p_size")
        # 把 AI 生成的文案同步进输入框：必须在下面的 widget 实例化之前执行一次
        if st.session_state.get("ai_synced"):
            st.session_state["p_title"] = st.session_state.get("ai_title", "")
            st.session_state["p_sub"] = st.session_state.get("ai_sub", "")
            st.session_state["p_tags"] = st.session_state.get("ai_tags", "")
            st.session_state["ai_synced"] = False
        p_title = st.text_input("大标题（留空则 AI 写 / 用主题）", key="p_title",
                                placeholder="例如：3 个被低估的赚钱小工具")
        p_sub = st.text_input("小标签（可选，如：信息差·副业）", key="p_sub")
        p_tags = st.text_input("标签（用逗号分隔，可选）", key="p_tags", placeholder="副业,AI工具,信息差")
        p_footer = st.text_input("底部署名（可选，如：by 小工具站）", key="p_footer")
        # 若 AI 已生成文案，给出明确提示，让用户知道内容已填入
        if st.session_state.get("ai_title"):
            st.success("✅ AI 已为你写好文案，已自动填入上方输入框，可直接修改后点「生成配图」。")
        st.markdown('</div>', unsafe_allow_html=True)

    # AI 写文案（可选）
    with st.expander("✨ 让 AI 帮我写文案（可选，填 Key 联网）"):
        p_provider = st.selectbox("选择大模型", list(PRESETS.keys()), index=0, key="p_provider")
        p_cloud_key = get_cloud_key(p_provider)
        p_key = st.text_input("API Key（留空则用自己的标题）", type="password", value=p_cloud_key, key="p_key")
        if p_cloud_key:
            st.success("✅ 已检测到云端密钥，可直接联网写文案")
        pu, pm = st.columns(2)
        with pu:
            p_url = st.text_input("API Base URL", value=PRESETS[p_provider]["base_url"], key="p_url")
        with pm:
            p_model = st.text_input("模型名", value=PRESETS[p_provider]["model"], key="p_model")
        if st.button("🤖 AI 生成文案", key="p_ai", use_container_width=True):
            if not p_topic.strip():
                st.warning("先填一下内容主题～")
            else:
                with st.spinner("AI 写文案中…"):
                    try:
                        cp = poster_copy(p_topic, p_key, p_url, p_model,
                                         mock=(not p_key.strip()) or (not p_url.strip()))
                        # 写到独立 key（不用 widget 绑定的 p_title/p_sub/p_tags，
                        # 否则 Streamlit 报错 "cannot be modified after widget instantiated"）
                        st.session_state["ai_title"] = cp.get("title", p_topic)
                        st.session_state["ai_sub"] = cp.get("subtitle", "")
                        st.session_state["ai_tags"] = ",".join(cp.get("tags", []))
                        # 标记需要把 AI 文案同步进输入框（在 widget 实例化之前执行一次）
                        st.session_state["ai_synced"] = True
                        st.rerun()
                    except Exception as e:
                        st.error(f"AI 文案失败：{e}")

    if st.button("🚀 生成配图", use_container_width=True, key="p_go"):
        # 优先用 AI 生成的文案，fallback 到用户手填
        title = (st.session_state.get("ai_title") or p_title or p_topic or "我的配图").strip()
        sub = (st.session_state.get("ai_sub") or p_sub).strip()
        _ai_tags = st.session_state.get("ai_tags", "")
        tags = [t.strip() for t in (_ai_tags if _ai_tags else p_tags).split(",") if t.strip()] if (_ai_tags or p_tags) else []
        footer = p_footer.strip()
        poster = generate_poster_html(title, sub, p_theme, p_size, footer, tags)
        # 取真实尺寸用于预览缩放
        from core import POSTER_SIZES as _PS
        _w, _h = _PS[p_size]
        st.markdown(poster_preview_html(poster, _w, _h), unsafe_allow_html=True)
        st.success("✅ 配图已生成！下方下载 HTML，用浏览器打开即可截图 / 打印成图片。")
        st.download_button("⬇️ 下载配图 HTML", data=poster, file_name="我的配图.html", mime="text/html")
        st.caption("提示：下载后用浏览器打开，右键图片区域 → 截图，或按 Ctrl+P 打印成 PDF/PNG，即得高清配图。")

# ---------------- Tab 9: 图片处理工具箱 ----------------
if st.session_state["active_tool"] == "image":
    import io as _io
    import zipfile
    st.markdown("### 🖼️ 图片处理工具箱")
    st.caption("上传图片，一键压缩体积、改尺寸、转格式、加水印。纯本地处理，文件不上传服务器。对标海外图片压缩类小工具（$5K-20K/月）。")
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        imgs = st.file_uploader("拖入或点击上传图片（可多选）", type=["png", "jpg", "jpeg", "webp", "bmp", "gif"],
                                accept_multiple_files=True, key="img_up")
        if1, if2 = st.columns(2)
        with if1:
            fmt = st.selectbox("输出格式", image_tool_formats(), key="img_fmt")
        with if2:
            quality = st.slider("压缩质量 (JPG 生效)", 10, 95, 82, key="img_q")
        ie1, ie2 = st.columns(2)
        with ie1:
            max_edge = st.number_input("最长边像素 (0=不缩放)", min_value=0, max_value=4000, value=0, step=50, key="img_edge")
        with ie2:
            wm = st.text_input("水印文字（可选，如 © 小工具站）", key="img_wm")
        st.markdown('</div>', unsafe_allow_html=True)
    if imgs:
        st.success(f"✅ 已选择 {len(imgs)} 张图片")
        if st.button("🚀 处理并打包下载", use_container_width=True, key="img_go"):
            results = []
            with st.spinner("处理中…"):
                for f in imgs:
                    raw = f.read()
                    try:
                        out, ext, w, h = image_process(raw, fmt=fmt, quality=quality,
                                                       max_edge=max_edge, watermark=wm)
                        results.append((f.name, out, ext, w, h, len(raw), len(out)))
                    except Exception as e:
                        st.error(f"处理失败 {f.name}：{e}")
            if results:
                for name, out, ext, w, h, orig, new in results:
                    ratio = (1 - new / float(orig)) * 100 if orig else 0
                    st.markdown(f"**{name}** → {w}×{h}px，{ext.upper()}，体积 {orig // 1024}KB → {new // 1024}KB（省 {ratio:.0f}%）")
                    st.image(out, caption=name, use_container_width=True)
                    st.download_button(f"⬇️ 下载 {name}", data=out,
                                       file_name=name.rsplit('.', 1)[0] + '.' + ext,
                                       mime=f"image/{ext}", key="dl_" + name)
                if len(results) > 1:
                    zbuf = _io.BytesIO()
                    with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
                        for name, out, ext, w, h, orig, new in results:
                            z.writestr(name.rsplit('.', 1)[0] + '.' + ext, out)
                    st.download_button("⬇️ 打包下载全部 (ZIP)", data=zbuf.getvalue(),
                                       file_name="小工具站_图片.zip", mime="application/zip")

# ---------------- Tab 9: 二维码生成器 ----------------
if st.session_state["active_tool"] == "qr":
    st.markdown("### 🔳 二维码生成器")
    st.markdown("输入文字或网址，一键生成可下载的二维码。纯本地、免 Key。")

    q_text = st.text_area("二维码内容（网址 / 文字 / 微信名片均可）", key="qr_text",
                          height=90, placeholder="例如：https://pandy-hu-tools-hub-app-ux8bug.streamlit.app")
    c1, c2 = st.columns(2)
    with c1:
        q_ec = st.selectbox("容错率", ["M", "L", "Q", "H"], index=0,
                            help="越高越抗污损，但图案越密。带 Logo 建议选 H。", key="qr_ec")
        q_box = st.slider("清晰度（像素密度）", 4, 16, 10, key="qr_box")
    with c2:
        q_fg = st.color_picker("前景色", "#000000", key="qr_fg")
        q_bg = st.color_picker("背景色", "#FFFFFF", key="qr_bg")
    q_logo = st.file_uploader("中心 Logo（可选，PNG/JPG）", type=["png", "jpg", "jpeg"],
                              key="qr_logo")

    if st.button("🚀 生成二维码", use_container_width=True, key="qr_go"):
        if not q_text.strip():
            st.warning("先填一下二维码内容～")
        else:
            with st.spinner("生成中…"):
                try:
                    logo_b = q_logo.getvalue() if q_logo else None
                    png, size = qr_generate(q_text, box_size=q_box, error_correction=q_ec,
                                            fg_color=q_fg, bg_color=q_bg, logo_bytes=logo_b)
                    st.image(png, width=min(360, size))
                    st.download_button("⬇️ 下载 PNG", data=png, file_name="qrcode.png",
                                       mime="image/png", use_container_width=True)
                except Exception as e:
                    st.error(f"生成失败：{e}")

# ---------------- Tab 7: 关于 / 升级 ----------------
if st.session_state["active_tool"] == "about":
    st.markdown("### 💡 关于这个小工具站")
    st.markdown("""
    这是一组**对标海外成功案例**的小工具（灵感来自 StarterStory）：

    - 📄 **PDF 转 Excel** —— 对标海外 $40K/月案例
    - 📊 **Excel 公式机器人** —— 对标海外 $20K/月案例
    - 📋 **中文 AI 简历生成器** —— 对标 Rezi $215K/月（中文差异化版）
    - 🛂 **中国护照签证政策聚合** —— 对标海外 $20K/月案例（信息差版）
    - 📈 **数据可视化小工具** —— 对标海外 $10K/月案例
    - 🗄️ **CSV → Airtable 导入** —— 对标海外 $20K/月案例（你带自己的 Token）
    - 🎨 **AI 配图 / 海报生成器** —— 对标社媒配图类小工具（$5K-20K/月，本地生成免 Key）
    - 🖼️ **图片处理工具箱** —— 对标海外图片压缩类小工具（$5K-20K/月，本地免 Key）
    - 🔳 **二维码生成器** —— 对标海外引流/营销类小工具（本地免 Key）

    全部免费可用。PDF / 签证 / 数据 / 导入 工具本地运行、无需本站 Key；公式 / 简历 填 Key 即联网、不填也能看演示。
    """)
    upgrade_url = st.secrets.get("UPGRADE_URL", "")
    st.markdown("### 🔓 升级高级版")
    if upgrade_url:
        st.markdown(f"解锁 **批量处理 / API 接口 / 去演示限制 / 定制功能**，点这里 👉 [{upgrade_url}]({upgrade_url})")
    else:
        st.info("高级版（批量处理、API 接口、去演示限制、定制功能）即将推出。\n\n"
                 "想抢先体验或提需求：可在部署平台的 Secrets 里配置 `UPGRADE_URL` 指向你的购买/联系方式链接，"
                 "本页会自动显示升级按钮。")

    # ---------------- 赞助 / 打赏 ----------------
    st.markdown("### ☕ 赞助 / 打赏")
    sponsor_url = st.secrets.get("SPONSOR_URL", "")
    sponsor_text = st.secrets.get("SPONSOR_TEXT", "如果这个工具帮到了你，请我喝杯奶茶 ☕")
    sponsor_icon = st.secrets.get("SPONSOR_ICON", "☕")
    if sponsor_url:
        st.markdown(f"""
        <div class="card" style="text-align:center;">
          <div style="font-size:15px;color:#1a2b4a;font-weight:600;">{sponsor_icon} {sponsor_text}</div>
          <a href="{sponsor_url}" target="_blank" rel="noopener">
            <button style="margin-top:12px;background:{PRIMARY};color:#fff;border:none;border-radius:10px;padding:11px 28px;font-weight:700;cursor:pointer;font-size:15px;">👉 去赞助</button>
          </a>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.info("若想开通赞助收款：在部署平台 Secrets 配置 `SPONSOR_URL`（你的收款页链接，如爱发电 / 奶茶）"
                 "与 `SPONSOR_TEXT`（按钮文案）、`SPONSOR_ICON`（图标 emoji），本页会自动出现赞助按钮。")
    st.markdown('<div class="foot">MVP · 由 WorkBuddy 生成 · 多个工具合成一站</div>', unsafe_allow_html=True)
