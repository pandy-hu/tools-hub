# 🧰 小工具站 (Tools Hub)

一组对标海外成功案例的小工具（含中国护照信息差版），合成一个网站，可一键部署到云端获得公开网址。

| 工具 | 对标案例 | 是否需要 Key |
|------|----------|--------------|
| 📄 PDF 发票/账单 → Excel | 海外 $40K/月案例 | 不需要（本地运行） |
| 📊 Excel 公式生成机器人 | 海外 $20K/月案例 | 填 Key 联网，不填可演示 |
| 📋 中文 AI 简历生成器 | Rezi $215K/月（中文差异化） | 填 Key 联网，不填可演示 |
| 🛂 中国护照签证政策聚合 | 海外 $20K/月案例（中文信息差版） | 不需要（内置数据集） |
| 📈 数据可视化小工具 | 海外 $10K/月案例 | 不需要（本地运行，上传 CSV/Excel 出图） |
| 🗄️ CSV → Airtable 导入 | 海外 $20K/月案例 | 需要你自己的 Airtable Token（数据直连 Airtable） |
| 🎨 AI 配图 / 海报生成器 | 社媒配图类小工具（$5K-20K/月） | 不需要（本地生成）；填 Key 可让 AI 写文案 |
| 🖼️ 图片处理工具箱 | 图片压缩类小工具（$5K-20K/月） | 不需要（本地压缩/转格式/改尺寸/水印） |
| 🔳 二维码生成器 | 引流/营销类小工具 | 不需要（本地生成，可选中心 Logo / 配色 / 容错率） |
| 🔊 文字转语音 TTS | 内容/音频类小工具 | 不需要（微软免费接口，免 Key，联网即可用） |
| 📑 PDF 合并 / 加水印 | PDF 类小工具（$40K/月级） | 不需要（本地合并多个 PDF / 批量加平铺文字水印） |
| ✍️ 多平台文案改写 | 内容营销 / 引流类小工具 | 填 Key 联网改写，不填可演示（小红书/公众号/抖音/微博/知乎一键切换） |
| 🔗 短链接生成器 | 引流 / 分享类小工具 | 不需要（调用免费公共接口 tinyurl / cleanuri，免 Key 免注册） |
| 🧹 图片智能去背景 | 封面 / 电商图 / 贴纸类小工具 | 不需要（本地 AI 模型 rembg 抠图，文件不上传） |
| 🧩 九宫格切图 | 小红书引流 / 内容发布类小工具 | 不需要（纯本地 Pillow 切图，零依赖、零上传） |

> 数据来源灵感：StarterStory.com 案例库。

---

## 一、本地运行（先在自己电脑上试）

```bash
# 用配好依赖的 Python 环境
C:/Users/NING MEI/.workbuddy/binaries/python/envs/default/Scripts/python.exe -m pip install -r requirements.txt
C:/Users/NING MEI/.workbuddy/binaries/python/envs/default/Scripts/python.exe -m streamlit run app.py --server.port 8501
```

浏览器打开 http://localhost:8501 即可使用。

---

## 二、部署到 Streamlit Cloud（获得公开网址，免费）

### 第 1 步：把代码推到 GitHub
- 仓库需包含：`app.py`、`core.py`、`requirements.txt`
- 在 GitHub 新建一个**公开**仓库（如 `tools-hub`）

### 第 2 步：登录 Streamlit Cloud
- 打开 https://share.streamlit.io
- 用你的 **GitHub 账号**登录（pandy-hu）
- 点 **"New app"** → 选择刚推的仓库 → 主文件填 `app.py`
- 点 **Deploy** → 等待 1~2 分钟

### 第 3 步：拿到网址
部署完成后会给你一个公开网址，形如：
```
https://xxxx-tools-hub.streamlit.app
```
任何人都能打开使用。把链接发出去就是流量入口。

---

## 三、配置 Secrets（可选但重要）

在 Streamlit Cloud 的 **App → Manage app → Settings → Secrets** 里可填。
格式是 TOML（每行 `键名 = "值"`）：

```toml
# —— AI 密钥（填了之后全站访客免粘贴、直接联网；不填则访客需自己粘贴或走演示）——
DEEPSEEK_API_KEY = "sk-你的deepseek密钥"
OPENAI_API_KEY   = "sk-你的openai密钥"
DASHSCOPE_API_KEY = "你的通义密钥"   # 仅用通义时填

# —— 赞助 / 打赏链接（填了站内自动出现赞助按钮）——
SPONSOR_URL   = "https://afdian.com/a/你的主页"   # 爱发电/奶茶等收款页
SPONSOR_TEXT  = "如果这个工具帮到了你，请我喝杯奶茶 ☕"
SPONSOR_ICON  = "☕"

# —— 升级高级版入口（可选）——
UPGRADE_URL   = "https://你的购买或联系页"
```

| Secret 键名 | 对应服务商 | 作用 |
|-----|------|------|
| `DEEPSEEK_API_KEY` | DeepSeek（默认推荐） | 公式/简历工具联网生成 |
| `OPENAI_API_KEY` | OpenAI | 同上，用 OpenAI 时 |
| `DASHSCOPE_API_KEY` | 通义千问 | 同上，用通义时 |
| `CUSTOM_API_KEY` | 自定义 | 选"自定义"服务商时 |
| `SPONSOR_URL` | —— | 站内「赞助」按钮跳转的收款页 |
| `SPONSOR_TEXT` / `SPONSOR_ICON` | —— | 赞助区文案与图标（可选） |
| `UPGRADE_URL` | —— | 「升级高级版」按钮链接（可选） |

> 不填 Secrets 也能正常运行：PDF 工具免 Key；公式/简历工具访客自己粘贴 Key，或留空走演示模式。
> 把 Key 放进 Secrets 后，**所有访客都能直接联网用**，但会消耗你的额度——按需选择。

---

## 四、怎么"收钱"（商业闭环）

部署上线只是拿到流量入口，**收钱需要额外设计**：

1. **免费引流 + 高级版付费**（推荐起步）
   - 基础功能免费（现在的样子）
   - 高级版收费：批量处理、API 接口、去演示限制、定制模板
   - 在 Secrets 配置 `UPGRADE_URL` 指向你的支付/联系页

2. **引流到私域再成交**（零代码）
   - 工具站放公众号二维码 / 微信
   - 免费工具吸引精准用户（想转 Excel、写简历、学公式的人）
   - 在社群/朋友圈卖课、卖服务、接定制

3. **广告 / 联盟**（流量大了再做）
   - 接 Google AdSense 或国内联盟广告

> 下一步可加：用户系统、用量统计、Stripe / 微信支付对接、批量任务队列。

---

## 五、目录结构

```
tools_hub/
├── app.py            # 统一界面（多工具按钮导航 + 关于/升级页）
├── core.py           # 各工具核心逻辑（与 UI 解耦）
├── requirements.txt  # 依赖
└── README.md         # 本文件
```
