# PPTX Studio 🎨

[简体中文](README.md) | [English](README_EN.md)

> 高度定制化的 PowerPoint 生成 API —— 一条 JSON，生成精美 PPTX。

基于 **Flask + python-pptx**，提供 **32+ 可视化组件**、**12 套专业主题**、**12 种页面装饰**与自由坐标布局，无需任何前端工具，即可用纯代码（或 curl）生成媲美人工设计的演示文稿。

---

## ✨ 特性

- 🧩 **32 个可视化组件**：图表（柱状/折线/雷达/漏斗/仪表/饼图/条形）、卡片/KPI/网格、定价卡、时间线、步骤条、流程图、标签云、图片画廊、团队卡、对比统计等
- 🎨 **12 套专业主题**：深空科技 / 极光幻彩 / 日落橙紫 / 森林 / 海洋 / 赛博霓虹 / 蔚蓝科技 / 红金商务 / 商务简约 / 黑白极简 / 糖果甜心 / 纸感复古，支持自定义覆盖
- ✨ **12 种页面装饰**：顶条 / 底条 / 点阵 / 斜线 / 色块 / 光斑 / 网格线 / 侧边光 / 波浪 / 折角等
- 🖼️ **16:9 宽屏**输出（支持 4:3 与自定义尺寸）
- 🌈 **渐变背景**、霓虹强调色、圆角、阴影、透明度、中文字体（微软雅黑）
- 🚀 **REST API**：一行 curl 即可渲染；内置业务模板开箱即用

---

## 效果预览

<img width="1920" height="985" alt="image" src="https://github.com/user-attachments/assets/48b37ba6-c48f-4e26-b8c4-bf4f8446ca3a" />



## 🚀 快速开始

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 启动服务（默认 http://0.0.0.0:5002）
python app.py

# 3. 渲染内置模板
curl -X POST http://127.0.0.1:5002/api/templates/demo_ai_weekly \
  -H "Content-Type: application/json" -d '{"theme":"deep_space"}'

# 4. 自由定制（使用示例 payload）
curl -X POST http://127.0.0.1:5002/api/pptx/render \
  -H "Content-Type: application/json" -d @examples/payload.json
```

> 响应会返回 `filename` 与 `download_url`，浏览器打开即可下载生成的 `.pptx`。

---

## 📡 API 一览

| 方法 | 路径 | 说明 |
|---|---|---|
| GET | `/api/health` | 健康检查 |
| GET | `/api/themes` | 列出全部主题 |
| GET | `/api/components` | 列出全部组件 |
| GET | `/api/templates` | 列出内置示例模板 |
| POST | `/api/templates/<name>` | 用内置模板快速渲染 |
| POST | `/api/pptx/render` | **渲染高度定制 PPTX** |
| GET | `/api/download/<filename>` | 下载生成的文件 |
| GET | `/api/files` | 列出已生成的文件 |
| DELETE | `/api/files/<filename>` | 删除文件 |

---

## 📄 请求体结构

```jsonc
{
  "theme": "deep_space",          // 主题名（12 选 1）或自定义 dict
  "size": "16:9",                 // "16:9" | "4:3" | {"w":13.33,"h":7.5}
  "filename_prefix": "my_ppt",
  "pages": [
    {
      "type": "cover",            // cover | section | content | blank
      "kicker": "小标题",          // 封面/内容页顶部小字
      "title": "主标题",
      "subtitle": "副标题",
      "index": "01",              // 章节页大序号
      "background": {},           // 覆盖背景（gradient/solid/image）
      "decorations": [],          // 装饰数组
      "page_number": true,
      "footer": "© 2026 Veshu",
      "components": [ /* 组件数组 */ ]
    }
  ]
}
```

### 自定义主题

```json
{
  "theme": {
    "colors": {
      "primary": "FF5722",
      "secondary": "9C27B0"
    },
    "background": {"type": "gradient", "colors": ["1A1A2E", "16213E"], "angle": 120}
  }
}
```

### 组件通用字段

所有组件支持定位 `x` / `y` / `w` / `h`（英寸，或 `"50%"` / `"fill"`）；颜色字段支持主题键
（`primary` / `secondary` / `accent` / `success` / `text` / `text_muted` …）或 HEX（`"FF5722"`）。

---

## 🧩 组件清单（32 个）

| 分类 | 组件 |
|---|---|
| **图表** | `bar_chart` 条形图 · `column_chart` 柱状图 · `line_chart` 折线图 · `radar_chart` 雷达图 · `funnel_chart` 漏斗图 · `gauge` 仪表盘 · `pie_chart` 饼图/环形图 |
| **布局** | `card` 卡片 · `kpi` 指标卡 · `grid_cards` 卡片网格 · `pricing` 定价卡 · `timeline` 时间线 · `steps` 步骤条 · `flow` 流程图 |
| **视觉** | `title` 标题 · `text` 富文本 · `list` 列表 · `badge` 标签 · `divider` 分隔线 · `shape` 形状 · `icon` 图标 · `progress` 进度条 · `quote` 引用 · `highlight` 高亮数字 · `cloud` 标签云 |
| **媒体/团队** | `image` 图片 · `carousel` 图片画廊 · `team` 团队成员 · `stat_compare` 对比统计 |
| **页面** | `footer` 页脚 · `page_number` 页码 |

> `GET /api/components` 可获取每个组件的字段说明。

---

## 🎨 主题清单（12 套）

| 主题 id | 风格 | 适用场景 |
|---|---|---|
| `deep_space` | 深色 · 科技 | AI / 技术周报 |
| `aurora` | 深色 · 梦幻 | 创意 / 品牌 |
| `sunset` | 深色 · 暖色 | 营销 / 发布会 |
| `forest` | 深色 · 自然 | 环保 / 农业 |
| `ocean` | 深色 · 清新 | 项目复盘 / 政务 |
| `neon` | 深色 · 赛博 | 游戏 / 极客 |
| `azure_blue` | 深色 · 商务蓝 | 产品发布 / 科技 |
| `crimson_gold` | 深色 · 华贵 | 商业计划 / 融资 |
| `corporate` | 浅色 · 商务 | 日常汇报 |
| `minimal` | 浅色 · 极简 | 设计 / 方案 |
| `candy` | 浅色 · 活泼 | 教育 / 活动 |
| `paper` | 浅色 · 复古 | 文化 / 历史 |

## ✨ 装饰清单（12 种）

`top_bar` 顶条 · `bottom_bar` 底条 · `circle` 圆点 · `dot_grid` 点阵 · `lines` 斜线 · `blob` 色块 · `corner` 折角 · `glow` 光斑 · `grid` 网格线 · `side_glow` 侧边光 · `waves` 波浪

---

## 📁 内置模板

| 模板 | 说明 | 页数 |
|---|---|---|
| `demo_ai_weekly` | AI 时代周报（封面 KPI / 目录卡片 / 章节 / 数据图表 / 团队 / 行动） | 11 |
| `demo_product_launch` | 产品发布会（光斑装饰 / 流程时间线 / 规格表 / 性能图表 / 路线图） | 7 |
| `demo_business_plan` | 商业计划书（市场图表 / 定价卡 / 财务预测 / 融资用途） | 10 |
| `demo_project_review` | 项目复盘（KPI / 折线雷达 / 得失对比 / 行动步骤 / 波浪装饰） | 8 |

```bash
# 渲染指定模板，可覆盖主题
curl -X POST http://127.0.0.1:5002/api/templates/demo_product_launch \
  -H "Content-Type: application/json" -d '{"theme":"azure_blue"}'
```

---

## 📁 目录结构

```
pptx-studio/
├── app.py                    # Flask API 入口
├── requirements.txt          # 依赖
├── README.md
├── README_EN.md              # English documentation
├── LICENSE                   # MIT
├── .gitignore
├── core/
│   ├── themes.py             # 12 套主题
│   ├── utils.py              # 底层渲染工具（渐变/圆角/阴影/字体/装饰）
│   ├── components.py         # 32 个可视化组件
│   └── engine.py             # 渲染引擎与页面编排
├── examples/
│   ├── payload.json          # 完整示例请求体
│   └── templates/            # 内置业务模板 JSON
└── generated/                # 输出目录（自动创建，已 gitignore）
```

---

## 🛠️ 技术栈

- [Flask](https://flask.palletsprojects.com/) —— Web 框架
- [python-pptx](https://python-pptx.readthedocs.io/) —— PPTX 生成
- OOXML 原生 XML 操作（渐变 / 阴影 / 圆角 / 透明度）

## 📄 License

[MIT](LICENSE)
