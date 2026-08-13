# PPTX Studio 🎨

[简体中文](README.md) | [English](README_EN.md)

> A highly customizable PowerPoint generation API — one JSON, stunning PPTX.

Built on **Flask + python-pptx**, it provides **32+ visualization components**, **12 professional themes**, **12 page decorations**, and free-form coordinate layout. Without any frontend tooling, you can generate presentations that rival hand-crafted designs using pure code (or curl).

---

## ✨ Features

- 🧩 **32 visualization components**: charts (bar / column / line / radar / funnel / gauge / pie / horizontal bar), cards / KPI / grids, pricing cards, timelines, step bars, flow diagrams, tag clouds, image galleries, team cards, comparison stats, and more
- 🎨 **12 professional themes**: Deep Space / Aurora / Sunset / Forest / Ocean / Cyber Neon / Azure Blue / Crimson Gold / Corporate / Minimal / Candy / Paper, with custom overrides supported
- ✨ **12 page decorations**: top bar / bottom bar / dot grid / slanted lines / color blob / glow / grid lines / side glow / waves / folded corner, etc.
- 🖼️ **16:9 widescreen** output (4:3 and custom sizes supported)
- 🌈 **Gradient backgrounds**, neon accent colors, rounded corners, shadows, transparency, CJK fonts (Microsoft YaHei)
- 🚀 **REST API**: render with a single curl command; built-in business templates out of the box

---

## 🚀 Quick Start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Start the service (default http://0.0.0.0:5002)
python app.py

# 3. Render a built-in template
curl -X POST http://127.0.0.1:5002/api/templates/demo_ai_weekly \
  -H "Content-Type: application/json" -d '{"theme":"deep_space"}'

# 4. Fully customize (use the sample payload)
curl -X POST http://127.0.0.1:5002/api/pptx/render \
  -H "Content-Type: application/json" -d @examples/payload.json
```

> The response returns a `filename` and `download_url`; open the latter in a browser to download the generated `.pptx`.

---

## 📡 API Overview

| Method | Path | Description |
|---|---|---|
| GET | `/api/health` | Health check |
| GET | `/api/themes` | List all themes |
| GET | `/api/components` | List all components |
| GET | `/api/templates` | List built-in sample templates |
| POST | `/api/templates/<name>` | Quickly render with a built-in template |
| POST | `/api/pptx/render` | **Render a highly customized PPTX** |
| GET | `/api/download/<filename>` | Download a generated file |
| GET | `/api/files` | List generated files |
| DELETE | `/api/files/<filename>` | Delete a file |

---

## 📄 Request Body Structure

```jsonc
{
  "theme": "deep_space",          // theme name (choose one of 12) or a custom dict
  "size": "16:9",                 // "16:9" | "4:3" | {"w":13.33,"h":7.5}
  "filename_prefix": "my_ppt",
  "pages": [
    {
      "type": "cover",            // cover | section | content | blank
      "kicker": "Kicker",         // small text at the top of cover / content pages
      "title": "Main Title",
      "subtitle": "Subtitle",
      "index": "01",              // large section number on section pages
      "background": {},           // override background (gradient/solid/image)
      "decorations": [],          // decoration array
      "page_number": true,
      "footer": "© 2026 Veshu",
      "components": [ /* component array */ ]
    }
  ]
}
```

### Custom Theme

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

### Common Component Fields

All components support positioning via `x` / `y` / `w` / `h` (in inches, or `"50%"` / `"fill"`); color fields accept theme keys
(`primary` / `secondary` / `accent` / `success` / `text` / `text_muted` …) or HEX values (`"FF5722"`).

---

## 🧩 Component List (32)

| Category | Components |
|---|---|
| **Charts** | `bar_chart` horizontal bar · `column_chart` column chart · `line_chart` line chart · `radar_chart` radar chart · `funnel_chart` funnel chart · `gauge` gauge · `pie_chart` pie / donut chart |
| **Layout** | `card` card · `kpi` KPI card · `grid_cards` card grid · `pricing` pricing card · `timeline` timeline · `steps` step bar · `flow` flow diagram |
| **Visual** | `title` title · `text` rich text · `list` list · `badge` badge · `divider` divider · `shape` shape · `icon` icon · `progress` progress bar · `quote` quote · `highlight` highlighted number · `cloud` tag cloud |
| **Media / Team** | `image` image · `carousel` image gallery · `team` team member · `stat_compare` comparison stats |
| **Page** | `footer` footer · `page_number` page number |

> `GET /api/components` returns the field documentation for each component.

---

## 🎨 Theme List (12)

| Theme id | Style | Best for |
|---|---|---|
| `deep_space` | Dark · Tech | AI / tech weekly reports |
| `aurora` | Dark · Dreamy | Creative / branding |
| `sunset` | Dark · Warm | Marketing / product launches |
| `forest` | Dark · Natural | Environmental / agriculture |
| `ocean` | Dark · Fresh | Project reviews / government |
| `neon` | Dark · Cyber | Gaming / geek culture |
| `azure_blue` | Dark · Business blue | Product launches / tech |
| `crimson_gold` | Dark · Luxurious | Business plans / fundraising |
| `corporate` | Light · Business | Daily reporting |
| `minimal` | Light · Minimal | Design / proposals |
| `candy` | Light · Playful | Education / events |
| `paper` | Light · Vintage | Culture / history |

## ✨ Decoration List (12)

`top_bar` top bar · `bottom_bar` bottom bar · `circle` dots · `dot_grid` dot grid · `lines` slanted lines · `blob` color blob · `corner` folded corner · `glow` glow · `grid` grid lines · `side_glow` side glow · `waves` waves

---

## 📁 Built-in Templates

| Template | Description | Pages |
|---|---|---|
| `demo_ai_weekly` | AI era weekly report (cover KPI / table of contents cards / sections / data charts / team / actions) | 11 |
| `demo_product_launch` | Product launch (glow decorations / process timeline / spec table / performance charts / roadmap) | 7 |
| `demo_business_plan` | Business plan (market charts / pricing cards / financial forecast / use of funds) | 10 |
| `demo_project_review` | Project review (KPI / line & radar charts / win-loss comparison / action steps / wave decorations) | 8 |

```bash
# Render a specific template, optionally overriding the theme
curl -X POST http://127.0.0.1:5002/api/templates/demo_product_launch \
  -H "Content-Type: application/json" -d '{"theme":"azure_blue"}'
```

---

## 📁 Directory Structure

```
pptx-studio/
├── app.py                    # Flask API entry point
├── requirements.txt          # dependencies
├── README.md                 # documentation (Chinese)
├── README_EN.md              # documentation (English)
├── LICENSE                   # MIT
├── .gitignore
├── core/
│   ├── themes.py             # 12 themes
│   ├── utils.py              # low-level rendering utilities (gradient/rounded corners/shadows/fonts/decorations)
│   ├── components.py         # 32 visualization components
│   └── engine.py             # rendering engine & page orchestration
├── examples/
│   ├── payload.json          # full sample request body
│   └── templates/            # built-in business template JSON
└── generated/                # output directory (auto-created, gitignored)
```

---

## 🛠️ Tech Stack

- [Flask](https://flask.palletsprojects.com/) — web framework
- [python-pptx](https://python-pptx.readthedocs.io/) — PPTX generation
- Native OOXML XML manipulation (gradients / shadows / rounded corners / transparency)

## 📄 License

[MIT](LICENSE)
