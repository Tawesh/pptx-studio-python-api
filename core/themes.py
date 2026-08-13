# -*- coding: utf-8 -*-
"""主题系统：10 套预置主题，每套含渐变背景 + 完整配色 + 装饰偏好。

主题结构：
{
  "id":         主题标识（API 中 theme 字段使用）
  "label":      中文名
  "style":      "dark" | "light"   （用于判断文字默认用浅色还是深色）
  "background": 默认背景（gradient / solid / image）
  "colors":     配色：
      primary      主色
      secondary    辅色
      accent       强调色（点缀/高亮）
      success/warning/danger 语义色
      text         正文文字
      text_muted   次级文字
      text_invert  深色/浅色主题下用于主色上的文字
      card_bg      卡片底色
      card_bg2     卡片交替底色
      card_border  卡片描边
      overlay      图片遮罩色
  "fonts":      {"heading": 标题字体, "body": 正文字体}
  "decoration": 页面默认装饰
}
"""
import copy

THEMES = {}

def _reg(t):
    THEMES[t["id"]] = t


# ---------------------------------------------------------------- 深空科技
_reg({
    "id": "deep_space",
    "label": "深空科技",
    "style": "dark",
    "background": {"type": "gradient", "colors": ["0B1026", "261B52"], "angle": 120},
    "colors": {
        "primary": "00E5FF", "secondary": "B44CFF", "accent": "FFD54F",
        "success": "4CD964", "warning": "FFB300", "danger": "FF5252",
        "text": "FFFFFF", "text_muted": "9AA3C7", "text_invert": "0B1026",
        "card_bg": "151B3C", "card_bg2": "1E2652", "card_border": "2E3A6E",
        "overlay": "0B1026",
    },
    "fonts": {"heading": "微软雅黑", "body": "微软雅黑"},
    "decoration": {"type": "dot_grid", "color": "00E5FF", "alpha": 18, "density": 0.07},
})

# ---------------------------------------------------------------- 极光
_reg({
    "id": "aurora",
    "label": "极光幻彩",
    "style": "dark",
    "background": {"type": "gradient", "colors": ["031B1B", "0E3B43", "24305E"], "angle": 135},
    "colors": {
        "primary": "3DDC97", "secondary": "7FD8BE", "accent": "E58BFF",
        "success": "3DDC97", "warning": "FFC857", "danger": "FF6B6B",
        "text": "F0FFF9", "text_muted": "9FC7BC", "text_invert": "031B1B",
        "card_bg": "0C2B30", "card_bg2": "12393F", "card_border": "1D4A52",
        "overlay": "031B1B",
    },
    "fonts": {"heading": "微软雅黑", "body": "微软雅黑"},
    "decoration": {"type": "blob", "color": "3DDC97", "alpha": 14, "count": 2},
})

# ---------------------------------------------------------------- 日落
_reg({
    "id": "sunset",
    "label": "日落橙紫",
    "style": "dark",
    "background": {"type": "gradient", "colors": ["2B1055", "5B1E5E", "B24A4A"], "angle": 150},
    "colors": {
        "primary": "FF9E64", "secondary": "FF6B9D", "accent": "FFD166",
        "success": "7AE582", "warning": "FFC857", "danger": "FF5252",
        "text": "FFF5EC", "text_muted": "D9A79E", "text_invert": "2B1055",
        "card_bg": "3A1660", "card_bg2": "472070", "card_border": "5A2F82",
        "overlay": "2B1055",
    },
    "fonts": {"heading": "微软雅黑", "body": "微软雅黑"},
    "decoration": {"type": "lines", "color": "FF9E64", "alpha": 16, "count": 3},
})

# ---------------------------------------------------------------- 森林
_reg({
    "id": "forest",
    "label": "森林自然",
    "style": "dark",
    "background": {"type": "gradient", "colors": ["0A1F1C", "16483B"], "angle": 110},
    "colors": {
        "primary": "7AC74F", "secondary": "A8E063", "accent": "FFD166",
        "success": "7AC74F", "warning": "FFC857", "danger": "FF6B6B",
        "text": "F2FBE8", "text_muted": "A9C9A0", "text_invert": "0A1F1C",
        "card_bg": "12332B", "card_bg2": "174038", "card_border": "25584B",
        "overlay": "0A1F1C",
    },
    "fonts": {"heading": "微软雅黑", "body": "微软雅黑"},
    "decoration": {"type": "blob", "color": "7AC74F", "alpha": 12, "count": 2},
})

# ---------------------------------------------------------------- 海洋
_reg({
    "id": "ocean",
    "label": "海洋深蓝",
    "style": "dark",
    "background": {"type": "gradient", "colors": ["04222E", "0E5A7A"], "angle": 120},
    "colors": {
        "primary": "4DD0FF", "secondary": "7CFFB2", "accent": "FFD166",
        "success": "7CFFB2", "warning": "FFC857", "danger": "FF6B6B",
        "text": "EAF8FF", "text_muted": "8FBCCF", "text_invert": "04222E",
        "card_bg": "0C3A50", "card_bg2": "10445E", "card_border": "1B5C7C",
        "overlay": "04222E",
    },
    "fonts": {"heading": "微软雅黑", "body": "微软雅黑"},
    "decoration": {"type": "dot_grid", "color": "4DD0FF", "alpha": 15, "density": 0.06},
})

# ---------------------------------------------------------------- 赛博霓虹
_reg({
    "id": "neon",
    "label": "赛博霓虹",
    "style": "dark",
    "background": {"type": "gradient", "colors": ["0D0221", "2A0E4A", "0D0221"], "angle": 100},
    "colors": {
        "primary": "FF2E88", "secondary": "00F0FF", "accent": "B7FF00",
        "success": "00F0FF", "warning": "FFD166", "danger": "FF5252",
        "text": "FFFFFF", "text_muted": "B7A6D9", "text_invert": "0D0221",
        "card_bg": "1B0A3A", "card_bg2": "240F4D", "card_border": "3D1F6E",
        "overlay": "0D0221",
    },
    "fonts": {"heading": "微软雅黑", "body": "微软雅黑"},
    "decoration": {"type": "lines", "color": "FF2E88", "alpha": 20, "count": 4},
})

# ---------------------------------------------------------------- 商务
_reg({
    "id": "corporate",
    "label": "商务简约",
    "style": "light",
    "background": {"type": "solid", "color": "FFFFFF"},
    "colors": {
        "primary": "1F4E79", "secondary": "2E74B5", "accent": "FFB900",
        "success": "2E9E5B", "warning": "F39C12", "danger": "D9534F",
        "text": "2B2B2B", "text_muted": "6B7A90", "text_invert": "FFFFFF",
        "card_bg": "F4F7FB", "card_bg2": "EAEFF7", "card_border": "D8E1EE",
        "overlay": "1F4E79",
    },
    "fonts": {"heading": "微软雅黑", "body": "微软雅黑"},
    "decoration": {"type": "top_bar", "color": "1F4E79", "alpha": 100},
})

# ---------------------------------------------------------------- 极简
_reg({
    "id": "minimal",
    "label": "黑白极简",
    "style": "light",
    "background": {"type": "solid", "color": "FFFFFF"},
    "colors": {
        "primary": "111111", "secondary": "555555", "accent": "E53935",
        "success": "2E7D32", "warning": "F9A825", "danger": "C62828",
        "text": "1A1A1A", "text_muted": "8A8A8A", "text_invert": "FFFFFF",
        "card_bg": "F5F5F5", "card_bg2": "EBEBEB", "card_border": "DDDDDD",
        "overlay": "111111",
    },
    "fonts": {"heading": "微软雅黑", "body": "微软雅黑"},
    "decoration": {"type": "top_bar", "color": "111111", "alpha": 100},
})

# ---------------------------------------------------------------- 糖果
_reg({
    "id": "candy",
    "label": "糖果甜心",
    "style": "light",
    "background": {"type": "gradient", "colors": ["FFE3EC", "FDF0D5"], "angle": 120},
    "colors": {
        "primary": "FF4D6D", "secondary": "7A2EFF", "accent": "FF9E00",
        "success": "3EC300", "warning": "FFB300", "danger": "FF5252",
        "text": "5C1130", "text_muted": "C76A8A", "text_invert": "FFFFFF",
        "card_bg": "FFFFFF", "card_bg2": "FFF0F3", "card_border": "FFD9E2",
        "overlay": "FF4D6D",
    },
    "fonts": {"heading": "微软雅黑", "body": "微软雅黑"},
    "decoration": {"type": "blob", "color": "FF4D6D", "alpha": 12, "count": 2},
})

# ---------------------------------------------------------------- 纸张
_reg({
    "id": "paper",
    "label": "纸感复古",
    "style": "light",
    "background": {"type": "solid", "color": "FBF6EC"},
    "colors": {
        "primary": "8B5E34", "secondary": "C08552", "accent": "D64541",
        "success": "6B8E23", "warning": "E9A00A", "danger": "C0392B",
        "text": "3E2F23", "text_muted": "9A8570", "text_invert": "FFFFFF",
        "card_bg": "FFFDF6", "card_bg2": "F6EFE0", "card_border": "E4D6BE",
        "overlay": "3E2F23",
    },
    "fonts": {"heading": "楷体", "body": "微软雅黑"},
    "decoration": {"type": "corner", "color": "8B5E34", "alpha": 100},
})


# ---------------------------------------------------------------- 蔚蓝科技
_reg({
    "id": "azure_blue",
    "label": "蔚蓝科技",
    "style": "dark",
    "background": {"type": "gradient", "colors": ["001B33", "003E6B", "0A5FA8"], "angle": 120},
    "colors": {
        "primary": "38BDF8", "secondary": "818CF8", "accent": "FBBF24",
        "success": "34D399", "warning": "FBBF24", "danger": "F87171",
        "text": "F0F9FF", "text_muted": "93B4CC", "text_invert": "001B33",
        "card_bg": "082F4F", "card_bg2": "0C3A60", "card_border": "14547F",
        "overlay": "001B33",
    },
    "fonts": {"heading": "微软雅黑", "body": "微软雅黑"},
    "decoration": {"type": "side_glow", "color": "38BDF8", "alpha": 16, "side": "left"},
})

# ---------------------------------------------------------------- 红金商务
_reg({
    "id": "crimson_gold",
    "label": "红金商务",
    "style": "dark",
    "background": {"type": "gradient", "colors": ["1A0A0A", "4A0E0E", "8C2B2B"], "angle": 130},
    "colors": {
        "primary": "F5C542", "secondary": "E8A33D", "accent": "FFD700",
        "success": "7FD962", "warning": "F5A623", "danger": "FF5252",
        "text": "FFF8E7", "text_muted": "D9B896", "text_invert": "1A0A0A",
        "card_bg": "2E1313", "card_bg2": "3A1818", "card_border": "5A2626",
        "overlay": "1A0A0A",
    },
    "fonts": {"heading": "微软雅黑", "body": "微软雅黑"},
    "decoration": {"type": "glow", "color": "F5C542", "alpha": 18, "x_frac": 0.85, "y_frac": 0.2},
})


def get_theme(name):
    """按名称取主题；不存在则回退到 deep_space。"""
    if isinstance(name, dict):
        # 支持内联自定义主题：深拷贝 deep_space 再覆盖
        base = copy.deepcopy(THEMES.get("deep_space"))
        if "colors" in name:
            base["colors"].update(name["colors"])
        if "background" in name:
            base["background"] = name["background"]
        if "fonts" in name:
            base["fonts"].update(name["fonts"])
        return base
    return THEMES.get(str(name or "").strip().lower(), THEMES["deep_space"])


def list_themes():
    """返回主题元信息列表（供 /api/themes）。"""
    return [
        {"id": t["id"], "label": t["label"], "style": t["style"]}
        for t in THEMES.values()
    ]
