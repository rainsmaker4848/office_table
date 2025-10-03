# -*- coding: utf-8 -*-
import os
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# ================== НАСТРОЙКИ ==================
EXCEL_PATH = "people.xlsx"            # A: ФИО, B: Организация, C: Должность
BACKGROUND_IMAGE = "background.jpg"   # или .jpg
OUT_DIR = "out"
IMG_DIR = os.path.join(OUT_DIR, "img")

# Плашка 210×80 мм
PLAQUE_W_MM = 210
PLAQUE_H_MM = 80

# DPI для рендеринга (300 — печать)
DPI = 300

# ---- ШРИФТ ----
# Укажи файл шрифта PP Right Grotesk Text Wide (OTF/TTF)
FONT_PATH_PRIMARY = "PPRightGroteskTextWide-Regular.otf"
FONT_PATH_FALLBACKS = ["bahnschrift.ttf", "arial.ttf"]

# Базовые размеры (pt) — будут авто-подгоняться под ширину контейнера
# Организация крупнее всех; “чуть больше на один разряд” сохраняем через SCALE
PT_ORG_BASE     = 100
PT_SURNAME_BASE = 160
PT_NAMEP_BASE   = 120
PT_ROLE_BASE    = 60

SIZE_SCALE = 1.10  # +10% к базовым размерам

# Цвета
COLOR_ORG     = (255, 255, 255)
COLOR_SURNAME = (255, 255, 255)
COLOR_NAMEP   = (255, 255, 255)
COLOR_ROLE    = (255, 255, 255)

# Координаты центров строк (мм от верхнего края)
LMARGIN_MM = 12
RMARGIN_MM = 12
ORG_Y_MM     = 20
SURNAME_Y_MM = 37
NAMEP_Y_MM   = 54
ROLE_Y_MM    = 72

# Доп. межсимвольное расстояние (px) — 0 = выкл
TRACKING = 0
# =================================================


def mm_to_px(mm, dpi=DPI):
    return int(round(mm * dpi / 25.4))


def load_font(size_pt):
    """Пытаемся загрузить основной шрифт, затем запасные."""
    candidates = [FONT_PATH_PRIMARY] + FONT_PATH_FALLBACKS
    for fp in candidates:
        try:
            return ImageFont.truetype(fp, size=int(round(size_pt)))
        except Exception:
            continue
    return ImageFont.load_default()


def text_size(draw, text, font):
    if not text:
        return (0, 0)
    bbox = draw.textbbox((0, 0), text, font=font, anchor="lt")
    return (bbox[2] - bbox[0], bbox[3] - bbox[1])


def fit_font_to_width(draw, text, max_width_px, initial_pt):
    """Подгоняем шрифт под максимальную ширину контейнера (бинпоиск)."""
    if not text:
        return load_font(initial_pt)
    lo, hi = 10, max(10, int(round(initial_pt)))
    f = load_font(hi)
    w, _ = text_size(draw, text, f)
    if w <= max_width_px:
        return f
    best = load_font(lo)
    while lo <= hi:
        mid = (lo + hi) // 2
        fm = load_font(mid)
        w, _ = text_size(draw, text, fm)
        if w <= max_width_px:
            best = fm
            lo = mid + 1
        else:
            hi = mid - 1
    return best


def draw_centered_text(draw, center_x, center_y, text, font, color):
    if not text:
        return
    w, h = text_size(draw, text, font)
    x = int(center_x - w / 2)
    y = int(center_y - h / 2)
    if TRACKING and len(text) > 1:
        curr_x = x
        for ch in text:
            cw, _ = text_size(draw, ch, font)
            draw.text((curr_x, y), ch, font=font, fill=color)
            curr_x += cw + TRACKING
    else:
        draw.text((x, y), text, font=font, fill=color)


def split_fio_for_lines(fio_raw):
    """
    Делим ФИО на:
      1) Фамилия
      2) Имя Отчество (в одной строке)
    """
    parts = str(fio_raw).strip().split()
    if not parts:
        return ("", "")
    surname = parts[0].upper()
    name_patronymic = " ".join(parts[1:]).upper() if len(parts) > 1 else ""
    return surname, name_patronymic


def render_plate(background_path, fio, org, role, out_path):
    # размеры
    W = mm_to_px(PLAQUE_W_MM)
    H = mm_to_px(PLAQUE_H_MM)

    # фон
    bg = Image.open(background_path).convert("RGBA")
    bg = bg.resize((W, H), Image.LANCZOS)
    draw = ImageDraw.Draw(bg)

    # контейнер по ширине
    max_line_width = W - mm_to_px(LMARGIN_MM + RMARGIN_MM)
    cx = W // 2

    # тексты
    org_text = str(org).strip().upper()
    surname, namep = split_fio_for_lines(fio)
    role_text = str(role).strip()

    # базовые размеры с масштабом
    pt_org     = PT_ORG_BASE     * SIZE_SCALE
    pt_surname = PT_SURNAME_BASE * SIZE_SCALE
    pt_namep   = PT_NAMEP_BASE   * SIZE_SCALE
    pt_role    = PT_ROLE_BASE    * SIZE_SCALE

    # подбор кеглей под ширину
    f_org     = fit_font_to_width(draw, org_text,  max_line_width, pt_org)
    f_surname = fit_font_to_width(draw, surname,   max_line_width, pt_surname)
    f_namep   = fit_font_to_width(draw, namep,     max_line_width, pt_namep)
    f_role    = fit_font_to_width(draw, role_text, max_line_width, pt_role)

    # координаты по Y
    y_org     = mm_to_px(ORG_Y_MM)
    y_surname = mm_to_px(SURNAME_Y_MM)
    y_namep   = mm_to_px(NAMEP_Y_MM)
    y_role    = mm_to_px(ROLE_Y_MM)

    # рендер строк
    draw_centered_text(draw, cx, y_org,     org_text,  f_org,     COLOR_ORG)
    draw_centered_text(draw, cx, y_surname, surname,   f_surname, COLOR_SURNAME)
    draw_centered_text(draw, cx, y_namep,   namep,     f_namep,   COLOR_NAMEP)
    draw_centered_text(draw, cx, y_role,    role_text, f_role,    COLOR_ROLE)

    # сохранение
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    bg.save(out_path, format="PNG", dpi=(DPI, DPI))


def main():
    # читаем 3 колонки: ФИО, Организация, Должность
    df = pd.read_excel(EXCEL_PATH, header=None, usecols=[0, 1, 2])
    df.columns = ["fio", "org", "role"]

    os.makedirs(IMG_DIR, exist_ok=True)
    for i, row in df.iterrows():
        fio  = str(row["fio"]).strip()
        org  = str(row["org"]).strip()
        role = str(row["role"]).strip()
        safe = "_".join([p for p in (org + " " + fio).split() if p]).upper()
        img_path = os.path.join(IMG_DIR, f"{i:03d}_{safe}.png")
        render_plate(BACKGROUND_IMAGE, fio, org, role, img_path)

    print(f"✅ PNG готово: {len(df)} шт. → {IMG_DIR}")


if __name__ == "__main__":
    main()
