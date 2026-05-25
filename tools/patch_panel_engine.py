#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
NetDTL Patch Panel Engine
Moteur générique de génération PDF A3 depuis inventaire Excel.
"""

import os
import math
import textwrap
import unicodedata
from typing import Dict, List, Any

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.pagesizes import A3
from reportlab.pdfgen import canvas as rl_canvas


# ============================================================
# CONFIG
# ============================================================

DPI = 150
A3_W_PT, A3_H_PT = A3
IMG_W = int(A3_W_PT / 72 * DPI)
IMG_H = int(A3_H_PT / 72 * DPI)

PORT_W = 62
PORT_H = 52
PORT_GAP_X = 5
PORT_GAP_Y = 12

MARGIN_LEFT = 80
MARGIN_RIGHT = 60
MARGIN_TOP = 130
MARGIN_BOTTOM = 80

PANEL_GAP = 80

BUBBLES_PER_ROW = 6
BUBBLE_H = 130
BUBBLE_GAP = 8

BG_COLOR = (28, 32, 38)
PANEL_BG = (40, 46, 56)
PANEL_BORDER = (60, 70, 88)

PORT_EMPTY = (55, 62, 75)
PORT_EMPTY_BORDER = (80, 95, 115)

PORT_ACTIVE_RJ45 = (30, 120, 200)
PORT_ACTIVE_RJ45_B = (100, 180, 255)

PORT_ACTIVE_RJ11 = (160, 100, 20)
PORT_ACTIVE_RJ11_B = (220, 160, 60)

PORT_INACTIVE = (80, 90, 105)
PORT_INACTIVE_B = (80, 95, 115)

BUBBLE_BG = (48, 56, 70)
BUBBLE_BORDER = (90, 130, 200)
BUBBLE_BORDER_RJ11 = (200, 150, 60)

CONNECTOR_FILL = (20, 25, 35)
CONNECTOR_BORDER = (100, 115, 135)

WHITE = (255, 255, 255)
TEXT_IP = (120, 200, 150)
TEXT_MULTI_BADGE = (255, 200, 80)
SEPARATOR_COLOR = (55, 65, 82)

PANEL_CONFIG = {
    "A": {"cols": 18, "slots": 18},
    "B": {"cols": 24, "slots": 48},
    "X": {"cols": 12, "slots": 24}
}


# ============================================================
# HELPERS
# ============================================================

class PatchPanelError(Exception):
    pass


def clean_value(value):
    if pd.isna(value):
        return ""

    s = str(value).strip()

    if s.lower() in ("nan", "nat", "none"):
        return ""

    return s


def normalize_text(text):
    text = clean_value(text)

    if not text:
        return ""

    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = text.lower().strip()

    return text


def wrap_text(text, max_chars):
    if not text:
        return []

    return textwrap.wrap(str(text), width=max_chars)


def detect_port_group(port_name):
    if not port_name:
        return "X"

    prefix = port_name[0].upper()

    if prefix in ("A", "B", "X"):
        return prefix

    return "X"


def parse_port_number(port_name):
    if not port_name:
        return 9999

    port_name = clean_value(port_name)

    if not port_name:
        return 9999

    prefix = detect_port_group(port_name)

    digits = "".join(filter(str.isdigit, port_name))

    if not digits:
        return 9999

    number = int(digits)

    base = {
        "A": 0,
        "B": 100,
        "X": 200
    }.get(prefix, 500)

    return base + number


# ============================================================
# FONTS
# ============================================================

def find_font(candidates, size):
    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                continue

    return ImageFont.load_default()


def load_fonts():
    bold = [
        "C:/Windows/Fonts/segoeuib.ttf",
        "C:/Windows/Fonts/arialbd.ttf",
    ]

    regular = [
        "C:/Windows/Fonts/segoeui.ttf",
        "C:/Windows/Fonts/arial.ttf",
    ]

    mono = [
        "C:/Windows/Fonts/consola.ttf",
        "C:/Windows/Fonts/cour.ttf",
    ]

    return {
        "title": find_font(bold, 26),
        "port": find_font(bold, 13),
        "small": find_font(regular, 11),
        "bubble": find_font(bold, 12),
        "ip": find_font(mono, 11),
        "label": find_font(regular, 12),
        "switch": find_font(bold, 12),
    }


# ============================================================
# VALIDATION
# ============================================================

REQUIRED_MAPPING = [
    "prise",
    "hostname",
    "ip",
]


OPTIONAL_MAPPING = [
    "type",
    "local",
    "notes",
]


def validate_mapping(df, mapping):
    if not mapping:
        raise PatchPanelError("Mapping absent.")

    columns = list(df.columns)

    for key in REQUIRED_MAPPING:
        if key not in mapping:
            raise PatchPanelError(f"Champ obligatoire absent : {key}")

        col = mapping[key]

        if not col:
            raise PatchPanelError(f"Colonne non définie : {key}")

        if col not in columns:
            raise PatchPanelError(
                f"Colonne introuvable dans Excel : {col}"
            )


# ============================================================
# LOAD INVENTORY
# ============================================================

def merge_duplicate_ports(existing, new_entry):
    if "_multi" not in existing:
        existing["_multi"] = [dict(existing)]

    existing["_multi"].append(new_entry)

    if new_entry["hostname"] and not existing["hostname"]:
        existing["hostname"] = new_entry["hostname"]

    if new_entry["ip"] and not existing["ip"]:
        existing["ip"] = new_entry["ip"]

    if new_entry["notes"] and not existing["notes"]:
        existing["notes"] = new_entry["notes"]


def load_inventory(excel_path, sheet_name, mapping):
    if not os.path.exists(excel_path):
        raise PatchPanelError("Fichier Excel introuvable.")

    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    if df.empty:
        raise PatchPanelError("Feuille Excel vide.")

    validate_mapping(df, mapping)

    result = {
        "A": [],
        "B": [],
        "X": []
    }

    seen_ports = {}

    for _, row in df.iterrows():
        prise = clean_value(row[mapping["prise"]])

        if not prise:
            continue

        entry = {
            "prise": prise,
            "hostname": clean_value(row[mapping["hostname"]]),
            "ip": clean_value(row[mapping["ip"]]),
            "type": clean_value(row[mapping.get("type", "")]) if mapping.get("type") else "",
            "local": clean_value(row[mapping.get("local", "")]) if mapping.get("local") else "",
            "notes": clean_value(row[mapping.get("notes", "")]) if mapping.get("notes") else "",
        }

        group = detect_port_group(prise)

        if prise in seen_ports:
            merge_duplicate_ports(seen_ports[prise], entry)
        else:
            seen_ports[prise] = entry
            result[group].append(entry)

    for group in result:
        result[group] = sorted(
            result[group],
            key=lambda x: parse_port_number(x["prise"])
        )

    return result
	# ============================================================
# SLOT ASSIGNMENT / DATA PREPARATION
# ============================================================

def assign_slots(port_list, total_slots):
    """
    Place les prises dans un tableau fixe de slots.
    """

    slots = [None] * total_slots

    for idx, port in enumerate(port_list):
        if idx >= total_slots:
            break

        slots[idx] = port

    return slots


def has_display_data(port):
    """
    Détermine si une prise mérite une bulle descriptive.
    """

    if not port:
        return False

    return any([
        port.get("hostname"),
        port.get("ip"),
        port.get("notes"),
        port.get("local"),
    ])


def prepare_panels(inventory):
    """
    Transforme l'inventaire en structure exploitable
    pour le moteur de dessin.
    """

    panels = []

    for group in ("A", "B", "X"):
        ports = inventory.get(group, [])

        if not ports:
            continue

        config = PANEL_CONFIG[group]

        slots = assign_slots(
            ports,
            config["slots"]
        )

        panels.append({
            "group": group,
            "label": f"Panneau {group}",
            "cols": config["cols"],
            "slots": config["slots"],
            "slot_data": slots,
        })

    return panels


def compute_panel_height(panel):
    """
    Calcule hauteur totale nécessaire.
    """

    total_slots = panel["slots"]
    cols = panel["cols"]
    slots = panel["slot_data"]

    n_rows = math.ceil(total_slots / cols)

    panel_height = (
        n_rows * (PORT_H + PORT_GAP_Y)
        - PORT_GAP_Y
        + 24
    )

    annotated = [
        s for s in slots
        if has_display_data(s)
    ]

    n_bubble_rows = (
        math.ceil(len(annotated) / BUBBLES_PER_ROW)
        if annotated else 0
    )

    bubble_height = (
        n_bubble_rows * (BUBBLE_H + BUBBLE_GAP + 24)
        if n_bubble_rows else 0
    )

    total_height = 45 + panel_height + 35 + bubble_height

    return {
        "panel_height": panel_height,
        "bubble_height": bubble_height,
        "total_height": total_height,
        "annotated_count": len(annotated)
    }


def compute_layout(panels):
    """
    Positionnement vertical complet.
    """

    cursor_y = MARGIN_TOP
    layout = []

    for panel in panels:
        metrics = compute_panel_height(panel)

        layout.append({
            "panel": panel,
            "y": cursor_y,
            "metrics": metrics
        })

        cursor_y += metrics["total_height"] + PANEL_GAP

    return layout


def normalize_inventory_for_render(inventory):
    """
    Pipeline complète avant rendu.
    """

    panels = prepare_panels(inventory)

    if not panels:
        raise PatchPanelError(
            "Aucune prise exploitable trouvée."
        )

    return compute_layout(panels)
	# ============================================================
# DRAWING ENGINE
# ============================================================

def rr(draw, xy, radius, fill=None, outline=None, width=1):
    """
    Rounded rectangle helper
    """
    draw.rounded_rectangle(
        xy,
        radius=radius,
        fill=fill,
        outline=outline,
        width=width
    )


def port_colors(port):
    """
    Détermine le style graphique d'une prise.
    """

    if port is None:
        return PORT_EMPTY, PORT_EMPTY_BORDER

    ptype = clean_value(port.get("type"))
    has_host = bool(port.get("hostname") or port.get("ip"))

    if ptype.upper() == "RJ45" and has_host:
        return PORT_ACTIVE_RJ45, PORT_ACTIVE_RJ45_B

    if ptype.upper() == "RJ11":
        return PORT_ACTIVE_RJ11, PORT_ACTIVE_RJ11_B

    if port.get("local") or port.get("notes"):
        return PORT_INACTIVE, PORT_INACTIVE_B

    return PORT_EMPTY, PORT_EMPTY_BORDER


def draw_port(draw, fonts, port, px, py):
    """
    Dessin d'une prise individuelle.
    """

    fill_col, border_col = port_colors(port)

    rr(
        draw,
        [px, py, px + PORT_W - 2, py + PORT_H - 2],
        radius=5,
        fill=fill_col,
        outline=border_col,
        width=1
    )

    cx = px + PORT_W // 2
    cy = py + PORT_H // 2

    cw = 22
    ch = 12

    draw.rectangle(
        [
            cx - cw // 2,
            cy - ch // 2 - 5,
            cx + cw // 2,
            cy + ch // 2 - 5
        ],
        fill=CONNECTOR_FILL,
        outline=CONNECTOR_BORDER,
        width=1
    )

    label = port["prise"] if port else "—"

    lw = draw.textlength(label, font=fonts["port"])

    draw.text(
        (cx - lw / 2, py + PORT_H - 18),
        label,
        fill=WHITE,
        font=fonts["port"]
    )


def draw_bubble(draw, fonts, port, bx, by, bubble_w):
    """
    Dessin bulle descriptive.
    """

    ptype = clean_value(port.get("type"))

    border = (
        BUBBLE_BORDER_RJ11
        if ptype.upper() == "RJ11"
        else BUBBLE_BORDER
    )

    rr(
        draw,
        [bx, by, bx + bubble_w, by + BUBBLE_H],
        radius=6,
        fill=BUBBLE_BG,
        outline=border,
        width=1
    )

    prise = port["prise"]
    local = port.get("local", "")
    hostname = port.get("hostname", "")
    ip = port.get("ip", "")
    notes = port.get("notes", "")

    draw.text(
        (bx + 9, by + 10),
        prise,
        fill=WHITE,
        font=fonts["bubble"]
    )

    if local:
        lw = draw.textlength(prise, font=fonts["bubble"])

        draw.text(
            (bx + 12 + lw, by + 11),
            f" {local}",
            fill=WHITE,
            font=fonts["small"]
        )

    if hostname:
        hl = wrap_text(hostname, 24)

        draw.text(
            (bx + 9, by + 36),
            hl[0],
            fill=WHITE,
            font=fonts["ip"]
        )

    if ip:
        draw.text(
            (bx + 9, by + 60),
            ip,
            fill=TEXT_IP,
            font=fonts["ip"]
        )

    if notes:
        nl = wrap_text(notes, 27)

        draw.text(
            (bx + 9, by + 84),
            nl[0],
            fill=WHITE,
            font=fonts["small"]
        )

    multi = port.get("_multi", [])

    if multi and len(multi) > 1:
        draw.text(
            (bx + bubble_w - 26, by + 10),
            f"+{len(multi)-1}",
            fill=TEXT_MULTI_BADGE,
            font=fonts["small"]
        )


def draw_panel(draw, fonts, panel_info):
    """
    Dessin complet d'un panneau.
    """

    panel = panel_info["panel"]
    y = panel_info["y"]

    slots = panel["slot_data"]
    cols = panel["cols"]
    label = panel["label"]

    total_slots = len(slots)

    panel_w = cols * (PORT_W + PORT_GAP_X) - PORT_GAP_X + 24

    rows = math.ceil(total_slots / cols)

    panel_h = (
        rows * (PORT_H + PORT_GAP_Y)
        - PORT_GAP_Y
        + 24
    )

    panel_x = MARGIN_LEFT
    panel_y = y

    draw.text(
        (panel_x, panel_y - 22),
        label,
        fill=WHITE,
        font=fonts["switch"]
    )

    rr(
        draw,
        [
            panel_x - 12,
            panel_y,
            panel_x + panel_w,
            panel_y + panel_h
        ],
        radius=8,
        fill=PANEL_BG,
        outline=PANEL_BORDER,
        width=2
    )

    for col in range(cols):
        cx = panel_x + col * (PORT_W + PORT_GAP_X) + PORT_W // 2
        lbl = str(col + 1)

        lw = draw.textlength(lbl, font=fonts["small"])

        draw.text(
            (cx - lw / 2, panel_y - 42),
            lbl,
            fill=WHITE,
            font=fonts["small"]
        )

    for idx, port in enumerate(slots):
        row = idx // cols
        col = idx % cols

        px = panel_x + col * (PORT_W + PORT_GAP_X)
        py = panel_y + row * (PORT_H + PORT_GAP_Y) + 12

        draw_port(draw, fonts, port, px, py)

    bubble_top = panel_y + panel_h + 35

    ref_panel_w = 24 * (PORT_W + PORT_GAP_X) - PORT_GAP_X + 24

    bubble_w = (
        ref_panel_w
        - (BUBBLES_PER_ROW - 1) * BUBBLE_GAP
    ) // BUBBLES_PER_ROW

    annotated = [
        p for p in slots
        if has_display_data(p)
    ]

    for idx, port in enumerate(annotated):
        bx = (
            panel_x - 12
            + (idx % BUBBLES_PER_ROW) * (bubble_w + BUBBLE_GAP)
        )

        by = (
            bubble_top
            + (idx // BUBBLES_PER_ROW) * (BUBBLE_H + BUBBLE_GAP + 24)
        )

        draw_bubble(draw, fonts, port, bx, by, bubble_w)
		# ============================================================
# PDF GENERATION
# ============================================================

def draw_header(draw, fonts, title):
    """
    En-tête du document.
    """

    draw.text(
        (MARGIN_LEFT, 28),
        title,
        fill=WHITE,
        font=fonts["title"]
    )

    draw.text(
        (MARGIN_LEFT, 64),
        "NetDTL - Tableau de brassage",
        fill=WHITE,
        font=fonts["label"]
    )

    draw.line(
        [
            (MARGIN_LEFT, 90),
            (IMG_W - MARGIN_RIGHT, 90)
        ],
        fill=SEPARATOR_COLOR,
        width=1
    )


def draw_legend(draw, fonts):
    """
    Légende graphique.
    """

    legend_y = IMG_H - MARGIN_BOTTOM + 10

    items = [
        (PORT_ACTIVE_RJ45, PORT_ACTIVE_RJ45_B, "RJ45 identifié"),
        (PORT_ACTIVE_RJ11, PORT_ACTIVE_RJ11_B, "RJ11"),
        (PORT_INACTIVE, PORT_INACTIVE_B, "Affecté / non identifié"),
        (PORT_EMPTY, PORT_EMPTY_BORDER, "Libre"),
    ]

    lx = MARGIN_LEFT

    for fill, border, label in items:
        rr(
            draw,
            [lx, legend_y, lx + 18, legend_y + 12],
            radius=2,
            fill=fill,
            outline=border,
            width=1
        )

        draw.text(
            (lx + 24, legend_y),
            label,
            fill=WHITE,
            font=fonts["small"]
        )

        lx += 210


def draw_footer(draw, fonts):
    """
    Pied de page.
    """

    draw.text(
        (IMG_W - 260, IMG_H - MARGIN_BOTTOM + 10),
        "Generated by NetDTL",
        fill=WHITE,
        font=fonts["small"]
    )


def render_image(layout, title):
    """
    Génère l'image Pillow complète.
    """

    fonts = load_fonts()

    img = Image.new(
        "RGB",
        (IMG_W, IMG_H),
        BG_COLOR
    )

    draw = ImageDraw.Draw(img)

    draw_header(draw, fonts, title)

    for panel_info in layout:
        draw_panel(draw, fonts, panel_info)

    draw_legend(draw, fonts)
    draw_footer(draw, fonts)

    return img


def save_pdf_from_image(img, output_path):
    """
    Encapsulation image -> PDF A3
    """

    temp_png = os.path.join(
        os.environ.get("TEMP", "."),
        "patch_panel_tmp.png"
    )

    img.save(
        temp_png,
        dpi=(DPI, DPI)
    )

    pdf = rl_canvas.Canvas(
        output_path,
        pagesize=A3
    )

    pdf.drawImage(
        temp_png,
        0,
        0,
        width=A3_W_PT,
        height=A3_H_PT
    )

    pdf.save()

    if os.path.exists(temp_png):
        os.remove(temp_png)


# ============================================================
# PUBLIC API
# ============================================================

def generate_pdf(
    excel_path,
    sheet_name,
    mapping,
    output_path,
    title="Tableau de brassage"
):
    """
    API publique principale.
    """

    inventory = load_inventory(
        excel_path=excel_path,
        sheet_name=sheet_name,
        mapping=mapping
    )

    layout = normalize_inventory_for_render(
        inventory
    )

    img = render_image(
        layout=layout,
        title=title
    )

    save_pdf_from_image(
        img=img,
        output_path=output_path
    )


def list_excel_sheets(excel_path):
    """
    Retourne les feuilles Excel disponibles.
    """

    if not os.path.exists(excel_path):
        raise PatchPanelError(
            "Fichier Excel introuvable."
        )

    xl = pd.ExcelFile(excel_path)

    return xl.sheet_names


def get_excel_columns(excel_path, sheet_name):
    """
    Retourne les colonnes d'une feuille.
    """

    df = pd.read_excel(
        excel_path,
        sheet_name=sheet_name
    )

    if df.empty:
        raise PatchPanelError(
            "Feuille vide."
        )

    return list(df.columns)
	generate_pdf(
    excel_path="inventaire.xlsx",
    sheet_name="Machines",
    mapping={
        "prise": "PortBrassage",
        "hostname": "NomMachine",
        "ip": "AdresseIP",
        "type": "TypePrise",
        "local": "Bureau",
        "notes": "Commentaire"
    },
    output_path="brassage.pdf",
    title="Baie informatique"
)
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from patch_panel_engine import (
    generate_pdf,
    list_excel_sheets,
    get_excel_columns,
    PatchPanelError
)

APP_TITLE = "NetDTL - Générateur tableau de brassage"
CONFIG_FILE = "patch_panel_config.json"


# ============================================================
# HELPERS
# ============================================================

def resource_path(filename):
    return os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        filename
    )


def load_config():
    cfg = resource_path(CONFIG_FILE)

    if not os.path.exists(cfg):
        return {}

    try:
        with open(cfg, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(data):
    cfg = resource_path(CONFIG_FILE)

    with open(cfg, "w", encoding="utf-8") as f:
        json.dump(
            data,
            f,
            indent=2,
            ensure_ascii=False
        )


def normalize(text):
    return str(text).strip().lower().replace(" ", "")


def guess_column(columns, candidates):
    normalized = {
        normalize(col): col
        for col in columns
    }

    for candidate in candidates:
        c = normalize(candidate)

        if c in normalized:
            return normalized[c]

    return ""


# ============================================================
# COLUMN DETECTION
# ============================================================

COLUMN_HINTS = {
    "prise": [
        "prise",
        "port",
        "portbrassage",
        "numero prise",
        "port brassage",
    ],
    "hostname": [
        "hostname",
        "host",
        "machine",
        "ordinateur",
        "nommachine",
        "nom poste",
        "computer",
    ],
    "ip": [
        "ip",
        "adresseip",
        "ipaddress",
        "adresse ip",
    ],
    "type": [
        "type",
        "typeprise",
        "rj45",
    ],
    "local": [
        "bureau",
        "local",
        "emplacement",
        "site",
        "piece",
    ],
    "notes": [
        "notes",
        "commentaire",
        "description",
        "remarque",
    ]
}


# ============================================================
# GUI
# ============================================================

class PatchPanelApp:

    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("820x520")
        self.root.resizable(False, False)

        self.config = load_config()

        self.excel_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.title_var = tk.StringVar(value="Tableau de brassage")
        self.save_cfg_var = tk.BooleanVar(value=True)

        self.mapping_vars = {
            key: tk.StringVar()
            for key in COLUMN_HINTS.keys()
        }

        self.sheet_combo = None
        self.mapping_combos = {}

        self.build_ui()

    def build_ui(self):
        frame = tk.Frame(self.root, padx=20, pady=20)
        frame.pack(fill="both", expand=True)

        # Excel
        tk.Label(frame, text="Inventaire Excel").grid(
            row=0, column=0, sticky="w"
        )

        tk.Entry(
            frame,
            textvariable=self.excel_var,
            width=70
        ).grid(row=1, column=0)

        tk.Button(
            frame,
            text="Parcourir",
            command=self.select_excel
        ).grid(row=1, column=1, padx=10)

        # Sheet
        tk.Label(frame, text="Feuille").grid(
            row=2, column=0, sticky="w", pady=(20, 0)
        )

        self.sheet_combo = ttk.Combobox(
            frame,
            textvariable=self.sheet_var,
            state="readonly",
            width=67
        )

        self.sheet_combo.grid(row=3, column=0)
        self.sheet_combo.bind(
            "<<ComboboxSelected>>",
            lambda e: self.load_columns()
        )

        # Mapping
        row = 4

        labels = [
            ("prise", "Numéro de prise"),
            ("hostname", "Hostname"),
            ("ip", "Adresse IP"),
            ("type", "Type prise"),
            ("local", "Localisation"),
            ("notes", "Notes"),
        ]

        for key, label in labels:
            tk.Label(frame, text=label).grid(
                row=row,
                column=0,
                sticky="w",
                pady=(12, 0)
            )

            combo = ttk.Combobox(
                frame,
                textvariable=self.mapping_vars[key],
                state="readonly",
                width=67
            )

            combo.grid(row=row + 1, column=0)

            self.mapping_combos[key] = combo

            row += 2

        # Output
        tk.Label(frame, text="PDF de sortie").grid(
            row=row,
            column=0,
            sticky="w",
            pady=(15, 0)
        )

        tk.Entry(
            frame,
            textvariable=self.output_var,
            width=70
        ).grid(row=row + 1, column=0)

        tk.Button(
            frame,
            text="Choisir",
            command=self.select_output
        ).grid(row=row + 1, column=1)

        # Options
        tk.Checkbutton(
            frame,
            text="Mémoriser cette configuration",
            variable=self.save_cfg_var
        ).grid(
            row=row + 2,
            column=0,
            sticky="w",
            pady=12
        )

        tk.Button(
            frame,
            text="Générer PDF",
            height=2,
            command=self.generate
        ).grid(
            row=row + 3,
            column=0,
            pady=20
        )

    def select_excel(self):
        path = filedialog.askopenfilename(
            title="Choisir inventaire Excel",
            filetypes=[
                ("Excel", "*.xlsx *.xls")
            ]
        )

        if not path:
            return

        self.excel_var.set(path)

        try:
            sheets = list_excel_sheets(path)

            self.sheet_combo["values"] = sheets

            preferred = self.config.get("sheet", "")

            if preferred in sheets:
                self.sheet_var.set(preferred)
            else:
                self.sheet_var.set(sheets[0])

            self.load_columns()

        except Exception as e:
            messagebox.showerror(APP_TITLE, str(e))

    def load_columns(self):
        excel = self.excel_var.get()
        sheet = self.sheet_var.get()

        if not excel or not sheet:
            return

        cols = get_excel_columns(excel, sheet)

        for key, combo in self.mapping_combos.items():
            combo["values"] = [""] + cols

            saved = self.config.get("mapping", {}).get(key)

            if saved in cols:
                self.mapping_vars[key].set(saved)
            else:
                guess = guess_column(
                    cols,
                    COLUMN_HINTS[key]
                )
                self.mapping_vars[key].set(guess)

    def select_output(self):
        path = filedialog.asksaveasfilename(
            title="PDF de sortie",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")]
        )

        if path:
            self.output_var.set(path)

    def generate(self):
        try:
            mapping = {
                key: var.get()
                for key, var in self.mapping_vars.items()
                if var.get()
            }

            if self.save_cfg_var.get():
                save_config({
                    "sheet": self.sheet_var.get(),
                    "mapping": mapping
                })

            generate_pdf(
                excel_path=self.excel_var.get(),
                sheet_name=self.sheet_var.get(),
                mapping=mapping,
                output_path=self.output_var.get(),
                title=self.title_var.get()
            )

            os.startfile(self.output_var.get())

            messagebox.showinfo(
                APP_TITLE,
                "PDF généré avec succès."
            )

        except PatchPanelError as e:
            messagebox.showerror(APP_TITLE, str(e))

        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Erreur : {e}")


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    root = tk.Tk()
    app = PatchPanelApp(root)
    root.mainloop()