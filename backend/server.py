
# FULL BACKEND PROVIDED BY USER (unchanged except Render-safe run)
import os
import tempfile
import logging
import html
import re
from pathlib import Path
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.oxml import parse_xml

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

THEME_COLOR_MAP = {
    "ACCENT1": "#4472C4", "ACCENT_1": "#4472C4",
    "ACCENT2": "#ED7D31", "ACCENT_2": "#ED7D31",
    "ACCENT3": "#A5A5A5", "ACCENT_3": "#A5A5A5",
    "ACCENT4": "#FFC000", "ACCENT_4": "#FFC000",
    "ACCENT5": "#5B9BD5", "ACCENT_5": "#5B9BD5",
    "ACCENT6": "#70AD47", "ACCENT_6": "#70AD47",
    "TEXT1": "#000000", "TEXT_1": "#000000",
    "TEXT2": "#FFFFFF", "TEXT_2": "#FFFFFF",
    "BACKGROUND1": "#FFFFFF", "BACKGROUND_1": "#FFFFFF",
    "BACKGROUND2": "#000000", "BACKGROUND_2": "#000000",
}

GENERATED_CSS = set()

def sanitize_name(s, maxlen=40):
    if not s: return ""
    s = re.sub(r'[^a-zA-Z0-9]+', '-', s).strip('-').lower()
    return s[:maxlen]

def rgb_to_hex(rgb):
    try:
        return f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
    except:
        return None

def extract_theme_colors(prs):
    theme_colors = {}
    try:
        theme_part = prs.part.theme_part
        root = parse_xml(theme_part.blob)
        clr_scheme = root.find(".//a:clrScheme", root.nsmap)
        if clr_scheme is not None:
            for elem in clr_scheme:
                name = elem.tag.split("}")[-1].upper()
                srgb = elem.find(".//a:srgbClr", root.nsmap)
                if srgb is not None:
                    val = srgb.get("val")
                    if val:
                        theme_colors[name] = f"#{val}"
                        theme_colors[name.replace("_","")] = f"#{val}"
    except:
        pass
    for k,v in THEME_COLOR_MAP.items():
        theme_colors.setdefault(k, v)
    return theme_colors

def add_css(selector, body):
    GENERATED_CSS.add(f"{selector} {{ {body} }}")

def process_table(table):
    rows = []
    for row in table.rows:
        cells = []
        for cell in row.cells:
            cells.append(f"<td>{html.escape(cell.text)}</td>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    add_css(".ppt-table", "border-collapse:collapse;width:100%")
    return "<table class='ppt-table'>" + "".join(rows) + "</table>"

app = Flask(__name__)
CORS(app)

@app.route("/")
def home():
    return "PPTX to HTML API running"

@app.route("/convert", methods=["POST"])
def convert():
    if "pptFile" not in request.files:
        return jsonify({"error":"No file"}),400
    f = request.files["pptFile"]
    tmp = tempfile.NamedTemporaryFile(delete=False,suffix=".pptx")
    f.save(tmp.name)

    prs = Presentation(tmp.name)
    blocks = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_table:
                blocks.append(process_table(shape.table))

    css = "<style>" + "\\n".join(GENERATED_CSS) + "</style>"
    os.remove(tmp.name)
    return jsonify({"html": css + "<br/>".join(blocks)})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
