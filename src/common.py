"""Shared paths and the report colour palette (same palette as the Excel workbook)."""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
RAW = DATA / "raw"
PUBLIC = DATA / "public"
SIM = DATA / "simulated"
OUT = DATA / "output"
DOCS = ROOT / "docs"
IMG = DOCS / "img"

# Palette (taken from the light case-study theme)
BLUE = "#2A78D6"        # main data colour
BLUE_DARK = "#1C5CAB"
BLUE_LIGHT = "#86B6EF"
BLUE_PALE = "#CDE2FB"
GREY = "#898781"        # comparison
GREY_LIGHT = "#C9C7C0"
RED = "#D03B3B"         # alerts, margin below floor
GREEN = "#0CA30C"
ORANGE = "#EB6834"
INK = "#0B0B0B"
INK_SOFT = "#52514E"
PAGE_BG = "#F4F4F2"
CARD_BG = "#FCFCFB"
BORDER = "#E1E0D9"
FONT = "Segoe UI"
