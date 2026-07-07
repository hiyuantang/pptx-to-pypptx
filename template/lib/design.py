from pptx.util import Inches

# Slide canvas size (16:9) in EMUs to avoid float truncation.
# Inches(13.333) truncates to 12191695, which triggers PowerPoint repair.
SLIDE_W = 12192000
SLIDE_H = 6858000

# Chrome layout constants
TITLE_BAR = (Inches(0.382), Inches(0.237), Inches(12.569), Inches(0.839))
SEP_Y = Inches(0.864)
FOOTER = {"x": Inches(0.222), "y": Inches(7.0), "w": Inches(1.5), "h": Inches(0.286)}
SLIDE_NUM = {"x": Inches(10.0), "y": Inches(7.0), "w": Inches(3.111), "h": Inches(0.399)}
CONTENT_TOP = Inches(1.05)
CONTENT_BOTTOM = Inches(6.7)
CONTENT_LEFT = Inches(0.382)
CONTENT_RIGHT = Inches(12.95)

# Color palette — edit to match your deck
COL = {
    "title": "1F497D",
    "body": "464646",
    "bg": "F9F9F9",
    "white": "FFFFFF",
    "black": "000000",
    "sep": "BFBFBF",
    "footer": "808080",
    "blue": "4F81BD",
    "green": "9BBB59",
    "red": "C0504D",
    "dark_red": "C00000",
    "yellow": "FFF2CC",
    "purple": "8064A2",
    "orange": "F79646",
    "table_alt": "F2F2F2",
}

# Font families — edit to match your deck
FONT = {
    "medium": "Arial",
    "regular": "Arial",
    "bold": "Arial Bold",
}

# Footer text — auto-detected at scaffold time; edit to change it on every slide.
FOOTER_TEXT = __FOOTER_TEXT__
