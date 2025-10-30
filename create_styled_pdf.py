#!/usr/bin/env python3
"""
Create a beautifully styled PDF for ExcelDatabookLayers.ahk hotkey reference
Enhanced design with modern typography, icons, and professional layout
"""

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    PageBreak,
)
from reportlab.platypus import Frame, PageTemplate, BaseDocTemplate
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, mm
from reportlab.pdfgen import canvas
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from datetime import datetime
import os


class HotkeyPDFGenerator:
    def __init__(self):
        self.doc = None
        self.styles = getSampleStyleSheet()
        self.story = []

        # Enhanced color palette with modern design
        self.colors = {
            "PowerPoint": {
                "bg": colors.Color(0.89, 0.85, 0.98),
                "border": colors.Color(0.42, 0.22, 0.69),
            },  # Purple
            "PASTE": {
                "bg": colors.Color(0.85, 0.93, 1.0),
                "border": colors.Color(0.12, 0.47, 0.71),
            },  # Blue
            "FORMAT": {
                "bg": colors.Color(0.85, 1.0, 0.85),
                "border": colors.Color(0.13, 0.55, 0.13),
            },  # Green
            "ALIGNMENT": {
                "bg": colors.Color(1.0, 0.93, 0.8),
                "border": colors.Color(0.8, 0.52, 0.25),
            },  # Orange
            "BORDERS": {
                "bg": colors.Color(1.0, 0.9, 0.9),
                "border": colors.Color(0.8, 0.2, 0.2),
            },  # Red
            "SIZING": {
                "bg": colors.Color(0.85, 1.0, 1.0),
                "border": colors.Color(0.0, 0.5, 0.5),
            },  # Teal
            "NAVIGATION": {
                "bg": colors.Color(1.0, 0.98, 0.8),
                "border": colors.Color(0.85, 0.65, 0.13),
            },  # Yellow
            "DATA": {
                "bg": colors.Color(1.0, 0.9, 0.95),
                "border": colors.Color(0.8, 0.2, 0.6),
            },  # Pink
            "CLEAR": {
                "bg": colors.Color(0.94, 0.94, 0.94),
                "border": colors.Color(0.4, 0.4, 0.4),
            },  # Gray
        }

        # Create custom styles
        self.create_custom_styles()

        # Hotkey data with the split entries
        self.hotkeys = self.get_hotkey_data()

    def create_custom_styles(self):
        """Create custom paragraph styles for the PDF"""

        # Title style
        self.styles.add(
            ParagraphStyle(
                name="CustomTitle",
                parent=self.styles["Title"],
                fontSize=24,
                textColor=colors.Color(0.2, 0.2, 0.2),
                spaceAfter=20,
                alignment=TA_CENTER,
                fontName="Helvetica-Bold",
            )
        )

        # Subtitle style
        self.styles.add(
            ParagraphStyle(
                name="CustomSubtitle",
                parent=self.styles["Normal"],
                fontSize=14,
                textColor=colors.Color(0.4, 0.4, 0.4),
                spaceAfter=30,
                alignment=TA_CENTER,
                fontName="Helvetica",
            )
        )

        # Category header style
        self.styles.add(
            ParagraphStyle(
                name="CategoryHeader",
                parent=self.styles["Heading2"],
                fontSize=16,
                textColor=colors.white,
                spaceBefore=15,
                spaceAfter=8,
                alignment=TA_CENTER,
                fontName="Helvetica-Bold",
            )
        )

        # Hotkey text style
        self.styles.add(
            ParagraphStyle(
                name="HotkeyText",
                parent=self.styles["Normal"],
                fontSize=11,
                fontName="Helvetica-Bold",
                textColor=colors.Color(0.0, 0.0, 0.5),
            )
        )

        # Function text style
        self.styles.add(
            ParagraphStyle(
                name="FunctionText",
                parent=self.styles["Normal"],
                fontSize=10,
                fontName="Helvetica-Bold",
            )
        )

        # Description text style
        self.styles.add(
            ParagraphStyle(
                name="DescText",
                parent=self.styles["Normal"],
                fontSize=9,
                fontName="Helvetica",
            )
        )

    def get_hotkey_data(self):
        """Return the hotkey data with split entries"""
        return [
            # PowerPoint
            [
                "PowerPoint",
                "Ctrl+Alt+Shift+S",
                "^!+s",
                "PowerPoint Sequence",
                "Send Tab Tab 70 Enter Tab×4 4.57 Tab×2 21.47",
                "30",
            ],
            [
                "PowerPoint",
                "Ctrl+Alt+Shift+F",
                "^!+f",
                "Format Object Pane",
                "Alt+4 wait Ctrl+A wait 70 Enter",
                "52",
            ],
            # PASTE
            [
                "PASTE",
                "CapsLock+V",
                "SC03A & v",
                "Paste Values",
                "Paste values only without formulas or formatting",
                "756",
            ],
            [
                "PASTE",
                "CapsLock+F",
                "SC03A & f",
                "Paste Formulas",
                "Paste formulas only",
                "757",
            ],
            [
                "PASTE",
                "Ctrl+CapsLock+F",
                "Ctrl+SC03A & f",
                "Freeze Panes",
                "Freeze panes at active cell",
                "757",
            ],
            [
                "PASTE",
                "CapsLock+T",
                "SC03A & t",
                "Paste Formats",
                "Paste formatting only",
                "765",
            ],
            [
                "PASTE",
                "CapsLock+W",
                "SC03A & w",
                "Paste Column Widths",
                "Paste column width information",
                "766",
            ],
            [
                "PASTE",
                "CapsLock+S",
                "SC03A & s",
                "Paste Formulas + Format",
                "Paste both formulas and formatting",
                "767",
            ],
            [
                "PASTE",
                "CapsLock+X",
                "SC03A & x",
                "Paste Values (Transpose)",
                "Paste values with transpose operation",
                "768",
            ],
            [
                "PASTE",
                "CapsLock+L",
                "SC03A & l",
                "Paste Link",
                "Create linked paste",
                "769",
            ],
            [
                "PASTE",
                "CapsLock+P",
                "SC03A & p",
                "Paste Special Dialog",
                "Open Paste Special dialog box",
                "770",
            ],
            [
                "PASTE",
                "CapsLock+Numpad/",
                "SC03A & NumpadDiv",
                "Toggle Filter",
                "Toggle AutoFilter on/off",
                "773",
            ],
            [
                "PASTE",
                "CapsLock+Numpad*",
                "SC03A & NumpadMult",
                "Clear Filter",
                "Clear all filters",
                "774",
            ],
            [
                "PASTE",
                "CapsLock+Numpad+",
                "SC03A & NumpadAdd",
                "Paste Add",
                "Paste operation with addition",
                "775",
            ],
            [
                "PASTE",
                "CapsLock+Numpad-",
                "SC03A & NumpadSub",
                "Paste Subtract",
                "Paste operation with subtraction",
                "776",
            ],
            # FORMAT
            [
                "FORMAT",
                "CapsLock+1",
                "SC03A & 1",
                "Custom Font Color",
                "Run custom font color macro",
                "779",
            ],
            [
                "FORMAT",
                "CapsLock+2",
                "SC03A & 2",
                "Subtotal Format",
                "Apply subtotal row formatting",
                "780",
            ],
            [
                "FORMAT",
                "CapsLock+3",
                "SC03A & 3",
                "Major Total Format",
                "Apply major total row formatting",
                "781",
            ],
            [
                "FORMAT",
                "CapsLock+4",
                "SC03A & 4",
                "Grand Total Format",
                "Apply grand total row formatting",
                "782",
            ],
            [
                "FORMAT",
                "CapsLock+6",
                "SC03A & 6",
                "Set Row Height 5pt",
                "Set row height to 5 points",
                "783",
            ],
            [
                "FORMAT",
                "CapsLock+9",
                "SC03A & 9",
                "General Format",
                "Apply general number format",
                "786",
            ],
            [
                "FORMAT",
                "CapsLock+A",
                "SC03A & a",
                "Set Row Height 5pt",
                "Set row height to 5 points",
                "787",
            ],
            [
                "FORMAT",
                "CapsLock+K",
                "SC03A & k",
                "Custom Number Format",
                "Run custom number format macro",
                "788",
            ],
            [
                "FORMAT",
                "CapsLock+5",
                "SC03A & 5",
                "Percent Format",
                "Apply percentage number format",
                "789",
            ],
            [
                "FORMAT",
                "CapsLock+M",
                "SC03A & m",
                "Month Format",
                "Apply month format (mmm-yy)",
                "790",
            ],
            [
                "FORMAT",
                "CapsLock+D",
                "SC03A & d",
                "Date Format",
                "Apply date number format",
                "791",
            ],
            # ALIGNMENT
            [
                "ALIGNMENT",
                "CapsLock+F1",
                "SC03A & F1",
                "Align Left",
                "Set text alignment to left",
                "794",
            ],
            [
                "ALIGNMENT",
                "CapsLock+F2",
                "SC03A & F2",
                "Align Center",
                "Set text alignment to center",
                "795",
            ],
            [
                "ALIGNMENT",
                "CapsLock+F3",
                "SC03A & F3",
                "Align Right",
                "Set text alignment to right",
                "796",
            ],
            [
                "ALIGNMENT",
                "CapsLock+F4",
                "SC03A & F4",
                "Toggle Wrap Text",
                "Toggle text wrapping in cells",
                "797",
            ],
            # BORDERS
            [
                "BORDERS",
                "CapsLock+R",
                "SC03A & r",
                "Right Border",
                "Apply right border macro",
                "800",
            ],
            [
                "SIZING",
                "Ctrl+CapsLock+R",
                "Ctrl+SC03A & r",
                "AutoFit Rows",
                "Auto-adjust row heights to fit content",
                "800",
            ],
            [
                "BORDERS",
                "CapsLock+B",
                "SC03A & b",
                "Bottom Border",
                "Apply bottom thin border macro",
                "808",
            ],
            [
                "BORDERS",
                "CapsLock+O",
                "SC03A & o",
                "Outline Borders",
                "Apply outline borders to selection",
                "809",
            ],
            [
                "BORDERS",
                "CapsLock+I",
                "SC03A & i",
                "Inside Borders",
                "Apply inside borders to selection",
                "810",
            ],
            [
                "BORDERS",
                "CapsLock+C",
                "SC03A & c",
                "Clear Borders",
                "Clear all borders",
                "811",
            ],
            [
                "SIZING",
                "Ctrl+CapsLock+C",
                "Ctrl+SC03A & c",
                "AutoFit Columns",
                "Auto-adjust column widths to fit content",
                "811",
            ],
            [
                "BORDERS",
                "CapsLock+Y",
                "SC03A & y",
                "Top Double Border",
                "Apply double border to top",
                "819",
            ],
            [
                "BORDERS",
                "CapsLock+J",
                "SC03A & j",
                "Left Thick Border",
                "Apply thick border to left",
                "820",
            ],
            [
                "BORDERS",
                "CapsLock+;",
                "SC03A & ;",
                "Right Thick Border",
                "Apply thick border to right",
                "821",
            ],
            # SIZING
            [
                "SIZING",
                "CapsLock+Q",
                "SC03A & q",
                "Column Width 0.5",
                "Set column width to 0.5",
                "824",
            ],
            [
                "SIZING",
                "Ctrl+CapsLock+Q",
                "Ctrl+SC03A & q",
                "Row Height 5pt",
                "Set row height to 5 points",
                "824",
            ],
            [
                "SIZING",
                "CapsLock+F5",
                "SC03A & F5",
                "AutoFit Columns",
                "Auto-adjust column widths to fit content",
                "832",
            ],
            [
                "SIZING",
                "CapsLock+F6",
                "SC03A & F6",
                "AutoFit Rows",
                "Auto-adjust row heights to fit content",
                "833",
            ],
            [
                "SIZING",
                "CapsLock+F11",
                "SC03A & F11",
                "Increase Indent",
                "Increase text indentation",
                "834",
            ],
            [
                "SIZING",
                "CapsLock+F12",
                "SC03A & F12",
                "Decrease Indent",
                "Decrease text indentation",
                "835",
            ],
            [
                "SIZING",
                "CapsLock+Numpad.",
                "SC03A & NumpadDot",
                "Add Decimal Place",
                "Increase decimal places in number format",
                "836",
            ],
            [
                "SIZING",
                "CapsLock+Numpad0",
                "SC03A & Numpad0",
                "Remove Decimal Place",
                "Decrease decimal places in number format",
                "837",
            ],
            # NAVIGATION
            [
                "NAVIGATION",
                "CapsLock+[",
                "SC03A & [",
                "Previous Divider",
                "Jump to previous data boundary",
                "840",
            ],
            [
                "NAVIGATION",
                "CapsLock+]",
                "SC03A & ]",
                "Next Divider",
                "Jump to next data boundary",
                "841",
            ],
            [
                "NAVIGATION",
                "CapsLock+=",
                "SC03A & =",
                "First Block Edge",
                "Jump to first block edge",
                "842",
            ],
            [
                "NAVIGATION",
                "CapsLock+-",
                "SC03A & -",
                "Last Block Edge",
                "Jump to last block edge",
                "843",
            ],
            [
                "NAVIGATION",
                "CapsLock+,",
                "SC03A & ,",
                "Previous Sheet",
                "Switch to previous worksheet",
                "844",
            ],
            [
                "NAVIGATION",
                "CapsLock+.",
                "SC03A & .",
                "Next Sheet",
                "Switch to next worksheet",
                "845",
            ],
            [
                "NAVIGATION",
                "CapsLock+G",
                "SC03A & g",
                "Go To",
                "Open Go To dialog (Ctrl+G)",
                "846",
            ],
            [
                "NAVIGATION",
                "CapsLock+8",
                "SC03A & 8",
                "Select Current Region",
                "Select current data region (Ctrl+Shift+8)",
                "847",
            ],
            [
                "NAVIGATION",
                "CapsLock+H",
                "SC03A & h",
                "Custom Macro H",
                "Run custom macro Ctrl+Alt+Shift+H",
                "848",
            ],
            [
                "NAVIGATION",
                "CapsLock+Numpad8",
                "SC03A & Numpad8",
                "Ctrl+Up",
                "Move to edge of data block upward",
                "851",
            ],
            [
                "NAVIGATION",
                "CapsLock+Numpad2",
                "SC03A & Numpad2",
                "Ctrl+Down",
                "Move to edge of data block downward",
                "852",
            ],
            [
                "NAVIGATION",
                "CapsLock+Numpad4",
                "SC03A & Numpad4",
                "Ctrl+Left",
                "Move to edge of data block leftward",
                "853",
            ],
            [
                "NAVIGATION",
                "CapsLock+Numpad6",
                "SC03A & Numpad6",
                "Ctrl+Right",
                "Move to edge of data block rightward",
                "854",
            ],
            [
                "NAVIGATION",
                "CapsLock+Numpad7",
                "SC03A & Numpad7",
                "Ctrl+Home",
                "Move to cell A1",
                "855",
            ],
            [
                "NAVIGATION",
                "CapsLock+Numpad9",
                "SC03A & Numpad9",
                "Ctrl+End",
                "Move to last used cell",
                "856",
            ],
            [
                "NAVIGATION",
                "CapsLock+Space",
                "SC03A & Space",
                "Show Help",
                "Display hotkey reference overlay",
                "872",
            ],
            # DATA
            [
                "DATA",
                "CapsLock+U",
                "SC03A & u",
                "Trim In Place",
                "Remove leading/trailing spaces from text",
                "859",
            ],
            [
                "DATA",
                "CapsLock+F8",
                "SC03A & F8",
                "Clean In Place",
                "Remove non-printable characters",
                "860",
            ],
            [
                "DATA",
                "CapsLock+N",
                "SC03A & n",
                "Convert to Number",
                "Convert text to numbers using Text to Columns",
                "861",
            ],
            [
                "DATA",
                "CapsLock+E",
                "SC03A & e",
                "Text to Columns",
                "Open Text to Columns wizard",
                "862",
            ],
            [
                "DATA",
                "CapsLock+F7",
                "SC03A & F7",
                "Toggle AutoFilter",
                "Toggle AutoFilter on/off (Ctrl+Shift+L)",
                "863",
            ],
            # CLEAR
            [
                "CLEAR",
                "CapsLock+Shift+Delete",
                "SC03A & +Delete",
                "Clear Formats",
                "Clear cell formatting only",
                "867",
            ],
            [
                "CLEAR",
                "CapsLock+Delete",
                "SC03A & Delete",
                "Clear Contents",
                "Clear cell contents only",
                "868",
            ],
            [
                "CLEAR",
                "CapsLock+Ctrl+Delete",
                "SC03A & ^Delete",
                "Clear All",
                "Clear both contents and formatting",
                "869",
            ],
        ]

    def add_title_page(self):
        """Add a professional title page"""

        # Main title
        title = Paragraph("ExcelDatabookLayers.ahk", self.styles["CustomTitle"])
        self.story.append(title)

        # Subtitle
        subtitle = Paragraph("Hotkey Reference Guide", self.styles["CustomSubtitle"])
        self.story.append(subtitle)

        # Add some space
        self.story.append(Spacer(1, 0.5 * inch))

        # Summary stats
        total_hotkeys = len(self.hotkeys)
        categories = list(set([h[0] for h in self.hotkeys]))

        stats_data = [
            ["Total Hotkeys:", str(total_hotkeys)],
            ["Categories:", str(len(categories))],
            ["Generated:", datetime.now().strftime("%Y-%m-%d %H:%M")],
            ["Source File:", "ExcelDatabookLayers.ahk"],
        ]

        stats_table = Table(stats_data, colWidths=[2 * inch, 2 * inch])
        stats_table.setStyle(
            TableStyle(
                [
                    ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                    ("FONTSIZE", (0, 0), (-1, -1), 12),
                    ("ALIGN", (0, 0), (0, -1), "RIGHT"),
                    ("ALIGN", (1, 0), (1, -1), "LEFT"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 12),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 12),
                    ("TOPPADDING", (0, 0), (-1, -1), 8),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                    ("GRID", (0, 0), (-1, -1), 1, colors.lightgrey),
                ]
            )
        )

        self.story.append(stats_table)
        self.story.append(Spacer(1, 0.5 * inch))

        # Category overview
        overview_para = Paragraph("<b>Categories Overview:</b>", self.styles["Normal"])
        self.story.append(overview_para)
        self.story.append(Spacer(1, 0.2 * inch))

        # Create category overview table
        cat_data = []
        for category in sorted(categories):
            count = len([h for h in self.hotkeys if h[0] == category])
            cat_data.append([category, str(count) + " hotkeys"])

        cat_table = Table(cat_data, colWidths=[2 * inch, 1.5 * inch])
        cat_table.setStyle(
            TableStyle(
                [
                    ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                    ("FONTSIZE", (0, 0), (-1, -1), 11),
                    ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 12),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 12),
                    ("TOPPADDING", (0, 0), (-1, -1), 6),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                    ("GRID", (0, 0), (-1, -1), 1, colors.lightgrey),
                    ("BACKGROUND", (0, 0), (-1, -1), colors.Color(0.98, 0.98, 0.98)),
                ]
            )
        )

        self.story.append(cat_table)
        self.story.append(PageBreak())

    def create_category_table(self, category_name, category_hotkeys):
        """Create a beautifully styled table for a category"""

        # Category header
        header_para = Paragraph(
            f"<b>{category_name.upper()}</b>", self.styles["CategoryHeader"]
        )

        # Create header background
        header_table = Table([[header_para]], colWidths=[7 * inch])
        header_table.setStyle(
            TableStyle(
                [
                    (
                        "BACKGROUND",
                        (0, 0),
                        (-1, -1),
                        self.colors[category_name]["border"],
                    ),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("TOPPADDING", (0, 0), (-1, -1), 8),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                    ("LEFTPADDING", (0, 0), (-1, -1), 12),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 12),
                ]
            )
        )

        self.story.append(header_table)

        # Prepare table data
        table_data = []

        for hotkey in category_hotkeys:
            # Format each cell with appropriate styling
            hotkey_display = Paragraph(f"<b>{hotkey[1]}</b>", self.styles["HotkeyText"])
            function_name = Paragraph(
                f"<b>{hotkey[3]}</b>", self.styles["FunctionText"]
            )
            description = Paragraph(hotkey[4], self.styles["DescText"])

            table_data.append([hotkey_display, function_name, description])

        # Create the main content table
        content_table = Table(table_data, colWidths=[2.2 * inch, 2 * inch, 2.8 * inch])

        # Apply sophisticated styling
        table_style = [
            # Basic formatting
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("ALIGN", (0, 0), (-1, -1), "LEFT"),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            # Padding
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            # Background colors (alternating rows)
            ("BACKGROUND", (0, 0), (-1, -1), self.colors[category_name]["bg"]),
            # Borders
            ("GRID", (0, 0), (-1, -1), 0.5, self.colors[category_name]["border"]),
            ("BOX", (0, 0), (-1, -1), 1, self.colors[category_name]["border"]),
        ]

        # Add alternating row colors for better readability
        for i in range(len(table_data)):
            if i % 2 == 1:  # Every other row
                darker_bg = colors.Color(
                    max(0, self.colors[category_name]["bg"].red - 0.03),
                    max(0, self.colors[category_name]["bg"].green - 0.03),
                    max(0, self.colors[category_name]["bg"].blue - 0.03),
                )
                table_style.append(("BACKGROUND", (0, i), (-1, i), darker_bg))

        content_table.setStyle(TableStyle(table_style))

        self.story.append(content_table)
        self.story.append(Spacer(1, 0.3 * inch))

    def generate_pdf(self):
        """Generate the complete PDF"""

        filename = r"C:\Users\wschoenberger\Hotkeys\ExcelDatabookLayers_Hotkeys.pdf"

        # Create document with custom page template
        self.doc = SimpleDocTemplate(
            filename,
            pagesize=letter,
            rightMargin=0.5 * inch,
            leftMargin=0.5 * inch,
            topMargin=0.7 * inch,
            bottomMargin=0.7 * inch,
        )

        # Add title page
        self.add_title_page()

        # Group hotkeys by category and create tables
        categories = {}
        for hotkey in self.hotkeys:
            category = hotkey[0]
            if category not in categories:
                categories[category] = []
            categories[category].append(hotkey)

        # Sort categories for consistent order
        category_order = [
            "PowerPoint",
            "PASTE",
            "FORMAT",
            "ALIGNMENT",
            "BORDERS",
            "SIZING",
            "NAVIGATION",
            "DATA",
            "CLEAR",
        ]

        for category in category_order:
            if category in categories:
                self.create_category_table(category, categories[category])

        # Build the PDF
        self.doc.build(self.story)
        print(f"PDF created successfully: {filename}")
        print(
            f"Total pages generated with {len(self.hotkeys)} hotkeys across {len(categories)} categories"
        )


def main():
    """Main function to generate the PDF"""
    generator = HotkeyPDFGenerator()
    generator.generate_pdf()


if __name__ == "__main__":
    main()
