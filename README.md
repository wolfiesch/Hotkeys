# Excel Databook Layers - AutoHotkey Hotkeys

A comprehensive AutoHotkey script that provides layered hotkey functionality for Excel and PowerPoint, designed to streamline common data manipulation and formatting tasks.

## 🚀 Features

### Excel Automation
- **CapsLock Layer System**: Hold CapsLock + key combinations for quick Excel operations
- **PASTE Operations**: Quick access to paste values, formulas, formats, and special operations
- **FORMAT Operations**: Streamlined formatting for rows, numbers, alignment, and borders
- **NAVIGATION**: Efficient navigation through data blocks and sheets
- **DATA/CLEANUP**: Text processing, filtering, and data cleaning tools
- **Ribbon Command Integration**: Uses Excel's ribbon commands for reliable automation

### PowerPoint Automation
- **Format Sequence**: Quick formatting for PowerPoint objects
- **Object Pane Control**: Streamlined object formatting workflows

## 📋 Quick Start

1. **Install AutoHotkey**: Download and install [AutoHotkey v1.1](https://www.autohotkey.com/download/)
2. **Download Script**: Clone this repository or download `ExcelDatabookLayers.ahk`
3. **Run Script**: Double-click the `.ahk` file to start the hotkey system
4. **Use in Excel**: Hold CapsLock + any key combination while in Excel

## 🎯 Key Hotkeys

### PASTE Layer (CapsLock + key)
- `CapsLock + V` - Paste Values
- `CapsLock + F` - Paste Formulas  
- `CapsLock + T` - Paste Formats
- `CapsLock + W` - Paste Column Widths
- `CapsLock + S` - Paste Values Skip Blanks

### FORMAT Layer (CapsLock + key)
- `CapsLock + 1-4` - Row classification (Section, Subtotal, Major Total, Grand Total)
- `CapsLock + R` - Right Border
- `CapsLock + B` - Bottom Border
- `CapsLock + F1-F3` - Alignment (Left, Center, Right)
- `CapsLock + 5` - Percent Format
- `CapsLock + M` - Month Format
- `CapsLock + D` - Date Format

### NAVIGATION Layer (CapsLock + key)
- `CapsLock + [` - Previous Divider Column
- `CapsLock + ]` - Next Divider Column
- `CapsLock + ,` - Previous Sheet
- `CapsLock + .` - Next Sheet
- `CapsLock + G` - Go To Dialog

### DATA/CLEANUP Layer (CapsLock + key)
- `CapsLock + U` - Trim In Place
- `CapsLock + N` - Convert to Number
- `CapsLock + E` - Text to Columns
- `CapsLock + F7` - Toggle AutoFilter
- `CapsLock + Z` - Clear Formats

## 📚 Documentation

- **[Complete Hotkey Reference](docs/HOTKEYS.md)** - Detailed documentation of all hotkeys
- **[Mapping Documentation](mapping.md)** - Technical implementation details
- **[Archive](archive/)** - Deprecated files and migration notes

## 🛠️ Requirements

- **AutoHotkey v1.1** or later
- **Microsoft Excel** (tested with Excel 2016+)
- **Microsoft PowerPoint** (for PowerPoint hotkeys)

## 📁 Project Structure

```
├── ExcelDatabookLayers.ahk    # Main AutoHotkey script
├── docs/
│   └── HOTKEYS.md             # Complete hotkey documentation
├── archive/                    # Deprecated files
├── create_styled_hotkeys.py   # Python utility for documentation
├── create_styled_pdf.py       # PDF generation utility
├── ExcelDatabookLayers_Hotkeys.xlsx  # Excel reference
├── ExcelDatabookLayers_Hotkeys.pdf   # PDF reference
└── mapping.md                 # Technical documentation
```

## 🔧 Customization

The script is designed to be easily customizable:

1. **Add New Hotkeys**: Edit the main script file and add new hotkey definitions
2. **Modify Existing Actions**: Update the action blocks for existing hotkeys
3. **Add New Applications**: Extend the `#IfWinActive` blocks for other applications

## 🤝 Contributing

Contributions are welcome! Please see [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🐛 Troubleshooting

### Common Issues

1. **Hotkeys not working**: Ensure AutoHotkey is running and the script is loaded
2. **Excel not responding**: Check if Excel is in the correct mode (not in edit mode)
3. **Conflicts with other software**: Some hotkeys may conflict with other applications

### Getting Help

- Check the [documentation](docs/HOTKEYS.md) for detailed hotkey information
- Review the [mapping documentation](mapping.md) for technical details
- Open an issue on GitHub for bugs or feature requests

## 📈 Version History

### v1.0.0 (2025-01-09)
- Initial release with comprehensive Excel automation
- CapsLock layer system implementation
- PowerPoint hotkey support
- Complete documentation and reference materials

## 🙏 Acknowledgments

- Built with [AutoHotkey](https://www.autohotkey.com/)
- Designed for Excel power users and data analysts
- Inspired by the need for efficient data manipulation workflows

---

**Note**: This script is designed to work with Excel's ribbon interface and may require adjustments for different Excel versions or configurations.
