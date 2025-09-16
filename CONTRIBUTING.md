# Contributing to Excel Databook Layers

Thank you for your interest in contributing to the Excel Databook Layers project! This document provides guidelines for contributing to this AutoHotkey-based Excel automation tool.

## 🤝 How to Contribute

### Reporting Issues

Before creating an issue, please:
1. Check if the issue already exists
2. Test with the latest version of the script
3. Provide detailed information about your environment

When reporting issues, include:
- **AutoHotkey version**: Run `AutoHotkey.exe /Version` in command prompt
- **Excel version**: Help > About Microsoft Excel
- **Operating System**: Windows version
- **Steps to reproduce**: Detailed steps to recreate the issue
- **Expected behavior**: What should happen
- **Actual behavior**: What actually happens
- **Screenshots**: If applicable

### Suggesting Features

We welcome feature suggestions! Please:
1. Check existing issues and discussions first
2. Describe the use case and benefit
3. Provide examples of how the feature would work
4. Consider backward compatibility

### Code Contributions

#### Getting Started

1. **Fork the repository** on GitHub
2. **Clone your fork** locally:
   ```bash
   git clone https://github.com/YOUR_USERNAME/ExcelDatabookLayers.git
   cd ExcelDatabookLayers
   ```
3. **Create a feature branch**:
   ```bash
   git checkout -b feature/your-feature-name
   ```

#### Development Guidelines

**Code Style:**
- Use consistent indentation (tabs or spaces, but be consistent)
- Add comments for complex logic
- Follow AutoHotkey naming conventions
- Keep functions focused and modular

**Testing:**
- Test all new hotkeys thoroughly in Excel
- Verify compatibility with different Excel versions
- Test edge cases and error conditions
- Ensure no conflicts with existing hotkeys

**Documentation:**
- Update `docs/HOTKEYS.md` for new hotkeys
- Update `mapping.md` for technical changes
- Add comments in the code
- Update README.md if needed

#### Pull Request Process

1. **Test your changes** thoroughly
2. **Update documentation** as needed
3. **Commit your changes** with clear commit messages
4. **Push to your fork**:
   ```bash
   git push origin feature/your-feature-name
   ```
5. **Create a Pull Request** on GitHub

#### Commit Message Format

Use clear, descriptive commit messages:
```
Add: New hotkey for paste special operations
Fix: CapsLock layer not working in Excel 2019
Update: Documentation for new formatting hotkeys
```

## 📋 Development Setup

### Prerequisites

- **AutoHotkey v1.1+**: Download from [autohotkey.com](https://www.autohotkey.com/download/)
- **Microsoft Excel**: Any version 2016 or later
- **Git**: For version control
- **Text Editor**: VS Code, Notepad++, or any editor with AutoHotkey syntax support

### Testing Environment

1. **Create a test Excel file** with sample data
2. **Test each hotkey** systematically
3. **Verify ribbon commands** work as expected
4. **Check for conflicts** with other applications

### Code Organization

The main script (`ExcelDatabookLayers.ahk`) is organized into sections:
- **PowerPoint hotkeys** (lines ~1-50)
- **Excel CapsLock layer** (lines ~51-200)
- **Helper functions** (lines ~201+)

When adding new hotkeys:
- Place them in the appropriate section
- Follow the existing pattern
- Add comments explaining the functionality

## 🎯 Areas for Contribution

### High Priority
- **Bug fixes** for existing hotkeys
- **Performance improvements**
- **Better error handling**
- **Additional Excel versions support**

### Medium Priority
- **New hotkey combinations**
- **Additional applications** (Word, Outlook, etc.)
- **Configuration options**
- **User interface improvements**

### Low Priority
- **Documentation improvements**
- **Code refactoring**
- **Additional utilities**
- **Examples and tutorials**

## 🚫 What Not to Contribute

- **Breaking changes** without discussion
- **Hotkeys that conflict** with existing ones
- **Platform-specific code** without Windows fallback
- **Code without proper testing**

## 📞 Getting Help

If you need help contributing:
- **Open a discussion** on GitHub
- **Check existing issues** for similar problems
- **Review the documentation** in the `docs/` folder
- **Look at existing code** for examples

## 🏆 Recognition

Contributors will be:
- Listed in the README.md
- Mentioned in release notes
- Credited in the documentation

## 📄 License

By contributing, you agree that your contributions will be licensed under the same MIT License that covers the project.

---

Thank you for contributing to Excel Databook Layers! Your contributions help make Excel automation more accessible and efficient for everyone.
