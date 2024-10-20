# File Renamer Tool

![Python](https://img.shields.io/badge/Python-3.x-blue.svg) ![Tkinter](https://img.shields.io/badge/Tkinter-GUI-green.svg) ![PyInstaller](https://img.shields.io/badge/PyInstaller-Executable-yellow.svg)

A Python application with a graphical user interface (GUI) that renames files based on data provided in an Excel file. The user selects the Excel file and a directory, and the application matches file names, then renames them in the format `prd_code Prd_name Number`.

## Features

- Simple GUI built using Tkinter for ease of use.
- Supports Excel files (`.xlsx`) for renaming.
- Progress bar for real-time feedback on renaming progress.
- Preview functionality for listing files before renaming.
- Filter for specific file types (images, documents, etc.).
- Cross-platform support (Windows, macOS, Linux).
- Easily packaged as an executable for sharing (no need for users to install Python).

## Table of Contents

- [Installation](#installation)
- [Usage](#usage)
- [How it Works](#how-it-works)
- [Screenshots](#screenshots)
- [Development](#development)
- [Contributing](#contributing)
- [License](#license)

## Installation

### Option 1: Run via Python (For Developers)
To run the application directly from Python, follow these steps:

1. **Clone the repository**:
   ```bash
   git clone https://github.com/yourusername/file-renamer-tool.git
   cd file-renamer-tool
