# 🔍 WebShot Checker

<div align="center">

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Playwright](https://img.shields.io/badge/Playwright-Async-2EAD33?style=for-the-badge&logo=playwright&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-yellow?style=for-the-badge)
![Platform](https://img.shields.io/badge/Platform-Linux%20%7C%20Windows%20%7C%20macOS-lightgrey?style=for-the-badge)

**Bulk domain & IP screenshot scanner with professional HTML report and Excel export.**  
Designed to scale from 1 to tens of thousands of targets with high performance.

[Features](#-features) · [Installation](#-installation) · [Usage](#-usage) · [Output](#-output) · [Options](#%EF%B8%8F-options)

</div>

---

## ✨ Features

- 🚀 **Async & concurrent** — scan thousands of targets in parallel using Playwright async engine
- 🌐 **Auto HTTP + HTTPS** — every target is checked on both protocols automatically
- 🎯 **Flexible target format** — supports `domain`, `IP`, `domain:port`, `IP:port`, and full URLs
- 📸 **Auto screenshot** — screenshots named after the domain/URL, saved as PNG files
- 📊 **Dual output** — interactive dark-mode HTML report + Excel spreadsheet
- 🔍 **Filter & search** — HTML report includes per-status filters and live search
- 🖼️ **Embedded screenshots** — screenshots embedded directly in HTML, click to zoom
- 📋 **Screenshot filename reference** — filename shown in Excel column for easy auditing
- ⚡ **Optimized for scale** — handles tens of thousands of targets efficiently

---

## 📦 Installation

### 1. Clone the repository

```bash
git clone https://github.com/threatlabindonesia/webshot.git
cd webshot-checker
```

### 2. Install dependencies

```bash
pip install playwright openpyxl aiohttp
```

### 3. Install Chromium browser

```bash
playwright install chromium
```

---

## 🚀 Usage

### Single target

```bash
python webshot_checker.py --target google.com
```

### Bulk from file

```bash
python webshot_checker.py --file targets.txt
```

### HTML output + screenshots

```bash
python webshot_checker.py --file targets.txt --out report --shots-dir screenshots --format html
```

### HTML + Excel output together

```bash
python webshot_checker.py --file targets.txt --out report --shots-dir screenshots --format both
```

### Custom port target

```bash
python webshot_checker.py --target 192.168.1.1:8080 --format both
```

### Fast scan for large target lists

```bash
python webshot_checker.py --file targets.txt --concurrency 30 --timeout 10 --format both
```

### HTTPS only, skip HTTP

```bash
python webshot_checker.py --file targets.txt --no-http --format html
```

---

## 📄 Target File Format

Create a `targets.txt` file with one target per line. Lines starting with `#` are treated as comments.

```
# Web servers
google.com
github.com
cloudflare.com

# IPs with custom port
192.168.1.1:8080
10.0.0.1:9090

# Domain with non-standard port
example.com:8443

# Full URL (used as-is)
https://custom.domain/admin
```

---

## 📊 Output

### HTML Report (Dark Mode)

The generated HTML report is **self-contained** — all screenshots are embedded as base64, so it can be opened without an internet connection and shared as a single file.

HTML report features:
- 🎨 Professional dark mode UI
- 📊 Summary cards (Total, OK, Redirect, Timeout, Error, Success Rate)
- 🔍 Live search & per-status filtering
- 🖼️ Thumbnail screenshots with click-to-zoom modal
- 📱 Responsive layout

### Excel Report

The Excel report consists of 2 sheets:

| Sheet | Contents |
|-------|----------|
| `Hasil Scan` | Full data for all targets including screenshot filenames |
| `Ringkasan` | Summary statistics (total, OK, errors, success rate) |

Columns available in the main sheet:

| Column | Description |
|--------|-------------|
| Target | Original input target |
| URL Checked | Full URL accessed (http/https) |
| Status | OK / Redirect / Timeout / Error |
| HTTP Code | HTTP status code (200, 301, 403, etc.) |
| Page Title | Page title from HTML |
| Screenshot File | PNG screenshot filename |
| Final URL | URL after any redirects |
| Content-Type | Response content type |
| Response Time | Response time in ms |
| Error | Error message if failed |
| Checked At | Timestamp of the check |

### Screenshots

Screenshots are saved in the `screenshots/` folder (or the path set via `--shots-dir`), named after the target URL:

```
screenshots/
├── google.com.png
├── https_github.com.png
├── http_192.168.1.1_8080.png
└── ...
```

---

## ⚙️ Options

```
usage: webshot_checker.py [-h] [--target TARGET] [--file FILE]
                           [--out OUT] [--format {excel,html,both}]
                           [--shots-dir SHOTS_DIR]
                           [--concurrency CONCURRENCY]
                           [--timeout TIMEOUT]
                           [--no-http] [--no-https]
```

| Argument | Default | Description |
|----------|---------|-------------|
| `--target` | — | Single target (domain, IP, IP:port) |
| `--file` | — | TXT file containing list of targets |
| `--out` | `webshot_report` | Output filename prefix |
| `--format` | `both` | Output format: `excel` \| `html` \| `both` |
| `--shots-dir` | `screenshots` | Folder to save screenshots |
| `--concurrency` | `10` | Number of parallel browser tabs |
| `--timeout` | `15` | Per-request timeout in seconds |
| `--no-http` | `false` | Skip HTTP, check HTTPS only |
| `--no-https` | `false` | Skip HTTPS, check HTTP only |

### Performance tuning for large target lists

| Target Count | Recommended `--concurrency` | Recommended `--timeout` |
|--------------|----------------------------|------------------------|
| < 100 | 10 | 15 |
| 100 – 1,000 | 20 | 12 |
| 1,000 – 10,000 | 30–40 | 10 |
| > 10,000 | 50 | 8 |

> **Note:** Very high concurrency values may increase memory usage. Adjust based on your machine's available resources.

---

## 🛠️ Requirements

- Python 3.10+
- [`playwright`](https://playwright.dev/python/) — browser automation
- [`openpyxl`](https://openpyxl.readthedocs.io/) — Excel export
- [`aiohttp`](https://docs.aiohttp.org/) — async HTTP

---

## 📁 Output Structure

```
project/
├── webshot_checker.py
├── targets.txt
├── report.html              ← Self-contained HTML report
├── report.xlsx              ← Excel report
└── screenshots/
    ├── google.com.png
    ├── https_github.com.png
    └── ...
```

---

## 📝 License

MIT License — free to use and modify.

---

<div align="center">
Made with ☕ and 🐍
</div>
