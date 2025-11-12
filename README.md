# 個人工作狀態管理工具 (Personal Work Status Management Tool)

這是一個自動化工具，旨在掃描個人工作項目、更新進度，並在個人頁面上生成一份清晰的工作狀態報告。它可以幫助您和團隊成員快速了解目前的工作負載、進度以及已完成的項目，並將所有狀態變更記錄在版本控制中。

## ✨ 核心功能 (Features)

- **自動計算工作日**: 根據專案的 `start_date` 和 `due_date` 自動計算並更新 `estimated_workdays`。
- **視覺化甘特圖**: 自動為有明確日期的專案產生 [Mermaid](https://mermaid-js.github.io/mermaid/#/gantt) 甘特圖，並在圖表下方顯示生成日期，將專案時程視覺化。
- **狀態分類**: 將所有專案按狀態 (`In-Progress`, `Pending`, `Completed`) 清晰分類。
- **詳細進度報告**: 自動提取 `In-Progress` 專案的「進度報告」內容，並顯示最新的五個項目在主列表中。
- **自動排序**: 列表中的專案會根據檔案名稱自動降冪排序，確保最新的項目永遠在最上面。

## 📂 專案結構 (Project Structure)

本工具依賴特定的目錄和檔案結構來運作：

```
.
├── auto-update.py       # 主要的自動化腳本
├── people/
│   └── YourName.md      # 自動產生的個人工作狀態報告
└── projects/
    └── YYYY-MM-DD-property.md # 每個專案一個 Markdown 檔案 property: dev/research/bug
```

### 專案檔案格式

每個位於 `projects/` 目錄下的 `.md` 檔案都代表一個獨立的專案。檔案的元數據 (metadata) 必須以 YAML Frontmatter 的形式定義在檔案開頭。

**範例 (`projects/YYYY-MM-DD-property.md`):**
```yaml
---
id: 2025-11-01-bug
title: 專案標題
project: 所屬的大專案或產品線
owner: 負責人
team: 所屬團隊
status: In-Progress # 可選: Pending, In-Progress, Completed
start_date: 2025-11-01
due_date: 2025-11-30
actual_end_date: "" # 完成後填寫
---

# 專案描述標題

## 項目描述
關於這個專案的詳細說明。

## 進度報告
- **YYYY-MM-DD:** 更新事項...
- **YYYY-MM-DD:** 更新事項...
```

## ⚙️ 安裝需求 (Requirements)

- Python 3
- PyYAML

您可以使用 pip 來安裝所需的套件：
```bash
pip install PyYAML python-pptx mermaid-py cairosvg
```

## 🚀 使用方法 (Usage)

直接在專案根目錄下執行 `auto-update.py` 腳本即可。腳本會自動掃描 `projects/` 目錄並更新 `people/` 目錄下的報告檔案。

```bash
python3 auto-update.py
```

## 📊 PPTX 報告生成工具 (PPTX Report Generation Tool)

`create_presentation.py` 腳本用於將 `people/YourName.md` 檔案轉換為單頁的 PowerPoint 簡報 (`.pptx`)。這份簡報會將所有相關資訊濃縮在一個頁面中，便於快速概覽。

### 核心功能 (Features)

- **單頁簡報**: 將所有內容（包括甘特圖和專案列表）濃縮在單一投影片中。
- **甘特圖圖片嵌入**: 自動將 Mermaid 甘特圖渲染為圖片並嵌入投影片，而非原始程式碼。
- **16:10 比例**: 投影片比例設定為主流的 16:10，並確保內容佈局合理不重疊。
- **中文字體支援**: 確保簡報中的中文字體正常顯示。
- **內容格式化**: 自動移除連結、多餘空白，並將專案狀態標題設為粗體第一層級，同時在內容開頭加入生成日期。

### 安裝需求 (Requirements)

除了上述的 Python 3 和 PyYAML，此工具還需要以下套件：
- `python-pptx`
- `mermaid-py`
- `cairosvg`

您可以使用 pip 來安裝所需的套件：
```bash
pip install python-pptx mermaid-py cairosvg
```

### 使用方法 (Usage)

在專案根目錄下執行 `create_presentation.py` 腳本即可。腳本會讀取 `people/` 目錄下的 Markdown 報告，並在同一個目錄中生成 `.pptx` 檔案。

```bash
python3 create_presentation.py
```
