# 個人工作狀態管理工具 (Personal Work Status Management Tool)

這是一個自動化工具，旨在掃描個人工作項目、更新進度，並在個人頁面上生成一份清晰的工作狀態報告。它可以幫助您和團隊成員快速了解目前的工作負載、進度以及已完成的項目，並將所有狀態變更記錄在版本控制中。

## ✨ 核心功能 (Features)

- **自動計算工作日**: 根據專案的 `start_date` 和 `due_date` 自動計算並更新 `estimated_workdays`。
- **視覺化甘特圖**: 自動為有明確日期的專案產生 [Mermaid](https://mermaid-js.github.io/mermaid/#/gantt) 甘特圖，將專案時程視覺化。
- **狀態分類**: 將所有專案按狀態 (`In-Progress`, `Pending`, `Completed`) 清晰分類。
- **詳細進度報告**: 自動提取 `In-Progress` 專案的「進度報告」內容，並顯示在主列表中。
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
pip install PyYAML
```

## 🚀 使用方法 (Usage)

直接在專案根目錄下執行 `auto-update.py` 腳本即可。腳本會自動掃描 `projects/` 目錄並更新 `people/` 目錄下的報告檔案。

```bash
python3 auto-update.py
```
