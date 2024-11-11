[English Version](README_EN.md)

# 機車位抽籤系統

文華天際社區機車停車位的抽籤系統，使用 Python 開發，包含 `tkinter` 圖形介面，並透過 `xlwings` 實現 Excel 文件的即時更新，讓抽籤結果即時顯示於開啟的 Excel 文件中。

## 功能
- 支援社區內不同棟別的機車位抽籤（例如 A棟、B棟）。
- 自動抽取參加者並分配停車位。
- 抽籤結果直接寫入 Excel 檔案，開啟文件時即時顯示更新。

## 測試影片
[測試影片連結](https://drive.google.com/file/d/1J9sSs46bFs52iwFfptUgivsCy5NFHmKE/view?usp=sharing)

## 系統需求
- Python 3.x
- 相關套件：
  - `openpyxl`（用於操作 Excel 檔案）
  - `xlwings`（用於即時更新 Excel）
  - `tkinter`（用於圖形介面，Python 預設包含）

## 安裝步驟

1. 克隆此倉庫：
   ```bash
   git clone https://github.com/leo851001/motor_lottery.git
   cd motor_lottery
   ```

2. 安裝所需套件：
   ```bash
   pip install openpyxl xlwings
   ```

## 使用說明

1. 執行應用程式：
   ```bash
   python motor_lottery_gui.py
   ```
2. 點擊「A棟抽籤」或「B棟抽籤」以開始對應棟別的抽籤。
3. 完成一棟的抽籤後，您可以立即點選另一棟按鈕繼續抽籤。

## 注意事項

- 請確保 Excel 檔案保持開啟，以便即時查看更新結果。
- 應用程式會自動將抽籤結果寫入指定的 Excel 檔案。
