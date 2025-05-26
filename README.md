# MCP Quote Bot - 報價單生成工具

這是一個基於 MCP (Model Context Protocol) 的報價單生成工具，可以在 Cursor 中直接使用。

## 🚀 快速開始

### 方法一：從 GitHub 安裝

```bash
# 1. 克隆專案
git clone https://github.com/nccu231015/Invoice-Word-MCP.git
cd Invoice-Word-MCP

# 2. 安裝依賴
pip install -r requirements.txt
```

### 方法二：手動下載

1. 下載專案 ZIP 檔案並解壓縮
2. 進入專案目錄並安裝依賴：
```bash
pip install -r requirements.txt
```

### 配置 Cursor MCP

在您的 `~/.cursor/mcp.json` 文件中添加以下配置：

```json
{
  "mcpServers": {
    "quote-bot-word": {
      "command": "python",
      "args": ["/path/to/your/Invoice-Word-MCP/mcp_server_stdio.py"],
      "cwd": "/path/to/your/Invoice-Word-MCP"
    }
  }
}
```

**重要**：請將 `/path/to/your/Invoice-Word-MCP` 替換為您實際的專案路徑。

### 重啟 Cursor

配置完成後，重啟 Cursor 即可使用。

## 🤖 專業報價機器人模式

### 觸發條件
當您在 Cursor 中提到以下關鍵詞時，會自動啟用專業報價機器人模式：
- 「報價單」、「quote」、「pricing」、「MCP」、「報價」、「估價」

### 工作流程
1. **確認客戶需求和功能規格**
2. **搜尋相關案例和市場價格參考**
   - 使用 Astra DB MCP 工具搜尋 invoice 資料庫
   - 若找不到，使用 Tavily AI MCP 工具搜尋台灣市場平均報價
3. **生成標準 JSON 格式報價數據**
4. **使用 MCP 工具生成 Word 文檔**
5. **提供本地文件路徑**

### 定價原則
- 基於市場調研和案例參考
- 提供多個方案選擇（基礎/進階/豪華）
- 考慮技術複雜度和開發時間
- 包含適當的利潤空間

## 📝 使用方法

在 Cursor 中，您可以使用以下工具：

### `generate_quote_docs`
根據 JSON 數據生成報價單 Word 文檔。

**參數**：
- `json_content`: JSON 格式的報價單數據
- `json_file_path`: JSON 文件路徑（可選）

**示例**：
```
請幫我生成報價單，包含以下功能：
1. 使用者註冊與登入
2. 首頁與推薦內容
3. 看板管理
4. 互動功能
```

## 📋 JSON 數據格式

### 標準結構
```json
{
  "quotes": [
    {
      "header": {
        "companyName": "亦式數位互動有限公司",
        "companyContact": "0988363357",
        "companyEmail": "istudiodesign.tw@gmail.com",
        "quoteNumber": "Q-YYYY-MMDD",
        "start_date": "YYYY/MM/DD",
        "end_date": "YYYY/MM/DD",
        "staff": "亦式數位互動有限公司（負責人）",
        "key": "96790278",
        "recipient": "客戶 先生/小姐",
        "Title": "方案標題"
      },
      "details": [
        {
          "category": "類別名稱",
          "items": "項目描述",
          "unit": 1,
          "quantity": 1,
          "amount": 50000
        }
      ],
      "total_without_tax": 50000,
      "discount": 0,
      "tax_rate": 2500,
      "total_with_tax": 52500,
      "notes": "備註說明"
    }
  ]
}
```

### 格式要求
1. **quoteNumber 格式**：必須以 "Q-" 開頭，如 "Q-2024-0523"
2. **日期格式**：使用 "YYYY/MM/DD" 格式，如 "2024/05/23"
3. **tax_rate**：必須是實際稅額數值，非百分比
4. **companyContact**：必須是電話號碼
5. **數字欄位**：必須是數字，不能含逗點或其他格式符號
6. **多方案**：如需產出兩個不同方案，請用 array 方式隔開

## 📁 輸出文件

生成的 Word 文檔會保存在 `temp/` 目錄中。

## 🛠️ 故障排除

### 問題：Cursor 顯示 "no tools available"
**解決方案**：
1. 確認路徑配置正確
2. 確認已安裝所有依賴
3. 重啟 Cursor

### 問題：生成文件失敗
**解決方案**：
1. 確認 `報價單.docx` 模板文件存在
2. 確認 `temp/` 目錄有寫入權限
3. 檢查 JSON 數據格式是否正確

### 問題：JSON 格式錯誤
**解決方案**：
1. 確保生成純 JSON 數據，不要添加 markdown 標記
2. 檢查所有必填欄位是否完整
3. 確認數字格式正確

## 📞 支援

如有問題，請聯絡：istudiodesign.tw@gmail.com 