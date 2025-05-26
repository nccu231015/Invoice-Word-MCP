import os
import json
import asyncio
import logging
from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp import types

# 設置日誌
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger("mcp-server-stdio")

# 導入報價單生成功能
from generate_quote_docs import generate_docs

# 確保 temp 目錄存在
def ensure_temp_dir():
    """確保臨時目錄存在"""
    temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "temp")
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir, exist_ok=True)
    return temp_dir

# 建立 MCP Server
app_server = Server("quote-bot-word")

# 註冊工具處理程序
@app_server.call_tool()
async def handle_tool_call(name: str, arguments: dict | None) -> list[types.TextContent | types.ImageContent | types.EmbeddedResource]:
    """處理工具調用"""
    if name == "generate_quote_docs":
        try:
            # 如果 arguments 為 None，設為空字典
            if arguments is None:
                arguments = {}
                
            logger.info(f"=== MCP 工具調用開始 ===")
            logger.info(f"接收報價單生成請求，工具名稱: {name}")
            logger.info(f"原始參數類型: {type(arguments)}")
            logger.info(f"原始參數內容: {json.dumps(arguments, ensure_ascii=False, indent=2)}")
            
            # 初始化文件數據
            file_data = None
            
            # 方法1：從文件路徑讀取
            if "json_file_path" in arguments and arguments["json_file_path"]:
                file_path = arguments["json_file_path"]
                logger.info(f"嘗試從文件路徑讀取: {file_path}")
                
                if os.path.exists(file_path):
                    try:
                        with open(file_path, 'r', encoding='utf-8') as f:
                            file_data = json.load(f)
                        logger.info(f"成功從文件讀取數據: {file_path}")
                    except Exception as e:
                        logger.error(f"讀取文件失敗: {e}")
                        return [types.TextContent(type="text", text=f"讀取JSON文件失敗: {str(e)}")]
                else:
                    logger.warning(f"文件不存在: {file_path}")
            
            # 方法2：從JSON內容讀取
            elif "json_content" in arguments and arguments["json_content"]:
                json_content = arguments["json_content"]
                logger.info(f"嘗試解析JSON內容，長度: {len(json_content)} 字符")
                
                try:
                    file_data = json.loads(json_content)
                    logger.info(f"成功解析JSON內容")
                except json.JSONDecodeError as e:
                    logger.error(f"解析JSON內容失敗: {e}")
                    return [types.TextContent(type="text", text=f"解析JSON內容失敗: {str(e)}")]
            
            # 方法3：如果沒有提供文件，使用備用數據
            if file_data is None:
                logger.info("未提供有效的JSON文件或內容，使用備用數據")
                backup_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "input.json")
                if os.path.exists(backup_path):
                    try:
                        with open(backup_path, 'r', encoding='utf-8') as f:
                            file_data = json.load(f)
                        logger.info(f"已從備用文件載入數據: {backup_path}")
                    except Exception as e:
                        logger.error(f"讀取備用文件失敗: {e}")
                        return [types.TextContent(type="text", text=f"讀取備用數據失敗: {str(e)}")]
                else:
                    logger.error(f"無法找到備用數據文件: {backup_path}")
                    return [types.TextContent(type="text", text="未提供JSON文件且無備用數據")]
            
            # 驗證數據格式
            if not isinstance(file_data, dict):
                logger.error(f"JSON數據必須是字典格式，實際類型: {type(file_data)}")
                return [types.TextContent(type="text", text="JSON數據格式錯誤：必須是字典格式")]
            
            # 檢查是否包含quotes字段
            if "quotes" not in file_data:
                logger.error("JSON數據缺少'quotes'字段")
                return [types.TextContent(type="text", text="JSON數據缺少'quotes'字段")]
            
            if not isinstance(file_data["quotes"], list) or len(file_data["quotes"]) == 0:
                logger.error("'quotes'必須是非空列表")
                return [types.TextContent(type="text", text="'quotes'必須是非空列表")]
            
            # 記錄報價單信息
            logger.info(f"JSON文件包含 {len(file_data['quotes'])} 個報價單:")
            for i, quote in enumerate(file_data["quotes"]):
                if isinstance(quote, dict) and "header" in quote:
                    header = quote["header"]
                    logger.info(f"Quote {i+1}: {header.get('quoteNumber', 'N/A')} - {header.get('Title', 'N/A')}")
                    logger.info(f"Quote {i+1}: recipient={header.get('recipient', 'N/A')}")
                    logger.info(f"Quote {i+1}: details count={len(quote.get('details', []))}")
                else:
                    logger.warning(f"Quote {i+1}: 格式不正確或缺少header字段")
            
            # 確保 temp 目錄存在
            ensure_temp_dir()
            
            # 直接將文件數據傳遞給 generate_docs 函數
            doc_paths = generate_docs(file_data)
            
            # 確保生成的文檔存在
            if not doc_paths or len(doc_paths) == 0:
                logger.error("未能生成任何報價單文檔")
                return [types.TextContent(type="text", text="未能生成任何報價單文檔")]
            
            # 構建結果 - STDIO 模式下只返回本地文件路徑
            result_content = []
            for path in doc_paths:
                filename = os.path.basename(path)
                
                # 確認文件確實存在
                if not os.path.exists(path):
                    logger.warning(f"生成的文件不存在: {path}")
                    continue
                
                logger.info(f"已生成報價單: {filename}, 文件路徑: {path}")
                # 在 STDIO 模式下，只提供本地文件路徑
                result_content.append(types.TextContent(
                    type="text", 
                    text=f"已生成報價單文檔: {filename}\n文件路徑: {path}"
                ))
            
            if not result_content:
                return [types.TextContent(type="text", text="生成的報價單文件無法訪問")]
            
            logger.info(f"=== MCP 工具調用完成，生成了 {len(result_content)} 個文檔 ===")
            return result_content
            
        except Exception as e:
            logger.error(f"工具執行失敗: {str(e)}", exc_info=True)
            return [types.TextContent(type="text", text=f"文件生成失敗: {str(e)}")]
    else:
        return [types.TextContent(type="text", text=f"不支援的工具: {name}")]

# 註冊工具列表
@app_server.list_tools()
async def list_tools():
    """返回可用的工具列表"""
    return [
        types.Tool(
            name="generate_quote_docs",
            description="根據提供的JSON文件生成報價單 Word 文檔",
            inputSchema={
                "type": "object",
                "properties": {
                    "json_file_path": {
                        "type": "string",
                        "description": "包含報價單數據的JSON文件路徑"
                    },
                    "json_content": {
                        "type": "string", 
                        "description": "JSON文件的内容（如果无法传递文件路径）"
                    }
                },
                "required": []  # 两个参数至少需要一个
            }
        )
    ]

# 主函數
async def main():
    """運行 STDIO MCP server"""
    logger.info("啟動 MCP Server (STDIO 模式)")
    
    # 確保 temp 目錄存在
    ensure_temp_dir()
    
    # 使用 stdio_server 運行
    async with stdio_server() as (read_stream, write_stream):
        await app_server.run(
            read_stream,
            write_stream,
            app_server.create_initialization_options()
        )

if __name__ == "__main__":
    asyncio.run(main()) 