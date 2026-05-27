# Agent Excel Demo

通过自然语言操作 Excel 的 Multi-Agent 演示项目。

## 架构

```
┌─ 浏览器 ──────────────────────────────────────────────┐
│                                                        │
│  ┌──────────┐  ┌──────────────┐  ┌─────────────────┐  │
│  │ 左侧菜单  │  │  内容区域     │  │  右侧 AI 助手   │  │
│  │ 仪表盘   │  │  Dashboard   │  │  💬 聊天面板     │  │
│  │ Excel    │  │  Excel编辑器 │  │  推荐问题        │  │
│  │ 笔记     │  │  (Luckysheet)│  │  历史会话        │  │
│  │ 设置     │  │  笔记/设置   │  │  预览面板        │  │
│  └──────────┘  └──────────────┘  └─────────────────┘  │
│                      ↕ Tool Runtime                     │
│                      ↕ WebSocket                        │
└──────────────────────┬──────────────────────────────────┘
                       │
┌──────────────────────▼──────────────────────────────────┐
│  Server (Node.js)                                       │
│                                                         │
│  ┌─────────────┐    ┌──────────────┐    ┌───────────┐  │
│  │ Master Agent│───▶│ Excel Agent  │───▶│ LLM API   │  │
│  │  (路由)     │    │ (意图→工具)  │    │(DeepSeek) │  │
│  └─────────────┘    └──────────────┘    └───────────┘  │
│         │                                                  │
│         └──────────▶ General Agent ──▶ LLM API           │
│                                                         │
│  REST API: /api/dashboard, /api/files, /api/notes       │
└─────────────────────────────────────────────────────────┘
```

---

## 关键流程顺序图

### 1. 用户聊天 → Agent 操作 Excel

用户输入自然语言（如"增加4月销售额"），Excel Agent 解析意图、调用工具、返回结果。

```mermaid
sequenceDiagram
    actor User
    participant FE as 前端 (index.html)
    participant WS as WebSocket
    participant Server as Server (Node.js)
    participant Master as Master Agent
    participant Excel as Excel Agent
    participant LLM as LLM API (DeepSeek)
    participant TR as Tool Runtime (浏览器)

    User->>FE: 输入 "增加4月销售额"
    FE->>FE: getPageContext() 获取页面数据+选区
    FE->>WS: {type:"chat", content, pageContext}
    WS->>Server: 转发消息
    Server->>Master: masterAgent(msg, sessionId, pageContext)

    Note over Master: 路由判断：excel 还是 general
    Master->>LLM: 路由 prompt + 用户消息
    LLM-->>Master: "excel"
    Master->>Excel: excelAgent(task, sessionId, pageContext)

    Note over Excel: 多轮工具调用循环 (max 10 轮)
    
    Excel->>LLM: system prompt + 选区上下文 + 用户消息
    LLM-->>Excel: get_sheet_info (tool_call)
    Excel->>WS: {type:"tool_call", method:"get_sheet_info"}
    WS->>FE: 转发 tool_call
    FE->>TR: executeTool() → luckysheet.getSheetData()
    TR-->>FE: {rows, cols, headers:["姓名","1月销售额",...]}
    FE->>WS: {type:"tool_result", result}
    WS->>Excel: 返回工具结果

    Excel->>LLM: 消息 + 工具结果
    LLM-->>Excel: insert_column(col:"E", header:"4月销售额")
    Excel->>WS: {type:"tool_call", method:"insert_column"}
    WS->>FE: 转发 tool_call
    FE->>TR: executeTool() → celldata 右移 + luckysheet.destroy/create
    TR-->>FE: {ok:true}
    FE->>WS: {type:"tool_result", result}

    Excel->>LLM: 消息 + 插入列结果
    LLM-->>Excel: set_range_values(startRow:2, startCol:"E", values:[[25800],[28100]...])
    Excel->>WS: {type:"tool_call", method:"set_range_values"}
    WS->>FE: 转发
    FE->>TR: luckysheet.setCellValue() × N
    TR-->>FE: {ok:true}
    FE->>WS: {type:"tool_result", result}

    Excel->>LLM: 消息 + 填入数据结果
    LLM-->>Excel: (无 tool_calls，最终回复)
    Excel-->>Server: {reply: "已完成！新增4月销售额列..."}
    Server->>WS: {type:"chat_reply", content}
    WS->>FE: 显示 agent 回复
    FE->>User: 💬 "已完成！新增4月销售额列..."
```

### 2. 打开 Excel 文件 → 自动总结

用户进入 Excel 页面，自动加载默认数据并生成 AI 摘要。

```mermaid
sequenceDiagram
    actor User
    participant FE as 前端
    participant Server as Server
    participant LLM as LLM API
    participant LS as Luckysheet

    User->>FE: 点击"Excel 编辑器"菜单
    FE->>Server: GET /api/files
    Server-->>FE: [{id:1, name:"销售数据_Q1.xlsx"}, ...]
    FE->>FE: renderExcelFileList() 渲染左侧文件列表

    Note over FE: 默认打开第一个文件
    FE->>FE: openExcel(1, "销售数据_Q1.xlsx")
    FE->>FE: getFileCelldata() 获取预置数据
    FE->>LS: luckysheet.create({celldata: [...]})
    LS-->>FE: Excel 渲染完成

    Note over FE: 延迟 1.5s 后自动总结
    FE->>FE: autoSummarizeExcel("销售数据_Q1.xlsx")
    FE->>FE: 读取表头 + 前5行数据
    FE->>Server: WS {type:"chat", content:"请简要总结..."}
    Server->>LLM: generalAgent(总结请求)
    LLM-->>Server: "这是一个销售数据表格..."
    Server-->>FE: {type:"chat_reply", content}
    FE->>FE: isAutoSummary=true → 显示为蓝色系统提示
    FE->>User: 📋 销售数据_Q1.xlsx - 这是一个销售数据表格...
```

### 3. 全屏聊天 → 点击文件名预览

全屏模式下，agent 回复中的文件名可点击，右侧弹出预览。

```mermaid
sequenceDiagram
    actor User
    participant FE as 前端
    participant Server as Server
    participant LLM as LLM API

    User->>FE: 点击 ⛶ 全屏按钮
    FE->>FE: syncToFullscreen() → 同步当前消息到全屏
    FE->>FE: renderHistoryList() → 左侧显示会话列表

    User->>FE: 输入 "有哪些文件"
    FE->>Server: WS {type:"chat", content}
    Server->>LLM: generalAgent(文件列表请求)
    LLM-->>Server: "当前有3个文件：销售数据_Q1.xlsx..."
    Server-->>FE: {type:"chat_reply", content}

    FE->>FE: renderPreviewLinks() → 文件名 → 🔗可点击链接
    FE->>User: "当前有3个文件：🔗销售数据_Q1.xlsx..."

    User->>FE: 点击 🔗销售数据_Q1.xlsx
    FE->>FE: openPreview('excel', 1)

    alt 全屏模式
        FE->>FE: getFileCelldata() → 构建 HTML 表格
        FE->>User: 右侧预览面板弹出 Excel 表格
    else 侧边栏模式
        FE->>FE: openExcel(1, "销售数据_Q1.xlsx")
        FE->>User: 跳转到 Excel 编辑器页面
    end
```

### 4. 选区感知 → AI 理解用户选中了什么

用户在 Excel 中选中一个区域，聊天时 AI 自动获取选区数据。

```mermaid
sequenceDiagram
    actor User
    participant FE as 前端
    participant LS as Luckysheet
    participant Server as Server
    participant LLM as LLM API

    User->>LS: 鼠标选中 B2:B7
    LS-->>FE: 选区变化

    User->>FE: 输入 "这些数据求和"
    FE->>FE: getPageContext()
    FE->>LS: luckysheet.getRange()
    LS-->>FE: {row:[1,6], column:[1,1]}
    FE->>LS: luckysheet.getCellValue() × 6
    LS-->>FE: [15200, 23800, 18500, 32100, 19800, 27400]

    Note over FE: pageContext.data = {selectedRange:"B2:B7", selectedData:[[15200],...]}
    FE->>Server: WS {type:"chat", pageContext:{...}}
    Server->>LLM: excelAgent(task, sessionId, pageContext)
    Note over LLM: system prompt 包含选区信息:<br/>选中区域 B2:B7<br/>数据: [15200, 23800, ...]
    LLM-->>Server: 理解"这些数据"= B2:B7, 返回操作
    Server-->>FE: chat_reply
    FE->>User: "已对选中区域 B2:B7 求和..."
```

### 5. 切换 Excel 文件

用户在左侧文件列表点击不同文件，Excel 编辑器切换数据。

```mermaid
sequenceDiagram
    actor User
    participant FE as 前端
    participant LS as Luckysheet

    User->>FE: 点击左侧 "项目预算.xlsx"
    FE->>FE: openExcel(2, "项目预算.xlsx")
    FE->>FE: renderExcelFileList() → 高亮切换
    FE->>FE: getFileCelldata("项目预算.xlsx")
    FE->>LS: luckysheet.destroy()
    LS-->>FE: 销毁旧实例
    FE->>LS: luckysheet.create({celldata: 项目预算数据})
    LS-->>FE: 新数据渲染完成
    FE->>FE: autoSummarizeExcel("项目预算.xlsx")
    FE->>User: 📋 项目预算.xlsx - 包含5个项目的预算使用情况...
```

---

## 快速开始

```bash
# 1. 克隆
git clone https://github.com/marlonyao/agent-excel-demo.git
cd agent-excel-demo

# 2. 安装依赖
npm install

# 3. 配置环境变量
cp .env.example .env
# 编辑 .env，填入你的 API Key

# 4. 启动
npm run dev

# 5. 打开浏览器
# http://localhost:3000
```

## 使用示例

在聊天框中输入：

- `帮我填入以下数据：A2到A5填入张三、李四、王五、赵六，B2到B5填入1500、2300、1800、3200`
- `B列求和放到C1`
- `增加4月销售额，数据分别是：张三25800，李四28100...`
- `把标题行加粗标黄`
- `这些数据求个和`（先选中 Excel 区域）
- `有哪些文件`（文件名可点击）

## 技术栈

- **前端**: 纯 HTML/CSS/JS + Luckysheet (在线 Excel) + WebSocket
- **后端**: Express + ws (WebSocket)
- **LLM**: 任何 OpenAI 兼容 API (通过 function calling)
- **Agent 模式**: Master Agent (路由) + Excel Agent (工具调用) + General Agent (通用)

## 项目结构

```
agent-excel-demo/
├── server.js          # 服务端：Master Agent + Excel Agent + WebSocket
├── public/
│   └── index.html     # 前端：Luckysheet + Tool Runtime + Chat UI + 全屏模式
├── package.json
└── README.md
```

## 扩展思路

1. **加更多 Agent**: 比如图表 Agent（调用 ECharts）、数据分析 Agent
2. **加 A2A 协议**: 把 Excel Agent 改为 A2A Remote Agent
3. **加 MCP**: 把 Tool Runtime 封装为 MCP Server
4. **加权限控制**: 限制 Agent 可操作的单元格范围
5. **加操作历史**: 记录 Agent 的每一步操作，支持撤销
