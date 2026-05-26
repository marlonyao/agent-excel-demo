# Agent Excel Demo

通过自然语言操作 Excel 的 Multi-Agent 演示项目。

## 架构

```
┌─ 浏览器 ─────────────────────────────────────┐
│                                               │
│   Luckysheet (Excel 组件)                      │
│        ↕ Tool Runtime                         │
│        ↕ WebSocket                            │
│                                               │
└───────────────┬───────────────────────────────┘
                │
┌───────────────▼───────────────────────────────┐
│  Server (Node.js)                             │
│                                               │
│   Master Agent ──▶ Excel Agent ──▶ LLM API    │
│     (路由)          (意图→工具)                  │
│                                               │
└───────────────────────────────────────────────┘
```

### 通讯流

1. 用户在浏览器输入自然语言（如 "B列求和放到C1"）
2. 前端通过 WebSocket 发给 Server
3. Master Agent 接收，交给 Excel Agent
4. Excel Agent 调用 LLM，LLM 返回 tool_calls
5. tool_calls 通过 WebSocket 发给浏览器执行
6. 浏览器端 Tool Runtime 调用 Luckysheet API 执行操作
7. 结果返回 Excel Agent，Excel Agent 再调 LLM 生成回复
8. 最终回复通过 WebSocket 返回给用户

## 快速开始

```bash
# 1. 安装依赖
cd agent-excel-demo
npm install

# 2. 配置 LLM API（支持任何 OpenAI 兼容 API）
export LLM_BASE_URL="https://api.openai.com/v1"  # 或你的 LiteLLM 地址
export LLM_API_KEY="sk-xxx"
export LLM_MODEL="gpt-4o-mini"  # 或 glm-4-flash 等

# 3. 启动
npm run dev

# 4. 打开浏览器
# http://localhost:3000
```

## 使用示例

在聊天框中输入：

- `帮我填入以下数据：A2到A5填入张三、李四、王五、赵六，B2到B5填入1500、2300、1800、3200`
- `B列求和放到C1`
- `计算B列的平均值放到C2`
- `把标题行加粗标黄`
- `找出销售额最高的人是谁`
- `在C列标注，大于2000的标"优"，其他标"良"`

## 技术栈

- **前端**: Luckysheet (开源在线 Excel 组件) + WebSocket
- **后端**: Express + ws (WebSocket)
- **LLM**: 任何 OpenAI 兼容 API (通过 function calling)
- **Agent 模式**: Master Agent (路由) + Excel Agent (工具调用)

## 项目结构

```
agent-excel-demo/
├── server.js          # 服务端：Master Agent + Excel Agent + WebSocket
├── public/
│   └── index.html     # 前端：Luckysheet + Tool Runtime + Chat UI
├── package.json
└── README.md
```

## 扩展思路

1. **加更多 Agent**: 比如图表 Agent（调用 ECharts）、数据分析 Agent
2. **加 A2A 协议**: 把 Excel Agent 改为 A2A Remote Agent
3. **加 MCP**: 把 Tool Runtime 封装为 MCP Server
4. **加权限控制**: 限制 Agent 可操作的单元格范围
5. **加操作历史**: 记录 Agent 的每一步操作，支持撤销
