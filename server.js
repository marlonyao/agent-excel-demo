/**
 * Agent Excel Demo - Server
 *
 * 架构:
 *   用户浏览器 ◀──WebSocket──▶ Server ◀──HTTP──▶ LLM API
 *        │                        │
 *   Luckysheet (Excel组件)    Master Agent + Excel Agent
 *
 * 通讯流:
 *   1. 用户在浏览器输入自然语言
 *   2. 前端通过 WebSocket 发给 Server
 *   3. Server 的 Master Agent 判断意图
 *   4. 如果是 Excel 操作，转发给 Excel Agent
 *   5. Excel Agent 调用 LLM 生成 tool_calls
 *   6. tool_calls 通过 WebSocket 发回浏览器执行
 *   7. 执行结果返回给 Excel Agent，最终回复用户
 */

const express = require('express');
const http = require('http');
const { WebSocketServer } = require('ws');
const path = require('path');

const app = express();
const server = http.createServer(app);
const wss = new WebSocketServer({ server });

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ==========================================
// LLM 配置 - 支持任何 OpenAI 兼容 API
// ==========================================
const LLM_CONFIG = {
  // 修改为你的 LLM API 地址
  baseURL: process.env.LLM_BASE_URL || 'https://api.openai.com/v1',
  apiKey: process.env.LLM_API_KEY || 'your-api-key',
  model: process.env.LLM_MODEL || 'gpt-4o-mini',
};

async function callLLM(messages, tools = null) {
  const body = {
    model: LLM_CONFIG.model,
    messages,
    temperature: 0.1,
  };
  if (tools) {
    body.tools = tools;
    body.tool_choice = 'auto';
  }

  const resp = await fetch(`${LLM_CONFIG.baseURL}/chat/completions`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${LLM_CONFIG.apiKey}`,
    },
    body: JSON.stringify(body),
  });

  if (!resp.ok) {
    const text = await resp.text();
    throw new Error(`LLM API error: ${resp.status} ${text}`);
  }

  return resp.json();
}

// ==========================================
// Excel Tools 定义 - Excel Agent 能调用的工具
// ==========================================
const EXCEL_TOOLS = [
  {
    type: 'function',
    function: {
      name: 'get_cell_value',
      description: '获取指定单元格的值。row 为行号(数字，从1开始)，col 为列号(字母，如 A, B, C)',
      parameters: {
        type: 'object',
        properties: {
          row: { type: 'number', description: '行号，从1开始' },
          col: { type: 'string', description: '列号字母，如 A, B, C' },
        },
        required: ['row', 'col'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'set_cell_value',
      description: '设置指定单元格的值或公式。row 为行号(数字，从1开始)，col 为列号(字母)，value 为要设置的值或公式(如 =SUM(B1:B10))',
      parameters: {
        type: 'object',
        properties: {
          row: { type: 'number', description: '行号，从1开始' },
          col: { type: 'string', description: '列号字母，如 A, B, C' },
          value: { type: 'string', description: '要设置的值，公式以=开头' },
        },
        required: ['row', 'col', 'value'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'get_range_values',
      description: '获取一个矩形区域的值。返回二维数组。',
      parameters: {
        type: 'object',
        properties: {
          startRow: { type: 'number', description: '起始行号' },
          startCol: { type: 'string', description: '起始列号字母' },
          endRow: { type: 'number', description: '结束行号' },
          endCol: { type: 'string', description: '结束列号字母' },
        },
        required: ['startRow', 'startCol', 'endRow', 'endCol'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'set_range_values',
      description: '批量设置一个区域的值。从指定起始单元格开始，按二维数组填充。',
      parameters: {
        type: 'object',
        properties: {
          startRow: { type: 'number', description: '起始行号' },
          startCol: { type: 'string', description: '起始列号字母' },
          values: {
            type: 'array',
            items: { type: 'array', items: { type: 'string' } },
            description: '二维数组，每个元素是要设置的值或公式',
          },
        },
        required: ['startRow', 'startCol', 'values'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'get_sheet_info',
      description: '获取当前工作表的基本信息：总行数、总列数、已有数据的范围',
      parameters: { type: 'object', properties: {} },
    },
  },
  {
    type: 'function',
    function: {
      name: 'apply_cell_style',
      description: '设置单元格样式（加粗、背景色、字体大小等）',
      parameters: {
        type: 'object',
        properties: {
          row: { type: 'number', description: '行号' },
          col: { type: 'string', description: '列号字母' },
          bold: { type: 'boolean', description: '是否加粗' },
          bgColor: { type: 'string', description: '背景色，如 #ffff00' },
          fontColor: { type: 'string', description: '字体颜色' },
          fontSize: { type: 'number', description: '字体大小' },
        },
        required: ['row', 'col'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'insert_formula',
      description: '在指定单元格插入公式，如 =SUM, =AVERAGE, =MAX 等',
      parameters: {
        type: 'object',
        properties: {
          row: { type: 'number', description: '行号' },
          col: { type: 'string', description: '列号字母' },
          formula: { type: 'string', description: '公式，如 =SUM(B1:B10)' },
        },
        required: ['row', 'col', 'formula'],
      },
    },
  },
];

// ==========================================
// Browser Connections - 浏览器端 Tool Runtime 连接
// ==========================================
const browserClients = new Map(); // sessionId -> ws

// Pending tool call requests
const pendingRequests = new Map(); // requestId -> { resolve, reject, timer }

// ==========================================
// Master Agent - 调度器
// ==========================================
async function masterAgent(userMessage, sessionId) {
  // Master Agent 判断用户意图
  // 这里简化：所有操作都交给 Excel Agent
  // 实际项目中，Master Agent 会有自己的 LLM 来做路由判断
  console.log(`[Master] 收到用户消息: ${userMessage}`);

  const result = await excelAgent(userMessage, sessionId);
  return result;
}

// ==========================================
// Excel Agent - Excel 操作专家
// ==========================================
async function excelAgent(task, sessionId) {
  console.log(`[ExcelAgent] 处理任务: ${task}`);

  const browserWs = browserClients.get(sessionId);
  if (!browserWs) {
    return '错误：浏览器未连接，无法操作 Excel';
  }

  // Excel Agent 的对话历史（多轮 tool calling）
  const messages = [
    {
      role: 'system',
      content: `你是一个 Excel 操作助手。用户会用自然语言描述想要对 Excel 表格做的操作。
你需要将用户的意图转换为具体的 Excel 工具调用。

规则：
1. 先了解当前数据（必要时调用 get_sheet_info 或 get_range_values）
2. 然后执行操作
3. 操作完成后，用简洁的中文告诉用户你做了什么
4. 对于求和、平均值等计算，使用公式（如 =SUM(B1:B10)）而不是手动计算`,
    },
    { role: 'user', content: task },
  ];

  // 多轮 tool calling loop
  const MAX_ROUNDS = 10;
  for (let round = 0; round < MAX_ROUNDS; round++) {
    console.log(`[ExcelAgent] LLM 调用轮次 ${round + 1}`);
    const response = await callLLM(messages, EXCEL_TOOLS);
    const choice = response.choices[0];
    const assistantMsg = choice.message;

    messages.push(assistantMsg);

    // 如果没有 tool_calls，说明 LLM 给出了最终回复
    if (!assistantMsg.tool_calls || assistantMsg.tool_calls.length === 0) {
      console.log(`[ExcelAgent] 最终回复: ${assistantMsg.content}`);
      return assistantMsg.content;
    }

    // 执行每个 tool_call
    for (const toolCall of assistantMsg.tool_calls) {
      const { name, arguments: argsStr } = toolCall.function;
      const args = JSON.parse(argsStr);
      console.log(`[ExcelAgent] 调用工具: ${name}`, args);

      try {
        // 通过 WebSocket 发给浏览器端执行
        const result = await executeOnBrowser(sessionId, name, args);
        console.log(`[ExcelAgent] 工具结果: ${JSON.stringify(result).slice(0, 200)}`);

        messages.push({
          role: 'tool',
          tool_call_id: toolCall.id,
          content: typeof result === 'string' ? result : JSON.stringify(result),
        });
      } catch (err) {
        console.error(`[ExcelAgent] 工具执行失败: ${err.message}`);
        messages.push({
          role: 'tool',
          tool_call_id: toolCall.id,
          content: `错误: ${err.message}`,
        });
      }
    }
  }

  return '操作轮次过多，请简化你的请求。';
}

// ==========================================
// 通过 WebSocket 调用浏览器端 Tool Runtime
// ==========================================
function executeOnBrowser(sessionId, method, params) {
  return new Promise((resolve, reject) => {
    const ws = browserClients.get(sessionId);
    if (!ws) {
      reject(new Error('浏览器未连接'));
      return;
    }

    const requestId = `${Date.now()}-${Math.random().toString(36).slice(2)}`;
    const timeout = setTimeout(() => {
      pendingRequests.delete(requestId);
      reject(new Error('浏览器执行超时'));
    }, 30000);

    pendingRequests.set(requestId, { resolve, reject, timeout });

    ws.send(JSON.stringify({ type: 'tool_call', id: requestId, method, params }));
  });
}

// ==========================================
// WebSocket 连接处理
// ==========================================
wss.on('connection', (ws) => {
  console.log('[WS] 浏览器客户端已连接');
  let sessionId = null;

  ws.on('message', async (data) => {
    const msg = JSON.parse(data);

    // 浏览器端注册
    if (msg.type === 'register') {
      sessionId = msg.sessionId || `session-${Date.now()}`;
      browserClients.set(sessionId, ws);
      ws.send(JSON.stringify({ type: 'registered', sessionId }));
      console.log(`[WS] 注册会话: ${sessionId}`);
      return;
    }

    // 浏览器端返回工具执行结果
    if (msg.type === 'tool_result') {
      const pending = pendingRequests.get(msg.id);
      if (pending) {
        clearTimeout(pending.timer);
        pendingRequests.delete(msg.id);
        if (msg.error) {
          pending.reject(new Error(msg.error));
        } else {
          pending.resolve(msg.result);
        }
      }
      return;
    }

    // 用户聊天消息（从前端发来）
    if (msg.type === 'chat') {
      try {
        ws.send(JSON.stringify({ type: 'status', message: '正在思考...' }));
        const reply = await masterAgent(msg.content, sessionId);
        ws.send(JSON.stringify({ type: 'chat_reply', content: reply }));
      } catch (err) {
        console.error('[WS] 处理消息失败:', err);
        ws.send(JSON.stringify({ type: 'chat_reply', content: `出错了: ${err.message}` }));
      }
      return;
    }
  });

  ws.on('close', () => {
    if (sessionId) {
      browserClients.delete(sessionId);
      console.log(`[WS] 断开会话: ${sessionId}`);
    }
  });
});

// ==========================================
// 启动
// ==========================================
const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
  console.log(`\n🚀 Agent Excel Demo 已启动!`);
  console.log(`   打开浏览器访问: http://localhost:${PORT}`);
  console.log(`\n📋 配置:`);
  console.log(`   LLM API: ${LLM_CONFIG.baseURL}`);
  console.log(`   Model:   ${LLM_CONFIG.model}`);
  console.log(`\n💡 环境变量:`);
  console.log(`   LLM_BASE_URL  - LLM API 地址`);
  console.log(`   LLM_API_KEY   - API Key`);
  console.log(`   LLM_MODEL     - 模型名称`);
  console.log(`   PORT          - 服务端口 (默认 3000)\n`);
});
