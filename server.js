/**
 * Agent Platform Demo - Server
 *
 * 架构:
 *   用户浏览器 ◀──WebSocket──▶ Server ◀──HTTP──▶ LLM API
 *        │                        │
 *   页面内容 + 操作               Master Agent (路由) + 子 Agent (执行)
 *
 * 页面:
 *   /dashboard - 仪表盘（概览）
 *   /files     - 文件管理
 *   /notes     - 笔记
 *   /excel     - Excel 编辑器
 *   /settings  - 设置
 *
 * Agent 架构:
 *   Master Agent → 分析页面上下文，生成推荐问题，路由到子 Agent
 *     ├── Excel Agent → 操作 Excel（function calling）
 *     ├── Summary Agent → 总结摘要（纯 LLM）
 *     └── Chat Agent → 通用对话（纯 LLM）
 */

require('dotenv').config();
const express = require('express');
const http = require('http');
const { WebSocketServer } = require('ws');
const path = require('path');

const app = express();
const server = http.createServer(app);
const wss = new WebSocketServer({ server });

app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// 截图接收接口
app.post('/api/screenshot', (req, res) => {
  const fs = require('fs');
  const { data } = req.body;
  if (!data) return res.status(400).json({ error: 'no data' });
  const base64 = data.replace(/^data:image\/\w+;base64,/, '');
  fs.writeFileSync('/tmp/screenshot-fs-capture.png', Buffer.from(base64, 'base64'));
  console.log('Screenshot saved:', Buffer.from(base64, 'base64').length, 'bytes');
  res.json({ ok: true });
});

// ==========================================
// LLM 配置
// ==========================================
const LLM_CONFIG = {
  baseURL: process.env.LLM_BASE_URL || 'https://api.openai.com/v1',
  apiKey: process.env.LLM_API_KEY || 'your-api-key',
  model: process.env.LLM_MODEL || 'gpt-4o-mini',
};

async function callLLM(messages, tools = null) {
  const body = { model: LLM_CONFIG.model, messages, temperature: 0.1 };
  if (tools) { body.tools = tools; body.tool_choice = 'auto'; }

  const resp = await fetch(`${LLM_CONFIG.baseURL}/chat/completions`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${LLM_CONFIG.apiKey}` },
    body: JSON.stringify(body),
  });

  if (!resp.ok) throw new Error(`LLM API error: ${resp.status}`);
  return resp.json();
}

// ==========================================
// Excel Tools 定义
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
      description: '设置指定单元格的值或公式。公式以 = 开头，如 =SUM(B1:B10)',
      parameters: {
        type: 'object',
        properties: {
          row: { type: 'number', description: '行号，从1开始' },
          col: { type: 'string', description: '列号字母' },
          value: { type: 'string', description: '要设置的值或公式' },
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
      description: '批量设置一个区域的值。',
      parameters: {
        type: 'object',
        properties: {
          startRow: { type: 'number', description: '起始行号' },
          startCol: { type: 'string', description: '起始列号字母' },
          values: { type: 'array', items: { type: 'array', items: { type: 'string' } }, description: '二维数组' },
        },
        required: ['startRow', 'startCol', 'values'],
      },
    },
  },
  {
    type: 'function',
    function: {
      name: 'get_sheet_info',
      description: '获取当前工作表的基本信息',
      parameters: { type: 'object', properties: {} },
    },
  },
  {
    type: 'function',
    function: {
      name: 'apply_cell_style',
      description: '设置单元格样式（加粗、背景色等）',
      parameters: {
        type: 'object',
        properties: {
          row: { type: 'number', description: '行号' },
          col: { type: 'string', description: '列号字母' },
          bold: { type: 'boolean' },
          bgColor: { type: 'string', description: '背景色，如 #ffff00' },
          fontColor: { type: 'string', description: '字体颜色' },
          fontSize: { type: 'number', description: '字体大小' },
        },
        required: ['row', 'col'],
      },
    },
  },
];

// ==========================================
// 模拟数据
// ==========================================
const MOCK_FILES = [
  { id: 1, name: '销售数据_Q1.xlsx', type: 'excel', size: '24KB', modified: '2026-05-25', owner: '姚磊' },
  { id: 2, name: '项目预算.xlsx', type: 'excel', size: '18KB', modified: '2026-05-24', owner: '姚磊' },
  { id: 3, name: '员工花名册.xlsx', type: 'excel', size: '32KB', modified: '2026-05-23', owner: '姚磊' },
];

const MOCK_NOTES = [
  { id: 1, title: '项目周会纪要 - 第21周', content: '本周完成了用户认证模块的重构，性能提升约 40%。下周计划：1）完成支付集成测试 2）优化数据库查询 3）前端国际化支持。', date: '2026-05-24' },
  { id: 2, title: '技术选型笔记：消息队列', content: '对比了 Kafka、RabbitMQ 和 Pulsar。最终选择 RabbitMQ，原因：轻量、支持多种消息模式、Python/Go 客户端成熟。', date: '2026-05-22' },
  { id: 3, title: '读书笔记：Designing Data-Intensive Applications', content: '第三章笔记：存储引擎分为 log-structured 和 page-oriented 两大类。LSM-Tree 适合写多读少，B-Tree 适合读多写少。', date: '2026-05-20' },
];

const MOCK_DASHBOARD = {
  stats: { files: 3, notes: 3, tasks: 12, completedTasks: 8 },
  recentActivity: [
    { action: '编辑了', target: '销售数据_Q1.xlsx', time: '2小时前' },
    { action: '创建了', target: '项目周会纪要 - 第21周', time: '昨天' },
    { action: '上传了', target: '员工花名册.xlsx', time: '3天前' },
  ],
};

// ==========================================
// 连接管理
// ==========================================
const browserClients = new Map();
const pendingRequests = new Map();

// ==========================================
// Master Agent - 有自己的 LLM 做路由
// ==========================================
async function masterAgent(userMessage, sessionId, pageContext) {
  const browserWs = browserClients.get(sessionId);
  if (!browserWs) return { reply: '错误：浏览器未连接', suggestions: [] };

  const { page, data } = pageContext;

  // 1. Master Agent 判断意图
  console.log(`[Master] 页面=${page}, 消息=${userMessage}`);

  const routePrompt = `你是一个智能助手系统的路由器。根据用户消息和当前页面上下文，判断应该交给哪个子 Agent 处理。

当前页面: ${page}
页面数据摘要: ${data ? JSON.stringify(data).slice(0, 500) : '无'}

可选子 Agent:
- "excel": Excel 数据操作（读写单元格、公式、样式）— 只有当用户明确要操作 Excel 数据时才选
- "general": 通用对话（总结、问答、建议、分析）— 其他所有情况

回复格式（只回复一个词）:
excel 或 general`;

  const routeResponse = await callLLM([
    { role: 'system', content: routePrompt },
    { role: 'user', content: userMessage },
  ]);

  const agentType = routeResponse.choices[0].message.content.trim().toLowerCase();
  console.log(`[Master] 路由到: ${agentType}`);

  // 2. 路由到对应子 Agent
  if (agentType === 'excel' && page === 'excel') {
    return await excelAgent(userMessage, sessionId);
  } else {
    return await generalAgent(userMessage, pageContext);
  }
}

// ==========================================
// 生成页面推荐问题
// ==========================================
async function generateSuggestions(pageContext) {
  const { page, data } = pageContext;

  const prompt = `你是一个智能助手。根据当前页面内容和类型，推荐 3-4 个用户可能会问的问题。

当前页面: ${page}
页面数据: ${data ? JSON.stringify(data).slice(0, 800) : '无数据'}

${page === 'excel' ? `对于 Excel 页面，推荐的问题应该包括：
1. 对数据内容的总结分析
2. 常见的 Excel 操作建议（如求和、排序、条件标注等）
请用中文回复，每行一个问题，不要编号，不要其他内容。` :
`推荐一些实用的、基于页面内容的问题。
请用中文回复，每行一个问题，不要编号，不要其他内容。`}`;

  const response = await callLLM([
    { role: 'system', content: prompt },
    { role: 'user', content: '请推荐问题' },
  ]);

  const text = response.choices[0].message.content.trim();
  return text.split('\n').filter(s => s.trim());
}

// ==========================================
// Excel Agent - 子 Agent
// ==========================================
async function excelAgent(task, sessionId) {
  console.log(`[ExcelAgent] 处理: ${task}`);

  const messages = [
    {
      role: 'system',
      content: `你是一个 Excel 操作助手。用户会用自然语言描述想要对 Excel 表格做的操作。
你需要将用户的意图转换为具体的 Excel 工具调用。
规则：
1. 先了解当前数据（必要时调用 get_sheet_info 或 get_range_values）
2. 然后执行操作
3. 操作完成后，用简洁的中文告诉用户你做了什么
4. 对于求和、平均值等计算，使用公式（如 =SUM(B1:B10)）`,
    },
    { role: 'user', content: task },
  ];

  const MAX_ROUNDS = 10;
  for (let round = 0; round < MAX_ROUNDS; round++) {
    const response = await callLLM(messages, EXCEL_TOOLS);
    const assistantMsg = response.choices[0].message;
    messages.push(assistantMsg);

    if (!assistantMsg.tool_calls || assistantMsg.tool_calls.length === 0) {
      return { reply: assistantMsg.content, suggestions: [] };
    }

    for (const toolCall of assistantMsg.tool_calls) {
      const { name, arguments: argsStr } = toolCall.function;
      const args = JSON.parse(argsStr);
      console.log(`[ExcelAgent] 工具: ${name}`, JSON.stringify(args));

      try {
        const result = await executeOnBrowser(sessionId, name, args);
        messages.push({ role: 'tool', tool_call_id: toolCall.id, content: typeof result === 'string' ? result : JSON.stringify(result) });
      } catch (err) {
        messages.push({ role: 'tool', tool_call_id: toolCall.id, content: `错误: ${err.message}` });
      }
    }
  }

  return { reply: '操作轮次过多，请简化请求。', suggestions: [] };
}

// ==========================================
// General Agent - 子 Agent（通用对话）
// ==========================================
async function generalAgent(userMessage, pageContext) {
  console.log(`[GeneralAgent] 处理: ${userMessage}`);

  const { page, data } = pageContext;

  const response = await callLLM([
    {
      role: 'system',
      content: `你是一个智能助手。当前用户在浏览 "${page}" 页面。
页面数据: ${data ? JSON.stringify(data).slice(0, 1000) : '无数据'}
请根据上下文用中文回答用户的问题。回答要简洁有用。`,
    },
    { role: 'user', content: userMessage },
  ]);

  return { reply: response.choices[0].message.content, suggestions: [] };
}

// ==========================================
// 通过 WebSocket 调用浏览器端 Tool Runtime
// ==========================================
function executeOnBrowser(sessionId, method, params) {
  return new Promise((resolve, reject) => {
    const ws = browserClients.get(sessionId);
    if (!ws) { reject(new Error('浏览器未连接')); return; }

    const requestId = `${Date.now()}-${Math.random().toString(36).slice(2)}`;
    const timeout = setTimeout(() => { pendingRequests.delete(requestId); reject(new Error('超时')); }, 30000);
    pendingRequests.set(requestId, { resolve, reject, timeout });

    ws.send(JSON.stringify({ type: 'tool_call', id: requestId, method, params }));
  });
}

// ==========================================
// WebSocket
// ==========================================
wss.on('connection', (ws) => {
  let sessionId = null;

  ws.on('message', async (raw) => {
    const msg = JSON.parse(raw);

    if (msg.type === 'register') {
      sessionId = msg.sessionId || `s-${Date.now()}`;
      browserClients.set(sessionId, ws);
      ws.send(JSON.stringify({ type: 'registered', sessionId }));
      return;
    }

    if (msg.type === 'tool_result') {
      const p = pendingRequests.get(msg.id);
      if (p) { clearTimeout(p.timeout); pendingRequests.delete(msg.id); msg.error ? p.reject(new Error(msg.error)) : p.resolve(msg.result); }
      return;
    }

    if (msg.type === 'chat') {
      try {
        ws.send(JSON.stringify({ type: 'status', message: '正在思考...' }));
        const result = await masterAgent(msg.content, sessionId, msg.pageContext || {});
        ws.send(JSON.stringify({ type: 'chat_reply', content: result.reply }));
      } catch (err) {
        ws.send(JSON.stringify({ type: 'chat_reply', content: `出错了: ${err.message}` }));
      }
      return;
    }

    if (msg.type === 'get_suggestions') {
      try {
        const suggestions = await generateSuggestions(msg.pageContext || {});
        ws.send(JSON.stringify({ type: 'suggestions', suggestions }));
      } catch (err) {
        ws.send(JSON.stringify({ type: 'suggestions', suggestions: [] }));
      }
      return;
    }
  });

  ws.on('close', () => { if (sessionId) browserClients.delete(sessionId); });
});

// ==========================================
// API routes（供前端获取模拟数据）
// ==========================================
app.get('/api/dashboard', (req, res) => res.json(MOCK_DASHBOARD));
app.get('/api/files', (req, res) => res.json(MOCK_FILES));
app.get('/api/notes', (req, res) => res.json(MOCK_NOTES));

// ==========================================
// 启动
// ==========================================
const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
  console.log(`\n🚀 Agent Platform 已启动!`);
  console.log(`   http://localhost:${PORT}`);
  console.log(`   LLM: ${LLM_CONFIG.baseURL} / ${LLM_CONFIG.model}\n`);
});
