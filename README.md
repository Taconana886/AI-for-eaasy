# AI 论文阅读生成 PPT

一个本地 Web 工具，输入论文文本或上传 PDF，自动生成可编辑的 PowerPoint 汇报幻灯片。支持 AI 生成大纲 + 本地降级双模式，内置 RAG 检索、多智能体流程、视觉质量检测。

## 系统架构

```
用户浏览器 (index.html)
       │  HTTP (GET/POST)
       ▼
server.py  (ThreadingHTTPServer, 端口 8765)
       │
       ├── ai_client.py       → 调用 DeepSeek / OpenAI 兼容 API
       ├── local_parser.py    → 本地提取章节、关键词、指标
       ├── agent_pipeline.py  → 三阶段智能体：策略 → 执行 → 评审
       ├── rag_index.py       → TF-IDF 分块检索，为每页提供上下文
       ├── pdf_extractor.py   → PyMuPDF 解析，提取文本+图片
       ├── ppt_builder.py     → python-pptx 构建 PPTX，4 种模板+模板导入
       ├── visual_qa.py       → 生成后检测文字溢出/重叠/对比度
       ├── outputs/           → 生成的文件
       ├── uploads/           → 上传文件暂存
       ├── checkpoints/       → 任务断点（可恢复）
       ├── history/           → 生成记录 JSON
       └── static/            → 前端页面
```

## 核心工作流程

```
用户提交 ──▶ 接收论文（TXT/PDF 自动提取）
              │
              ├─ (可选) AI 翻译全文
              │
              ├─ RAG 索引 ──▶ 论文分块，TF-IDF 向量化
              │
              ├─ 多智能体流程
              │     ├─ Strategist（策略规划）─ 分析论文，规划每页内容
              │     ├─ Executor（内容执行）─ 用 RAG 检索填充各页证据
              │     └─ Critic（质量评审）  ─ 检查内容完整性
              │
              ├─ 构建 PPTX ──▶ 插入 PDF 图片 + 可视化图表
              │
              ├─ 视觉质检 ──▶ 检测溢出/重叠/对比度，自动修复
              │
              └─ 返回下载 + 保存历史
```

## 模块说明

### server.py
- 基于 Python 标准库 `ThreadingHTTPServer`，无外部 Web 框架依赖
- **检查点机制**：每步（rag/outline/pptx）自动保存，失败可恢复
- 模板上传 API：上传 PPTX 自动提取颜色/字体作为生成模板
- 异步任务 + 前端轮询进度

### ai_client.py
- `call_ai_outline()` — 调用 LLM 生成结构化 PPT 大纲 JSON
- `call_translate()` — 翻译论文为中文，含图文引用智能处理
- 兼容 OpenAI 格式（DeepSeek / OpenAI / vLLM 等）
- API 不可用时自动降级到本地解析

### local_parser.py
- 正则匹配论文章节（Abstract、Introduction、Method 等）
- 提取标题、关键词、实验指标、引用
- `generate_outline()` — 无 AI 时生成基础大纲

### pdf_extractor.py
- 基于 PyMuPDF 替代 pdfminer+pypdf
- 保留排版布局，逐页提取文本
- **图片提取**：自动识别 PDF 内的图表并提取
- 通过标题匹配将图片分配到对应幻灯片

### rag_index.py
- 纯 Python 实现（scikit-learn TF-IDF），无需向量数据库
- 论文自动分块（含重叠），TfidfVectorizer 索引
- `retrieve(query)` 返回最相关段落，为每页幻灯片提供论文证据
- 单文档场景下关键词匹配降级

### agent_pipeline.py
- **Strategist（策略规划）**：分析论文结构，利用 RAG 规划每页主题和图表类型
- **Executor（内容执行）**：用 RAG 检索的结果填充各页要点
- **Critic（质量评审）**：检查空页、缺图表、内容完整性

### ppt_builder.py
- python-pptx 生成 16:9 宽屏 PPTX
- 内置 4 种风格：
  | 风格 | 适用场景 |
  |---|---|
  | 学术期刊风 (academic) | 标准论文汇报 |
  | 科技蓝图风 (tech_blueprint) | 计算机/工程类 |
  | 实验报告风 (experiment) | 实验/数据科学类 |
  | 中文组会风 (chinese_meeting) | 中文组会 |
- **模板导入**：`import_template()` 分析外部 PPTX 颜色/字体，生成匹配风格
- **图片插入**：`_insert_images()` 将 PDF 提取的图表放入 PPT
- 每页内含可视化（流程图/柱状图/矩阵/时间线），原生形状可编辑

### visual_qa.py
- **文字溢出检测**：估算文本行数 vs 形状容量
- **元素重叠检测**：计算各形状边界框交叉面积
- **颜色对比度检测**：WCAG 相对亮度公式，标记低对比度
- 自动缩小超大字号修复溢出问题

## API 文档

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/` 或 `/index.html` | 前端页面 |
| GET | `/static/*` | 静态资源 |
| POST | `/api/generate` | 提交生成任务 |
| POST | `/api/upload-template` | 上传 PPTX 模板 |
| GET | `/api/jobs/{job_id}` | 查询任务进度 |
| GET | `/api/checkpoint/{job_id}` | 查询已保存的断点 |
| GET | `/api/download/{job_id}` | 下载 PPTX |
| GET | `/api/download-translate/{job_id}` | 下载中文翻译 TXT |
| GET | `/api/download-translate-pdf/{job_id}` | 下载中文翻译 PDF |
| GET | `/api/history` | 历史记录 |
| GET | `/api/styles` | 模板风格列表（含导入模板） |

### POST /api/generate

`Content-Type: multipart/form-data`

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `paper_text` | string | 否 | 论文文本 |
| `paper_file` | file | 否 | .txt / .md / .pdf（PDF 自动提取文本+图片） |
| `template_file` | file | 否 | .pptx（用作模板，提取颜色/字体） |
| `api_base` | string | 推荐 | API 地址，如 `https://api.deepseek.com` |
| `api_key` | string | 推荐 | API Key |
| `model` | string | 否 | 模型名 |
| `slides` | int | 否 | 页数 4-12（默认 8） |
| `language` | string | 否 | `中文` 或 `English` |
| `style` | string | 否 | 模板风格 key |
| `translate` | bool | 否 | 同时翻译为中文 |
| `mode` | string | 否 | `ppt`（默认）或 `translate`（仅翻译） |

### POST /api/upload-template

`Content-Type: multipart/form-data`

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| `template` | file | 是 | .pptx 文件，上传后自动分析并加入风格列表 |

## 启动

```bash
cd /mnt/d/AI论文阅读生成PPT
pip install -r requirements.txt
python app.py
```

打开 http://127.0.0.1:8765

### 保持在后台运行

```bash
# nohup
nohup python3 app.py &

# tmux（推荐）
tmux new -s ppt
python3 app.py
# Ctrl+B, D 分离；tmux attach -t ppt 重新连接
```

## 配置

`.env` 文件预设：

```env
HOST=0.0.0.0
PORT=8765
```

## 技术栈

- Python 标准库（HTTP 服务、线程、JSON）— 零 Web 框架依赖
- PyMuPDF — PDF 文本+图片提取
- scikit-learn — TF-IDF 向量化与余弦相似度检索
- python-pptx — PowerPoint 生成
- python-dotenv — 环境变量管理
