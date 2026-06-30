import html
import json
import os
import re
import threading
import time
import uuid
import zipfile
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
from pathlib import Path
from urllib import request, error
from xml.sax.saxutils import escape
import cgi


BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "outputs"
UPLOAD_DIR = BASE_DIR / "uploads"
HOST = "127.0.0.1"
PORT = int(os.environ.get("PAPER_PPT_PORT", "8765"))

JOBS = {}
JOB_LOCK = threading.Lock()


HTML_PAGE = r"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>AI 论文阅读生成 PPT</title>
  <style>
    :root {
      color-scheme: light;
      --ink: #172033;
      --muted: #647085;
      --line: #d8dee8;
      --paper: #f7f3ec;
      --panel: #ffffff;
      --accent: #2454a6;
      --accent-2: #0f8a7c;
      --warn: #b56b13;
      --shadow: 0 16px 38px rgba(31, 42, 68, .12);
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Microsoft YaHei", "Segoe UI", Arial, sans-serif;
      color: var(--ink);
      background: linear-gradient(180deg, #f9fbff 0%, var(--paper) 100%);
      min-height: 100vh;
    }
    header {
      padding: 28px 36px 16px;
      border-bottom: 1px solid rgba(23, 32, 51, .08);
      background: rgba(255,255,255,.72);
      backdrop-filter: blur(10px);
      position: sticky;
      top: 0;
      z-index: 5;
    }
    h1 { margin: 0 0 8px; font-size: 28px; letter-spacing: 0; }
    .sub { color: var(--muted); font-size: 14px; }
    main {
      display: grid;
      grid-template-columns: minmax(420px, 1fr) 420px;
      gap: 22px;
      padding: 24px 36px 40px;
      max-width: 1360px;
      margin: 0 auto;
    }
    section {
      background: var(--panel);
      border: 1px solid var(--line);
      border-radius: 8px;
      box-shadow: var(--shadow);
      padding: 22px;
    }
    h2 { margin: 0 0 16px; font-size: 18px; }
    label { display: block; font-weight: 700; margin: 14px 0 8px; font-size: 13px; }
    textarea, input, select {
      width: 100%;
      border: 1px solid var(--line);
      border-radius: 6px;
      padding: 11px 12px;
      color: var(--ink);
      font: inherit;
      background: #fff;
      outline: none;
    }
    textarea { min-height: 270px; resize: vertical; line-height: 1.55; }
    input:focus, textarea:focus, select:focus { border-color: var(--accent); box-shadow: 0 0 0 3px rgba(36,84,166,.12); }
    .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }
    .hint { color: var(--muted); font-size: 12px; margin-top: 7px; line-height: 1.45; }
    .actions { display: flex; gap: 12px; align-items: center; margin-top: 18px; }
    button {
      border: 0;
      border-radius: 6px;
      padding: 12px 18px;
      background: var(--accent);
      color: #fff;
      cursor: pointer;
      font-weight: 800;
      min-width: 132px;
    }
    button:disabled { opacity: .55; cursor: not-allowed; }
    .ghost { background: #edf2fa; color: var(--accent); }
    .meter {
      height: 12px;
      background: #e7ebf2;
      border-radius: 999px;
      overflow: hidden;
      margin: 8px 0 18px;
    }
    .bar {
      height: 100%;
      width: 0;
      background: linear-gradient(90deg, var(--accent), var(--accent-2));
      transition: width .35s ease;
    }
    .status {
      min-height: 48px;
      border-left: 4px solid var(--accent);
      padding: 10px 12px;
      background: #f4f7fb;
      color: var(--ink);
      line-height: 1.5;
      border-radius: 0 6px 6px 0;
      margin-bottom: 16px;
    }
    .steps { display: grid; gap: 9px; margin-top: 16px; }
    .step {
      display: flex;
      align-items: center;
      gap: 10px;
      color: var(--muted);
      font-size: 13px;
    }
    .dot {
      width: 10px;
      height: 10px;
      border-radius: 50%;
      border: 2px solid #aeb8c8;
      flex: 0 0 auto;
    }
    .step.done { color: var(--ink); }
    .step.done .dot { background: var(--accent-2); border-color: var(--accent-2); }
    .result {
      display: none;
      margin-top: 18px;
      padding: 14px;
      border: 1px solid #b8d8d2;
      background: #effaf7;
      border-radius: 8px;
    }
    .result a { color: var(--accent); font-weight: 800; }
    .api-box {
      margin-top: 18px;
      padding: 14px;
      border: 1px solid #d9c89e;
      background: #fff9eb;
      border-radius: 8px;
      color: #4b3b17;
      font-size: 13px;
      line-height: 1.6;
    }
    code { background: rgba(23,32,51,.08); padding: 2px 5px; border-radius: 4px; }
    @media (max-width: 980px) {
      main { grid-template-columns: 1fr; padding: 18px; }
      header { padding: 22px 18px 14px; }
      .grid { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <header>
    <h1>AI 论文阅读生成 PPT</h1>
    <div class="sub">输入论文内容或上传 TXT/PDF，配置 OpenAI 兼容接口，生成可编辑 PowerPoint 文件。</div>
  </header>
  <main>
    <section>
      <h2>论文与接口</h2>
      <form id="form">
        <label for="paper_text">论文文本</label>
        <textarea id="paper_text" name="paper_text" placeholder="粘贴摘要、正文、实验结果、结论等内容。文本越完整，PPT越具体。"></textarea>
        <label for="paper_file">上传论文文件</label>
        <input id="paper_file" name="paper_file" type="file" accept=".txt,.md,.pdf" />
        <div class="hint">TXT/MD 可直接读取；PDF 会尽力提取文本，若本机没有 PDF 解析库，请复制论文文本到上方输入框。</div>
        <div class="grid">
          <div>
            <label for="api_base">AI 接口地址</label>
            <input id="api_base" name="api_base" value="https://api.deepseek.com" placeholder="https://api.deepseek.com" />
          </div>
          <div>
            <label for="model">模型名称</label>
            <select id="model" name="model">
              <option value="deepseek-v4-flash">deepseek-v4-flash</option>
              <option value="deepseek-v4-pro">deepseek-v4-pro</option>
              <option value="deepseek-chat">deepseek-chat（旧兼容名）</option>
              <option value="deepseek-reasoner">deepseek-reasoner（旧兼容名）</option>
              <option value="gpt-4o-mini">gpt-4o-mini</option>
            </select>
          </div>
        </div>
        <label for="api_key">API Key</label>
        <input id="api_key" name="api_key" type="password" placeholder="sk-..." />
        <div class="hint">还没有 Key？打开 <a href="https://platform.deepseek.com/api_keys" target="_blank" rel="noopener">DeepSeek API Keys</a> 创建或复制。</div>
        <div class="grid">
          <div>
            <label for="slides">页数</label>
            <select id="slides" name="slides">
              <option value="0">不限（AI 按内容自动决定）</option>
              <option value="8">8 页</option>
              <option value="10">10 页</option>
              <option value="12">12 页</option>
              <option value="15">15 页</option>
              <option value="20">20 页</option>
            </select>
          </div>
          <div>
            <label for="language">语言</label>
            <select id="language" name="language">
              <option value="中文">中文</option>
              <option value="English">English</option>
            </select>
          </div>
        </div>
        <div class="actions">
          <button id="submit" type="submit">生成 PPT</button>
          <button class="ghost" type="button" id="fillDemo">填入示例</button>
        </div>
      </form>
    </section>
    <section>
      <h2>生成进度</h2>
      <div class="meter"><div class="bar" id="bar"></div></div>
      <div class="status" id="status">等待提交论文。</div>
      <div class="steps" id="steps">
        <div class="step" data-p="10"><span class="dot"></span>接收论文与接口参数</div>
        <div class="step" data-p="25"><span class="dot"></span>提取论文核心信息</div>
        <div class="step" data-p="45"><span class="dot"></span>调用 AI 生成大纲</div>
        <div class="step" data-p="70"><span class="dot"></span>设计幻灯片内容</div>
        <div class="step" data-p="90"><span class="dot"></span>写入 PowerPoint 文件</div>
        <div class="step" data-p="100"><span class="dot"></span>完成并生成下载链接</div>
      </div>
      <div class="result" id="result"></div>
      <div class="api-box">
        <strong>本地接入口链接</strong><br />
        DeepSeek：<code>https://api.deepseek.com</code>，模型建议 <code>deepseek-v4-flash</code><br />
        获取 Key：<a href="https://platform.deepseek.com/api_keys" target="_blank" rel="noopener">https://platform.deepseek.com/api_keys</a><br />
        页面：<code id="pageUrl"></code><br />
        生成接口：<code id="apiUrl"></code><br />
        进度查询：<code>/api/jobs/{job_id}</code>
      </div>
    </section>
  </main>
  <script>
    const form = document.getElementById('form');
    const submit = document.getElementById('submit');
    const statusBox = document.getElementById('status');
    const bar = document.getElementById('bar');
    const result = document.getElementById('result');
    const steps = [...document.querySelectorAll('.step')];
    document.getElementById('pageUrl').textContent = location.href;
    document.getElementById('apiUrl').textContent = location.origin + '/api/generate';

    function setProgress(progress, message) {
      bar.style.width = `${progress || 0}%`;
      statusBox.textContent = message || '';
      steps.forEach(s => s.classList.toggle('done', Number(s.dataset.p) <= progress));
    }

    async function poll(jobId) {
      const res = await fetch(`/api/jobs/${jobId}`);
      const job = await res.json();
      setProgress(job.progress, job.message);
      if (job.status === 'done') {
        submit.disabled = false;
        result.style.display = 'block';
        result.innerHTML = `已完成：<a href="${job.download_url}">下载 PPTX</a><div class="hint">文件也保存在项目 outputs 目录中。</div>`;
        return;
      }
      if (job.status === 'error') {
        submit.disabled = false;
        result.style.display = 'block';
        result.innerHTML = `<strong>生成失败：</strong>${job.error || job.message}`;
        return;
      }
      setTimeout(() => poll(jobId), 900);
    }

    form.addEventListener('submit', async (event) => {
      event.preventDefault();
      submit.disabled = true;
      result.style.display = 'none';
      setProgress(5, '正在提交任务...');
      const data = new FormData(form);
      const res = await fetch('/api/generate', { method: 'POST', body: data });
      const payload = await res.json();
      if (!res.ok) {
        submit.disabled = false;
        setProgress(0, payload.error || '提交失败');
        return;
      }
      poll(payload.job_id);
    });

    document.getElementById('fillDemo').addEventListener('click', () => {
      document.getElementById('paper_text').value = `Title: Retrieval-Augmented Generation for Scientific Paper Reading\nAbstract: This paper proposes an AI-assisted workflow that extracts claims, methods, experiments, limitations, and future work from academic papers, then generates editable presentation slides.\nMethod: The system combines document parsing, section-aware summarization, evidence ranking, and slide planning. It keeps citations and highlights experimental results.\nExperiments: On a benchmark of research articles, the approach improves summary coverage and reduces slide preparation time for graduate students.\nConclusion: RAG-based paper reading can transform dense papers into structured presentations while preserving evidence and uncertainty.`;
    });
  </script>
</body>
</html>
"""


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def set_job(job_id, **updates):
    with JOB_LOCK:
        JOBS.setdefault(job_id, {}).update(updates)


def read_upload_text(field):
    if field is None or not getattr(field, "filename", ""):
        return ""
    filename = Path(field.filename).name
    raw = field.file.read()
    UPLOAD_DIR.mkdir(exist_ok=True)
    (UPLOAD_DIR / f"{uuid.uuid4().hex}_{filename}").write_bytes(raw)
    suffix = Path(filename).suffix.lower()
    if suffix in {".txt", ".md"}:
        return raw.decode("utf-8", errors="ignore")
    if suffix == ".pdf":
        return extract_pdf_text(raw)
    return raw.decode("utf-8", errors="ignore")


def extract_pdf_text(raw):
    try:
        import pypdf
        reader = pypdf.PdfReader(BytesIO(raw))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception:
        return ""


def compact_text(text, limit=12000):
    text = re.sub(r"\s+", " ", text or "").strip()
    return text[:limit]


def local_outline(paper_text, slide_count, language):
    title = guess_title(paper_text)
    sentences = split_sentences(paper_text)
    keywords = top_keywords(paper_text)
    data_points = extract_data_points(paper_text)
    is_english = is_mostly_english(paper_text)
    sections = [
        ("研究背景", "这篇论文关注的问题、应用场景与研究动机。", "flow"),
        ("核心问题", "现有方法的限制，以及本文试图解决的关键缺口。", "matrix"),
        ("方法框架", "模型、流程或系统架构的主要组成部分。", "flow"),
        ("关键创新", "论文相对已有工作的改进点和技术贡献。", "matrix"),
        ("实验设置", "数据集、评价指标、对比方法与实验协议。", "timeline"),
        ("主要结果", "最重要的实验发现、性能变化和可解释结论。", "bars"),
        ("结果分析", "对实验数据的深入分析、消融实验和可视化结果解读。", "bars"),
        ("局限与风险", "适用边界、潜在失败模式和作者承认的不足。", "matrix"),
        ("汇报总结", "适合在组会或课堂中强调的结论与后续问题。", "flow"),
        ("可复现清单", "代码、数据、参数、消融实验和阅读时需要核查的证据。", "timeline"),
        ("延伸思考", "可以继续做的研究方向或工程落地机会。", "flow"),
        ("讨论问题", "给听众的讨论点和需要进一步确认的假设。", "matrix"),
        ("参考信息", "论文关键词、术语和可追溯信息。", "bars"),
    ]
    if slide_count <= 0:
        slide_count = max(6, min(len(sections), len(sentences) + 2))
    slides = []
    for i in range(slide_count):
        heading, fallback, visual_type = sections[i % len(sections)]
        evidence = sentences[i % len(sentences)] if sentences else fallback
        bullets = [
            evidence[:150],
            f"关键词：{', '.join(keywords[:5])}" if keywords else "提炼论文中的定义、假设和实验依据。",
            fallback,
        ]
        if data_points and i < len(data_points):
            bullets.insert(1, f"原始数据：{data_points[i][:120]}")
        slide = {
            "title": heading,
            "subtitle": fallback,
            "bullets": bullets,
            "zh_summary": fallback if not is_english else f"中文解读：{fallback}",
            "visual_type": visual_type,
            "visual_items": keywords[:4] or ["问题", "方法", "实验", "结论"],
        }
        if is_english:
            slide["english_terms"] = [evidence[:90], "Chinese translation: 自动提炼论文含义，生成中文讲解。"]
        slides.append(slide)
    slides[0] = {
        "title": title or "论文阅读汇报",
        "subtitle": "AI 自动生成的论文汇报 PPT",
        "bullets": ["研究主题概览", f"生成时间：{now()}", f"语言：{language}", "内置图示：流程图、矩阵图、柱状图、路线图"],
        "zh_summary": "这份 PPT 会把论文的英文信息转成中文讲解，并用图示组织方法、贡献、实验和局限。",
        "visual_type": "flow",
        "visual_items": ["论文输入", "结构化阅读", "图示化表达", "PPT 汇报"],
    }
    return {"deck_title": title or "论文阅读汇报", "slides": slides}


def is_mostly_english(text):
    latin = len(re.findall(r"[A-Za-z]", text or ""))
    cjk = len(re.findall(r"[\u4e00-\u9fff]", text or ""))
    return latin > max(120, cjk * 2)


def guess_title(text):
    lines = [line.strip() for line in (text or "").splitlines() if line.strip()]
    for line in lines[:8]:
        line = re.sub(r"^(title|题目)\s*[:：]\s*", "", line, flags=re.I)
        if 6 <= len(line) <= 90 and not line.lower().startswith(("abstract", "摘要")):
            return line
    return ""


def split_sentences(text):
    parts = re.split(r"(?<=[。！？.!?])\s+", compact_text(text, 8000))
    return [p.strip() for p in parts if len(p.strip()) > 25][:40] or [compact_text(text, 300)]


def top_keywords(text):
    words = re.findall(r"[A-Za-z][A-Za-z0-9\-]{3,}|[\u4e00-\u9fff]{2,}", text or "")
    stop = {"this", "that", "with", "from", "paper", "method", "result", "using", "based", "研究", "论文", "方法", "结果"}
    counts = {}
    for word in words:
        w = word.lower()
        if w in stop:
            continue
        counts[w] = counts.get(w, 0) + 1
    return [w for w, _ in sorted(counts.items(), key=lambda item: item[1], reverse=True)[:10]]


def extract_data_points(text):
    lines = (text or "").splitlines()
    data_lines = []
    for line in lines:
        line = line.strip()
        if re.search(r"\d+[\.\d]*%|\d+\.\d+\s*[±±]", line) or re.search(r"(准确率|精度|召回|F1|AUC|mAP|BLEU|ROUGE|Perplexity|loss|acc|accuracy|precision|recall|score)", line, re.I):
            data_lines.append(line[:150])
    return data_lines[:20]


def call_ai_outline(api_base, api_key, model, paper_text, slide_count, language):
    if not api_base or not api_key:
        return None
    url = api_base.strip()
    if not url.endswith("/chat/completions"):
        url = url.rstrip("/") + "/chat/completions"
    source_language = "英文" if is_mostly_english(paper_text) else "中文或混合语言"
    slide_instruction = f"生成 {slide_count} 页 {language} PPT 大纲" if slide_count > 0 else f"根据论文内容决定页数（不限数量），生成 {language} PPT 大纲，以讲清楚论文为准"
    prompt = f"""
请阅读论文内容，{slide_instruction}，质量要像论文组会汇报。
如果原文是英文，每页必须提供中文翻译/中文解读，不能只复制英文原文。
每页必须规划一个图示，图示类型从 flow、bars、matrix、timeline 中选择。
尽量使用论文中的原始实验数据（如准确率、指标数值、对比数据等）填充到内容中。
对于论文中重要的图表、表格数据，必须在对应幻灯片中描述和分析。
只返回 JSON，不要 Markdown。格式：
{{
  "deck_title": "标题",
  "slides": [
    {{
      "title": "中文页标题，必须像结论",
      "subtitle": "可选中文副标题",
      "bullets": ["中文要点1（含原始数据）", "中文要点2（含原始数据）", "中文要点3（含原始数据）"],
      "zh_summary": "对英文原文的中文翻译或中文讲解，2句以内",
      "english_terms": ["英文关键术语 -> 中文解释", "英文原句短摘 -> 中文含义"],
      "visual_type": "flow|bars|matrix|timeline",
      "visual_items": ["图中节点或指标1", "图中节点或指标2", "图中节点或指标3", "图中节点或指标4"]
    }}
  ]
}}
每页标题要像汇报结论，不要只是章节名；要点必须来自论文证据，不能编造数据。
如果论文中有实验数据，务必在 bullets 中包含具体数值和对比结果。
若没有明确数值，bars 图只表达相对强弱或证据权重，不要虚构真实指标。
原文语言判断：{source_language}

论文内容：
{compact_text(paper_text)}
"""
    payload = {
        "model": model or "deepseek-v4-flash",
        "messages": [
            {"role": "system", "content": "You are an expert research paper analyst and PowerPoint outline writer."},
            {"role": "user", "content": prompt},
        ],
        "temperature": 0.35,
        "response_format": {"type": "json_object"},
    }
    req = request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}",
        },
        method="POST",
    )
    with request.urlopen(req, timeout=90) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    content = data["choices"][0]["message"]["content"]
    return json.loads(content)


def safe_filename(name):
    name = re.sub(r"[^\w\u4e00-\u9fff\-]+", "_", name or "论文阅读汇报").strip("_")
    return (name[:36] or "论文阅读汇报") + ".pptx"


def build_pptx(outline, output_path):
    slides = outline.get("slides", [])
    if not slides:
        slides = local_outline("论文阅读汇报", 8, "中文")["slides"]
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types(len(slides)))
        z.writestr("_rels/.rels", root_rels())
        z.writestr("ppt/presentation.xml", presentation_xml(len(slides)))
        z.writestr("ppt/_rels/presentation.xml.rels", presentation_rels(len(slides)))
        z.writestr("ppt/theme/theme1.xml", theme_xml())
        z.writestr("ppt/slideMasters/slideMaster1.xml", slide_master_xml())
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", slide_master_rels())
        z.writestr("ppt/slideLayouts/slideLayout1.xml", slide_layout_xml())
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slide_layout_rels())
        z.writestr("docProps/core.xml", core_xml())
        z.writestr("docProps/app.xml", app_xml(len(slides)))
        for idx, slide in enumerate(slides, start=1):
            z.writestr(f"ppt/slides/slide{idx}.xml", slide_xml(slide, idx, len(slides)))
            z.writestr(f"ppt/slides/_rels/slide{idx}.xml.rels", slide_rels())


def tx_body(lines, font_size=2200, color="172033"):
    paras = []
    for line in lines:
        line = str(line or "").strip()
        if not line:
            continue
        paras.append(
            f'<a:p><a:r><a:rPr lang="zh-CN" sz="{font_size}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:rPr>'
            f'<a:t>{escape(line[:240])}</a:t></a:r></a:p>'
        )
    return "".join(paras) or '<a:p><a:r><a:t></a:t></a:r></a:p>'


def shape_text(shape_id, name, x, y, cx, cy, lines, font_size=2200, color="172033", fill=None):
    fill_xml = '<a:noFill/>' if not fill else f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
    return f"""
    <p:sp>
      <p:nvSpPr><p:cNvPr id="{shape_id}" name="{escape(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{fill_xml}<a:ln><a:noFill/></a:ln></p:spPr>
      <p:txBody><a:bodyPr wrap="square" lIns="140000" tIns="80000" rIns="140000" bIns="80000"/><a:lstStyle/>{tx_body(lines, font_size, color)}</p:txBody>
    </p:sp>"""


def rect_shape(shape_id, name, x, y, cx, cy, fill="FFFFFF", line="D8DEE8"):
    line_xml = '<a:noFill/>' if not line else f'<a:solidFill><a:srgbClr val="{line}"/></a:solidFill>'
    return f"""
    <p:sp>
      <p:nvSpPr><p:cNvPr id="{shape_id}" name="{escape(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
      <p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill><a:ln w="12000">{line_xml}</a:ln></p:spPr>
    </p:sp>"""


def clean_items(items, fallback):
    if isinstance(items, str):
        items = re.split(r"[;；,，\n]+", items)
    values = []
    for item in items or []:
        text = str(item or "").strip()
        if text:
            values.append(text[:34])
    return (values or fallback)[:4]


def clean_lines(value, limit=3):
    if isinstance(value, str):
        parts = re.split(r"[;；\n]+", value)
    else:
        parts = value or []
    lines = []
    for item in parts:
        text = str(item or "").strip()
        if text:
            lines.append(text[:95])
    return lines[:limit]


def visual_shapes(slide, idx):
    visual_type = (slide.get("visual_type") or "").lower()
    if visual_type not in {"flow", "bars", "matrix", "timeline"}:
        visual_type = ["flow", "matrix", "bars", "timeline"][idx % 4]
    items = clean_items(slide.get("visual_items"), ["问题", "方法", "实验", "结论"])
    base_id = 20
    if visual_type == "bars":
        return bars_visual(base_id, items)
    if visual_type == "matrix":
        return matrix_visual(base_id, items)
    if visual_type == "timeline":
        return timeline_visual(base_id, items)
    return flow_visual(base_id, items)


def flow_visual(base_id, items):
    shapes = [shape_text(base_id, "visual-title", 650000, 2050000, 3900000, 360000, ["论文逻辑流程图"], 1600, "2454A6")]
    y = 2600000
    for i, item in enumerate(items):
        x = 650000 + i * 2450000
        shapes.append(shape_text(base_id + 1 + i, f"flow-node-{i+1}", x, y, 1900000, 780000, [f"{i+1}. {item}"], 1650, "FFFFFF", "2454A6" if i % 2 == 0 else "0F8A7C"))
        if i < len(items) - 1:
            shapes.append(rect_shape(base_id + 10 + i, f"flow-link-{i+1}", x + 1950000, y + 345000, 420000, 45000, "B56B13", None))
    shapes.append(shape_text(base_id + 18, "visual-note", 780000, 3800000, 9000000, 720000, ["用图把论文从问题、方法、证据到结论的阅读路径串起来。"], 1550, "647085", "FFFFFF"))
    return shapes


def bars_visual(base_id, items):
    shapes = [shape_text(base_id, "visual-title", 650000, 2050000, 3900000, 360000, ["证据权重图"], 1600, "2454A6")]
    widths = [3300000, 2750000, 2200000, 1650000]
    for i, item in enumerate(items):
        y = 2600000 + i * 560000
        shapes.append(shape_text(base_id + 1 + i, f"bar-label-{i+1}", 700000, y, 1900000, 380000, [item], 1400, "172033"))
        shapes.append(rect_shape(base_id + 8 + i, f"bar-track-{i+1}", 2700000, y + 90000, 3600000, 210000, "E7EBF2", None))
        shapes.append(rect_shape(base_id + 14 + i, f"bar-value-{i+1}", 2700000, y + 90000, widths[i], 210000, "0F8A7C" if i % 2 else "2454A6", None))
    shapes.append(shape_text(base_id + 22, "visual-note", 700000, 5000000, 5600000, 520000, ["无明确数值时表示相对证据强弱，不虚构实验指标。"], 1350, "647085", "FFFFFF"))
    return shapes


def matrix_visual(base_id, items):
    row_labels = clean_items(items[:2], ["现有问题", "本文方案"])
    col_labels = clean_items(items[2:], ["技术贡献", "实验证据"])
    while len(row_labels) < 2:
        row_labels.append("对比维度")
    while len(col_labels) < 2:
        col_labels.append("证据维度")
    shapes = [shape_text(base_id, "visual-title", 650000, 2050000, 3900000, 360000, ["贡献对照矩阵"], 1600, "2454A6")]
    x0, y0 = 850000, 2650000
    shapes.append(shape_text(base_id + 1, "matrix-col-1", x0 + 1900000, y0, 1900000, 420000, [col_labels[0]], 1350, "FFFFFF", "2454A6"))
    shapes.append(shape_text(base_id + 2, "matrix-col-2", x0 + 3850000, y0, 1900000, 420000, [col_labels[1]], 1350, "FFFFFF", "0F8A7C"))
    for r in range(2):
        y = y0 + 470000 + r * 780000
        shapes.append(shape_text(base_id + 3 + r, f"matrix-row-{r+1}", x0, y, 1700000, 700000, [row_labels[r]], 1350, "172033", "F7F3EC"))
        shapes.append(shape_text(base_id + 5 + r * 2, f"matrix-cell-{r+1}-1", x0 + 1900000, y, 1900000, 700000, ["✓", "论文证据"], 1350, "172033", "FFFFFF"))
        shapes.append(shape_text(base_id + 6 + r * 2, f"matrix-cell-{r+1}-2", x0 + 3850000, y, 1900000, 700000, ["→", "汇报重点"], 1350, "172033", "FFFFFF"))
    return shapes


def timeline_visual(base_id, items):
    shapes = [shape_text(base_id, "visual-title", 650000, 2050000, 3900000, 360000, ["阅读路线图"], 1600, "2454A6")]
    y = 3300000
    shapes.append(rect_shape(base_id + 1, "timeline-line", 850000, y + 240000, 7600000, 52000, "2454A6", None))
    for i, item in enumerate(items):
        x = 850000 + i * 2500000
        shapes.append(shape_text(base_id + 2 + i, f"timeline-node-{i+1}", x, y, 780000, 520000, [str(i + 1)], 1800, "FFFFFF", "B56B13" if i % 2 else "0F8A7C"))
        shapes.append(shape_text(base_id + 8 + i, f"timeline-label-{i+1}", x - 220000, y + 640000, 1300000, 620000, [item], 1300, "172033", "FFFFFF"))
    return shapes


def slide_xml(slide, idx, total):
    title = slide.get("title", f"Slide {idx}")
    subtitle = slide.get("subtitle", "")
    bullets = slide.get("bullets", [])
    zh_summary = slide.get("zh_summary", "")
    english_terms = clean_lines(slide.get("english_terms", []), 3)
    bg = "F7F3EC" if idx % 2 else "F9FBFF"
    accent = "2454A6" if idx % 3 else "0F8A7C"
    shapes = [
        f'<p:bg><p:bgPr><a:solidFill><a:srgbClr val="{bg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>',
        shape_text(2, "kicker-label", 520000, 320000, 3100000, 330000, [f"论文阅读图解 / {idx:02d}"], 1300, accent),
        shape_text(3, "title", 520000, 760000, 11200000, 920000, [title], 3000 if idx > 1 else 3800, "172033"),
    ]
    if subtitle:
        shapes.append(shape_text(4, "subtitle", 620000, 1600000, 10400000, 430000, [subtitle], 1650, "647085"))
    bullet_lines = [f"• {b}" for b in bullets[:5]]
    shapes.extend(visual_shapes(slide, idx))
    shapes.append(shape_text(5, "content", 7000000, 2200000, 4300000, 2250000, bullet_lines, 1550, "172033", "FFFFFF"))
    if zh_summary:
        shapes.append(shape_text(6, "zh-summary", 7000000, 4650000, 4300000, 780000, [f"中文解读：{zh_summary}"], 1450, "172033", "EFFAF7"))
    if english_terms:
        term_lines = [str(t) for t in english_terms[:3]]
        shapes.append(shape_text(7, "translation-terms", 7000000, 5520000, 4300000, 640000, term_lines, 1220, "4B3B17", "FFF9EB"))
    shapes.append(shape_text(8, "footer", 540000, 6450000, 10800000, 260000, [f"AI 论文阅读生成 PPT  |  图示化阅读  |  {idx}/{total}"], 950, "647085"))
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    {"".join(shapes[:1])}
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      {"".join(shapes[1:])}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>"""


def content_types(n):
    overrides = "".join(f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' for i in range(1, n + 1))
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
{overrides}
</Types>"""


def root_rels():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""


def presentation_xml(n):
    slide_ids = "".join(f'<p:sldId id="{255+i}" r:id="rId{i}"/>' for i in range(1, n + 1))
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId{n+1}"/></p:sldMasterIdLst>
<p:sldIdLst>{slide_ids}</p:sldIdLst>
<p:sldSz cx="12192000" cy="6858000" type="wide"/>
<p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>"""


def presentation_rels(n):
    rels = "".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>' for i in range(1, n + 1))
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{rels}
<Relationship Id="rId{n+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
<Relationship Id="rId{n+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>"""


def slide_rels():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"""


def slide_master_rels():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"""


def slide_layout_rels():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"""


def slide_master_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>"""


def slide_layout_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
<p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>"""


def theme_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="PaperPPT">
<a:themeElements><a:clrScheme name="PaperPPT"><a:dk1><a:srgbClr val="172033"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="2454A6"/></a:dk2><a:lt2><a:srgbClr val="F7F3EC"/></a:lt2><a:accent1><a:srgbClr val="2454A6"/></a:accent1><a:accent2><a:srgbClr val="0F8A7C"/></a:accent2><a:accent3><a:srgbClr val="B56B13"/></a:accent3><a:accent4><a:srgbClr val="647085"/></a:accent4><a:accent5><a:srgbClr val="D8DEE8"/></a:accent5><a:accent6><a:srgbClr val="172033"/></a:accent6><a:hlink><a:srgbClr val="2454A6"/></a:hlink><a:folHlink><a:srgbClr val="0F8A7C"/></a:folHlink></a:clrScheme><a:fontScheme name="PaperPPT"><a:majorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:majorFont><a:minorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:minorFont></a:fontScheme><a:fmtScheme name="PaperPPT"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements>
</a:theme>"""


def core_xml():
    created = datetime.utcnow().isoformat(timespec="seconds") + "Z"
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>AI 论文阅读生成 PPT</dc:title><dc:creator>Paper PPT AI</dc:creator><cp:lastModifiedBy>Paper PPT AI</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">{created}</dcterms:modified>
</cp:coreProperties>"""


def app_xml(slide_count):
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>Paper PPT AI</Application><PresentationFormat>宽屏</PresentationFormat><Slides>{slide_count}</Slides><Company></Company>
</Properties>"""


def worker(job_id, form, file_field):
    try:
        set_job(job_id, status="running", progress=10, message="接收论文与接口参数...")
        paper_text = (form.get("paper_text") or "") + "\n" + read_upload_text(file_field)
        paper_text = paper_text.strip()
        if len(paper_text) < 40:
            raise ValueError("论文内容太少。请粘贴论文文本，或上传可读取的 TXT/PDF。")
        slide_count = int(form.get("slides") or 0)
        language = form.get("language") or "中文"
        set_job(job_id, progress=25, message="正在提取论文核心信息...")
        time.sleep(0.4)
        set_job(job_id, progress=45, message="正在调用 AI 生成大纲；接口为空时使用本地摘要逻辑...")
        try:
            outline = call_ai_outline(form.get("api_base"), form.get("api_key"), form.get("model"), paper_text, slide_count, language)
        except (error.URLError, TimeoutError, KeyError, json.JSONDecodeError, ValueError) as exc:
            outline = None
            set_job(job_id, message=f"AI 接口未成功返回，已切换本地摘要逻辑：{exc}")
            time.sleep(0.8)
        if not outline:
            outline = local_outline(paper_text, slide_count, language)
        set_job(job_id, progress=70, message="正在设计幻灯片内容...")
        time.sleep(0.4)
        OUTPUT_DIR.mkdir(exist_ok=True)
        filename = safe_filename(outline.get("deck_title"))
        output_path = OUTPUT_DIR / f"{job_id}_{filename}"
        set_job(job_id, progress=90, message="正在写入 PowerPoint 文件...")
        build_pptx(outline, output_path)
        set_job(
            job_id,
            status="done",
            progress=100,
            message="完成。",
            output=str(output_path),
            download_url=f"/api/download/{job_id}",
        )
    except Exception as exc:
        set_job(job_id, status="error", progress=100, message="生成失败。", error=str(exc))


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        print(f"[{now()}] {self.address_string()} {fmt % args}")

    def send_json(self, payload, status=HTTPStatus.OK):
        raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def do_GET(self):
        if self.path == "/" or self.path.startswith("/?"):
            raw = HTML_PAGE.encode("utf-8")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
            return
        if self.path.startswith("/api/jobs/"):
            job_id = self.path.rsplit("/", 1)[-1]
            with JOB_LOCK:
                job = JOBS.get(job_id)
            if not job:
                self.send_json({"error": "job not found"}, HTTPStatus.NOT_FOUND)
                return
            self.send_json(job)
            return
        if self.path.startswith("/api/download/"):
            job_id = self.path.rsplit("/", 1)[-1]
            with JOB_LOCK:
                job = JOBS.get(job_id)
            if not job or not job.get("output"):
                self.send_json({"error": "file not found"}, HTTPStatus.NOT_FOUND)
                return
            path = Path(job["output"])
            raw = path.read_bytes()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation")
            self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{request.pathname2url(path.name)}")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
            return
        self.send_json({"error": "not found"}, HTTPStatus.NOT_FOUND)

    def do_POST(self):
        if self.path != "/api/generate":
            self.send_json({"error": "not found"}, HTTPStatus.NOT_FOUND)
            return
        ctype, pdict = cgi.parse_header(self.headers.get("content-type"))
        if ctype != "multipart/form-data":
            self.send_json({"error": "请使用 multipart/form-data 提交。"}, HTTPStatus.BAD_REQUEST)
            return
        pdict["boundary"] = bytes(pdict["boundary"], "utf-8")
        form_data = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ={"REQUEST_METHOD": "POST"})
        form = {}
        for key in ["paper_text", "api_base", "api_key", "model", "slides", "language"]:
            item = form_data[key] if key in form_data else None
            form[key] = item.value if item is not None and not item.filename else ""
        file_field = form_data["paper_file"] if "paper_file" in form_data else None
        job_id = uuid.uuid4().hex[:12]
        set_job(job_id, status="queued", progress=0, message="任务已创建。")
        threading.Thread(target=worker, args=(job_id, form, file_field), daemon=True).start()
        self.send_json({"job_id": job_id, "status_url": f"/api/jobs/{job_id}"}, HTTPStatus.ACCEPTED)


def main():
    OUTPUT_DIR.mkdir(exist_ok=True)
    UPLOAD_DIR.mkdir(exist_ok=True)
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"AI 论文阅读生成 PPT 已启动： http://{HOST}:{PORT}")
    print(f"生成接口： http://{HOST}:{PORT}/api/generate")
    server.serve_forever()


if __name__ == "__main__":
    main()
