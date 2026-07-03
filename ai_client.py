import json
import logging
import re
import time
from urllib import request, error

logger = logging.getLogger(__name__)


def compact_text(text, limit=None):
    text = re.sub(r"\s+", " ", text or "").strip()
    return text if limit is None else text[:limit]


def call_ai_outline(api_base, api_key, model, paper_text, slide_count, language, sections=None):
    if not api_base or not api_key:
        logger.warning(f"AI call skipped: api_base={'set' if api_base else 'empty'}, api_key={'set' if api_key else 'empty'}")
        return None

    url = api_base.strip()
    if not url.endswith("/chat/completions"):
        url = url.rstrip("/") + "/chat/completions"

    source_language = "英文" if _is_mostly_english(paper_text) else "中文或混合语言"

    section_context = ""
    if sections:
        for sec_name, sec_content in sections.items():
            if sec_content and sec_name != "full_text":
                section_context += f"\n### {sec_name}\n{sec_content[:300]}\n"

    prompt = f"""请阅读以下论文内容，生成 {slide_count} 页 {language} PPT 大纲，质量要达到论文组会汇报水准。

要求：
1. 如果原文是英文，每页必须提供中文翻译/中文解读，不能只复制英文原文。
2. 每页必须规划一个图示，图示类型从 flow、bars、matrix、timeline 中选择。
3. 只返回 JSON，不要 Markdown。
4. 每页标题要像汇报结论，不要只是章节名。
5. 要点必须来自论文证据，不能编造数据。
6. 若没有明确数值，bars 图只表达相对强弱或证据权重。

输出 JSON 格式：
{{
  "deck_title": "论文标题",
  "slides": [
    {{
      "title": "中文页标题",
      "subtitle": "可选中文副标题",
      "bullets": ["中文要点1", "中文要点2", "中文要点3"],
      "zh_summary": "对英文原文的中文翻译或中文讲解，2句以内",
      "english_terms": ["英文关键术语 -> 中文解释"],
      "visual_type": "flow|bars|matrix|timeline",
      "visual_items": ["图中节点或指标1", "图中节点或指标2", "图中节点或指标3"]
    }}
  ]
}}

原文语言判断：{source_language}

论文内容：
{compact_text(paper_text, 12000)}

按章节提取的关键内容：
{section_context}
"""
    payload = {
        "model": model or "deepseek-v4-flash",
        "messages": [
            {
                "role": "system",
                "content": "You are an expert research paper analyst and PowerPoint outline writer. Always respond with valid JSON.",
            },
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

    try:
        with request.urlopen(req, timeout=120) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        content = data["choices"][0]["message"]["content"]
        result = json.loads(content)
        logger.info("AI outline generated successfully")
        return result
    except error.HTTPError as e:
        body = e.read().decode("utf-8", errors="ignore")
        logger.warning(f"AI API HTTP error {e.code}: {body[:200]}")
        raise
    except (error.URLError, TimeoutError, KeyError, json.JSONDecodeError, ValueError) as e:
        logger.warning(f"AI API error: {e}")
        raise


def _mark_captions(text):
    lines = text.split("\n")
    marked = []
    for line in lines:
        stripped = line.strip()
        if re.search(
            r"^\s*(?:Fig(?:ure)?\.?\s*[A-Z0-9]|TABLE\s+[A-Z0-9]|Table\s+[A-Z0-9]|图\s*[0-9]|表\s*[0-9])",
            stripped,
            re.IGNORECASE,
        ):
            marked.append(f"【CAPTION】{stripped}【/CAPTION】")
        elif re.search(
            r"(?:as\s+shown\s+in\s+Fig(?:ure)?\.?\s*[0-9]|as\s+presented\s+in\s+Table\s+[0-9]|"
            r"illustrated\s+in\s+Fig(?:ure)?\.?\s*[0-9]|see\s+Fig(?:ure)?\.?\s*[0-9]|"
            r"如图\s*[0-9]|如表\s*[0-9])",
            stripped,
            re.IGNORECASE,
        ):
            marked.append(f"【REF】{stripped}【/REF】")
        else:
            marked.append(line)
    return "\n".join(marked)


def call_translate(api_base, api_key, model, paper_text):
    if not api_base or not api_key:
        logger.warning("Translate skipped: API not configured")
        return None

    url = api_base.strip()
    if not url.endswith("/chat/completions"):
        url = url.rstrip("/") + "/chat/completions"

    marked_text = _mark_captions(compact_text(paper_text))

    prompt = f"""You are an expert academic translator and document reconstructor. Your job is twofold:
1. **Reconstruct** the original paper structure from raw PDF-extracted text
2. **Translate** it into precise, publication-quality Chinese

The text below was extracted from a PDF automatically. PDF extraction often produces:
- Garbled table data (columns/rows scrambled, missing separators)
- Missing or merged words (no spaces between some words)
- Split words across lines (hyphenation artifacts)
- Mathematical symbols as garbled characters
- Figure/table content mixed with body text
- Wrong paragraph ordering (multi-column layouts)

## Step 1: Reconstruct the document
First, understand the logical structure. Identify:
- Section headings (Introduction, Method, Results, etc.)
- Figure/table captions and their associated content
- Table structures (rows, columns, data cells) — even if they appear garbled
- Paragraph boundaries that PDF extraction destroyed
- Lists, bullet points, numbered items

Reconstruct tables carefully: identify the column headers and row labels even when PDF extraction jumbles them. Use your understanding of academic paper conventions to place text correctly.

## Step 2: Translate to Chinese
After understanding the structure, translate EVERYTHING completely.

### Terminology Precision
- Each technical term must use the standard Chinese academic translation
- On first occurrence in each section, format as: 中文术语 (English term)
- Maintain term consistency throughout
- For widely known acronyms (CNN, LSTM, BERT, GPU, etc.), keep the English acronym

### Figures & Tables (图/表) — CRITICAL

Figure/table captions are marked with 【CAPTION】...【/CAPTION】. In-text references are marked with 【REF】...【/REF】.

**Handle ALL caption patterns:**
- "Figure 1:", "Fig. 2:", "Figure S1:", "Figure 1a:", "Figure 1(a):"
- "Table 1:", "Table S1:", "TABLE I:", "Tab. 1:"
- "Supplementary Figure 1:", "Supplementary Table 1:"
- "图 1:", "表 1:", "Figure 1. "
- Also detect captions NOT marked — even partial/fragmented ones like "Fig 1 . Overview"

**Caption translation rules (with examples):**
- "Figure 1: Architecture of the proposed model" → "图 1: 所提模型的架构"
- "Fig. 2: Experimental results on ImageNet" → "图 2: 在 ImageNet 上的实验结果"
- "Table 1: Comparison with state-of-the-art methods" → "表 1: 与最先进方法的对比"
- "Figure S1: Supplementary analysis of..." → "附图 1: 关于...的补充分析"
- "Figure 3a: Training loss curve" → "图 3a: 训练损失曲线"
- "TABLE I: PARAMETER CONFIGURATIONS" → "表 I: 参数配置"
- "Figure 2: Distribution of (a) accuracy, (b) F1 score" → "图 2: (a) 准确率、(b) F1 分数的分布"
- Never renumber figures/tables
- Preserve all sub-figure labels (a, b, c, 1, 2, 3) as-is

**In-text reference translation rules (with examples):**
- "as shown in Figure 3" → "如图 3 所示"
- "as presented in Table 2" → "如表 2 所示"
- "illustrated in Fig. 5" → "如图 5 所示"
- "see Figure 1a" → "见图 1a"
- "in Figures 3 and 4" → "在图 3 和图 4 中"
- "as shown in Figs. 3-5" → "如图 3-5 所示"
- "Table 2 summarizes the results" → "表 2 总结了实验结果"
- "Figure 1 compares different methods" → "图 1 比较了不同方法"
- "(see Figure 2, Table 1)" → "（见图 2 和表 1）"
- "As can be seen from Fig. 3, the..." → "从图 3 可以看出，..."
- "We report results in Table 4" → "我们在表 4 中报告了结果"

**Table reconstruction:** PDF tables often lose their structure. You MUST reconstruct them:
- Identify column headers (they are usually the first row)
- Group data correctly under each column
- Translate only headers and row labels, keep numerical data as-is
- Output tables in a clean format using | separators
- Example garbled input: "Accuracy 95.2 Precision 93.1 F1 94.5 Method CNN"
- Should be reconstructed as: "| Method | Accuracy | Precision | F1 |\n| CNN | 95.2 | 93.1 | 94.5 |"

### Formulas & Equations
- Preserve ALL mathematical expressions and equations EXACTLY
- If PDF extraction garbled a formula (e.g., "E = mc^2" → "E m c 2"), reconstruct it correctly
- Do NOT translate variable names (x, y, θ, λ, etc.)
- Keep equation numbers and labels intact
- Translate surrounding text

### Citations & References
- Preserve ALL citation markers: [1], [2,3], (Smith et al., 2020)
- Keep author names in original form
- Preserve reference list format and numbering
- If references are garbled, reconstruct them using citation patterns

### Style & Tone
- Academic, formal Chinese
- Maintain the original paragraph flow — do not reorder content
- Preserve all numerical values, percentages, statistical measures exactly

### Preserve Academic Special Characters
- Keep ALL mathematical symbols: ∑ ∫ ∂ √ ∈ ∀ ∃ ∞ ≈ ≠ ≤ ≥ ± ∓ × ÷
- Keep ALL Greek letters: α β γ δ ε ζ η θ ι κ λ μ ν ξ ο π ρ σ τ υ φ χ ψ ω
- Keep ALL arrows: → ← ↑ ↓ ⇒ ⇔ ↔ ↦
- Keep ALL subscripts/superscripts: ₂ ³ ⁿ ⁻ etc.
- Keep ALL set/logical symbols: ∩ ∪ ⊆ ⊂ ⊃ ⊇ ¬ ∧ ∨
- For formulas reconstructed from garbled PDF text, use proper Unicode math symbols
- Do NOT replace math symbols with ASCII equivalents (e.g. keep "α" not "alpha")

### Table Formatting
- Use `|` `-` `+` for table borders
- Separate tables from surrounding text with blank lines

### What NOT to do
- Do NOT add explanations or meta-commentary
- Do NOT omit any content
- Do NOT invent data
- Do NOT translate code or algorithm pseudocode

---

## Paper to Translate (raw PDF extraction — needs reconstruction):

{marked_text}
"""
    payload = {
        "model": model or "deepseek-v4-flash",
        "messages": [
            {"role": "system", "content": "You are a precise academic translator. You only output the Chinese translation, never add extra text or commentary."},
            {"role": "user", "content": prompt},
        ],
        "temperature": 0.1,
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

    try:
        with request.urlopen(req, timeout=300) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        content = data["choices"][0]["message"]["content"]
        logger.info("Translation completed successfully")
        return content.strip()
    except error.HTTPError as e:
        body = e.read().decode("utf-8", errors="ignore")
        logger.warning(f"Translate API HTTP error {e.code}: {body[:200]}")
        raise
    except (error.URLError, TimeoutError, KeyError, json.JSONDecodeError) as e:
        logger.warning(f"Translate API error: {e}")
        raise


def _is_mostly_english(text):
    latin = len(re.findall(r"[A-Za-z]", text or ""))
    cjk = len(re.findall(r"[\u4e00-\u9fff]", text or ""))
    return latin > max(120, cjk * 2)
