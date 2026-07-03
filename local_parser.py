import re
from datetime import datetime


SECTION_PATTERNS = [
    (r"(?:^|\n)\s*(?:Abstract|摘要)\s*\n", "abstract"),
    (r"(?:^|\n)\s*(?:Introduction|引言|介绍|背景)\s*\n", "introduction"),
    (r"(?:^|\n)\s*(?:Related Work|相关工作|研究背景)\s*\n", "related_work"),
    (r"(?:^|\n)\s*(?:Method|Methods|Methodology|Approach|方法|方法论|方案)\s*\n", "method"),
    (r"(?:^|\n)\s*(?:Proposed|Framework|System|Model|Architecture|提出|框架|系统|模型)\s*\n", "method"),
    (r"(?:^|\n)\s*(?:Experiment|Experiments|Experimental|Evaluation|实验|评估)\s*\n", "experiments"),
    (r"(?:^|\n)\s*(?:Result|Results|实验[结果]?|结果)\s*\n", "results"),
    (r"(?:^|\n)\s*(?:Discussion|讨论)\s*\n", "discussion"),
    (r"(?:^|\n)\s*(?:Conclusion|Conclusions|总结|结论|展望)\s*\n", "conclusion"),
    (r"(?:^|\n)\s*(?:References|参考文献|Reference)\s*\n", "references"),
    (r"(?:^|\n)\s*(?:Appendix|附录)\s*\n", "appendix"),
]


def extract_sections(text):
    sections = {}
    positions = []
    for pattern, name in SECTION_PATTERNS:
        for m in re.finditer(pattern, text, re.IGNORECASE):
            positions.append((m.start(), name, m.end()))
    positions.sort()
    if not positions:
        return {"full_text": text}
    for i, (start, name, end) in enumerate(positions):
        next_start = positions[i + 1][0] if i + 1 < len(positions) else len(text)
        content = text[end:next_start].strip()
        sections[name] = content
    return sections


def extract_title(text):
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    for line in lines[:10]:
        cleaned = re.sub(r"^(?:Title|题目|论文题目)\s*[:：]\s*", "", line, flags=re.I)
        if 6 <= len(cleaned) <= 120 and not cleaned.lower().startswith(
            ("abstract", "introduction", "摘要", "引言")
        ):
            has_upper = bool(re.search(r"[A-Z]", cleaned))
            has_cjk = bool(re.search(r"[\u4e00-\u9fff]", cleaned))
            if has_upper or has_cjk:
                return cleaned
    return ""


def extract_keywords(text, top_n=10):
    words = re.findall(r"[A-Z][a-z]+(?:[\s-][A-Z][a-z]+)*|[\u4e00-\u9fff]{2,}", text or "")
    stop = {
        "this", "that", "with", "from", "paper", "method", "result", "using",
        "based", "study", "approach", "proposed", "research", "model", "data",
        "also", "show", "can", "experiment", "performance", "different",
        "however", "thus", "well", "analysis", "problem", "task", "results",
        "method", "methods", "work", "system", "approach",
    }
    counts = {}
    for word in words:
        w = word.lower().strip()
        if len(w) < 3 or w in stop:
            continue
        counts[w] = counts.get(w, 0) + 1
    sorted_words = sorted(counts.items(), key=lambda x: x[1], reverse=True)
    return [w for w, c in sorted_words[:top_n]]


def extract_metrics(text):
    metrics = []
    patterns = [
        r"(\d+[.,]?\d*)\s*[%％]",
        r"(?:accuracy|precision|recall|f1|f-score|bleu|rouge|perplexity|AUC|mAP|IoU)\s*(?::|is|of|=)\s*(\d+[.,]?\d*)",
        r"(?:达到|提升|降低|减少|增长)[了约]?\s*(\d+[.,]?\d*)\s*[%％]",
    ]
    for pattern in patterns:
        for m in re.finditer(pattern, text, re.IGNORECASE):
            context_start = max(0, m.start() - 40)
            context_end = min(len(text), m.end() + 40)
            context = text[context_start:context_end].strip()
            metrics.append(context)
    return metrics[:8]


def extract_citations(text):
    citations = set()
    patterns = [
        r"\[(\d+(?:\s*,\s*\d+)*)\]",
        r"\(([A-Z][a-z]+(?:\s+et\s+al\.?)?\s*[,，]\s*\d{4})\)",
    ]
    for pattern in patterns:
        for m in re.finditer(pattern, text):
            citations.add(m.group(1))
    return list(citations)[:12]


def split_sentences(text, max_len=8000):
    text = re.sub(r"\s+", " ", text).strip()[:max_len]
    parts = re.split(r"(?<=[。！？.!?])\s+", text)
    return [p.strip() for p in parts if len(p.strip()) > 25][:30]


def is_mostly_english(text):
    latin = len(re.findall(r"[A-Za-z]", text or ""))
    cjk = len(re.findall(r"[\u4e00-\u9fff]", text or ""))
    return latin > max(120, cjk * 2)


def generate_outline(text, slide_count=8, language="中文"):
    sections = extract_sections(text)
    title = extract_title(text) or "论文阅读汇报"
    keywords = extract_keywords(text)
    metrics = extract_metrics(text)
    citations = extract_citations(text)
    sentences = split_sentences(text)
    is_eng = is_mostly_english(text)

    section_content = {}
    for s in ["abstract", "introduction", "method", "experiments", "results", "conclusion"]:
        section_content[s] = sections.get(s, "")[:500]

    template_slides = [
        ("title", "研究背景", "这篇论文关注的问题与应用场景", "flow"),
        ("problem", "核心问题", "现有方法的限制与本文要解决的关键缺口", "matrix"),
        ("method", "方法框架", "模型、流程或系统架构的主要组成部分", "flow"),
        ("innovation", "关键创新", "论文相对已有工作的改进点和技术贡献", "matrix"),
        ("experiment", "实验设置", "数据集、评价指标与实验协议", "timeline"),
        ("results", "主要结果", "最重要的实验发现和结论", "bars"),
        ("limitation", "局限与展望", "适用边界与未来工作方向", "matrix"),
        ("conclusion", "汇报总结", "适合在汇报中强调的结论与后续问题", "flow"),
    ]

    extra_slides = [
        ("discussion", "讨论问题", "给听众的讨论点和需进一步确认的假设", "matrix"),
        ("reference", "参考信息", "论文关键词、术语和引用信息", "timeline"),
        ("reproduce", "可复现清单", "代码、数据、参数设置与消融实验", "bars"),
        ("thinking", "延伸思考", "可继续做的研究方向或工程落地机会", "flow"),
    ]

    all_slides = template_slides + extra_slides
    slides = []

    for i in range(slide_count):
        slide_type, heading, fallback, visual_type = all_slides[i % len(all_slides)]

        evidence = ""
        if slide_type == "abstract" and section_content.get("abstract"):
            evidence = section_content["abstract"][:200]
        elif slide_type == "method" and section_content.get("method"):
            evidence = section_content["method"][:200]
        elif slide_type == "results" and section_content.get("results"):
            evidence = section_content["results"][:200]
        elif slide_type == "experiment" and section_content.get("experiments"):
            evidence = section_content["experiments"][:200]
        elif slide_type == "conclusion" and section_content.get("conclusion"):
            evidence = section_content["conclusion"][:200]
        elif sentences:
            evidence = sentences[i % len(sentences)]

        bullets = [
            evidence[:160] if evidence else fallback,
        ]
        if keywords:
            kw_str = ", ".join(keywords[:5])
            bullets.append(f"关键词：{kw_str}")

        if slide_type == "results" and metrics:
            bullets.append(f"实验指标：{metrics[0][:120]}")
        elif slide_type == "reference" and citations:
            bullets.append(f"引用文献：{', '.join(citations[:5])}")
        else:
            bullets.append(fallback)

        slide = {
            "title": heading,
            "subtitle": fallback,
            "bullets": [b for b in bullets if b],
            "zh_summary": f"中文解读：{fallback}" if is_eng else f"要点总结：{fallback}",
            "english_terms": [],
            "visual_type": visual_type,
            "visual_items": keywords[:4] or ["问题", "方法", "实验", "结论"],
        }

        if is_eng and evidence:
            slide["english_terms"] = [
                evidence[:90],
                "Chinese translation: 自动提炼论文含义，生成中文讲解。",
            ]

        slides.append(slide)

    slides[0] = {
        "title": title,
        "subtitle": f"AI 自动生成的论文汇报 PPT（{language}）",
        "bullets": [
            "研究主题概览",
            f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}",
            f"语言：{language}",
            "内置图示：流程图、矩阵图、柱状图、路线图",
        ],
        "zh_summary": "本 PPT 将论文内容转化为结构化汇报，包含方法、实验、结论等核心模块。",
        "english_terms": [],
        "visual_type": "flow",
        "visual_items": ["论文输入", "结构化阅读", "图示化表达", "PPT 汇报"],
    }

    return {"deck_title": title, "slides": slides, "keywords": keywords}
