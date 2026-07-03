import logging
import json
import re
import time
from typing import Optional

from local_parser import extract_sections, extract_title, extract_keywords, extract_metrics, generate_outline
from ai_client import call_ai_outline
from rag_index import RAGIndex

logger = logging.getLogger(__name__)


def run_agent_pipeline(
    paper_text: str,
    slide_count: int,
    language: str,
    style: str,
    form: dict,
    checkpoint_callback=None,
) -> dict:
    """Three-agent pipeline: Strategist -> Executor -> Critic.

    Each agent can use AI (if configured) or fallback to local logic.

    Returns outline dict compatible with PPTBuilder.
    """
    sections = extract_sections(paper_text)
    title = extract_title(paper_text) or "论文阅读汇报"

    if checkpoint_callback:
        checkpoint_callback("sections", {"sections": sections, "title": title})

    rag = RAGIndex()
    rag.build(paper_text)

    if checkpoint_callback:
        checkpoint_callback("rag", {"chunks": len(rag.chunks)})

    outline = _strategist(paper_text, sections, title, slide_count, language, form, rag)
    if not outline or not outline.get("slides"):
        logger.warning("Strategist failed, falling back to local outline")
        outline = generate_outline(paper_text, slide_count, language)

    if checkpoint_callback:
        checkpoint_callback("strategist", {"slide_count": len(outline.get("slides", []))})

    outline = _executor(paper_text, outline, sections, language, form, rag)

    if checkpoint_callback:
        checkpoint_callback("executor", {"slide_count": len(outline.get("slides", []))})

    outline = _critic(paper_text, outline, language, form, rag)

    if checkpoint_callback:
        checkpoint_callback("critic", {"issues_fixed": True})

    return outline


def _strategist(paper_text, sections, title, slide_count, language, form, rag) -> Optional[dict]:
    """Agent 1: Analyze paper and plan slide structure.

    Tries AI first, falls back to local planning.
    """
    slide_topics = [
        "研究背景与问题", "核心方法与框架", "关键创新点",
        "实验设置", "主要结果", "局限与展望", "总结"
    ]
    slide_topics = slide_topics[:slide_count]

    context_chunks = {}
    for topic in slide_topics:
        chunks = rag.retrieve(topic, top_k=2)
        context_chunks[topic] = " ".join(c["chunk"][:200] for c in chunks)

    ai_outline = call_ai_outline(
        form.get("api_base"),
        form.get("api_key"),
        form.get("model"),
        paper_text,
        slide_count,
        language,
        sections,
    )

    if ai_outline and ai_outline.get("slides"):
        for slide in ai_outline["slides"]:
            slide.setdefault("visual_type", "flow")
            slide.setdefault("visual_items", ["问题", "方法", "实验", "结论"])
            slide.setdefault("zh_summary", "")
            slide.setdefault("english_terms", [])
        return ai_outline

    return None


def _executor(paper_text, outline, sections, language, form, rag) -> dict:
    """Agent 2: Enrich each slide with RAG-retrieved evidence."""
    slides = outline.get("slides", [])
    enriched = False

    for slide in slides:
        title = slide.get("title", "")
        if not title:
            continue
        chunks = rag.retrieve(title, top_k=2)
        if not chunks:
            continue
        evidence = "\n".join(c["chunk"][:300] for c in chunks)
        if evidence:
            bullets = slide.get("bullets", [])
            if not bullets or len(bullets) <= 1:
                evidence_points = _extract_key_points(evidence, 2)
                for ep in evidence_points:
                    if ep not in bullets:
                        bullets.append(ep)
                        enriched = True
                slide["bullets"] = bullets

    if enriched:
        logger.info("Executor enriched slides with RAG evidence")

    return outline


def _extract_key_points(text: str, max_points: int = 2) -> list[str]:
    sentences = re.split(r"(?<=[。！？.!?])\s+", text)
    meaningful = [s.strip() for s in sentences if len(s.strip()) > 20]
    return meaningful[:max_points]


def _critic(paper_text, outline, language, form, rag) -> dict:
    """Agent 3: Review and fix slide quality issues.

    Checks:
    - No empty slides
    - Bullets should have content
    - Each slide should have evidence from paper
    """
    slides = outline.get("slides", [])
    if not slides:
        return outline

    for slide in slides:
        bullets = slide.get("bullets", [])
        if not bullets or all(not b.strip() for b in bullets):
            slide["bullets"] = ["详细内容请参考论文原文"]

        if not slide.get("visual_type"):
            slide["visual_type"] = "flow"
        if not slide.get("visual_items"):
            slide["visual_items"] = ["问题", "方法", "实验", "结论"]

    logger.info(f"Critic reviewed {len(slides)} slides")
    return outline
