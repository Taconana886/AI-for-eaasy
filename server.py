import json
import logging
import os
import re
import threading
import time
import uuid
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from io import BytesIO
from pathlib import Path
from urllib.parse import urlparse, quote
from urllib import request as url_request

import cgi

from local_parser import extract_sections, extract_title, generate_outline
from ai_client import call_ai_outline, call_translate
from ppt_builder import PPTBuilder
from rag_index import RAGIndex
from agent_pipeline import run_agent_pipeline
from visual_qa import check_pptx, auto_fix_pptx
from pdf_extractor import extract_text_and_images

logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "outputs"
UPLOAD_DIR = BASE_DIR / "uploads"
HISTORY_DIR = BASE_DIR / "history"
HISTORY_FILE = HISTORY_DIR / "records.json"
STATIC_DIR = BASE_DIR / "static"

JOBS = {}
JOB_LOCK = threading.Lock()


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
    unique_name = f"{uuid.uuid4().hex}_{filename}"
    (UPLOAD_DIR / unique_name).write_bytes(raw)
    suffix = Path(filename).suffix.lower()
    if suffix in {".txt", ".md"}:
        return raw.decode("utf-8", errors="ignore")
    if suffix == ".pdf":
        return _extract_pdf_text(raw)
    return raw.decode("utf-8", errors="ignore")


def _extract_pdf_text(raw):
    try:
        from io import BytesIO as _BytesIO
        from pdfminer.high_level import extract_text as pdfminer_extract
        from pdfminer.layout import LAParams
        laparams = LAParams(
            detect_vertical=True,
            all_texts=True,
            line_margin=0.5,
            word_margin=0.1,
            char_margin=2.0,
        )
        text = pdfminer_extract(_BytesIO(raw), laparams=laparams)
        if text and len(text.strip()) > 50:
            return text.strip()
    except Exception:
        pass
    try:
        import pypdf
        reader = pypdf.PdfReader(BytesIO(raw))
        texts = []
        for page in reader.pages:
            t = page.extract_text()
            if t:
                texts.append(t)
        return "\n".join(texts)
    except Exception:
        return ""


def safe_filename(title, ext=".pptx"):
    name = re.sub(r"[^\w\u4e00-\u9fff\-]+", "_", title or "论文阅读汇报").strip("_")
    return (name[:36] or "论文阅读汇报") + ext


def save_history(record):
    try:
        HISTORY_DIR.mkdir(parents=True, exist_ok=True)
        if HISTORY_FILE.exists():
            data = json.loads(HISTORY_FILE.read_text("utf-8"))
        else:
            data = []
        data.insert(0, record)
        if len(data) > 50:
            data = data[:50]
        HISTORY_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), "utf-8")
    except Exception as e:
        logger.warning(f"Failed to save history: {e}")


def load_history():
    try:
        if HISTORY_FILE.exists():
            return json.loads(HISTORY_FILE.read_text("utf-8"))
    except Exception:
        pass
    return []


def _do_translate(job_id, paper_text, form):
    translated = call_translate(
        form.get("api_base"),
        form.get("api_key"),
        form.get("model"),
        paper_text,
    )
    if not translated:
        return None
    title_for_fn = extract_title(paper_text) or "论文"
    base = re.sub(r"[^\w\u4e00-\u9fff\-]+", "_", title_for_fn).strip("_")[:30] or "论文"
    txt_name = f"{base}_中文版.txt"
    pdf_name = f"{base}_中文版.pdf"
    txt_path = OUTPUT_DIR / f"{job_id}_{txt_name}"
    pdf_path = OUTPUT_DIR / f"{job_id}_{pdf_name}"
    txt_path.write_text(translated, encoding="utf-8")
    _generate_pdf(translated, pdf_path)
    return txt_path


CHECKPOINT_DIR = BASE_DIR / "checkpoints"


def _save_checkpoint(job_id, stage, data):
    try:
        CHECKPOINT_DIR.mkdir(parents=True, exist_ok=True)
        cp_path = CHECKPOINT_DIR / f"{job_id}_{stage}.json"
        serializable = _make_serializable(data)
        cp_path.write_text(json.dumps(serializable, ensure_ascii=False), "utf-8")
    except Exception as e:
        logger.warning(f"Checkpoint save failed ({stage}): {e}")


def _make_serializable(obj):
    if isinstance(obj, dict):
        return {k: _make_serializable(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_make_serializable(v) for v in obj]
    if isinstance(obj, (str, int, float, bool)):
        return obj
    try:
        return str(obj)
    except Exception:
        return None


def _load_checkpoint(job_id, stage):
    cp_path = CHECKPOINT_DIR / f"{job_id}_{stage}.json"
    if cp_path.exists():
        try:
            return json.loads(cp_path.read_text("utf-8"))
        except Exception:
            pass
    return None


def _make_checkpoint_callback(job_id):
    def callback(stage, data):
        _save_checkpoint(job_id, stage, data)
    return callback


def worker(job_id, form, file_field):
    try:
        set_job(job_id, status="running", progress=10, message="接收论文与接口参数...")

        paper_text = (form.get("paper_text") or "") + "\n" + read_upload_text(file_field)
        paper_text = paper_text.strip()

        if len(paper_text) < 40:
            raise ValueError("论文内容太少。请粘贴论文文本，或上传可读取的 TXT/PDF。")

        extracted_images = []
        if file_field is not None and getattr(file_field, "filename", ""):
            suffix = Path(file_field.filename).suffix.lower()
            if suffix == ".pdf":
                set_job(job_id, progress=12, message="正在用 PDF 提取器解析文档...")
                try:
                    pdf_result = extract_text_and_images(file_field.file)
                    if pdf_result.get("text"):
                        paper_text = pdf_result["text"].strip()
                    extracted_images = pdf_result.get("images", [])
                    set_job(job_id, message=f"PDF 解析完成，提取 {len(extracted_images)} 张图片")
                except Exception as e:
                    logger.warning(f"PDF extraction failed: {e}")
                    file_field.file.seek(0)

        _save_checkpoint(job_id, "input", {
            "text_length": len(paper_text),
            "images_count": len(extracted_images),
        })

        mode = form.get("mode") or "ppt"
        translate = form.get("translate") == "true"

        if mode == "translate":
            set_job(job_id, progress=30, message="正在翻译论文为中文...")
            translate_path = _do_translate(job_id, paper_text, form)
            if not translate_path or not translate_path.exists():
                raise ValueError("翻译失败，请检查 AI 接口配置。")
            extra = {
                "translate_download_url": f"/api/download-translate/{job_id}",
                "translate_filename": translate_path.name,
            }
            pdf_path = translate_path.with_suffix(".pdf")
            if pdf_path.exists():
                extra["translate_pdf_url"] = f"/api/download-translate-pdf/{job_id}"
                extra["translate_pdf_filename"] = pdf_path.name
            set_job(job_id, status="done", progress=100, message="翻译完成！", **extra)
            save_history({
                "job_id": job_id,
                "title": extract_title(paper_text) or "论文",
                "language": "中文翻译",
                "created_at": now(),
            })
            return

        slide_count = max(4, min(12, int(form.get("slides") or 8)))
        language = form.get("language") or "中文"
        style = form.get("style") or "academic"

        imported_style = None
        if form.get("template_pptx"):
            try:
                from ppt_builder import PPTBuilder as PBuilder
                imported_style = PBuilder.import_template(form["template_pptx"])
                if imported_style:
                    imported_style["imported"] = True
                    logger.info(f"Template imported: {imported_style.get('name')}")
            except Exception as e:
                logger.warning(f"Template import failed: {e}")

        translate_path = None
        if translate:
            set_job(job_id, progress=15, message="正在翻译论文为中文...")
            try:
                translate_path = _do_translate(job_id, paper_text, form)
                if translate_path:
                    set_job(job_id, message="论文翻译完成")
            except Exception as exc:
                logger.warning(f"Translation failed: {exc}")
                set_job(job_id, message=f"翻译未成功（将跳过翻译）：{exc}")
                time.sleep(0.5)

        _save_checkpoint(job_id, "translate", {"done": translate_path is not None})

        set_job(job_id, progress=25, message="正在提取论文核心信息并构建 RAG 索引...")
        sections = extract_sections(paper_text)
        rag = RAGIndex()
        rag.build(paper_text)
        _save_checkpoint(job_id, "rag", {"chunks": len(rag.chunks)})
        time.sleep(0.2)

        set_job(job_id, progress=45, message="正在通过多智能体流程生成大纲...")
        outline = None
        ai_used = False

        try:
            outline = run_agent_pipeline(
                paper_text, slide_count, language, style, form,
                checkpoint_callback=_make_checkpoint_callback(job_id),
            )
            if outline and outline.get("slides"):
                ai_used = bool(form.get("api_key"))
                set_job(job_id, message="多智能体大纲生成成功")
        except Exception as exc:
            logger.warning(f"Agent pipeline failed: {exc}")
            set_job(job_id, message=f"智能体流程未成功，使用本地解析：{exc}")

        if not outline or not outline.get("slides"):
            outline = generate_outline(paper_text, slide_count, language)
            set_job(job_id, message="已使用本地解析生成大纲")

        if extracted_images and outline.get("slides"):
            import re as _re
            for slide in outline["slides"]:
                slide.setdefault("images", [])
                related = _match_images_to_slide(slide, extracted_images)
                slide["images"].extend(related[:2])

        _save_checkpoint(job_id, "outline", {"slide_count": len(outline.get("slides", []))})

        set_job(job_id, progress=70, message=f"正在使用「{style}」风格设计幻灯片...")
        time.sleep(0.3)

        OUTPUT_DIR.mkdir(exist_ok=True)
        deck_title = outline.get("deck_title") or extract_title(paper_text) or "论文阅读汇报"
        filename = safe_filename(deck_title)
        output_path = OUTPUT_DIR / f"{job_id}_{filename}"

        set_job(job_id, progress=85, message="正在写入 PowerPoint 文件...")
        if imported_style:
            from ppt_builder import PPTBuilder as PBuilder, STYLES
            STYLES["imported"] = imported_style
            builder = PPTBuilder(style="imported")
        else:
            builder = PPTBuilder(style=style)
        builder.build(outline, output_path)

        _save_checkpoint(job_id, "pptx", {"path": str(output_path)})

        set_job(job_id, progress=92, message="正在检查视觉质量...")
        try:
            qa_report = check_pptx(str(output_path))
            if not qa_report.passed or qa_report.issues:
                logger.info(f"Visual QA: {qa_report.summary()}")
                fixed_path = auto_fix_pptx(str(output_path))
                if fixed_path != str(output_path):
                    output_path = Path(fixed_path)
                    set_job(job_id, message="已完成视觉优化")
        except Exception as e:
            logger.warning(f"Visual QA failed: {e}")

        extra = {}
        if translate_path and translate_path.exists():
            extra["translate_download_url"] = f"/api/download-translate/{job_id}"
            extra["translate_filename"] = translate_path.name
            pdf_path = translate_path.with_suffix(".pdf")
            if pdf_path.exists():
                extra["translate_pdf_url"] = f"/api/download-translate-pdf/{job_id}"
                extra["translate_pdf_filename"] = pdf_path.name

        set_job(
            job_id,
            status="done",
            progress=100,
            message="完成！",
            output=str(output_path),
            download_url=f"/api/download/{job_id}",
            filename=filename,
            **extra,
        )

        save_history({
            "job_id": job_id,
            "title": deck_title,
            "style": style,
            "language": language,
            "slide_count": slide_count,
            "ai_used": ai_used,
            "filename": filename,
            "created_at": now(),
            "images_count": len(extracted_images),
        })

    except Exception as exc:
        logger.exception("Worker failed")
        set_job(job_id, status="error", progress=100, message="生成失败。", error=str(exc))
        cp_data = _load_checkpoint(job_id, "outline")
        if cp_data:
            set_job(job_id, message="已保存部分进度，可尝试恢复")


def _match_images_to_slide(slide, images):
    title = (slide.get("title") or "").lower()
    if not images or not title:
        return []
    matched = []
    for img in images:
        alt = (img.get("alt_text") or "").lower()
        if any(word in alt for word in title.split() if len(word) > 3):
            matched.append(img)
        if len(matched) >= 2:
            break
    return matched or images[:1]


def _generate_pdf(text, output_path):
    try:
        from fpdf import FPDF
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        font_added = False
        for p in [
            "/mnt/c/Windows/Fonts/msyh.ttc",
            "/mnt/c/Windows/Fonts/simsun.ttc",
            "C:/Windows/Fonts/msyh.ttc",
            "C:/Windows/Fonts/simsun.ttc",
            "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
            "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
        ]:
            if Path(p).exists():
                try:
                    pdf.add_font("CJK", "", str(p))
                    font_added = True
                    break
                except Exception:
                    continue
        if font_added:
            pdf.set_font("CJK", "", 11)
        else:
            pdf.set_font("Courier", "", 11)
        for para in text.split("\n"):
            para = para.strip()
            if not para:
                pdf.ln(4)
                continue
            if font_added:
                safe = re.sub(r"[^\u0020-\u007e\u00a0-\u00ff\u0300-\u036f\u0370-\u03ff\u2000-\u206f\u2070-\u209f\u20a0-\u20cf\u2100-\u214f\u2150-\u218f\u2190-\u21ff\u2200-\u22ff\u2300-\u23ff\u2460-\u24ff\u2500-\u257f\u2580-\u259f\u25a0-\u25ff\u2600-\u26ff\u2e80-\u2eff\u3000-\u303f\u3040-\u309f\u30a0-\u30ff\u3100-\u312f\u31f0-\u31ff\u3200-\u33ff\u3400-\u4dbf\u4e00-\u9fff\ua000-\ua4cf\uf900-\ufaff\ufe30-\ufe4f\uff00-\uffef\U0001d400-\U0001d7ff]", "\ufffd", para)
                pdf.multi_cell(0, 6, safe)
            else:
                pdf.multi_cell(0, 6, para.encode("ascii", errors="replace").decode("ascii"))
            pdf.ln(2)
        pdf.output(str(output_path))
        return True
    except Exception as e:
        logger.warning(f"PDF generation failed: {e}")
        return False


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        logger.info(f"{self.address_string()} {fmt % args}")

    def _send_json(self, payload, status=HTTPStatus.OK):
        raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(raw)

    def _find_file_by_job(self, job_id):
        with JOB_LOCK:
            job = JOBS.get(job_id)
        if job and job.get("output"):
            path = Path(job["output"])
            if path.exists():
                return path
        pattern = f"{job_id}_*.pptx"
        for f in OUTPUT_DIR.glob(pattern):
            return f
        return None

    def _send_file(self, path, content_type=None, download_name=None):
        path = Path(path)
        if not path.exists() or not path.is_file():
            self._send_json({"error": "file not found"}, HTTPStatus.NOT_FOUND)
            return
        raw = path.read_bytes()
        if not content_type:
            suffix = path.suffix.lower()
            content_type = {
                ".html": "text/html; charset=utf-8",
                ".css": "text/css; charset=utf-8",
                ".js": "application/javascript; charset=utf-8",
                ".png": "image/png",
                ".jpg": "image/jpeg",
                ".svg": "image/svg+xml",
            }.get(suffix, "application/octet-stream")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(raw)))
        self.send_header("Access-Control-Allow-Origin", "*")
        if download_name:
            ascii_name = download_name.encode("ascii", errors="replace").decode("ascii")
            encoded = quote(download_name.encode("utf-8"))
            self.send_header("Content-Disposition", f'attachment; filename="{ascii_name}"; filename*=UTF-8\'\'{encoded}')
        self.end_headers()
        self.wfile.write(raw)

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path

        if path == "/" or path == "/index.html":
            html_path = STATIC_DIR / "index.html"
            if html_path.exists():
                self._send_file(html_path)
            else:
                self._send_json({"error": "frontend not found"}, HTTPStatus.NOT_FOUND)
            return

        if path.startswith("/static/"):
            file_path = STATIC_DIR / path[len("/static/"):]
            self._send_file(file_path)
            return

        if path.startswith("/api/jobs/"):
            job_id = path.rsplit("/", 1)[-1]
            with JOB_LOCK:
                job = JOBS.get(job_id)
            if not job:
                self._send_json({"error": "job not found"}, HTTPStatus.NOT_FOUND)
                return
            self._send_json(job)
            return

        if path.startswith("/api/download/"):
            job_id = path.rsplit("/", 1)[-1]
            with JOB_LOCK:
                job = JOBS.get(job_id)
                filename = job.get("filename") if job else None
            file_path = self._find_file_by_job(job_id)
            if not file_path:
                self._send_json({"error": "file not found"}, HTTPStatus.NOT_FOUND)
                return
            self._send_file(
                file_path,
                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                download_name=filename or file_path.name,
            )
            return

        if path.startswith("/api/download-translate/"):
            job_id = path.rsplit("/", 1)[-1]
            with JOB_LOCK:
                job = JOBS.get(job_id)
            if not job or not job.get("translate_download_url"):
                self._send_json({"error": "translation not found"}, HTTPStatus.NOT_FOUND)
                return
            for f in OUTPUT_DIR.glob(f"{job_id}_*_中文版.txt"):
                self._send_file(f, "text/plain; charset=utf-8", download_name=f.name)
                return
            self._send_json({"error": "file not found"}, HTTPStatus.NOT_FOUND)
            return

        if path.startswith("/api/download-translate-pdf/"):
            job_id = path.rsplit("/", 1)[-1]
            with JOB_LOCK:
                job = JOBS.get(job_id)
            if not job or not job.get("translate_pdf_url"):
                self._send_json({"error": "PDF not found"}, HTTPStatus.NOT_FOUND)
                return
            for f in OUTPUT_DIR.glob(f"{job_id}_*_中文版.pdf"):
                self._send_file(f, "application/pdf", download_name=f.name)
                return
            self._send_json({"error": "file not found"}, HTTPStatus.NOT_FOUND)
            return

        if path == "/api/history":
            self._send_json(load_history())
            return

        if path == "/api/styles":
            from ppt_builder import STYLES
            styles_info = {k: {"name": v["name"]} for k, v in STYLES.items()}
            self._send_json(styles_info)
            return

        if path.startswith("/api/checkpoint/"):
            job_id = path.rsplit("/", 1)[-1]
            stages = ["input", "translate", "rag", "outline", "pptx"]
            available = {}
            for stage in stages:
                cp = _load_checkpoint(job_id, stage)
                if cp:
                    available[stage] = cp
            self._send_json({"job_id": job_id, "checkpoints": available})
            return

        self._send_json({"error": "not found"}, HTTPStatus.NOT_FOUND)

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path

        if path == "/api/upload-template":
            ctype, pdict = cgi.parse_header(self.headers.get("content-type"))
            if ctype != "multipart/form-data":
                self._send_json({"error": "请使用 multipart/form-data"}, HTTPStatus.BAD_REQUEST)
                return
            pdict["boundary"] = bytes(pdict["boundary"], "utf-8")
            form_data = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ={"REQUEST_METHOD": "POST"})
            if "template" not in form_data:
                self._send_json({"error": "请上传 PPTX 文件"}, HTTPStatus.BAD_REQUEST)
                return
            template_file = form_data["template"]
            if not template_file.filename:
                self._send_json({"error": "请上传 PPTX 文件"}, HTTPStatus.BAD_REQUEST)
                return
            UPLOAD_DIR.mkdir(exist_ok=True)
            dest = UPLOAD_DIR / f"template_{uuid.uuid4().hex[:8]}.pptx"
            with open(dest, "wb") as f:
                f.write(template_file.file.read())
            from ppt_builder import PPTBuilder
            style = PPTBuilder.import_template(str(dest))
            if style:
                from ppt_builder import STYLES
                STYLES["imported"] = style
                self._send_json({"success": True, "style": {"name": style.get("name", "导入模板"), "key": "imported"}})
            else:
                self._send_json({"error": "无法解析模板"}, HTTPStatus.BAD_REQUEST)
            return

        if path != "/api/generate":
            self._send_json({"error": "not found"}, HTTPStatus.NOT_FOUND)
            return

        ctype, pdict = cgi.parse_header(self.headers.get("content-type"))
        if ctype != "multipart/form-data":
            self._send_json({"error": "请使用 multipart/form-data 提交。"}, HTTPStatus.BAD_REQUEST)
            return

        pdict["boundary"] = bytes(pdict["boundary"], "utf-8")
        form_data = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ={"REQUEST_METHOD": "POST"})

        form = {}
        field_keys = ["paper_text", "api_base", "api_key", "model", "slides", "language", "style", "translate", "mode", "template_style"]
        for key in field_keys:
            item = form_data[key] if key in form_data else None
            form[key] = item.value if item is not None and not item.filename else ""

        if not form.get("style"):
            form["style"] = "academic"

        if "template_file" in form_data:
            template_file = form_data["template_file"]
            if template_file.filename:
                UPLOAD_DIR.mkdir(exist_ok=True)
                tpl_dest = UPLOAD_DIR / f"tpl_{uuid.uuid4().hex[:8]}.pptx"
                with open(tpl_dest, "wb") as f:
                    f.write(template_file.file.read())
                form["template_pptx"] = str(tpl_dest)

        file_field = form_data["paper_file"] if "paper_file" in form_data else None

        job_id = uuid.uuid4().hex[:12]
        set_job(job_id, status="queued", progress=0, message="任务已创建。")

        thread = threading.Thread(target=worker, args=(job_id, form, file_field), daemon=True)
        thread.start()

        self._send_json({"job_id": job_id, "status_url": f"/api/jobs/{job_id}"}, HTTPStatus.ACCEPTED)

    def do_OPTIONS(self):
        self.send_response(HTTPStatus.NO_CONTENT)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()


def create_server(host="127.0.0.1", port=8765):
    OUTPUT_DIR.mkdir(exist_ok=True)
    UPLOAD_DIR.mkdir(exist_ok=True)
    HISTORY_DIR.mkdir(exist_ok=True)
    STATIC_DIR.mkdir(exist_ok=True)
    server = ThreadingHTTPServer((host, port), Handler)
    return server
