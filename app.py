"""
PDF 英文→中文 翻译工具 - Web版
Flask 后端服务
"""
import os
import re
import uuid
import shutil
import traceback
from pathlib import Path
from datetime import datetime
from threading import Lock

from flask import (
    Flask, request, render_template, jsonify,
    send_from_directory, send_file
)
import fitz
import openpyxl

# ============================================================
# 配置
# ============================================================
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB
UPLOAD_DIR = Path('/tmp/pdf_translator_uploads')
OUTPUT_DIR = Path('/tmp/pdf_translator_outputs')
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# 任务状态存储
tasks = {}
tasks_lock = Lock()

# ============================================================
# 字体配置
# ============================================================
_CJK_FONT_FILES = [
    # Windows
    "C:/Windows/Fonts/simsun.ttc",
    "C:/Windows/Fonts/simhei.ttf",
    "C:/Windows/Fonts/simkai.ttf",
    "C:/Windows/Fonts/msyh.ttc",
    "C:/Windows/Fonts/msyhbd.ttc",
    # Linux (Render.com / Ubuntu)
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/noto-cjk/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
    "/usr/share/fonts/opentype/noto/NotoSansSC-Regular.otf",
    # macOS
    "/System/Library/Fonts/STHeiti Light.ttc",
    "/System/Library/Fonts/STHeiti Medium.ttc",
    "/System/Library/Fonts/PingFang.ttc",
    "/Library/Fonts/Arial Unicode.ttf",
]

_cjk_font_cache = None


def get_cjk_fontfile():
    global _cjk_font_cache
    if _cjk_font_cache and os.path.exists(_cjk_font_cache):
        return _cjk_font_cache
    for fp in _CJK_FONT_FILES:
        if os.path.exists(fp) and os.path.getsize(fp) > 500_000:
            _cjk_font_cache = fp
            return fp
    return ""


# ============================================================
# Excel 词库
# ============================================================
class ExcelDictionary:
    def __init__(self):
        self._exact = {}
        self._align = {}

    @staticmethod
    def _make_key(text):
        return re.sub(r'[^a-zA-Z0-9]', '', text.strip()).lower()

    def load(self, excel_path):
        self._exact.clear()
        self._align.clear()
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
        count = 0
        for row in ws.iter_rows(min_row=2, max_col=3):
            if not row or len(row) < 2:
                continue
            en_raw = row[0].value
            zh_raw = row[1].value
            if en_raw is None or zh_raw is None:
                continue
            en = str(en_raw).strip()
            zh = str(zh_raw).strip()
            if en and zh:
                key = self._make_key(en)
                self._exact[key] = zh
                count += 1
                if len(row) >= 3 and row[2].value:
                    self._align[key] = str(row[2].value).strip().lower()
        wb.close()
        return count

    def translate(self, text):
        if not re.search(r'[a-zA-Z]', text):
            return text
        key = self._make_key(text)
        return self._exact.get(key, text)

    def get_align(self, text):
        key = self._make_key(text)
        return self._align.get(key, "center")


# ============================================================
# PDF 翻译引擎
# ============================================================
class PDFTranslator:
    def __init__(self):
        self.dictionary = ExcelDictionary()
        self._found_words = set()

    def load_dictionary(self, excel_path):
        return self.dictionary.load(excel_path)

    def translate_pdf(self, pdf_path, output_path, progress_cb=None):
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        fontfile = get_cjk_fontfile()

        for page_idx in range(total_pages):
            if progress_cb:
                progress_cb(page_idx, total_pages)
            page = doc[page_idx]
            self._process_page(page, page_idx, fontfile)

        doc.save(output_path, garbage=4, deflate=True)
        doc.close()
        return True

    def _process_page(self, page, page_idx, fontfile):
        try:
            blocks = page.getText("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
        except AttributeError:
            blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]

        replaced_spans = []

        for block in blocks:
            if block["type"] != 0:
                continue
            for line in block.get("lines", []):
                line_spans = []
                for span in line.get("spans", []):
                    text = span["text"].strip()
                    if not text or not re.search(r'[a-zA-Z]', text):
                        continue
                    line_spans.append({
                        "text": text,
                        "bbox": tuple(span["bbox"]),
                        "size": span.get("size", 10.0),
                        "color": self._rgb_from_span(span),
                    })

                if not line_spans:
                    continue

                if len(line_spans) > 1:
                    sx = sorted(line_spans, key=lambda s: s["bbox"][0])
                    for i, s in enumerate(sx):
                        s["cell_center"] = (
                            ((sx[i-1]["bbox"][2] + s["bbox"][0])/2 if i > 0 else s["bbox"][0])
                            + (s["bbox"][2] + (sx[i+1]["bbox"][0] if i < len(sx)-1 else s["bbox"][2]))/2
                        ) / 2

                # 整行匹配
                full_text = " ".join(s["text"] for s in line_spans)
                translated = self.dictionary.translate(full_text)

                if translated == full_text:
                    joined = "".join(s["text"] for s in line_spans)
                    t2 = self.dictionary.translate(joined)
                    if t2 != joined:
                        translated = t2
                        full_text = joined

                # 逐个 span
                if translated == full_text and len(line_spans) > 1:
                    for s in line_spans:
                        st = self.dictionary.translate(s["text"])
                        if st != s["text"]:
                            s["translated"] = st
                            self._add_redact(page, fitz.Rect(*s["bbox"]))
                            replaced_spans.append(s)
                            self._found_words.add(s["text"].strip())
                    continue

                if translated != full_text:
                    for s in line_spans:
                        self._found_words.add(s["text"].strip())
                    if not translated or not translated.strip():
                        continue
                    for i, s in enumerate(line_spans):
                        if len(line_spans) == 1:
                            s["translated"] = translated
                        else:
                            char_total = sum(len(x["text"]) for x in line_spans)
                            if char_total > 0:
                                zh_chars = len(translated)
                                start = sum(int(len(x["text"])/char_total*zh_chars) for x in line_spans[:i])
                                end = start + int(len(s["text"])/char_total*zh_chars) if i < len(line_spans)-1 else zh_chars
                                s["translated"] = translated[start:end]
                            else:
                                s["translated"] = translated
                        s["align"] = self.dictionary.get_align(s["text"]) or "center"
                        t = s.get("translated", "")
                        if t and t.strip() and t != s["text"]:
                            self._add_redact(page, fitz.Rect(*s["bbox"]))
                            replaced_spans.append(s)

        # 对齐中心聚类 (单span)
        lone = [s for s in replaced_spans if "cell_center" not in s]
        if lone:
            clusters = []
            for s in sorted(lone, key=lambda s: s["bbox"][0]):
                placed = False
                for c in clusters:
                    if abs(s["bbox"][0] - c["avg_x0"]) < 20:
                        c["spans"].append(s)
                        c["avg_x0"] = sum(x["bbox"][0] for x in c["spans"]) / len(c["spans"])
                        c["avg_x1"] = sum(x["bbox"][2] for x in c["spans"]) / len(c["spans"])
                        placed = True
                        break
                if not placed:
                    clusters.append({"spans": [s], "avg_x0": s["bbox"][0], "avg_x1": s["bbox"][2]})
            for c in clusters:
                cc = (c["avg_x0"] + c["avg_x1"]) / 2
                for s in c["spans"]:
                    s["cell_center"] = cc

        if replaced_spans:
            page.apply_redactions()
        for s in replaced_spans:
            self._replace_text(page, s, fontfile)

    def _replace_text(self, page, span_data, fontfile):
        x0, y0, x1, y1 = span_data["bbox"]
        translated = span_data["translated"].strip()
        original_text = span_data["text"].strip()

        if not translated or translated == original_text:
            return

        original_size = span_data["size"]
        color = span_data["color"]
        center = span_data.get("cell_center") or ((x0 + x1) / 2)
        span_w = x1 - x0
        page_w = page.rect.width

        # 自适应字号
        font_size = original_size
        MIN_FONT = 4.0
        while font_size > MIN_FONT:
            est_w = len(translated) * font_size
            nx0 = center - est_w / 2
            if nx0 < 10 or (nx0 + est_w) > page_w - 10 or est_w > span_w * 1.5:
                font_size -= 1.0
            else:
                break

        est_w = len(translated) * font_size
        new_x0 = max(10, min(center - est_w / 2, page_w - est_w - 10))

        kw = dict(fontname="china-ss", fontsize=font_size, color=color)
        if fontfile:
            kw["fontfile"] = fontfile

        try:
            try:
                page.insertText(fitz.Point(new_x0, y1 - 1), translated, **kw)
            except AttributeError:
                page.insert_text(fitz.Point(new_x0, y1 - 1), translated, **kw)
        except Exception:
            pass

    @staticmethod
    def _add_redact(page, rect):
        try:
            page.addRedactAnnot(rect, fill=None, cross_out=False)
        except AttributeError:
            page.add_redact_annot(rect, fill=None, cross_out=False)

    @staticmethod
    def _rgb_from_span(span):
        c = span.get("color", 0)
        if isinstance(c, (int, float)):
            c_int = int(c)
            return (c_int >> 16 & 255)/255.0, (c_int >> 8 & 255)/255.0, (c_int & 255)/255.0
        if isinstance(c, (list, tuple)) and len(c) >= 3:
            return c[0]/255.0, c[1]/255.0, c[2]/255.0
        return (0, 0, 0)


# ============================================================
# Web 路由
# ============================================================
@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/translate', methods=['POST'])
def api_translate():
    """上传文件并开始翻译"""
    pdf_file = request.files.get('pdf')
    excel_file = request.files.get('excel')

    if not pdf_file or not pdf_file.filename.endswith('.pdf'):
        return jsonify({'error': '请上传 PDF 文件'}), 400

    task_id = str(uuid.uuid4())[:8]
    task_dir = UPLOAD_DIR / task_id
    task_dir.mkdir(parents=True, exist_ok=True)

    # 保存文件
    pdf_path = task_dir / 'input.pdf'
    pdf_file.save(str(pdf_path))

    excel_path = None
    if excel_file and excel_file.filename.endswith('.xlsx'):
        excel_path = task_dir / 'dictionary.xlsx'
        excel_file.save(str(excel_path))

    output_path = OUTPUT_DIR / f'{task_id}_translated.pdf'

    # 记录任务
    with tasks_lock:
        tasks[task_id] = {
            'status': 'processing',
            'progress': 0,
            'total': 1,
            'created': datetime.now().isoformat(),
        }

    # 同步处理（小文件）
    try:
        translator = PDFTranslator()
        dict_count = 0
        if excel_path:
            dict_count = translator.load_dictionary(str(excel_path))

        def progress_cb(current, total):
            pct = int(current / total * 100) if total > 0 else 0
            with tasks_lock:
                if task_id in tasks:
                    tasks[task_id]['progress'] = pct
                    tasks[task_id]['total'] = total

        translator.translate_pdf(str(pdf_path), str(output_path), progress_cb)

        stat = {
            'status': 'done',
            'progress': 100,
            'total': 1,
            'download_url': f'/api/download/{task_id}',
            'filename': f'{pdf_path.stem}_中文翻译.pdf',
            'dict_count': dict_count,
            'matched': len(translator._found_words),
        }
        with tasks_lock:
            tasks[task_id] = stat
        stat['task_id'] = task_id
        return jsonify(stat)

    except Exception as e:
        with tasks_lock:
            tasks[task_id] = {'status': 'error', 'error': str(e)}
        return jsonify({'error': str(e), 'detail': traceback.format_exc()}), 500


@app.route('/api/status/<task_id>')
def api_status(task_id):
    with tasks_lock:
        task = tasks.get(task_id, {})
    return jsonify(task)


@app.route('/api/download/<task_id>')
def api_download(task_id):
    output_path = OUTPUT_DIR / f'{task_id}_translated.pdf'
    if not output_path.exists():
        return jsonify({'error': '文件不存在或已过期'}), 404
    return send_file(
        str(output_path),
        as_attachment=True,
        download_name=f'translated_{task_id}.pdf',
        mimetype='application/pdf'
    )


# ============================================================
# 启动
# ============================================================
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
