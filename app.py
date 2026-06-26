"""
PDF 英文→中文 翻译工具 - Web版
Flask 后端服务 - 支持批量处理
"""
import os
import re
import uuid
import zipfile
import io
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
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB
UPLOAD_DIR = Path('/tmp/pdf_translator_uploads')
OUTPUT_DIR = Path('/tmp/pdf_translator_outputs')
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

tasks = {}
tasks_lock = Lock()

# ============================================================
# 字体配置
# ============================================================
_CJK_FONT_FILES = [
    "C:/Windows/Fonts/simsun.ttc",
    "C:/Windows/Fonts/simhei.ttf",
    "C:/Windows/Fonts/msyh.ttc",
    "C:/Windows/Fonts/msyhbd.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansSC-Regular.otf",
    "/System/Library/Fonts/STHeiti Light.ttc",
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

                full_text = " ".join(s["text"] for s in line_spans)
                translated = self.dictionary.translate(full_text)

                if translated == full_text:
                    joined = "".join(s["text"] for s in line_spans)
                    t2 = self.dictionary.translate(joined)
                    if t2 != joined:
                        translated = t2

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
        orig_text = span_data["text"].strip()
        if not translated or translated == orig_text:
            return

        original_size = span_data["size"]
        color = span_data["color"]
        center = span_data.get("cell_center") or ((x0 + x1) / 2)
        span_w, page_w = x1 - x0, page.rect.width
        font_size = original_size

        while font_size > 4.0:
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
    """批量翻译：支持单文件或多文件上传"""
    pdf_files = request.files.getlist('pdfs')
    # 兼容旧版单文件字段名
    if not pdf_files:
        pf = request.files.get('pdf')
        if pf and pf.filename.endswith('.pdf'):
            pdf_files = [pf]

    if not pdf_files:
        return jsonify({'error': '请上传 PDF 文件'}), 400

    excel_file = request.files.get('excel')
    task_id = str(uuid.uuid4())[:8]
    task_dir = UPLOAD_DIR / task_id
    task_dir.mkdir(parents=True, exist_ok=True)
    out_dir = OUTPUT_DIR / task_id
    out_dir.mkdir(parents=True, exist_ok=True)

    # 加载词库
    excel_path = None
    dict_count = 0
    if excel_file and excel_file.filename.endswith('.xlsx'):
        excel_path = task_dir / 'dictionary.xlsx'
        excel_file.save(str(excel_path))

    translator = PDFTranslator()
    if excel_path:
        dict_count = translator.load_dictionary(str(excel_path))

    # 保存所有 PDF
    pdf_info = []
    for i, pf in enumerate(pdf_files):
        if not pf.filename.endswith('.pdf'):
            continue
        stem = Path(pf.filename).stem
        saved = task_dir / f'{i}_{pf.filename}'
        pf.save(str(saved))
        pdf_info.append({
            'index': i,
            'original_name': pf.filename,
            'stem': stem,
            'saved_path': str(saved),
            'output_name': f'{stem}_中文.pdf',
            'output_path': str(out_dir / f'{stem}_中文.pdf'),
        })

    if not pdf_info:
        return jsonify({'error': '没有有效的 PDF 文件'}), 400

    total_files = len(pdf_info)
    total_matched = 0

    with tasks_lock:
        tasks[task_id] = {
            'status': 'processing', 'progress': 0,
            'total': total_files, 'files': [],
            'current_file': '', 'created': datetime.now().isoformat(),
        }

    try:
        for fi, info in enumerate(pdf_info):
            def progress_cb(current, total):
                pct = int(current / total * 100) if total > 0 else 0
                overall = int((fi / total_files) * 100 + pct / total_files)
                with tasks_lock:
                    if task_id in tasks:
                        tasks[task_id]['progress'] = overall
                        tasks[task_id]['current_file'] = f'[{fi+1}/{total_files}] {info["original_name"]}'

            translator._found_words = set()
            translator.translate_pdf(info['saved_path'], info['output_path'], progress_cb)
            m = len(translator._found_words)
            total_matched += m

            with tasks_lock:
                if task_id in tasks:
                    tasks[task_id]['files'].append({
                        'name': info['output_name'],
                        'download_url': f'/api/download-file/{task_id}/{info["output_name"]}',
                        'matched': m,
                    })

        result = {
            'status': 'done', 'progress': 100,
            'total': total_files, 'files': [],
            'download_zip': f'/api/download-zip/{task_id}',
            'dict_count': dict_count, 'matched': total_matched,
        }
        with tasks_lock:
            result['files'] = tasks[task_id].get('files', [])
            tasks[task_id] = {**tasks[task_id], **result}
        return jsonify(result)

    except Exception as e:
        with tasks_lock:
            tasks[task_id] = {'status': 'error', 'error': str(e)}
        return jsonify({'error': str(e), 'detail': traceback.format_exc()}), 500


@app.route('/api/status/<task_id>')
def api_status(task_id):
    with tasks_lock:
        return jsonify(tasks.get(task_id, {}))


@app.route('/api/download-zip/<task_id>')
def api_download_zip(task_id):
    out_dir = OUTPUT_DIR / task_id
    if not out_dir.exists():
        return jsonify({'error': '文件不存在或已过期'}), 404
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for f in sorted(out_dir.glob('*.pdf')):
            zf.write(str(f), f.name)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f'pdf_translated_{task_id}.zip',
                     mimetype='application/zip')


@app.route('/api/download-file/<task_id>/<path:filename>')
def api_download_file(task_id, filename):
    fp = OUTPUT_DIR / task_id / filename
    if not fp.exists():
        return jsonify({'error': '文件不存在'}), 404
    return send_file(str(fp), as_attachment=True,
                     download_name=filename, mimetype='application/pdf')


# ============================================================
# 启动
# ============================================================
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
