import os
import re
import json
import glob
import argparse

try:
    import docx
except ImportError:
    docx = None

import fitz


# =========================
# 基础正则
# =========================

DOCX_QUESTION_RE = re.compile(r'^\s*\d+[\.、．\s]')
DOCX_OPTION_RE = re.compile(r'^\s*[A-D][\.、．\s]')

# 兼容：
# 1.（25 联考）
# 2. （21 浙江选调）
# 16（. 25 联考事业）
# 14（. 24 联考事业）
PDF_QUESTION_RE = re.compile(
    r'^\s*(\d{1,3})\s*(?:[\.、．]\s*|[（(]\s*\.?\s*)'
)

OPTION_RE = re.compile(r'(?<![A-Za-z0-9])([A-D])\s*[\.、．]\s*')


# =========================
# DOCX 提取
# =========================

def extract_docx_questions(file_path):
    """
    解析 DOCX 文字题。
    保留你原本正确的 DOCX 提取逻辑。
    """
    if docx is None:
        raise ImportError("缺少 python-docx，请运行：python -m pip install python-docx")

    print(f"正在解析 DOCX: {file_path}")

    doc = docx.Document(file_path)
    questions = []
    current_q = None

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        if DOCX_QUESTION_RE.match(text):
            if current_q:
                questions.append(current_q)

            current_q = {
                "source": os.path.basename(file_path),
                "question_text": text,
                "options": [],
                "images": []
            }

        elif DOCX_OPTION_RE.match(text) and current_q:
            current_q["options"].append(text)

        elif current_q:
            if not current_q["options"]:
                current_q["question_text"] += "\n" + text
            else:
                current_q["options"][-1] += "\n" + text

    if current_q:
        questions.append(current_q)

    return questions


# =========================
# PDF 文本工具
# =========================

def clean_line(line):
    """
    清理页眉、页脚、章节标题等无关文本。
    """
    line = str(line or "").strip()

    if not line:
        return ""

    # 页眉
    if line == "超格学员专用":
        return ""

    # 页码
    if re.fullmatch(r"\d{1,3}", line):
        return ""

    skip_patterns = [
        r"^第一部分\s+图形推理$",
        r"^第二部分\s+逻辑判断、定义判断、类比推理$",
        r"^梯度刷题[（(].*[）)]$",
        r"^专项刷题.*$",
        r"^必然性推理.*$",
        r"^识别：.*$",
        r"^作答情况复盘$",
        r"^情况统计.*$",
        r"^暴露问题$",
        r"^第[一二三四五六七八九十]+套(?:（难度[★☆]+）)?$",
        r"^第[一二三四五六七八九十]+套（特色刷题）$",
        r"^【温馨提示】.*$",
    ]

    for pattern in skip_patterns:
        if re.fullmatch(pattern, line):
            return ""

    return line


def get_page_lines(page):
    """
    获取 PDF 页面中的文本行，保留坐标。
    """
    result = []
    data = page.get_text("dict")

    for block in data.get("blocks", []):
        if block.get("type") != 0:
            continue

        for line in block.get("lines", []):
            spans = line.get("spans", [])
            text = "".join(span.get("text", "") for span in spans).strip()
            text = clean_line(text)

            if not text:
                continue

            bbox = fitz.Rect(line.get("bbox"))

            result.append({
                "text": text,
                "x0": bbox.x0,
                "y0": bbox.y0,
                "x1": bbox.x1,
                "y1": bbox.y1,
            })

    result.sort(key=lambda d: (round(d["y0"], 1), d["x0"]))
    return result


def get_page_text(page):
    lines = get_page_lines(page)
    return "\n".join(line["text"] for line in lines).strip()


def is_graph_page(page):
    """
    判断是否属于图形推理页。

    这份精简 PDF 的结构比较稳定：
    - 页码 1-14 是图形推理
    - 从“第二部分 逻辑判断、定义判断、类比推理”开始是文字题
    """
    raw_text = page.get_text("text")

    if "第一部分 图形推理" in raw_text:
        return True

    if "梯度刷题（一）" in raw_text:
        return True

    if "梯度刷题（二）" in raw_text:
        return True

    if "梯度刷题（四）" in raw_text:
        return True

    if "第二部分 逻辑判断、定义判断、类比推理" in raw_text:
        return False

    # 通过页脚页码判断：当前精简 PDF 的图形页是原资料页码 1-12、32-33
    lines = get_page_lines(page)
    footer_numbers = []
    for line in lines:
        if line["y0"] > page.rect.height * 0.88 and re.fullmatch(r"\d{1,3}", line["text"]):
            footer_numbers.append(int(line["text"]))

    if footer_numbers:
        n = footer_numbers[-1]
        if 1 <= n <= 14 or 32 <= n <= 33:
            return True

    return False


def find_question_starts(doc):
    """
    在整个 PDF 中寻找题目起点。
    """
    starts = []

    for page_index in range(len(doc)):
        page = doc.load_page(page_index)
        lines = get_page_lines(page)

        for line in lines:
            text = line["text"]
            m = PDF_QUESTION_RE.match(text)

            if not m:
                continue

            q_no = int(m.group(1))

            starts.append({
                "page_index": page_index,
                "page_num": page_index + 1,
                "original_question_no": q_no,
                "y": line["y0"],
                "text": text,
                "is_graph": is_graph_page(page),
            })

    # 去重：避免同一题被重复识别
    unique = []
    for item in starts:
        duplicated = False

        for old in unique:
            same_page = old["page_index"] == item["page_index"]
            same_no = old["original_question_no"] == item["original_question_no"]
            close_y = abs(old["y"] - item["y"]) < 3

            if same_page and same_no and close_y:
                duplicated = True
                break

        if not duplicated:
            unique.append(item)

    unique.sort(key=lambda d: (d["page_index"], d["y"]))
    return unique


def get_question_segments(doc, start, end):
    """
    根据当前题起点和下一题起点，得到该题所在区域。
    支持跨页文字题。
    """
    segments = []

    start_page = start["page_index"]
    end_page = end["page_index"] if end else start_page

    for page_index in range(start_page, end_page + 1):
        page = doc.load_page(page_index)
        page_height = page.rect.height

        # 避开页眉页脚
        top_y = 42
        bottom_y = page_height - 42

        if page_index == start_page:
            y0 = max(top_y, start["y"] - 3)
        else:
            y0 = top_y

        if end and page_index == end_page:
            y1 = min(bottom_y, end["y"] - 5)
        else:
            y1 = bottom_y

        if y1 - y0 < 20:
            continue

        rect = fitz.Rect(0, y0, page.rect.width, y1)

        segments.append({
            "page_index": page_index,
            "page_num": page_index + 1,
            "rect": rect
        })

    return segments


def extract_text_from_segments(doc, segments):
    """
    从题目区域中提取文本。
    """
    lines = []

    for seg in segments:
        page = doc.load_page(seg["page_index"])
        rect = seg["rect"]

        for line in get_page_lines(page):
            line_rect = fitz.Rect(line["x0"], line["y0"], line["x1"], line["y1"])

            if line_rect.intersects(rect):
                lines.append(line["text"])

    return "\n".join(lines).strip()


def split_question_and_options(text):
    """
    拆分文字题题干和 A/B/C/D 四个选项。
    必须识别出完整 A/B/C/D，否则返回空选项，避免出现只有 A/B 的残缺题。
    """
    text = str(text or "").strip()

    if not text:
        return "", []

    matches = list(OPTION_RE.finditer(text))

    if len(matches) < 4:
        return text, []

    start_idx = -1

    for i in range(len(matches) - 3):
        letters = "".join(m.group(1) for m in matches[i:i + 4])
        if letters == "ABCD":
            start_idx = i
            break

    if start_idx == -1:
        return text, []

    selected = matches[start_idx:start_idx + 4]

    question_text = text[:selected[0].start()].strip()
    options = []

    for i, m in enumerate(selected):
        opt_start = m.start()
        opt_end = selected[i + 1].start() if i + 1 < len(selected) else len(text)
        opt = text[opt_start:opt_end].strip()
        opt = re.sub(r"\s+", " ", opt)

        if len(opt) < 2:
            return text, []

        options.append(opt)

    return question_text, options


# =========================
# PDF 图片题截图
# =========================

def save_graph_question_image(doc, segments, image_output_dir, question_no, zoom=2.8):
    """
    图形推理题截图。

    这份 PDF 的图形题排版稳定：
    - 每道题图形和 A/B/C/D 选项都在题目起点到下一题起点之间
    - 所以直接按题目区域截图最稳
    """
    image_paths = []

    for part_index, seg in enumerate(segments, start=1):
        page = doc.load_page(seg["page_index"])
        rect = fitz.Rect(seg["rect"])

        # 左右裁掉页边空白和页眉区域
        clip = fitz.Rect(rect)
        clip.x0 = max(page.rect.x0, clip.x0 + 18)
        clip.x1 = min(page.rect.x1, clip.x1 - 18)
        clip.y0 = max(page.rect.y0, clip.y0 - 4)
        clip.y1 = min(page.rect.y1, clip.y1 + 4)

        if clip.width < 80 or clip.height < 50:
            continue

        filename = f"pdf_q{question_no:04d}_part{part_index}.png"
        filepath = os.path.join(image_output_dir, filename)

        pix = page.get_pixmap(
            matrix=fitz.Matrix(zoom, zoom),
            clip=clip,
            alpha=False
        )
        pix.save(filepath)

        image_paths.append(os.path.join(image_output_dir, filename).replace("\\", "/"))

    return image_paths


# =========================
# PDF 主提取逻辑
# =========================

def extract_pdf_questions(file_path, image_output_dir):
    """
    针对当前精简版判断推理 PDF 的提取逻辑。

    规则：
    1. 图形推理页：
       - 题目区域截图
       - options 为空
       - 前端显示固定 A/B/C/D 按钮
    2. 文字判断 / 类比推理页：
       - 从文本中拆分题干和 A/B/C/D
       - 只有完整四个选项才保留
    """
    print(f"正在解析 PDF: {file_path}")

    os.makedirs(image_output_dir, exist_ok=True)

    doc = fitz.open(file_path)
    starts = find_question_starts(doc)

    print(f"识别到候选题目起点：{len(starts)} 个")

    questions = []
    skipped = 0

    for idx, start in enumerate(starts):
        end = starts[idx + 1] if idx + 1 < len(starts) else None
        segments = get_question_segments(doc, start, end)

        if not segments:
            skipped += 1
            continue

        global_q_no = len(questions) + 1

        raw_text = extract_text_from_segments(doc, segments)

        if not raw_text:
            skipped += 1
            continue

        # 图形推理题：截图为主
        if start["is_graph"]:
            image_paths = save_graph_question_image(
                doc=doc,
                segments=segments,
                image_output_dir=image_output_dir,
                question_no=global_q_no
            )

            if not image_paths:
                skipped += 1
                continue

            item = {
                "source": os.path.basename(file_path),
                "page": start["page_num"],
                "question_no": global_q_no,
                "original_question_no": start["original_question_no"],
                "question_text": "请观察图片中的题干和 A / B / C / D 选项，并点击下方按钮作答。",
                "options": [],
                "images": image_paths,
                "image": image_paths[0],
                "render_mode": "image_question",
                "category": "判断推理"
            }

            questions.append(item)

        # 文字题：解析 A/B/C/D
        else:
            question_text, options = split_question_and_options(raw_text)

            if len(options) < 4:
                skipped += 1
                print(f"跳过第 {start['page_num']} 页附近题目：未识别到完整 A/B/C/D。")
                continue

            item = {
                "source": os.path.basename(file_path),
                "page": start["page_num"],
                "question_no": global_q_no,
                "original_question_no": start["original_question_no"],
                "question_text": question_text,
                "options": options,
                "images": [],
                "category": "判断推理"
            }

            questions.append(item)

    print(f"PDF 成功提取：{len(questions)} 道")
    print(f"PDF 跳过：{skipped} 道")

    return questions


# =========================
# 文件与主函数
# =========================

def first_existing(patterns):
    for pattern in patterns:
        matches = glob.glob(pattern)
        if matches:
            return matches[0]
    return None


def clean_old_pdf_images(image_dir):
    """
    清理旧 PDF 图片，避免旧图残留。
    """
    if not os.path.exists(image_dir):
        return

    for name in os.listdir(image_dir):
        if name.startswith("pdf_q") and name.lower().endswith(".png"):
            try:
                os.remove(os.path.join(image_dir, name))
            except Exception:
                pass


def main():
    parser = argparse.ArgumentParser(description="提取职业能力 DOCX 与判断推理 PDF 题库")
    parser.add_argument("--docx", default=None, help="DOCX 文件路径")
    parser.add_argument("--pdf", default=None, help="PDF 文件路径")
    parser.add_argument("--image-dir", default="images", help="图片输出目录")
    parser.add_argument("--output", default="questions_database.json", help="JSON 输出文件")
    parser.add_argument("--no-clean-images", action="store_true", help="不清理旧 PDF 图片")
    args = parser.parse_args()

    docx_file = args.docx or first_existing([
        "职业能力测试.docx",
        "职业能力测试*.docx"
    ])

    pdf_file = args.pdf or first_existing([
        "判断推理.pdf",
        "判断推理*.pdf"
    ])

    image_dir = args.image_dir
    output_json = args.output

    os.makedirs(image_dir, exist_ok=True)

    if not args.no_clean_images:
        clean_old_pdf_images(image_dir)

    all_questions = []

    if docx_file and os.path.exists(docx_file):
        docx_questions = extract_docx_questions(docx_file)
        all_questions.extend(docx_questions)
        print(f"从 DOCX 中提取：{len(docx_questions)} 道")
    else:
        print("未找到 DOCX 文件，跳过 DOCX 提取。")

    if pdf_file and os.path.exists(pdf_file):
        pdf_questions = extract_pdf_questions(pdf_file, image_dir)
        all_questions.extend(pdf_questions)
        print(f"从 PDF 中提取：{len(pdf_questions)} 道")
    else:
        print("未找到 PDF 文件，跳过 PDF 提取。")

    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(all_questions, f, ensure_ascii=False, indent=4)

    print("\n提取完成")
    print(f"输出 JSON：{output_json}")
    print(f"图片目录：{image_dir}")
    print(f"总题数：{len(all_questions)}")


if __name__ == "__main__":
    main()