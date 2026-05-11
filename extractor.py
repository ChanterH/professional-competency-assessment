import os
import json
import re
import glob
import argparse

try:
    import docx
except ImportError as e:
    raise ImportError("缺少 python-docx 库，请运行：python -m pip install python-docx") from e

import fitz


DOCX_QUESTION_RE = re.compile(r'^\s*\d+[\.、．\s]')
DOCX_OPTION_RE = re.compile(r'^\s*[A-D][\.、．\s]')
PDF_QUESTION_RE = re.compile(r'^\s*(\d{1,3})\s*[\.、．]\s*')


def extract_docx_questions(file_path):
    """
    解析 Word 文档中的职业能力题。
    保留原始 DOCX 提取逻辑。
    """
    print(f"正在解析: {file_path}")

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


def _clean_pdf_text(text):
    lines = []

    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue

        if line == "超格学员专用":
            continue

        if re.fullmatch(r"\d+", line):
            continue

        if re.fullmatch(r"第[一二三四五六七八九十]+套(?:（难度[★☆]+）)?", line):
            continue

        lines.append(line)

    return "\n".join(lines).strip()


def _get_pdf_question_starts(page):
    """
    获取 PDF 当前页中每道题的起始位置。
    """
    starts = []

    for block in page.get_text("blocks"):
        x0, y0, x1, y1, text = block[:5]

        if y1 < 55 or y0 > page.rect.height - 45:
            continue

        text = text.strip()
        if not text:
            continue

        first_line = text.splitlines()[0].strip()
        m = PDF_QUESTION_RE.match(first_line)

        if m:
            starts.append({
                "no": int(m.group(1)),
                "x": x0,
                "y": y0,
                "text": first_line
            })

    dedup = {}
    for item in sorted(starts, key=lambda d: (d["y"], d["x"])):
        if item["no"] not in dedup:
            dedup[item["no"]] = item

    return sorted(dedup.values(), key=lambda d: d["y"])


def _get_text_lines_in_rect(page, rect):
    items = []

    for block in page.get_text("blocks"):
        x0, y0, x1, y1, text = block[:5]
        block_rect = fitz.Rect(x0, y0, x1, y1)

        if not block_rect.intersects(rect):
            continue

        if y1 < 55 or y0 > page.rect.height - 45:
            continue

        clean_text = _clean_pdf_text(text)
        if clean_text:
            items.append((y0, x0, clean_text))

    items.sort(key=lambda x: (round(x[0], 1), x[1]))
    return [x[2] for x in items]


def _split_question_and_options(lines):
    """
    拆分 PDF 文本题干和文字选项。
    只用于没有图片的纯文字题。
    """
    text = "\n".join(lines).strip()
    if not text:
        return "", []

    option_pattern = re.compile(r'(?<![A-Za-z])([A-D])[\.\、．]\s*')
    matches = list(option_pattern.finditer(text))

    if len(matches) < 4:
        return text, []

    first_four = matches[:4]
    letters = "".join(m.group(1) for m in first_four)

    if letters != "ABCD":
        return text, []

    question_text = text[:first_four[0].start()].strip()
    options = []

    for i, m in enumerate(first_four):
        start = m.start()
        end = first_four[i + 1].start() if i + 1 < len(first_four) else len(text)
        opt = text[start:end].strip()
        options.append(opt)

    return question_text, options


def _get_image_rects(page, question_rect):
    """
    获取题目区域内的图片对象位置。
    """
    rects = []
    seen = set()

    for img in page.get_images(full=True):
        xref = img[0]

        try:
            img_rects = page.get_image_rects(xref)
        except Exception:
            continue

        for rect in img_rects:
            if rect.width < 20 or rect.height < 20:
                continue

            center_y = (rect.y0 + rect.y1) / 2
            if not (question_rect.y0 <= center_y <= question_rect.y1):
                continue

            if not rect.intersects(question_rect):
                continue

            key = (
                round(rect.x0, 1),
                round(rect.y0, 1),
                round(rect.x1, 1),
                round(rect.y1, 1)
            )

            if key not in seen:
                seen.add(key)
                rects.append(rect)

    return rects


def _get_drawing_rects(page, question_rect):
    """
    获取题目区域内的矢量图形 / 表格线条位置。
    用于处理 PDF 中不是图片对象的图形题、表格题。
    """
    rects = []

    try:
        drawings = page.get_drawings()
    except Exception:
        return rects

    for drawing in drawings:
        rect = drawing.get("rect")
        if not rect:
            continue

        rect = fitz.Rect(rect)

        if rect.width < 15 or rect.height < 8:
            continue

        center_y = (rect.y0 + rect.y1) / 2
        if not (question_rect.y0 <= center_y <= question_rect.y1):
            continue

        if not rect.intersects(question_rect):
            continue

        rects.append(rect)

    return rects


def _merge_rects(rects):
    if not rects:
        return None

    merged = fitz.Rect(rects[0])
    for rect in rects[1:]:
        merged |= rect

    return merged


def _save_visual_region(page, rects, image_output_dir, page_num, q_no, zoom=3.0, padding=10):
    """
    只保存题目中的图形 / 表格 / 图片区域。
    不保存整页，不保存页码，不保存上一题残留。
    """
    merged = _merge_rects(rects)
    if not merged:
        return None

    clip = fitz.Rect(merged)

    clip.x0 = max(page.rect.x0, clip.x0 - padding)
    clip.y0 = max(page.rect.y0, clip.y0 - padding)
    clip.x1 = min(page.rect.x1, clip.x1 + padding)
    clip.y1 = min(page.rect.y1, clip.y1 + padding)

    filename = f"pdf_page{page_num:03d}_q{q_no:03d}.png"
    filepath = os.path.join(image_output_dir, filename)

    pix = page.get_pixmap(
        matrix=fitz.Matrix(zoom, zoom),
        clip=clip,
        alpha=False
    )

    pix.save(filepath)
    return filepath


def extract_pdf_questions(file_path, image_output_dir):
    """
    解析 PDF 判断推理题。

    修复原则：
    1. 不再整道题截图，避免把页码、上一题残留、无关内容截进去。
    2. 只截取题目区域内的图片 / 图形 / 表格部分。
    3. 有图形的 PDF 题统一作为图片题处理，options 留空，前端显示 A/B/C/D 按钮。
    4. 没有图片、没有完整选项的 PDF 残缺题直接跳过。
    """
    print(f"正在解析: {file_path}")

    os.makedirs(image_output_dir, exist_ok=True)

    doc = fitz.open(file_path)
    questions = []

    for page_index in range(len(doc)):
        page = doc.load_page(page_index)
        page_num = page_index + 1

        starts = _get_pdf_question_starts(page)
        if not starts:
            continue

        for i, start in enumerate(starts):
            q_no = start["no"]

            y0 = max(55, start["y"])

            if i + 1 < len(starts):
                y1 = starts[i + 1]["y"] - 4
            else:
                y1 = page.rect.height - 55

            if y1 <= y0:
                continue

            question_rect = fitz.Rect(0, y0, page.rect.width, y1)

            text_lines = _get_text_lines_in_rect(page, question_rect)
            question_text, text_options = _split_question_and_options(text_lines)

            image_rects = _get_image_rects(page, question_rect)
            drawing_rects = _get_drawing_rects(page, question_rect)

            visual_rects = image_rects + drawing_rects

            image_path = None
            if visual_rects:
                image_path = _save_visual_region(
                    page=page,
                    rects=visual_rects,
                    image_output_dir=image_output_dir,
                    page_num=page_num,
                    q_no=q_no
                )

            if image_path:
                item = {
                    "source": os.path.basename(file_path),
                    "page": page_num,
                    "question_no": q_no,
                    "question_text": question_text,
                    "options": [],
                    "images": [image_path],
                    "image": image_path,
                    "render_mode": "image_question"
                }

                questions.append(item)

            else:
                if len(text_options) >= 4:
                    item = {
                        "source": os.path.basename(file_path),
                        "page": page_num,
                        "question_no": q_no,
                        "question_text": question_text,
                        "options": text_options,
                        "images": []
                    }

                    questions.append(item)

                else:
                    print(f"跳过 PDF 第 {page_num} 页第 {q_no} 题：没有图片，也没有完整选项。")

    return questions


def _first_existing(patterns):
    for pattern in patterns:
        matches = glob.glob(pattern)
        if matches:
            return matches[0]
    return None


def main():
    parser = argparse.ArgumentParser(description="提取 DOCX / PDF 题库数据")
    parser.add_argument("--docx", default=None, help="DOCX 文件路径")
    parser.add_argument("--pdf", default=None, help="PDF 文件路径")
    parser.add_argument("--image-dir", default="images", help="图片输出目录")
    parser.add_argument("--output", default="questions_database.json", help="JSON 输出文件")
    args = parser.parse_args()

    docx_file = args.docx or _first_existing([
        "职业能力测试.docx",
        "职业能力测试*.docx"
    ])

    pdf_file = args.pdf or _first_existing([
        "判断推理.pdf",
        "判断推理*.pdf"
    ])

    image_dir = args.image_dir
    output_json = args.output

    os.makedirs(image_dir, exist_ok=True)

    all_questions = []

    if docx_file and os.path.exists(docx_file):
        docx_questions = extract_docx_questions(docx_file)
        all_questions.extend(docx_questions)
        print(f"从 DOCX 中提取了 {len(docx_questions)} 道题目。")
    else:
        print("未找到 DOCX 文件，跳过 DOCX 提取。")

    if pdf_file and os.path.exists(pdf_file):
        pdf_questions = extract_pdf_questions(pdf_file, image_dir)
        all_questions.extend(pdf_questions)
        print(f"从 PDF 中提取了 {len(pdf_questions)} 道题目。")
    else:
        print("未找到 PDF 文件，跳过 PDF 提取。")

    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(all_questions, f, ensure_ascii=False, indent=4)

    print(f"\n所有数据已成功导出至 {output_json}")
    print(f"图片已保存到：{image_dir}")


if __name__ == "__main__":
    main()