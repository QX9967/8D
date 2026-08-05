"""Generate an editable 8D report from D0-D8 folders and a PowerPoint template."""

from __future__ import annotations

import argparse
import base64
import copy
import json
import re
import sys
from pathlib import Path
from typing import Any

import keyring
from json_repair import repair_json
from openai import OpenAI
from PIL import Image
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


BASE_URL = "https://www.packyapi.ai/v1"
MODEL = "MiniMax-M3"
KEYRING_SERVICE = "Adayo8D-AI"
KEYRING_ACCOUNT = "Packy API Key"
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".webp"}
STAGE_CODES = ("d0", "d1", "d2", "d3", "d4", "d5", "d6", "d7", "d8")
STAGE_NAMES = [
    "D0 相关信息", "D1 团队成立", "D2 问题描述", "D3 临时措施", "D4 原因分析",
    "D5 长期对策", "D6 效果验证", "D7 预防措施", "D8 结案总结",
]


def ask_for_key() -> str:
    """Ask once through a Windows dialog; the returned key is saved in Credential Manager."""
    try:
        import tkinter as tk
        from tkinter import simpledialog

        root = tk.Tk()
        root.withdraw()
        value = simpledialog.askstring("Adayo 8D AI", "请输入 Packy API Key（仅首次需要）：", show="*")
        root.destroy()
    except Exception:
        from getpass import getpass

        value = getpass("Packy API Key: ")
    if not value or not value.strip():
        raise RuntimeError("未输入 API Key，已取消。")
    return value.strip()


def get_api_key(reset: bool = False) -> str:
    if reset:
        try:
            keyring.delete_password(KEYRING_SERVICE, KEYRING_ACCOUNT)
        except keyring.errors.PasswordDeleteError:
            pass
    value = keyring.get_password(KEYRING_SERVICE, KEYRING_ACCOUNT)
    if not value:
        value = ask_for_key()
        keyring.set_password(KEYRING_SERVICE, KEYRING_ACCOUNT, value)
    return value


def read_text(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "gb18030", "utf-16"):
        try:
            return raw.decode(encoding).strip()
        except UnicodeDecodeError:
            pass
    return raw.decode("utf-8", errors="replace").strip()


def collect_materials(materials_root: Path) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    folders = [item for item in materials_root.iterdir() if item.is_dir()]
    for stage_index, stage_name in enumerate(STAGE_NAMES):
        folder = next(
            (item for item in folders if item.name.lower().startswith(f"d{stage_index}") or item.name.startswith(f"{stage_index:02}")),
            None,
        )
        if folder is None:
            result[stage_name] = {"notes": "", "images": []}
            continue
        notes = "\n\n".join(read_text(item) for item in folder.rglob("*.txt") if item.stat().st_size)
        images = sorted(item for item in folder.rglob("*") if item.suffix.lower() in IMAGE_EXTENSIONS)
        result[stage_name] = {"notes": notes, "images": images}
    return result


def executable_folder() -> Path:
    return Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent


def bundled_template() -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / "8D模板.pptx"


def pause(message: str = "\nPress Enter to close...") -> None:
    if getattr(sys, "frozen", False):
        try:
            input(message)
        except EOFError:
            pass


def image_part(path: Path) -> dict[str, Any]:
    mime = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".bmp": "image/bmp", ".webp": "image/webp"}[path.suffix.lower()]
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{encoded}"}}


def build_prompt(materials: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    content: list[dict[str, Any]] = [{"type": "text", "text": "以下为 D0-D8 分组材料。TXT 是人员记录，图片是证据。请严格只依据材料输出。"}]
    for stage, payload in materials.items():
        content.append({"type": "text", "text": f"\n【{stage}】\n文字记录：\n{payload['notes'] or '（无）'}\n图片文件：{', '.join(image.name for image in payload['images']) or '（无）'}"})
        for image in payload["images"]:
            content.append(image_part(image))
    return content


SYSTEM_PROMPT = r"""
你是汽车电子质量 8D 报告助手。基于用户给出的分组文字和现场证据图片，输出单个 JSON 对象，不要 Markdown，不要解释。

原则：
1. 不得编造日期、数据、责任人、验证结果、根因或措施。没有证据的值写“待确认”。
2. 图片只能支持可直接观察到的现象；不能凭图片断言真实根因。
3. 使用正式、简洁的中文。每项最多两句。team 最多 6 人；actions 最多 7 项；countermeasures、validations 各最多 5 项。
4. evidence_summary 是每个阶段图片想证明的简短说明；没有图片则为空字符串。

按以下 JSON 结构输出，所有键必须存在：
{
 "cover":{"title":"","prepared_by":"","reviewed_by":"","approved_by":"","date":""},
 "d0":{"fault_date":"","model":"","trace_code":"","station":"","description":"","summary":"","evidence_summary":""},
 "d1":{"team":[{"name":"","role":"","duty":"","contact":"","module":""}],"summary":"","evidence_summary":""},
 "d2":{"customer":"","date":"","supplier":"","model":"","vehicle_model":"","material":"","quantity":"","failure_type":"","lot":"","location":"","production_date":"","vin":"","complaint_source":"","symptom":"","confirmed_symptom":"","summary":"","evidence_summary":""},
 "d3":{"actions":[{"category":"","action":"","owner":"","date":"","status":""}],"summary":"","evidence_summary":""},
 "d4":{"five_whys":[{"why":"","answer":""}],"root_cause":"","escape_cause":"","summary":"","evidence_summary":""},
 "d5":{"countermeasures":[{"action":"","owner":"","date":"","status":""}],"summary":"","evidence_summary":""},
 "d6":{"validations":[{"method":"","result":"","owner":"","date":""}],"summary":"","evidence_summary":""},
 "d7":{"preventions":[{"item":"","yes_no":"","comment":"","owner":"","date":"","status":""}],"summary":"","evidence_summary":""},
 "d8":{"conclusion":"","summary":"","evidence_summary":""}
}
""".strip()
SYSTEM_PROMPT += """

Additionally include a root-level image_pages object with d2, d4, d5, d6, d7 and d8 arrays.
Each page must contain title, summary, layout (single, two, three or four), and images (exact image file names).
Use single for screenshots, comparison images, or images that require inspection individually. Only group images that show one same activity.
Every supplied image name must appear exactly once in its own D-stage list.
"""
SYSTEM_PROMPT += """

For d4, return occurrence_why_rows with exactly five ordered rows (Why1 to Why5), and escape_why_rows with exactly three ordered rows (Why1 to Why3).
Every row must contain level, problem and cause. `cause` is the answer/reason for that Why level, never a rewording of the question. The first `problem` is the original fault; later rows' `problem` should be the previous row's cause, so the chain is continuous.
The PPT has three columns: level, problem and cause. Put the answer in `cause`, not only in the `whys` list. Do not make up unsupported facts: if a level genuinely has no evidence, leave that row's cause empty.
"""


def parse_json(text: str) -> dict[str, Any]:
    cleaned = text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", cleaned, flags=re.S)
    match = re.search(r"\{.*\}", cleaned, flags=re.S)
    if not match:
        raise ValueError("模型没有返回 JSON 对象。")
    candidate = match.group(0)
    try:
        return json.loads(candidate)
    except json.JSONDecodeError:
        repaired = repair_json(candidate, return_objects=True)
        if not isinstance(repaired, dict):
            raise ValueError("模型返回内容无法修复为 JSON 对象。")
        return repaired


def generate_content(materials: dict[str, dict[str, Any]], api_key: str) -> dict[str, Any]:
    client = OpenAI(base_url=BASE_URL, api_key=api_key)
    response = client.chat.completions.create(
        model=MODEL,
        temperature=0.1,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": build_prompt(materials)},
        ],
    )
    text = response.choices[0].message.content or ""
    return parse_json(text)


def text(value: Any) -> str:
    return str(value or "待确认").strip()


def set_run_font(run: Any, size: int = 10) -> None:
    """Set Latin and East-Asian font metadata; PowerPoint otherwise falls back to SimSun."""
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(size)
    rpr = run._r.get_or_add_rPr()
    rpr.set("latin", "Microsoft YaHei")
    rpr.set("ea", "Microsoft YaHei")
    rpr.set("cs", "Microsoft YaHei")


def put_cell(table: Any, row: int, col: int, value: Any) -> None:
    cell = table.cell(row, col)
    cell.text = "" if value in (None, "", "待确认") else str(value).strip()
    for paragraph in cell.text_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.LEFT
        for run in paragraph.runs:
            set_run_font(run)


def first_table(slide: Any) -> Any:
    table_shape = next((shape for shape in slide.shapes if shape.has_table), None)
    if table_shape is None:
        raise ValueError("模板页面未找到可填写的原生表格。")
    return table_shape


def replace_tokens(shape: Any, replacements: dict[str, Any]) -> None:
    if not shape.has_text_frame:
        return
    original = shape.text
    updated = original
    for token, value in replacements.items():
        updated = updated.replace(token, text(value))
    if updated != original:
        shape.text = updated
        shape.name = "AI_DYNAMIC"


def set_paragraph_text(paragraph: Any, value: Any) -> None:
    """Replace a paragraph while retaining a deterministic Chinese font setting."""
    paragraph.text = "" if value in (None, "", "待确认") else str(value).strip()
    for run in paragraph.runs:
        set_run_font(run)


def add_text(slide: Any, value: str, left: float, top: float, width: float, height: float, size: int = 10) -> None:
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    box.name = "AI_DYNAMIC"
    frame = box.text_frame
    frame.word_wrap = True
    frame.text = value
    for paragraph in frame.paragraphs:
        for run in paragraph.runs:
            set_run_font(run, size)


def crop_picture(slide: Any, image: Path, left: float, top: float, width: float, height: float) -> None:
    if left < 0 or top < 0 or left + width > 10.0 or top + height > 5.625:
        raise ValueError(f"image layout exceeds slide bounds: {image.name}")
    with Image.open(image) as source:
        image_ratio = source.width / source.height
    frame_ratio = width / height
    if image_ratio > frame_ratio:
        render_width = width
        render_height = width / image_ratio
        render_left = left
        render_top = top + (height - render_height) / 2
    else:
        render_height = height
        render_width = height * image_ratio
        render_left = left + (width - render_width) / 2
        render_top = top
    slide.shapes.add_picture(str(image), Inches(render_left), Inches(render_top), width=Inches(render_width), height=Inches(render_height))


def clone_slide(presentation: Presentation, source_index: int) -> Any:
    source = presentation.slides[source_index]
    return clone_slide_from(presentation, source)


def clone_slide_from(presentation: Presentation, source: Any) -> Any:
    target = presentation.slides.add_slide(source.slide_layout)
    for shape in source.shapes:
        target.shapes._spTree.insert_element_before(copy.deepcopy(shape.element), "p:extLst")
    return target


def remove_shape(shape: Any) -> None:
    shape.element.getparent().remove(shape.element)


def move_slide_after(presentation: Presentation, slide: Any, after_index: int) -> None:
    slide_ids = presentation.slides._sldIdLst
    slide_id = slide_ids[-1]
    slide_ids.remove(slide_id)
    slide_ids.insert(after_index + 1, slide_id)


def move_slide_before(presentation: Presentation, slide: Any, before_index: int) -> None:
    slide_ids = presentation.slides._sldIdLst
    slide_id = slide_ids[-1]
    slide_ids.remove(slide_id)
    slide_ids.insert(before_index, slide_id)


def expand_team_table(shape: Any, member_count: int, slide_height: int) -> None:
    """Extend the native PowerPoint table instead of moving extra team members outside it."""
    table = shape.table
    wanted_rows = max(2, member_count)
    existing_rows = len(table.rows) - 1
    if wanted_rows <= existing_rows:
        return
    row_height = table.rows[len(table.rows) - 1].height
    for _ in range(wanted_rows - existing_rows):
        table._tbl.append(copy.deepcopy(table.rows[len(table.rows) - 1]._tr))
    maximum_height = slide_height - shape.top - Inches(0.55)
    shape.height = min(shape.height + row_height * (wanted_rows - existing_rows), maximum_height)


def image_slots(count: int) -> list[tuple[float, float, float, float]]:
    if count == 1:
        return [(0.55, 1.15, 8.90, 3.65)]
    if count == 2:
        return [(0.55, 1.15, 4.25, 3.65), (5.15, 1.15, 4.25, 3.65)]
    if count == 3:
        return [(0.55, 1.15, 5.35, 3.65), (6.15, 1.15, 3.25, 1.70), (6.15, 3.10, 3.25, 1.70)]
    return [(0.55, 1.15, 4.25, 1.70), (5.15, 1.15, 4.25, 1.70), (0.55, 3.10, 4.25, 1.70), (5.15, 3.10, 4.25, 1.70)]


def plan_image_pages(images: list[Path], plans: list[dict[str, Any]], default_summary: str) -> list[dict[str, Any]]:
    """Honor model-selected grouping. Unassigned images remain one image per page."""
    by_name = {image.name: image for image in images}
    used: set[str] = set()
    result: list[dict[str, Any]] = []
    maximum = {"single": 1, "two": 2, "three": 3, "four": 4}
    for plan in plans:
        layout = str(plan.get("layout", "single")).lower()
        selected = [by_name[name] for name in plan.get("images", []) if name in by_name and name not in used]
        if not selected:
            continue
        limit = maximum.get(layout, 1)
        while selected:
            group, selected = selected[:limit], selected[limit:]
            used.update(image.name for image in group)
            result.append({"images": group, "title": text(plan.get("title")), "summary": text(plan.get("summary") or default_summary)})
    for image in images:
        if image.name not in used:
            result.append({"images": [image], "title": "证据材料", "summary": default_summary})
    return result


def insert_evidence_pages(presentation: Presentation, source_index: int, stage: str, pages: list[dict[str, Any]], before_source: bool = False) -> None:
    insertion_index = source_index
    source = presentation.slides[source_index]
    for page_number, page in enumerate(pages, start=1):
        slide = clone_slide_from(presentation, source)
        for shape in list(slide.shapes):
            if shape.has_table or shape.name == "AI_DYNAMIC" or (shape.has_text_frame and "{{Content}}" in shape.text):
                remove_shape(shape)
        add_text(slide, page["summary"], 0.55, 5.03, 8.90, 0.25)
        for image, slot in zip(page["images"], image_slots(len(page["images"]))):
            crop_picture(slide, image, *slot)
        if before_source:
            move_slide_before(presentation, slide, insertion_index)
        else:
            move_slide_after(presentation, slide, insertion_index)
        insertion_index += 1


def fill_template(template: Path, output: Path, data: dict[str, Any], materials: dict[str, dict[str, Any]]) -> None:
    presentation = Presentation(str(template))
    if len(presentation.slides) < 11:
        raise ValueError("模板页数不足，需包含封面和 D0-D8 页面。")

    cover, d0, d1, d2 = data["cover"], data["d0"], data["d1"], data["d2"]
    slide = presentation.slides[0]
    # The title still supports {{Content}}, while cover approvals can also be ordinary
    # labels such as "编制：" after a customer has customized the template.
    for shape in slide.shapes:
        replace_tokens(shape, {"{{Content}}": cover.get("title", "")})
        if not shape.has_text_frame:
            continue
        label_values = (("编制", cover.get("prepared_by")), ("审核", cover.get("reviewed_by")),
                        ("批准", cover.get("approved_by")), ("日期", cover.get("date")))
        for paragraph in shape.text_frame.paragraphs:
            source = paragraph.text.replace(" ", "")
            for label, value in label_values:
                if label in source:
                    display = "" if value in (None, "", "待确认") else str(value).strip()
                    separator = "：" if "：" in paragraph.text or ":" not in paragraph.text else ":"
                    set_paragraph_text(paragraph, f"{label}{separator}{display}")
                    shape.name = "AI_DYNAMIC"
                    break

    slide = presentation.slides[2]
    d0_lines = [d0["fault_date"], d0["model"], d0["trace_code"], d0["station"], d0["description"]]
    for index, paragraph in enumerate(slide.shapes[1].text_frame.paragraphs):
        if index < len(d0_lines) and "{{Content}}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{{Content}}", text(d0_lines[index]))

    team_shape = first_table(presentation.slides[3])
    expand_team_table(team_shape, len(d1.get("team", [])), presentation.slide_height)
    team_table = team_shape.table
    for row, member in enumerate(d1.get("team", []), start=1):
        put_cell(team_table, row, 0, row)
        for col, key in enumerate(("name", "role", "duty", "contact", "module"), start=1):
            put_cell(team_table, row, col, member.get(key))

    d2_table = first_table(presentation.slides[4]).table
    d2_map = {
        (0, 1): d2["customer"], (0, 3): d2["date"], (0, 5): d2["supplier"],
        (1, 1): d2["model"], (1, 3): d2["vehicle_model"], (1, 5): d2["material"],
        (2, 1): d2["quantity"], (2, 3): d2["failure_type"], (2, 5): d2["lot"],
        (3, 1): d2["location"], (3, 3): d2["production_date"], (3, 5): d2["vin"],
        (4, 1): d2["complaint_source"], (5, 1): d2["symptom"], (6, 1): d2["confirmed_symptom"],
    }
    for (row, col), value in d2_map.items():
        put_cell(d2_table, row, col, value)

    d3_table = first_table(presentation.slides[5]).table
    by_category = {text(item.get("category")): item for item in data["d3"].get("actions", [])}
    for row in range(1, 8):
        category = d3_table.cell(row, 1).text.strip()
        item = by_category.get(category, {})
        for col, key in ((2, "action"), (3, "owner"), (4, "date"), (5, "status")):
            put_cell(d3_table, row, col, item.get(key, "待确认"))

    # Fill the two D4 5WHY tables from model-generated occurrence and escape reason chains.
    d4 = data.get("d4", {})
    def reason_chain(kind: str) -> list[tuple[str, str]]:
        """Return the problem-and-cause pair for every individual WHY row."""
        direct = d4.get(f"{kind}_why", {})
        problem = str(direct.get("problem", "")).strip()
        values = [str(value).strip() for value in direct.get("whys", []) if str(value or "").strip()]
        rows = d4.get(f"{kind}_why_rows", [])
        if rows:
            return [
                (str(row.get("problem", "")).strip(), str(row.get("cause", "")).strip())
                for row in rows
            ]
        if not values:
            values = [str(item.get("answer", "")).strip() for item in d4.get("five_whys", []) if str(item.get("answer", "")).strip()]
        return [(problem if index == 0 else "", value) for index, value in enumerate(values)]

    occurrence_table = first_table(presentation.slides[7]).table
    occurrence_rows = reason_chain("occurrence")
    for row_index, (problem, cause) in enumerate(occurrence_rows[:len(occurrence_table.rows) - 1], start=1):
        put_cell(occurrence_table, row_index, 1, problem)
        put_cell(occurrence_table, row_index, 2, cause)
    escape_table = first_table(presentation.slides[8]).table
    escape_rows = reason_chain("escape")
    for row_index, (problem, cause) in enumerate(escape_rows[:len(escape_table.rows) - 1], start=1):
        put_cell(escape_table, row_index, 1, problem)
        put_cell(escape_table, row_index, 2, cause)

    d5_table = first_table(presentation.slides[9]).table
    d5_lines = [f"{index + 1}. {text(x.get('action'))}（{text(x.get('owner'))}，{text(x.get('date'))}，{text(x.get('status'))}）" for index, x in enumerate(data["d5"].get("countermeasures", []))]
    put_cell(d5_table, 1, 0, "1")
    put_cell(d5_table, 1, 1, "\n".join(d5_lines))
    put_cell(d5_table, 1, 2, "待确认")
    put_cell(d5_table, 1, 3, "待确认")
    put_cell(d5_table, 1, 4, "待确认")

    d6_table = first_table(presentation.slides[10]).table
    d6_lines = [f"{index + 1}. {text(x.get('method'))}" for index, x in enumerate(data["d6"].get("validations", []))]
    result_lines = [f"{index + 1}. {text(x.get('result'))}" for index, x in enumerate(data["d6"].get("validations", []))]
    put_cell(d6_table, 1, 0, "1")
    put_cell(d6_table, 1, 1, "\n".join(d6_lines))
    put_cell(d6_table, 1, 2, "\n".join(result_lines))
    put_cell(d6_table, 1, 3, "待确认")
    put_cell(d6_table, 1, 4, "待确认")

    d7_table = first_table(presentation.slides[11]).table
    prevention_map = {text(item.get("item")): item for item in data["d7"].get("preventions", [])}
    for row in range(1, len(d7_table.rows)):
        item = prevention_map.get(d7_table.cell(row, 0).text.strip(), {})
        for col, key in ((1, "yes_no"), (2, "comment"), (3, "owner"), (4, "date"), (5, "status")):
            # Do not touch unmatched or empty cells: this preserves the customer's
            # original row heights, widths, borders and default formatting.
            value = item.get(key) if item else ""
            if value not in (None, "", "待确认"):
                put_cell(d7_table, row, col, value)

    replace_tokens(presentation.slides[12].shapes[1], {"{{Content}}": data["d8"].get("conclusion")})

    # Insert stage evidence immediately after its own slide. Descending order keeps source indexes stable.
    stage_sources = (("D8 结案总结", "d8", 12), ("D7 预防措施", "d7", 11), ("D6 效果验证", "d6", 10), ("D5 长期对策", "d5", 9), ("D4 原因分析", "d4", 6), ("D2 问题描述", "d2", 4))
    page_data = data.get("image_pages", {})
    for stage, key, source_index in stage_sources:
        images = materials[stage]["images"]
        if images:
            pages = plan_image_pages(images, page_data.get(key, []), data.get(key, {}).get("evidence_summary", ""))
            insert_evidence_pages(presentation, source_index, stage, pages)

    output.parent.mkdir(parents=True, exist_ok=True)
    presentation.save(str(output))


def main() -> int:
    parser = argparse.ArgumentParser(description="Adayo 8D AI 报告生成器")
    parser.add_argument("--template", type=Path, help="8D 模板 .pptx")
    parser.add_argument("--materials", type=Path, help="含 D0-D8 子文件夹的材料根目录")
    parser.add_argument("--output", type=Path, help="输出 .pptx；省略时写入 EXE 同级目录")
    parser.add_argument("--reset-key", action="store_true", help="删除凭据管理器中保存的 API Key")
    parser.add_argument("--draft-json", type=Path, help="跳过模型调用，使用已有 JSON 草稿生成 PPT")
    args = parser.parse_args()
    auto_mode = not any((args.template, args.materials, args.output, args.draft_json, args.reset_key))
    try:
        if args.reset_key:
            get_api_key(reset=True)
            print("旧密钥已清除，新密钥已保存。")
            return 0

        # Double-click mode: the EXE is placed directly in the material folder.
        if auto_mode:
            args.template = bundled_template()
            args.materials = executable_folder()
            args.output = executable_folder() / "8D报告_AI生成.pptx"
        else:
            if not args.template or not args.materials:
                parser.error("必须同时提供 --template 和 --materials。")
            if args.output is None:
                args.output = executable_folder() / "8D报告_AI生成.pptx"

        if not args.template.exists():
            raise FileNotFoundError(f"内置模板不存在：{args.template}")
        print("[1/4] 正在读取 00-08（或 D0-D8）材料...")
        materials = collect_materials(args.materials)
        found = sum(len(item["images"]) for item in materials.values())
        if found == 0:
            raise FileNotFoundError("未找到 00-08 或 D0-D8 文件夹中的图片材料。")
        stage_counts = ", ".join(f"{stage[:2]}={len(item['images'])}" for stage, item in materials.items())
        print(f"      已发现 {found} 张图片：{stage_counts}")
        if args.draft_json:
            print("[2/4] 正在读取本地 AI 草稿...")
            data = json.loads(args.draft_json.read_text(encoding="utf-8"))
        else:
            print("[2/4] 正在上传图片并由 MiniMax 生成 8D 内容；图片较多时此步骤需要等待...")
            data = generate_content(materials, get_api_key())
            print("[3/4] AI 内容已返回，正在保存 AI 草稿...")
            args.output.with_suffix(".json").write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        print("[4/4] 正在写入表格、文字并插入图片，请稍候...")
        fill_template(args.template, args.output, data, materials)
        print(f"已生成可编辑 PPT：{args.output.resolve()}")
        return 0
    except Exception as exc:
        print(f"生成失败：{exc}")
        return 1
    finally:
        if auto_mode or getattr(sys, "frozen", False):
            pause()


if __name__ == "__main__":
    sys.exit(main())
