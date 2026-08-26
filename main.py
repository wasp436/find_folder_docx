import os
from datetime import date

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff"}


def output_basename(label):
    return f"{date.today():%Y-%m-%d}_{label}"

CJK_FONT_CANDIDATES = [
    r"C:\Windows\Fonts\msjh.ttc",
    r"C:\Windows\Fonts\mingliu.ttc",
    r"C:\Windows\Fonts\kaiu.ttf",
    "/System/Library/Fonts/PingFang.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/arphic/uming.ttc",
]


def find_cjk_font():
    for path in CJK_FONT_CANDIDATES:
        if os.path.isfile(path):
            return path
    return None


def find_thumbs_files(root):
    found = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if not d.startswith(".")]
        for f in filenames:
            if f.lower() == "thumbs.db":
                found.append(os.path.join(dirpath, f))
    return found


def find_empty_dirs(root):
    empty = []
    for dirpath, dirnames, filenames in os.walk(root, topdown=True):
        dirnames[:] = [d for d in dirnames if not d.startswith(".")]
        if dirpath == root:
            continue
        if not dirnames and not filenames:
            empty.append(dirpath)
    return empty


def find_docx_only_dirs(root):
    result = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if not d.startswith(".")]

        is_leaf = len(dirnames) == 0
        if not is_leaf:
            continue

        has_docx = any(f.lower().endswith(".docx") for f in filenames)
        if not has_docx:
            continue

        has_image = any(
            os.path.splitext(f)[1].lower() in IMAGE_EXTENSIONS for f in filenames
        )
        if not has_image:
            result.append(dirpath)
    return result


def rel_path(root, path):
    return os.path.join(".", os.path.relpath(path, root)).replace("\\", "/")


def confirm_delete(prompt):
    ans = input(prompt).strip().lower()
    return ans == "y"


def generate_missing_docx_image(root, dirs_missing_docx, label):
    if not dirs_missing_docx:
        return None

    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        print("⚠️ 未安裝 Pillow，略過圖片產生 (pip install pillow)")
        return None

    font_reg_path = find_cjk_font()
    if not font_reg_path:
        print("⚠️ 找不到可用的中文字型，略過圖片產生")
        return None

    names = [os.path.relpath(d, root).replace("\\", "/") for d in dirs_missing_docx]
    root_parts = [p for p in root.replace("\\", "/").split("/") if p]
    root_display = "/".join(root_parts[-2:])

    width = 720
    pad = 36
    line_h = 56
    path_line_h = 28

    bg = (17, 20, 24)
    card = (26, 30, 36)
    white = (235, 238, 242)
    gray = (150, 158, 168)

    dummy_img = Image.new("RGB", (10, 10))
    dummy_draw = ImageDraw.Draw(dummy_img)

    f_path = ImageFont.truetype(font_reg_path, 20)
    f_sub = ImageFont.truetype(font_reg_path, 24)
    f_item = ImageFont.truetype(font_reg_path, 26)

    max_text_width = width - pad * 2
    path_lines = []
    current = ""
    for ch in root_display:
        candidate = current + ch
        if not current or dummy_draw.textlength(candidate, font=f_path) <= max_text_width:
            current = candidate
        else:
            path_lines.append(current)
            current = ch
    if current:
        path_lines.append(current)

    header_h = pad + len(path_lines) * path_line_h + 16 + 46
    footer_h = 20
    height = header_h + line_h * len(names) + footer_h

    img = Image.new("RGB", (width, height), bg)
    draw = ImageDraw.Draw(img)

    y = pad
    for line in path_lines:
        draw.text((pad, y), line, font=f_path, fill=gray)
        y += path_line_h
    y += 16
    draw.text((pad, y), f"共 {len(names)} 個", font=f_sub, fill=gray)
    y += 46

    for name in names:
        draw.rounded_rectangle(
            [pad, y, width - pad, y + line_h - 10], radius=10, fill=card
        )
        draw.text((pad + 18, y + 9), name, font=f_item, fill=white)
        y += line_h

    out_path = os.path.join(root, f"{output_basename(label)}.png")
    img.save(out_path)
    return out_path


def generate_missing_docx_ods(root, dirs_missing_docx, label, sheet_name):
    if not dirs_missing_docx:
        return None

    try:
        from pathlib import Path

        from odf.opendocument import OpenDocumentSpreadsheet
        from odf.style import Style, TableColumnProperties
        from odf.table import Table, TableCell, TableColumn, TableRow
        from odf.text import P
    except ImportError:
        print("⚠️ 未安裝 odfpy，略過 ods 產生 (pip install odfpy)")
        return None

    def formula_literal(text):
        return '"' + text.replace('"', '""') + '"'

    def char_width(text):
        return sum(2 if ord(ch) > 0x2E80 else 1 for ch in text)

    def column_width_cm(texts, min_cm=1.5, max_cm=25.0):
        max_chars = max((char_width(t) for t in texts), default=0)
        return max(min_cm, min(max_cm, max_chars * 0.19 + 0.5))

    rows = []
    for dirpath in dirs_missing_docx:
        rel_parts = os.path.relpath(dirpath, root).replace("\\", "/").split("/")
        top_name = rel_parts[0]
        leaf_name = rel_parts[-1]
        uri = Path(dirpath).resolve().as_uri()
        if not uri.endswith("/"):
            uri += "/"
        rows.append((top_name, leaf_name, uri))

    doc = OpenDocumentSpreadsheet()

    name_col_style = Style(name="NameCol", family="table-column")
    name_col_style.addElement(
        TableColumnProperties(
            columnwidth=f"{column_width_cm(['名稱'] + [r[0] for r in rows]):.2f}cm"
        )
    )
    doc.automaticstyles.addElement(name_col_style)

    path_col_style = Style(name="PathCol", family="table-column")
    path_col_style.addElement(
        TableColumnProperties(
            columnwidth=f"{column_width_cm(['路徑'] + [r[1] for r in rows]):.2f}cm"
        )
    )
    doc.automaticstyles.addElement(path_col_style)

    table = Table(name=sheet_name)
    table.addElement(TableColumn(stylename=name_col_style))
    table.addElement(TableColumn(stylename=path_col_style))
    doc.spreadsheet.addElement(table)

    header_row = TableRow()
    for header in ("名稱", "路徑"):
        cell = TableCell()
        cell.addElement(P(text=header))
        header_row.addElement(cell)
    table.addElement(header_row)

    for top_name, leaf_name, uri in rows:
        name_cell = TableCell()
        name_cell.addElement(P(text=top_name))
        table_row = TableRow()
        table_row.addElement(name_cell)

        formula = f"of:=HYPERLINK({formula_literal(uri)};{formula_literal(leaf_name)})"
        link_cell = TableCell(formula=formula, valuetype="string", stringvalue=leaf_name)
        link_cell.addElement(P(text=leaf_name))
        table_row.addElement(link_cell)

        table.addElement(table_row)

    out_path = os.path.join(root, f"{output_basename(label)}.ods")
    doc.save(out_path)
    return out_path


def main():
    root = os.getcwd()

    print("===== 檢查 Thumbs.db =====")
    thumbs_files = find_thumbs_files(root)
    if thumbs_files:
        print("找到以下 Thumbs.db 檔案:")
        for f in thumbs_files:
            print(f"  {rel_path(root, f)}")
        if confirm_delete("是否刪除以上 Thumbs.db 檔案? (y/N): "):
            for f in thumbs_files:
                try:
                    os.remove(f)
                    print(f"✅ 已刪除: {rel_path(root, f)}")
                except OSError:
                    pass
    else:
        print("✅ 沒有 Thumbs.db")
    print()

    print("===== 檢查空資料夾 =====")
    empty_dirs = find_empty_dirs(root)
    if empty_dirs:
        print("找到以下空資料夾:")
        for d in empty_dirs:
            print(f"  {rel_path(root, d)}")
        if confirm_delete("是否刪除以上空資料夾? (y/N): "):
            for d in empty_dirs:
                try:
                    os.rmdir(d)
                    print(f"✅ 已刪除: {rel_path(root, d)}")
                except OSError:
                    pass
    else:
        print("✅ 沒有空資料夾")
    print()

    print("===== 檢查 .docx =====")
    dirs_with_docx = []
    dirs_missing_docx = []

    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if not d.startswith(".")]

        is_leaf = len(dirnames) == 0
        if not is_leaf:
            continue

        has_image = any(
            os.path.splitext(f)[1].lower() in IMAGE_EXTENSIONS for f in filenames
        )
        if not has_image:
            continue

        has_docx = any(f.lower().endswith(".docx") for f in filenames)
        if has_docx:
            dirs_with_docx.append(dirpath)
        else:
            dirs_missing_docx.append(dirpath)

    for dirpath in dirs_with_docx:
        print(f"✅ {rel_path(root, dirpath)}")
    for dirpath in dirs_missing_docx:
        print(f"❌ {rel_path(root, dirpath)}")

    leaf_dirs_with_image = len(dirs_with_docx) + len(dirs_missing_docx)
    leaf_dirs_missing_docx = len(dirs_missing_docx)
    print()

    leaf_dirs_with_docx = leaf_dirs_with_image - leaf_dirs_missing_docx
    completion_rate = (
        leaf_dirs_with_docx / leaf_dirs_with_image * 100
        if leaf_dirs_with_image > 0
        else 0.0
    )

    print("===== 統計結果 =====")
    print(f"根目錄: {root.replace('\\', '/')}")
    print(f"有圖片的目錄數: {leaf_dirs_with_image}")
    print(f"有圖片但無 .docx 的目錄數: {leaf_dirs_missing_docx}")
    print(
        f"完成度: {completion_rate:.1f}% ({leaf_dirs_with_docx}/{leaf_dirs_with_image})"
    )

    image_path = generate_missing_docx_image(root, dirs_missing_docx, "照片不完整清單")
    if image_path:
        print(f"已產生圖片: {rel_path(root, image_path)}")

    ods_path = generate_missing_docx_ods(
        root, dirs_missing_docx, "照片不完整清單", "缺少docx清單"
    )
    if ods_path:
        print(f"已產生 ods 清單: {rel_path(root, ods_path)}")
    print()

    print("===== 檢查只有 .docx 沒有圖片的資料夾 =====")
    docx_only_dirs = find_docx_only_dirs(root)
    if docx_only_dirs:
        print("找到以下只有 .docx 沒有圖片的資料夾:")
        for d in docx_only_dirs:
            print(f"  {rel_path(root, d)}")
    else:
        print("✅ 沒有只有 .docx 沒有圖片的資料夾")

    docx_only_image_path = generate_missing_docx_image(
        root, docx_only_dirs, "只有docx沒有圖片清單"
    )
    if docx_only_image_path:
        print(f"已產生圖片: {rel_path(root, docx_only_image_path)}")

    docx_only_ods_path = generate_missing_docx_ods(
        root, docx_only_dirs, "只有docx沒有圖片清單", "只有docx沒有圖片清單"
    )
    if docx_only_ods_path:
        print(f"已產生 ods 清單: {rel_path(root, docx_only_ods_path)}")


if __name__ == "__main__":
    main()
    input("\n按 Enter 鍵結束...")
