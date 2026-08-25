import os
import sys

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff"}

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


def rel_path(root, path):
    return os.path.join(".", os.path.relpath(path, root)).replace("\\", "/")


def confirm_delete(prompt):
    ans = input(prompt).strip().lower()
    return ans == "y"


def generate_missing_docx_image(root, dirs_missing_docx):
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

    out_path = os.path.join(root, "缺少docx清單.png")
    img.save(out_path)
    return out_path


def generate_missing_docx_ods(root, dirs_missing_docx):
    if not dirs_missing_docx:
        return None

    try:
        from pathlib import Path

        from odf.opendocument import OpenDocumentSpreadsheet
        from odf.table import Table, TableCell, TableRow
        from odf.text import P
    except ImportError:
        print("⚠️ 未安裝 odfpy，略過 ods 產生 (pip install odfpy)")
        return None

    def formula_literal(text):
        return '"' + text.replace('"', '""') + '"'

    doc = OpenDocumentSpreadsheet()

    table = Table(name="缺少docx清單")
    doc.spreadsheet.addElement(table)

    header_row = TableRow()
    for header in ("名稱", "路徑"):
        cell = TableCell()
        cell.addElement(P(text=header))
        header_row.addElement(cell)
    table.addElement(header_row)

    for dirpath in dirs_missing_docx:
        name = os.path.relpath(dirpath, root).replace("\\", "/")
        uri = Path(dirpath).resolve().as_uri()
        if not uri.endswith("/"):
            uri += "/"

        name_cell = TableCell()
        name_cell.addElement(P(text=name))
        table_row = TableRow()
        table_row.addElement(name_cell)

        formula = f"of:=HYPERLINK({formula_literal(uri)};{formula_literal(name)})"
        link_cell = TableCell(formula=formula, valuetype="string", stringvalue=name)
        link_cell.addElement(P(text=name))
        table_row.addElement(link_cell)

        table.addElement(table_row)

    out_path = os.path.join(root, "缺少docx清單.ods")
    doc.save(out_path)
    return out_path


def get_root():
    cwd = os.getcwd()
    if getattr(sys, "frozen", False):
        # onedir 打包後 exe 位於子資料夾中，實際目標根目錄在上一層
        return os.path.dirname(cwd)
    return cwd


def main():
    root = get_root()

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

    image_path = generate_missing_docx_image(root, dirs_missing_docx)
    if image_path:
        print(f"已產生圖片: {rel_path(root, image_path)}")

    ods_path = generate_missing_docx_ods(root, dirs_missing_docx)
    if ods_path:
        print(f"已產生 ods 清單: {rel_path(root, ods_path)}")


if __name__ == "__main__":
    main()
    input("\n按 Enter 鍵結束...")
