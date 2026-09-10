import os
from datetime import date

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff"}


def output_basename(label):
    return f"{date.today():%Y-%m-%d}_{label}"


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


def path_is_within(path, ancestor):
    return path == ancestor or path.startswith(ancestor + os.sep)


def find_dirs_by_name_keyword(root, keyword):
    return find_dirs_by_name_keywords(root, [keyword])


def find_dirs_by_name_keywords(root, keywords):
    result = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if not d.startswith(".")]
        if dirpath == root:
            continue
        name = os.path.basename(dirpath)
        if any(keyword in name for keyword in keywords):
            result.append(dirpath)
    return result


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


def _add_dirs_sheet(doc, root, dirs, sheet_name):
    from pathlib import Path

    from odf.style import Style, TableColumnProperties
    from odf.table import Table, TableCell, TableColumn, TableRow
    from odf.text import P

    def formula_literal(text):
        return '"' + text.replace('"', '""') + '"'

    def char_width(text):
        return sum(2 if ord(ch) > 0x2E80 else 1 for ch in text)

    def column_width_cm(texts, min_cm=1.5, max_cm=25.0):
        max_chars = max((char_width(t) for t in texts), default=0)
        return max(min_cm, min(max_cm, max_chars * 0.19 + 0.5))

    rows = []
    for dirpath in dirs:
        rel_parts = os.path.relpath(dirpath, root).replace("\\", "/").split("/")
        leaf_name = rel_parts[-1]
        if len(rel_parts) >= 3:
            top_name = "/".join(rel_parts[:2])
        else:
            top_name = rel_parts[0]
        uri = Path(dirpath).resolve().as_uri()
        if not uri.endswith("/"):
            uri += "/"
        rows.append((top_name, leaf_name, uri))

    name_col_style = Style(name=f"{sheet_name}NameCol", family="table-column")
    name_col_style.addElement(
        TableColumnProperties(
            columnwidth=f"{column_width_cm(['名稱'] + [r[0] for r in rows]):.2f}cm"
        )
    )
    doc.automaticstyles.addElement(name_col_style)

    path_col_style = Style(name=f"{sheet_name}PathCol", family="table-column")
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
        link_cell = TableCell(
            formula=formula, valuetype="string", stringvalue=leaf_name
        )
        link_cell.addElement(P(text=leaf_name))
        table_row.addElement(link_cell)

        table.addElement(table_row)


def generate_combined_ods(root, sections, label):
    sections = [(dirs, sheet_name) for dirs, sheet_name in sections if dirs]
    if not sections:
        return None

    try:
        from odf.opendocument import OpenDocumentSpreadsheet
    except ImportError:
        return None

    doc = OpenDocumentSpreadsheet()
    for dirs, sheet_name in sections:
        _add_dirs_sheet(doc, root, dirs, sheet_name)

    out_path = os.path.join(root, f"{output_basename(label)}.ods")
    doc.save(out_path)
    return out_path


def main():
    root = os.getcwd()

    thumbs_files = find_thumbs_files(root)
    for f in thumbs_files:
        try:
            os.remove(f)
        except OSError:
            pass

    empty_dirs = find_empty_dirs(root)

    missing_material_dirs = find_dirs_by_name_keyword(root, "領料單")
    has_or_dirs = find_dirs_by_name_keywords(root, ["(", ")"])

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
        if not has_docx:
            dirs_missing_docx.append(dirpath)

    dirs_missing_docx = [
        d
        for d in dirs_missing_docx
        if not any(path_is_within(d, m) for m in missing_material_dirs)
    ]

    docx_only_dirs = find_docx_only_dirs(root)

    generate_combined_ods(
        root,
        [
            (dirs_missing_docx, "缺少圖片"),
            (missing_material_dirs, "領料單"),
            (has_or_dirs, "名稱含(或)"),
            (empty_dirs, "空資料夾"),
            (docx_only_dirs, "只有docx沒有圖片(需要把圖片另存出來)"),
        ],
        "檢查清單",
    )


if __name__ == "__main__":
    main()
