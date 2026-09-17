#!/usr/bin/env python3
"""掃描目前資料夾結構與圖片檔名日期，輸出單一 Excel（.xlsx）檢查清單。"""

import os
import re
import sys
import zipfile
from datetime import date
from pathlib import Path
from xml.sax.saxutils import escape, quoteattr

IMAGE_EXTENSIONS = {
    ".jpg",
    ".jpeg",
    ".png",
    ".gif",
    ".bmp",
    ".tiff",
    ".tif",
    ".webp",
    ".heic",
    ".heif",
}

# 民國日期格式：115.12.5 / 115-12-5 / 115_12_05 等（年 2~3 碼，月、日 1~2 碼）
DATE_PATTERN = re.compile(r"(?<!\d)(\d{2,3})[.\-_](\d{1,2})[.\-_](\d{1,2})(?!\d)")

# 資料夾名稱裡的月.日格式：08.01 / 8.1 等（不含年份）
FOLDER_DATE_PATTERN = re.compile(r"(?<!\d)(\d{1,2})[.\-_](\d{1,2})(?!\d)")

# docx 檔名開頭的西元日期格式：20260801_XXXX
DOCX_NAME_DATE_PATTERN = re.compile(r"^(\d{4})(\d{2})(\d{2})_")

# 人名名單檔案，需與 main.py（或打包後的 exe）放在同一個要檢查的資料夾內
NAMES_FILENAME = "123.txt"


def load_name_list(root):
    """讀取人名名單（每行一個名字）。找不到檔案時回傳 None，表示不進行人名檢查。"""
    names_path = Path(root) / NAMES_FILENAME
    if not names_path.is_file():
        return None
    names = set()
    with open(names_path, encoding="utf-8-sig") as f:
        for line in f:
            name = line.strip()
            if name:
                names.add(name)
    return names


def parse_folder_month_day(folder_name: str):
    """從資料夾名稱找出「月.日」（例如 08.01），回傳 (month, day) 或 None。"""
    match = FOLDER_DATE_PATTERN.search(folder_name)
    if not match:
        return None
    month, day = (int(g) for g in match.groups())
    return (month, day)


def find_ancestor_month_day(folder: Path, root: Path):
    """從 folder 往上（含 folder 本身）找最近一層名稱含「月.日」的祖先資料夾，回傳其 (month, day)。

    例如 08.01/aaa 底下的圖片，會往上比對到 08.01 這一層，而不是只看直接上層的 aaa。
    """
    current = folder
    while True:
        month_day = parse_folder_month_day(current.name)
        if month_day is not None:
            return month_day
        if current == root or current.parent == current:
            return None
        current = current.parent


def find_grandparent_month_day(docx_path: Path, root: Path):
    """找出 docx 檔案上上一層資料夾（含往上找）的「月.日」，回傳 (month, day) 或 None。

    例如 08.01/工作項目/xxx.docx，會比對到 08.01 這一層。
    若上上一層已經在 root 之外（docx 放太淺），則不進行比對。
    """
    grandparent = docx_path.parent.parent
    if grandparent != root and root not in grandparent.parents:
        return None
    return find_ancestor_month_day(grandparent, root)


def output_basename(label):
    return f"{date.today():%Y-%m-%d}_{label}"


_NUM_SPLIT = re.compile(r"(\d+)")


def natural_sort_key(text: str):
    """自然排序鍵：讓字串中的數字依數值大小排序（例如 08.01 排在 08.15 之前），
    而非逐字元比較（會讓 08.15 排在 08.2 之前）。"""
    return [
        int(part) if part.isdigit() else part.lower() for part in _NUM_SPLIT.split(text)
    ]


class Hyperlink:
    """代表一個超連結儲存格：點擊後開啟 href，顯示文字為 display。"""

    def __init__(self, href: str, display: str):
        self.href = href
        self.display = display


def path_link(path: Path, display: str) -> Hyperlink:
    """建立指向本機檔案／資料夾的超連結（Windows、Ubuntu 上的 Excel／LibreOffice 皆可點擊開啟）。"""
    uri = path.resolve().as_uri()
    if path.is_dir() and not uri.endswith("/"):
        uri += "/"
    return Hyperlink(uri, display)


# ---------- 資料夾結構檢查 ----------


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


def dir_list_rows(root, dirs):
    """把資料夾清單轉成表格列：單一「資料夾路徑」欄，內容為可點擊超連結，並依自然排序小到大排列。"""
    root_path = Path(root)
    rows = []
    for dirpath in sorted(dirs, key=natural_sort_key):
        folder = Path(dirpath)
        try:
            display = str(folder.relative_to(root_path))
        except ValueError:
            display = str(folder)
        rows.append({"資料夾路徑": path_link(folder, display)})
    return rows


# ---------- 圖片檔名日期檢查 ----------


def find_date_in_filename(filename: str):
    """在檔名（不含副檔名）中尋找民國日期格式。

    回傳 dict（含原始字串、年月日、是否為有效日期、對應西元日期、比對到的位置）或 None（找不到符合格式的字串）。
    """
    match = DATE_PATTERN.search(filename)
    if not match:
        return None

    raw = match.group(0)
    roc_year, month, day = (int(g) for g in match.groups())

    is_valid = True
    western_date_str = ""
    try:
        western_year = roc_year + 1911
        d = date(western_year, month, day)
        western_date_str = d.strftime("%Y-%m-%d")
    except ValueError:
        is_valid = False

    return {
        "raw": raw,
        "roc_year": roc_year,
        "month": month,
        "day": day,
        "is_valid": is_valid,
        "western_date": western_date_str,
        "match_start": match.start(),
        "match_end": match.end(),
    }


def find_person_name(stem: str, result, names):
    """在檔名（不含副檔名）中，日期前面或後面尋找符合名單的人名，回傳找到的名字或 None。"""
    prefix = stem[: result["match_start"]]
    suffix = stem[result["match_end"] :]
    for name in names:
        if name in prefix or name in suffix:
            return name
    return None


def scan_images(root):
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if not d.startswith(".")]
        for f in filenames:
            if os.path.splitext(f)[1].lower() in IMAGE_EXTENSIONS:
                yield Path(dirpath) / f


def build_date_check_sheets(root):
    root_path = Path(root)
    names = load_name_list(root)
    if names is None:
        print(
            f"警告：找不到人名名單 {NAMES_FILENAME}，將略過人名檢查。",
            file=sys.stderr,
        )

    folder_stats = {}
    for path in sorted(scan_images(root), key=lambda p: natural_sort_key(str(p))):
        stem = path.stem
        result = find_date_in_filename(stem)
        folder = path.parent
        stats = folder_stats.setdefault(
            folder,
            {
                "圖片總數": 0,
                "含有效日期圖片數": 0,
                "符合日期檔名": [],
                "日期與資料夾不符檔名": [],
                "人名不在名單檔名": [],
            },
        )
        stats["圖片總數"] += 1

        if result is None or not result["is_valid"]:
            continue

        date_mismatch = False
        folder_month_day = find_ancestor_month_day(folder, root_path)
        if (
            folder_month_day is not None
            and (result["month"], result["day"]) != folder_month_day
        ):
            stats["日期與資料夾不符檔名"].append(path.name)
            date_mismatch = True

        name_mismatch = False
        if names is not None and find_person_name(stem, result, names) is None:
            stats["人名不在名單檔名"].append(path.name)
            name_mismatch = True

        if not date_mismatch and not name_mismatch:
            stats["含有效日期圖片數"] += 1
            stats["符合日期檔名"].append(path.name)

    folder_rows = []
    for folder in sorted(folder_stats, key=lambda p: natural_sort_key(str(p))):
        stats = folder_stats[folder]
        qualified = (
            stats["含有效日期圖片數"] > 0
            and not stats["日期與資料夾不符檔名"]
            and not stats["人名不在名單檔名"]
        )
        try:
            display = str(folder.relative_to(root_path))
        except ValueError:
            display = str(folder)
        folder_rows.append(
            {
                "資料夾路徑": path_link(folder, display if display != "." else ""),
                "圖片總數": stats["圖片總數"],
                "含有效日期圖片數": stats["含有效日期圖片數"],
                "是否合格": "合格" if qualified else "不合格",
                "符合日期檔名": "、".join(stats["符合日期檔名"]),
                "日期與資料夾不符檔名": "、".join(stats["日期與資料夾不符檔名"]),
                "人名不在名單檔名": "、".join(stats["人名不在名單檔名"]),
            }
        )

    dir_fieldnames = [
        "資料夾路徑",
        "圖片總數",
        "含有效日期圖片數",
        "是否合格",
        "符合日期檔名",
        "日期與資料夾不符檔名",
        "人名不在名單檔名",
    ]

    return [
        ("資料夾合格檢核(日期)", dir_fieldnames, folder_rows),
    ]


# ---------- docx 檔名與資料夾日期比對 ----------


def scan_docx_files(root):
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if not d.startswith(".")]
        for f in filenames:
            if f.lower().endswith(".docx"):
                yield Path(dirpath) / f


def check_docx_name_against_folder(path: Path, root: Path):
    """比對單一 docx 檔名開頭日期，是否與其上上一層「月.日」資料夾相符。

    回傳 dict：資料夾日期顯示字串、檔名日期顯示字串、是否合格、不合格原因。
    """
    folder_month_day = find_grandparent_month_day(path, root)
    folder_date_display = (
        f"{folder_month_day[0]:02d}.{folder_month_day[1]:02d}"
        if folder_month_day is not None
        else ""
    )

    match = DOCX_NAME_DATE_PATTERN.match(path.stem)
    if match is None:
        return {
            "folder_date": folder_date_display,
            "file_date": "",
            "qualified": False,
            "reason": "檔名開頭非「YYYYMMDD_」格式",
        }

    year, month, day = (int(g) for g in match.groups())
    file_date_display = f"{year:04d}{month:02d}{day:02d}"

    if folder_month_day is None:
        return {
            "folder_date": folder_date_display,
            "file_date": file_date_display,
            "qualified": False,
            "reason": "找不到上層「月.日」資料夾",
        }

    if year != date.today().year:
        return {
            "folder_date": folder_date_display,
            "file_date": file_date_display,
            "qualified": False,
            "reason": f"年份不是 {date.today().year}",
        }

    if (month, day) != folder_month_day:
        return {
            "folder_date": folder_date_display,
            "file_date": file_date_display,
            "qualified": False,
            "reason": "月日與資料夾不符",
        }

    return {
        "folder_date": folder_date_display,
        "file_date": file_date_display,
        "qualified": True,
        "reason": "",
    }


def build_docx_name_check_sheets(root):
    root_path = Path(root)
    rows = []
    for path in sorted(scan_docx_files(root), key=lambda p: natural_sort_key(str(p))):
        result = check_docx_name_against_folder(path, root_path)
        try:
            display = str(path.relative_to(root_path))
        except ValueError:
            display = str(path)
        rows.append(
            {
                "docx路徑": path_link(path, display),
                "資料夾日期(月.日)": result["folder_date"],
                "檔名日期": result["file_date"],
                "是否合格": "合格" if result["qualified"] else "不合格",
                "不合格原因": result["reason"],
            }
        )

    fieldnames = ["docx路徑", "資料夾日期(月.日)", "檔名日期", "是否合格", "不合格原因"]

    return [
        ("docx檔名與資料夾比對", fieldnames, rows),
    ]


# ---------- XLSX 寫出（純標準庫，不依賴 openpyxl 等第三方套件） ----------


def _col_letter(index: int) -> str:
    """0-based 欄位索引轉成 Excel 欄位字母（0->A, 25->Z, 26->AA ...）。"""
    index += 1
    letters = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _display_text(value) -> str:
    if isinstance(value, Hyperlink):
        return value.display
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return str(value)
    return "" if value is None else str(value)


def _display_width(text: str) -> float:
    """粗略估算顯示寬度：中日韓字元視為 2 個字元寬，其餘視為 1。"""
    width = 0.0
    for ch in text:
        width += 2.0 if ord(ch) > 0x2E80 else 1.0
    return width


def _build_sheet_xml(fieldnames: list, rows: list):
    """組出單一分頁的 sheetN.xml 內容，回傳 (sheet_xml, sheet_rels_xml_or_None)。"""
    col_count = len(fieldnames)

    col_widths = [_display_width(name) for name in fieldnames]
    for row in rows:
        for i, name in enumerate(fieldnames):
            text = _display_text(row.get(name))
            col_widths[i] = max(col_widths[i], _display_width(text))
    col_widths = [max(8.0, min(60.0, w + 2.0)) for w in col_widths]

    cols_xml = "".join(
        f'<col min="{i + 1}" max="{i + 1}" width="{w:.2f}" customWidth="1"/>'
        for i, w in enumerate(col_widths)
    )

    hyperlink_entries = []  # (cell_ref, target_href)

    def cell_xml(row_idx: int, col_idx: int, value) -> str:
        ref = f"{_col_letter(col_idx)}{row_idx}"
        if row_idx == 1:
            style_attr = ' s="2"'
        elif isinstance(value, Hyperlink):
            style_attr = ' s="1"'
        else:
            style_attr = ""

        if isinstance(value, Hyperlink):
            hyperlink_entries.append((ref, value.href))
            text = escape(value.display)
            return f'<c r="{ref}"{style_attr} t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return f'<c r="{ref}"{style_attr}><v>{value}</v></c>'
        text = "" if value is None else str(value)
        if text == "":
            return f'<c r="{ref}"{style_attr}/>'
        text = escape(text)
        return f'<c r="{ref}"{style_attr} t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'

    row_xml_parts = []
    header_cells = "".join(cell_xml(1, i, name) for i, name in enumerate(fieldnames))
    row_xml_parts.append(f'<row r="1">{header_cells}</row>')
    for r, row in enumerate(rows, start=2):
        cells = "".join(
            cell_xml(r, i, row.get(name)) for i, name in enumerate(fieldnames)
        )
        row_xml_parts.append(f'<row r="{r}">{cells}</row>')
    sheet_data_xml = "".join(row_xml_parts)

    hyperlinks_xml = ""
    rels_entries = []
    if hyperlink_entries:
        links_xml_parts = []
        for i, (ref, href) in enumerate(hyperlink_entries, start=1):
            rid = f"rId{i}"
            links_xml_parts.append(f'<hyperlink ref="{ref}" r:id="{rid}"/>')
            rels_entries.append(
                f'<Relationship Id="{rid}" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" '
                f'Target={quoteattr(href)} TargetMode="External"/>'
            )
        hyperlinks_xml = f"<hyperlinks>{''.join(links_xml_parts)}</hyperlinks>"

    last_col = _col_letter(col_count - 1)
    last_row = len(rows) + 1

    sheet_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:{last_col}{last_row}"/>
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <cols>{cols_xml}</cols>
  <sheetData>{sheet_data_xml}</sheetData>
  {hyperlinks_xml}
</worksheet>
"""

    sheet_rels_xml = None
    if rels_entries:
        sheet_rels_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(rels_entries)
            + "</Relationships>\n"
        )

    return sheet_xml, sheet_rels_xml


def write_xlsx(path: Path, sheets: list):
    """用純標準庫（zipfile + 手刻 OOXML）寫出 .xlsx，不依賴 openpyxl 等第三方套件。

    sheets: [(sheet_name, fieldnames, rows), ...]，依序成為工作簿中的各分頁。
    """
    content_type_overrides = []
    workbook_sheet_entries = []
    workbook_rels_entries = []
    sheet_files = {}

    for i, (sheet_name, fieldnames, rows) in enumerate(sheets, start=1):
        sheet_xml, sheet_rels_xml = _build_sheet_xml(fieldnames, rows)
        sheet_files[f"xl/worksheets/sheet{i}.xml"] = sheet_xml
        if sheet_rels_xml:
            sheet_files[f"xl/worksheets/_rels/sheet{i}.xml.rels"] = sheet_rels_xml
        content_type_overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
        rid = f"rId{i}"
        workbook_sheet_entries.append(
            f'<sheet name={quoteattr(sheet_name)} sheetId="{i}" r:id="{rid}"/>'
        )
        workbook_rels_entries.append(
            f'<Relationship Id="{rid}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{i}.xml"/>'
        )

    styles_rid = f"rId{len(sheets) + 1}"
    workbook_rels_entries.append(
        f'<Relationship Id="{styles_rid}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )

    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        + "".join(content_type_overrides)
        + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        "</Types>\n"
    )

    root_rels_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"""

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<sheets>" + "".join(workbook_sheet_entries) + "</sheets>"
        "</workbook>\n"
    )

    workbook_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(workbook_rels_entries)
        + "</Relationships>\n"
    )

    styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="2" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
"""

    if path.exists():
        path.unlink()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", root_rels_xml)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        zf.writestr("xl/styles.xml", styles_xml)
        for name, data in sheet_files.items():
            zf.writestr(name, data)


def generate_combined_xlsx(root, sections, date_check_sheets, label):
    sheets = [
        (sheet_name, ["資料夾路徑"], dir_list_rows(root, dirs))
        for dirs, sheet_name in sections
        if dirs
    ]
    sheets.extend(date_check_sheets)
    if not sheets:
        return None

    out_path = Path(root) / f"{output_basename(label)}.xlsx"
    write_xlsx(out_path, sheets)
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

    date_check_sheets = build_date_check_sheets(root)
    docx_name_check_sheets = build_docx_name_check_sheets(root)

    generate_combined_xlsx(
        root,
        [
            (dirs_missing_docx, "缺少圖片"),
            (missing_material_dirs, "領料單"),
            (has_or_dirs, "名稱含(或)"),
            (empty_dirs, "空資料夾"),
            (docx_only_dirs, "只有docx沒有圖片(需要另存圖片)"),
        ],
        date_check_sheets + docx_name_check_sheets,
        "檢查清單",
    )


if __name__ == "__main__":
    main()
