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

# 資料夾名稱「整段」需恰好是月.日格式：08.01 / 8.1 等（不含年份），避免資料夾名稱中的
# 型號、設備編號等子字串（例如「EF-3-1」）被誤判為日期
FOLDER_DATE_PATTERN = re.compile(r"(\d{1,2})[.\-_](\d{1,2})")

# docx 檔名開頭的西元日期格式：20260801_XXXX
DOCX_NAME_DATE_PATTERN = re.compile(r"^(\d{4})(\d{2})(\d{2})_")

# 人名名單檔案，需與 main.py（或打包後的 exe）放在同一個要檢查的資料夾內
NAMES_FILENAME = "123.txt"

# 掃描時要略過的資料夾名稱：隱藏資料夾與 Python 快取資料夾，避免被誤判為工作項目資料夾
IGNORED_DIR_NAMES = {"__pycache__"}

# 資料夾名稱含有此關鍵字時，預設不納入統計，除非執行時輸入 yes 明確納入
CHLORINE_TABLET_KEYWORD = "氯錠"

# 是否將名稱含「氯錠」的資料夾納入統計；由 main() 開頭詢問使用者後設定
_include_chlorine_tablet_dirs = False

# 根目錄底下第一層資料夾名稱須恰好是 01~12 這種月份格式，其他名稱（例如 test）一律略過不掃描
MONTH_DIR_PATTERN = re.compile(r"^(0[1-9]|1[0-2])$")


def _filter_dirnames(dirnames, dirpath, root):
    filtered = [
        d for d in dirnames if not d.startswith(".") and d not in IGNORED_DIR_NAMES
    ]
    if not _include_chlorine_tablet_dirs:
        filtered = [d for d in filtered if CHLORINE_TABLET_KEYWORD not in d]
    if dirpath == root:
        filtered = [d for d in filtered if MONTH_DIR_PATTERN.fullmatch(d)]
    return filtered


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
    """判斷資料夾名稱整段是否恰好是「月.日」格式（例如 08.01），回傳 (month, day) 或 None。

    只接受整段名稱完全符合，避免名稱中夾雜的型號、設備編號等子字串
    （例如「中正7F-3-1保養」裡的「3-1」）被誤判為日期。
    """
    match = FOLDER_DATE_PATTERN.fullmatch(folder_name.strip())
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
    """取出 docx 檔案「上上一層」資料夾名稱的「月.日」，回傳 (month, day) 或 None。

    固定路徑結構為 root/月/月.日/工作項目/xxx.docx，因此直接比對 docx 上上一層
    （即月.日資料夾本身）的名稱，不往上額外搜尋，避免比對到非預期的祖先資料夾。
    """
    grandparent = docx_path.parent.parent
    if grandparent != root and root not in grandparent.parents:
        return None
    return parse_folder_month_day(grandparent.name)


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
        dirnames[:] = _filter_dirnames(dirnames, dirpath, root)
        for f in filenames:
            if f.lower() == "thumbs.db":
                found.append(os.path.join(dirpath, f))
    return found


def find_empty_dirs(root):
    empty = []
    for dirpath, dirnames, filenames in os.walk(root, topdown=True):
        dirnames[:] = _filter_dirnames(dirnames, dirpath, root)
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
        dirnames[:] = _filter_dirnames(dirnames, dirpath, root)
        if dirpath == root:
            continue
        name = os.path.basename(dirpath)
        if any(keyword in name for keyword in keywords):
            result.append(dirpath)
    return result


def find_docx_only_dirs(root):
    result = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = _filter_dirnames(dirnames, dirpath, root)

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
        dirnames[:] = _filter_dirnames(dirnames, dirpath, root)
        for f in filenames:
            if os.path.splitext(f)[1].lower() in IMAGE_EXTENSIONS:
                yield Path(dirpath) / f


def compute_date_check_folder_stats(root):
    """依資料夾分組，統計圖片日期／人名比對結果，回傳 {folder(Path): stats}。"""
    root_path = Path(root)
    names = load_name_list(root)

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

    return folder_stats


def is_date_check_folder_qualified(stats) -> bool:
    return (
        stats["含有效日期圖片數"] > 0
        and not stats["日期與資料夾不符檔名"]
        and not stats["人名不在名單檔名"]
    )


def build_date_check_sheets(root, folder_stats):
    root_path = Path(root)
    folder_rows = []
    for folder in sorted(folder_stats, key=lambda p: natural_sort_key(str(p))):
        stats = folder_stats[folder]
        qualified = is_date_check_folder_qualified(stats)
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
        dirnames[:] = _filter_dirnames(dirnames, dirpath, root)
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


def compute_docx_check_results(root):
    """回傳 [(docx路徑, 檢查結果dict), ...]，供分頁與統計共用。"""
    root_path = Path(root)
    return [
        (path, check_docx_name_against_folder(path, root_path))
        for path in sorted(
            scan_docx_files(root), key=lambda p: natural_sort_key(str(p))
        )
    ]


def build_docx_name_check_sheets(root, docx_results):
    root_path = Path(root)
    rows = []
    for path, result in docx_results:
        folder = path.parent
        try:
            display = str(folder.relative_to(root_path))
        except ValueError:
            display = str(folder)
        rows.append(
            {
                "資料夾路徑": path_link(folder, display if display != "." else ""),
                "資料夾日期(月.日)": result["folder_date"],
                "檔名日期": result["file_date"],
                "是否合格": "合格" if result["qualified"] else "不合格",
                "不合格原因": result["reason"],
            }
        )

    fieldnames = [
        "資料夾路徑",
        "資料夾日期(月.日)",
        "檔名日期",
        "是否合格",
        "不合格原因",
    ]

    return [
        ("docx檔名與資料夾比對", fieldnames, rows),
    ]


# ---------- 人名＋日期檔名格式檢核 ----------


def check_name_date_format(stem: str, result, matched_name: str):
    """檢查「人名＋日期」在檔名中的先後順序、是否緊鄰，以及日期是否月、日皆補零成二碼。

    合格：人名115.05.01（人名緊接在日期前面，中間無其他字元，且月、日皆二碼）。
    不合格：人名115.5.1（月或日未補零）、115.05.01人名（日期在人名前面）、
           或人名與日期中間夾了其他字元（例如空白，如「人名 115.05.01」）。
    """
    prefix = stem[: result["match_start"]]

    if not prefix.endswith(matched_name):
        if matched_name in prefix:
            return False, "人名與日期之間有其他字元（例如空白）"
        return False, "日期出現在人名前面（順序相反）"

    date_parts = re.split(r"[.\-_]", result["raw"])
    month_str, day_str = date_parts[1], date_parts[2]
    is_zero_padded = len(month_str) == 2 and len(day_str) == 2

    if not is_zero_padded:
        return False, "月或日未補零成二碼"
    return True, ""


def suggest_qualified_filename(path: Path, result, matched_name: str) -> str:
    """組出「若修正成合格格式」的檔名參考：人名緊接補零後的日期，其餘文字保留在後面。"""
    corrected_date = f"{result['roc_year']}.{result['month']:02d}.{result['day']:02d}"
    extra_text = (
        path.stem.replace(result["raw"], "", 1).replace(matched_name, "", 1).strip()
    )
    return f"{matched_name}{corrected_date}{extra_text}{path.suffix}"


def scan_name_date_format_issues(root):
    """為每張「同時符合人名名單與日期格式」的圖片，產出檢查結果。

    yield (path, matched_name, result, qualified, reason, suggested_name)。
    """
    names = load_name_list(root)
    if not names:
        return

    for path in sorted(scan_images(root), key=lambda p: natural_sort_key(str(p))):
        stem = path.stem
        result = find_date_in_filename(stem)
        if result is None or not result["is_valid"]:
            continue

        matched_name = find_person_name(stem, result, names)
        if matched_name is None:
            continue

        qualified, reason = check_name_date_format(stem, result, matched_name)
        suggested_name = suggest_qualified_filename(path, result, matched_name)
        yield path, matched_name, result, qualified, reason, suggested_name


def build_name_date_format_check_sheets(root, issues):
    root_path = Path(root)

    rows = []
    for (
        path,
        matched_name,
        result,
        qualified,
        reason,
        suggested_name,
    ) in issues:
        folder = path.parent
        try:
            display = str(folder.relative_to(root_path))
        except ValueError:
            display = str(folder)

        rows.append(
            {
                "資料夾路徑": path_link(folder, display if display != "." else ""),
                "比對到的人名": matched_name,
                "比對到的日期": result["raw"],
                "是否合格": "合格" if qualified else "不合格",
                "不合格原因": reason,
                "原始檔名": path.name,
                "修正成合格後的檔名(參考)": suggested_name,
            }
        )

    fieldnames = [
        "資料夾路徑",
        "比對到的人名",
        "比對到的日期",
        "是否合格",
        "不合格原因",
        "原始檔名",
        "修正成合格後的檔名(參考)",
    ]

    return [
        ("人名日期格式檢核", fieldnames, rows),
    ]


def prompt_and_rename_unqualified_images(root, issues):
    """列出不合格的「人名＋日期」圖片檔名，顯示修改前後對照，需輸入 yes 才會實際重新命名。"""
    root_path = Path(root)
    plan = []
    for (
        path,
        _matched_name,
        _result,
        qualified,
        _reason,
        suggested_name,
    ) in issues:
        if qualified:
            continue
        new_path = path.with_name(suggested_name)
        if new_path == path:
            continue
        plan.append((path, new_path))

    if not plan:
        return

    print()
    print("=== 以下圖片檔名不符合「人名+日期」合格格式，可修改為建議檔名 ===")
    for old_path, new_path in plan:
        try:
            display = str(old_path.parent.relative_to(root_path))
        except ValueError:
            display = str(old_path.parent)
        print(f"路徑：{display}")
        print(f"  修改前：{old_path.name}")
        print(f"  修改後：{new_path.name}")

    answer = input(
        f"\n共 {len(plan)} 個檔案，是否要套用以上重新命名？輸入 yes 才會執行："
    )
    if answer.strip().lower() != "yes":
        print("已取消，未修改任何檔案。")
        return

    renamed = 0
    for old_path, new_path in plan:
        if new_path.exists():
            print(f"略過（目標檔名已存在）：{new_path}")
            continue
        old_path.rename(new_path)
        renamed += 1
    print(f"已完成，共重新命名 {renamed} 個檔案。")


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


# 試算表公式觸發字元：儲存格內容若以這些字元開頭，Excel/LibreOffice 在某些情境下
# （例如先匯出成 CSV 再重新開啟）可能會把純文字誤判成公式並執行
_FORMULA_TRIGGER_CHARS = ("=", "+", "-", "@", "\t", "\r")


def _neutralize_formula_prefix(text: str) -> str:
    """若文字開頭是公式觸發字元，補一個單引號讓試算表軟體視為純文字。"""
    if text and text[0] in _FORMULA_TRIGGER_CHARS:
        return "'" + text
    return text


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
            text = escape(_neutralize_formula_prefix(value.display))
            return f'<c r="{ref}"{style_attr} t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return f'<c r="{ref}"{style_attr}><v>{value}</v></c>'
        text = "" if value is None else str(value)
        if text == "":
            return f'<c r="{ref}"{style_attr}/>'
        text = escape(_neutralize_formula_prefix(text))
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

    if path.is_symlink() or path.exists():
        path.unlink()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", root_rels_xml)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        zf.writestr("xl/styles.xml", styles_xml)
        for name, data in sheet_files.items():
            zf.writestr(name, data)


# ---------- 統計總覽 ----------


def find_all_leaf_dirs(root):
    """回傳所有葉資料夾（沒有子資料夾）的路徑字串，作為「資料夾總數」的統計基準。"""
    leaf_dirs = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = _filter_dirnames(dirnames, dirpath, root)
        if dirpath == root:
            continue
        if not dirnames:
            leaf_dirs.append(dirpath)
    return leaf_dirs


def _norm_dir(path) -> str:
    return os.path.normpath(str(path))


def build_summary_sheet(total_dirs, unqualified_dirs):
    """統計「資料夾總數」與「不合格資料夾數」，不合格資料夾以路徑去重，不重複計算。"""
    total = {_norm_dir(d) for d in total_dirs}
    unqualified = {_norm_dir(d) for d in unqualified_dirs} & total
    qualified = len(total) - len(unqualified)
    completion_rate = f"{qualified / len(total) * 100:.1f}%" if total else "N/A"

    rows = [
        {"項目": "資料夾總數", "數量": len(total)},
        {"項目": "合格資料夾數", "數量": qualified},
        {"項目": "不合格資料夾數", "數量": len(unqualified)},
        {"項目": "完成率", "數量": completion_rate},
    ]
    fieldnames = ["項目", "數量"]

    return [("統計總覽", fieldnames, rows)]


def generate_combined_xlsx(
    root, sections, date_check_sheets, label, leading_sheets=None
):
    sheets = list(leading_sheets or [])
    sheets.extend(
        (sheet_name, ["資料夾路徑"], dir_list_rows(root, dirs))
        for dirs, sheet_name in sections
        if dirs
    )
    sheets.extend(date_check_sheets)
    if not sheets:
        return None

    out_path = Path(root) / f"{output_basename(label)}.xlsx"
    write_xlsx(out_path, sheets)
    return out_path


def main():
    global _include_chlorine_tablet_dirs
    answer = input(
        f"是否要將資料夾名稱含有「{CHLORINE_TABLET_KEYWORD}」的資料夾納入統計？輸入 yes 才會納入："
    )
    _include_chlorine_tablet_dirs = answer.strip().lower() == "yes"

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
        dirnames[:] = _filter_dirnames(dirnames, dirpath, root)

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

    if load_name_list(root) is None:
        print(
            f"警告：找不到人名名單 {NAMES_FILENAME}，將略過人名檢查。",
            file=sys.stderr,
        )

    date_check_folder_stats = compute_date_check_folder_stats(root)
    date_check_sheets = build_date_check_sheets(root, date_check_folder_stats)

    docx_check_results = compute_docx_check_results(root)
    docx_name_check_sheets = build_docx_name_check_sheets(root, docx_check_results)

    name_date_format_issues = list(scan_name_date_format_issues(root))
    name_date_format_sheets = build_name_date_format_check_sheets(
        root, name_date_format_issues
    )

    unqualified_dirs = set(dirs_missing_docx) | set(empty_dirs) | set(docx_only_dirs)
    unqualified_dirs |= {
        folder
        for folder, stats in date_check_folder_stats.items()
        if not is_date_check_folder_qualified(stats)
    }
    unqualified_dirs |= {
        path.parent for path, result in docx_check_results if not result["qualified"]
    }
    unqualified_dirs |= {
        path.parent
        for path, _matched_name, _result, qualified, _reason, _suggested_name in (
            name_date_format_issues
        )
        if not qualified
    }

    summary_sheet = build_summary_sheet(find_all_leaf_dirs(root), unqualified_dirs)

    generate_combined_xlsx(
        root,
        [
            (dirs_missing_docx, "缺少圖片"),
            (missing_material_dirs, "領料單"),
            (has_or_dirs, "名稱含(或)"),
            (empty_dirs, "空資料夾"),
            (docx_only_dirs, "只有docx沒有圖片(需要另存圖片)"),
        ],
        date_check_sheets + docx_name_check_sheets + name_date_format_sheets,
        "檢查清單",
        leading_sheets=summary_sheet,
    )

    prompt_and_rename_unqualified_images(root, name_date_format_issues)


if __name__ == "__main__":
    main()
