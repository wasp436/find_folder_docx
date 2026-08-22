import os

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff"}


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
    return os.path.join(".", os.path.relpath(path, root))


def confirm_delete(prompt):
    while True:
        ans = input(prompt).strip().lower()
        if ans == "yes":
            return True
        if ans == "no":
            return False
        print("請輸入 yes 或 no")


def main():
    root = os.getcwd()

    print("===== 檢查 Thumbs.db =====")
    thumbs_files = find_thumbs_files(root)
    if thumbs_files:
        print("找到以下 Thumbs.db 檔案:")
        for f in thumbs_files:
            print(f"  {rel_path(root, f)}")
        if confirm_delete("是否刪除以上 Thumbs.db 檔案? (輸入 yes 刪除): "):
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
        if confirm_delete("是否刪除以上空資料夾? (輸入 yes 刪除): "):
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
    print(f"根目錄: {root}")
    print(f"有圖片的目錄數: {leaf_dirs_with_image}")
    print(f"有圖片但無 .docx 的目錄數: {leaf_dirs_missing_docx}")
    print(
        f"完成度: {completion_rate:.1f}% ({leaf_dirs_with_docx}/{leaf_dirs_with_image})"
    )


if __name__ == "__main__":
    main()
    input("\n按 Enter 鍵結束...")
