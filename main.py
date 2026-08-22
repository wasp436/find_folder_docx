import os


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
            print(f"  {f}")
        if confirm_delete("是否刪除以上 Thumbs.db 檔案? (輸入 yes 刪除): "):
            for f in thumbs_files:
                try:
                    os.remove(f)
                    print(f"✅ 已刪除: {f}")
                except OSError:
                    pass
    if not thumbs_files:
        print("✅ 沒有 Thumbs.db")

    print("\n===== 檢查空資料夾 =====")
    empty_dirs = find_empty_dirs(root)
    if empty_dirs:
        print("找到以下空資料夾:")
        for d in empty_dirs:
            print(f"  {d}")
        if confirm_delete("是否刪除以上空資料夾? (輸入 yes 刪除): "):
            for d in empty_dirs:
                try:
                    os.rmdir(d)
                    print(f"✅ 已刪除: {d}")
                except OSError:
                    pass
    if not empty_dirs:
        print("✅ 沒有空資料夾")

    print("\n===== 檢查 .docx =====")
    leaf_dirs = 0
    leaf_dirs_with_docx = 0

    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if not d.startswith(".")]

        is_leaf = len(dirnames) == 0
        if not is_leaf:
            continue

        docx_files = [f for f in filenames if f.lower().endswith(".docx")]
        has_docx = len(docx_files) > 0

        leaf_dirs += 1
        if has_docx:
            leaf_dirs_with_docx += 1

        if not has_docx:
            print(f"{dirpath}  [❌ 無 .docx]")

    completion_rate = leaf_dirs_with_docx / leaf_dirs * 100 if leaf_dirs > 0 else 0.0

    print("\n===== 統計結果 =====")
    print(f"根目錄: {root}")
    print(f"最後一層目錄數: {leaf_dirs}")
    print(f"有 .docx 的目錄數: {leaf_dirs_with_docx}")
    print(f"完成度: {completion_rate:.1f}% ({leaf_dirs_with_docx}/{leaf_dirs})")


if __name__ == "__main__":
    main()
    input("\n按 Enter 鍵結束...")
