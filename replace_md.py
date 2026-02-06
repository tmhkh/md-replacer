# =============================================================
# ライブラリのインポート
# =============================================================
import os
import glob
import openpyxl

# =============================================================
# 定数定義
# =============================================================
REPLACE_EXCEL_PATH = "./replace.xlsx"   # 置換定義ファイルのパス
MD_FOLDER_PATH     = "./md"             # 置換対象の .md ファイルが入っているフォルダ
MD_EXTENSION       = ".md"              # 対象ファイルの拡張子
SHEET_INDEX        = 0                  # replace.xlsx の読み込みシート (先頭シート)
COL_BEFORE         = 0                  # A列 (置換前)
COL_AFTER          = 1                  # B列 (置換後)
ENCODING           = "utf-8"            # .md ファイルのエンコーディング

# =============================================================
# 1. replace.xlsx から置換ペアを読み込む
# =============================================================
wb = openpyxl.load_workbook(REPLACE_EXCEL_PATH)
ws = wb.worksheets[SHEET_INDEX]

replace_pairs = []
for row in ws.iter_rows(min_row=2, values_only=True):   # 1行目はヘッダーと想定
    before = row[COL_BEFORE]
    after  = row[COL_AFTER]
    if before is None:
        continue
    # after が None（空セル）の場合は空文字に置換
    replace_pairs.append((str(before), str(after) if after is not None else ""))

print(f"置換ペア数: {len(replace_pairs)}")
for b, a in replace_pairs:
    print(f"  「{b}」 → 「{a}」")

# =============================================================
# 2. md フォルダ内の .md ファイルを取得
# =============================================================
md_pattern = os.path.join(MD_FOLDER_PATH, f"*{MD_EXTENSION}")
md_files = sorted(glob.glob(md_pattern))

print(f"\n対象ファイル数: {len(md_files)}")
for f in md_files:
    print(f"  {f}")

# =============================================================
# 3. 各 .md ファイルに対して置換を実行
# =============================================================
for md_file in md_files:
    # ファイル読み込み
    with open(md_file, "r", encoding=ENCODING) as f:
        content = f.read()

    original_content = content
    count_total = 0

    # 全ての置換ペアを順番に適用
    for before, after in replace_pairs:
        cnt = content.count(before)
        if cnt > 0:
            content = content.replace(before, after)
            count_total += cnt

    # 変更があった場合のみ書き込み
    if content != original_content:
        with open(md_file, "w", encoding=ENCODING) as f:
            f.write(content)
        print(f"✅ {os.path.basename(md_file)} — {count_total} 箇所置換しました")
    else:
        print(f"⏭️  {os.path.basename(md_file)} — 置換対象なし（変更なし）")

print("\n🎉 すべての処理が完了しました。")