"""
================================================================================
【イベントチェック名簿作成システム - 最終仕様】

1. 自動フィルタリング（2段階）:
   - 【新規】「対象外」列が TRUE のデータを抽出し「excluded_target_out.csv」へ出力。
   - 【新規】「チェックOK」列が TRUE でないデータを抽出し「excluded_check_not_ok.csv」へ出力。
   - 上記いずれにも該当しない（名簿掲載対象）のみを処理。
2. 賢いカテゴリ分け:
   - 参加権A・B・Cを判定。複数所持や該当なしは「複合・その他」へ自動集約。
3. こだわりの並び替え:
   - 参加権A・B・複合・その他：枚数が多い順（降順） → 受付番号順（昇順）。
   - 参加権C：純粋な 受付番号順（昇順）。
4. 2列・高密度レイアウト:
   - 通常カテゴリは1ページ20名、情報量の多い「複合・その他」は1ページ10名の省スペース設計。
5. 完全連動の通し番号:
   - PDFのチェックボックス下と、出力CSVの1列目に共通の「通し番号」を自動付与。
6. ミニマルデザイン:
   - 住所や氏名の「ラベル」を排除し、データのみをスッキリ表示して視認性を向上。
================================================================================
"""

import pandas as pd
import glob
import os
import re
import math
from fpdf import FPDF
from fpdf.enums import XPos, YPos

# --- 基本設定（パス・項目名） ---
BASE_DIR = "/Users/suzukiarato/PycharmProjects/20260310_Check_pdf_crerater"
INPUT_DIR = os.path.join(BASE_DIR, 'csv_files')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')

# 出力ファイルパスの定義
OUTPUT_CSV_ALL    = os.path.join(OUTPUT_DIR, 'merged_data_all.csv')
OUTPUT_CSV_NOT_OK = os.path.join(OUTPUT_DIR, 'excluded_check_not_ok.csv') # チェックOK=FALSE用
OUTPUT_CSV_TARGET = os.path.join(OUTPUT_DIR, 'excluded_target_out.csv')   # 対象外=TRUE用

FONT_NAME = 'ipaexg.ttf'
FONT_PATH = os.path.join(BASE_DIR, FONT_NAME)

# 項目名定義
COL_SERIAL  = '通し番号'
COL_ID      = '受付番号'
COL_ZIP     = '元の郵便番号'
COL_ADDR    = '元の住所'
COL_NAME    = '氏名'
COL_TICKET  = 'イベントチケット'
COL_OPT     = 'イベントオプションチケット'
COL_CHECK   = 'チェックOK'
COL_EXCLUDE = '対象外'

def clean_text(text):
    if pd.isna(text) or str(text).lower() == 'nan' or str(text).strip() == '':
        return "なし"
    text = str(text)
    text = re.sub(r'[\u200e\u200b\u208b]', '', text)
    return text

def extract_ticket_info(row):
    text = str(row[COL_TICKET])
    count = 0
    match = re.search(r'[:：](\d+)枚', text)
    if match:
        count = int(match.group(1))

    has_a = "参加権A" in text
    has_b = "参加権B" in text
    has_c = "参加権C" in text

    matches = [has_a, has_b, has_c]
    if sum(matches) == 1:
        if has_a: category = "参加権A"
        elif has_b: category = "参加権B"
        else: category = "参加権C"
    else:
        category = "複合・その他"

    return pd.Series([category, count])

def load_and_merge_csv(directory):
    all_files = glob.glob(os.path.join(directory, "*.csv"))
    if not all_files:
        print("【！】CSVファイルが見つかりません。")
        return None
    df_list = []
    for file in all_files:
        try:
            try:
                temp_df = pd.read_csv(file, encoding='utf-8-sig')
            except UnicodeDecodeError:
                temp_df = pd.read_csv(file, encoding='shift_jis')
            df_list.append(temp_df)
            print(f"読み込み成功: {os.path.basename(file)}")
        except Exception as e:
            print(f"【！】読み込み失敗: {file} ({e})")
    return pd.concat(df_list, ignore_index=True) if df_list else None

def create_pdf_check_sheet(df, category_name):
    if df.empty:
        return

    if category_name == "複合・その他":
        row_height = 45
        items_per_col = 5
        line_h = 4.0
        font_size_body = 7.5
    else:
        row_height = 27
        items_per_col = 10
        line_h = 3.6
        font_size_body = 7

    items_per_page = items_per_col * 2
    total_pages = math.ceil(len(df) / items_per_page)

    output_path = os.path.join(OUTPUT_DIR, f"check_sheet_{category_name}.pdf")
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=False)

    if os.path.exists(FONT_PATH):
        pdf.add_font('JP', '', FONT_PATH)
        pdf.add_font('JP', 'B', FONT_PATH)
    else:
        print(f"【！】フォントなし: {FONT_PATH}")
        return

    col_width = 90
    col_space = 10

    def add_header(title, current_page, total_p):
        pdf.add_page()
        pdf.set_font('JP', 'B', 14)
        header_text = f"チェック用名簿 【{title}】 ({current_page} / {total_p} ページ)"
        pdf.cell(0, 10, header_text, align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        return pdf.get_y()

    current_p_num = 1
    start_y_base = add_header(category_name, current_p_num, total_pages)

    for i, (_, row) in enumerate(df.iterrows()):
        item_in_page = i % items_per_page
        col_idx = item_in_page // items_per_col
        row_idx = item_in_page % items_per_col

        if i > 0 and item_in_page == 0:
            current_p_num += 1
            start_y_base = add_header(category_name, current_p_num, total_pages)

        curr_x = 10 + (col_idx * (col_width + col_space))
        curr_y = start_y_base + (row_idx * row_height)

        pdf.set_draw_color(0)
        pdf.rect(curr_x, curr_y + 1, 5, 5)

        pdf.set_xy(curr_x, curr_y + 6)
        pdf.set_font('JP', '', 7)
        pdf.cell(5, 4, str(row[COL_SERIAL]), align='C')

        inner_x = curr_x + 7
        content_w = col_width - 8

        pdf.set_xy(inner_x, curr_y)
        pdf.set_font('JP', 'B', 9)
        pdf.cell(content_w, line_h, f"【{row[COL_ID]}】", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        pdf.set_x(inner_x)
        pdf.set_font('JP', '', 8.5)
        pdf.cell(content_w, line_h, f"〒{row[COL_ZIP]}    {row[COL_NAME]}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        pdf.set_font('JP', '', font_size_body)
        for col in [COL_ADDR, COL_TICKET, COL_OPT]:
            val = clean_text(row[col])
            if val != "なし":
                pdf.set_x(inner_x)
                pdf.multi_cell(content_w, line_h, val)

        pdf.set_draw_color(220, 220, 220)
        pdf.line(curr_x, curr_y + row_height - 1.5, curr_x + col_width, curr_y + row_height - 1.5)

    pdf.output(output_path)
    print(f"✅ PDF生成完了: {os.path.basename(output_path)}")

def main():
    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    df = load_and_merge_csv(INPUT_DIR)
    if df is not None:
        # --- 1. 「対象外」フィルタと出力 ---
        if COL_EXCLUDE in df.columns:
            # TRUE（除外対象）を抽出して保存
            df_target_out = df[df[COL_EXCLUDE].astype(str).str.upper() == 'TRUE'].copy()
            if not df_target_out.empty:
                df_target_out.to_csv(OUTPUT_CSV_TARGET, index=False, encoding='utf-8-sig')
                print(f"✅ 対象外データを保存しました: {os.path.basename(OUTPUT_CSV_TARGET)} ({len(df_target_out)}件)")

            # TRUEでない人だけを次に進める
            df = df[df[COL_EXCLUDE].astype(str).str.upper() != 'TRUE'].copy()

        # --- 2. 「チェックOK」フィルタと出力 ---
        if COL_CHECK in df.columns:
            # TRUEでない（未完了）人を抽出して保存
            df_check_not_ok = df[df[COL_CHECK].astype(str).str.upper() != 'TRUE'].copy()
            if not df_check_not_ok.empty:
                df_check_not_ok.to_csv(OUTPUT_CSV_NOT_OK, index=False, encoding='utf-8-sig')
                print(f"✅ チェック未完了データを保存しました: {os.path.basename(OUTPUT_CSV_NOT_OK)} ({len(df_check_not_ok)}件)")

            # TRUEの人だけを名簿対象にする
            df = df[df[COL_CHECK].astype(str).str.upper() == 'TRUE'].copy()

        # --- 3. データ補完と全統合CSV保存 ---
        required_cols = [COL_ID, COL_ZIP, COL_ADDR, COL_NAME, COL_TICKET, COL_OPT]
        for col in required_cols:
            if col not in df.columns:
                df[col] = "なし"

        df[['カテゴリ', '枚数']] = df.apply(extract_ticket_info, axis=1)
        df.to_csv(OUTPUT_CSV_ALL, index=False, encoding='utf-8-sig')

        # --- 4. カテゴリ別出力フロー ---
        categories = ["参加権A", "参加権B", "参加権C", "複合・その他"]
        for cat in categories:
            sub_df = df[df['カテゴリ'] == cat].copy()
            if not sub_df.empty:
                if cat == "参加権C":
                    sub_df = sub_df.sort_values(by=[COL_ID], ascending=[True])
                else:
                    sub_df = sub_df.sort_values(by=['枚数', COL_ID], ascending=[False, True])

                sub_df.insert(0, COL_SERIAL, range(1, len(sub_df) + 1))

                cat_csv_path = os.path.join(OUTPUT_DIR, f"data_{cat}.csv")
                sub_df.to_csv(cat_csv_path, index=False, encoding='utf-8-sig')
                create_pdf_check_sheet(sub_df, cat)

if __name__ == "__main__":
    main()