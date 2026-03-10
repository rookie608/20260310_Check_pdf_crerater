import pandas as pd
import glob
import os
import re
from fpdf import FPDF
from fpdf.enums import XPos, YPos

# --- 設定項目 ---
BASE_DIR = "/Users/suzukiarato/PycharmProjects/20260310_Check_pdf_crerater"
INPUT_DIR = os.path.join(BASE_DIR, 'csv_files')
OUTPUT_CSV = os.path.join(BASE_DIR, 'merged_data.csv')
OUTPUT_PDF = os.path.join(BASE_DIR, 'check_sheet.pdf')

FONT_NAME = 'ipaexg.ttf'
FONT_PATH = os.path.join(BASE_DIR, FONT_NAME)


def clean_text(text):
    """NaN対策と不要な記号の除去"""
    if pd.isna(text) or str(text).lower() == 'nan' or str(text).strip() == '':
        return "なし"
    text = str(text)
    # PDFエラーの原因になる制御文字を削除
    text = re.sub(r'[\u200e\u200b\u208b]', '', text)
    return text


def load_and_merge_csv(directory):
    all_files = glob.glob(os.path.join(directory, "*.csv"))
    if not all_files:
        print("【！】csv_files フォルダにCSVファイルが見つかりません。")
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
            print(f"【！】ファイル {file} の読み込みに失敗しました: {e}")

    return pd.concat(df_list, ignore_index=True) if df_list else None


def create_pdf_check_sheet(df):
    target_cols = ['受付番号', '郵便番号', '住所', '名前', 'イベントチケット', 'イベントオプションチケット']
    for col in target_cols:
        if col not in df.columns:
            df[col] = "なし"
        df[col] = df[col].apply(clean_text)

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=False)

    if os.path.exists(FONT_PATH):
        pdf.add_font('JP', '', FONT_PATH)
        pdf.add_font('JP', 'B', FONT_PATH)
    else:
        print(f"【！】エラー: フォントが見つかりません: {FONT_PATH}")
        return

    # 1ページ20名用 (2列 x 10行)
    items_per_col = 10
    col_width = 90
    col_space = 10
    row_height = 27  # 1人あたりの高さ
    line_h = 3.6  # 行間の高さ

    pdf.add_page()
    pdf.set_font('JP', 'B', 14)
    pdf.cell(0, 10, "チェック用名簿", align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    start_y_base = pdf.get_y()

    for i, (_, row) in enumerate(df.iterrows()):
        item_in_page = i % (items_per_col * 2)
        col_idx = item_in_page // items_per_col
        row_idx = item_in_page % items_per_col

        if i > 0 and item_in_page == 0:
            pdf.add_page()
            pdf.set_font('JP', 'B', 14)
            pdf.cell(0, 10, f"チェック用名簿", align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            start_y_base = pdf.get_y()

        curr_x = 10 + (col_idx * (col_width + col_space))
        curr_y = start_y_base + (row_idx * row_height)

        # 1. チェックボックス
        pdf.set_draw_color(0)
        pdf.rect(curr_x, curr_y + 1, 5, 5)

        inner_x = curr_x + 7
        content_w = col_width - 8

        # 2. 受付番号（太字）
        pdf.set_xy(inner_x, curr_y)
        pdf.set_font('JP', 'B', 9)
        pdf.cell(content_w, line_h, f"【{row['受付番号']}】", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        # 3. 郵便番号 & 名前（ラベルなし、間にスペースを確保）
        pdf.set_x(inner_x)
        pdf.set_font('JP', '', 8.5)
        # 郵便番号と名前の間に全角スペース2つ分程度の隙間を挿入
        pdf.cell(content_w, line_h, f"〒{row['郵便番号']}    {row['名前']}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        # 4. 住所（ラベルなし）
        if row['住所'] != "なし":
            pdf.set_x(inner_x)
            pdf.set_font('JP', '', 7)
            pdf.multi_cell(content_w, line_h, row['住所'])

        # 5. チケット情報（ラベルなし）
        if row['イベントチケット'] != "なし":
            pdf.set_x(inner_x)
            pdf.set_font('JP', '', 7)
            pdf.multi_cell(content_w, line_h, row['イベントチケット'])

        # 6. オプション情報（ラベルなし）
        if row['イベントオプションチケット'] != "なし":
            pdf.set_x(inner_x)
            pdf.multi_cell(content_w, line_h, row['イベントオプションチケット'])

        # 7. 区切り線
        pdf.set_draw_color(220, 220, 220)
        pdf.line(curr_x, curr_y + row_height - 1.5, curr_x + col_width, curr_y + row_height - 1.5)

    pdf.output(OUTPUT_PDF)
    print(f"✅ PDFを出力しました: {OUTPUT_PDF} (20名/2列/ラベルなし版)")


def main():
    if not os.path.exists(INPUT_DIR):
        os.makedirs(INPUT_DIR)
        print(f"フォルダを作成しました。{INPUT_DIR} にCSVを入れてください。")
        return
    df = load_and_merge_csv(INPUT_DIR)
    if df is not None:
        df.to_csv(OUTPUT_CSV, index=False, encoding='utf-8-sig')
        create_pdf_check_sheet(df)


if __name__ == "__main__":
    main()