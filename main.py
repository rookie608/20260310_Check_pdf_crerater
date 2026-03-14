import pandas as pd
import glob
import os
import re
import math
import unicodedata
from fpdf import FPDF
from fpdf.enums import XPos, YPos

# ==========================================
# 1. 基本設定（パス・項目名）
# ==========================================
BASE_DIR = "/Users/suzukiarato/PycharmProjects/20260310_Check_pdf_crerater"
INPUT_DIR = os.path.join(BASE_DIR, 'csv_files')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')

# 出力ファイル名
OUTPUT_CSV_ALL = os.path.join(OUTPUT_DIR, 'merged_data_all.csv')
OUTPUT_CSV_EXCLUDED = os.path.join(OUTPUT_DIR, 'excluded_all.csv')  # 除外データを1つに統合

FONT_NAME = 'ipaexg.ttf'
FONT_PATH = os.path.join(BASE_DIR, FONT_NAME)

# 内部共通項目名
COL_SERIAL = '通し番号'
COL_ID = '受付番号'
COL_ZIP = '郵便番号'
COL_ADDR = '住所'
COL_NAME = '氏名'
COL_TICKET = 'チケット情報'
COL_OPT = 'オプション情報'

# 学長招待客リスト専用の元列名
GUEST_COL_TICKET_A = '人数(チケットAのみ)'
GUEST_COL_WAITING = '控室入れる券'
GUEST_COL_ADDRESS = '郵送先住所'


# ==========================================
# 2. 補助関数
# ==========================================

def normalize_text(text):
    if not isinstance(text, str): return text
    return unicodedata.normalize('NFKC', text).strip()


def clean_text(text):
    if pd.isna(text) or str(text).lower() == 'nan' or str(text).strip() == '':
        return "なし"
    text = str(text)
    text = re.sub(r'[\u200e\u200b\u208b]', '', text)
    return text


def parse_guest_address_v3(full_text):
    if pd.isna(full_text) or str(full_text).strip() == "":
        return "なし", ["なし"], "なし"
    lines = [l.strip() for l in str(full_text).split('\n') if l.strip()]
    if not lines: return "なし", ["なし"], "なし"

    zip_code = "なし"
    name = "なし"

    if re.match(r'^\d{3}-\d{4}', lines[0]):
        zip_code = lines[0]
        remaining = lines[1:]
    else:
        remaining = lines

    filtered = []
    for l in remaining:
        clean_l = re.sub(r'[-\s\(\)]', '', l)
        if clean_l.isdigit() and clean_l.startswith('0') and 9 <= len(clean_l) <= 11:
            continue
        filtered.append(l)

    if filtered:
        name = filtered[-1]
        address_parts = filtered[:-1]
        if not address_parts: address_parts = ["なし"]
    else:
        address_parts = ["住所不明"]

    return zip_code, address_parts, name


def extract_ticket_info(row):
    text = str(row.get('チケット情報', ''))
    count = 0
    match = re.search(r'[:：](\d+)枚', text)
    if match: count = int(match.group(1))
    has_a, has_b, has_c = "参加権A" in text, "参加権B" in text, "参加権C" in text
    if sum([has_a, has_b, has_c]) == 1:
        category = "参加権A" if has_a else "参加権B" if has_b else "参加権C"
    else:
        category = "複合・その他"
    return pd.Series([category, count])


# ==========================================
# 3. PDF生成ロジック
# ==========================================

def create_pdf_check_sheet(df, category_name):
    if df.empty: return

    if category_name == "学長招待客":
        row_height, items_per_col, line_h = 45, 5, 5.2
        f_size_id, f_size_name, f_size_body = 11, 10, 9.5
    elif category_name == "複合・その他":
        row_height, items_per_col, line_h = 45, 5, 4.0
        f_size_id, f_size_name, f_size_body = 9, 8.5, 7.5
    else:
        row_height, items_per_col, line_h = 27, 10, 3.6
        f_size_id, f_size_name, f_size_body = 9, 8.5, 7.0

    total_pages = math.ceil(len(df) / (items_per_col * 2))
    output_path = os.path.join(OUTPUT_DIR, f"check_sheet_{category_name}.pdf")
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.set_auto_page_break(auto=False)
    if os.path.exists(FONT_PATH):
        pdf.add_font('JP', '', FONT_PATH);
        pdf.add_font('JP', 'B', FONT_PATH)
    else:
        return

    def add_header(title, curr_p, total_p):
        pdf.add_page()
        pdf.set_font('JP', 'B', 14)
        pdf.cell(0, 10, f"チェック用名簿 【{title}】 ({curr_p} / {total_p} ページ)", align='C', new_x=XPos.LMARGIN,
                 new_y=YPos.NEXT)
        return pdf.get_y()

    curr_p, start_y = 1, 0
    start_y = add_header(category_name, curr_p, total_pages)

    for i, (_, row) in enumerate(df.iterrows()):
        item_idx = i % (items_per_col * 2)
        if i > 0 and item_idx == 0:
            curr_p += 1
            start_y = add_header(category_name, curr_p, total_pages)

        # z 字 (左右交互) に並べる
        row_idx = item_idx // 2  # 何段目か（0〜items_per_col-1）
        col_idx = item_idx % 2  # 0:左, 1:右

        cx = 10 + (col_idx * 100)
        cy = start_y + (row_idx * row_height)

        pdf.set_draw_color(0);
        pdf.rect(cx, cy + 1, 5, 5)
        pdf.set_xy(cx, cy + 6);
        pdf.set_font('JP', '', 7);
        pdf.cell(5, 4, str(row[COL_SERIAL]), align='C')

        ix, cw = cx + 7, 82
        pdf.set_xy(ix, cy);
        pdf.set_font('JP', 'B', f_size_id)
        pdf.cell(cw, line_h, f"【{str(row.get(COL_ID, '招待'))}】", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_x(ix);
        pdf.set_font('JP', '', f_size_name)
        pdf.cell(cw, line_h, f"〒{row[COL_ZIP]}    {row[COL_NAME]}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_font('JP', '', f_size_body)
        for c in [COL_ADDR, COL_TICKET, COL_OPT]:
            val = clean_text(row.get(c, "なし"))
            if val != "なし":
                pdf.set_x(ix);
                pdf.multi_cell(cw, line_h, val)

        pdf.set_draw_color(220, 220, 220)
        pdf.line(cx, cy + row_height - 1.5, cx + 90, cy + row_height - 1.5)

    pdf.output(output_path)
    print(f"✅ PDF生成完了: {os.path.basename(output_path)}")


# ==========================================
# 4. メイン処理
# ==========================================

def main():
    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    all_files = glob.glob(os.path.join(INPUT_DIR, "*.csv"))

    guest_files = [f for f in all_files if "学長招待客リスト" in os.path.basename(f)]
    normal_files = [f for f in all_files if "学長招待客リスト" not in os.path.basename(f)]

    all_excluded_list = []  # 全ての除外データを蓄積するリスト

    # --- 1. 学長招待客リスト処理 ---
    for gf in guest_files:
        try:
            try:
                gdf = pd.read_csv(gf, encoding='utf-8-sig')
            except:
                gdf = pd.read_csv(gf, encoding='shift_jis')
            gdf.columns = [normalize_text(c) for c in gdf.columns]

            # フィルタリング（除外データを保存用に分ける）
            col_ex = normalize_text('対象外')
            col_ok = normalize_text('最終チェックOK')

            if col_ex in gdf.columns:
                ex_target = gdf[gdf[col_ex].astype(str).str.upper() == 'TRUE'].copy()
                ex_target['除外理由'] = '対象外'
                all_excluded_list.append(ex_target)
                gdf = gdf[gdf[col_ex].astype(str).str.upper() != 'TRUE']

            if col_ok in gdf.columns:
                ex_not_ok = gdf[gdf[col_ok].astype(str).str.upper() != 'TRUE'].copy()
                ex_not_ok['除外理由'] = 'チェック未完了'
                all_excluded_list.append(ex_not_ok)
                gdf = gdf[gdf[col_ok].astype(str).str.upper() == 'TRUE']

            addr_col = normalize_text(GUEST_COL_ADDRESS)
            if addr_col in gdf.columns:
                gdf = gdf[gdf[addr_col].astype(str).str.strip().replace('nan', '') != ''].copy()
                parsed = gdf[addr_col].apply(parse_guest_address_v3)
                gdf[COL_ZIP] = [x[0] for x in parsed]
                gdf[COL_NAME] = [x[2] for x in parsed]
                addr_lists = [x[1] for x in parsed]
                gdf[COL_ADDR] = [" ".join(l) for l in addr_lists]

                # CSV用住所列作成
                max_l = max(len(l) for l in addr_lists) if addr_lists else 1
                addr_cols = []
                for j in range(max_l):
                    cn = f'住所{j + 1}';
                    gdf[cn] = [l[j] if j < len(l) else "" for l in addr_lists]
                    addr_cols.append(cn)

                gdf = gdf[~gdf[COL_ADDR].isin(["なし", "住所不明"])]

                col_a = normalize_text(GUEST_COL_TICKET_A);
                col_w = normalize_text(GUEST_COL_WAITING)
                gdf['カテゴリ'] = "学長招待客"
                gdf['枚数'] = pd.to_numeric(gdf[col_a], errors='coerce').fillna(0).astype(
                    int) if col_a in gdf.columns else 0
                gdf[COL_TICKET] = "参加権A: " + gdf[col_a].astype(str) + "枚" if col_a in gdf.columns else "参加権A: なし"
                gdf[COL_OPT] = "控え室権: " + gdf[col_w].astype(str) + "枚" if col_w in gdf.columns else "控え室権: なし"
                gdf[COL_ID] = "招待客"

                if not gdf.empty:
                    gdf.insert(0, COL_SERIAL, range(1, len(gdf) + 1))
                    out_cols = [COL_SERIAL, COL_ID, COL_ZIP] + addr_cols + [COL_NAME, COL_TICKET, COL_OPT, 'カテゴリ',
                                                                            '枚数']
                    gdf[out_cols].to_csv(os.path.join(OUTPUT_DIR, "data_学長招待客.csv"), index=False,
                                         encoding='utf-8-sig')
                    create_pdf_check_sheet(gdf, "学長招待客")

        except Exception as e:
            print(f"【！】招待客リストエラー: {e}")

    # --- 2. 通常CSV処理 ---
    df_list = []
    for f in normal_files:
        try:
            try:
                tmp = pd.read_csv(f, encoding='utf-8-sig')
            except:
                tmp = pd.read_csv(f, encoding='shift_jis')
            df_list.append(tmp)
        except:
            pass

    if df_list:
        df = pd.concat(df_list, ignore_index=True)
        df.columns = [normalize_text(c) for c in df.columns]

        # フィルタリング（統合保存用）
        c_ex, c_ok = normalize_text('対象外'), normalize_text('最終チェックOK')
        if c_ex in df.columns:
            ex_t = df[df[c_ex].astype(str).str.upper() == 'TRUE'].copy()
            ex_t['除外理由'] = '対象外'
            all_excluded_list.append(ex_t)
            df = df[df[c_ex].astype(str).str.upper() != 'TRUE']
        if c_ok in df.columns:
            ex_n = df[df[c_ok].astype(str).str.upper() != 'TRUE'].copy()
            ex_n['除外理由'] = 'チェック未完了'
            all_excluded_list.append(ex_n)
            df = df[df[c_ok].astype(str).str.upper() == 'TRUE']

        m_cols = {normalize_text('元の郵便番号'): COL_ZIP, normalize_text('元の住所'): COL_ADDR,
                  normalize_text('氏名'): COL_NAME, normalize_text('イベントチケット'): 'チケット情報',
                  normalize_text('イベントオプションチケット'): COL_OPT}
        df = df.rename(columns=m_cols)
        df[['カテゴリ', '枚数']] = df.apply(extract_ticket_info, axis=1)
        df = df.rename(columns={'チケット情報': COL_TICKET})
        df.to_csv(OUTPUT_CSV_ALL, index=False, encoding='utf-8-sig')

        for cat in ["参加権A", "参加権B", "参加権C", "複合・その他"]:
            sub = df[df['カテゴリ'] == cat].copy()
            if not sub.empty:
                sub = sub.sort_values(by=[COL_ID]) if cat == "参加権C" else sub.sort_values(by=['枚数', COL_ID],
                                                                                            ascending=[False, True])
                sub.insert(0, COL_SERIAL, range(1, len(sub) + 1))
                sub.to_csv(os.path.join(OUTPUT_DIR, f"data_{cat}.csv"), index=False, encoding='utf-8-sig')
                create_pdf_check_sheet(sub, cat)

    # --- 3. 除外データの統合出力 ---
    if all_excluded_list:
        combined_excluded = pd.concat(all_excluded_list, ignore_index=True)
        combined_excluded.to_csv(OUTPUT_CSV_EXCLUDED, index=False, encoding='utf-8-sig')
        print(f"✅ 除外データを統合保存しました: {os.path.basename(OUTPUT_CSV_EXCLUDED)}")


if __name__ == "__main__":
    main()