import streamlit as st
import pandas as pd
import re
import calendar
from datetime import datetime
import traceback

# ==========================================
# 1. データ抽出・クレンジング用関数
# ==========================================
def clean_student_name(name):
    if name is None or pd.isna(name): return None
    if name == "チューター": return "チューター"
    parts = re.split(r'[A-Za-zＡ-Ｚａ-ｚ][0-9０-９]', str(name))
    cleaned_name = parts[0].strip()
    return cleaned_name if cleaned_name else str(name)

def get_val(row, idx):
    if idx < len(row):
        v = row.iloc[idx]
        if pd.notna(v) and str(v).strip() != "" and str(v).strip().lower() != "nan":
            return str(v).strip()
    return None

def process_booth_excel(file_data):
    # Webでも確実に開けるよう engine="openpyxl" を明記
    all_sheets = pd.read_excel(file_data, sheet_name=None, header=None, engine="openpyxl")
    final_data = []
    
    for sheet_name, df in all_sheets.items():
        if sheet_name == "今月ベース": continue
        for row_idx in range(len(df)):
            row_values = [str(x) for x in df.iloc[row_idx].tolist()]
            # 日付行の特定
            if '年' in row_values and '月' in row_values and '日' in row_values:
                date_indices = [i for i, x in enumerate(row_values) if '年' in x]
                for i, start_col in enumerate(date_indices):
                    try:
                        end_col = date_indices[i + 1] if i + 1 < len(date_indices) else len(df.columns)
                        year = get_val(df.iloc[row_idx], start_col - 1) or get_val(df.iloc[row_idx], start_col)
                        month = get_val(df.iloc[row_idx], start_col + 1)
                        day = get_val(df.iloc[row_idx], start_col + 3)
                        dow = get_val(df.iloc[row_idx], start_col + 5)
                        date_str = f"{year}/{month}/{day}"
                        
                        time_row_idx = row_idx + 2
                        times_row = df.iloc[time_row_idx, start_col:end_col]
                        time_slots = []
                        for c_offset, t_val in enumerate(times_row):
                            if isinstance(t_val, str) and ('～' in t_val or '~' in t_val):
                                actual_col = start_col + c_offset
                                time_slots.append({'time': t_val, 'col1': actual_col, 'col2': actual_col + 1})
                                
                        current_row = time_row_idx + 1
                        while current_row < len(df):
                            row_list = df.iloc[current_row].astype(str).tolist()
                            row_head = str(df.iloc[current_row, 0]).strip()
                            label = str(df.iloc[current_row, 1]).strip()
                            
                            if '年' in "".join(row_list): break
                            
                            if '生徒' in label:
                                student_row = df.iloc[current_row]
                                teacher_row = df.iloc[current_row + 1]
                                for ts in time_slots:
                                    s1 = clean_student_name(get_val(student_row, ts['col1']))
                                    s2 = clean_student_name(get_val(student_row, ts['col2']))
                                    t1 = get_val(teacher_row, ts['col1'])
                                    t2 = get_val(teacher_row, ts['col2'])
                                    if t1 and not t2: t2 = t1
                                    elif t2 and not t1: t1 = t2
                                    
                                    if s1: final_data.append({"日付": date_str, "曜日": dow, "時間": ts['time'], "講師": t1, "内容": s1})
                                    if s2: final_data.append({"日付": date_str, "曜日": dow, "時間": ts['time'], "講師": t2, "内容": s2})
                                current_row += 2
                                
                            elif 'チューター' in label or 'チューター' in row_head:
                                row_options = [df.iloc[current_row], df.iloc[current_row + 1]]
                                for ts in time_slots:
                                    for r in row_options:
                                        t_name = get_val(r, ts['col1']) or get_val(r, ts['col2'])
                                        if t_name and "生徒" not in t_name and "チューター" not in t_name:
                                            final_data.append({"日付": date_str, "曜日": dow, "時間": ts['time'], "講師": t_name, "内容": "チューター"})
                                            break
                                current_row += 2
                            else: 
                                current_row += 1
                    except Exception as inner_e:
                        # ここでエラーを無視せず、画面に出力して原因を特定する
                        st.error(f"データ解析中に内部エラー発生（行 {row_idx}）: {inner_e}")
                        continue
    return pd.DataFrame(final_data)

def get_start_minutes(time_str):
    normalized = str(time_str).replace('：', ':')
    match = re.search(r'(\d{1,2}):(\d{1,2})', normalized)
    return int(match.group(1)) * 60 + int(match.group(2)) if match else 9999

# ==========================================
# 2. Webアプリの画面デザインと処理
# ==========================================
st.set_page_config(page_title="講師別カレンダー作成ツール", layout="centered")

st.title("📅 講師別カレンダー作成ツール")
st.write("ブース表（Excel）をアップロードすると、印刷用のカレンダー（HTML）を作成します。")

uploaded_file = st.file_uploader("ブース表のExcelファイルを選択してください", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("データを読み込んでカレンダーを作成中..."):
        try:
            result_df = process_booth_excel(uploaded_file)
            
            if result_df.empty:
                st.error("データが抽出できませんでした。ファイル形式やシート名を確認してください。")
                
                # --- デバッグ表示ゾーン（Webサーバーがエクセルをどう見ているか） ---
                st.warning("⚠️【原因調査用データ】Webサーバーがエクセルを以下のように読み込んでしまっています。")
                uploaded_file.seek(0) # ファイルを先頭に戻す
                debug_sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None, engine="openpyxl")
                for sheet_name, df in debug_sheets.items():
                    if sheet_name != "今月ベース":
                        st.write(f"▼ シート名: {sheet_name} の最初の15行")
                        st.dataframe(df.head(15))
                        break # 1シートだけ表示してストップ
                # --------------------------------------------------------
                
            else:
                result_df['date_obj'] = pd.to_datetime(result_df['日付']).dt.date
                all_times = sorted(result_df['時間'].unique(), key=get_start_minutes)
                target_date = result_df['date_obj'].mode()[0]
                target_year, target_month = target_date.year, target_date.month
                cal = calendar.Calendar(firstweekday=calendar.SUNDAY)
                weeks = cal.monthdatescalendar(target_year, target_month)
                dow_chars = ["月", "火", "水", "木", "金", "土", "日"]

                html_content = f"""
                <html>
                <head>
                <meta charset="utf-8">
                <style>
                    @page {{ size: A4 portrait; margin: 5mm; }}
                    body {{ font-family: 'Hiragino Kaku Gothic ProN', 'Meiryo', sans-serif; font-size: 6.5pt; line-height: 1.1; color: #333; margin: 0; }}
                    .teacher-page {{ page-break-after: always; width: 100%; overflow: hidden; }}
                    h1 {{ text-align: center; color: #2e5a27; border-bottom: 1px solid #2e5a27; margin: 0 0 5px 0; font-size: 11pt; }}
                    table {{ width: 100%; border-collapse: collapse; table-layout: fixed; margin-bottom: 3px; }}
                    th, td {{ border: 0.5px solid #666; padding: 1px 2px; text-align: center; vertical-align: middle; word-break: normal; }}
                    th {{ background-color: #e2efda; font-weight: bold; }}
                    .content-cell {{ font-size: 5.5pt; }}
                    .time-col {{ background-color: #f2f2f2 !important; width: 70px !important; white-space: nowrap; font-size: 6pt; font-weight: bold; }}
                    th:nth-child(2), td:nth-child(2) {{ background-color: #fff5f5; color: #c00; }}
                    th:nth-child(8), td:nth-child(8) {{ background-color: #f5faff; color: #00c; }}
                    .other-month {{ color: #bbb; }}
                    .day-header {{ font-weight: normal; font-size: 6pt; }}
                    @media print {{ * {{ -webkit-print-color-adjust: exact !important; color-adjust: exact !important; }} }}
                </style>
                </head>
                <body>
                """

                teachers = sorted([t for t in result_df['講師'].unique() if pd.notna(t) and str(t).strip() != ""])

                for teacher in teachers:
                    t_df = result_df[result_df['講師'] == teacher]
                    html_content += f'<div class="teacher-page"><h1>{target_year}年{target_month}月 講師勤務表：{teacher} 先生</h1>'
                    
                    for week in weeks:
                        html_content += '<table>'
                        html_content += '<tr><th class="time-col">時間帯</th>'
                        for d in week:
                            m_class = "" if d.month == target_month else 'class="other-month"'
                            html_content += f'<th {m_class}>{d.month}/{d.day} <span class="day-header">({dow_chars[d.weekday()]})</span></th>'
                        html_content += '</tr>'
                        
                        for t in all_times:
                            html_content += f'<tr><td class="time-col">{t}</td>'
                            for d in week:
                                c_list = t_df[(t_df['date_obj'] == d) & (t_df['時間'] == t)]['内容'].tolist()
                                c_list = list(dict.fromkeys(c_list))
                                cell_val = " / ".join(c_list) if c_list else ""
                                html_content += f'<td class="content-cell">{cell_val}</td>'
                            html_content += '</tr>'
                        html_content += '</table>'
                    html_content += '</div>'

                html_content += "</body></html>"

                st.success(f"✅ {target_year}年{target_month}月のカレンダー作成が完了しました！")
                
                file_name = f"講師別勤務表_{target_year}年{target_month}月.html"
                st.download_button(
                    label="📥 カレンダー（HTML）をダウンロード",
                    data=html_content,
                    file_name=file_name,
                    mime="text/html"
                )

        except Exception as e:
            st.error(f"深刻なエラーが発生しました: {e}")
            st.code(traceback.format_exc())
