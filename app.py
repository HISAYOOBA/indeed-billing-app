import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime
import pytz
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(
    page_title="Indeed請求明細ジェネレーター",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;700&family=Inter:wght@400;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Noto Sans JP', sans-serif; }
[data-testid="stSidebar"] { display: none !important; }
[data-testid="collapsedControl"] { display: none !important; }
.title-bar { background: linear-gradient(135deg, #1F4E79, #2E75B6); padding: 20px 28px 16px 28px; border-radius: 12px; margin-bottom: 10px; }
.main-title { font-family: 'Inter', sans-serif; font-size: 1.8rem; font-weight: 700; color: #FFFFFF; margin: 0; padding: 0; }
.sub-title { color: #CBD5E1; font-size: 0.88rem; margin-top: 4px; }
.section-card { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px; padding: 12px 18px; margin-bottom: 10px; }
.stButton > button { background-color: #1F4E79; color: white; border: none; border-radius: 8px; font-size: 1rem; font-weight: 700; padding: 10px 28px; width: 100%; transition: all 0.2s; }
.stButton > button:hover { background-color: #2E75B6; transform: translateY(-1px); box-shadow: 0 4px 12px rgba(31,78,121,0.3); }
.result-box { background: #e8f5e9; border: 1px solid #a5d6a7; border-radius: 8px; padding: 14px; color: #2e7d32; font-weight: 600; margin-top: 10px; }
.error-box { background: #ffebee; border: 1px solid #ef9a9a; border-radius: 8px; padding: 14px; color: #c62828; margin-top: 10px; }
.login-box { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px; padding: 12px 18px; margin-bottom: 10px; }
.login-title { font-size: 1.2rem; font-weight: 700; color: #1F4E79; margin-bottom: 10px; }
hr { margin: 8px 0 !important; border-color: #e2e8f0 !important; }
</style>
""", unsafe_allow_html=True)

# ==================== パスワード認証 ====================
CORRECT_PASSWORD = st.secrets.get("APP_PASSWORD", "rs5489-4191")

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("""
    <div class="title-bar">
        <div class="main-title">📊 Indeed請求明細ジェネレーター</div>
        <div class="sub-title">Google DriveのデータからクライアントごとのIndeed請求明細Excelを自動生成します</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="login-box">', unsafe_allow_html=True)
    st.markdown('<div class="login-title">🔐 ログイン</div>', unsafe_allow_html=True)
    password_input = st.text_input("パスワードを入力してください", type="password")
    if st.button("ログイン", use_container_width=True):
        if password_input == CORRECT_PASSWORD:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("パスワードが正しくありません")
    st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

# ==================== Google Drive / Sheets 接続 ====================
def get_credentials():
    creds_info = dict(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
    if "private_key" in creds_info:
        creds_info["private_key"] = creds_info["private_key"].replace("\\n", "\n")
    return service_account.Credentials.from_service_account_info(
        creds_info,
        scopes=[
            "https://www.googleapis.com/auth/drive.readonly",
            "https://www.googleapis.com/auth/spreadsheets"
        ]
    )

def get_drive_service():
    return build("drive", "v3", credentials=get_credentials())

def get_sheets_service():
    return build("sheets", "v4", credentials=get_credentials())

def list_files_in_folder(service, folder_id):
    results = service.files().list(
        q=f"'{folder_id}' in parents and trashed=false",
        fields="files(id, name, mimeType)", orderBy="name"
    ).execute()
    return results.get("files", [])

def download_file(service, file_id):
    request = service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    buf.seek(0)
    return buf

# ==================== Indeedデータ 万能ローダー ====================
# どの形式で来ても統一された縦持ちDataFrameに正規化する。
# 対応: UTF-8 / UTF-8-sig / UTF-16(TSV) / cp932、csv / xlsx、
#       縦持ち(メジャー ネーム) / 横持ち(合計費用列)、セル結合、"￥240,000"表記

SCHEMA = [
    '対象年月', '契約代理店コード', '契約代理店名', '委託先コード', '委託先社名', '企業名',
    '企業ランキング (前四半期合計費用降順)', 'アカウントID', 'アカウント名', 'アカウント - 種別',
    'アカウント - 利用状況 (当四半期)', 'アカウントメール', 'キャンペーンID', 'キャンペーン名',
    'フィード', 'フィード種別', 'キャンペーンの目標', 'キャンペーンステータス',
    '最新のキャンペーンステータス', '予算種別', '予算総額', '利用済み予算', '予算消化率',
    'キャンペーン開始日', 'キャンペーン終了日 (指定した日付)', 'キャンペーン終了日 (目標とする日付)',
    'キャンペーン進行ステータス', 'メジャー ネーム', 'メジャー バリュー',
]

FFILL_COLS = [
    '対象年月', '契約代理店コード', '契約代理店名', '委託先コード', '委託先社名',
    '企業名', '企業ランキング (前四半期合計費用降順)', 'アカウントID', 'アカウント名',
    'アカウント - 種別', 'アカウント - 利用状況 (当四半期)', 'アカウントメール',
]

def parse_yen(val):
    if pd.isna(val):
        return 0
    s = re.sub(r'[￥¥,\s]', '', str(val))
    try:
        return int(float(s))
    except (ValueError, TypeError):
        return 0

def _read_any(raw_bytes, filename=''):
    name = (filename or '').lower()
    if name.endswith(('.xlsx', '.xls')):
        return pd.read_excel(io.BytesIO(raw_bytes)), 'Excel'
    if raw_bytes[:2] in (b'\xff\xfe', b'\xfe\xff'):
        for sep in ('\t', ','):
            try:
                df = pd.read_csv(io.BytesIO(raw_bytes), encoding='utf-16', sep=sep)
                if len(df.columns) > 5:
                    return df, 'UTF-16'
            except Exception:
                continue
    for enc in ('utf-8-sig', 'utf-8', 'cp932'):
        for sep in (',', '\t'):
            try:
                df = pd.read_csv(io.BytesIO(raw_bytes), encoding=enc, sep=sep)
                if len(df.columns) > 5:
                    return df, enc.upper()
            except Exception:
                continue
    raise ValueError('ファイル形式を判定できませんでした')

def _normalize(df):
    df = df.copy()
    n = len(df)
    for c in FFILL_COLS:
        if c in df.columns and 0 < df[c].notna().sum() < n:
            df[c] = df[c].ffill()

    if 'メジャー ネーム' in df.columns and 'メジャー バリュー' in df.columns:
        layout = '縦持ち'
        out = df[df['メジャー ネーム'] == '合計費用'].copy()
        out['メジャー バリュー'] = out['メジャー バリュー'].apply(parse_yen)
    elif '合計費用' in df.columns:
        layout = '横持ち'
        out = df.copy()
        out['メジャー ネーム'] = '合計費用'
        out['メジャー バリュー'] = out['合計費用'].apply(parse_yen)
    else:
        raise ValueError('合計費用の列が見つかりません')

    for c in SCHEMA:
        if c not in out.columns:
            out[c] = pd.NA
    out = out[SCHEMA]

    def fix_month(v):
        if pd.isna(v):
            return v
        s = str(v).strip()
        m = re.match(r'^(\d{4})[-/](\d{1,2})', s)
        return f'{m.group(1)}-{int(m.group(2)):02d}-01' if m else s
    out['対象年月'] = out['対象年月'].apply(fix_month)
    return out, layout

def load_campaign_data(raw_bytes, filename=''):
    df, enc = _read_any(raw_bytes, filename)
    out, layout = _normalize(df)
    return out, f'{enc} / {layout}'


# ==================== 請求データ 万能ローダー ====================
# 未加工の billing_statement_YYYY-MM[_dr|_jo].csv（英語列）と
# 旧来の Indeed_YYYY年M月.xlsx（日本語列）の両方に対応する。

MEDIA_MAP = {
    'aw': 'IndeedPLUS（Airワーク）',
    'dr': 'Indeed（ドラディス）',
    'jo': 'Indeed（ジョブオプ）',
}

# 英語列 → 日本語列
BILL_COLMAP = {
    'Employer ID': 'Employer ID',
    'Client name': 'Client name',
    'Document number': 'Document number',
    'Invoice Date': 'Invoice Date',
    'Txn currency': 'Txn currency',
    'Service amount': '費消額',
    'Pre-Invoice adjustment amount': '請求前調整額',
    'Promotion amount': 'プロモーション分費消額',
    'Discount rate': '手数料率',
    'Discount amount': '手数料(=粗利)',
    'Tax amount': '消費税',
    'Invoice amount': 'Indeedからの請求額',
}

BILL_MONTH_RE = re.compile(r'(20\d{2})[-_年]?\s*(\d{1,2})\s*月?')
SEG_RE = re.compile(r'(?:^|[_\-])(aw|dr|jo)(?:[_\-]|\.|$)', re.IGNORECASE)

def detect_billing_file(filename):
    """ファイル名から (対象年月キー, 媒体コード) を推定。請求ファイルでなければ None"""
    n = filename
    low = n.lower()
    if not (low.endswith(('.csv', '.xlsx', '.xls'))):
        return None
    is_raw = 'billing_statement' in low
    is_xlsx = low.startswith('indeed_') and low.endswith(('.xlsx', '.xls'))
    if not (is_raw or is_xlsx):
        return None
    m = BILL_MONTH_RE.search(n)
    if not m:
        return None
    key = f"{m.group(1)}-{int(m.group(2)):02d}-01"
    seg_m = SEG_RE.search(n)
    if seg_m:
        seg = seg_m.group(1).lower()
    else:
        # 既知の接尾辞が無い場合：月の直後に余分なトークンが付いていれば未知区分とみなす
        tail = n[m.end():]
        tail = re.sub(r'\.(csv|xlsx|xls)$', '', tail, flags=re.IGNORECASE)
        tail = re.sub(r'[_\-]?csv$', '', tail, flags=re.IGNORECASE)
        unknown = re.search(r'[_\-]([A-Za-z]{2,})$', tail)
        seg = unknown.group(1).lower() if unknown else 'aw'
    return key, seg

def load_billing_data(raw_bytes, filename, seg):
    """請求ファイル1本 → 日本語列に正規化したDataFrame"""
    low = filename.lower()
    if low.endswith(('.xlsx', '.xls')):
        df = pd.read_excel(io.BytesIO(raw_bytes))
    else:
        df = None
        for enc in ('utf-8-sig', 'utf-8', 'cp932'):
            try:
                df = pd.read_csv(io.BytesIO(raw_bytes), encoding=enc)
                if len(df.columns) > 3:
                    break
            except Exception:
                continue
        if df is None:
            raise ValueError('請求CSVを読み込めませんでした')

    df = df.rename(columns={k: v for k, v in BILL_COLMAP.items() if k in df.columns})
    df.columns = [str(c).replace('\n', '') for c in df.columns]

    if 'Employer ID' not in df.columns or '費消額' not in df.columns:
        raise ValueError('Employer ID / 費消額(Service amount) の列が見つかりません')

    df['費消額'] = pd.to_numeric(df['費消額'], errors='coerce').fillna(0).astype(int)
    # 手数料は未加工版では負値で入っているため絶対値に揃える
    if '手数料(=粗利)' in df.columns:
        df['手数料(=粗利)'] = pd.to_numeric(df['手数料(=粗利)'], errors='coerce').fillna(0).abs().astype(int)
    df['媒体'] = MEDIA_MAP.get(seg, f'未分類（{seg}）')
    return df

# ==================== ログ記録 ====================
SPREADSHEET_ID = "1ge34r5lmRi6st9hFOJvzMEvdh4MRC8csHQHvhXHs5oY"

def write_log(client_name, months, diff_results):
    try:
        sheets = get_sheets_service()
        jst = pytz.timezone("Asia/Tokyo")
        now = datetime.now(jst).strftime("%Y/%m/%d %H:%M:%S")
        month_str = "・".join(months)
        diff_str = "　".join([
            f"{m}:{'差異なし' if d == 0 else f'差異¥{abs(int(d)):,}'}"
            for m, d in diff_results
        ])
        row = [[now, client_name, month_str, diff_str]]

        result = sheets.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range="アクセスログ!A1:D1"
        ).execute()
        if not result.get("values"):
            sheets.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range="アクセスログ!A1:D1",
                valueInputOption="RAW",
                body={"values": [["日時", "クライアント名", "対象月", "突合結果"]]}
            ).execute()

        sheets.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range="アクセスログ!A:D",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": row}
        ).execute()
        return None
    except Exception as e:
        return str(e)

# ==================== Excel生成 ====================
FONT_NAME = 'メイリオ'

def all_border(color='AAAAAA', style='thin'):
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def header_border():
    s = Side(style='medium', color='1F4E79')
    return Border(left=s, right=s, top=s, bottom=s)

def total_border():
    s = Side(style='medium', color='1F4E79')
    return Border(left=s, right=s, top=s, bottom=s)

def create_billing_excel(client_name, inv_df, csv_df, month_label):
    hits_inv = inv_df[inv_df['Client name'].str.contains(client_name, na=False, regex=False)].copy()
    if len(hits_inv) == 0:
        return None, f"「{client_name}」に一致するクライアントが見つかりませんでした"
    # 同一Employer IDが請求Excelに複数行ある場合の二重計上を防ぐ
    hits_inv = hits_inv.drop_duplicates(subset=['Employer ID', '媒体'], keep='first') if '媒体' in hits_inv.columns else hits_inv.drop_duplicates(subset=['Employer ID'], keep='first')
    if '媒体' not in hits_inv.columns:
        hits_inv['媒体'] = ''
    target_ids = hits_inv['Employer ID'].tolist()

    cost_df = csv_df[(csv_df['アカウントID'].isin(target_ids)) & (csv_df['メジャー ネーム'] == '合計費用')].copy()
    cost_df['合計費用_数値'] = cost_df['メジャー バリュー'].fillna(0).astype(float).astype(int)

    merged = hits_inv.merge(
        cost_df[['アカウントID','アカウント名','キャンペーン名','キャンペーン開始日','キャンペーン終了日 (指定した日付)','キャンペーンステータス','合計費用_数値']],
        left_on='Employer ID', right_on='アカウントID', how='left'
    )
    diff = hits_inv['費消額'].sum() - cost_df['合計費用_数値'].sum()
    HEADER_BG, WHITE, GRAY, TOTAL_BG = '1F4E79', 'FFFFFF', 'F0F4F8', 'FFF2CC'
    wb = Workbook()
    ws = wb.active
    ws.title = '請求明細'
    ws.merge_cells('A1:H1')
    ws['A1'] = f'{client_name}　Indeed請求明細　{month_label}'
    ws['A1'].font = Font(name=FONT_NAME, size=13, bold=True, color=WHITE)
    ws['A1'].fill = PatternFill('solid', fgColor=HEADER_BG)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws['A1'].border = header_border()
    ws.row_dimensions[1].height = 30
    headers = ['アカウント名','媒体','キャンペーン名','開始日','終了日','ステータス','キャンペーン費消額（円）','アカウント合計費消額（円）']
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col, value=h)
        cell.font = Font(name=FONT_NAME, size=10, bold=True, color=WHITE)
        cell.fill = PatternFill('solid', fgColor='2E75B6')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = all_border(color='1F4E79', style='medium')
    ws.row_dimensions[2].height = 28
    account_order = merged['Employer ID'].unique()
    fill_white = PatternFill('solid', fgColor=WHITE)
    fill_gray = PatternFill('solid', fgColor=GRAY)
    row = 3
    for i, acc_id in enumerate(account_order):
        grp = merged[merged['Employer ID'] == acc_id].reset_index(drop=True)
        acc_name = grp['Client name'].iloc[0]
        fill = fill_gray if i % 2 == 0 else fill_white
        start_row = row
        for j, r in grp.iterrows():
            camp_fee = r['合計費用_数値'] if pd.notna(r['合計費用_数値']) else 0
            values = [acc_name if j == 0 else '', r.get('媒体','') if j == 0 else '', r.get('キャンペーン名',''), r.get('キャンペーン開始日',''), r.get('キャンペーン終了日 (指定した日付)',''), r.get('キャンペーンステータス',''), int(camp_fee), '']
            for col, val in enumerate(values, 1):
                cell = ws.cell(row=row, column=col, value=val)
                cell.font = Font(name=FONT_NAME, size=9)
                cell.fill = fill
                cell.border = all_border(color='AAAAAA', style='thin')
                cell.alignment = Alignment(vertical='center', wrap_text=True)
                if col in (7, 8):
                    cell.number_format = '#,##0'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                if col in (4, 5):
                    cell.number_format = 'YYYY/MM/DD'
                    cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[row].height = 22
            row += 1
        end_row = row - 1
        if start_row < end_row:
            ws.merge_cells(f'A{start_row}:A{end_row}')
            ws.merge_cells(f'B{start_row}:B{end_row}')
            ws.merge_cells(f'H{start_row}:H{end_row}')
        ws[f'A{start_row}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws[f'A{start_row}'].font = Font(name=FONT_NAME, size=9, bold=True)
        ws[f'B{start_row}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        tc = ws.cell(row=start_row, column=8)
        tc.value = f'=SUM(G{start_row}:G{end_row})'
        tc.number_format = '#,##0'
        tc.alignment = Alignment(horizontal='right', vertical='center')
        tc.font = Font(name=FONT_NAME, size=9, bold=True)
        for r2 in range(start_row, end_row + 1):
            for c2 in [1, 2, 8]:
                ws.cell(row=r2, column=c2).border = all_border(color='AAAAAA', style='thin')
    total_row = row
    ws.merge_cells(f'A{total_row}:F{total_row}')
    ws[f'A{total_row}'] = '合　計'
    ws[f'A{total_row}'].font = Font(name=FONT_NAME, size=10, bold=True)
    ws[f'A{total_row}'].fill = PatternFill('solid', fgColor=TOTAL_BG)
    ws[f'A{total_row}'].alignment = Alignment(horizontal='center', vertical='center')
    ws[f'A{total_row}'].border = total_border()
    for col in range(2, 7):
        cell = ws.cell(row=total_row, column=col)
        cell.fill = PatternFill('solid', fgColor=TOTAL_BG)
        cell.border = total_border()
    for col in (7, 8):
        cell = ws.cell(row=total_row, column=col)
        cell.value = f'=SUM(G3:G{total_row-1})'
        cell.font = Font(name=FONT_NAME, size=10, bold=True)
        cell.fill = PatternFill('solid', fgColor=TOTAL_BG)
        cell.number_format = '#,##0'
        cell.alignment = Alignment(horizontal='right', vertical='center')
        cell.border = total_border()
    ws.row_dimensions[total_row].height = 24
    ws.column_dimensions['A'].width = 32
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 44
    ws.column_dimensions['D'].width = 13
    ws.column_dimensions['E'].width = 13
    ws.column_dimensions['F'].width = 12
    ws.column_dimensions['G'].width = 22
    ws.column_dimensions['H'].width = 22
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.freeze_panes = 'A3'
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf, diff

# ==================== フォルダID ====================
try:
    folder_id_inv = st.secrets["FOLDER_IDS"]["FOLDER_ID_INV"]
except Exception:
    folder_id_inv = ""
try:
    folder_id_csv = st.secrets["FOLDER_IDS"]["FOLDER_ID_CSV"]
except Exception:
    folder_id_csv = ""

# ==================== Driveデータの取得（キャッシュ付き） ====================
MONTH_RE = re.compile(r'(20\d{2})\s*年\s*(\d{1,2})\s*月')

@st.cache_data(ttl=600, show_spinner=False)
def fetch_all_data():
    """Driveから全データを取得し、(キャンペーンdf, 請求df辞書, 診断) を返す"""
    service = get_drive_service()
    diag = []

    # ---------- キャンペーンデータ ----------
    all_df = None
    for f in list_files_in_folder(service, folder_id_csv):
        name = f['name']
        if not name.lower().endswith(('.csv', '.xlsx', '.xls')):
            continue
        if detect_billing_file(name):        # 請求ファイルはここでは扱わない
            continue
        try:
            raw = download_file(service, f['id']).read()
            df, fmt = load_campaign_data(raw, name)
            all_df = df if all_df is None else pd.concat([all_df, df], ignore_index=True)
            months = ', '.join(sorted(df['対象年月'].dropna().unique()))
            diag.append({'種別': 'キャンペーン', 'ファイル': name, '判定': fmt,
                         '行数': len(df), '対象月': months, '状態': '✅ OK'})
        except Exception as e:
            diag.append({'種別': 'キャンペーン', 'ファイル': name, '判定': '-',
                         '行数': 0, '対象月': '-', '状態': f'⚠ {str(e)[:50]}'})

    if all_df is not None:
        # 同一キャンペーンが複数ファイルに現れた場合は後勝ちで排除
        all_df = all_df.drop_duplicates(
            subset=['対象年月', 'アカウントID', 'キャンペーンID'], keep='last')

    # ---------- 請求データ ----------
    raw_months, xlsx_months = {}, {}
    for f in list_files_in_folder(service, folder_id_inv):
        name = f['name']
        det = detect_billing_file(name)
        if not det:
            continue
        key, seg = det
        is_legacy = name.lower().startswith('indeed_')
        try:
            raw = download_file(service, f['id']).read()
            df = load_billing_data(raw, name, seg)
            bucket = xlsx_months if is_legacy else raw_months
            bucket.setdefault(key, []).append(df)
            note = '' if seg in MEDIA_MAP else f'　⚠ 未知の区分「{seg}」'
            diag.append({'種別': '請求', 'ファイル': name,
                         '判定': ('旧xlsx' if is_legacy else '未加工CSV') + ' / ' + MEDIA_MAP.get(seg, seg),
                         '行数': len(df), '対象月': key,
                         '状態': ('✅ OK' + note)})
        except Exception as e:
            diag.append({'種別': '請求', 'ファイル': name, '判定': '-',
                         '行数': 0, '対象月': key, '状態': f'⚠ {str(e)[:50]}'})

    # 未加工CSVがある月は、旧xlsxを使わない（二重計上の防止）
    inv_data = {}
    for key, parts in raw_months.items():
        inv_data[key] = pd.concat(parts, ignore_index=True)
    for key, parts in xlsx_months.items():
        if key in inv_data:
            for d in diag:
                if d['種別'] == '請求' and d['対象月'] == key and '旧xlsx' in str(d['判定']):
                    d['状態'] = 'ℹ 未加工CSVを優先（この旧ファイルは未使用）'
            continue
        inv_data[key] = pd.concat(parts, ignore_index=True)

    # 同一Employer IDの重複を除去（媒体が同じなら1件に）
    for key, df in inv_data.items():
        inv_data[key] = df.drop_duplicates(subset=['Employer ID', '媒体'], keep='first')

    return all_df, inv_data, diag


# ==================== UI ====================
st.markdown("""
<div class="title-bar">
    <div class="main-title">📊 Indeed請求明細ジェネレーター</div>
    <div class="sub-title">Google DriveのデータからクライアントごとのIndeed請求明細Excelを自動生成します</div>
</div>
""", unsafe_allow_html=True)

if not folder_id_inv or not folder_id_csv:
    st.error("フォルダIDが設定されていません。Secretsを確認してください")
    st.stop()

col_r1, col_r2 = st.columns([4, 1])
with col_r2:
    if st.button("🔄 データを再読込", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

try:
    with st.spinner("Google Driveからデータを取得中..."):
        all_csv_df, inv_data, diag = fetch_all_data()
except Exception as e:
    st.error(f"Google Driveへの接続に失敗しました：{e}")
    st.stop()

csv_months = set(all_csv_df['対象年月'].dropna().unique()) if all_csv_df is not None else set()
ready = sorted(set(inv_data.keys()) & csv_months, reverse=True)

def month_label(key):
    y, m, _ = key.split('-')
    return f"{y}年{int(m)}月"

with st.expander("🔍 読み込んだデータの状況", expanded=(len(ready) == 0)):
    if diag:
        st.dataframe(pd.DataFrame(diag), use_container_width=True, hide_index=True)
    only_csv = sorted(csv_months - set(inv_data.keys()))
    only_inv = sorted(set(inv_data.keys()) - csv_months)
    if only_csv:
        st.warning("キャンペーンデータのみ（請求データが未アップロード）：" + ", ".join(month_label(k) for k in only_csv))
    if only_inv:
        st.warning("請求データのみ（キャンペーンデータに該当月なし）：" + ", ".join(month_label(k) for k in only_inv))
    unknown = [d for d in diag if '未知の区分' in str(d.get('状態', ''))]
    if unknown:
        st.error("未知の媒体区分が見つかりました。正しい媒体名を設定するため、接尾辞をご確認ください。")

if not ready:
    st.error("生成可能な月がありません。キャンペーンデータと請求データの両方をDriveにアップロードしてください。")
    st.stop()

col1, col2 = st.columns([1, 1], gap="large")

with col1:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("① 対象月を選択")
    selected_keys = st.multiselect(
        "対象月（複数選択可）",
        ready,
        default=[ready[0]],
        format_func=month_label
    )
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("② クライアント名を入力")
    client_name = st.text_input("クライアント名（部分一致）", placeholder="例：JSS、ノムラメディアス、ORES など")
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("③ 請求明細を生成")

    if st.button("📥 請求明細Excelを生成", use_container_width=True):
        if not client_name:
            st.error("クライアント名を入力してください")
        elif not selected_keys:
            st.error("対象月を1つ以上選択してください")
        else:
            try:
                results, diff_results = [], []

                for key in selected_keys:
                    label = month_label(key)
                    inv_df = inv_data[key]
                    csv_month = all_csv_df[all_csv_df['対象年月'] == key]
                    result_buf, diff = create_billing_excel(client_name, inv_df, csv_month, label)
                    if result_buf is None:
                        st.error(diff)
                    else:
                        results.append((label, result_buf, diff))
                        diff_results.append((label, diff))

                if results:
                    log_error = write_log(client_name, [month_label(k) for k in selected_keys], diff_results)
                    if log_error:
                        st.warning(f"⚠️ ログ記録エラー：{log_error}")

                    st.success(f"{len(results)}件のExcelを生成しました")
                    for label, result_buf, diff in results:
                        st.download_button(
                            label=f"⬇️ {label}　Excelをダウンロード",
                            data=result_buf,
                            file_name=f"{client_name}_Indeed請求明細_{label}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                            key=f"dl_{label}"
                        )
                        if diff == 0:
                            st.markdown(f'<div class="result-box">✅ {label}　突合完了・差異ゼロ</div>', unsafe_allow_html=True)
                        else:
                            st.markdown(
                                f'<div class="error-box">⚠️ {label}　差異あり：¥{abs(int(diff)):,}<br>'
                                '<span style="font-size:0.85rem">Indeedのデータが月末最終日まで反映されていない可能性があります。'
                                'Indeed管理画面で対象月のデータを再エクスポートしてください。</span></div>',
                                unsafe_allow_html=True)

            except Exception as e:
                st.error(f"エラーが発生しました：{e}")

    st.markdown('</div>', unsafe_allow_html=True)

st.markdown("---")
st.caption("© Indeed請求明細ジェネレーター")
