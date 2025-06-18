# app.py
from flask import Flask, request, render_template, flash, redirect
import pandas as pd
import re
import io
import base64
from fugashi import Tagger

app = Flask(__name__)
app.secret_key = 'your-secret-key'

# ── 定数 ───────────────────────────────────────────
CATALOG_PATH     = 'textbooks_list.xlsx'
GUIDELINES_PATH  = 'かな及び漢字等の書き表し方.xlsx'
# ───────────────────────────────────────────────────

# MeCab 互換形態素解析器
tagger = Tagger()

# ── 起動時処理：カタログ読み込み & 正規化 ────────────
df_catalog = pd.read_excel(CATALOG_PATH, dtype=str)
df_catalog = df_catalog.apply(
    lambda col: col.str.strip() if col.dtype == 'object' else col,
    axis=0
)

# ── 起動時処理：ガイドライン読み込み & パターン生成 ────
df_guidelines = pd.read_excel(
    GUIDELINES_PATH,
    sheet_name='かな及び漢字等の書き表し方',
    dtype=str
).apply(
    lambda col: col.str.strip() if col.dtype == 'object' else col,
    axis=0
)

patterns = []
last_idx = df_guidelines.index.max()
for idx, row in df_guidelines.iterrows():
    correct = row['使用する表現']
    note    = row['備考']
    if pd.isna(note):
        continue
    if idx == last_idx:
        incorrects = re.findall(r'「(.+?)」', correct)
        patterns.append({'incorrects': incorrects, 'correct': correct, 'quote_required': True})
    else:
        incs = note.lstrip('×').split('、')
        incs = [inc.strip() for inc in incs if inc.strip()]
        patterns.append({'incorrects': incs, 'correct': correct, 'quote_required': False})

# ── カタログ照合用マッピング ────────────────────────────────
CHECK_TO_CATALOG = {
    '教科':         '教科名',
    '種目':         '種目',
    '発行者の略称': '発行者略称',
    '教科書の番号': '教科書番号',
    '書名':         '書籍名',
}


@app.route('/', methods=['GET', 'POST'])
def index():
    download_link = None

    if request.method == 'POST':
        # アップロードファイル取得
        f = request.files.get('file')
        if not f:
            flash('ファイルをアップロードしてください。')
            return redirect(request.url)

        # シート名自動検出
        xls = pd.ExcelFile(f)
        sheet = next((s for s in xls.sheet_names if '別紙様式２' in s), None)
        if not sheet:
            flash(f"シート '別紙様式２' が見つかりません。利用可能なシート: {xls.sheet_names}")
            return redirect(request.url)

        # チェック用データ読み込み & ヘッダー整形
        df_check = pd.read_excel(f, sheet_name=sheet, dtype=str)
        df_check.columns = df_check.columns.map(lambda x: re.sub(r'\s+', '', x) if isinstance(x, str) else x)

        # 必須列チェック
        missing = [col for col in CHECK_TO_CATALOG if col not in df_check.columns]
        if missing:
            flash(f"列 '{missing[0]}' が見つかりません。利用可能な列: {list(df_check.columns)}")
            return redirect(request.url)

        # カタログ列チェック
        missing_cat = [CHECK_TO_CATALOG[col] for col in CHECK_TO_CATALOG
                       if CHECK_TO_CATALOG[col] not in df_catalog.columns]
        if missing_cat:
            flash(f"カタログの列 '{missing_cat[0]}' が見つかりません。")
            return redirect(request.url)

        # 前処理：空白除去
        for col in CHECK_TO_CATALOG:
            df_check[col] = df_check[col].fillna('').astype(str).str.strip()

        # ① 目録照合チェック
        for col, cat_col in CHECK_TO_CATALOG.items():
            catalog_set = set(df_catalog[cat_col].astype(str).values)
            df_check[f'{col}_check'] = df_check[col].apply(
                lambda v: 'OK' if '\n' in v or v in catalog_set else '要確認'
            )
        def combined_ok(row):
            if any('\n' in str(row[col]) for col in CHECK_TO_CATALOG):
                return 'OK'
            cond = pd.Series(True, index=df_catalog.index)
            for col, cat_col in CHECK_TO_CATALOG.items():
                cond &= (df_catalog[cat_col] == row[col])
            return 'OK' if cond.any() else '要確認'
        df_check['総合チェック'] = df_check.apply(combined_ok, axis=1)

        # ② 選定理由不正表記チェック
        reason_col = next((c for c in df_check.columns if '選定理由' in c), None)
        violations = []
        for idx, row in df_check.iterrows():
            text = str(row.get(reason_col, ''))
            # ガイドライン違反
            for pat in patterns:
                if pat['quote_required']:
                    for inc in pat['incorrects']:
                        if inc in text and f'「{inc}」' not in text:
                            violations.append({'行番号': idx+1, '教科': row['教科'], '種目': row['種目'],
                                                '違反候補': f'科目名「{inc}」は引用符で囲まれていません'})
                else:
                    for inc in pat['incorrects']:
                        if inc in text:
                            violations.append({'行番号': idx+1, '教科': row['教科'], '種目': row['種目'],
                                                '違反候補': f'「{inc}」は不正です。正しくは「{pat["correct"]}」'})
            # 他者比較 or １者のみ
            if '他者と比較して' not in text and '１者のみの発行' not in text:
                violations.append({'行番号': idx+1, '教科': row['教科'], '種目': row['種目'],
                                   '違反候補': '「他者と比較して」または「１者のみの発行」の記載が必要です'})
            # 自校生徒言及
            if '本校生徒' not in text and '自校の生徒' not in text:
                violations.append({'行番号': idx+1, '教科': row['教科'], '種目': row['種目'],
                                   '違反候補': '自校の生徒の実態を踏まえた文言を含めてください'})
        df_violations = pd.DataFrame(violations)

        # ③ 誤字脱字チェック（未知語検出）
        typo_list = []
        for idx, row in df_check.iterrows():
            text = str(row.get(reason_col, ''))
            for token in tagger(text):
                if getattr(token, 'is_unknown', False):
                    typo_list.append({'行番号': idx+1, '教科': row['教科'], '種目': row['種目'],
                                      '候補': token.surface})
        df_typos = pd.DataFrame(typo_list)

        # ── Excel ファイルをメモリ上で作成 ───────────────────
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_check.to_excel(writer, sheet_name='目録照合チェック', index=False)
            df_violations.to_excel(writer, sheet_name='不正表記チェック', index=False)
            df_typos.to_excel(writer, sheet_name='誤字脱字チェック', index=False)
        data = output.getvalue()
        b64 = base64.b64encode(data).decode()
        download_link = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + b64

        # テンプレートへ結果を渡す
        return render_template(
            'index.html',
            table_catalog = df_check.to_html(classes='table table-sm table-bordered', index=False, na_rep=''),
            table_reason  = df_violations.to_html(classes='table table-sm table-bordered', index=False, na_rep=''),
            table_typos   = df_typos.to_html(classes='table table-sm table-bordered', index=False, na_rep=''),
            download_link = download_link
        )

    return render_template('index.html')


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
