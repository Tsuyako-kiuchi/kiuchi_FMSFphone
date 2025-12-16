
# make_index.py
# data.xlsx から PINなし・検索/クリア・初期はカテゴリ展開・URL表示・カード/テーブル切替付きの index.html を生成

import os
import re
import pandas as pd
from html import escape

CATEGORY_ORDER = [
    "発注者", "コンサル", "設計監理", "別途業者", "施工図", "躯体工事",
    "仕上げ工事", "その他業者", "職員"
]

COL_MAP = {
    'カテゴリ': 'カテゴリ', 'カテゴリー': 'カテゴリ', 'category': 'カテゴリ',
    '業種': 'カテゴリ', '業種ギョウシュ': 'カテゴリ',
    '氏名': '氏名', '名前': '氏名', 'name': '氏名',
    '会社名': '会社名', '会社': '会社名', '所属': '会社名', 'company': '会社名',
    '役職': '役職', '肩書': '役職', 'position': '役職',
    '住所': '住所', 'address': '住所',
    '電話': '電話', 'TEL': '電話', 'tel': '電話', '電話番号': '電話',
    'FAX': 'FAX', 'fax': 'FAX', 'FAX番号': 'FAX',
    '携帯': '携帯', 'mobile': '携帯', '携帯電話': '携帯', '携帯番号': '携帯',
    'メール': 'メール', 'mail': 'メール', 'email': 'メール', 'E-mail': 'メール', 'アドレス': 'メール'
}

EXPECTED_COLS = ['カテゴリ','氏名','会社名','役職','住所','電話','FAX','携帯','メール']

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """列名の表記ゆれを吸収して標準列に整形"""
    target_sources = {t: [] for t in EXPECTED_COLS}
    for c in df.columns:
        if c in COL_MAP:
            target_sources[COL_MAP[c]].append(c); continue
        lc = str(c).lower()
        for k, v in COL_MAP.items():
            if k.lower() == lc:
                target_sources[v].append(c)
                break
    out = pd.DataFrame()
    for t in EXPECTED_COLS:
        srcs = target_sources.get(t, [])
        if not srcs:
            out[t] = ''
        else:
            s = df[srcs[0]].astype(str)
            s = s.where(~s.str.lower().eq('nan'), '').fillna('')
            for src in srcs[1:]:
                s2 = df[src].astype(str)
                s2 = s2.where(~s2.str.lower().eq('nan'), '').fillna('')
                mask = (s.str.strip() == '') & (s2.str.strip() != '')
                s = s.mask(mask, s2)
            out[t] = s
    return out[EXPECTED_COLS]

def normalize_category(value) -> str:
    s = '' if pd.isna(value) else str(value).strip()
    if not s: return ''
    if '発注者' in s: return '発注者'
    if 'コンサル' in s: return 'コンサル'
    if '設計監理' in s: return '設計監理'
    if '施工図' in s: return '施工図'
    if ('躯体' in s) or ('クタイ' in s): return '躯体工事'
    if ('仕上げ' in s) or ('仕上' in s) or ('シア' in s): return '仕上げ工事'
    if ('その他' in s) or ('他' in s): return 'その他業者'
    if '職員' in s: return '職員'
    if '別途' in s: return '別途業者'
    return s

def tel_link(num: str) -> str:
    num = (str(num) if not pd.isna(num) else '').strip()
    if not num: return ''
    cleaned = ''.join(re.findall(r"[0-9+]+", num))
    if not cleaned: cleaned = num
    return f"tel:{cleaned}"

def mail_link(addr: str) -> str:
    addr = (str(addr) if not pd.isna(addr) else '').strip()
    if not addr: return ''
    return f"mailto:{addr}"

def build_entry_search_text(row: dict) -> str:
    parts = [row.get(k, '') for k in EXPECTED_COLS]
    return ' '.join([str(p) for p in parts if p]).lower()

def build_card_html(row: dict) -> str:
    name = escape(str(row.get('氏名','') or ''))
    company = escape(str(row.get('会社名','') or ''))
    role = escape(str(row.get('役職','') or ''))
    addr = escape(str(row.get('住所','') or ''))
    tel = str(row.get('電話','') or '')
    fax = str(row.get('FAX','') or '')
    mobile = str(row.get('携帯','') or '')
    mail = str(row.get('メール','') or '')

    tel_href = tel_link(tel)
    mobile_href = tel_link(mobile)
    mail_href = mail_link(mail)

    contact_parts = []
    if tel.strip():
        contact_parts.append(f'<span class="label">電話:</span> {escape(tel_href)}{escape(tel)}</a>')
    if fax.strip():
        contact_parts.append(f'<span class="label">FAX:</span> {escape(fax)}')
    if mobile.strip():
        contact_parts.append(f'<span class="label">携帯:</span> {escape(mobile_href)}{escape(mobile)}</a>')
    contact_html = ' / '.join(contact_parts) if contact_parts else ''

    email_html = f'<span class="label">メール:</span> {escape(mail_href)}{escape(mail)}</a>' if mail.strip() else ''
    company_role = ' / '.join([p for p in [company, role] if p])

    return f'''<div class="card entry" data-search="{escape(build_entry_search_text(row))}">
  <div class="card-header">
    <div class="name">{name}</div>
    <div class="company-role">{company_role}</div>
  </div>
  <div class="card-body">
    {('<div class="contact">' + contact_html + '</div>') if contact_html else ''}
    {('<div class="email">' + email_html + '</div>') if email_html else ''}
    {('<div class="addr"><span class="label">住所:</span> ' + addr + '</div>') if addr else ''}
  </div>
</div>'''

def build_table_row_html(row: dict) -> str:
    name = escape(str(row.get('氏名','') or ''))
    company = escape(str(row.get('会社名','') or ''))
    role = escape(str(row.get('役職','') or ''))
    addr = escape(str(row.get('住所','') or ''))
    tel = str(row.get('電話','') or '')
    fax = str(row.get('FAX','') or '')
    mobile = str(row.get('携帯','') or '')
    mail = str(row.get('メール','') or '')

    tel_href = tel_link(tel)
    mobile_href = tel_link(mobile)
    mail_href = mail_link(mail)

    tel_cell = f'{escape(tel_href)}{escape(tel)}</a>' if tel.strip() else ''
    fax_cell = escape(fax) if fax.strip() else ''
    mobile_cell = f'{escape(mobile_href)}{escape(mobile)}</a>' if mobile.strip() else ''
    mail_cell = f'{escape(mail_href)}{escape(mail)}</a>' if mail.strip() else ''

    return f'''<tr class="entry" data-search="{escape(build_entry_search_text(row))}">
  <td class="td-name">{name}</td>
  <td class="td-company">{company}</td>
  <td class="td-role">{role}</td>
  <td class="td-addr">{addr}</td>
  <td class="td-tel">{tel_cell}</td>
  <td class="td-fax">{fax_cell}</td>
  <td class="td-mobile">{mobile_cell}</td>
  <td class="td-mail">{mail_cell}</td>
</tr>'''

def build_category_section(cat: str, df_cat: pd.DataFrame) -> str:
    cards = "\n".join([build_card_html(row) for _, row in df_cat.iterrows()])
    rows = "\n".join([build_table_row_html(row) for _, row in df_cat.iterrows()])
    return f'''<details class="category" data-category="{escape(cat)}">
  <summary>{escape(cat)}</summary>
  <div class="entries card-view">
    {cards if cards.strip() else '<div class="no-data">該当データがありません</div>'}
  </div>
  <div class="entries table-view">
    <table class="table">
      <thead>
        <tr>
          <th>氏名</th><th>会社名</th><th>役職</th><th>住所</th><th>電話</th><th>FAX</th><th>携帯</th><th>メール</th>
        </tr>
      </thead>
      <tbody>
        {rows if rows.strip() else '<tr class="no-data"><td colspan="8">該当データがありません</td></tr>'}
      </tbody>
    </table>
  </div>
</details>'''

def build_html(df: pd.DataFrame) -> str:
    df['カテゴリ'] = df['カテゴリ'].apply(normalize_category)
    df['カテゴリ'] = df['カテゴリ'].fillna('').astype(str)

    present_cats = [c for c in CATEGORY_ORDER if c in set(df['カテゴリ'])]
    other_cats = sorted([c for c in set(df['カテゴリ']) if c not in CATEGORY_ORDER and c])
    all_cats = present_cats + other_cats

    sections = []
    for cat in all_cats:
        df_cat = df[df['カテゴリ'] == cat].copy()
        df_cat['会社名'] = df_cat['会社名'].fillna('')
        df_cat['氏名'] = df_cat['氏名'].fillna('')
        df_cat = df_cat.sort_values(by=['会社名','氏名'])
        sections.append(build_category_section(cat, df_cat))

    sections_html = "\n\n".join(sections)

    html_prefix = '''<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>電話帳（PINなし版）</title>
  <style>
    :root {
      --bg: #f9fafb; --fg: #111827; --muted: #6b7280; --primary: #2563eb; --border: #e5e7eb;
    }
    html, body { margin: 0; padding: 0; background: var(--bg); color: var(--fg);
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial,
                   "Noto Sans JP", "Hiragino Kaku Gothic ProN", "Yu Gothic", sans-serif; }
    header.searchbar { position: sticky; top: 0; z-index: 1000; background: white; border-bottom: 1px solid var(--border);
      padding: 12px 16px; box-shadow: 0 2px 8px rgba(0,0,0,0.04); }
    header .title { font-size: 20px; font-weight: 600; margin: 0 0 6px 0; }
    header .page-url { font-size: 12px; color: var(--muted); word-break: break-all; margin-bottom: 8px; }
    .controls { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }
    .controls input[type="search"] { flex: 1 1 280px; padding: 8px 10px; border: 1px solid var(--border);
      border-radius: 6px; font-size: 14px; }
    .btn { padding: 8px 10px; border: 1px solid var(--border); background: #fff; border-radius: 6px; font-size: 13px; cursor: pointer; }
    .btn.active { background: var(--primary); color:#fff; border-color:var(--primary); }

    main { padding: 16px; }
    details.category { border: 1px solid var(--border); border-radius: 8px; background: #fff; margin-bottom: 12px; }
    details.category > summary { cursor: pointer; list-style: none; padding: 12px 16px; font-weight: 600; }
    details.category[open] > summary { border-bottom: 1px solid var(--border); }
    .entries { padding: 12px 16px; }

    /* Card view */
    .card-view { display: block; }
    .card { border: 1px solid var(--border); border-radius: 8px; padding: 10px 12px; margin: 8px 0; background: #fafafa; }
    .card-header { display: flex; justify-content: space-between; gap: 8px; flex-wrap: wrap; }
    .card .name { font-weight: 700; }
    .card .company-role { color: var(--muted); }
    .card .label { color: var(--muted); margin-right: 4px; }
    .card .contact, .card .email, .card .addr { margin-top: 6px; }

    /* Table view */
    .table-view { display: none; }
    .table { width: 100%; border-collapse: collapse; }
    .table th, .table td { border-bottom: 1px solid var(--border); padding: 8px; text-align: left; font-size: 14px; }
    .table th { background: #f3f4f6; }
    .no-data { color: var(--muted); font-size: 13px; padding: 8px 0; }

    /* View toggling */
    #content.view-card .card-view { display: block; }
    #content.view-card .table-view { display: none; }
    #content.view-table .card-view { display: none; }
    #content.view-table .table-view { display: block; }
  </style>
</head>
<body>
  <header class="searchbar">
    <h1 class="title">電話帳（PINなし）</h1>
    <div class="page-url"><span id="page-url"></span></div>
    <div class="controls">
      <input id="search" type="search" placeholder="全文検索（氏名・会社・役職・住所・電話・FAX・携帯・メール）" aria-label="検索">
      <button id="clear-btn" class="btn" type="button">クリア</button>
      <div class="view-toggle" role="group" aria-label="表示切り替え">
        <button id="btn-card" class="btn active" type="button">カード表示</button>
        <button id="btn-table" class="btn" type="button">テーブル表示</button>
      </div>
    </div>
  </header>
  <main id="content" class="view-card">
'''
    html_suffix = '''
  </main>

  <script>
    (function(){
      var el = document.getElementById('page-url');
      if (el) el.textContent = window.location.href;
    })();

    const content = document.getElementById('content');
    const searchInput = document.getElementById('search');
    const clearBtn = document.getElementById('clear-btn');
    const btnCard = document.getElementById('btn-card');
    const btnTable = document.getElementById('btn-table');

    function filter(q) {
      q = (q || '').toLowerCase();
      const entries = content.querySelectorAll('.entry');
      const catHasMatch = new Map();

      entries.forEach(function(el) {
        const text = (el.getAttribute('data-search') || '').toLowerCase();
        const match = !q || text.indexOf(q) !== -1;
        el.style.display = match ? '' : 'none';
        const cat = el.closest('details.category');
        if (match && cat) catHasMatch.set(cat, true);
      });

      const cats = content.querySelectorAll('details.category');
      cats.forEach(function(d) {
        if (q) { d.open = !!catHasMatch.get(d); }
        else   { d.open = true; }  // ← 初期は展開（開く）
        d.querySelectorAll('.entries').forEach(function(wrapper) {
          if (q) {
            const visible = wrapper.querySelectorAll('.entry:not([style*="display: none"])').length;
            wrapper.style.display = visible ? '' : 'none';
          } else {
            wrapper.style.display = '';  // ← 初期は常に表示
          }
        });
      });
    }

    searchInput.addEventListener('input', function() { filter(searchInput.value); });
    clearBtn.addEventListener('click', function() { searchInput.value = ''; filter(''); });

    btnCard.addEventListener('click', function() {
      content.classList.add('view-card'); content.classList.remove('view-table');
      btnCard.classList.add('active'); btnTable.classList.remove('active');
    });
    btnTable.addEventListener('click', function() {
      content.classList.add('view-table'); content.classList.remove('view-card');
      btnTable.classList.add('active'); btnCard.classList.remove('active');
    });

    // 初期化（検索なし → カテゴリは展開）
    filter('');
  </script>
</body>
</html>'''
    return html_prefix + sections_html + html_suffix

def main():
    # .nojekyll（空ファイル）
    with open('.nojekyll', 'w', encoding='utf-8') as f:
        f.write('')

    # 必要に応じて sheet_name を指定（最初のシートなら省略可）
    df = pd.read_excel('data.xlsx', engine='openpyxl')  # 例: sheet_name='電話帳'

    df = normalize_columns(df)
    df = df.dropna(how='all').fillna('')

    # まったく空の行を除外
    mask_valid = df[['氏名','会社名','メール','電話','携帯','住所']].astype(str)\
                  .apply(lambda r: any(v.strip() for v in r), axis=1)
    df = df[mask_valid]

    html = build_html(df)
    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(html)
    print('index.html と .nojekyll を生成しました')

ifif __name__ == '__main__':
