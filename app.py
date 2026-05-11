"""
追善法要のしおり 生成システム
実行: streamlit run app.py
"""
import streamlit as st
from datetime import date, timedelta
import re, io, zipfile, os

st.set_page_config(
    page_title='追善法要のしおり 生成',
    page_icon='📄',
    layout='centered'
)

# ─── テンプレート読み込み ─────────────────────────────────────────────────────
TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'template.docx')

@st.cache_data
def load_template():
    with open(TEMPLATE_PATH, 'rb') as f:
        return f.read()

try:
    TEMPLATE_BYTES = load_template()
except FileNotFoundError:
    st.error('⚠ template.docx が見つかりません。app.py と同じフォルダに置いてください。')
    st.stop()

# ─── 日付計算 ─────────────────────────────────────────────────────────────────
def calc_dates(d: date):
    def ann(years):
        try:    return d.replace(year=d.year + years)
        except: return d.replace(year=d.year + years, day=28)
    ignichi = {
        '初七日': d + timedelta(6),  '二七日': d + timedelta(13),
        '三七日': d + timedelta(20), '四七日': d + timedelta(27),
        '五七日': d + timedelta(34), '六七日': d + timedelta(41),
        '尽七日': d + timedelta(48), '百カ日': d + timedelta(99),
    }
    nenki = {
        '一周忌':  ann(1),  '三回忌':  ann(2),
        '七回忌':  ann(6),  '十三回忌': ann(12),
        '十七回忌': ann(16), '廿三回忌': ann(22),
        '廿七回忌': ann(26), '卅三回忌': ann(32),
        '卅七回忌': ann(36), '五十回忌': ann(49),
    }
    return ignichi, nenki

# ─── 数値変換 ─────────────────────────────────────────────────────────────────
def reiwa(year: int) -> str:
    r = year - 2018
    fw = '０１２３４５６７８９'
    if r <= 0: return '？'
    if r < 10: return fw[r]
    return fw[r // 10] + fw[r % 10]

def zen(n: int) -> str:
    """1桁→全角（縦書きで横倒し防止）、2桁以上→半角"""
    return '０１２３４５６７８９'[n] if n < 10 else str(n)

def fw_year(y: int, upper: int, lower: int) -> tuple:
    """
    西暦年を全角に変換し、テンプレートの分割数に合わせて2分割する
    upper: 上位ランの文字数, lower: 下位ランの文字数
    例: fw_year(2027, 3, 1) → ('２０２', '７')
        fw_year(2052, 2, 2) → ('２０', '５２')
    """
    fw = '０１２３４５６７８９'
    s = ''.join(fw[int(c)] for c in str(y))
    return s[:upper], s[upper:]

# テンプレートの各年回忌における run 分割（上位桁数, 下位桁数）
# ※ テンプレート固有の Word 分割構造に対応
NENKI_SPLITS = {
    '一周忌':   (3, 1),   # 2025 → ２０２|５
    '廿三回忌': (3, 1),   # 2046 → ２０４|６
    '三回忌':   (3, 1),   # 2026 → ２０２|６
    '廿七回忌': (2, 2),   # 2050 → ２０|５０
    '七回忌':   (2, 2),   # 2030 → ２０|３０
    '卅三回忌': (2, 2),   # 2056 → ２０|５６
    '十三回忌': (3, 1),   # 2036 → ２０３|６
    '卅七回忌': (2, 2),   # 2060 → ２０|６０
    '十七回忌': (2, 2),   # 2040 → ２０|４０
    '五十回忌': (3, 1),   # 2073 → ２０７|３
}

# ─── XML操作 ──────────────────────────────────────────────────────────────────
def get_runs(xml: str):
    return [(m.start(), m.end(), m.group(1), m.group(2))
            for m in re.finditer(r'<w:t([^>]*)>(.*?)</w:t>', xml, re.DOTALL)]

def apply_simple(xml: str, replacements: list) -> str:
    """インデックス指定テキスト置換をbyte降順で一括適用"""
    runs = get_runs(xml)
    for s, e, attrs, new in sorted(
        [(runs[i][0], runs[i][1], runs[i][2], v) for i, v in replacements],
        key=lambda x: -x[0]
    ):
        xml = xml[:s] + f'<w:t{attrs}>{new}</w:t>' + xml[e:]
    return xml

def find_ruby_run(xml: str, kana: str) -> tuple:
    """kana を含む ruby の外側 <w:r...> の (start, end) を返す
    生docxでは <w:r> に属性がつくため <w:r[ >] パターンで検索"""
    pos   = xml.index(f'>{kana}<')
    ruby  = xml.rindex('<w:ruby>', 0, pos)
    # <w:r> または <w:r + 属性 を検索（<w:rPr><w:rFonts>等を除外するため [ >] を使う）
    r_positions = [m.start() for m in re.finditer(r'<w:r[ >]', xml[:ruby])]
    start = r_positions[-1]
    end   = pos + re.search(r'</w:ruby>\s*</w:r>', xml[pos:]).end()
    return start, end

def make_choshu_mei_run(mei_k: str, mei_r: str) -> str:
    """弔主の名（漢字あり→ruby、なし→plain text）"""
    if mei_k and mei_k != mei_r:
        return f'''<w:r>
      <w:rPr><w:sz w:val="36"/><w:szCs w:val="36"/>
        <w:highlight w:val="yellow"/></w:rPr>
      <w:ruby>
        <w:rubyPr><w:rubyAlign w:val="distributeSpace"/>
          <w:hps w:val="18"/><w:hpsRaise w:val="34"/>
          <w:hpsBaseText w:val="36"/><w:lid w:val="ja-JP"/></w:rubyPr>
        <w:rt><w:r><w:rPr>
          <w:rFonts w:ascii="游明朝" w:eastAsia="游明朝" w:hAnsi="游明朝"/>
          <w:sz w:val="18"/><w:szCs w:val="36"/>
          <w:highlight w:val="yellow"/></w:rPr>
          <w:t>{mei_r}</w:t></w:r></w:rt>
        <w:rubyBase><w:r><w:rPr>
          <w:sz w:val="36"/><w:szCs w:val="36"/>
          <w:highlight w:val="yellow"/></w:rPr>
          <w:t>{mei_k}</w:t></w:r></w:rubyBase>
      </w:ruby>
    </w:r>'''
    return f'''<w:r>
      <w:rPr><w:rFonts w:hint="eastAsia"/>
        <w:sz w:val="36"/><w:szCs w:val="36"/>
        <w:highlight w:val="yellow"/></w:rPr>
      <w:t>{mei_r}</w:t>
    </w:r>'''

def make_hogo_ruby(pairs: list) -> str:
    """法号 ruby XML を生成（最後のペアはハイライトなし）"""
    parts = []
    for i, (k, r) in enumerate(pairs):
        last = (i == len(pairs) - 1)
        hl = '' if last else '<w:highlight w:val="yellow"/>'
        parts.append(f'''<w:r>
      <w:rPr>
        <w:rFonts w:ascii="HG正楷書体-PRO" w:eastAsia="HG正楷書体-PRO" w:hint="eastAsia"/>
        <w:sz w:val="60"/><w:szCs w:val="60"/>{hl}
      </w:rPr>
      <w:ruby>
        <w:rubyPr><w:rubyAlign w:val="distributeSpace"/>
          <w:hps w:val="30"/><w:hpsRaise w:val="58"/>
          <w:hpsBaseText w:val="60"/><w:lid w:val="ja-JP"/></w:rubyPr>
        <w:rt><w:r><w:rPr>
          <w:rFonts w:ascii="HG正楷書体-PRO" w:eastAsia="HG正楷書体-PRO"
            w:hAnsi="游明朝" w:hint="eastAsia"/>
          <w:sz w:val="30"/><w:szCs w:val="60"/>{hl}
        </w:rPr><w:t>{r}</w:t></w:r></w:rt>
        <w:rubyBase><w:r><w:rPr>
          <w:rFonts w:ascii="HG正楷書体-PRO" w:eastAsia="HG正楷書体-PRO" w:hint="eastAsia"/>
          <w:sz w:val="60"/><w:szCs w:val="60"/>{hl}
        </w:rPr><w:t>{k}</w:t></w:r></w:rubyBase>
      </w:ruby>
    </w:r>''')
    return '\n    '.join(parts)

# ─── Word 生成 ────────────────────────────────────────────────────────────────
def generate_docx(data: dict) -> bytes:
    with zipfile.ZipFile(io.BytesIO(TEMPLATE_BYTES)) as z:
        xml = z.read('word/document.xml').decode('utf-8')

    d = date.fromisoformat(data['death_date'])
    ig, nk = calc_dates(d)
    def md(dt): return zen(dt.month), zen(dt.day)

    def yr(key):
        u, l = NENKI_SPLITS[key]
        return fw_year(nk[key].year, u, l)

    n1, n23 = nk['一周忌'], nk['廿三回忌']
    y1h,y1l   = yr('一周忌');  y23h,y23l = yr('廿三回忌')
    y3h,y3l   = yr('三回忌');  y27h,y27l = yr('廿七回忌')
    y7h,y7l   = yr('七回忌');  y33h,y33l = yr('卅三回忌')
    y13h,y13l = yr('十三回忌'); y37h,y37l = yr('卅七回忌')
    y17h,y17l = yr('十七回忌'); y50h,y50l = yr('五十回忌')

    # ── テンプレート生XML のランインデックス（template.docx 固有）──
    simple = [
        # 弔主姓
        (5,  data['choshu_sei_kana']),
        (6,  data['choshu_sei_kanji']),
        # 帰寂日
        (36, reiwa(d.year)),
        (39, zen(d.month)),
        (42, zen(d.day)),
        # 俗名・行年
        (51, data['name_sei_kana']),
        (52, data['name_sei_kanji']),
        (54, data['name_mei_kana']),
        (55, data['name_mei_kanji']),
        (59, str(data['age'])),
        # 忌日
        (65, md(ig['初七日'])[0]), (67, md(ig['初七日'])[1]),
        (75, md(ig['五七日'])[0]), (77, md(ig['五七日'])[1]),
        (82, md(ig['二七日'])[0]), (84, md(ig['二七日'])[1]),
        (92, md(ig['六七日'])[0]), (94, md(ig['六七日'])[1]),
        (99, md(ig['三七日'])[0]), (101,md(ig['三七日'])[1]),
        (109,md(ig['尽七日'])[0]), (111,md(ig['尽七日'])[1]),
        (116,md(ig['四七日'])[0]), (118,md(ig['四七日'])[1]),
        (126,md(ig['百カ日'])[0]), (128,md(ig['百カ日'])[1]),
        # 年回忌
        (134,y1h),  (135,y1l),
        (137,zen(n1.month)),  (139,zen(n1.day)),
        (144,y23h), (145,y23l),
        (147,zen(n23.month)), (149,zen(n23.day)),
        (153,y3h),  (154,y3l),
        (158,y27h), (159,y27l),
        (164,y7h),  (165,y7l),
        (169,y33h), (170,y33l),
        (174,y13h), (175,y13l),
        (179,y37h), (180,y37l),
        (184,y17h), (185,y17l),
        (189,y50h), (190,y50l),
    ]
    xml = apply_simple(xml, simple)

    # 帰寂日の月が1桁の場合、縦中横（eastAsianLayout）を除去
    if d.month < 10:
        xml = re.sub(r'\s*<w:eastAsianLayout w:id="-868065536"[^/]*/>', '', xml, count=1)

    # 弔主 名 置換
    ts, te = find_ruby_run(xml, 'たろう')
    xml = xml[:ts] + make_choshu_mei_run(
        data.get('choshu_mei_kanji', ''),
        data.get('choshu_mei_kana', '')
    ) + xml[te:]

    # 法号 置換
    hs, _  = find_ruby_run(xml, 'ほんしょう')
    _, he  = find_ruby_run(xml, 'れい')
    xml = xml[:hs] + make_hogo_ruby(data['hogo_pairs']) + xml[he:]

    # 黄色ハイライト（編集マーカー）を除去
    xml = re.sub(r'<w:highlight w:val="yellow"/>', '', xml)

    # docx に書き戻す（メモリ上で完結）
    buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(TEMPLATE_BYTES), 'r') as zin:
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'word/document.xml':
                    zout.writestr(item, xml.encode('utf-8'))
                else:
                    zout.writestr(item, zin.read(item.filename))
    return buf.getvalue()

# ─── Streamlit UI ─────────────────────────────────────────────────────────────
st.title('📄 追善法要のしおり 生成システム')
st.caption('故人・法号・弔主の情報を入力し、Wordファイルをダウンロードしてください。')

tab_input, tab_preview = st.tabs(['① 入力', '② 忌日・年回忌 確認'])

with tab_input:
    with st.form('main_form'):

        # ── 故人情報 ───────────────────────────────────────────────────────────
        st.subheader('故人情報')
        c1, c2 = st.columns(2)
        name_sei_k = c1.text_input('姓（漢字）', placeholder='例：村川')
        name_sei_r = c2.text_input('姓（よみ）', placeholder='例：むらかわ')
        c3, c4 = st.columns(2)
        name_mei_k = c3.text_input('名（漢字）', placeholder='例：照子')
        name_mei_r = c4.text_input('名（よみ）', placeholder='例：てるこ')
        c5, c6 = st.columns(2)
        age = c5.number_input('行年（歳）', min_value=1, max_value=130, value=None,
                              placeholder='例：77')
        death_date = c6.date_input('帰寂日', value=None,
                                   min_value=date(2019, 1, 1),
                                   max_value=date(2099, 12, 31),
                                   format='YYYY/MM/DD')

        st.divider()

        # ── 法号 ───────────────────────────────────────────────────────────────
        st.subheader('法号（戒名）')
        st.caption('漢字と読みをペアで入力してください。最後のペアは必ず「霊」「れい」にしてください。')

        if 'hogo_count' not in st.session_state:
            st.session_state.hogo_count = 5

        hogo_pairs = []
        defaults = [
            ('', ''), ('', ''), ('', ''), ('', ''), ('霊', 'れい')
        ]
        for i in range(st.session_state.hogo_count):
            ha, hb = st.columns(2)
            dk = defaults[i][0] if i < len(defaults) else ''
            dr = defaults[i][1] if i < len(defaults) else ''
            hk = ha.text_input(f'漢字 {i+1}', value=dk, key=f'hk{i}',
                               placeholder='例：昇進院')
            hr = hb.text_input(f'よみ {i+1}', value=dr, key=f'hr{i}',
                               placeholder='例：しょうしんいん')
            if hk or hr:
                hogo_pairs.append((hk, hr))

        ca, cb = st.columns(2)
        if ca.form_submit_button('＋ ペアを追加'):
            st.session_state.hogo_count += 1
            st.rerun()
        if cb.form_submit_button('－ ペアを削除') and st.session_state.hogo_count > 1:
            st.session_state.hogo_count -= 1
            st.rerun()

        st.divider()

        # ── 弔主 ───────────────────────────────────────────────────────────────
        st.subheader('弔主')
        d1, d2 = st.columns(2)
        choshu_sei_k = d1.text_input('姓（漢字）', placeholder='例：村久', key='cs_k')
        choshu_sei_r = d2.text_input('姓（よみ）', placeholder='例：むらく', key='cs_r')
        d3, d4 = st.columns(2)
        choshu_mei_k = d3.text_input('名（漢字）', placeholder='例：物部（ひらがなのみの場合は空欄）', key='cm_k')
        choshu_mei_r = d4.text_input('名（よみ）', placeholder='例：もののべ', key='cm_r')
        relation = st.text_input('続柄', placeholder='例：長男・妻・兄')

        st.divider()
        submitted = st.form_submit_button('✅ 確認して生成する', use_container_width=True,
                                          type='primary')

# ─── 送信後の処理 ─────────────────────────────────────────────────────────────
if submitted:
    errors = []
    if not name_sei_k: errors.append('故人の姓（漢字）')
    if not name_sei_r: errors.append('故人の姓（よみ）')
    if not name_mei_k: errors.append('故人の名（漢字）')
    if not name_mei_r: errors.append('故人の名（よみ）')
    if not age:        errors.append('行年')
    if not death_date: errors.append('帰寂日')
    if not hogo_pairs: errors.append('法号')
    if not choshu_sei_k: errors.append('弔主の姓（漢字）')
    if not choshu_mei_r: errors.append('弔主の名（よみ）')

    if errors:
        st.error('以下の項目を入力してください：' + '、'.join(errors))
    else:
        data = {
            'name_sei_kanji': name_sei_k,
            'name_sei_kana':  name_sei_r,
            'name_mei_kanji': name_mei_k,
            'name_mei_kana':  name_mei_r,
            'age':            int(age),
            'death_date':     death_date.isoformat(),
            'choshu_sei_kanji': choshu_sei_k,
            'choshu_sei_kana':  choshu_sei_r,
            'choshu_mei_kanji': choshu_mei_k,
            'choshu_mei_kana':  choshu_mei_r or choshu_mei_k,
            'relation':         relation,
            'hogo_pairs':       hogo_pairs,
        }

        ig, nk = calc_dates(death_date)

        with tab_preview:
            st.subheader('忌日一覧')
            ig_data = {k: v.strftime('%-m月%-d日') for k, v in ig.items()}
            col_a, col_b = st.columns(2)
            items = list(ig_data.items())
            for k, v in items[:4]:
                col_a.metric(k + '忌', v)
            for k, v in items[4:]:
                col_b.metric(k + '忌', v)

            st.subheader('年回忌一覧')
            nk_data = {k: v.strftime('%Y年%-m月%-d日') for k, v in nk.items()}
            for k, v in nk_data.items():
                st.text(f'{k}：{v}')

        with st.spinner('Wordファイルを生成中...'):
            try:
                docx_bytes = generate_docx(data)
                fname = f"追善法要のしおり_{name_sei_k}{name_mei_k}.docx"
                st.success('✅ 生成完了！')
                st.download_button(
                    label='📥 Word ファイルをダウンロード',
                    data=docx_bytes,
                    file_name=fname,
                    mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    use_container_width=True
                )
            except Exception as e:
                st.error(f'生成エラー: {e}')
                st.exception(e)

with tab_preview:
    if not submitted:
        st.info('① 入力タブで情報を入力して「確認して生成する」を押してください。')
