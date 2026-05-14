"""
追善法要のしおり＆言上文 生成システム
実行: streamlit run app.py
"""
import streamlit as st
from datetime import date, timedelta
import re, io, zipfile, os

st.set_page_config(page_title='法要書類 生成', page_icon='📄', layout='centered')

# ─── テンプレート読み込み ─────────────────────────────────────────────────────
DIR = os.path.dirname(__file__)

@st.cache_data
def load_template(fname):
    with open(os.path.join(DIR, fname), 'rb') as f:
        return f.read()

try:
    SHIORI_BYTES   = load_template('template_shiori.docx')
    GENJOU_BYTES   = load_template('template_genjou.docx')
except FileNotFoundError as e:
    st.error(f'テンプレートファイルが見つかりません: {e}')
    st.stop()

# ─── 定数 ─────────────────────────────────────────────────────────────────────
YOUBI = ['月', '火', '水', '木', '金', '土', '日']

LAST_CHAR_KANA = {
    '霊':  'れい',
    '沙弥': 'しゃみ',
    '法師': 'ほっし',
    '法尼': 'ほうに',
    '大徳': 'だいとく',
    '上人': 'しょうにん',
}

# 年回忌テンプレートの上位桁数/下位桁数（テンプレート固有）
NENKI_SPLITS = {
    '一周忌':   (3,1), '廿三回忌': (3,1),
    '三回忌':   (3,1), '廿七回忌': (2,2),
    '七回忌':   (2,2), '卅三回忌': (2,2),
    '十三回忌': (3,1), '卅七回忌': (2,2),
    '十七回忌': (2,2), '五十回忌': (3,1),
}

# ─── 日付・数値変換 ───────────────────────────────────────────────────────────
def calc_dates(d: date):
    def ann(y):
        try:    return d.replace(year=d.year+y)
        except: return d.replace(year=d.year+y, day=28)
    ig = {
        '初七日': d+timedelta(6),  '二七日': d+timedelta(13),
        '三七日': d+timedelta(20), '四七日': d+timedelta(27),
        '五七日': d+timedelta(34), '六七日': d+timedelta(41),
        '尽七日': d+timedelta(48), '百カ日': d+timedelta(99),
    }
    nk = {
        '一周忌': ann(1), '三回忌': ann(2), '七回忌': ann(6),
        '十三回忌': ann(12), '十七回忌': ann(16), '廿三回忌': ann(22),
        '廿七回忌': ann(26), '卅三回忌': ann(32), '卅七回忌': ann(36), '五十回忌': ann(49),
    }
    return ig, nk

def zen(n: int) -> str:
    """1桁→全角、2桁以上→半角（縦書き横倒し防止）"""
    return '０１２３４５６７８９'[n] if n < 10 else str(n)

def reiwa_zen(year: int) -> str:
    """西暦年→令和全角年（例: 2026→８）"""
    r = year - 2018
    fw = '０１２３４５６７８９'
    if r <= 0: return '？'
    return fw[r] if r < 10 else fw[r//10] + fw[r%10]

def fw_year(y: int, upper: int, lower: int) -> tuple:
    """年を全角4桁にしてテンプレートの分割数で2分割"""
    s = ''.join('０１２３４５６７８９'[int(c)] for c in str(y))
    return s[:upper], s[upper:]

def to_kanji(n: int) -> str:
    """整数→漢数字（1-99、言上文用）"""
    k = ['', '一','二','三','四','五','六','七','八','九']
    if n <= 0:  return '零'
    if n < 10:  return k[n]
    t, o = n // 10, n % 10
    s = ('' if t == 1 else k[t]) + '十'
    if o: s += k[o]
    return s

def reiwa_kanji(year: int) -> str:
    """西暦年→令和漢数字（例: 2026→八）"""
    return to_kanji(year - 2018)

def kanji_month_split(month: int) -> tuple:
    """月を漢字2ランに分割（言上文の十二月→十,二 形式）"""
    k = ['','一','二','三','四','五','六','七','八','九']
    if month <= 9:   return k[month], ''
    elif month == 10: return '十', ''
    else:             return '十', k[month - 10]

# ─── XML操作 ──────────────────────────────────────────────────────────────────
# 修正済みregex（w:textAlignment等を誤マッチしない）
RUNS_RE = re.compile(r'<w:t(?=[>\s])[^>]*>(.*?)</w:t>', re.DOTALL)

def get_runs(xml: str):
    return [(i, m.start(), m.end(), m.group(1))
            for i, m in enumerate(RUNS_RE.finditer(xml))]

def apply_simple(xml: str, replacements: list) -> str:
    runs = get_runs(xml)
    subs = sorted(
        [(runs[i][1], runs[i][2], v) for i, v in replacements],
        key=lambda x: -x[0]
    )
    for s, e, new in subs:
        orig = xml[s:e]
        m = re.match(r'<w:t([^>]*)>', orig)
        attrs = m.group(1) if m else ''
        xml = xml[:s] + f'<w:t{attrs}>{new}</w:t>' + xml[e:]
    return xml

def find_ruby_run(xml: str, kana: str) -> tuple:
    """kana を含む ruby の外側 <w:r...> の (start, end) を返す"""
    pos   = xml.index(f'>{kana}<')
    ruby  = xml.rindex('<w:ruby>', 0, pos)
    r_pos = [m.start() for m in re.finditer(r'<w:r[ >]', xml[:ruby])]
    start = r_pos[-1]
    end   = pos + re.search(r'</w:ruby>\s*</w:r>', xml[pos:]).end()
    return start, end

def make_hogo_ruby_simple(main_k, main_r, last_k, last_r, sz=60, hps=30, hpsRaise=58):
    """
    法号を2ブロックのrubyに簡略化:
      メイン部分(ハイライトあり) + 末尾(ハイライトなし)
    """
    def one(kanji, kana, hl):
        h = '<w:highlight w:val="yellow"/>' if hl else ''
        return f'''<w:r>
      <w:rPr>
        <w:rFonts w:ascii="HG正楷書体-PRO" w:eastAsia="HG正楷書体-PRO" w:hint="eastAsia"/>
        <w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>{h}
      </w:rPr>
      <w:ruby>
        <w:rubyPr>
          <w:rubyAlign w:val="distributeSpace"/>
          <w:hps w:val="{hps}"/>
          <w:hpsRaise w:val="{hpsRaise}"/>
          <w:hpsBaseText w:val="{sz}"/>
          <w:lid w:val="ja-JP"/>
        </w:rubyPr>
        <w:rt><w:r><w:rPr>
          <w:rFonts w:ascii="HG正楷書体-PRO" w:eastAsia="HG正楷書体-PRO"
            w:hAnsi="游明朝" w:hint="eastAsia"/>
          <w:sz w:val="{hps}"/><w:szCs w:val="{sz}"/>{h}
        </w:rPr><w:t>{kana}</w:t></w:r></w:rt>
        <w:rubyBase><w:r><w:rPr>
          <w:rFonts w:ascii="HG正楷書体-PRO" w:eastAsia="HG正楷書体-PRO" w:hint="eastAsia"/>
          <w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>{h}
        </w:rPr><w:t>{kanji}</w:t></w:r></w:rubyBase>
      </w:ruby>
    </w:r>'''
    return one(main_k, main_r, True) + '\n    ' + one(last_k, last_r, False)

def make_choshu_mei_run(mei_k: str, mei_r: str) -> str:
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

def strip_yellow(xml: str) -> str:
    return re.sub(r'<w:highlight w:val="yellow"/>', '', xml)

def pack_docx(template_bytes: bytes, xml: str) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(template_bytes), 'r') as zin:
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                zout.writestr(item,
                    xml.encode('utf-8') if item.filename == 'word/document.xml'
                    else zin.read(item.filename))
    return buf.getvalue()

# ─── しおり 生成 ──────────────────────────────────────────────────────────────
def generate_shiori(data: dict) -> bytes:
    with zipfile.ZipFile(io.BytesIO(SHIORI_BYTES)) as z:
        xml = z.read('word/document.xml').decode('utf-8')

    d      = date.fromisoformat(data['death_date'])
    ig, nk = calc_dates(d)
    is_rei = (data['last_char'] == '霊')

    def md(dt): return zen(dt.month), zen(dt.day), YOUBI[dt.weekday()]
    def yr(key): u, l = NENKI_SPLITS[key]; return fw_year(nk[key].year, u, l)

    n1, n23 = nk['一周忌'], nk['廿三回忌']
    y1h, y1l   = yr('一周忌');  y23h,y23l = yr('廿三回忌')
    y3h, y3l   = yr('三回忌');  y27h,y27l = yr('廿七回忌')
    y7h, y7l   = yr('七回忌');  y33h,y33l = yr('卅三回忌')
    y13h,y13l  = yr('十三回忌'); y37h,y37l = yr('卅七回忌')
    y17h,y17l  = yr('十七回忌'); y50h,y50l = yr('五十回忌')

    # ── テンプレートインデックス（template_shiori.docx 固有）──────────────
    # 忌日 月/日/曜
    i0,i1,i2,i3 = [ig[k] for k in ('初七日','五七日','二七日','六七日')]
    i4,i5,i6,i7 = [ig[k] for k in ('三七日','尽七日','四七日','百カ日')]

    simple = [
        # 弔主 姓
        (4,  data['choshu_sei_kana']),
        (5,  data['choshu_sei_kanji']),
        # 続柄
        (11, data['relation']),
        # 帰寂日
        (39, reiwa_zen(d.year)),
        (42, zen(d.month)),
        (45, zen(d.day)),
        # 帰寂 or 遷化
        (48, '帰寂' if is_rei else '遷化'),
        # 俗名
        (54, data['name_sei_kana']),
        (55, data['name_sei_kanji']),
        (57, data['name_mei_kana']),
        (58, data['name_mei_kanji']),
        # 行年 or 法寿
        (61, '行年' if is_rei else '法寿'),
        (62, str(data['age'])),
        # 忌日（月/日/曜）
        (68, md(i0)[0]), (70, md(i0)[1]), (73, md(i0)[2]),
        (80, md(i1)[0]), (82, md(i1)[1]), (85, md(i1)[2]),
        (90, md(i2)[0]), (92, md(i2)[1]), (95, md(i2)[2]),
        (102,md(i3)[0]), (104,md(i3)[1]), (107,md(i3)[2]),
        (112,md(i4)[0]), (114,md(i4)[1]), (117,md(i4)[2]),
        (123,md(i5)[0]), (125,md(i5)[1]), (128,md(i5)[2]),
        (133,md(i6)[0]), (135,md(i6)[1]), (138,md(i6)[2]),
        (143,md(i7)[0]), (145,md(i7)[1]), (148,md(i7)[2]),
        # 年回忌
        (154,y1h),  (155,y1l),
        (157,zen(n1.month)),  (159,zen(n1.day)),
        (164,y23h), (165,y23l),
        (167,zen(n23.month)), (169,zen(n23.day)),
        (173,y3h),  (174,y3l),
        (178,y27h), (179,y27l),
        (184,y7h),  (185,y7l),
        (189,y33h), (190,y33l),
        (194,y13h), (195,y13l),
        (199,y37h), (200,y37l),
        (204,y17h), (205,y17l),
        (209,y50h), (210,y50l),
    ]
    xml = apply_simple(xml, simple)

    # 月が1桁の場合 eastAsianLayout（縦中横）を除去
    if d.month < 10:
        xml = re.sub(r'\s*<w:eastAsianLayout w:id="-868065536"[^/]*/>', '', xml, count=1)

    # 弔主 名
    ts, te = find_ruby_run(xml, 'はなこ')  # 弔主 名 kana
    xml = xml[:ts] + make_choshu_mei_run(
        data.get('choshu_mei_kanji', ''),
        data.get('choshu_mei_kana', '')) + xml[te:]

    # 法号：ほんせんいん 〜 れい を2ブロックrubyに置換
    hs, _  = find_ruby_run(xml, 'ほんせんいん')
    _, he  = find_ruby_run(xml, 'れい')
    last_k = data['last_char']
    last_r = LAST_CHAR_KANA[last_k]
    xml = xml[:hs] + make_hogo_ruby_simple(
        data['hogo_kanji'], data['hogo_kana'],
        last_k, last_r, sz=60, hps=30, hpsRaise=58) + xml[he:]

    xml = strip_yellow(xml)
    return pack_docx(SHIORI_BYTES, xml)

# ─── 言上文 生成 ──────────────────────────────────────────────────────────────
def generate_genjou(data: dict) -> bytes:
    with zipfile.ZipFile(io.BytesIO(GENJOU_BYTES)) as z:
        xml = z.read('word/document.xml').decode('utf-8')

    d      = date.fromisoformat(data['death_date'])
    is_rei = (data['last_char'] == '霊')
    m_hi, m_lo = kanji_month_split(d.month)

    simple = [
        # 俗名 姓名
        (3,  data['name_sei_kana']),
        (4,  data['name_sei_kanji']),
        (6,  data['name_mei_kana']),
        (7,  data['name_mei_kanji']),
        # 行年 or 法寿 ラベル
        (8,  '行年' if is_rei else '法寿'),
        # 行年数（漢数字）
        (10, to_kanji(int(data['age']))),
        # 令和年（漢数字）
        (13, reiwa_kanji(d.year)),
        # 月（十二→十,二 形式）
        (16, m_hi),
        (17, m_lo),
        # 日（漢数字）
        (19, to_kanji(d.day)),
        # 帰寂 or 遷化（2文字分割）
        (21, '帰' if is_rei else '遷'),
        (22, '寂' if is_rei else '化'),
        # 弔主 姓
        (43, data['choshu_sei_kana']),
        (44, data['choshu_sei_kanji']),
        # 続柄
        (48, data['relation']),
    ]
    xml = apply_simple(xml, simple)

    # 弔主 名
    ts, te = find_ruby_run(xml, 'はなこ')
    xml = xml[:ts] + make_choshu_mei_run(
        data.get('choshu_mei_kanji', ''),
        data.get('choshu_mei_kana', '')) + xml[te:]

    # 法号：ほんせんいん 〜 れい を2ブロックrubyに置換（言上文用サイズ）
    hs, _  = find_ruby_run(xml, 'ほんせんいん')
    _, he  = find_ruby_run(xml, 'れい')
    last_k = data['last_char']
    last_r = LAST_CHAR_KANA[last_k]
    xml = xml[:hs] + make_hogo_ruby_simple(
        data['hogo_kanji'], data['hogo_kana'],
        last_k, last_r, sz=72, hps=36, hpsRaise=70) + xml[he:]

    xml = strip_yellow(xml)
    return pack_docx(GENJOU_BYTES, xml)

# ─── Streamlit UI ─────────────────────────────────────────────────────────────
st.title('📄 法要書類 生成システム')
st.caption('情報を入力して「しおり」と「言上文」を同時に生成します。')

with st.form('main_form'):

    # ── 故人情報 ───────────────────────────────────────────────────────────────
    st.subheader('故人情報')

    c1, c2 = st.columns(2)
    name_sei_k = c1.text_input('姓（漢字）', placeholder='例：山田')
    name_mei_k = c2.text_input('名（漢字）', placeholder='例：太郎')

    c3, c4 = st.columns(2)
    name_sei_r = c3.text_input('姓（よみ）', placeholder='例：やまだ')
    name_mei_r = c4.text_input('名（よみ）', placeholder='例：たろう')

    c5, c6 = st.columns(2)
    age = c5.number_input('行年（歳）', min_value=1, max_value=130, value=None,
                          placeholder='例：93')
    death_date = c6.date_input('帰寂日', value=None,
                               min_value=date(2019,1,1), max_value=date(2099,12,31),
                               format='YYYY/MM/DD')

    st.divider()

    # ── 法号 ───────────────────────────────────────────────────────────────────
    st.subheader('法号（戒名）')
    st.caption('末尾（霊・沙弥など）を除いた本体部分を入力してください。')

    hc1, hc2 = st.columns(2)
    hogo_kanji = hc1.text_input('法号（漢字）', placeholder='例：本山院浄勲法田日太居士')
    hogo_kana  = hc2.text_input('法号（よみ）', placeholder='例：ほんせんいんじょうくんほうでんにったこじ')

    last_char = st.selectbox('末尾', list(LAST_CHAR_KANA.keys()))
    is_rei_preview = (last_char == '霊')
    if not is_rei_preview:
        st.info('「霊」以外が選択されました。帰寂 → **遷化**、行年 → **法寿** に自動変更されます。')

    st.divider()

    # ── 弔主 ───────────────────────────────────────────────────────────────────
    st.subheader('弔主')

    d1, d2 = st.columns(2)
    choshu_sei_k = d1.text_input('姓（漢字）', placeholder='例：鈴木', key='cs_k')
    choshu_mei_k = d2.text_input('名（漢字）', placeholder='例：花子（漢字なければ空欄）', key='cm_k')

    d3, d4 = st.columns(2)
    choshu_sei_r = d3.text_input('姓（よみ）', placeholder='例：すずき', key='cs_r')
    choshu_mei_r = d4.text_input('名（よみ）', placeholder='例：はなこ', key='cm_r')

    relation = st.text_input('続柄', placeholder='例：長男・妻・長女')

    st.divider()
    submitted = st.form_submit_button('✅ Word ファイルを生成する',
                                      use_container_width=True, type='primary')

# ─── 生成処理 ─────────────────────────────────────────────────────────────────
if submitted:
    errors = []
    if not name_sei_k: errors.append('故人の姓（漢字）')
    if not name_sei_r: errors.append('故人の姓（よみ）')
    if not name_mei_k: errors.append('故人の名（漢字）')
    if not name_mei_r: errors.append('故人の名（よみ）')
    if not age:        errors.append('行年')
    if not death_date: errors.append('帰寂日')
    if not hogo_kanji: errors.append('法号（漢字）')
    if not hogo_kana:  errors.append('法号（よみ）')
    if not choshu_sei_k: errors.append('弔主の姓（漢字）')
    if not choshu_mei_r and not choshu_mei_k: errors.append('弔主の名')

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
            'hogo_kanji':     hogo_kanji,
            'hogo_kana':      hogo_kana,
            'last_char':      last_char,
            'choshu_sei_kanji': choshu_sei_k,
            'choshu_sei_kana':  choshu_sei_r,
            'choshu_mei_kanji': choshu_mei_k,
            'choshu_mei_kana':  choshu_mei_r or choshu_mei_k,
            'relation':         relation,
        }

        ig, _ = calc_dates(death_date)

        # 忌日プレビュー
        with st.expander('忌日・年回忌を確認する'):
            cols = st.columns(2)
            for j, (k, v) in enumerate(ig.items()):
                cols[j % 2].metric(k+'忌', f'{v.month}月{v.day}日（{YOUBI[v.weekday()]}）')

        col_s, col_g = st.columns(2)
        with st.spinner('生成中...'):
            try:
                name = f"{name_sei_k}{name_mei_k}"
                s_bytes = generate_shiori(data)
                col_s.success('しおり 完成')
                col_s.download_button(
                    '📥 しおりをダウンロード',
                    data=s_bytes,
                    file_name=f'追善法要のしおり_{name}.docx',
                    mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    use_container_width=True)

                g_bytes = generate_genjou(data)
                col_g.success('言上文 完成')
                col_g.download_button(
                    '📥 言上文をダウンロード',
                    data=g_bytes,
                    file_name=f'言上文_{name}.docx',
                    mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    use_container_width=True)

            except Exception as e:
                st.error(f'生成エラー: {e}')
                st.exception(e)
