"""
excel_writer.py  ─  xlsxwriter 기반 주간보고서 Excel 생성
핵심 변경사항:
  • openpyxl → xlsxwriter: inlineStr → sharedStrings (XML 오류 완전 해결)
  • write_rich_string으로 * 이후 붉은색 텍스트 (정상 동작)
  • 데이터 행 병합 없음 (안정성 우선)
"""
from __future__ import annotations
import re
from datetime import datetime, date as date_type

import xlsxwriter

# ── 상수 ─────────────────────────────────────────────────────────────
DATA_ROW  = 13      # 0-indexed (Excel 14행)
DATA_EROW = 81      # 0-indexed (Excel 82행)
# 컬럼 0-indexed: A=0 B=1 C=2 D=3 E=4 F=5 G=6 H=7 I=8 J=9 K=10 L=11 M=12 N=13 O=14 P=15

# ── 제목 정리: [#코드#코드] 블록 제거 (H열 표시용) ─────────────────
_CODE_BLK_RE = re.compile(r"\s*\[\s*#\s*[A-Za-z0-9]{5,8}\s*#\s*[A-Z]{2}\s*\d{2}\s*-\s*\d{2}\s*\]\s*")

def _clean_subject(s: str) -> str:
    """[#PC#FC] 코드블록만 제거. 나머지 [Non-KPI]/[KPI]/[메뉴얼] 등은 유지."""
    return _CODE_BLK_RE.sub(" ", s).strip()


# ════════════════════════════════════════════════════════════════════
# 포맷 정의
# ════════════════════════════════════════════════════════════════════

def _fmts(wb: xlsxwriter.Workbook) -> dict:
    """모든 셀 포맷을 미리 생성해 dict로 반환"""
    f = wb.add_format   # 단축

    THIN, THICK = 1, 5
    BLU_DARK = '#1F4E79'
    BLU_NEXT = '#000000'   # 차주 글씨: 검정 (이미지 기준)
    HDR_FILL = '#595959'   # 헤더 배경: 진회색
    HDR_TXT  = '#FFFFFF'   # 헤더 글씨: 흰색
    TOT_FILL = '#FFE699'
    RED      = '#FF0000'

    B1 = {'border': THIN}

    def cell(**kw):
        base = {'font_name':'Calibri','font_size':11,'valign':'vcenter', **B1}
        return f({**base, **kw})

    return {
        # ── 템플릿 행 ─────────────────────────────────────────────
        'title_week'  : f({'font_name':'Calibri','font_size':16,'bold':True,
                            'font_color':BLU_DARK,'align':'center','valign':'vcenter',
                            'bottom':THICK}),
        'title_main'  : f({'font_name':'Calibri','font_size':26,'bold':True,
                            'valign':'vcenter','bottom':THICK}),
        'title_dept'  : f({'font_name':'Calibri','font_size':12,'bold':True,
                            'align':'right','valign':'vcenter','bottom':THICK}),
        'title_blank' : f({'bottom':THICK}),

        # M/H 집계
        'mh_label'    : cell(bold=True, align='center'),
        'mh_next_lbl' : cell(bold=True, align='center'),   # 차주 라벨도 검정
        'mh_val'      : cell(align='center', num_format='0.0'),
        'mh_pct'      : cell(align='center', num_format='0%'),
        'mh_nval'     : cell(align='center', num_format='0.0', font_color=BLU_NEXT),
        'mh_npct'     : cell(align='center', num_format='0%',  font_color=BLU_NEXT),
        'total_lbl'   : cell(bold=True, align='center', bg_color=TOT_FILL),
        'total_val'   : cell(bold=True, align='center', num_format='0.0', bg_color=TOT_FILL),
        'total_pct'   : cell(bold=True, align='center', num_format='0%',  bg_color=TOT_FILL),
        'total_nval'  : cell(bold=True, align='center', num_format='0.0',
                             font_color=BLU_NEXT, bg_color=TOT_FILL),
        'total_npct'  : cell(bold=True, align='center', num_format='0%',
                             font_color=BLU_NEXT, bg_color=TOT_FILL),
        'merged_empty': cell(),  # 병합 빈 셀

        # 컬럼 헤더 (12~13행)
        'hdr'         : cell(bold=True, align='center', text_wrap=True,
                             bg_color=HDR_FILL, font_color=HDR_TXT),
        'hdr_next'    : cell(bold=True, align='center', text_wrap=True,
                             bg_color=HDR_FILL, font_color=HDR_TXT),
        'hdr_none'    : f({'bg_color':'#FFFFFF'}),  # L열 구분선

        # 데이터 행
        'data_c'      : cell(align='center', text_wrap=True),   # C 구분
        'data_text'   : cell(align='center', num_format='@'),   # D 코드
        'data_fc'     : cell(align='center'),                   # E FC
        'data_name'   : cell(align='left',   text_wrap=True),   # F 업무명
        'data_date'   : cell(align='left',   num_format='yyyy-mm-dd'),
        'data_h'      : cell(align='left',   text_wrap=True),   # H 제목
        'data_i'      : cell(align='left',   text_wrap=True),   # I 상세(plain)
        'data_j'      : cell(align='center', num_format='0.0'), # J 금주
        'data_k'      : cell(align='center', num_format='0%'),  # K 금주%
        'data_m'      : cell(align='center', num_format='0.0',  # M 차주
                             font_color=BLU_NEXT),
        'data_n'      : cell(align='center', num_format='0%',   # N 차주%
                             font_color=BLU_NEXT),
        'data_blank'  : cell(),                                 # O,P 빈칸
        'data_l'      : f({'bg_color':'#FFFFFF'}),              # L 구분선

        # Rich text run 포맷 (border/fill 없음 - inline 전용)
        'run_blk'     : wb.add_format({'font_name':'Calibri','font_size':11}),
        'run_red'     : wb.add_format({'font_name':'Calibri','font_size':11,
                                        'color': RED}),
    }


# ════════════════════════════════════════════════════════════════════
# Rich text 헬퍼
# ════════════════════════════════════════════════════════════════════

def _write_body(ws, row: int, col: int, body: str, fmts: dict,
                merge_end_row: int = None):
    """
    I열(수행업무 상세) 쓰기
      1) \r 제거 → _x000D_ 방지
      2) * 이후 모든 줄 붉은색 (found_star 플래그)
      3) merge_end_row 지정 시 해당 행까지 I열 병합
    """
    # ── 병합 처리 ─────────────────────────────────────────────────
    do_merge = merge_end_row is not None and merge_end_row > row

    if not body:
        if do_merge:
            ws.merge_range(row, col, merge_end_row, col, '', fmts['data_i'])
        else:
            ws.write_blank(row, col, fmts['data_i'])
        return

    # ── \r 제거: 아웃룩 본문의 \r\n → \n 변환 (_x000D_ 방지) ──
    body = body.replace('\r\n', '\n').replace('\r', '\n')

    if '*' not in body:
        if do_merge:
            ws.merge_range(row, col, merge_end_row, col, body, fmts['data_i'])
        else:
            ws.write(row, col, body, fmts['data_i'])
        return

    lines = body.split('\n')
    parts: list[tuple] = []
    found_star = False   # True가 된 이후 모든 줄은 red

    for i, line in enumerate(lines):
        nl = '\n' if i < len(lines) - 1 else ''

        if found_star:
            # ★ * 이후 줄: 전부 붉은색
            t = line + nl
            if t:
                parts.append((fmts['run_red'], t))
        else:
            idx = line.find('*')
            if idx >= 0:
                found_star = True
                before = line[:idx]
                after  = line[idx:] + nl
                if before:
                    parts.append((fmts['run_blk'], before))
                if after:
                    parts.append((fmts['run_red'], after))
            else:
                t = line + nl
                if t:
                    parts.append((fmts['run_blk'], t))

    if not parts:
        ws.write_blank(row, col, fmts['data_i'])
        return

    # 파트 1개: 단색 write
    if len(parts) == 1:
        fmt, txt = parts[0]
        if fmt is fmts['run_red']:
            red_cell = ws.workbook.add_format({
                'font_name':'Calibri','font_size':11,'color':'#FF0000',
                'border':1,'valign':'vcenter','align':'left','text_wrap':True,
            })
            ws.write(row, col, txt, red_cell)
        else:
            ws.write(row, col, txt, fmts['data_i'])
        return

    # 파트 2개+: write_rich_string
    tokens = []
    for fmt, txt in parts:
        tokens += [fmt, txt]
    tokens.append(fmts['data_i'])

    try:
        if do_merge:
            ws.merge_range(row, col, merge_end_row, col, '')
        ws.write_rich_string(row, col, *tokens)
    except Exception:
        if do_merge:
            ws.merge_range(row, col, merge_end_row, col, body, fmts['data_i'])
        else:
            ws.write(row, col, body, fmts['data_i'])


# ════════════════════════════════════════════════════════════════════
# 공개 진입점
# ════════════════════════════════════════════════════════════════════

# ── 구분 정렬 순서 (create_excel 정렬용) ──────────────────────────
_GUBUN_SORT = {"수행 (KPI)": 0, "입찰": 1, "기타 (KPI)": 2, "기타": 3}

def _sort_body_key(d: dict) -> str:
    """I열 병합 기준용 본문 키: * 이후 제거"""
    body = (d.get('body') or '').replace('\r\n', '\n').replace('\r', '\n')
    m = re.search(r'\*', body)
    return body[:m.start()].rstrip() if m else body


def create_excel(rows: list[dict], output_path: str) -> None:
    # ── 정렬: 날짜 우선, 같은 내용은 연속 배치 ───────────────────
    # 방식: 1차 날짜순 정렬 → 2차 OrderedDict로 내용별 그룹화
    # (g_min defaultdict 키 미스매치 문제 방지용 안정적 구현)
    import datetime as _dt_mod
    from collections import OrderedDict

    def _row_date(r):
        dt = r.get('date')
        if hasattr(dt, 'date'): dt = dt.date()
        return dt or _dt_mod.date.min

    def _content_key(r):
        """병합/그룹 기준 키: (구분, PC, FC, 제목, 본문)"""
        return (r.get('gubun',''), r.get('project_code',''),
                r.get('func_code',''), r.get('subject',''),
                _sort_body_key(r))

    def _stable_group_sort(rlist):
        """
        1) 날짜 우선 정렬 (같은 날짜면 수행KPI 우선)
        2) OrderedDict로 같은 내용 그룹화 (첫 등장 순서 유지)
        3) 그룹 내 날짜 오름차순 정렬
        → 결과: 가장 이른 날짜의 내용이 맨 앞, 같은 내용은 연속 배치
        """
        # 1차: 날짜 기준 정렬 (첫 등장 순서 결정)
        by_date = sorted(rlist, key=lambda r: (
            _row_date(r),
            _GUBUN_SORT.get(r.get('gubun',''), 9),
            r.get('project_code',''),
            r.get('func_code',''),
            r.get('subject',''),
            _sort_body_key(r),
        ))
        # 2차: 같은 내용 그룹화 (첫 등장 날짜 순서 유지)
        groups = OrderedDict()
        for r in by_date:
            ck = _content_key(r)
            if ck not in groups:
                groups[ck] = []
            groups[ck].append(r)
        # 3차: 각 그룹 내 날짜 오름차순 + 순서대로 병합
        result = []
        for rows_in_group in groups.values():
            result.extend(sorted(rows_in_group, key=_row_date))
        return result

    this_rows = [r for r in rows if r.get('source') != 'next_week_only']
    next_rows  = [r for r in rows if r.get('source') == 'next_week_only']
    rows = _stable_group_sort(this_rows) + _stable_group_sort(next_rows)

    wb = xlsxwriter.Workbook(output_path, {'strings_to_numbers': False})
    fmts = _fmts(wb)

    ws = wb.add_worksheet('Weekly Report')
    _set_dims(ws)
    _template(ws, fmts, len(rows))
    _data(ws, rows, fmts, wb)
    _plant_mh(wb, rows, fmts)

    wb.close()


# ════════════════════════════════════════════════════════════════════
# 시트 치수
# ════════════════════════════════════════════════════════════════════

def _set_dims(ws):
    widths = [2.5, 2.625, 13.75, 9.375, 12.75, 31.5, 12.625,
              34.625, 46.375, 11.75, 11.75, 0.875, 11.75, 13.0, 38.0, 15.625]
    for i, w in enumerate(widths):
        ws.set_column(i, i, w)

    row_heights = {0:42, 1:9, 2:24, 3:26.25, 4:26.25, 5:26.25,
                   6:26.25, 7:9.75, 8:26.25, 9:26.25, 10:17.25, 11:22.5, 12:16.5}
    for r, h in row_heights.items():
        ws.set_row(r, h)


# ════════════════════════════════════════════════════════════════════
# 템플릿 (1~13행)
# ════════════════════════════════════════════════════════════════════
# 수식에서 Excel 행번호(1-based) 사용: DATA_ROW+1=14, DATA_EROW+1=82

DS  = DATA_ROW  + 1   # Excel row 14
DE  = DATA_EROW + 1   # Excel row 82

def _template(ws, fmts: dict, num_rows: int = 68):
    """
    num_rows: 실제 데이터 행 수 (SUMIF 끝 행 동적 계산에 사용)
    """
    f = fmts
    _de = DATA_ROW + max(num_rows, 1) + 1  # 동적 끝 행 (최소 DATA_ROW+2)

    # ── 1행: 타이틀 ──
    ws.merge_range(0, 2, 0, 3,
        f'=IFERROR(TEXT(G{DS},"m")&"월 "&INT((DAY(G{DS})-1)/7)+1&"주차","")',
        f['title_week'])
    ws.write(0, 4, 'Weekly Report', f['title_main'])
    for c in range(5, 14):
        ws.write_blank(0, c, f['title_blank'])
    ws.merge_range(0, 14, 0, 15, '플랜트사업본부 설계실', f['title_dept'])

    # ── 3행: M/H 집계 헤더 ──
    ws.write(2, 8, '주간 M/H 실적 집계', f['mh_label'])
    ws.merge_range(2, 9, 2, 10, '금주', f['mh_label'])
    ws.merge_range(2, 12, 2, 13, '차주', f['mh_next_lbl'])

    # ── 4~7행: 카테고리 집계 ──
    cats = ['수행 (KPI)', '입찰', '기타 (KPI)', '기타']
    for i, cat in enumerate(cats):
        r = 3 + i   # 0-indexed row 3~6
        er = r + 1  # Excel 1-based
        ws.write(r, 8,  cat,  f['mh_label'])
        ws.write_formula(r, 9,
            f'=SUMIF($C${DS}:$C${_de},I{er},$J${DS}:$J${_de})', f['mh_val'], 0)
        ws.write_formula(r, 10,
            f'=IFERROR(J{er}/$J$9,0)', f['mh_pct'], 0)
        ws.write_blank(r, 11, f['mh_label'])
        ws.write_formula(r, 12,
            f'=SUMIF($C${DS}:$C${DE},I{er},$M${DS}:$M${DE})', f['mh_nval'], 0)
        ws.write_formula(r, 13,
            f'=IFERROR(M{er}/$M$9,0)', f['mh_npct'], 0)

    # ── 9행: 총계 ──
    ws.merge_range(8, 2, 9, 3, '', f['merged_empty'])
    ws.write(8, 8,  '총   계',          f['total_lbl'])
    ws.write_formula(8, 9,  '=SUM(J4:J7)', f['total_val'], 0)
    ws.write_formula(8, 10, '=SUM(K4:K7)', f['total_pct'], 0)
    ws.write_blank(8, 11, f['total_lbl'])
    ws.write_formula(8, 12, '=SUM(M4:M7)', f['total_nval'], 0)
    ws.write_formula(8, 13, '=SUM(N4:N7)', f['total_npct'], 0)

    # ── 12~13행: 컬럼 헤더 ──
    headers_12 = {
        2: '구분', 3: 'Project\nCode', 4: 'Function Code',
        6: 'Date', 7: '수행 업무 (제목)', 8: '수행 업무 (상세)',
        9: '금주', 12: '차주', 14: '보직자 Feedback', 15: '비고',
    }
    for c, txt in headers_12.items():
        if c == 12:
            ws.merge_range(11, c, 11, c+1, txt, f['hdr_next'])
        elif c == 9:
            ws.merge_range(11, c, 11, c+1, txt, f['hdr'])
        else:
            ws.merge_range(11, c, 12, c, txt, f['hdr']) if c in (2,3,6,7,8,14,15) \
                else ws.write(11, c, txt, f['hdr'])

    # 13행 서브헤더
    ws.merge_range(11, 4, 11, 5, '', f['hdr'])   # Function Code 병합
    ws.write(12, 4, 'Code',        f['hdr'])
    ws.write(12, 5, 'Description', f['hdr'])
    ws.write(12, 9, '시간',        f['hdr'])
    ws.write(12, 10,'비중',        f['hdr'])
    ws.write(12, 12,'시간',        f['hdr_next'])
    ws.write(12, 13,'비중',        f['hdr_next'])
    ws.write_blank(11, 11, f['hdr_none'])
    ws.write_blank(12, 11, f['hdr_none'])


# ════════════════════════════════════════════════════════════════════
# 데이터 행 (14행~)
# ════════════════════════════════════════════════════════════════════

def _est_h(body: str, subject: str = "") -> float:
    """
    행 높이 추정 (pt 단위)
    개선 포인트:
      - col_i=32: I열(상세) 실효 폭. 맑은 고딕 11pt 한글은 Excel
        'char unit' 기준 실제보다 넓어서 46→32로 보수적으로 설정
      - col_h=24: H열(제목) 실효 폭 (34.625칸 → 보수적 24)
      - line_pt=16: 11pt 폰트의 실제 줄 높이(행간 포함)
      - margin=1.25: 여유 마진 25% 추가
    """
    COL_I   = 42    # I열 실효 문자폭 (46.375칸, 한글 1.8배 기준)
    COL_H   = 30    # H열 실효 문자폭 (34.625칸 기준)
    LINE_PT = 13.5  # 줄당 높이(pt) - Calibri 11pt 기준
    MARGIN  = 1.10  # 여유 마진 10%
    MIN_H   = 28.0
    MAX_H   = 300.0

    def count_lines(text: str, col_w: int) -> int:
        if not text:
            return 0
        total = 0
        for ln in text.split("\n"):
            ln_len = sum(2 if ord(c) > 127 else 1 for c in ln)
            total += max(1, -(-ln_len // col_w))
        return total

    lines_i = count_lines(body,    COL_I)
    lines_h = count_lines(subject, COL_H)
    total   = max(lines_i, lines_h, 1)

    h = total * LINE_PT * MARGIN
    return min(MAX_H, max(MIN_H, h))


def _data(ws, rows: list[dict], fmts: dict, wb):
    f = fmts
    DS_str = f'C{DATA_ROW+1}:C{DATA_EROW+1}'

    # 드롭다운 데이터 유효성 검사
    ws.data_validation(DS_str, {
        'validate': 'list',
        'source': ['수행 (KPI)', '입찰', '기타 (KPI)', '기타'],
        'error_message': '목록에서 선택하세요',
    })

    # ── 병합을 위해 행 재정렬 (같은 내용이 연속되도록) ─────────────
    # 정렬 기준: 구분 → PC → FC → 제목 → 본문 → 날짜
    # 이렇게 해야 같은 내용의 여러 날짜 행이 연속되어 I열 병합 가능
    GUBUN_ORD = {"수행 (KPI)": 0, "입찰": 1, "기타 (KPI)": 2, "기타": 3}

    def _body_key(d):
        body = (d.get('body') or '').replace('\r\n','\n').replace('\r','\n')
        m = re.search(r'\*', body)
        return body[:m.start()].rstrip() if m else body

    def _sort_key(d):
        dt = d.get('date')
        if hasattr(dt, 'date'): dt = dt.date()
        return (
            GUBUN_ORD.get(d.get('gubun',''), 9),
            d.get('project_code',''),
            d.get('func_code',''),
            _clean_subject(d.get('subject','')),
            _body_key(d),
            dt or __import__('datetime').date.min,
        )

    # this_week / next_week 분리 → 각각 정렬 → 합치기
    this_rows = sorted([r for r in rows if r.get('source') != 'next_week_only'], key=_sort_key)
    next_rows = sorted([r for r in rows if r.get('source') == 'next_week_only'], key=_sort_key)
    rows = this_rows + next_rows

    # ── I열 병합 그룹 사전 계산 ──────────────────────────────────
    # 같은 (gubun, pc, fc, subject, 별표제거_body) 가 연속될 때 I열 병합

    # merge_map[row_idx] = (r_excel_start, r_excel_end)
    #   r_excel_start 행에만 body 쓰고, 나머지는 blank merge
    merge_map: dict[int, tuple[int,int]] = {}
    i = 0
    while i < len(rows):
        key_i = (_body_key(rows[i]),
                 rows[i].get('gubun',''),
                 rows[i].get('project_code',''),
                 rows[i].get('func_code',''),
                 rows[i].get('subject',''))
        if not key_i[0]:          # 빈 body는 병합 안 함
            i += 1
            continue
        j = i + 1
        while j < len(rows):
            key_j = (_body_key(rows[j]),
                     rows[j].get('gubun',''),
                     rows[j].get('project_code',''),
                     rows[j].get('func_code',''),
                     rows[j].get('subject',''))
            if key_j == key_i:
                j += 1
            else:
                break
        if j > i + 1:             # 2행 이상 동일 → 병합 대상
            r_start = DATA_ROW + i
            r_end   = DATA_ROW + j - 1
            for k in range(i, j):
                merge_map[k] = (r_start, r_end)
        i = j

    # ── 행 높이 사전 계산 ────────────────────────────────────────
    # 병합 그룹은 총 높이를 행 수로 균등 분배
    MIN_ROW_H = 20.0
    processed_merges: set[tuple] = set()
    row_heights: dict[int, float] = {}

    for i, d in enumerate(rows):
        body    = d.get('body','') or ''
        subject = d.get('subject','') or ''
        if i in merge_map:
            r_start, r_end = merge_map[i]
            key = (r_start, r_end)
            if key not in processed_merges:
                processed_merges.add(key)
                total_h  = _est_h(body, subject)
                num_rows = r_end - r_start + 1
                each_h   = max(MIN_ROW_H, total_h / num_rows)
                for k in range(r_start, r_end + 1):
                    row_heights[k - DATA_ROW] = each_h
        else:
            row_heights[i] = _est_h(body, subject)

    # ── 데이터 출력 ──────────────────────────────────────────────
    for i, d in enumerate(rows):
        r    = DATA_ROW + i
        er   = r + 1
        ws.set_row(r, row_heights.get(i, 28.0))

        # C 구분
        ws.write(r, 2,  d['gubun'],        f['data_c'])
        # D Project Code
        ws.write(r, 3,  d['project_code'], f['data_text'])
        # E Function Code
        ws.write(r, 4,  d['func_code'],    f['data_fc'])
        # F 업무명
        ws.write(r, 5,  d['func_name'],    f['data_name'])
        # G 날짜
        dt = d.get('date')
        if isinstance(dt, datetime):
            dt = datetime(dt.year, dt.month, dt.day)
        elif isinstance(dt, date_type):
            dt = datetime(dt.year, dt.month, dt.day)
        ws.write_datetime(r, 6, dt, f['data_date']) if dt else ws.write_blank(r, 6, f['data_date'])
        # H 수행업무(제목)
        ws.write(r, 7,  _clean_subject(d.get('subject','') or ''), f['data_h'])
        # I 수행업무(상세) - 병합 처리
        if i in merge_map:
            r_start, r_end = merge_map[i]
            if r == r_start:          # 첫 행: body 쓰기 + 병합
                _write_body(ws, r, 8, d.get('body','') or '', f,
                            merge_end_row=r_end)
            # 나머지 행: 이미 병합됨 → 아무것도 안 씀
        else:
            _write_body(ws, r, 8, d.get('body','') or '', f)
        # J 금주 시간
        ws.write(r, 9,  d.get('this_week_h', 0.0), f['data_j'])
        # K 금주 비중
        ws.write_formula(r, 10, f'=IFERROR(J{er}/$J$9,0)', f['data_k'], 0)
        # L 구분선
        ws.write_blank(r, 11, f['data_l'])
        # M 차주 시간
        ws.write(r, 12, d.get('next_week_h', 0.0), f['data_m'])
        # N 차주 비중
        ws.write_formula(r, 13, f'=IFERROR(M{er}/$M$9,0)', f['data_n'], 0)
        # O 보직자 Feedback
        ws.write_blank(r, 14, f['data_blank'])
        # P 비고
        ws.write_blank(r, 15, f['data_blank'])


# ════════════════════════════════════════════════════════════════════
# Plant M/H 탭
# ════════════════════════════════════════════════════════════════════

def _plant_mh(wb, rows: list[dict], fmts: dict):
    ws = wb.add_worksheet('Plant MH')

    this_rows = [r for r in rows if r.get('source') == 'this_week' and r.get('this_week_h', 0) > 0]
    if not this_rows:
        ws.write(0, 0, '이번주 [실적] 데이터 없음')
        return

    dates = sorted(set(
        datetime(r['date'].year, r['date'].month, r['date'].day).date()
        if isinstance(r['date'], datetime) else r['date']
        for r in this_rows
    ))

    DAY_NAMES = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun']
    hdr_fmt = wb.add_format({'font_name':'Calibri','font_size':10,'bold':True,
                              'border':1,'align':'center','valign':'vcenter',
                              'bg_color':'#D9E1F2','text_wrap':True})
    dat_fmt = wb.add_format({'font_name':'Calibri','font_size':10,
                              'border':1,'valign':'vcenter'})
    num_fmt = wb.add_format({'font_name':'Calibri','font_size':10,
                              'border':1,'align':'center','valign':'vcenter','num_format':'0.0'})
    tot_fmt = wb.add_format({'font_name':'Calibri','font_size':10,'bold':True,
                              'border':1,'align':'center','valign':'vcenter',
                              'bg_color':'#FFE699','num_format':'0.0'})
    tot_lbl = wb.add_format({'font_name':'Calibri','font_size':10,'bold':True,
                              'border':1,'align':'center','valign':'vcenter','bg_color':'#FFE699'})

    # 열 너비
    ws.set_column(0, 0, 12)
    ws.set_column(1, 1, 12)
    ws.set_column(2, 2, 12)
    ws.set_column(3, 3, 30)
    for ci in range(4, 4 + len(dates) + 1):
        ws.set_column(ci, ci, 8)

    # 헤더
    ws.set_row(0, 36)
    for ci, hd in enumerate(['구분','Project Code','Func. Code','Description']):
        ws.write(0, ci, hd, hdr_fmt)
    for di, d in enumerate(dates):
        ws.write(0, 4+di, f'{d.month}/{d.day}\n({DAY_NAMES[d.weekday()]})', hdr_fmt)
    ws.write(0, 4+len(dates), '합계', hdr_fmt)

    # 집계
    agg: dict[tuple, dict] = {}
    order: list[tuple] = []
    for r in this_rows:
        key = (r['gubun'], r['project_code'], r['func_code'], r['func_name'])
        dt  = r['date'].date() if isinstance(r['date'], datetime) else r['date']
        if key not in agg:
            agg[key] = {d: 0.0 for d in dates}
            order.append(key)
        agg[key][dt] = round(agg[key].get(dt, 0.0) + r.get('this_week_h', 0.0), 1)

    total_by_date = {d: 0.0 for d in dates}
    grand_total   = 0.0

    for ri, key in enumerate(order, 1):
        ws.set_row(ri, 18)
        gubun, pc, fc, desc = key
        hm = agg[key]
        row_total = round(sum(hm[d] for d in dates), 1)
        grand_total = round(grand_total + row_total, 1)
        ws.write(ri, 0, gubun, dat_fmt)
        ws.write(ri, 1, pc,    dat_fmt)
        ws.write(ri, 2, fc,    dat_fmt)
        ws.write(ri, 3, desc,  dat_fmt)
        for di, d in enumerate(dates):
            v = hm[d]
            if v:
                ws.write(ri, 4+di, v, num_fmt)
                total_by_date[d] = round(total_by_date[d] + v, 1)
            else:
                ws.write_blank(ri, 4+di, num_fmt)
        ws.write(ri, 4+len(dates), row_total if row_total else '', num_fmt)

    # 합계 행
    tr = len(order) + 1
    ws.set_row(tr, 20)
    ws.write(tr, 0, '합계', tot_lbl)
    ws.write_blank(tr, 1, tot_lbl)
    ws.write_blank(tr, 2, tot_lbl)
    ws.write_blank(tr, 3, tot_lbl)
    for di, d in enumerate(dates):
        v = total_by_date[d]
        ws.write(tr, 4+di, v if v else '', tot_fmt)
    ws.write(tr, 4+len(dates), grand_total if grand_total else '', tot_fmt)
