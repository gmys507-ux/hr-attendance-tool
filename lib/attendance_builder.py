"""
근태정리 핵심 로직 — v16 산출물과 동일 구조(5시트 + 수식 기반) 재현

시트 구성:
  1. 근태정리            : 메인 (SUMIFS/COUNTIF 수식, 조건부서식)
  2. 일자별_병합RAW      : Single Source of Truth (이름·부서·주차·날짜·요일·근무분·야간분·출처)
  3. ERP_원본_10명       : FLEX 전환자 원본 데이터
  4. 우이솔_야간상세     : 야간비율 검증용 (주간요약 + 일별 SUMIFS + 합계)
  5. 초과근로자_일자별RAW: 초과>0 인원 일자별 (사람별 블록, 우이솔 제외)

원티드스페이스(기본) + FLEX(ERP 전환자) 병합. 전환일 이후는 FLEX 우선(이중카운트 방지).
"""
import pandas as pd
from datetime import datetime, time, timedelta
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule

RAW_SHEET = '일자별_병합RAW'          # 메인 시트 SUMIFS가 참조하는 시트명
DAYS_SHORT = ['월', '화', '수', '목', '금', '토', '일']
DAYS_FULL = ['월요일', '화요일', '수요일', '목요일', '금요일', '토요일', '일요일']
NIGHT_PERSON = '우이솔'                # 야간상세 검증 대상


# ===== 시간 변환 =====
def to_min(v):
    """다양한 형식의 시간 값을 '분' 정수로 변환"""
    if pd.isna(v):
        return 0
    if isinstance(v, str):
        if v in ('0:00', '00:00', '-', ''):
            return 0
        try:
            parts = v.split(':')
            return int(parts[0]) * 60 + int(parts[1])
        except Exception:
            return 0
    if isinstance(v, time):
        return v.hour * 60 + v.minute
    if isinstance(v, datetime):
        return v.hour * 60 + v.minute
    if isinstance(v, (int, float)):
        return int(round(v * 24 * 60)) if v < 1 else int(v)
    return 0


def get_person_rows(df, name):
    """WORKING 우선, 없으면 ABSENCE도 사용 (박세진·황몽흔 케이스)"""
    sub_all = df[df['Name'] == name]
    sub_w = sub_all[sub_all['Status'] == 'WORKING']
    return sub_w if len(sub_w) > 0 else sub_all


def calc_weeks(start_date):
    """시작일부터 4주차 자동 계산. 주차 라벨은 '{월}월 {주차}주차'."""
    d = pd.Timestamp(start_date)
    weeks = []
    for i in range(4):
        w_start = d + timedelta(days=i * 7)
        w_end = w_start + timedelta(days=6)
        m = w_start.month
        # 주차 번호 = 그 달의 N번째 7일 구간 (v16 라벨 규칙과 일치)
        wk_num = (w_start.day - 1) // 7 + 1
        weeks.append({
            'name': f'{m}월 {wk_num}주차',
            'range': f'{w_start.month}/{w_start.day}~{w_end.month}/{w_end.day}',
            'start': w_start,
            'end': w_end,
        })
    return weeks


def _source_label(sources):
    """출처 라벨: 원티드 우선 순서로 정렬 (v16과 동일 표기)"""
    ordered = [s for s in ['원티드스페이스', 'FLEX'] if s in sources]
    return ' / '.join(ordered) if ordered else '-'


# ===== 일자별 병합 데이터 빌드 =====
def build_daily_data(targets, dept_map, erp_switchers, df1, df2, df3, weeks):
    """각 인원 × 28일 일자별 데이터. df3(FLEX) 전환일 이후 우선."""
    all_daily = {}
    for name in targets:
        daily = []
        sub2 = get_person_rows(df2, name).copy()
        if len(sub2):
            sub2['From'] = pd.to_datetime(sub2['From'])

        erp_info = erp_switchers.get(name)
        switch_date = pd.Timestamp(erp_info['switch_date']) if erp_info else None
        erp_name = erp_info['erp_name'] if erp_info else None

        for week_idx, w in enumerate(weeks):
            wstart = w['start']
            for d_idx in range(7):
                cur_date = wstart + timedelta(days=d_idx)
                short = DAYS_SHORT[d_idx]
                use_erp = False

                if erp_name and switch_date and cur_date >= switch_date and df3 is not None:
                    erp_row = df3[(df3['이름'] == erp_name) & (df3['날짜'] == cur_date)]
                    if len(erp_row) and sum(to_min(v) for v in erp_row['총 근무']) > 0:
                        use_erp = True

                if use_erp:
                    erp_row = df3[(df3['이름'] == erp_name) & (df3['날짜'] == cur_date)]
                    work_min = sum(to_min(v) for v in erp_row['총 근무'])
                    night_min = sum(to_min(v) for v in erp_row['야간'])
                    source = 'FLEX'
                else:
                    week_row = sub2[sub2['From'] == wstart] if len(sub2) else pd.DataFrame()
                    if len(week_row):
                        work_min = max(0, to_min(week_row[f'{short}_근무시간'].iloc[0])
                                       if isinstance(week_row[f'{short}_근무시간'].iloc[0], str)
                                       else week_row[f'{short}_근무시간'].iloc[0])
                        nv = week_row[f'{short}_근무시간(22시후)'].iloc[0]
                        night_min = to_min(nv) if isinstance(nv, str) else (nv or 0)
                    else:
                        work_min = night_min = 0
                    source = '원티드스페이스' if (work_min or night_min) else '없음'

                daily.append({
                    'date': cur_date,
                    'date_str': cur_date.strftime('%Y-%m-%d'),
                    'weekday': DAYS_FULL[d_idx],
                    'week_idx': week_idx,
                    'week_name': w['name'],
                    'dept': dept_map.get(name, ''),
                    'work_min': int(work_min) if work_min else 0,
                    'night_min': int(night_min) if night_min else 0,
                    'source': source,
                })
        all_daily[name] = daily
    return all_daily


def compute_summary(targets, sojung_map, contract_map, all_daily, weeks):
    """근태정리 요약 행 리스트 반환 (웹 미리보기·구글시트 공유용).
    각 행에 'contract'(필터용)와 표시 컬럼 포함. No는 호출측에서 부여."""
    rows = []
    for name in targets:
        daily = all_daily.get(name, [])
        sj = sojung_map.get(name, 40)
        wh, nmin, exc, sources = _aggregate(daily, sj)
        row = {
            '이름': name,
            'contract': contract_map.get(name, '시급제'),
            '소정근무시간(주)': sj,
        }
        for i, w in enumerate(weeks):
            row[w['name']] = round(wh[i], 2)
        row['초과근무(시간)'] = round(exc, 2)
        row['야간근무(시간)'] = round(nmin / 60, 2)
        row['데이터 출처'] = _source_label(sources)
        rows.append(row)
    return rows


def _aggregate(daily, sojung):
    """주차별 시간, 야간 합계, 초과, 출처 집합 계산"""
    weekly_h = [0.0] * 4
    night_min = 0
    sources = set()
    for d in daily:
        weekly_h[d['week_idx']] += d['work_min'] / 60
        night_min += d['night_min']
        if d['source'] != '없음':
            sources.add(d['source'])
    # v16 수식과 동일하게 '반올림된 주차값' 기준으로 초과 계산 (반올림 순서 일치)
    excess = sum(max(0, round(h, 2) - sojung) for h in weekly_h)
    return weekly_h, night_min, excess, sources


# ===== 검증 12개 항목 =====
def validate_data(all_daily, sojung_map, contract_map, erp_switchers, targets, df3, weeks):
    issues, warnings_list = [], []

    # 1. 인원 수
    if len(targets) == 0:
        issues.append("대상 인원이 0명입니다.")

    # 2. 일자별RAW 행 수 = 인원수 × 28
    for name in targets:
        n = len(all_daily.get(name, []))
        if n != 28:
            issues.append(f"{name}: 일자 수가 28일이 아님 ({n}일)")

    # 3. 근태정리 ↔ RAW 합계 일치 (집계 자체가 RAW 기반이므로 항상 일치 — 표식)
    # (수식이 RAW를 직접 참조하므로 구조적으로 보장됨)

    # 4. ERP 전환자 출처 분기 정상 (전환일 이후 FLEX 사용 여부)
    for name, info in erp_switchers.items():
        if name not in all_daily:
            continue
        sd = pd.Timestamp(info['switch_date'])
        after = [d for d in all_daily[name] if d['date'] >= sd]
        if after and not any(d['source'] == 'FLEX' for d in after):
            warnings_list.append(f"ℹ {name}: 전환일({info['switch_date']}) 이후 FLEX 기록이 없음 (원티드로 처리됨)")

    # 5~7. 개별 전환자 분기는 4번에 포함

    # 8~9. ABSENCE/데이터 없음 인원 (모든 주 0h)
    for name in targets:
        if name not in all_daily:
            continue
        if all(d['work_min'] == 0 and d['night_min'] == 0 for d in all_daily[name]):
            warnings_list.append(f"⚠ {name}: 4주간 근무 기록이 전혀 없음 (명단 확인 필요)")

    # 10. 초과근무 주 12h 한도 위반
    for name in targets:
        if name not in all_daily:
            continue
        sj = sojung_map.get(name, 40)
        wh, _, _, _ = _aggregate(all_daily[name], sj)
        for w_idx, h in enumerate(wh):
            ex = max(0, h - sj)
            if ex > 12:
                warnings_list.append(f"⛔ {name} {weeks[w_idx]['name']}: 1주 12h 한도 초과 (초과 {ex:.2f}h)")

    # 11. ERP 같은 날 중복 행 합산 — 실제 사용하는 전환자만, 요약 1줄
    if df3 is not None and '이름' in df3.columns and '날짜' in df3.columns:
        used = {info['erp_name'] for info in erp_switchers.values()}
        sub = df3[df3['이름'].isin(used)]
        dup = sub.groupby(['이름', '날짜']).size()
        dup_cnt = int((dup > 1).sum())
        if dup_cnt:
            warnings_list.append(f"ℹ ERP 동일 날짜 중복행 {dup_cnt}건 자동 합산 처리됨")

    # 12. 야간 검증 대상(우이솔) 존재 여부
    if NIGHT_PERSON in targets:
        _, nmin, _, _ = _aggregate(all_daily[NIGHT_PERSON], sojung_map.get(NIGHT_PERSON, 40))
        if nmin == 0:
            warnings_list.append(f"ℹ {NIGHT_PERSON}: 야간근무 기록 0 (야간상세 시트 참고)")

    return issues, warnings_list


# ===== 스타일 헬퍼 =====
def _styles():
    return {
        'header_fill': PatternFill('solid', fgColor='B4C7E7'),
        'contract_fill': PatternFill('solid', fgColor='D9E1F2'),
        'sojung_fill': PatternFill('solid', fgColor='E2EFDA'),
        'ot_night_fill': PatternFill('solid', fgColor='FFF2CC'),
        'raw_fill': PatternFill('solid', fgColor='D9D9D9'),
        'title_fill': PatternFill('solid', fgColor='305496'),
        'count_head_fill': PatternFill('solid', fgColor='4472C4'),
        'count_val_fill': PatternFill('solid', fgColor='FFF2CC'),
        'header_font': Font(name='맑은 고딕', size=11, bold=True),
        'data_font': Font(name='맑은 고딕', size=11),
        'title_font': Font(name='맑은 고딕', size=11, bold=True, color='FFFFFF'),
        'center': Alignment(horizontal='center', vertical='center', wrap_text=True),
        'left': Alignment(horizontal='left', vertical='center'),
        'border': Border(*(Side(border_style='thin', color='808080'),) * 4),
    }


# ===== 시트 1: 근태정리 =====
def _build_main(wb, targets, sojung_map, contract_map, all_daily, weeks, S):
    ws = wb.active
    ws.title = '근태정리'

    headers = ['No', '이름', '계약종류', '소정근무시간(주)',
               weeks[0]['name'], weeks[1]['name'], weeks[2]['name'], weeks[3]['name'],
               '초과근무(시간)', '야간근무(시간)', '데이터 출처',
               '연봉제 인원', '시급제 인원', '사업소득제 인원']
    for c, h in enumerate(headers, 1):
        ws.cell(1, c, h)
    # 2행: E~H 기간, L~N COUNTIF
    for i in range(4):
        ws.cell(2, 5 + i, f"({weeks[i]['range']})")
    ws.cell(2, 12, '=COUNTIF(C:C,"연봉제")')
    ws.cell(2, 13, '=COUNTIF(C:C,"시급제")')
    ws.cell(2, 14, '=COUNTIF(C:C,"사업소득제")')
    for c in (1, 2, 3, 4, 9, 10, 11):
        ws.merge_cells(start_row=1, start_column=c, end_row=2, end_column=c)

    for i, name in enumerate(targets, 1):
        r = i + 2
        ws.cell(r, 1, i)
        ws.cell(r, 2, name)
        ws.cell(r, 3, contract_map.get(name, '시급제'))
        ws.cell(r, 4, sojung_map.get(name, 40))
        for wi in range(4):
            col = get_column_letter(5 + wi)
            ws.cell(r, 5 + wi,
                    f'=ROUND(SUMIFS({RAW_SHEET}!F:F,{RAW_SHEET}!A:A,B{r},{RAW_SHEET}!C:C,"{weeks[wi]["name"]}")/60,2)')
        ws.cell(r, 9, f'=ROUND(MAX(0,E{r}-D{r})+MAX(0,F{r}-D{r})+MAX(0,G{r}-D{r})+MAX(0,H{r}-D{r}),2)')
        ws.cell(r, 10, f'=ROUND(SUMIFS({RAW_SHEET}!G:G,{RAW_SHEET}!A:A,B{r})/60,2)')
        _, _, _, sources = _aggregate(all_daily.get(name, []), sojung_map.get(name, 40))
        ws.cell(r, 11, _source_label(sources))

    last = 2 + len(targets)

    # 헤더 서식
    for c in range(1, 15):
        for r in (1, 2):
            cell = ws.cell(r, c)
            cell.font = S['header_font']; cell.alignment = S['center']; cell.border = S['border']
            if c in (9, 10):
                cell.fill = S['ot_night_fill']
            elif c == 4:
                cell.fill = S['sojung_fill']
            elif c == 3:
                cell.fill = S['contract_fill']
            elif c >= 12:
                cell.fill = S['count_head_fill']; cell.font = S['title_font']
            else:
                cell.fill = S['header_fill']
    # COUNTIF 값 서식
    for c in (12, 13, 14):
        cell = ws.cell(2, c)
        cell.font = Font(name='맑은 고딕', size=14, bold=True, color='C00000')
        cell.fill = S['count_val_fill']; cell.alignment = S['center']; cell.border = S['border']
        cell.number_format = '0"명"'

    # 데이터 서식
    for r in range(3, last + 1):
        for c in range(1, 12):
            cell = ws.cell(r, c)
            cell.font = S['data_font']; cell.border = S['border']
            cell.alignment = S['left'] if c == 2 else S['center']
            if c == 4:
                cell.number_format = '0'
            if c in (5, 6, 7, 8, 9, 10):
                cell.number_format = '0.00'

    widths = {1: 6, 2: 16, 3: 12, 4: 14, 5: 14, 6: 14, 7: 14, 8: 14,
              9: 13, 10: 13, 11: 20, 12: 14, 13: 14, 14: 14}
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 22
    ws.freeze_panes = 'C3'

    # 조건부서식 — 주차별(E~H): 소정 대비 비율
    for col in ('E', 'F', 'G', 'H'):
        rng = f'{col}3:{col}{last}'
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'AND({col}3>0, {col}3 > $D3 * 1.1)'],
            fill=PatternFill('solid', fgColor='FFC7CE'),
            font=Font(name='맑은 고딕', size=11, color='9C0006', bold=True)))
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'AND({col}3>0, {col}3 > $D3, {col}3 <= $D3 * 1.1)'],
            fill=PatternFill('solid', fgColor='FFEB9C')))
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'AND({col}3>0, {col}3 < $D3 * 0.5)'],
            fill=PatternFill('solid', fgColor='DDEBF7')))

    # 조건부서식 — 초과(I)·야간(J): (소정×4) 대비 비율
    for col in ('I', 'J'):
        rng = f'{col}3:{col}{last}'
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'AND($D3>0, {col}3 >= $D3 * 4 * 0.75)'],
            fill=PatternFill('solid', fgColor='C00000'),
            font=Font(name='맑은 고딕', size=11, bold=True, color='FFFFFF')))
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'AND($D3>0, {col}3 >= $D3 * 4 * 0.4)'],
            fill=PatternFill('solid', fgColor='FFC7CE')))
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f'AND($D3>0, {col}3 >= $D3 * 4 * 0.15)'],
            fill=PatternFill('solid', fgColor='FFEB9C')))


# ===== 시트 2: 일자별_병합RAW =====
def _build_raw(wb, targets, all_daily, S):
    ws = wb.create_sheet(RAW_SHEET)
    ws['A1'] = f'{len(targets)}명 × 28일 병합 RAW 데이터 (원티드 + ERP)'
    ws['A1'].font = S['title_font']; ws['A1'].fill = S['title_fill']
    ws['A2'] = '※ 데이터 출처: FLEX=ERP 전환자, 원티드스페이스=기본, 없음=근무없음'
    ws['A2'].font = Font(name='맑은 고딕', size=9, italic=True)

    head = ['이름', '부서', '주차', '날짜', '요일', '근무(분)', '야간(분)', '출처']
    for c, h in enumerate(head, 1):
        cell = ws.cell(4, c, h)
        cell.font = S['header_font']; cell.fill = S['raw_fill']
        cell.alignment = S['center']; cell.border = S['border']

    r = 5
    for name in targets:
        for d in all_daily.get(name, []):
            vals = [name, d['dept'], d['week_name'], d['date_str'], d['weekday'],
                    d['work_min'], d['night_min'], d['source']]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(r, c, v)
                cell.font = S['data_font']; cell.border = S['border']
                cell.alignment = S['left'] if c == 1 else S['center']
                if c in (6, 7):
                    cell.number_format = '#,##0'
            r += 1

    for c, w in zip(range(1, 9), [16, 14, 12, 13, 9, 12, 12, 18]):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = 'A5'


# ===== 시트 3: ERP_원본_10명 =====
def _build_erp(wb, erp_switchers, df3, weeks, S, source_file=''):
    ws = wb.create_sheet('ERP_원본_10명')
    n = len(erp_switchers)
    ws['A1'] = f'ERP 시스템 원본 데이터 — 전환자 {n}명만' + (f' (출처: {source_file})' if source_file else '')
    ws['A1'].font = S['title_font']; ws['A1'].fill = S['title_fill']

    head = ['이름(한글)', '이름(ERP)', '사번', '조직', '날짜', '요일',
            '출근시각', '퇴근시각', '총 근무', '야간', '연장']
    for c, h in enumerate(head, 1):
        cell = ws.cell(3, c, h)
        cell.font = S['header_font']; cell.fill = S['raw_fill']
        cell.alignment = S['center']; cell.border = S['border']

    r = 4
    if df3 is not None:
        p_start, p_end = weeks[0]['start'], weeks[3]['end']
        col_map = {  # df3 컬럼명 → 우리 컬럼 (있으면 사용)
            '사번': '사번', '조직': '조직', '요일': '요일',
            '출근시각': '출근시각', '퇴근시각': '퇴근시각',
            '총 근무': '총 근무', '야간': '야간', '연장': '연장',
        }
        for name, info in erp_switchers.items():
            erp_name = info['erp_name']
            sub = df3[(df3['이름'] == erp_name)].copy()
            if '날짜' in sub.columns:
                sub = sub[(sub['날짜'] >= p_start) & (sub['날짜'] <= p_end)]
                sub = sub.sort_values('날짜')
            for _, row in sub.iterrows():
                def g(col):
                    return row[col] if col in sub.columns and not pd.isna(row.get(col)) else ''
                dt = row.get('날짜')
                vals = [name, erp_name, g('사번'), g('조직'),
                        pd.Timestamp(dt).strftime('%Y-%m-%d') if not pd.isna(dt) else '',
                        g('요일'), g('출근시각'), g('퇴근시각'),
                        g('총 근무'), g('야간'), g('연장')]
                for c, v in enumerate(vals, 1):
                    cell = ws.cell(r, c, v)
                    cell.font = S['data_font']; cell.border = S['border']
                    cell.alignment = S['center']
                r += 1

    for c, w in zip(range(1, 12), [12, 16, 11, 14, 12, 6, 10, 10, 9, 8, 8]):
        ws.column_dimensions[get_column_letter(c)].width = w


# ===== 시트 4: 우이솔_야간상세 =====
def _build_night(wb, all_daily, weeks, S, person=NIGHT_PERSON):
    if person not in all_daily:
        return
    ws = wb.create_sheet('우이솔_야간상세')
    ws['A1'] = f'{person} 야간근무 상세 검증 (4주 28일)'
    ws['A1'].font = S['title_font']; ws['A1'].fill = S['title_fill']
    ws['A2'] = '※ 야간근무 = 22시 이후 근무 합계 (요일별 22시후 컬럼 합산)'
    ws['A2'].font = Font(name='맑은 고딕', size=9, italic=True)

    # 주간 요약 (R4 헤더, R5~R8 4주, R9 합계)
    sum_head = ['주차', '기간', '총 근무(h)', '22시 이후(h)', '야간 비율(%)', '비고']
    for c, h in enumerate(sum_head, 1):
        cell = ws.cell(4, c, h)
        cell.font = S['header_font']; cell.fill = S['header_fill']
        cell.alignment = S['center']; cell.border = S['border']
    for wi, w in enumerate(weeks):
        r = 5 + wi
        wk_daily = [d for d in all_daily[person] if d['week_idx'] == wi]
        work_h = round(sum(d['work_min'] for d in wk_daily) / 60, 2)
        night_h = round(sum(d['night_min'] for d in wk_daily) / 60, 2)
        ratio = round(night_h / work_h * 100, 1) if work_h else 0
        note = '⚠ 야간 비율 매우 높음' if ratio >= 50 else ('저녁/야간 패턴' if ratio >= 20 else '')
        for c, v in enumerate([w['name'], w['range'], work_h, night_h, ratio, note], 1):
            cell = ws.cell(r, c, v)
            cell.font = S['data_font']; cell.border = S['border']
            cell.alignment = S['left'] if c in (1, 2, 6) else S['center']
    ws.cell(9, 1, '합계').font = S['header_font']
    ws.cell(9, 3, '=SUM(C5:C8)')
    ws.cell(9, 4, '=SUM(D5:D8)')
    ws.cell(9, 5, '=ROUND(D9/C9*100,1)')
    ws.cell(9, 6, '4주 평균')
    for c in range(1, 7):
        ws.cell(9, c).border = S['border']; ws.cell(9, c).font = S['header_font']

    # 일자별 상세 (R11 제목, R12 헤더, R13~R40 28일, R41 합계)
    ws.cell(11, 1, '[28일 일자별 야간 상세]').font = S['header_font']
    dhead = ['주차', '날짜', '요일', '총 근무(분)', '총 근무(h)', '22시후(분)', '22시후(h)', '비율(%)', '비고']
    for c, h in enumerate(dhead, 1):
        cell = ws.cell(12, c, h)
        cell.font = S['header_font']; cell.fill = S['raw_fill']
        cell.alignment = S['center']; cell.border = S['border']
    for i, d in enumerate(all_daily[person]):
        r = 13 + i
        ws.cell(r, 1, d['week_name'])
        ws.cell(r, 2, d['date_str'])
        ws.cell(r, 3, d['weekday'])
        ws.cell(r, 4, f'=SUMIFS({RAW_SHEET}!F:F,{RAW_SHEET}!A:A,"{person}",{RAW_SHEET}!D:D,B{r})')
        ws.cell(r, 5, f'=ROUND(D{r}/60,2)')
        ws.cell(r, 6, f'=SUMIFS({RAW_SHEET}!G:G,{RAW_SHEET}!A:A,"{person}",{RAW_SHEET}!D:D,B{r})')
        ws.cell(r, 7, f'=ROUND(F{r}/60,2)')
        ws.cell(r, 8, f'=IFERROR(ROUND(F{r}/D{r}*100,1),0)')
        ws.cell(r, 9, '🌙 야간 발생' if d['night_min'] > 0 else ('근무 없음' if d['work_min'] == 0 else ''))
        for c in range(1, 10):
            cell = ws.cell(r, c)
            cell.font = S['data_font']; cell.border = S['border']
            cell.alignment = S['left'] if c in (1, 9) else S['center']
    tot = 13 + len(all_daily[person])
    ws.cell(tot, 1, '합계').font = S['header_font']
    ws.cell(tot, 4, f'=SUM(D13:D{tot - 1})')
    ws.cell(tot, 5, f'=SUM(E13:E{tot - 1})')
    ws.cell(tot, 6, f'=SUM(F13:F{tot - 1})')
    ws.cell(tot, 7, f'=SUM(G13:G{tot - 1})')
    for c in range(1, 10):
        ws.cell(tot, c).border = S['border']; ws.cell(tot, c).font = S['header_font']

    for c, w in zip(range(1, 10), [12, 13, 9, 13, 12, 12, 11, 10, 14]):
        ws.column_dimensions[get_column_letter(c)].width = w


# ===== 시트 5: 초과근로자_일자별RAW =====
def _build_overtime(wb, targets, sojung_map, dept_map, all_daily, S, exclude=NIGHT_PERSON):
    # 초과>0 인원을 초과량 내림차순 정렬
    ranked = []
    for name in targets:
        if name == exclude or name not in all_daily:
            continue
        sj = sojung_map.get(name, 40)
        _, _, excess, _ = _aggregate(all_daily[name], sj)
        if excess > 0:
            ranked.append((name, sj, round(excess, 2)))
    ranked.sort(key=lambda x: -x[2])

    ws = wb.create_sheet('초과근로자_일자별RAW')
    ws['A1'] = '초과근로 발생자 일자별 RAW 데이터 (병합 후) — 우이솔 제외'
    ws['A1'].font = S['title_font']; ws['A1'].fill = S['title_fill']

    dhead = ['주차', '날짜', '요일', '근무(분)', '근무(h)', '야간(분)', '야간(h)', '초과(h)', '출처', '비고']
    r = 3
    for idx, (name, sj, excess) in enumerate(ranked, 1):
        dept = dept_map.get(name, '')
        ws.cell(r, 1, f'{idx}. {name}  |  부서: {dept}  |  소정 {sj}h/주  |  4주 합계 초과: {excess}h')
        ws.cell(r, 1).font = Font(name='맑은 고딕', size=11, bold=True, color='C00000')
        r += 1
        for c, h in enumerate(dhead, 1):
            cell = ws.cell(r, c, h)
            cell.font = S['header_font']; cell.fill = S['raw_fill']
            cell.alignment = S['center']; cell.border = S['border']
        r += 1
        for d in all_daily[name]:
            ws.cell(r, 1, d['week_name'])
            ws.cell(r, 2, d['date_str'])
            ws.cell(r, 3, d['weekday'])
            ws.cell(r, 4, f'=SUMIFS({RAW_SHEET}!F:F,{RAW_SHEET}!A:A,"{name}",{RAW_SHEET}!D:D,B{r})')
            ws.cell(r, 5, f'=ROUND(D{r}/60,2)')
            ws.cell(r, 6, f'=SUMIFS({RAW_SHEET}!G:G,{RAW_SHEET}!A:A,"{name}",{RAW_SHEET}!D:D,B{r})')
            ws.cell(r, 7, f'=ROUND(F{r}/60,2)')
            ws.cell(r, 9, d['source'])
            for c in range(1, 11):
                cell = ws.cell(r, c)
                cell.font = S['data_font']; cell.border = S['border']
                cell.alignment = S['left'] if c in (1, 9, 10) else S['center']
            r += 1
        r += 1  # 블록 간 빈 줄

    for c, w in zip(range(1, 11), [12, 13, 9, 12, 11, 12, 11, 10, 16, 12]):
        ws.column_dimensions[get_column_letter(c)].width = w


# ===== 메인 진입점 =====
def build_excel(targets, sojung_map, erp_switchers, contract_map,
                df1, df2, df3, period_start, output_path,
                dept_map=None, erp_source_file=''):
    """v16 동일 구조(5시트 + 수식)로 엑셀 빌드.
    반환: dict(output_path, num_people, issues, warnings, weeks)
    """
    dept_map = dept_map or {}
    weeks = calc_weeks(period_start)
    all_daily = build_daily_data(targets, dept_map, erp_switchers, df1, df2, df3, weeks)
    issues, warnings_list = validate_data(
        all_daily, sojung_map, contract_map, erp_switchers, targets, df3, weeks)

    S = _styles()
    wb = Workbook()
    _build_main(wb, targets, sojung_map, contract_map, all_daily, weeks, S)
    _build_raw(wb, targets, all_daily, S)
    _build_erp(wb, erp_switchers, df3, weeks, S, source_file=erp_source_file)
    _build_night(wb, all_daily, weeks, S)
    _build_overtime(wb, targets, sojung_map, dept_map, all_daily, S)

    wb.save(output_path)
    summary = compute_summary(targets, sojung_map, contract_map, all_daily, weeks)
    return {
        'output_path': str(output_path),
        'num_people': len(targets),
        'issues': issues,
        'warnings': warnings_list,
        'weeks': [w['name'] for w in weeks],
        'week_ranges': [w['range'] for w in weeks],
        'summary': summary,
        'period_month': weeks[1]['start'].month,  # 대표 월 (탭 이름 기본값)
    }
