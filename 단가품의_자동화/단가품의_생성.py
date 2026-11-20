# -*- coding: utf-8 -*-
"""
거래처별구매단가품의 분기별 자동 생성 프로그램
=====================================================================
입력 : ① 프로젝트구매입고품목조회(입고내역, 해당 분기)  ② 구매단가조회(등록단가)
출력 : ① {연도}년_{분기}분기_거래처별구매단가품의.xlsx  (2분기 양식과 동일)
       ② {연도}년_{분기}분기_품의_검토자료.xlsx        (제외/확인 대상 내역 - 검증용)

사용법
  1) 더블클릭(또는 python 단가품의_생성.py) → 파일 선택창에서 입고내역, 등록단가 선택
  2) 명령줄:
     python 단가품의_생성.py -r 입고내역.xlsx -m 등록단가.xlsx [-y 2026 -q 2] [-o 출력폴더]
        [--override 수기조정.xlsx] [--start-date 2026-10-01] [--compare 기존품의.xlsx] [--include-registered]

필요 패키지 : openpyxl  (pip install openpyxl)
"""
import argparse
import datetime as dt
import os
import re
import sys
from collections import defaultdict

import openpyxl
import openpyxl.worksheet.properties
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.colors import Color

# =====================================================================
# ★ 설정 (업무 규칙) - 규칙이 바뀌면 여기만 수정하세요
# =====================================================================
# (1) 제외 대상: PSC 90번대 자재 (품번이 P-9 로 시작 = PSC 시약/가공품/금형 등)
EXCLUDE_PART_PREFIXES = ("P-9",)
EXCLUDE_LARGE_CATEGORIES = ("PSC",)          # 품목대분류가 PSC 인 자재

# (2) 제외 대상: 수시로 단가가 바뀌는 비용성 품목 (품명/규격에 아래 단어가 있으면 제외)
#     ※ 2분기 실제 품의 파일 대조 결과: '수리', '운송비', '기업이윤', '일반관리비', '개조' 는 전량 제외,
#       2026-10 사용자 확인으로 '수정'(수정/가공 비용), '인건'(인건비) 도 제외 대상에 추가.
EXCLUDE_WORDS = ["수리", "운송비", "기업이윤", "일반관리비", "개조", "수정", "인건"]

# (3) 단가 비교 허용오차(원): USD 환산 소수점 단가를 정수로 등록한 경우의 반올림 차이 흡수
PRICE_TOLERANCE = 0.5

# (4) 등록단가 파일이 '품의 등록 이후' 상태일 때를 위한 판별:
#     적용시작일이 분기말 다음날 ~ +POST_WINDOW_DAYS 일 사이 && 1일자 인 건 = 이번 품의로 등록된 건
POST_WINDOW_DAYS = 62

# (5) 출력 고정값 (2분기 양식 기준)
OUT_ASSET_CLASS = "원재료"
OUT_CURRENCY = "KRW"
TITLE = "거래처별구매단가품의"
HEADERS = ["적용시작일", "구매거래처", "품목자산분류", "품번", "품명", "규격", "통화", "단가", "단위", "구분", "증감률"]
OUT_UNIT = "EA"
PCT_FMT = '"▲"0.0%;"▼"0.0%;0.0%'   # 올랐으면 ▲, 내렸으면 ▼
# =====================================================================


def nz(v):
    """None → '' , 문자열 앞뒤공백 제거(NBSP 포함)"""
    if v is None:
        return ""
    if isinstance(v, str):
        return v.replace(" ", " ").strip()
    return v


def num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").strip())
    except ValueError:
        return None


def read_table(path, key_header="품번"):
    """헤더행(품번 포함행)을 자동 탐색하여 [ {헤더:값} ] 반환. 원본 순서 유지."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = ws.iter_rows(values_only=True)
    header, out = None, []
    for i, row in enumerate(rows):
        if header is None:
            if row and key_header in [nz(c) for c in row]:
                header = [nz(c) for c in row]
            elif i > 10:
                raise ValueError(f"'{key_header}' 헤더를 찾지 못했습니다: {os.path.basename(path)}")
            continue
        if row is None or all(c in (None, "") for c in row):
            continue
        out.append({h: c for h, c in zip(header, row) if h})
    wb.close()
    if header is None:
        raise ValueError(f"'{key_header}' 헤더를 찾지 못했습니다: {os.path.basename(path)}")
    return header, out


def quarter_range(year, q):
    sm = 3 * (q - 1) + 1
    start = dt.date(year, sm, 1)
    if q == 4:
        end = dt.date(year, 12, 31)
    else:
        end = dt.date(year, sm + 3, 1) - dt.timedelta(days=1)
    return start, end


def to_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v or "").strip()[:10]
    try:
        return dt.date.fromisoformat(s)
    except ValueError:
        return None


def exclusion_reason(part, name, spec, large_cat):
    if any(part.startswith(p) for p in EXCLUDE_PART_PREFIXES) or large_cat in EXCLUDE_LARGE_CATEGORIES:
        return "PSC 90번대 자재"
    text = f"{name} {spec}"
    for w in EXCLUDE_WORDS:
        if w in text:
            return f"수시변동 비용({w})"
    return None


def ref_date(r):
    """단가가 정해진 시점 = 발주일(없으면 입고일). 입고단가는 발주단가에서 오기 때문"""
    return str(nz(r.get("발주일")) or nz(r.get("입고일")))[:10]


def build(receipts, master, year, quarter, overrides=None, prev_parts=None, include_registered=False):
    """
    prev_parts : (사용 중단) 예전 버전 호환용. 구분(신규/변경)은 '현재 등록단가의 품번' 기준으로만 판정합니다(거래처가 달라도 기존).
    include_registered : True 면 '이미 등록된 이번 분기 품의분'도 다시 포함(과거 품의 재생성/검증용)
    """
    q_start, q_end = quarter_range(year, quarter)

    # ---- 등록단가: 품번 기준(거래처 포함) 인덱스
    m_by_key, m_by_part = {}, defaultdict(list)
    for m in master:
        part, ven = nz(m.get("품번")), nz(m.get("구매거래처"))
        if not part:
            continue
        reg = str(nz(m.get("최종작업일시")))[:10]          # 이 단가가 등록(수정)된 날
        rec = {
            "vendor": ven, "price": num(m.get("단가")),
            "start": to_date(m.get("적용시작일")), "name": nz(m.get("품명")),
            "reg": reg if re.match(r"\d{4}-\d{2}-\d{2}$", reg) else "",
            "stopped": str(nz(m.get("중단여부"))).upper() == "TRUE",
        }
        m_by_key[(part, ven)] = rec
        m_by_part[part].append(rec)

    # ---- 입고내역: (품번, 거래처)별 전체 입고행
    rows_by_key, seq = defaultdict(list), 0
    n_all = 0
    out_of_q = 0
    for r in receipts:
        if nz(r.get("사업부문")) == "TOTAL":
            continue
        part, ven = nz(r.get("품번")), nz(r.get("구매거래처"))
        if not part:
            continue
        n_all += 1
        d = to_date(r.get("입고일"))
        if d is None or not (q_start <= d <= q_end):
            out_of_q += 1
        seq += 1
        # 최신 입고 = 입고일 → 입고번호 → (같은 입고번호 내 여러 단가 라인이면) 수량이 가장 큰 라인 → 원본순서
        order = (str(nz(r.get("입고일"))), str(nz(r.get("입고번호"))), num(r.get("금회입고수량")) or 0, seq)
        rows_by_key[(part, ven)].append((order, r))

    def pick(rows):
        return max(rows, key=lambda x: x[0])

    def unit_price(r):
        p = num(r.get("입고원화단가"))
        return p if p is not None else num(r.get("입고단가"))

    included, excluded, review = [], [], []
    for (part, ven), rows in rows_by_key.items():
        order, r = pick(rows)                              # 최신 입고(전체)
        name, spec = nz(r.get("품명")), nz(r.get("규격"))
        price = unit_price(r)
        large = nz(r.get("품목대분류"))
        m = m_by_key.get((part, ven))
        base = {"part": part, "vendor": ven, "name": name, "spec": spec, "L": price,
                "date": nz(r.get("입고일")), "large": large, "asset": nz(r.get("품목자산분류")),
                "order": order, "Rstart": m["start"] if m else None, "Rreg": m["reg"] if m else ""}

        why = exclusion_reason(part, name, spec, large)
        if why:
            excluded.append({**base, "reason": why})
            continue

        if m:
            R = m["price"]
            in_window = bool(m["start"] and q_end < m["start"] <= q_end + dt.timedelta(days=POST_WINDOW_DAYS)
                             and m["start"].day == 1)

            if include_registered:
                # [과거 품의 재생성 모드] 등록 이후 파일에서도 이번 품의분을 다시 만든다
                changed = R is None or abs(price - R) > PRICE_TOLERANCE
                if not (changed or in_window):
                    excluded.append({**base, "reason": "등록단가와 동일(변동없음)", "R": R})
                    continue
                final, tag = price, ""
                if in_window and changed:
                    final, tag = R, "수기조정 추정(등록단가 채택)"
                elif changed:
                    tag = "단가변경(등록단가 대비)"
                else:
                    tag = "이번 품의 등록분"
                item = {**base, "price": final, "R": R, "tag": tag}
                included.append(item)
                if tag.startswith("수기조정"):
                    review.append({**item, "note": f"입고 최신단가 {price:,.2f} ≠ 등록단가 {R:,.2f}"})
                continue

            # (A) 이번 분기 품의로 이미 등록된 건(적용시작일이 분기말 다음날~2개월 내 1일자) → 다시 올리지 않는다
            if in_window:
                excluded.append({**base, "reason": f"이번 분기 품의로 이미 등록됨(적용 {m['start']})", "R": R})
                if R is None or abs(price - R) > PRICE_TOLERANCE:
                    review.append({**base, "price": R, "R": R, "tag": "이미등록(입고단가 상이·제외됨)",
                                   "note": f"입고 최신단가 {price:,.2f} ≠ 이미 등록된 {fmt_num(R)} - 등록단가 우선, 변경 필요시 수기 추가"})
                continue

            # (B) 등록(수정)일 이전에 정해진 입고분은 '등록단가보다 옛 단가'이므로 변경 근거에서 제외
            valid = [x for x in rows if not m["reg"] or ref_date(x[1]) >= m["reg"]]
            if not valid:
                if R is not None and abs(price - R) <= PRICE_TOLERANCE:
                    excluded.append({**base, "reason": "등록단가와 동일(변동없음)", "R": R})
                else:
                    excluded.append({**base, "reason": "등록 이전 입고분 - 등록단가가 더 최신", "R": R})
                    review.append({**base, "price": R, "R": R, "tag": "등록 이전 입고분(제외됨)",
                                   "note": f"입고단가 {price:,.2f}는 등록({m['reg']}) 이전 발주/입고분, 등록단가 {fmt_num(R)} 우선 - 변경 필요시 수기 추가"})
                continue
            order, r = pick(valid)                         # 등록 이후 발주분 중 최신 입고
            price = unit_price(r)
            base.update({"L": price, "date": nz(r.get("입고일")), "order": order})

            # (C) 등록단가와 같으면 제외, 다르면 단가변경으로 포함
            if R is not None and abs(price - R) <= PRICE_TOLERANCE:
                excluded.append({**base, "reason": "등록단가와 동일(변동없음)", "R": R})
                continue
            included.append({**base, "price": price, "R": R, "tag": "단가변경(등록단가 대비)"})
        elif part not in m_by_part:
            # 등록단가에 없는 신규 품번 → 등록 필요
            item = {**base, "price": price, "R": None, "tag": "신규품번(미등록)"}
            included.append(item)
            if not price:
                review.append({**item, "note": "신규 품번인데 입고단가 0원 - 단가 확인 필요"})
        else:
            # 품번은 등록돼 있으나 이 거래처로는 미등록 → 거래처별 단가가 다를 수 있으므로 입고 거래처 단가로 신규 등록
            others = ", ".join(f"{x['vendor']}({x['price']:,.0f})" if x["price"] is not None else x["vendor"]
                               for x in m_by_part[part])
            reg_price = m_by_part[part][0]["price"]
            included.append({**base, "price": price, "R": None, "tag": "신규거래처(타 거래처 등록품번)",
                             "other": others, "Rother": reg_price})

    # ---- 수기조정 파일 적용(선택)
    if overrides:
        for it in included:
            ov = overrides.get((it["part"], it["vendor"])) or overrides.get((it["part"], ""))
            if ov is not None:
                it["price"], it["tag"] = ov, "수기조정(파일)"

    # ---- 신규/기존 구분 (현재 등록단가 기준)
    #  [품번 기준] 신규 = 현재 등록단가에 그 품번이 없음 / 변경(기존) = 품번이 이미 등록됨(거래처가 달라도 기존, ERP는 같은 품번을 타 거래처로 신규등록 불가)
    for it in included:
        it["kind"] = "신규" if it["part"] not in m_by_part else "변경"

    included.sort(key=lambda x: (x["part"], x["vendor"]))
    excluded.sort(key=lambda x: (x["reason"], x["part"], x["vendor"]))
    stats = {"n_receipt_rows": n_all, "out_of_quarter_rows": out_of_q, "n_keys": len(rows_by_key),
             "n_included": len(included), "n_excluded": len(excluded), "q_start": q_start, "q_end": q_end}
    return included, excluded, review, stats


def fmt_num(v):
    return "-" if v is None else f"{v:,.2f}"


# ------------------------------------------------------------------ 엑셀 출력
def _thin():
    return Side(style="thin")


def write_main(path, items, start_date=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    f_title = Font(name="맑은 고딕", size=15, bold=True, color="FF000000")
    f_head = Font(name="맑은 고딕", size=10, bold=True)
    f_body = Font(name="맑은 고딕", size=9)
    fill_head = PatternFill("solid", fgColor="FFE2EFDA")
    center = Alignment(horizontal="center", vertical="center")
    t = _thin()

    ws.merge_cells("A1:K1")
    ws["A1"] = TITLE
    ws["A1"].font, ws["A1"].alignment = f_title, center
    ws.row_dimensions[1].height = 27
    ws.row_dimensions[2].height = 22.15
    for i, h in enumerate(HEADERS, 1):
        c = ws.cell(2, i, h)
        c.font, c.fill = f_head, fill_head
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(left=t, right=t, top=t, bottom=Side(style="double"))
        if i == 8:
            c.number_format = '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-'

    NUMFMT = '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-'
    for n, it in enumerate(items):
        r = 3 + n
        ws.row_dimensions[r].height = 14.85
        old = it.get("R")
        pct = (it["price"] - old) / old if old and it["price"] is not None else None   # 신규(등록단가 없음)는 공란
        vals = [start_date, it["vendor"], OUT_ASSET_CLASS, it["part"], it["name"], it["spec"],
                OUT_CURRENCY, it["price"], OUT_UNIT, it.get("kind"), pct]
        for i, v in enumerate(vals, 1):
            c = ws.cell(r, i, v if v != "" else None)
            c.font = f_body
            c.border = Border(left=t, right=t, bottom=t, top=t if i in (3, 4) else None)
            if i == 1:
                c.number_format, c.alignment = "yy\\/mm\\/dd", center
            elif i in (3, 4):
                c.alignment = Alignment(horizontal="center", vertical="top", wrap_text=True)
            elif i == 8:
                c.number_format, c.alignment = NUMFMT, Alignment(vertical="center")
            elif i in (9, 10):
                c.alignment = center
            elif i == 11:
                c.number_format, c.alignment = PCT_FMT, center
            else:
                c.alignment = Alignment(vertical="center")
    for col, w in {"A": 13.5, "B": 23.66, "C": 16.66, "D": 15, "E": 35.83, "F": 13, "G": 6,
                   "H": 19.5, "I": 8, "J": 13.5, "K": 11}.items():
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:K{max(2 + len(items), 3)}"
    ws.page_setup.orientation = "landscape"
    ws.print_title_rows = "2:2"
    ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    wb.save(path)


def write_review(path, stats, included, excluded, review):
    wb = openpyxl.Workbook()
    hb = Font(bold=True)
    fill = PatternFill("solid", fgColor="FFE2EFDA")

    def sheet(title, headers, rows, first=False):
        ws = wb.active if first else wb.create_sheet()
        ws.title = title
        ws.append(headers)
        for c in ws[1]:
            c.font, c.fill = hb, fill
        for r in rows:
            ws.append(r)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for i, h in enumerate(headers, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = max(12, min(45, len(str(h)) * 2 + 4))
        return ws

    q = f"{stats['q_start']} ~ {stats['q_end']}"
    from collections import Counter
    ex_cnt = Counter(e["reason"] for e in excluded)
    in_cnt = Counter(i["tag"] for i in included)
    rows = [["대상 기간", q], ["입고내역 행수", stats["n_receipt_rows"]],
            ["분기 밖 입고행(확인)", stats["out_of_quarter_rows"]], ["품번+거래처 조합", stats["n_keys"]],
            ["품의 포함", stats["n_included"]], ["제외", stats["n_excluded"]], ["", ""], ["[포함 사유]", ""]]
    rows += [[k, v] for k, v in in_cnt.items()] + [["", ""], ["[제외 사유]", ""]] + [[k, v] for k, v in ex_cnt.items()]
    sheet("요약", ["항목", "값"], rows, first=True)
    sheet("확인필요", ["품번", "거래처", "품명", "구분", "입고 최신단가", "등록단가", "품의단가", "비고"],
          [[x["part"], x["vendor"], x["name"], x["tag"], x["L"], x.get("R"), x["price"], x["note"]] for x in review])
    nv = [x for x in included if x["tag"].startswith("신규거래처")]
    sheet("신규거래처(참고)", ["품번", "거래처(입고)", "품명", "규격", "신규등록 단가", "등록 거래처(단가)", "등록 거래처 단가", "차이율"],
          [[x["part"], x["vendor"], x["name"], x["spec"], x["price"], x["other"], x.get("Rother"),
            ((x["price"] - x["Rother"]) / x["Rother"]) if x.get("Rother") and x["price"] is not None else None] for x in nv])
    sheet("포함내역", ["품번", "거래처", "품명", "규격", "품의단가", "입고 최신단가", "등록단가", "포함사유", "최근입고일", "구분",
                     "등록 적용시작일", "등록 작업일"],
          [[x["part"], x["vendor"], x["name"], x["spec"], x["price"], x["L"], x.get("R"), x["tag"], x["date"], x.get("kind"),
            str(x["Rstart"] or ""), x.get("Rreg", "")] for x in included])
    sheet("제외내역", ["품번", "거래처", "품명", "규격", "입고 최신단가", "등록단가", "제외사유", "타거래처 등록정보", "최근입고일",
                     "등록 적용시작일", "등록 작업일"],
          [[x["part"], x["vendor"], x["name"], x["spec"], x["L"], x.get("R"), x["reason"], x.get("other", ""), x["date"],
            str(x.get("Rstart") or ""), x.get("Rreg", "")] for x in excluded])
    wb.save(path)


# ------------------------------------------------------------------ 검증(기존 품의 파일과 대조)
def compare(items, existing_path):
    _, rows = read_table(existing_path)
    ex = {(nz(r.get("품번")), nz(r.get("구매거래처"))): num(r.get("단가")) for r in rows if nz(r.get("품번"))}
    mine = {(i["part"], i["vendor"]): i["price"] for i in items}
    only_mine = sorted(set(mine) - set(ex))
    only_ex = sorted(set(ex) - set(mine))
    diff = [(k, mine[k], ex[k]) for k in mine if k in ex and abs((mine[k] or 0) - (ex[k] or 0)) > 0.5]
    return {"기존건수": len(ex), "생성건수": len(mine), "일치(품번+거래처)": len(set(mine) & set(ex)),
            "생성에만 있음": only_mine, "기존에만 있음": only_ex, "단가상이": diff}


# ------------------------------------------------------------------ 실행
def detect_year_quarter(receipts):
    ds = [to_date(r.get("입고일")) for r in receipts if nz(r.get("사업부문")) != "TOTAL"]
    ds = [d for d in ds if d]
    if not ds:
        raise ValueError("입고일을 읽을 수 없습니다.")
    from collections import Counter
    c = Counter((d.year, (d.month - 1) // 3 + 1) for d in ds)
    return c.most_common(1)[0][0]


def load_overrides(path):
    if not path:
        return None
    _, rows = read_table(path)
    out = {}
    for r in rows:
        p, v, pr = nz(r.get("품번")), nz(r.get("구매거래처")), num(r.get("단가"))
        if p and pr is not None:
            out[(p, v)] = pr
    return out


def run(receipts_path, master_path, year=None, quarter=None, outdir=None, override=None,
        start_date=None, compare_path=None, log=print, include_registered=False):
    _, receipts = read_table(receipts_path)
    _, master = read_table(master_path)
    dy, dq = detect_year_quarter(receipts)
    year, quarter = year or dy, quarter or dq
    if (year, quarter) != (dy, dq):
        log(f"[주의] 입고내역의 대부분은 {dy}년 {dq}분기인데 {year}년 {quarter}분기로 지정되었습니다.")
    included, excluded, review, stats = build(
        receipts, master, year, quarter, load_overrides(override), include_registered=include_registered)
    outdir = outdir or os.path.dirname(os.path.abspath(receipts_path))
    os.makedirs(outdir, exist_ok=True)
    main = os.path.join(outdir, f"{year}년_{quarter}분기_거래처별구매단가품의.xlsx")
    rev = os.path.join(outdir, f"{year}년_{quarter}분기_품의_검토자료.xlsx")
    sd = to_date(start_date) if start_date else None
    write_main(main, included, sd)
    write_review(rev, stats, included, excluded, review)
    log(f"{year}년 {quarter}분기 ({stats['q_start']} ~ {stats['q_end']})")
    log(f"  입고내역 {stats['n_receipt_rows']:,}행 → 품번+거래처 {stats['n_keys']:,}건")
    log(f"  품의 포함 {stats['n_included']:,}건 / 제외 {stats['n_excluded']:,}건 / 확인필요 {len(review)}건")
    if stats["out_of_quarter_rows"]:
        log(f"  [주의] 분기 밖 입고일 {stats['out_of_quarter_rows']}행 포함됨 - 입고내역 조회기간 확인")
    log(f"  → {main}\n  → {rev}")
    if compare_path:
        cmp_ = compare(included, compare_path)
        log("  [대조] " + ", ".join(f"{k}={v if not isinstance(v, list) else len(v)}" for k, v in cmp_.items()))
        return main, rev, cmp_
    return main, rev, None


def main():
    ap = argparse.ArgumentParser(description="거래처별구매단가품의 분기별 자동 생성")
    ap.add_argument("-r", "--receipts", help="프로젝트구매입고품목조회(입고내역) xlsx")
    ap.add_argument("-m", "--master", help="구매단가조회(등록단가) xlsx")
    ap.add_argument("-y", "--year", type=int)
    ap.add_argument("-q", "--quarter", type=int, choices=[1, 2, 3, 4])
    ap.add_argument("-o", "--outdir")
    ap.add_argument("--override", help="수기조정 단가 파일(품번, 구매거래처(선택), 단가)")
    ap.add_argument("--start-date", help="적용시작일 일괄 입력 (예: 2026-07-01), 미입력시 공란")
    ap.add_argument("--include-registered", action="store_true",
                    help="(검증용) 이미 등록된 이번 분기 품의분도 다시 포함 - 과거 품의 재생성 시에만 사용")
    ap.add_argument("--compare", help="기존에 수기로 만든 품의 파일과 대조(검증용)")
    a = ap.parse_args()

    if not (a.receipts and a.master):
        try:
            import tkinter as tk
            from tkinter import filedialog, messagebox, simpledialog
            root = tk.Tk()
            root.withdraw()
            a.receipts = filedialog.askopenfilename(title="① 입고내역(프로젝트구매입고품목조회) 선택",
                                                    filetypes=[("Excel", "*.xlsx")])
            a.master = filedialog.askopenfilename(title="② 등록단가(구매단가조회) 선택",
                                                  filetypes=[("Excel", "*.xlsx")])
            if not (a.receipts and a.master):
                return
            msgs = []
            try:
                run(a.receipts, a.master, a.year, a.quarter, a.outdir, log=msgs.append)
                messagebox.showinfo("완료", "\n".join(msgs))
            except Exception as e:  # noqa
                messagebox.showerror("오류", str(e))
            return
        except ImportError:
            ap.error("--receipts 와 --master 는 필수입니다.")
    run(a.receipts, a.master, a.year, a.quarter, a.outdir, a.override, a.start_date, a.compare,
        include_registered=a.include_registered)


if __name__ == "__main__":
    main()
