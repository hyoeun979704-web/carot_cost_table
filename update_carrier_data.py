#!/usr/bin/env python3
"""
통신사 단가표 이미지 → 엑셀 자동 업데이트 스크립트

이미지에서 SKT/KT/LG 끝단가·부가서비스·정책 차감 데이터를 추출하여
엑셀 템플릿의 아래 시트들을 자동으로 업데이트합니다.

  ▸ 입력용           : SKT·KT·LG 공통지원금, MNP지원금 (원 단위)
  ▸ SK합산단가입력    : 끝단가, 유통지원금, 공통추가, 부가서비스 (MNP·기변)
  ▸ KT 합산단가입력   : 끝단가, 유통지원금, S26정책, 부가추가, 저가MNP (MNP·기변)
  ▸ LG 합산단가입력   : 끝단가, 유통지원금, 담당정책, 777연합, 부가추가 (MNP·기변)

사용법:
    python update_carrier_data.py <이미지1> [이미지2 ...] -t <템플릿.xlsx> [-o <출력.xlsx>]

예시:
    python update_carrier_data.py skt.png kt.jpg lg.png -t 통큰단가표.xlsx
    python update_carrier_data.py all_carriers.png -t template.xlsx -o updated.xlsx
"""

import json
import sys
import os
import re
import argparse
import shutil
from pathlib import Path

from google import genai
from PIL import Image
import openpyxl


# ─────────────────────────────────────────────
# 1. 이미지 분석 — Gemini Vision
# ─────────────────────────────────────────────

EXTRACT_PROMPT = """이 이미지는 통신사(SKT/KT/LG U+) 휴대폰 단가표입니다.
아래 JSON 형식으로 데이터를 빠짐없이 추출해 주세요.

반환 형식:
{
  "carriers": {
    "SKT": {
      "models": [
        {
          "name": "모델명 (이미지에 표기된 그대로)",

          // ── 번호이동(MNP) ──
          "mnp_gongsi": 63,        // 공시 끝단가 (만원). 없으면 null
          "mnp_seonyak": 53,       // 선약 끝단가 (만원). 없으면 null
          "mnp_yutong": 5,         // 유통지원금 20이상 (만원). 없으면 null
          "mnp_gongtong": null,    // 공통추가지원금 (만원). 없으면 null
          "mnp_buga": null,        // 부가서비스 차감 (만원). 없으면 null

          // ── 기기변경(기변) ──
          "gibyon_gongsi": 33,     // 공시 끝단가 (만원). 없으면 null
          "gibyon_seonyak": 23,    // 선약 끝단가 (만원). 없으면 null
          "gibyon_yutong": 5,      // 유통지원금 20이상 (만원). 없으면 null
          "gibyon_gongtong": null, // 공통추가지원금 (만원). 없으면 null
          "gibyon_buga": null,     // 부가서비스 차감 (만원). 없으면 null

          // ── 공시지원금 (입력용 시트) ──
          "gongsi_jiwon": null     // 공시지원금 원 단위 (예: 450000). 없으면 null
        }
      ]
    },
    "KT": {
      "models": [
        {
          "name": "모델명",

          // ── 번호이동(MNP) ──
          "mnp_gongsi": 41,          // 공시 끝단가 (만원)
          "mnp_yutong": 3,           // 유통지원금 20이상 (만원)
          "mnp_s26_yeayak": null,    // S26 예약접수 추가 (만원)
          "mnp_s26_daeeung": null,   // S26 대응정책 추가 (만원)
          "mnp_buga": 4,             // 부가추가 차감 (만원)
          "mnp_jeokga_mnp": null,    // 저가MNP 구두지원 (만원)

          // ── 기기변경(기변) ──
          "gibyon_gongsi": 48,       // 공시 끝단가 (만원)
          "gibyon_yutong": 3,        // 유통지원금 20이상 (만원)
          "gibyon_s26_daeeung": null,// S26 대응정책 추가 (만원)
          "gibyon_buga": 4,          // 부가추가 차감 (만원)

          "gongsi_jiwon": null
        }
      ]
    },
    "LGU": {
      "models": [
        {
          "name": "모델명",

          // ── 번호이동(MNP) ──
          "mnp_gongsi": 53,          // 공시 끝단가 (만원)
          "mnp_yutong": 3,           // 유통망지원금 20이상 (만원)
          "mnp_13si": null,          // 13시(2시) 이후 추가 (만원)
          "mnp_damdang": null,       // 담당정책 추가 (만원)
          "mnp_777": 5,              // 777연합 추가 (만원)
          "mnp_buga": 6,             // 부가추가 차감 (만원)

          // ── 기기변경(기변) ──
          "gibyon_gongsi": 60,       // 공시 끝단가 (만원)
          "gibyon_yutong": 3,        // 유통망지원금 20이상 (만원)
          "gibyon_s26_early": null,  // S26 사전예약/얼리 추가 (만원)
          "gibyon_damdang": null,    // 담당정책 추가 (만원)
          "gibyon_777": 5,           // 777연합 추가 (만원)
          "gibyon_buga": 6,          // 부가추가 차감 (만원)

          "gongsi_jiwon": null
        }
      ]
    }
  }
}

주의사항:
- 이미지에 없는 통신사 키는 포함하지 마세요.
- 끝단가(고객 지불금, 만원)와 공시지원금(보조금, 원)을 구분하세요.
  · 끝단가: 고객이 내는 금액 (예: 63만원 → 63)
  · 공시지원금: 할인 보조금 (예: 450,000원 → 450000)
- 유통지원금·부가서비스·정책 금액은 모두 만원 단위입니다.
- 이미지에 값이 없거나 불명확한 필드는 반드시 null로 표기하세요.
- "신규" 컬럼은 무시하고 MNP/기변만 추출하세요.
- JSON 외의 텍스트를 절대 출력하지 마세요.

모델명: 이미지 표기 그대로 유지 (예: "갤럭시 S26", "아이폰17 에어 256G")
"""


def analyze_image(image_path: str) -> dict:
    print(f"  → 분석: {image_path}")
    client = genai.Client(api_key=os.environ.get("GOOGLE_API_KEY"))
    img = Image.open(image_path)
    response = client.models.generate_content(
        model="gemini-2.0-flash",
        contents=[img, EXTRACT_PROMPT],
    )
    text = response.text.strip()

    if text.startswith("```"):
        text = "\n".join(text.split("\n")[1:-1])

    try:
        return json.loads(text)
    except json.JSONDecodeError as e:
        print(f"    ⚠ JSON 파싱 오류: {e}\n    원본(처음 300자): {text[:300]}")
        return {"carriers": {}}


# ─────────────────────────────────────────────
# 2. 모델명 퍼지 매칭
# ─────────────────────────────────────────────

# 정규화: 공백·특수문자 제거, 소문자, 용량 표기 통일
_NORM_RE = re.compile(r"[\s\-_/·()（）]")
_CAP_RE = re.compile(r"(\d+)(g|gb)", re.IGNORECASE)


def normalize(name: str) -> str:
    s = _NORM_RE.sub("", name).lower()
    s = _CAP_RE.sub(lambda m: m.group(1) + "g", s)
    # 브랜드 별칭 통일
    s = s.replace("갤럭시", "galaxy")
    s = s.replace("아이폰", "iphone")
    s = s.replace("pro", "프로").replace("max", "맥스").replace("plus", "+")
    s = s.replace("ultra", "울트라").replace("air", "에어")
    s = s.replace("edge", "엣지").replace("flip", "플립").replace("fold", "폴드")
    return s


def best_match(img_name: str, excel_names) -> str | None:
    """img_name 과 가장 유사한 excel_names 항목 반환. 없으면 None."""
    nimg = normalize(img_name)
    # 정규화 결과를 한 번만 계산
    normed = [(en, normalize(en)) for en in excel_names]

    # 1) 완전 일치
    for en, nen in normed:
        if nen == nimg:
            return en

    # 2) 이미지명이 엑셀명의 부분집합인 경우 (용량 생략 등)
    for en, nen in normed:
        if nimg in nen or nen in nimg:
            return en

    # 3) 공통 토큰 기반 유사도
    img_tokens = set(re.findall(r"[가-힣a-z0-9]+", nimg))
    best_score, best_en = 0, None
    for en, nen in normed:
        en_tokens = set(re.findall(r"[가-힣a-z0-9]+", nen))
        if not en_tokens:
            continue
        score = len(img_tokens & en_tokens) / max(len(img_tokens), len(en_tokens))
        if score > best_score:
            best_score, best_en = score, en

    return best_en if best_score >= 0.5 else None


# ─────────────────────────────────────────────
# 3. 엑셀 구조 로드
# ─────────────────────────────────────────────

CARRIER_SHEET_MAP = {
    "SKT": "SK합산단가입력",
    "KT":  "KT 합산단가입력",
    "LGU": "LG 합산단가입력",
}

# 각 캐리어 시트의 업데이트 대상 열 인덱스 (0-based)
# 수식 셀은 제외: SKT E(4)/KT E(4)O(14)/LGU E(4)P(15) 등
CARRIER_COLS = {
    "SKT": {
        # 번호이동
        "mnp_gongsi":   3,   # D: MNP 공시 끝단가
        "mnp_seonyak":  4,   # E: MNP 선약 끝단가
        "mnp_yutong":   5,   # F: MNP 유통지원금 20이상
        "mnp_gongtong": 6,   # G: MNP 공통추가지원금
        "mnp_buga":     8,   # I: MNP 부가서비스
        # 기기변경
        "gibyon_gongsi":   12,  # M: 기변 공시 끝단가
        "gibyon_seonyak":  13,  # N: 기변 선약 끝단가
        "gibyon_yutong":   14,  # O: 기변 유통지원금 20이상
        "gibyon_gongtong": 15,  # P: 기변 공통추가지원금
        "gibyon_buga":     17,  # R: 기변 부가서비스
    },
    "KT": {
        # 번호이동 (E=D수식이라 D만 업데이트)
        "mnp_gongsi":      3,   # D: MNP 공시 끝단가
        "mnp_yutong":      5,   # F: MNP 유통지원금
        "mnp_s26_yeayak":  6,   # G: S26 예약접수 추가
        "mnp_s26_daeeung": 7,   # H: S26 대응정책 추가
        "mnp_buga":        8,   # I: MNP 부가추가
        "mnp_jeokga_mnp":  9,   # J: 저가MNP 구두지원
        # 기기변경 (O=N수식이라 N만 업데이트)
        "gibyon_gongsi":      13,  # N: 기변 공시 끝단가
        "gibyon_yutong":      15,  # P: 기변 유통지원금
        "gibyon_s26_daeeung": 17,  # R: 기변 S26 대응정책 추가
        "gibyon_buga":        18,  # S: 기변 부가추가
    },
    "LGU": {
        # 번호이동 (E=D-1수식이라 D만 업데이트)
        "mnp_gongsi":  3,   # D: MNP 공시 끝단가
        "mnp_yutong":  5,   # F: MNP 유통망지원금
        "mnp_13si":    7,   # H: 13시(2시) 이후 추가
        "mnp_damdang": 8,   # I: MNP 담당정책 추가
        "mnp_777":     9,   # J: MNP 777연합 추가
        "mnp_buga":    10,  # K: MNP 부가추가
        # 기기변경 (P=O-1수식이라 O만 업데이트)
        "gibyon_gongsi":    14,  # O: 기변 공시 끝단가
        "gibyon_yutong":    16,  # Q: 기변 유통망지원금
        "gibyon_s26_early": 17,  # R: 기변 S26 사전예약/얼리
        "gibyon_damdang":   18,  # S: 기변 담당정책 추가
        "gibyon_777":       19,  # T: 기변 777연합 추가
        "gibyon_buga":      20,  # U: 기변 부가추가
    },
}

# 입력용 시트 — 공시지원금·MNP추가지원금 열 (0-based)
IPRYONG_COLS = {
    "SKT": {"gongsi": 5, "mnp": 6},    # F, G
    "KT":  {"gongsi": 9, "mnp": 10},   # J, K
    "LGU": {"gongsi": 13, "mnp": 14},  # N, O
}


def load_excel_model_rows(wb: openpyxl.Workbook) -> dict:
    """
    각 시트에서 모델명 → 행 번호 매핑을 반환.
    반환 형식:
      {
        "입력용":          {"갤럭시 S26_256G": 16, ...},
        "SK합산단가입력":   {"갤럭시 S26_256G": 11, ...},
        "KT 합산단가입력":  {...},
        "LG 합산단가입력":  {...},
      }
    """
    def _scan_rows(ws, min_row, max_row, skip_vals):
        """C열(인덱스2)에서 모델명→행번호 매핑을 추출."""
        m = {}
        for row in ws.iter_rows(min_row=min_row, max_row=max_row):
            cell = row[2]
            val = cell.value
            if val and isinstance(val, str) and not val.startswith("=") \
               and val not in skip_vals:
                m[val] = cell.row
        return m

    result = {
        "입력용": _scan_rows(wb["입력용"], 13, 59, {"모델명"}),
    }
    for sheet_name in CARRIER_SHEET_MAP.values():
        if sheet_name in wb.sheetnames:
            result[sheet_name] = _scan_rows(
                wb[sheet_name], 8, 50, {"모델명", "끝단가", "액단가"}
            )
    return result


# ─────────────────────────────────────────────
# 4. 엑셀 업데이트
# ─────────────────────────────────────────────

def update_excel(wb: openpyxl.Workbook, all_extracted: list[dict],
                 model_rows: dict, log: list) -> int:
    """
    추출된 데이터로 엑셀을 업데이트하고 변경 셀 수를 반환.
    """
    updated = 0

    for extracted in all_extracted:
        carriers = extracted.get("carriers", {})

        for carrier_key, carrier_data in carriers.items():
            carrier_key = carrier_key.upper()
            # LG, LGU+, LGUP 등 변형 → LGU로 통일
            if carrier_key in ("LG", "LGU+", "LGUP", "LG U+", "LG U"):
                carrier_key = "LGU"

            models_data = carrier_data.get("models", [])
            sheet_name = CARRIER_SHEET_MAP.get(carrier_key)

            for item in models_data:
                img_name = item.get("name", "")
                if not img_name:
                    continue

                # ── A. 캐리어 합산단가 시트 업데이트 ──
                if sheet_name and sheet_name in wb.sheetnames:
                    sheet_models = model_rows.get(sheet_name, {})
                    matched = best_match(img_name, sheet_models)

                    if matched:
                        row_num = sheet_models[matched]
                        ws = wb[sheet_name]
                        col_map = CARRIER_COLS.get(carrier_key, {})

                        for field, col_idx in col_map.items():
                            val = item.get(field)
                            if val is not None:
                                cell = ws.cell(row=row_num, column=col_idx + 1)
                                old = cell.value
                                # 수식 셀은 건드리지 않음
                                if isinstance(old, str) and old.startswith("="):
                                    log.append(
                                        f"[{sheet_name}] 행{row_num} {matched} | "
                                        f"{field}: 수식 셀 스킵 ({old})"
                                    )
                                    continue
                                cell.value = val
                                log.append(
                                    f"[{sheet_name}] 행{row_num} {matched} | "
                                    f"{field}: {old} → {val}"
                                )
                                updated += 1
                    else:
                        log.append(f"[{sheet_name}] 매칭 실패: '{img_name}'")

                # ── B. 입력용 시트 공시지원금 업데이트 ──
                gongsi_jiwon = item.get("gongsi_jiwon")
                if gongsi_jiwon is not None:
                    ipyong_models = model_rows.get("입력용", {})
                    matched_ip = best_match(img_name, ipyong_models)

                    if matched_ip:
                        row_num = ipyong_models[matched_ip]
                        ws_ip = wb["입력용"]
                        col_map_ip = IPRYONG_COLS.get(carrier_key, {})
                        gongsi_col = col_map_ip.get("gongsi")

                        if gongsi_col is not None:
                            cell_ip = ws_ip.cell(row=row_num, column=gongsi_col + 1)
                            old = cell_ip.value
                            cell_ip.value = gongsi_jiwon
                            log.append(
                                f"[입력용] 행{row_num} {matched_ip} | "
                                f"{carrier_key} gongsi_jiwon: {old} → {gongsi_jiwon}"
                            )
                            updated += 1

    return updated


# ─────────────────────────────────────────────
# 5. 메인
# ─────────────────────────────────────────────

def parse_args():
    parser = argparse.ArgumentParser(
        description="통신사 단가표 이미지를 분석해 엑셀 템플릿을 업데이트합니다."
    )
    parser.add_argument("images", nargs="+",
                        help="분석할 이미지 파일 (JPG, PNG 등)")
    parser.add_argument("-t", "--template", required=True,
                        help="업데이트할 엑셀 템플릿 (.xlsx)")
    parser.add_argument("-o", "--output",
                        help="출력 파일명 (기본값: 템플릿 파일 덮어쓰기)")
    parser.add_argument("--dry-run", action="store_true",
                        help="실제 저장 없이 변경 내용만 출력")
    return parser.parse_args()


def main():
    args = parse_args()

    # 파일 검증
    for img in args.images:
        if not os.path.exists(img):
            print(f"❌ 이미지 파일 없음: {img}")
            sys.exit(1)
    if not os.path.exists(args.template):
        print(f"❌ 템플릿 파일 없음: {args.template}")
        sys.exit(1)

    api_key = os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        print("❌ GOOGLE_API_KEY 환경변수가 설정되지 않았습니다.")
        sys.exit(1)

    output_path = args.output or args.template

    # 템플릿 로드
    print(f"\n📂 템플릿 로드: {args.template}")
    wb = openpyxl.load_workbook(args.template)
    model_rows = load_excel_model_rows(wb)

    print(f"  ▸ 입력용 모델 수: {len(model_rows.get('입력용', {}))}")
    for carrier, sheet in CARRIER_SHEET_MAP.items():
        n = len(model_rows.get(sheet, {}))
        print(f"  ▸ {sheet} 모델 수: {n}")

    # 이미지 분석
    print(f"\n🔍 이미지 분석 ({len(args.images)}개)...")
    all_extracted = []
    for img_path in args.images:
        data = analyze_image(img_path)
        all_extracted.append(data)
        carriers_found = list(data.get("carriers", {}).keys())
        print(f"    ✓ {img_path} — 통신사 감지: {carriers_found}")

    # 엑셀 업데이트
    print("\n📝 엑셀 업데이트 중...")
    change_log = []
    total_updated = update_excel(wb, all_extracted, model_rows, change_log)

    # 변경 로그 출력
    if change_log:
        print(f"\n변경 내역 ({total_updated}셀):")
        for entry in change_log:
            marker = "  ✓" if "→" in entry else "  ⚠"
            print(f"{marker} {entry}")
    else:
        print("  ⚠ 업데이트할 데이터가 없습니다.")

    # 저장
    if args.dry_run:
        print("\n[dry-run] 파일을 저장하지 않습니다.")
    else:
        # 원본 백업
        if output_path == args.template:
            tp = Path(args.template)
            backup = str(tp.with_name(tp.stem + "_backup" + tp.suffix))
            shutil.copy2(args.template, backup)
            print(f"\n💾 원본 백업: {backup}")

        wb.save(output_path)
        print(f"✅ 저장 완료: {output_path}")
        print(f"   총 {total_updated}개 셀 업데이트")


if __name__ == "__main__":
    main()
