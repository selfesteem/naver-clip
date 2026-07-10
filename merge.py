#!/usr/bin/env python3
"""
결과 파일 합산 도구 (메인 PC에서 실행)

사용법:
    python merge.py ./results/                   # 오늘 날짜 상태 확인 + 완료 시 자동 합산
    python merge.py ./results/ --date 20260708   # 날짜 지정
    python merge.py ./results/ --force           # 미완료 포함 강제 합산
"""

import argparse
import sys
from datetime import date
from pathlib import Path

import pandas as pd

SECTION_COLS = [
    "키워드", "처리완료", "네이버 클립", "블로그", "인플루언서", "뉴스",
    "카페", "지식iN", "동영상", "VIEW", "웹문서", "기타 노출", "오류",
]


def find_result_files(directory: Path, date_str: str) -> list[Path]:
    return sorted(directory.glob(f"result_{date_str}_*.xlsx"))


def parse_filename(path: Path) -> tuple[int, int]:
    """result_YYYYMMDD_{start}_{count}.xlsx → (start, count)"""
    parts = path.stem.split("_")
    return int(parts[2]), int(parts[3])


def check_and_merge(directory: Path, date_str: str, force: bool, output_dir: Path):
    files = find_result_files(directory, date_str)

    if not files:
        print(f"[{date_str}] 결과 파일이 없습니다. (result_{date_str}_*.xlsx)")
        return

    print(f"[{date_str}] 결과 파일 {len(files)}개\n")

    all_ready = True
    dfs: list[pd.DataFrame] = []
    total_keywords = 0

    for f in files:
        try:
            start, count = parse_filename(f)
        except (IndexError, ValueError):
            print(f"  {f.name:<50} ⚠  파일명 형식 불일치 (건너뜀)")
            continue

        df = pd.read_excel(f, dtype=str)
        done = int((df.get("처리완료", pd.Series(dtype=str)) == "Y").sum())
        total = len(df)
        total_keywords += total

        if done == total:
            status = "✓ 완료"
        else:
            status = f"⏳ 진행 중"
            all_ready = False

        print(f"  {f.name:<50} {status} ({done}/{total})")
        dfs.append(df)

    print()

    if not all_ready and not force:
        print("미완료 PC가 있습니다. 완료 후 다시 실행하거나 --force 옵션을 사용하세요.")
        return

    if not dfs:
        print("합산할 파일이 없습니다.")
        return

    merged = pd.concat(dfs, ignore_index=True)

    # 컬럼 순서 정렬
    ordered = [c for c in SECTION_COLS if c in merged.columns]
    extra = [c for c in merged.columns if c not in ordered]
    merged = merged[ordered + extra]

    # 중복 키워드 제거 (처리완료 Y 우선)
    merged["_sort"] = merged["처리완료"].apply(lambda v: 0 if str(v).strip() == "Y" else 1)
    merged = merged.sort_values("_sort").drop_duplicates(subset=["키워드"], keep="first")
    merged = merged.drop(columns=["_sort"]).reset_index(drop=True)

    output_path = output_dir / f"merged_{date_str}.xlsx"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        merged.to_excel(writer, index=False, sheet_name="결과")
        ws = writer.sheets["결과"]
        for col_cells in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col_cells), default=8)
            ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 40)

    total_done = int((merged.get("처리완료", pd.Series(dtype=str)) == "Y").sum())
    label = "완료" if all_ready else "일부 미완료(강제 합산)"
    print(f"{label}: {total_done}/{len(merged)}개")
    print(f"→ {output_path}")


def main():
    parser = argparse.ArgumentParser(description="결과 파일 합산")
    parser.add_argument("directory", nargs="?", default=".", help="결과 파일 폴더 (기본: 현재)")
    parser.add_argument("--date", help="날짜 지정 (YYYYMMDD, 기본: 오늘)")
    parser.add_argument("--force", action="store_true", help="미완료 포함 강제 합산")
    parser.add_argument("--output-dir", help="합산 파일 저장 위치 (기본: 같은 폴더)")
    args = parser.parse_args()

    directory = Path(args.directory)
    if not directory.exists():
        sys.exit(f"폴더를 찾을 수 없습니다: {directory}")

    date_str = args.date or date.today().strftime("%Y%m%d")
    output_dir = Path(args.output_dir) if args.output_dir else directory

    check_and_merge(directory, date_str, args.force, output_dir)


if __name__ == "__main__":
    main()
