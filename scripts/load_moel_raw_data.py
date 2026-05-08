from __future__ import annotations

import csv
import json
import re
import ssl
import time
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path


BASE = Path(__file__).resolve().parents[1]
RAW_DIR = BASE / "public" / "data" / "raw"
OUT_DIR = BASE / "public" / "data"
MANIFEST = RAW_DIR / "manifest.json"

DATASETS = [
    ("15068774", "연도별 최저임금", "예상임금 하한선 설정"),
    ("3069913", "(고용노동통계)직종별사업체노동력조사", "직종별 빈일자리와 부족인원 기준선"),
    ("15154122", "청년일자리도약장려금 사업추진실적", "장려금 기업 경로의 채용확률 보정"),
    ("15118732", "일학습병행 참여기업", "훈련과 취업 결합 경로 후보"),
    ("15033150", "한권으로 통하는 대한민국 청년지원 프로그램 정보", "청년에게 가능한 정책 경로 후보 생성"),
    ("3069922", "(고용노동통계)지역별사업체노동력조사", "지역 노동수요와 경기 변화 반영"),
    ("15118639", "국가기술자격 통계연보", "자격취득 경로의 효용 추정"),
    ("15068742", "구직급여 소정급여일수", "구직급여 가능 기간에 따른 압박도 계산"),
    ("15066368", "고용복지플러스센터 현황", "오프라인 전달체계와 기관 접근성 계산"),
    ("15029716", "구직급여 수급 현황", "실업 상태 기준선과 구직기간 제약 추정"),
    ("15106008", "국가직무능력표준(NCS) 개발 개선 현황", "희망직무와 보유역량 간 스킬갭 계산"),
    ("15117780", "직업능력개발 사업현황", "훈련사업 유형과 정책수단 확장"),
    ("3069910", "(고용노동통계)사업체노동력조사", "산업별 고용수요 보정"),
    ("15068741", "고용24구인구직 취업동향", "지역과 직무별 수요 및 바로 구직 경로의 취업확률 기준선"),
    ("3038234", "국민취업지원제도 운영기관", "상담기관 연결과 상담 경로 효과 산정"),
    ("15120704", "지역산업맞춤형 일자리지원사업 접수처", "지역 맞춤사업 경로 확장"),
    ("15139460", "국민내일배움카드 정보통신분야 실시현황", "IT와 데이터 직무 훈련 경로 보정"),
    ("15133009", "300명이상 사업체의 산업대분류별 임금수준", "예상 초임과 임금 경로 보정"),
    ("15119494", "청년일자리도약장려금 사업 운영기관", "장려금 가능 기업과 운영기관 연결"),
    ("15113450", "국민내일배움카드 발급 현황", "훈련 참여 가능성 보정"),
]


def fetch_bytes(url: str) -> bytes:
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0",
            "Accept": "*/*",
            "Referer": "https://www.data.go.kr/",
        },
    )
    context = ssl.create_default_context()
    with urllib.request.urlopen(req, timeout=60, context=context) as res:
        return res.read()


def fetch_text(url: str) -> str:
    return fetch_bytes(url).decode("utf-8", "replace")


def clean_filename(value: str) -> str:
    value = re.sub(r"[\\/:*?\"<>|]+", "_", value)
    value = re.sub(r"\s+", "_", value).strip("_")
    return value[:120]


def extract_download_urls(html: str) -> list[str]:
    urls = re.findall(r'"contentUrl"\s*:\s*"([^"]+)"', html)
    decoded = []
    for url in urls:
        url = url.replace("\\/", "/")
        if url not in decoded:
            decoded.append(url)
    return decoded


def extract_detail_pk(html: str) -> str:
    match = re.search(r'id="publicDataDetailPk"[^>]*value="([^"]+)"', html)
    return match.group(1) if match else ""


def detect_extension(url: str, content: bytes, fallback: str = "dat") -> str:
    parsed = urllib.parse.urlparse(url)
    path = urllib.parse.unquote(parsed.path)
    suffix = Path(path).suffix.lower().lstrip(".")
    if suffix in {"csv", "xlsx", "xls", "json", "xml", "zip"}:
        return suffix
    head = content[:200].lstrip()
    if head.startswith(b"PK\x03\x04"):
        return "xlsx"
    if head.startswith(b"{") or head.startswith(b"["):
        return "json"
    if head.startswith(b"<?xml") or head.startswith(b"<"):
        return "xml"
    return fallback


def count_csv_rows(path: Path) -> int | None:
    for enc in ("utf-8-sig", "cp949", "euc-kr"):
        try:
            with path.open("r", encoding=enc, newline="") as f:
                return max(0, sum(1 for _ in csv.reader(f)) - 1)
        except UnicodeDecodeError:
            continue
        except Exception:
            return None
    return None


def main() -> None:
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    rows = []
    for dataset_id, title, role in DATASETS:
        detail_url = f"https://www.data.go.kr/data/{dataset_id}/fileData.do"
        print(f"[{dataset_id}] 상세 페이지 확인: {title}")
        try:
            html = fetch_text(detail_url)
            download_urls = extract_download_urls(html)
            detail_pk = extract_detail_pk(html)
            if not download_urls:
                rows.append({
                    "id": dataset_id,
                    "title": title,
                    "role": role,
                    "detailUrl": detail_url,
                    "publicDataDetailPk": detail_pk,
                    "status": "no_content_url",
                    "files": [],
                })
                continue

            files = []
            for index, download_url in enumerate(download_urls, start=1):
                content = fetch_bytes(download_url)
                ext = detect_extension(download_url, content, "csv")
                file_name = f"{dataset_id}_{index}_{clean_filename(title)}.{ext}"
                out_path = RAW_DIR / file_name
                out_path.write_bytes(content)
                row_count = count_csv_rows(out_path) if ext == "csv" else None
                files.append({
                    "file": f"public/data/raw/{file_name}",
                    "downloadUrl": download_url,
                    "bytes": len(content),
                    "extension": ext,
                    "rowCount": row_count,
                })
                print(f"  저장: {file_name} ({len(content)} bytes)")
                time.sleep(0.25)

            rows.append({
                "id": dataset_id,
                "title": title,
                "role": role,
                "detailUrl": detail_url,
                "publicDataDetailPk": detail_pk,
                "status": "downloaded",
                "files": files,
            })
        except Exception as exc:
            rows.append({
                "id": dataset_id,
                "title": title,
                "role": role,
                "detailUrl": detail_url,
                "status": "error",
                "error": str(exc),
                "files": [],
            })
            print(f"  실패: {exc}")
        time.sleep(0.5)

    manifest = {
        "generatedAt": datetime.now().isoformat(timespec="seconds"),
        "source": "공공데이터포털 data.go.kr 파일데이터 상세 페이지의 schema.org contentUrl",
        "datasetCount": len(rows),
        "downloadedCount": sum(1 for row in rows if row["status"] == "downloaded"),
        "datasets": rows,
    }
    MANIFEST.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\nmanifest: {MANIFEST}")
    print(f"downloaded: {manifest['downloadedCount']} / {manifest['datasetCount']}")


if __name__ == "__main__":
    main()
