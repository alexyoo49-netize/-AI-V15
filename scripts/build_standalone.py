from pathlib import Path
import json

BASE = Path(__file__).resolve().parents[1]
html = (BASE / "public" / "index.html").read_text(encoding="utf-8")
css = (BASE / "public" / "styles.css").read_text(encoding="utf-8")
js = (BASE / "public" / "app.js").read_text(encoding="utf-8")
data = json.loads((BASE / "public" / "data" / "datasets.json").read_text(encoding="utf-8"))

js = js.replace(
    '  const res = await fetch("/data/datasets.json");\n  dataStore = await res.json();',
    "  dataStore = window.__DATASETS__;",
)
js = js.replace(
    'async function extractPdf(file) {\n  const arrayBuffer = await file.arrayBuffer();\n  const base64 = btoa(String.fromCharCode(...new Uint8Array(arrayBuffer)));\n  const res = await fetch("/api/extract-pdf", {\n    method: "POST",\n    headers: { "content-type": "application/json" },\n    body: JSON.stringify({ base64 }),\n  });\n  const json = await res.json();\n  if (!res.ok) throw new Error(json.error || "PDF 추출 실패");\n  return json;\n}',
    'async function extractPdf() {\n  throw new Error("이 standalone 버전은 서버 없이 실행되어 PDF 자동 추출은 비활성화되어 있습니다. 이력서 내용을 텍스트로 붙여넣어 주세요.");\n}',
)

html = html.replace('    <link rel="stylesheet" href="/styles.css" />', f"    <style>\n{css}\n    </style>")
html = html.replace('    <script src="/app.js" type="module"></script>', f"    <script>window.__DATASETS__ = {json.dumps(data, ensure_ascii=False)};</script>\n    <script type=\"module\">\n{js}\n    </script>")
html = html.replace('PDF 이력서 업로드', 'PDF 이력서 업로드(서버 실행 시 사용)')

out = BASE / "내일경로AI_standalone.html"
out.write_text(html, encoding="utf-8")
print(out)
