"""
Conversor e Juntador de Word para PDF - Aplicação 100% Python Nativo
Utiliza exclusivamente bibliotecas nativas integradas ao Python:
- http.server, json, base64, tempfile, pathlib, subprocess, shutil, uuid, webbrowser
"""

import base64
import json
import mimetypes
import os
import shutil
import subprocess
import sys
import tempfile
import uuid
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
PUBLIC_DIR = BASE_DIR / "public"
TEMP_ROOT = Path(tempfile.gettempdir()) / "ConversorWordPDF_Python"
TEMP_ROOT.mkdir(parents=True, exist_ok=True)


def convert_with_word_batch(files_payload: list, merge_in_word: bool = False) -> dict:
    """
    Converte múltiplos arquivos Word (.docx, .doc, .rtf) em PDF na ordem escolhida
    utilizando uma única instância do Microsoft Word instalado no Windows,
    comandada pelo Python sem depender de pacotes externos (sem Node.js e sem pip).
    """
    job_dir = TEMP_ROOT / uuid.uuid4().hex
    job_dir.mkdir(parents=True, exist_ok=True)

    try:
        manifest = []
        for idx, item in enumerate(files_payload, start=1):
            orig_name = item.get("name", f"documento_{idx}.docx")
            ext = Path(orig_name).suffix.lower() or ".docx"
            raw_bytes = base64.b64decode(item["data"])

            input_path = job_dir / f"in_{idx:03d}{ext}"
            output_pdf_path = job_dir / f"out_{idx:03d}.pdf"
            input_path.write_bytes(raw_bytes)

            if ext == ".pdf":
                output_pdf_path.write_bytes(raw_bytes)

            manifest.append({
                "index": idx,
                "name": orig_name,
                "ext": ext,
                "input": str(input_path.resolve()),
                "output": str(output_pdf_path.resolve()),
            })

        manifest_path = job_dir / "manifest.json"
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False), encoding="utf-8")
        merged_word_pdf_path = job_dir / "merged_word.pdf"

        # Script de automação COM do Microsoft Word executado em lote único (rápido e silencioso)
        ps_script = job_dir / "word_worker.ps1"
        ps_code = f"""
$ErrorActionPreference = "Stop"
$manifest = Get-Content -Raw -Encoding UTF8 "{manifest_path.resolve()}" | ConvertFrom-Json
$mergeInWord = ${"true" if merge_in_word else "false"}

$word = $null
$mainDoc = $null
try {{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    if ($mergeInWord) {{
        $first = $manifest[0]
        $mainDoc = $word.Documents.Open($first.input, $false, $false, $false)
        $sel = $word.Selection
        for ($i = 1; $i -lt $manifest.Count; $i++) {{
            $sel.EndKey(6) | Out-Null
            $sel.InsertBreak(2) | Out-Null
            $sel.InsertFile($manifest[$i].input)
        }}
        $mainDoc.ExportAsFixedFormat("{merged_word_pdf_path.resolve()}", 17, $false, 0)
    }} else {{
        foreach ($item in $manifest) {{
            if ($item.ext -eq ".pdf") {{ continue }}
            $doc = $null
            try {{
                $doc = $word.Documents.Open($item.input, $false, $true, $false)
                $doc.ExportAsFixedFormat($item.output, 17, $false, 0)
            }} finally {{
                if ($null -ne $doc) {{
                    try {{ $doc.Close(0) }} catch {{}}
                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
                }}
            }}
        }}
    }}
}} finally {{
    if ($null -ne $mainDoc) {{
        try {{ $mainDoc.Close(0) }} catch {{}}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($mainDoc)
    }}
    if ($null -ne $word) {{
        try {{ $word.Quit(0) }} catch {{}}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}}
"""
        ps_script.write_text(ps_code, encoding="utf-8")

        print(f"[PYTHON] Convertendo {len(manifest)} arquivo(s) via Microsoft Word...")
        proc = subprocess.run(
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps_script)],
            capture_output=True,
            text=True,
        )
        if proc.returncode != 0:
            err_detail = (proc.stderr or proc.stdout or "Erro desconhecido no Microsoft Word").strip()
            raise RuntimeError(f"Falha ao converter no Microsoft Word: {err_detail}")

        if merge_in_word:
            if not merged_word_pdf_path.is_file():
                raise RuntimeError("O arquivo PDF unificado não pôde ser gerado pelo Word.")
            merged_b64 = base64.b64encode(merged_word_pdf_path.read_bytes()).decode("ascii")
            return {
                "ok": True,
                "mergedPdfBase64": merged_b64,
                "totalPages": 0,
            }

        converted_items = []
        for entry in manifest:
            out_file = Path(entry["output"])
            if not out_file.is_file():
                raise RuntimeError(f"Não foi possível gerar o PDF para '{entry['name']}'.")
            pdf_b64 = base64.b64encode(out_file.read_bytes()).decode("ascii")
            converted_items.append({
                "name": entry["name"],
                "pdfBase64": pdf_b64,
            })

        print(f"[PYTHON] Conversão concluída com sucesso ({len(converted_items)} arquivo(s)).")
        return {
            "ok": True,
            "files": converted_items,
            "totalPages": 0,
        }
    finally:
        shutil.rmtree(job_dir, ignore_errors=True)


class ConverterHTTPRequestHandler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        print(f"[PYTHON HTTP] {self.address_string()} - {fmt % args}")

    def _send_json(self, status_code: int, payload: dict):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = self.path.split("?", 1)[0]

        if path == "/api/status":
            self._send_json(200, {
                "ok": True,
                "service": "Conversor Word -> PDF (100% Python Nativo)",
                "python_version": sys.version.split()[0],
            })
            return

        if path == "/":
            path = "/index.html"

        rel_path = path.lstrip("/").replace("/", os.sep)
        file_path = (PUBLIC_DIR / rel_path).resolve()

        if not str(file_path).startswith(str(PUBLIC_DIR.resolve())) or not file_path.is_file():
            self._send_json(404, {"ok": False, "error": f"Arquivo não encontrado: {path}"})
            return

        content_type, _ = mimetypes.guess_type(str(file_path))
        if not content_type:
            content_type = "application/octet-stream"
        if content_type.startswith("text/") or content_type == "application/javascript":
            content_type += "; charset=utf-8"

        data = file_path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self):
        path = self.path.split("?", 1)[0]

        if path == "/api/convert":
            try:
                content_length = int(self.headers.get("Content-Length", "0"))
                raw_body = self.rfile.read(content_length)
                payload = json.loads(raw_body.decode("utf-8"))

                files = payload.get("files", [])
                merge_in_word = bool(payload.get("mergeInWord", False))

                if not files:
                    self._send_json(400, {"ok": False, "error": "Nenhum arquivo enviado para conversão."})
                    return

                result = convert_with_word_batch(files, merge_in_word=merge_in_word)
                self._send_json(200, result)
            except Exception as exc:
                print(f"[ERRO] {exc}")
                self._send_json(500, {"ok": False, "error": str(exc)})
            return

        self._send_json(404, {"ok": False, "error": f"Rota não encontrada: {path}"})


def main():
    port = 8080
    server = None
    for candidate_port in range(8080, 8095):
        try:
            server = ThreadingHTTPServer(("127.0.0.1", candidate_port), ConverterHTTPRequestHandler)
            port = candidate_port
            break
        except OSError:
            continue

    if server is None:
        print("[ERRO] Não foi possível iniciar o servidor Python nas portas 8080 a 8094.")
        sys.exit(1)

    url = f"http://localhost:{port}/"
    print("=" * 64)
    print("  CONVERSOR E JUNTADOR WORD -> PDF (100% PYTHON NATIVO)")
    print(f"  Python {sys.version.split()[0]} ativo em: {url}")
    print("  Mantenha esta janela aberta enquanto usa o conversor.")
    print("  Pressione Ctrl+C para encerrar.")
    print("=" * 64)

    if "--no-browser" not in sys.argv:
        webbrowser.open(url)

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nEncerrando servidor Python...")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
