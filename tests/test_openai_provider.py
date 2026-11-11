"""End-to-end build against a local fake OpenAI-compatible server."""

import json
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer

import pytest
from docx import Document

from docforge.cli import main
from test_render import RESULT


class Handler(BaseHTTPRequestHandler):
    requests = []

    def log_message(self, *a):
        pass

    def do_POST(self):
        body = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
        Handler.requests.append(body)
        if body.get("response_format", {}).get("type") == "json_schema":
            # Like servers without structured-output support: docforge should step down.
            data = json.dumps({"error": {"message": "json_schema not supported"}}).encode()
            self.send_response(400)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        text = "<think>planning</think>```json\n" + json.dumps(RESULT) + "\n```"
        self.send_response(200)
        self.send_header("Content-Type", "text/event-stream")
        self.end_headers()
        half = len(text) // 2
        for piece, finish in ((text[:half], None), (text[half:], "stop")):
            chunk = {"id": "x", "object": "chat.completion.chunk", "created": 0, "model": body["model"],
                     "choices": [{"index": 0, "delta": {"content": piece}, "finish_reason": finish}]}
            self.wfile.write(f"data: {json.dumps(chunk)}\n\n".encode())
        self.wfile.write(b"data: [DONE]\n\n")


@pytest.fixture
def server():
    Handler.requests = []
    srv = HTTPServer(("127.0.0.1", 0), Handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    yield f"http://127.0.0.1:{srv.server_port}/v1"
    srv.shutdown()


def test_build_with_openai_compatible_api(tmp_path, server, monkeypatch):
    for var in ("DOCFORGE_PROVIDER", "DOCFORGE_MODEL", "DOCFORGE_BASE_URL", "DOCFORGE_API_KEY", "OPENAI_API_KEY"):
        monkeypatch.delenv(var, raising=False)
    monkeypatch.setenv("NO_PROXY", "127.0.0.1")
    assert main(["init", str(tmp_path), "--sample"]) == 0
    (tmp_path / "docforge.toml").write_text(
        f'provider = "openai"\nbase_url = "{server}"\nmodel = "local-model"\n', encoding="utf-8")
    out = tmp_path / "out.docx"
    assert main(["build", "-C", str(tmp_path), "-o", str(out)]) == 0

    formats = [r.get("response_format", {}).get("type") for r in Handler.requests]
    assert formats == ["json_schema", "json_object"]
    sent = Handler.requests[-1]
    assert sent["model"] == "local-model" and "max_tokens" not in sent
    user_text = " ".join(p["text"] for p in sent["messages"][1]["content"] if p["type"] == "text")
    assert "project Aurora" in user_text and 'section id="s1"' in user_text
    text = "\n".join(p.text for p in Document(str(out)).paragraphs)
    assert "Aurora - Status Report" in text and "Replace the old system." in text
