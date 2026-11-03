"""Run from repo root: pytest -q tests/test_api.py (install httpx & pytest)."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'backend'))
from fastapi.testclient import TestClient
from main import app, docx_stats

ROOT = Path(__file__).resolve().parents[1]
client = TestClient(app)

def test_health():
    response = client.get('/api/health')
    assert response.status_code == 200
    assert response.json()['status'] == 'ok'

def test_invalid_docx_is_rejected():
    response = client.post('/api/normalize', data={'template_id':'rabbani'},
        files={'file': ('invalid.docx', b'not-a-zip', 'application/octet-stream')})
    assert response.status_code == 422

def test_unknown_template_is_rejected():
    sample = (ROOT / 'templates/rabbani/word2010_base.docx').read_bytes()
    response = client.post('/api/normalize', data={'template_id':'unknown'},
        files={'file': ('intermediate.docx', sample, 'application/octet-stream')})
    assert response.status_code == 422

def test_libreoffice_roundtrip():
    sample = (ROOT / 'templates/rabbani/word2010_base.docx').read_bytes()
    response = client.post('/api/normalize', data={'template_id':'rabbani'},
        files={'file': ('intermediate.docx', sample, 'application/octet-stream')})
    assert response.status_code == 200, response.text
    assert response.content != sample
    assert 'wordprocessingml.document' in response.headers['content-type']
    assert 'footnotes' in docx_stats(response.content)
