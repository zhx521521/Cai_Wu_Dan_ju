import os
from pathlib import Path

os.environ["FINANCE_DB"] = str(Path(__file__).parent / "test.db")

from fastapi.testclient import TestClient
from app.main import app


client = TestClient(app)


def test_health_and_sso():
    assert client.get("/api/health").json()["ok"] is True
    response = client.post("/api/auth/login", headers={"X-User": "reviewer@example.com"})
    assert response.status_code == 200
    assert response.json()["user"]["role"] == "finance_reviewer"


def test_applicant_cannot_read_another_role_write_surface():
    response = client.get("/api/configs", headers={"X-User": "applicant@example.com"})
    assert response.status_code == 403


def test_seeded_demo_and_rules():
    response = client.get("/api/documents", headers={"X-User": "reviewer@example.com"})
    assert response.status_code == 200
    docs = response.json()
    assert docs and docs[0]["type"] == "payment"
    doc_id = docs[0]["id"]
    checks = client.post(f"/api/documents/{doc_id}/run-checks", headers={"X-User": "reviewer@example.com"})
    assert checks.status_code == 200
    assert checks.json()["risk_level"] >= 1


def test_document_lifecycle_and_exports():
    applicant = {"X-User": "applicant@example.com"}
    reviewer = {"X-User": "reviewer@example.com"}
    admin = {"X-User": "admin@example.com"}
    created = client.post(
        "/api/documents",
        headers=applicant,
        json={
            "type": "payment",
            "title": "契约测试付款单",
            "department": "市场部",
            "payee": "测试供应商",
            "reason": "验收付款",
            "amount": 500000,
            "detail_total": 500000,
            "invoice_total": 500000,
            "contract_ratio": 0.4,
            "supplier": "测试供应商",
        },
    )
    assert created.status_code == 200
    document_id = created.json()["id"]
    uploaded = client.post(
        f"/api/documents/{document_id}/attachments",
        headers=applicant,
        files={"file": ("invoice.png", b"demo", "image/png")},
    )
    assert uploaded.status_code == 200
    assert client.post(f"/api/documents/{document_id}/actions", headers=applicant, json={"action": "submit"}).status_code == 409
    for field, value in (("contract_payment_ratio", "0.40"), ("payment_terms", "验收后付款")):
        assert client.post(
            f"/api/documents/{document_id}/evidence/{field}/correct",
            headers=applicant,
            json={"value": value, "reason": "契约测试人工校正"},
        ).status_code == 200
    assert client.post(f"/api/documents/{document_id}/actions", headers=applicant, json={"action": "submit"}).status_code == 200
    assert client.post(f"/api/documents/{document_id}/actions", headers=reviewer, json={"action": "start_review"}).status_code == 200
    assert client.post(f"/api/documents/{document_id}/actions", headers=reviewer, json={"action": "approve", "comment": "通过"}).status_code == 200
    for extension, content_type in (("csv", "text/csv"), ("xlsx", "spreadsheetml"), ("pdf", "application/pdf")):
        response = client.get(f"/api/reports/{document_id}/export.{extension}", headers=admin)
        assert response.status_code == 200
        assert content_type in response.headers["content-type"]
