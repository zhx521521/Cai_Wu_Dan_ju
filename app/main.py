from __future__ import annotations

import csv
import hashlib
import io
import json
import os
import sqlite3
import uuid
from contextlib import asynccontextmanager
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from typing import Any, Optional

from fastapi import Depends, FastAPI, File, Header, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from pydantic import BaseModel, Field

try:
    from openpyxl import load_workbook
except Exception:  # Optional at import time; requirements install it in normal use.
    load_workbook = None

BASE_DIR = Path(__file__).resolve().parent.parent
DB_PATH = Path(os.getenv("FINANCE_DB", BASE_DIR / "finance.db"))
STATIC_DIR = BASE_DIR / "frontend"

DOC_TYPES = {
    "payment": "对公付款单",
    "prepayment": "预付款单",
    "batch_payment": "批量付款单",
    "expense": "费用报销单",
    "travel": "差旅报销单",
}
STATUSES = {
    "draft": "草稿",
    "parsing": "解析中",
    "applicant_confirm": "待申请人确认",
    "pending_review": "待初审",
    "in_review": "审批中",
    "returned": "退回待补充",
    "approved": "已通过",
    "rejected": "已驳回",
    "voided": "已作废",
    "withdrawn": "已撤回",
}
RISK_LEVELS = {0: "无风险", 1: "低风险", 2: "中风险", 3: "高风险"}
ROLE_NAMES = {
    "applicant": "单据申请人",
    "finance_reviewer": "财务初审员",
    "approver": "财务主管/审批人",
    "admin": "系统管理员",
    "auditor": "审计只读用户",
}


def now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def money(value: Any) -> Decimal:
    try:
        return Decimal(str(value or 0)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return Decimal("0.00")


def json_default(value: Any) -> Any:
    if isinstance(value, Decimal):
        return float(value)
    raise TypeError(f"Unsupported value: {type(value)!r}")


class DocumentIn(BaseModel):
    type: str
    title: str = Field(min_length=1, max_length=200)
    department: str = Field(default="市场部", min_length=1)
    payee: str = Field(default="", max_length=200)
    reason: str = Field(default="", max_length=2000)
    amount: float = Field(default=0, ge=0)
    detail_total: float = Field(default=0, ge=0)
    invoice_total: float = Field(default=0, ge=0)
    contract_ratio: float = Field(default=1, ge=0, le=1)
    expense_category: str = Field(default="技术服务")
    city_grade: str = Field(default="二线")
    supplier: str = Field(default="示例供应商")
    due_date: Optional[str] = None
    metadata: dict[str, Any] = Field(default_factory=dict)


class DocumentPatch(BaseModel):
    title: Optional[str] = None
    department: Optional[str] = None
    payee: Optional[str] = None
    reason: Optional[str] = None
    amount: Optional[float] = Field(default=None, ge=0)
    detail_total: Optional[float] = Field(default=None, ge=0)
    invoice_total: Optional[float] = Field(default=None, ge=0)
    contract_ratio: Optional[float] = Field(default=None, ge=0, le=1)
    metadata: Optional[dict[str, Any]] = None


class ActionIn(BaseModel):
    action: str
    comment: str = Field(default="", max_length=2000)
    assignee: Optional[str] = None


class CorrectionIn(BaseModel):
    value: str
    reason: str = Field(min_length=1, max_length=500)


class ConfigPatch(BaseModel):
    key: str
    value: str


def connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db() -> None:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = connect()
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS users (
          id TEXT PRIMARY KEY, email TEXT UNIQUE NOT NULL, name TEXT NOT NULL,
          role TEXT NOT NULL, department TEXT NOT NULL, active INTEGER NOT NULL DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS documents (
          id TEXT PRIMARY KEY, number TEXT UNIQUE NOT NULL, type TEXT NOT NULL,
          title TEXT NOT NULL, department TEXT NOT NULL, applicant_id TEXT NOT NULL,
          payee TEXT, reason TEXT, amount NUMERIC NOT NULL DEFAULT 0,
          detail_total NUMERIC NOT NULL DEFAULT 0, invoice_total NUMERIC NOT NULL DEFAULT 0,
          contract_ratio NUMERIC NOT NULL DEFAULT 1, expense_category TEXT,
          city_grade TEXT, supplier TEXT, due_date TEXT, status TEXT NOT NULL,
          version INTEGER NOT NULL DEFAULT 1, risk_level INTEGER NOT NULL DEFAULT 0,
          risk_summary TEXT NOT NULL DEFAULT '', metadata TEXT NOT NULL DEFAULT '{}',
          created_at TEXT NOT NULL, updated_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS attachments (
          id TEXT PRIMARY KEY, document_id TEXT NOT NULL, filename TEXT NOT NULL,
          content_type TEXT NOT NULL, size INTEGER NOT NULL, pages INTEGER NOT NULL DEFAULT 1,
          storage_path TEXT, parse_status TEXT NOT NULL DEFAULT 'uploaded',
          parse_error TEXT NOT NULL DEFAULT '', uploaded_at TEXT NOT NULL,
          FOREIGN KEY(document_id) REFERENCES documents(id)
        );
        CREATE TABLE IF NOT EXISTS evidence_fields (
          id TEXT PRIMARY KEY, document_id TEXT NOT NULL, attachment_id TEXT,
          field_name TEXT NOT NULL, normalized_value TEXT NOT NULL, raw_text TEXT NOT NULL,
          page INTEGER NOT NULL DEFAULT 1, x1 REAL NOT NULL DEFAULT 0, y1 REAL NOT NULL DEFAULT 0,
          x2 REAL NOT NULL DEFAULT 0, y2 REAL NOT NULL DEFAULT 0, confidence REAL NOT NULL DEFAULT 0,
          source TEXT NOT NULL DEFAULT 'adapter', corrected INTEGER NOT NULL DEFAULT 0,
          updated_at TEXT NOT NULL, UNIQUE(document_id, field_name),
          FOREIGN KEY(document_id) REFERENCES documents(id)
        );
        CREATE TABLE IF NOT EXISTS risk_items (
          id TEXT PRIMARY KEY, document_id TEXT NOT NULL, code TEXT NOT NULL,
          category TEXT NOT NULL, level INTEGER NOT NULL, title TEXT NOT NULL,
          detail TEXT NOT NULL, input_value TEXT NOT NULL, calculated_value TEXT NOT NULL,
          threshold TEXT NOT NULL, rule_version TEXT NOT NULL, evidence_ref TEXT NOT NULL,
          disposition TEXT NOT NULL DEFAULT 'pending', reviewer_comment TEXT NOT NULL DEFAULT '',
          created_at TEXT NOT NULL, updated_at TEXT NOT NULL,
          FOREIGN KEY(document_id) REFERENCES documents(id)
        );
        CREATE TABLE IF NOT EXISTS approvals (
          id TEXT PRIMARY KEY, document_id TEXT NOT NULL, node TEXT NOT NULL,
          actor_id TEXT NOT NULL, action TEXT NOT NULL, comment TEXT NOT NULL,
          assignee TEXT, created_at TEXT NOT NULL, idempotency_key TEXT UNIQUE,
          FOREIGN KEY(document_id) REFERENCES documents(id)
        );
        CREATE TABLE IF NOT EXISTS audit_events (
          id INTEGER PRIMARY KEY AUTOINCREMENT, actor_id TEXT NOT NULL, action TEXT NOT NULL,
          object_type TEXT NOT NULL, object_id TEXT NOT NULL, before_json TEXT,
          after_json TEXT, reason TEXT NOT NULL DEFAULT '', timestamp TEXT NOT NULL,
          rule_version TEXT, snapshot_hash TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS external_checks (
          id TEXT PRIMARY KEY, document_id TEXT NOT NULL, adapter TEXT NOT NULL,
          status TEXT NOT NULL, request_summary TEXT NOT NULL, response_json TEXT NOT NULL,
          service_version TEXT NOT NULL, queried_at TEXT NOT NULL, error TEXT NOT NULL DEFAULT '',
          FOREIGN KEY(document_id) REFERENCES documents(id)
        );
        CREATE TABLE IF NOT EXISTS configs (
          key TEXT PRIMARY KEY, value TEXT NOT NULL, updated_by TEXT NOT NULL, updated_at TEXT NOT NULL
        );
        """
    )
    defaults = {
        "invoice_confidence_threshold": "0.85",
        "market_price_tolerance": "0.15",
        "travel_city_grade_limit": "二线",
        "default_approval_flow": "finance_reviewer,approver",
    }
    for key, value in defaults.items():
        conn.execute("INSERT OR IGNORE INTO configs(key,value,updated_by,updated_at) VALUES (?,?,?,?)", (key, value, "system", now()))
    conn.commit()
    seed_demo(conn)
    conn.close()


def seed_demo(conn: sqlite3.Connection) -> None:
    users = [
        ("u-applicant", "applicant@example.com", "林申请", "applicant", "市场部"),
        ("u-reviewer", "reviewer@example.com", "周初审", "finance_reviewer", "财务部"),
        ("u-approver", "approver@example.com", "陈审批", "approver", "财务部"),
        ("u-admin", "admin@example.com", "系统管理员", "admin", "财务部"),
        ("u-auditor", "auditor@example.com", "审计只读", "auditor", "审计部"),
    ]
    conn.executemany("INSERT OR IGNORE INTO users(id,email,name,role,department) VALUES (?,?,?,?,?)", users)
    if conn.execute("SELECT 1 FROM documents LIMIT 1").fetchone():
        conn.commit()
        return
    demo = DocumentIn(
        type="payment", title="技术服务费对公付款", department="市场部", payee="云启科技有限公司",
        reason="年度技术服务付款", amount=500000, detail_total=500000, invoice_total=500000,
        contract_ratio=0.4, supplier="云启科技有限公司", metadata={"invoice_code": "031001900111", "invoice_number": "88990011", "invoice_date": "2026-09-01"}
    )
    create_document(conn, demo, "u-applicant", status="pending_review")
    doc = conn.execute("SELECT id FROM documents LIMIT 1").fetchone()
    if doc:
        run_rules(conn, doc["id"])
        conn.execute("UPDATE documents SET status='pending_review' WHERE id=?", (doc["id"],))
    conn.commit()


def create_document(conn: sqlite3.Connection, payload: DocumentIn, applicant_id: str, status: str = "draft") -> str:
    if payload.type not in DOC_TYPES:
        raise HTTPException(400, "不支持的单据类型")
    doc_id = str(uuid.uuid4())
    number = f"DOC-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:6].upper()}"
    stamp = now()
    conn.execute(
        """INSERT INTO documents(id,number,type,title,department,applicant_id,payee,reason,amount,detail_total,invoice_total,
        contract_ratio,expense_category,city_grade,supplier,due_date,status,metadata,created_at,updated_at)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (doc_id, number, payload.type, payload.title, payload.department, applicant_id, payload.payee, payload.reason,
         payload.amount, payload.detail_total, payload.invoice_total, payload.contract_ratio, payload.expense_category,
         payload.city_grade, payload.supplier, payload.due_date, status, json.dumps(payload.metadata, ensure_ascii=False), stamp, stamp),
    )
    append_audit(conn, applicant_id, "create", "document", doc_id, None, document_dict(conn, doc_id), "创建单据")
    return doc_id


def document_dict(conn: sqlite3.Connection, doc_id: str) -> dict[str, Any]:
    row = conn.execute("SELECT * FROM documents WHERE id=?", (doc_id,)).fetchone()
    if not row:
        raise HTTPException(404, "单据不存在")
    data = dict(row)
    data["type_name"] = DOC_TYPES.get(data["type"], data["type"])
    data["status_name"] = STATUSES.get(data["status"], data["status"])
    data["risk_level_name"] = RISK_LEVELS.get(data["risk_level"], "未知")
    data["metadata"] = json.loads(data["metadata"] or "{}")
    data["attachments"] = [dict(x) for x in conn.execute("SELECT * FROM attachments WHERE document_id=? ORDER BY uploaded_at", (doc_id,))]
    data["evidence_fields"] = [dict(x) for x in conn.execute("SELECT * FROM evidence_fields WHERE document_id=? ORDER BY field_name", (doc_id,))]
    data["risks"] = [dict(x) for x in conn.execute("SELECT * FROM risk_items WHERE document_id=? ORDER BY level DESC,created_at", (doc_id,))]
    data["approvals"] = [dict(x) for x in conn.execute("SELECT * FROM approvals WHERE document_id=? ORDER BY created_at", (doc_id,))]
    data["external_checks"] = [dict(x) for x in conn.execute("SELECT * FROM external_checks WHERE document_id=? ORDER BY queried_at DESC", (doc_id,))]
    return data


def append_audit(conn: sqlite3.Connection, actor: str, action: str, object_type: str, object_id: str,
                 before: Any, after: Any, reason: str = "", rule_version: str | None = None) -> None:
    payload = json.dumps({"before": before, "after": after}, ensure_ascii=False, default=json_default, sort_keys=True)
    digest = hashlib.sha256(payload.encode("utf-8")).hexdigest()
    conn.execute(
        """INSERT INTO audit_events(actor_id,action,object_type,object_id,before_json,after_json,reason,timestamp,rule_version,snapshot_hash)
        VALUES (?,?,?,?,?,?,?,?,?,?)""",
        (actor, action, object_type, object_id, json.dumps(before, ensure_ascii=False, default=json_default) if before is not None else None,
         json.dumps(after, ensure_ascii=False, default=json_default) if after is not None else None, reason, now(), rule_version, digest),
    )


def get_user(conn: sqlite3.Connection, identifier: str) -> sqlite3.Row | None:
    return conn.execute("SELECT * FROM users WHERE (id=? OR email=?) AND active=1", (identifier, identifier)).fetchone()


def current_user(x_user: str | None = Header(default=None, alias="X-User"), sso_user: str | None = Query(default=None, alias="sso_user")) -> sqlite3.Row:
    conn = connect()
    identifier = x_user or sso_user or "applicant@example.com"
    try:
        user = get_user(conn, identifier)
    except sqlite3.OperationalError as exc:
        conn.close()
        if "no such table" not in str(exc).lower():
            raise
        init_db()
        conn = connect()
        user = get_user(conn, identifier)
    conn.close()
    if not user:
        raise HTTPException(401, "SSO 用户未映射或已禁用")
    return user


def require_roles(*roles: str):
    def dependency(user: sqlite3.Row = Depends(current_user)) -> sqlite3.Row:
        if user["role"] not in roles:
            raise HTTPException(403, f"角色 {ROLE_NAMES.get(user['role'], user['role'])} 无权执行此操作")
        return user
    return dependency


def own_or_privileged(conn: sqlite3.Connection, doc_id: str, user: sqlite3.Row, write: bool = False) -> sqlite3.Row:
    row = conn.execute("SELECT * FROM documents WHERE id=?", (doc_id,)).fetchone()
    if not row:
        raise HTTPException(404, "单据不存在")
    if user["role"] in {"admin", "finance_reviewer", "approver", "auditor"}:
        if write and user["role"] == "auditor":
            raise HTTPException(403, "审计只读用户不可写入")
        return row
    if row["applicant_id"] != user["id"]:
        raise HTTPException(403, "只能访问自己创建的单据")
    return row


def config_value(conn: sqlite3.Connection, key: str, default: str) -> str:
    row = conn.execute("SELECT value FROM configs WHERE key=?", (key,)).fetchone()
    return row["value"] if row else default


def add_risk(conn: sqlite3.Connection, doc_id: str, code: str, category: str, level: int, title: str, detail: str,
             input_value: Any, calculated_value: Any, threshold: Any, evidence_ref: str = "") -> None:
    stamp = now()
    conn.execute(
        """INSERT INTO risk_items(id,document_id,code,category,level,title,detail,input_value,calculated_value,threshold,rule_version,evidence_ref,created_at,updated_at)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (str(uuid.uuid4()), doc_id, code, category, level, title, detail, str(input_value), str(calculated_value), str(threshold), "rules-v1.0", evidence_ref, stamp, stamp),
    )


def run_rules(conn: sqlite3.Connection, doc_id: str) -> dict[str, Any]:
    doc = conn.execute("SELECT * FROM documents WHERE id=?", (doc_id,)).fetchone()
    if not doc:
        raise HTTPException(404, "单据不存在")
    conn.execute("DELETE FROM risk_items WHERE document_id=?", (doc_id,))
    amount = money(doc["amount"])
    detail = money(doc["detail_total"])
    invoice = money(doc["invoice_total"])
    ratio = Decimal(str(doc["contract_ratio"] or 1))
    payable = (amount * ratio).quantize(Decimal("0.01"))
    tolerance = money(config_value(conn, "market_price_tolerance", "0.15"))
    delta = abs(amount - invoice)
    if delta > Decimal("0.01"):
        pct = (delta / amount * 100).quantize(Decimal("0.01")) if amount else Decimal("0")
        add_risk(conn, doc_id, "AMOUNT_INVOICE", "金额交叉核对", 2 if pct < 5 else 3, "单据与发票金额不一致", f"差额 {delta} 元，偏差 {pct}%", amount, invoice, "差额 <= 0.01", "invoice")
    if abs(amount - detail) > Decimal("0.01"):
        add_risk(conn, doc_id, "DETAIL_TOTAL", "金额交叉核对", 2, "明细合计与单据总额不一致", f"明细合计差额 {abs(amount-detail)} 元", detail, amount, "差额 <= 0.01", "document.details")
    if doc["type"] in {"payment", "prepayment"} and payable < amount:
        add_risk(conn, doc_id, "CONTRACT_RATIO", "合同付款", 3, "超前/超比例付款", f"本阶段应付 {payable} 元，申请 {amount} 元", amount, payable, f"合同付款比例 {ratio}", "contract.payment_ratio")
    if doc["type"] in {"expense", "travel"} and amount > Decimal("5000"):
        add_risk(conn, doc_id, "EXPENSE_LIMIT", "费用标准", 2, "费用超过标准", f"申请金额 {amount} 元超过 5000 元标准", amount, amount, "<= 5000", "expense.standard")
    metadata = json.loads(doc["metadata"] or "{}")
    invoice_key = (metadata.get("invoice_code"), metadata.get("invoice_number"), metadata.get("invoice_date"), str(invoice))
    if any(invoice_key[:3]):
        duplicates = conn.execute("SELECT id,number FROM documents WHERE id<>? AND metadata LIKE ?", (doc_id, f"%{invoice_key[1] or 'never'}%"))
        for duplicate in duplicates.fetchall():
            add_risk(conn, doc_id, "DUPLICATE_INVOICE", "重复报销", 3, "疑似重复发票", f"与单据 {duplicate['number']} 的发票号重复", invoice_key, duplicate["number"], "全局唯一", "invoice")
    if doc["type"] == "batch_payment":
        payees = metadata.get("payees", [])
        if len(payees) != len(set(payees)):
            add_risk(conn, doc_id, "BATCH_DUPLICATE_PAYEE", "批量付款", 2, "批次内收款人重复", "同一收款人出现多次", payees, len(set(payees)), "无重复", "batch.payees")
    risk_rows = conn.execute("SELECT * FROM risk_items WHERE document_id=?", (doc_id,)).fetchall()
    level = max((r["level"] for r in risk_rows), default=0)
    summary = "；".join(r["title"] for r in risk_rows) or "规则通过，未发现异常"
    conn.execute("UPDATE documents SET risk_level=?,risk_summary=?,updated_at=? WHERE id=?", (level, summary, now(), doc_id))
    return {"risk_level": level, "risk_level_name": RISK_LEVELS[level], "summary": summary, "items": [dict(r) for r in risk_rows]}


def parse_adapter(filename: str, content_type: str, size: int, doc: sqlite3.Row, pages: int = 1) -> dict[str, Any]:
    lower = filename.lower()
    if size > 20 * 1024 * 1024:
        return {"status": "non_retryable_failure", "error": "附件超过 20 MB"}
    if pages > 50:
        return {"status": "non_retryable_failure", "error": "附件超过 50 页"}
    if not lower.endswith((".pdf", ".png", ".jpg", ".jpeg")):
        return {"status": "non_retryable_failure", "error": "仅支持 PDF、PNG、JPG"}
    if "timeout" in lower:
        return {"status": "retryable_failure", "error": "OCR 服务超时，可重试"}
    if "nodata" in lower:
        return {"status": "no_data", "error": "外部 OCR 未返回可用数据"}
    metadata = json.loads(doc["metadata"] or "{}")
    fields = {
        "invoice_code": metadata.get("invoice_code", "031001900111"),
        "invoice_number": metadata.get("invoice_number", "88990011"),
        "tax_inclusive_total": str(doc["invoice_total"] or doc["amount"]),
        "tax_amount": str(money(doc["invoice_total"] or doc["amount"]) * Decimal("0.06")),
        "invoice_date": metadata.get("invoice_date", datetime.now().date().isoformat()),
        "buyer": metadata.get("buyer", "示例购买方"),
        "seller": doc["supplier"] or "示例供应商",
        "contract_payment_ratio": str(doc["contract_ratio"]),
        "payment_terms": metadata.get("payment_terms", "验收后付款"),
    }
    return {"status": "success", "service_version": "ocr-adapter-demo-1.0", "fields": fields, "requested_at": now()}


@asynccontextmanager
async def lifespan(_: FastAPI):
    init_db()
    yield


app = FastAPI(title="财务单据智能风险审核系统", version="1.0.0", lifespan=lifespan)
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=False, allow_methods=["*"], allow_headers=["*"])


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return (STATIC_DIR / "index.html").read_text(encoding="utf-8")


@app.get("/static/{asset:path}")
def static_asset(asset: str):
    target = STATIC_DIR / asset
    if not target.exists() or not target.is_file():
        raise HTTPException(404, "资源不存在")
    return FileResponse(target)


@app.get("/api/health")
def health() -> dict[str, Any]:
    return {"ok": True, "service": "financial-risk-review", "version": app.version}


@app.post("/api/auth/login")
def login(user: sqlite3.Row = Depends(current_user)) -> dict[str, Any]:
    conn = connect()
    append_audit(conn, user["id"], "login", "user", user["id"], None, {"email": user["email"], "role": user["role"]})
    conn.commit(); conn.close()
    return {"user": dict(user), "message": "SSO 映射成功"}


@app.get("/api/auth/me")
def me(user: sqlite3.Row = Depends(current_user)) -> dict[str, Any]:
    return {"id": user["id"], "email": user["email"], "name": user["name"], "role": user["role"], "role_name": ROLE_NAMES[user["role"]], "department": user["department"]}


@app.get("/api/users")
def users(user: sqlite3.Row = Depends(require_roles("admin"))) -> list[dict[str, Any]]:
    conn = connect(); rows = [dict(x) for x in conn.execute("SELECT id,email,name,role,department,active FROM users ORDER BY name")]; conn.close(); return rows


@app.get("/api/documents")
def list_documents(status: str | None = None, type: str | None = None, department: str | None = None,
                   risk_level: int | None = None, q: str | None = None, user: sqlite3.Row = Depends(current_user)) -> list[dict[str, Any]]:
    conn = connect(); clauses: list[str] = []; params: list[Any] = []
    if user["role"] == "applicant": clauses.append("applicant_id=?"); params.append(user["id"])
    if status: clauses.append("status=?"); params.append(status)
    if type: clauses.append("type=?"); params.append(type)
    if department: clauses.append("department=?"); params.append(department)
    if risk_level is not None: clauses.append("risk_level=?"); params.append(risk_level)
    if q: clauses.append("(number LIKE ? OR title LIKE ? OR payee LIKE ?)"); params.extend([f"%{q}%"] * 3)
    where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    rows = [document_summary(x) for x in conn.execute(f"SELECT * FROM documents {where} ORDER BY updated_at DESC", params)]
    conn.close(); return rows


def document_summary(row: sqlite3.Row) -> dict[str, Any]:
    data = dict(row); data["type_name"] = DOC_TYPES.get(data["type"], data["type"]); data["status_name"] = STATUSES.get(data["status"], data["status"]); data["risk_level_name"] = RISK_LEVELS.get(data["risk_level"], "未知"); return data


@app.post("/api/documents")
def create_doc(payload: DocumentIn, user: sqlite3.Row = Depends(require_roles("applicant", "admin"))) -> dict[str, Any]:
    conn = connect(); doc_id = create_document(conn, payload, user["id"]); conn.commit(); data = document_dict(conn, doc_id); conn.close(); return data


@app.get("/api/documents/{doc_id}")
def get_doc(doc_id: str, user: sqlite3.Row = Depends(current_user)) -> dict[str, Any]:
    conn = connect(); own_or_privileged(conn, doc_id, user); data = document_dict(conn, doc_id); conn.close(); return data


@app.patch("/api/documents/{doc_id}")
def patch_doc(doc_id: str, payload: DocumentPatch, user: sqlite3.Row = Depends(require_roles("applicant", "admin"))) -> dict[str, Any]:
    conn = connect(); old = own_or_privileged(conn, doc_id, user, write=True)
    if old["status"] not in {"draft", "returned", "withdrawn"}:
        raise HTTPException(409, "当前状态不可编辑")
    changes = payload.model_dump(exclude_none=True)
    if "metadata" in changes: changes["metadata"] = json.dumps(changes["metadata"], ensure_ascii=False)
    if not changes: return document_dict(conn, doc_id)
    fields = ",".join(f"{key}=?" for key in changes); values = list(changes.values()) + [now(), doc_id]
    conn.execute(f"UPDATE documents SET {fields},version=version+1,updated_at=? WHERE id=?", values)
    append_audit(conn, user["id"], "update", "document", doc_id, dict(old), document_dict(conn, doc_id), "编辑单据")
    conn.commit(); data = document_dict(conn, doc_id); conn.close(); return data


@app.post("/api/documents/{doc_id}/attachments")
async def upload_attachment(doc_id: str, file: UploadFile = File(...), user: sqlite3.Row = Depends(require_roles("applicant", "admin"))) -> dict[str, Any]:
    conn = connect(); doc = own_or_privileged(conn, doc_id, user, write=True); body = await file.read(); size = len(body)
    pages = 1
    if (file.filename or "").lower().endswith(".pdf"):
        try:
            from pypdf import PdfReader
            pages = len(PdfReader(io.BytesIO(body)).pages)
        except Exception:
            pages = 1
    parse = parse_adapter(file.filename or "attachment", file.content_type or "application/octet-stream", size, doc, pages)
    attachment_id = str(uuid.uuid4()); stamp = now(); storage = BASE_DIR / "uploads"; storage.mkdir(exist_ok=True); path = storage / f"{attachment_id}_{Path(file.filename or 'attachment').name}"; path.write_bytes(body)
    conn.execute("INSERT INTO attachments(id,document_id,filename,content_type,size,pages,storage_path,parse_status,parse_error,uploaded_at) VALUES (?,?,?,?,?,?,?,?,?,?)", (attachment_id, doc_id, file.filename or "attachment", file.content_type or "application/octet-stream", size, pages, str(path), parse["status"], parse.get("error", ""), stamp))
    if parse["status"] == "success":
        for name, value in parse["fields"].items():
            confidence = 0.91 if name not in {"contract_payment_ratio", "payment_terms"} else 0.76
            conn.execute("INSERT OR REPLACE INTO evidence_fields(id,document_id,attachment_id,field_name,normalized_value,raw_text,page,x1,y1,x2,y2,confidence,source,updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (str(uuid.uuid4()), doc_id, attachment_id, name, str(value), str(value), 1, 80, 100, 420, 130, confidence, "ocr-adapter-demo-1.0", stamp))
        conn.execute("UPDATE documents SET status='applicant_confirm',updated_at=? WHERE id=? AND status='draft'", (stamp, doc_id))
    append_audit(conn, user["id"], "upload_attachment", "attachment", attachment_id, None, {"document_id": doc_id, "filename": file.filename, "parse_status": parse["status"]}, "上传并调用 OCR 适配器")
    conn.commit(); result = {"attachment": dict(conn.execute("SELECT * FROM attachments WHERE id=?", (attachment_id,)).fetchone()), "parse": parse, "document": document_dict(conn, doc_id)}; conn.close(); return result


@app.post("/api/documents/{doc_id}/analyze")
def analyze_doc(doc_id: str, user: sqlite3.Row = Depends(require_roles("applicant", "admin", "finance_reviewer"))) -> dict[str, Any]:
    conn = connect(); doc = own_or_privileged(conn, doc_id, user, write=True); attachment = conn.execute("SELECT * FROM attachments WHERE document_id=? ORDER BY uploaded_at DESC LIMIT 1", (doc_id,)).fetchone()
    if not attachment: raise HTTPException(400, "请先上传凭证")
    parse = parse_adapter(attachment["filename"], attachment["content_type"], attachment["size"], doc, attachment["pages"])
    conn.execute("UPDATE attachments SET parse_status=?,parse_error=? WHERE id=?", (parse["status"], parse.get("error", ""), attachment["id"]))
    if parse["status"] == "success":
        for name, value in parse["fields"].items():
            conn.execute("UPDATE evidence_fields SET normalized_value=?,raw_text=?,updated_at=? WHERE document_id=? AND field_name=?", (str(value), str(value), now(), doc_id, name))
    append_audit(conn, user["id"], "analyze", "document", doc_id, None, parse, "重新执行凭证解析")
    conn.commit(); data = document_dict(conn, doc_id); conn.close(); return {"parse": parse, "document": data}


@app.post("/api/documents/{doc_id}/evidence/{field_name}/correct")
def correct_evidence(doc_id: str, field_name: str, payload: CorrectionIn, user: sqlite3.Row = Depends(require_roles("applicant", "finance_reviewer", "admin"))) -> dict[str, Any]:
    conn = connect(); own_or_privileged(conn, doc_id, user, write=True); row = conn.execute("SELECT * FROM evidence_fields WHERE document_id=? AND field_name=?", (doc_id, field_name)).fetchone()
    if not row: raise HTTPException(404, "字段证据不存在")
    conn.execute("UPDATE evidence_fields SET normalized_value=?,source='human-correction',corrected=1,updated_at=? WHERE id=?", (payload.value, now(), row["id"]))
    append_audit(conn, user["id"], "correct_evidence", "evidence_field", row["id"], dict(row), {**dict(row), "normalized_value": payload.value, "corrected": 1}, payload.reason)
    conn.commit(); result = document_dict(conn, doc_id); conn.close(); return result


@app.post("/api/documents/{doc_id}/run-checks")
def checks(doc_id: str, user: sqlite3.Row = Depends(require_roles("applicant", "finance_reviewer", "approver", "admin"))) -> dict[str, Any]:
    conn = connect(); own_or_privileged(conn, doc_id, user, write=True); result = run_rules(conn, doc_id); append_audit(conn, user["id"], "run_rules", "document", doc_id, None, result, "执行规则引擎", "rules-v1.0"); conn.commit(); conn.close(); return result


@app.post("/api/documents/{doc_id}/external-checks")
def external_checks(doc_id: str, user: sqlite3.Row = Depends(require_roles("finance_reviewer", "approver", "admin"))) -> dict[str, Any]:
    conn = connect(); doc = own_or_privileged(conn, doc_id, user, write=True); metadata = json.loads(doc["metadata"] or "{}"); supplier = doc["supplier"] or ""
    if "unavailable" in supplier.lower(): credit = {"status": "retryable_failure", "error": "征信服务暂时不可用"}
    elif "nodata" in supplier.lower(): credit = {"status": "no_data", "error": "供应商无可用征信数据"}
    elif "timeout" in supplier.lower(): credit = {"status": "retryable_failure", "error": "服务超时"}
    else: credit = {"status": "success", "registry_status": "存续", "abnormal_operations": [], "dishonesty": False, "account_changed_recently": bool(metadata.get("account_changed")), "source": "configured-credit-adapter"}
    price = {"status": "no_data", "error": "无匹配市场基准"} if "无基准" in (doc["title"] or "") else {"status": "success", "low": 430000, "high": 560000, "samples": 8, "source": "configured-market-adapter", "valid_until": "2026-12-31"}
    for adapter, result in (("credit", credit), ("market_price", price)):
        conn.execute("INSERT INTO external_checks(id,document_id,adapter,status,request_summary,response_json,service_version,queried_at,error) VALUES (?,?,?,?,?,?,?,?,?)", (str(uuid.uuid4()), doc_id, adapter, result["status"], json.dumps({"supplier": supplier, "item": doc["title"]}, ensure_ascii=False), json.dumps(result, ensure_ascii=False), "adapter-demo-1.0", now(), result.get("error", "")))
    # Refresh deterministic rule findings first; external findings are layered on top and retained.
    risks = run_rules(conn, doc_id)
    conn.execute("DELETE FROM risk_items WHERE document_id=? AND code IN ('MARKET_PRICE','CREDIT_UNAVAILABLE','SUPPLIER_CREDIT')", (doc_id,))
    if price["status"] == "success" and not (price["low"] <= float(doc["amount"]) <= price["high"]): add_risk(conn, doc_id, "MARKET_PRICE", "市场公允价", 2, "市场价格偏离基准", f"申请金额不在 {price['low']} - {price['high']} 区间", doc["amount"], f"{price['low']}-{price['high']}", config_value(conn, "market_price_tolerance", "0.15"), "market-price")
    if credit["status"] != "success": add_risk(conn, doc_id, "CREDIT_UNAVAILABLE", "供应商征信", 2, "外部征信结果不可用", credit.get("error", "请人工核查"), supplier, credit["status"], "不得默认为无风险", "credit")
    elif credit.get("dishonesty") or credit.get("abnormal_operations"): add_risk(conn, doc_id, "SUPPLIER_CREDIT", "供应商征信", 3, "供应商存在信用风险", "工商异常或失信记录", supplier, json.dumps(credit, ensure_ascii=False), "无异常", "credit")
    risk_rows = conn.execute("SELECT * FROM risk_items WHERE document_id=?", (doc_id,)).fetchall(); overall = max((row["level"] for row in risk_rows), default=0); summary = "；".join(row["title"] for row in risk_rows) or "规则通过，未发现异常"; conn.execute("UPDATE documents SET risk_level=?,risk_summary=?,updated_at=? WHERE id=?", (overall, summary, now(), doc_id)); result = {"risk_level": overall, "risk_level_name": RISK_LEVELS[overall], "summary": summary, "items": [dict(row) for row in risk_rows]}; append_audit(conn, user["id"], "external_checks", "document", doc_id, None, {"credit": credit, "market_price": price}, "调用征信与市场适配器"); conn.commit(); data = document_dict(conn, doc_id); conn.close(); return {"credit": credit, "market_price": price, "risk": result, "document": data}


TRANSITIONS = {
    "submit": {"draft": "pending_review", "applicant_confirm": "pending_review", "returned": "pending_review", "withdrawn": "pending_review"},
    "withdraw": {"pending_review": "withdrawn"},
    "return": {"pending_review": "returned", "in_review": "returned"},
    "reject": {"pending_review": "rejected", "in_review": "rejected"},
    "approve": {"pending_review": "approved", "in_review": "approved"},
    "start_review": {"pending_review": "in_review"},
    "void": {"draft": "voided", "withdrawn": "voided", "rejected": "voided", "approved": "voided"},
}


@app.post("/api/documents/{doc_id}/actions")
def document_action(doc_id: str, payload: ActionIn, user: sqlite3.Row = Depends(current_user)) -> dict[str, Any]:
    conn = connect(); old = own_or_privileged(conn, doc_id, user, write=True); action = payload.action
    if action == "void" and user["role"] not in {"applicant", "admin"}: raise HTTPException(403, "只有申请人或管理员可作废")
    if action in {"submit", "withdraw"} and user["role"] not in {"applicant", "admin"}: raise HTTPException(403, "申请人操作权限不足")
    if action in {"approve", "reject", "return", "start_review"} and user["role"] not in {"finance_reviewer", "approver", "admin"}: raise HTTPException(403, "审批角色权限不足")
    if action in {"return", "reject", "void"} and not payload.comment.strip(): raise HTTPException(400, "退回、驳回、作废必须填写原因")
    target = TRANSITIONS.get(action, {}).get(old["status"])
    if not target: raise HTTPException(409, f"状态 {STATUSES.get(old['status'])} 不允许执行 {action}")
    if action == "submit":
        fields = conn.execute("SELECT field_name,confidence,corrected FROM evidence_fields WHERE document_id=?", (doc_id,)).fetchall()
        attachment = conn.execute("SELECT 1 FROM attachments WHERE document_id=? AND parse_status='success' LIMIT 1", (doc_id,)).fetchone()
        if not attachment:
            raise HTTPException(409, "提交前必须上传并成功解析至少一个凭证附件")
        threshold = float(config_value(conn, "invoice_confidence_threshold", "0.85"))
        low = [x["field_name"] for x in fields if x["confidence"] < threshold and not x["corrected"]]
        if low: raise HTTPException(409, f"关键字段置信度不足，请先人工校正: {', '.join(low)}")
        result = run_rules(conn, doc_id); target = "pending_review"
        if result["risk_level"] >= 3: target = "pending_review"
    if action in {"approve", "reject", "return"} and old["status"] not in {"pending_review", "in_review"}: raise HTTPException(409, "当前不在审批队列")
    stamp = now(); conn.execute("UPDATE documents SET status=?,version=version+1,updated_at=? WHERE id=?", (target, stamp, doc_id))
    key = f"{doc_id}:{old['version']}:{action}"
    try:
        conn.execute("INSERT INTO approvals(id,document_id,node,actor_id,action,comment,assignee,created_at,idempotency_key) VALUES (?,?,?,?,?,?,?,?,?)", (str(uuid.uuid4()), doc_id, old["status"], user["id"], action, payload.comment, payload.assignee, stamp, key))
    except sqlite3.IntegrityError:
        raise HTTPException(409, "重复提交已被幂等保护")
    append_audit(conn, user["id"], action, "document", doc_id, dict(old), document_dict(conn, doc_id), payload.comment)
    conn.commit(); data = document_dict(conn, doc_id); conn.close(); return data


@app.get("/api/dashboard")
def dashboard(user: sqlite3.Row = Depends(current_user)) -> dict[str, Any]:
    conn = connect(); scope = "" if user["role"] != "applicant" else "WHERE applicant_id=?"; params = [] if scope == "" else [user["id"]]
    rows = conn.execute(f"SELECT * FROM documents {scope}", params).fetchall(); pending = [x for x in rows if x["status"] in {"pending_review", "in_review", "returned"}]; high = [x for x in rows if x["risk_level"] >= 3]; amounts = sum(float(x["amount"]) for x in rows); blocked = sum(float(x["amount"]) for x in rows if x["risk_level"] >= 2)
    by_type = {DOC_TYPES[k]: sum(1 for x in rows if x["type"] == k) for k in DOC_TYPES}; audit_count = conn.execute("SELECT COUNT(*) c FROM audit_events").fetchone()["c"]; conn.close()
    return {"pending_count": len(pending), "high_risk_count": len(high), "total_amount": amounts, "blocked_amount": blocked, "documents": len(rows), "by_type": by_type, "recent_tasks": [document_summary(x) for x in pending[:8]], "audit_events": audit_count}


@app.get("/api/configs")
def configs(user: sqlite3.Row = Depends(require_roles("admin", "auditor", "finance_reviewer", "approver"))) -> list[dict[str, Any]]:
    conn = connect(); rows = [dict(x) for x in conn.execute("SELECT * FROM configs ORDER BY key")]; conn.close(); return rows


@app.patch("/api/configs")
def update_config(payload: ConfigPatch, user: sqlite3.Row = Depends(require_roles("admin"))) -> dict[str, Any]:
    conn = connect(); old = conn.execute("SELECT * FROM configs WHERE key=?", (payload.key,)).fetchone(); conn.execute("INSERT INTO configs(key,value,updated_by,updated_at) VALUES (?,?,?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value,updated_by=excluded.updated_by,updated_at=excluded.updated_at", (payload.key, payload.value, user["id"], now())); append_audit(conn, user["id"], "update_config", "config", payload.key, dict(old) if old else None, {"key": payload.key, "value": payload.value}, "更新规则配置"); conn.commit(); result = dict(conn.execute("SELECT * FROM configs WHERE key=?", (payload.key,)).fetchone()); conn.close(); return result


@app.get("/api/audit")
def audit(object_id: str | None = None, user: sqlite3.Row = Depends(require_roles("admin", "auditor", "finance_reviewer", "approver"))) -> list[dict[str, Any]]:
    conn = connect(); rows = [dict(x) for x in conn.execute("SELECT * FROM audit_events WHERE object_id=? ORDER BY id" if object_id else "SELECT * FROM audit_events ORDER BY id DESC LIMIT 500", (object_id,) if object_id else ())]; conn.close(); return rows


@app.get("/api/reports")
def reports(user: sqlite3.Row = Depends(require_roles("admin", "auditor", "finance_reviewer", "approver"))) -> list[dict[str, Any]]:
    conn = connect(); rows = [document_summary(x) for x in conn.execute("SELECT * FROM documents WHERE status IN ('approved','rejected','voided') ORDER BY updated_at DESC")]; conn.close(); return rows


@app.get("/api/reports/{doc_id}/export.csv")
def export_csv(doc_id: str, user: sqlite3.Row = Depends(require_roles("admin", "auditor", "finance_reviewer", "approver"))):
    conn = connect(); doc = own_or_privileged(conn, doc_id, user); risks = conn.execute("SELECT * FROM risk_items WHERE document_id=?", (doc_id,)).fetchall(); audits = conn.execute("SELECT * FROM audit_events WHERE object_id=? ORDER BY id", (doc_id,)).fetchall(); out = io.StringIO(); writer = csv.writer(out); writer.writerow(["单据编号", "类型", "标题", "申请人", "部门", "金额", "风险等级", "状态"]); writer.writerow([doc["number"], DOC_TYPES[doc["type"]], doc["title"], doc["applicant_id"], doc["department"], doc["amount"], RISK_LEVELS[doc["risk_level"]], STATUSES[doc["status"]]]); writer.writerow([]); writer.writerow(["风险代码", "风险项", "等级", "依据", "证据定位", "处置"]); [writer.writerow([r["code"], r["title"], RISK_LEVELS[r["level"]], r["detail"], r["evidence_ref"], r["disposition"]]) for r in risks]; writer.writerow([]); writer.writerow(["时间", "操作人", "动作", "原因", "快照指纹"]); [writer.writerow([a["timestamp"], a["actor_id"], a["action"], a["reason"], a["snapshot_hash"]]) for a in audits]; conn.close(); return StreamingResponse(iter(["\ufeff" + out.getvalue()]), media_type="text/csv", headers={"Content-Disposition": f"attachment; filename={doc['number']}-risk-report.csv"})


@app.post("/api/import/xlsx")
async def import_xlsx(file: UploadFile = File(...), user: sqlite3.Row = Depends(require_roles("applicant", "admin"))) -> dict[str, Any]:
    if load_workbook is None: raise HTTPException(500, "Excel 解析依赖未安装")
    body = await file.read(); errors: list[dict[str, Any]] = []; created: list[str] = []; conn = connect()
    try:
        wb = load_workbook(io.BytesIO(body), read_only=True, data_only=True); ws = wb.active; rows = list(ws.iter_rows(values_only=True))
        if not rows: raise HTTPException(400, "Excel 为空")
        headers = [str(x or "").strip() for x in rows[0]]; required = {"type", "title", "department", "amount"}; missing = required - set(headers)
        if missing: raise HTTPException(400, f"缺少必填列: {', '.join(sorted(missing))}")
        for row_no, values in enumerate(rows[1:], start=2):
            record = dict(zip(headers, values))
            try:
                payload = DocumentIn(type=str(record.get("type") or ""), title=str(record.get("title") or ""), department=str(record.get("department") or ""), payee=str(record.get("payee") or ""), reason=str(record.get("reason") or ""), amount=float(record.get("amount") or 0), detail_total=float(record.get("detail_total") or record.get("amount") or 0), invoice_total=float(record.get("invoice_total") or 0), contract_ratio=float(record.get("contract_ratio") or 1), supplier=str(record.get("supplier") or "示例供应商"))
                existing = conn.execute("SELECT id FROM documents WHERE title=? AND amount=? AND applicant_id=?", (payload.title, payload.amount, user["id"])).fetchone()
                if existing: errors.append({"row": row_no, "field": "title/amount", "error": "疑似重复数据，请确认后再导入"}); continue
                created.append(create_document(conn, payload, user["id"]))
            except Exception as exc: errors.append({"row": row_no, "field": "row", "error": str(exc)})
        append_audit(conn, user["id"], "import_xlsx", "import", file.filename or "xlsx", None, {"created": len(created), "errors": errors}, "固定模板批量导入")
        conn.commit()
    finally: conn.close()
    return {"created_count": len(created), "created_ids": created, "error_count": len(errors), "errors": errors}


@app.get("/api/e2e/demo")
def e2e_demo(user: sqlite3.Row = Depends(require_roles("admin", "auditor", "finance_reviewer", "approver"))) -> dict[str, Any]:
    conn = connect(); doc = conn.execute("SELECT * FROM documents ORDER BY created_at LIMIT 1").fetchone(); data = document_dict(conn, doc["id"]) if doc else None; conn.close(); return {"scenario": "大额技术服务费对公付款", "steps": ["创建", "上传发票/合同", "OCR 解析", "规则交叉核对", "征信与市场价查询", "人工复核", "审批归档", "报告导出"], "document": data}


@app.patch("/api/risk-items/{risk_id}")
def dispose_risk(risk_id: str, payload: dict[str, str], user: sqlite3.Row = Depends(require_roles("finance_reviewer", "approver", "admin"))) -> dict[str, Any]:
    disposition = payload.get("disposition", "pending")
    if disposition not in {"pending", "confirmed", "ignored"}:
        raise HTTPException(400, "不支持的风险处置状态")
    conn = connect(); row = conn.execute("SELECT * FROM risk_items WHERE id=?", (risk_id,)).fetchone()
    if not row: raise HTTPException(404, "风险项不存在")
    before = dict(row)
    conn.execute("UPDATE risk_items SET disposition=?,reviewer_comment=?,updated_at=? WHERE id=?", (disposition, payload.get("reviewer_comment", ""), now(), risk_id))
    append_audit(conn, user["id"], "dispose_risk", "risk_item", risk_id, before, {**before, "disposition": disposition}, payload.get("reviewer_comment", ""))
    conn.commit(); result = dict(conn.execute("SELECT * FROM risk_items WHERE id=?", (risk_id,)).fetchone()); conn.close(); return result


@app.get("/api/reports/{doc_id}/export.xlsx")
def export_xlsx(doc_id: str, user: sqlite3.Row = Depends(require_roles("admin", "auditor", "finance_reviewer", "approver"))):
    if load_workbook is None: raise HTTPException(500, "Excel 导出依赖未安装")
    from openpyxl import Workbook
    conn = connect(); doc = own_or_privileged(conn, doc_id, user); risks = conn.execute("SELECT * FROM risk_items WHERE document_id=?", (doc_id,)).fetchall(); audits = conn.execute("SELECT * FROM audit_events WHERE object_id=? ORDER BY id", (doc_id,)).fetchall(); wb = Workbook(); ws = wb.active; ws.title = "单据摘要"
    ws.append(["单据编号", "类型", "标题", "申请人", "部门", "金额", "风险等级", "状态"]); ws.append([doc["number"], DOC_TYPES[doc["type"]], doc["title"], doc["applicant_id"], doc["department"], float(doc["amount"]), RISK_LEVELS[doc["risk_level"]], STATUSES[doc["status"]]])
    wr = wb.create_sheet("风险明细"); wr.append(["风险代码", "风险项", "等级", "依据", "证据定位", "处置"])
    for risk in risks: wr.append([risk["code"], risk["title"], RISK_LEVELS[risk["level"]], risk["detail"], risk["evidence_ref"], risk["disposition"]])
    wa = wb.create_sheet("审计轨迹"); wa.append(["时间", "操作人", "动作", "原因", "快照指纹"])
    for event in audits: wa.append([event["timestamp"], event["actor_id"], event["action"], event["reason"], event["snapshot_hash"]])
    stream = io.BytesIO(); wb.save(stream); stream.seek(0); conn.close(); return StreamingResponse(stream, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f"attachment; filename={doc['number']}-risk-report.xlsx"})


@app.get("/api/reports/{doc_id}/export.pdf")
def export_pdf(doc_id: str, user: sqlite3.Row = Depends(require_roles("admin", "auditor", "finance_reviewer", "approver"))):
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from reportlab.pdfgen import canvas
    except Exception:
        raise HTTPException(500, "PDF 导出依赖未安装")
    conn = connect(); doc = own_or_privileged(conn, doc_id, user); risks = conn.execute("SELECT * FROM risk_items WHERE document_id=?", (doc_id,)).fetchall(); audits = conn.execute("SELECT * FROM audit_events WHERE object_id=? ORDER BY id", (doc_id,)).fetchall(); stream = io.BytesIO(); pdf = canvas.Canvas(stream, pagesize=A4); width, height = A4; y = height - 48
    font_regular, font_bold = "Helvetica", "Helvetica-Bold"
    font_candidates = [os.getenv("REPORT_FONT"), r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\simhei.ttf", "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"]
    for font_path in font_candidates:
        if font_path and Path(font_path).exists():
            try:
                pdfmetrics.registerFont(TTFont("AppCJK", font_path, subfontIndex=0 if font_path.lower().endswith(".ttc") else 0))
                font_regular = font_bold = "AppCJK"
                break
            except Exception:
                continue
    pdf.setTitle(f"{doc['number']} 风险审核报告"); pdf.setFont(font_bold, 15); pdf.drawString(42, y, "财务单据风险审核报告"); y -= 28; pdf.setFont(font_regular, 10)
    lines = [f"Document: {doc['number']} / {DOC_TYPES[doc['type']]}", f"Title: {doc['title']}", f"Department: {doc['department']}  Amount: RMB {float(doc['amount']):,.2f}", f"Risk: {RISK_LEVELS[doc['risk_level']]}  Status: {STATUSES[doc['status']]}", "", "Risk items:"] + [f"- {r['title']} [{RISK_LEVELS[r['level']]}] {r['detail']}" for r in risks] + ["", "Audit trail:"] + [f"- {a['timestamp']} {a['actor_id']} {a['action']} {a['reason']}" for a in audits]
    for line in lines:
        if y < 52: pdf.showPage(); y = height - 48; pdf.setFont(font_regular, 10)
        pdf.drawString(42, y, line[:125]); y -= 16
    pdf.save(); stream.seek(0); conn.close(); return StreamingResponse(stream, media_type="application/pdf", headers={"Content-Disposition": f"attachment; filename={doc['number']}-risk-report.pdf"})
