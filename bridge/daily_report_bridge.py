#!/usr/bin/env python3
"""HTTP integration layer for the Dify daily store report workflow.

Dify owns routing and user-facing orchestration.  This bridge executes the
deterministic integrations that do not belong in a sandboxed workflow node:
reading Excel files, querying MySQL, updating workbooks, and delivery to a
Dify knowledge base.

By default remote delivery is disabled and production database access requires
environment variables.  Set DAILY_FIXTURE_MODE=true for local regression
tests that must not touch business systems.
"""

from __future__ import annotations

import base64
import importlib.util
import json
import os
import re
import tempfile
import threading
import uuid
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, List, Optional
from urllib.parse import quote

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, PlainTextResponse
from pydantic import BaseModel, Field


ROOT = Path(__file__).resolve().parent.parent
RUNTIME_DIR = Path(os.environ.get("DAILY_RUNTIME_DIR", str(ROOT / "runtime"))).resolve()
OUTPUT_DIR = RUNTIME_DIR / "output"
LOG_DIR = RUNTIME_DIR / "logs"
UPLOAD_DIR = RUNTIME_DIR / "uploads"
LEGACY_MAIN = Path(
    os.environ.get(
        "DAILY_LEGACY_MAIN",
        "/Users/apple/Documents/trae_projects/daily/main.py",
    )
).resolve()
FIXTURE_PATH = Path(
    os.environ.get("DAILY_FIXTURE_PATH", str(ROOT / "fixtures" / "daily_query_fixture.json"))
).resolve()


def env_true(name: str, default: bool = False) -> bool:
    value = os.environ.get(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


FIXTURE_MODE = env_true("DAILY_FIXTURE_MODE")
ENABLE_EXTERNAL_DELIVERY = env_true("DAILY_ENABLE_EXTERNAL_DELIVERY")
ALLOW_LEGACY_DEFAULT_CREDENTIALS = env_true("DAILY_ALLOW_LEGACY_DEFAULT_CREDENTIALS")
KNOWLEDGE_PROVIDER = os.environ.get("DAILY_KNOWLEDGE_PROVIDER", "dify").strip().lower()
DIFY_DATASET_BASE_URL = os.environ.get("DIFY_DATASET_BASE_URL", "http://localhost/v1").rstrip("/")
DIFY_DATASET_ID = os.environ.get("DIFY_DATASET_ID", "")
DIFY_DATASET_API_KEY = os.environ.get("DIFY_DATASET_API_KEY", "")
ALLOWED_OPERATORS = {
    value.strip() for value in os.environ.get("DAILY_ALLOWED_OPERATORS", "").split(",") if value.strip()
}
REPORT_LOCK = threading.Lock()


def allowed_input_dirs() -> List[Path]:
    configured = os.environ.get("DAILY_ALLOWED_INPUT_DIRS", "")
    defaults = ["/Users/apple/Documents/trae_projects/daily/input", str(UPLOAD_DIR)]
    return [Path(value).expanduser().resolve() for value in (configured.split(":") if configured else defaults)]


class EmbeddedFile(BaseModel):
    filename: str
    content_base64: str


class WorkflowRequest(BaseModel):
    operation: str = Field(default="generate_report")
    data_date: Optional[str] = None
    role: str = "operator"
    username: Optional[str] = "运营主管"
    auto_generate_report: bool = True
    upload_to_knowledge_base: bool = False
    send_to_feishu: bool = False
    supplement_days_range: int = 1
    managed_shops_text: str = ""
    shop_list_path: str = ""
    supplement_file_paths: str = ""
    custom_supplements_text: str = ""
    shop_list_file: Optional[EmbeddedFile] = None
    supplement_files: List[EmbeddedFile] = Field(default_factory=list)


class ReportPayload(BaseModel):
    success: bool
    operation: str
    message: str
    operator_name: Optional[str] = None
    report_file_name: Optional[str] = None
    report_generated: bool = False
    data_date: Optional[str] = None
    shop_count: int = 0
    supplement_file_count: int = 0
    output_file: Optional[str] = None
    report_download_url: Optional[str] = None
    knowledge_document_id: Optional[str] = None
    weknora_file_id: Optional[str] = None
    feishu_message_id: Optional[str] = None
    shop_data: Dict[str, Any] = Field(default_factory=dict)
    shops: List[str] = Field(default_factory=list)
    supplements: Dict[str, Any] = Field(default_factory=dict)
    supplement_uploads: List[Dict[str, Any]] = Field(default_factory=list)
    reports: List[Dict[str, Any]] = Field(default_factory=list)
    delivery_skipped: bool = False


app = FastAPI(
    title="Daily Store Report Dify Bridge",
    description="Controlled integration service used by the Dify daily report workflow.",
    version="1.0.0",
)

_legacy_module = None


def load_legacy():
    global _legacy_module
    if _legacy_module is not None:
        return _legacy_module
    if not LEGACY_MAIN.exists():
        raise RuntimeError(f"Legacy implementation not found: {LEGACY_MAIN}")
    spec = importlib.util.spec_from_file_location("daily_store_report_legacy", LEGACY_MAIN)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to load legacy implementation: {LEGACY_MAIN}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    module.OUTPUT_DIR = str(OUTPUT_DIR)
    module.LOG_DIR = str(LOG_DIR)
    _legacy_module = module
    return module


def require_secure_runtime(module, needs_external_delivery: bool, needs_feishu: bool = False) -> None:
    if FIXTURE_MODE and not ENABLE_EXTERNAL_DELIVERY:
        return
    if needs_external_delivery and not ENABLE_EXTERNAL_DELIVERY:
        raise HTTPException(
            status_code=403,
            detail="External knowledge-base/Feishu delivery is disabled by bridge policy.",
        )
    if needs_external_delivery and KNOWLEDGE_PROVIDER == "dify":
        if not DIFY_DATASET_ID or not DIFY_DATASET_API_KEY:
            raise HTTPException(
                status_code=503,
                detail="DIFY_DATASET_ID and DIFY_DATASET_API_KEY are required for Dify knowledge delivery.",
            )
        if FIXTURE_MODE:
            return
    if FIXTURE_MODE:
        return
    db_vars = [
        "SM_DATA_SQL_HOST",
        "SM_DATA_SQL_PORT",
        "SM_DATA_SQL_USER",
        "SM_DATA_SQL_PASSWORD",
        "SM_DATA_SQL_DATABASE",
    ]
    missing = [key for key in db_vars if not os.environ.get(key)]
    if missing and not ALLOW_LEGACY_DEFAULT_CREDENTIALS:
        raise HTTPException(
            status_code=503,
            detail="Database environment configuration is incomplete; refusing legacy hardcoded credentials.",
        )
    if not missing:
        module.DB_CONFIG.update(
            {
                "host": os.environ["SM_DATA_SQL_HOST"],
                "port": int(os.environ["SM_DATA_SQL_PORT"]),
                "user": os.environ["SM_DATA_SQL_USER"],
                "password": os.environ["SM_DATA_SQL_PASSWORD"],
                "db": os.environ["SM_DATA_SQL_DATABASE"],
            }
        )
    if needs_external_delivery and KNOWLEDGE_PROVIDER != "dify":
        required_weknora = ["WEKNORA_API_KEY", "WEKNORA_KNOWLEDGE_BASE_ID"]
        missing_weknora = [name for name in required_weknora if not os.environ.get(name)]
        if missing_weknora and not ALLOW_LEGACY_DEFAULT_CREDENTIALS:
            raise HTTPException(
                status_code=503,
                detail="WEKNORA_API_KEY and WEKNORA_KNOWLEDGE_BASE_ID are required for enabled external delivery.",
            )
        if os.environ.get("WEKNORA_API_KEY"):
            module.WEKNORA_CONFIG["api_key"] = os.environ["WEKNORA_API_KEY"]
        if os.environ.get("WEKNORA_BASE_URL"):
            module.WEKNORA_CONFIG["api_url"] = os.environ["WEKNORA_BASE_URL"]
        if os.environ.get("WEKNORA_KNOWLEDGE_BASE_ID"):
            module.WEKNORA_CONFIG["knowledge_base_id"] = os.environ["WEKNORA_KNOWLEDGE_BASE_ID"]
        if os.environ.get("WEKNORA_OUTPUT_FILE_NAME"):
            module.WEKNORA_CONFIG["output_file_name"] = os.environ["WEKNORA_OUTPUT_FILE_NAME"]
    if needs_feishu and not all(os.environ.get(name) for name in ("FEISHU_APP_ID", "FEISHU_APP_SECRET", "FEISHU_RECEIVE_ID")) and not ALLOW_LEGACY_DEFAULT_CREDENTIALS:
        raise HTTPException(
            status_code=503,
            detail="FEISHU_APP_ID, FEISHU_APP_SECRET and FEISHU_RECEIVE_ID are required for Feishu delivery.",
        )
    if needs_feishu:
        for env_name, config_name in (
            ("FEISHU_APP_ID", "app_id"),
            ("FEISHU_APP_SECRET", "app_secret"),
            ("FEISHU_RECEIVE_ID", "receive_id"),
        ):
            if os.environ.get(env_name):
                module.FEISHU_CONFIG[config_name] = os.environ[env_name]


def validate_local_file(raw_path: str) -> Path:
    path = Path(raw_path.strip()).expanduser().resolve()
    if path.suffix.lower() not in {".xls", ".xlsx"}:
        raise ValueError(f"Only .xls and .xlsx files are accepted: {path.name}")
    if not path.exists() or not path.is_file():
        raise ValueError(f"Input file does not exist: {path}")
    if not any(path.is_relative_to(root) for root in allowed_input_dirs()):
        raise ValueError(f"Input file is outside DAILY_ALLOWED_INPUT_DIRS: {path}")
    return path


def decode_embedded_file(file: EmbeddedFile) -> Path:
    suffix = Path(file.filename).suffix.lower()
    if suffix not in {".xls", ".xlsx"}:
        raise ValueError(f"Only .xls and .xlsx files are accepted: {file.filename}")
    try:
        content = base64.b64decode(file.content_base64, validate=True)
    except ValueError as exc:
        raise ValueError(f"Invalid base64 file content for {file.filename}") from exc
    if len(content) > 10 * 1024 * 1024:
        raise ValueError(f"File exceeds 10 MB bridge limit: {file.filename}")
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    target = UPLOAD_DIR / f"{uuid.uuid4().hex}_{Path(file.filename).name}"
    target.write_bytes(content)
    return target


def split_lines(value: str) -> List[str]:
    return [line.strip() for line in (value or "").replace("，", "\n").splitlines() if line.strip()]


def input_paths(request: WorkflowRequest) -> tuple[Optional[Path], List[Path], List[Path]]:
    temporary: List[Path] = []
    shop_path = validate_local_file(request.shop_list_path) if request.shop_list_path.strip() else None
    supplement_paths = [validate_local_file(value) for value in split_lines(request.supplement_file_paths)]
    if request.shop_list_file:
        shop_path = decode_embedded_file(request.shop_list_file)
        temporary.append(shop_path)
    for file in request.supplement_files:
        path = decode_embedded_file(file)
        supplement_paths.append(path)
        temporary.append(path)
    return shop_path, supplement_paths, temporary


def parse_shop_list(module, path: Path) -> List[str]:
    if path.suffix.lower() == ".xls":
        if module.xlrd is None:
            raise ValueError("xlrd is required to read .xls shop lists")
        book = module.xlrd.open_workbook(str(path))
        rows = [book.sheet_by_index(0).row_values(row) for row in range(book.sheet_by_index(0).nrows)]
    else:
        workbook = module.load_workbook(str(path), data_only=True)
        rows = [list(row) for row in workbook.active.iter_rows(values_only=True)]
        workbook.close()
    skipped_headers = {"店铺", "店铺名", "shop", "shop_name", "店铺名称"}
    shops = {
        str(row[0]).strip()
        for row in rows
        if row and row[0] is not None and str(row[0]).strip().lower() not in skipped_headers
    }
    return sorted(shops)


def parse_operator_name(path: Path) -> str:
    match = re.search(r"(?:[0-9a-f]{32}_)?(.+)-负责店铺列表\.(?:xls|xlsx)$", path.name, re.IGNORECASE)
    if not match:
        raise ValueError("店铺列表文件名必须为：{运营姓名}-负责店铺列表.xlsx")
    operator_name = match.group(1).strip()
    if not operator_name:
        raise ValueError("无法从店铺列表文件名识别运营姓名")
    if ALLOWED_OPERATORS and operator_name not in ALLOWED_OPERATORS:
        raise ValueError(f"运营姓名不在允许名单中: {operator_name}")
    return operator_name


def parse_custom_supplements(value: str) -> Dict[str, Dict[str, float]]:
    data: Dict[str, Dict[str, float]] = {}
    for line in split_lines(value):
        fields = [field.strip() for field in line.split("|")]
        if len(fields) != 3:
            raise ValueError("custom_supplements_text must contain lines formatted as 店铺名称|订单数|金额")
        shop, count, amount = fields
        data[shop] = {"count": int(count), "amount": round(float(amount), 2)}
    return data


def parse_supplement_filename(path: Path) -> tuple[Optional[str], Optional[str]]:
    match = re.match(
        r"^(?:[0-9a-f]{32}_)?(\d{4}-\d{1,2}-\d{1,2})(.*?)(?:\s+\d+单)?\.(?:xls|xlsx)$",
        path.name,
        re.IGNORECASE,
    )
    if not match:
        return None, None
    date_text = match.group(1)
    parsed_date = module_datetime(date_text)
    return parsed_date, match.group(2).strip()


def module_datetime(value: str) -> Optional[str]:
    from datetime import datetime

    try:
        return datetime.strptime(value, "%Y-%m-%d").strftime("%Y-%m-%d")
    except ValueError:
        return None


def parse_supplements(module, paths: List[Path], data_date: Optional[str] = None) -> Dict[str, Dict[str, float]]:
    parser = module.DailyStoreReportAPI.__new__(module.DailyStoreReportAPI)
    totals: Dict[str, Dict[str, float]] = defaultdict(lambda: {"count": 0, "amount": 0.0})
    for path in paths:
        parsed_date, shop_name = parse_supplement_filename(path)
        if not parsed_date or not shop_name:
            raise ValueError(f"Unable to parse date/shop from supplement filename: {path.name}")
        if data_date and parsed_date != data_date:
            continue
        count, amount = parser._parse_supplement_file(str(path), path.suffix.lstrip("."))
        totals[shop_name]["count"] += int(count)
        totals[shop_name]["amount"] = round(totals[shop_name]["amount"] + float(amount), 2)
    return dict(totals)


def archive_supplement_files(module, paths: List[Path]) -> List[Dict[str, Any]]:
    """Archive source supplement files like the legacy upload endpoints."""
    if not paths:
        return []
    if not ENABLE_EXTERNAL_DELIVERY:
        return [{"filename": path.name, "status": "skipped_external_delivery_disabled"} for path in paths]
    require_secure_runtime(module, needs_external_delivery=True)
    reporter = module.DailyStoreReportAPI.__new__(module.DailyStoreReportAPI)
    reporter.log_messages = []
    reporter.log = lambda message: reporter.log_messages.append(message)
    uploads = []
    for path in paths:
        parsed_date, shop_name = parse_supplement_filename(path)
        if not parsed_date or not shop_name:
            uploads.append({"filename": path.name, "status": "skipped_invalid_filename"})
            continue
        count, _ = reporter._parse_supplement_file(str(path), path.suffix.lstrip("."))
        archive_name = f"{parsed_date}{shop_name} {count}单{path.suffix.lower()}"
        upload_url = (
            f"{module.WEKNORA_CONFIG['api_url']}/knowledge-bases/"
            f"{module.WEKNORA_CONFIG['knowledge_base_id']}/knowledge/file"
        )
        with path.open("rb") as file_obj:
            files = {"file": (archive_name, file_obj, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
            response = reporter._weknora_request(
                "POST",
                upload_url,
                files=files,
                data={"enable_multimodel": "false"},
                timeout=30,
            )
        file_id = None
        if response is not None and response.status_code == 200:
            try:
                payload = response.json()
                file_id = payload.get("data", {}).get("id") if payload.get("success") else None
            except ValueError:
                file_id = None
        uploads.append(
            {
                "filename": archive_name,
                "status": "uploaded" if file_id else "failed",
                "weknora_file_id": file_id,
            }
        )
    return uploads


def upload_workflow_reporter_class(module, base):
    class UploadWorkflowReporter(base):
        supplement_headers = ["是否剔除补单数据", "补单金额", "补单数量"]

        def query_daily_data(self, supplement_removed_status=None):
            try:
                daily_data = super().query_daily_data(supplement_removed_status)
            except TypeError:
                daily_data = super().query_daily_data()
            return {shop: daily_data[shop] for shop in self.managed_shops if shop in daily_data}

        def update_summary_sheet(self, wb, daily_data):
            # Ensure daily_data has entries for all managed shops (default to 0)
            # This ensures super().update_summary_sheet writes 0s for shops with no data
            for shop_name in self.managed_shops:
                clean_name = shop_name.strip() if isinstance(shop_name, str) else shop_name
                if clean_name not in daily_data:
                    daily_data[clean_name] = {
                        "order_count": 0, "net_sales": 0, "net_pay_ratio": 0,
                        "visitor_count": 0, "sales_amount": 0
                    }
            
            for shop_name in self.managed_shops:
                clean_name = shop_name.strip() if isinstance(shop_name, str) else shop_name
                if clean_name not in wb.sheetnames or clean_name not in daily_data:
                    continue
                sheet = wb[clean_name]
                row = next(
                    (item for item in range(3, sheet.max_row + 1) if sheet.cell(item, 1).value and str(sheet.cell(item, 1).value).split()[0] == self.date_yesterday),
                    None,
                )
                if row and sheet.cell(row, 12).value is True and shop_name in self.custom_supplement_data:
                    keys = [
                        "date", "visitor_count", "order_count", "sales_amount", "net_sales", "promotion_fee",
                        "pay_ratio", "net_pay_ratio", "daily_expense_rate", "service_fee", "score_rank",
                    ]
                    daily_data[shop_name].update(
                        {key: sheet.cell(row, column).value for key, column in zip(keys, range(1, 12))}
                    )
                    daily_data[shop_name]["is_supplement_removed"] = True
                    daily_data[shop_name]["supplement_amount"] = sheet.cell(row, 13).value or 0
                    daily_data[shop_name]["supplement_count"] = sheet.cell(row, 14).value or 0
            super().update_summary_sheet(wb, daily_data)
            ws = wb["店铺单量总表"]
            date_row = next(
                (
                    row
                    for row in range(3, ws.max_row + 1)
                    if ws.cell(row=row, column=1).value == self.date_yesterday
                ),
                None,
            )
            if not date_row:
                return
            # Preserve values just written by the base reporter; only clear
            # columns for shops no longer present in this upload.
            current_shops = {str(shop).strip() for shop in self.managed_shops}
            for col in range(4, ws.max_column + 1, 3):
                shop_name = ws.cell(row=1, column=col).value
                if shop_name and str(shop_name).strip() not in current_shops:
                    for offset in range(3):
                        ws.cell(row=date_row, column=col + offset).value = None

        def update_shop_sheet(self, wb, shop_name, shop_data):
            if shop_name in wb.sheetnames:
                ws = wb[shop_name]
                row = next(
                    (item for item in range(3, ws.max_row + 1) if ws.cell(item, 1).value and str(ws.cell(item, 1).value).split()[0] == self.date_yesterday),
                    None,
                )
                if row and ws.cell(row, 12).value is True and shop_name in self.custom_supplement_data:
                    self.log(f"[INFO] {shop_name} 在 {self.date_yesterday} 已剔除补单，保留原日报行")
                    return
            super().update_shop_sheet(wb, shop_name, shop_data)
            ws = wb[shop_name]
            for column, header in enumerate(self.supplement_headers, start=12):
                ws.cell(row=2, column=column, value=header)
            date_row = next(
                (row for row in range(3, ws.max_row + 1) if ws.cell(row, 1).value == self.date_yesterday),
                ws.max_row,
            )
            has_current_supplement = shop_name in self.custom_supplement_data
            ws.cell(date_row, 12, value=shop_data.get("is_supplement_removed", False) or has_current_supplement)
            ws.cell(date_row, 13, value=shop_data.get("supplement_amount", 0) if has_current_supplement else 0)
            ws.cell(date_row, 14, value=shop_data.get("supplement_count", 0) if has_current_supplement else 0)

    return UploadWorkflowReporter


def dify_dataset_reporter_class(module, base):
    class DifyDatasetReporter(base):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self._dify_session = module.requests.Session()
            self._dify_session.trust_env = False

        def _report_file_name(self):
            return module.WEKNORA_CONFIG["output_file_name"]

        def _dify_headers(self):
            return {"Authorization": f"Bearer {DIFY_DATASET_API_KEY}"}

        def _dify_documents(self):
            response = self._dify_session.get(
                f"{DIFY_DATASET_BASE_URL}/datasets/{DIFY_DATASET_ID}/documents",
                headers=self._dify_headers(),
                params={"keyword": self._report_file_name(), "page": 1, "limit": 20},
                timeout=20,
            )
            response.raise_for_status()
            return [
                document
                for document in response.json().get("data", [])
                if document.get("name") == self._report_file_name()
            ]

        def download_template_from_knowledge_base(self):
            documents = self._dify_documents()
            if not documents:
                self.log("[INFO] Dify 知识库中无历史日报，创建新工作簿")
                return None
            document_id = documents[0]["id"]
            response = self._dify_session.get(
                f"{DIFY_DATASET_BASE_URL}/datasets/{DIFY_DATASET_ID}/documents/{document_id}/download",
                headers=self._dify_headers(),
                timeout=20,
            )
            response.raise_for_status()
            download_url = response.json()["url"]
            if download_url.startswith("/"):
                download_url = f"{DIFY_DATASET_BASE_URL.split('/v1', 1)[0]}{download_url}"
            content = self._dify_session.get(download_url, timeout=30)
            content.raise_for_status()
            temporary = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            temporary.write(content.content)
            temporary.close()
            self.log(f"[OK] 从 Dify 知识库下载历史日报: {document_id}")
            return temporary.name

        def delete_old_file_from_knowledge_base(self):
            for document in self._dify_documents():
                response = self._dify_session.delete(
                    f"{DIFY_DATASET_BASE_URL}/datasets/{DIFY_DATASET_ID}/documents/{document['id']}",
                    headers=self._dify_headers(),
                    timeout=20,
                )
                response.raise_for_status()
                self.log(f"[OK] 删除 Dify 知识库旧日报: {document['id']}")

        def upload_to_knowledge_base(self, file_path):
            with open(file_path, "rb") as report_file:
                response = self._dify_session.post(
                    f"{DIFY_DATASET_BASE_URL}/datasets/{DIFY_DATASET_ID}/document/create-by-file",
                    headers=self._dify_headers(),
                    files={
                        "file": (
                            self._report_file_name(),
                            report_file,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                    },
                    data={
                        "data": json.dumps(
                            {"indexing_technique": "economy", "process_rule": {"mode": "automatic"}},
                            ensure_ascii=False,
                        )
                    },
                    timeout=60,
                )
            response.raise_for_status()
            document_id = response.json().get("document", {}).get("id")
            self.log(f"[OK] 上传最新版日报到 Dify 知识库: {document_id}")
            return document_id

    return DifyDatasetReporter


def fixture_reporter_class(module):
    fixture = json.loads(FIXTURE_PATH.read_text(encoding="utf-8"))
    base = upload_workflow_reporter_class(module, module.DailyStoreReportAPI)
    real_dify_delivery = KNOWLEDGE_PROVIDER == "dify" and ENABLE_EXTERNAL_DELIVERY
    if real_dify_delivery:
        base = dify_dataset_reporter_class(module, base)

    class FixtureReporter(base):
        def run_mysql_query(self, query):
            if "FROM sys_users" in query:
                return "fixture-user"
            if "FROM mapping_shops" in query:
                return "\n".join(fixture["managed_shops"])
            data = fixture["dates"].get(self.date_yesterday, {})
            if "FROM pdd_mall_daily_performance" in query:
                return "\n".join("\t".join(str(value) for value in row) for row in data.get("performance_rows", []))
            if "FROM pdd_mall_daily_summary" in query:
                return "\n".join("\t".join(str(value) for value in row) for row in data.get("expense_rows", []))
            return ""

        def get_supplement_amounts(self):
            return None

        def download_template_from_knowledge_base(self):
            if real_dify_delivery:
                return super().download_template_from_knowledge_base()
            existing = OUTPUT_DIR / module.WEKNORA_CONFIG["output_file_name"]
            if not existing.exists():
                return None
            temporary = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            temporary.write(existing.read_bytes())
            temporary.close()
            return temporary.name

        def delete_old_file_from_knowledge_base(self):
            if real_dify_delivery:
                return super().delete_old_file_from_knowledge_base()
            self.log("[FIXTURE] 模拟删除知识库旧日报文件")

        def upload_to_knowledge_base(self, file_path):
            if real_dify_delivery:
                return super().upload_to_knowledge_base(file_path)
            self.log("[FIXTURE] 模拟上传最新日报到知识库")
            return "fixture-latest-report"

    return FixtureReporter


def reporter_class(module):
    if FIXTURE_MODE:
        return fixture_reporter_class(module)
    base = upload_workflow_reporter_class(module, module.DailyStoreReportAPI)
    return dify_dataset_reporter_class(module, base) if KNOWLEDGE_PROVIDER == "dify" else base


def run_single_report(
    module,
    request: WorkflowRequest,
    data_date: Optional[str],
    managed_shops: Optional[List[str]],
    supplements: Optional[Dict[str, Dict[str, float]]],
) -> Dict[str, Any]:
    reporter = reporter_class(module)(
        role=request.role,
        username=request.username,
        test_date=data_date,
        supplement_days_range=request.supplement_days_range,
        custom_shops=managed_shops if managed_shops else None,
        custom_supplement_data=supplements or None,
    )
    reporter.is_custom_shops = False
    if KNOWLEDGE_PROVIDER == "dify" and ENABLE_EXTERNAL_DELIVERY:
        local_output = OUTPUT_DIR / reporter.output_file_name
        local_output.unlink(missing_ok=True)
    result = reporter.run(
        upload_to_knowledge_base=request.upload_to_knowledge_base,
        send_to_feishu=request.send_to_feishu,
    )
    return {
        "success": True,
        "report_generated": True,
        "data_date": reporter.date_yesterday,
        "shop_count": len(result["daily_data"]),
        "output_file": result["output_file"],
        "report_download_url": "/v1/report/download",
        "knowledge_document_id": result.get("weknora_file_id"),
        "weknora_file_id": result.get("weknora_file_id"),
        "feishu_message_id": result.get("feishu_message_id"),
        "shop_data": result["daily_data"],
    }


@app.get("/health")
async def health() -> Dict[str, Any]:
    return {
        "status": "healthy",
        "fixture_mode": FIXTURE_MODE,
        "external_delivery_enabled": ENABLE_EXTERNAL_DELIVERY,
        "legacy_main_exists": LEGACY_MAIN.exists(),
    }


@app.post("/v1/files/upload")
async def upload_input_file(file: UploadFile = File(...)) -> Dict[str, Any]:
    if not file.filename or Path(file.filename).suffix.lower() not in {".xls", ".xlsx"}:
        raise HTTPException(status_code=400, detail="Only .xls and .xlsx files are accepted.")
    data = await file.read()
    if len(data) > 10 * 1024 * 1024:
        raise HTTPException(status_code=413, detail="File exceeds 10 MB bridge limit.")
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    output = UPLOAD_DIR / f"{uuid.uuid4().hex}_{Path(file.filename).name}"
    output.write_bytes(data)
    return {"success": True, "path": str(output), "filename": file.filename}


async def embedded_upload(file: UploadFile) -> EmbeddedFile:
    if not file.filename or Path(file.filename).suffix.lower() not in {".xls", ".xlsx"}:
        raise HTTPException(status_code=400, detail="Only .xls and .xlsx files are accepted.")
    data = await file.read()
    if len(data) > 10 * 1024 * 1024:
        raise HTTPException(status_code=413, detail="File exceeds 10 MB bridge limit.")
    return EmbeddedFile(filename=file.filename, content_base64=base64.b64encode(data).decode("ascii"))


@app.post("/v1/workflow/run-files-download")
async def workflow_run_files_download(
    data_date: Optional[str] = Form(None),
    shop_list_file: UploadFile = File(...),
    supplement_files: Optional[List[UploadFile]] = File(None),
):
    from datetime import datetime, timedelta
    if not data_date:
        data_date = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    supplement_files_list = supplement_files or []
    try:
        request = WorkflowRequest(
            operation="generate_from_uploads",
            data_date=data_date,
            shop_list_file=await embedded_upload(shop_list_file),
            supplement_files=[await embedded_upload(file) for file in supplement_files_list],
        )
        report = await workflow_run(request)
    except HTTPException as exc:
        if 400 <= exc.status_code < 500:
            detail = str(exc.detail)
            return PlainTextResponse(
                f"日报生成失败：{detail}",
                status_code=200,
                headers={"X-Workflow-Success": "false"},
            )
        raise
    report_path = Path(report.output_file or "")
    if not report_path.exists():
        raise HTTPException(status_code=500, detail="Generated report file does not exist.")
    filename = report.report_file_name or report_path.name
    return FileResponse(
        str(report_path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "X-Workflow-Success": "true",
            "X-Report-Operator": quote(report.operator_name or ""),
            "X-Knowledge-Document-Id": report.knowledge_document_id or "",
        },
    )


@app.post("/v1/workflow/run", response_model=ReportPayload)
async def workflow_run(request: WorkflowRequest) -> ReportPayload:
    module = load_legacy()
    operation = request.operation.strip()
    permitted = {
        "generate_from_uploads",
        "generate_report",
        "upload_shop_list",
        "upload_shop_list_with_supplement",
        "upload_supplement",
        "batch_upload_supplement",
    }
    if operation not in permitted:
        raise HTTPException(status_code=400, detail=f"Unsupported operation: {operation}")
    needs_delivery = operation == "generate_from_uploads" or request.upload_to_knowledge_base or request.send_to_feishu
    require_secure_runtime(module, needs_external_delivery=needs_delivery, needs_feishu=request.send_to_feishu)

    temporary: List[Path] = []
    try:
        shop_path, supplement_paths, temporary = input_paths(request)
        managed_shops = split_lines(request.managed_shops_text)
        if shop_path:
            managed_shops = parse_shop_list(module, shop_path)
        supplements = parse_custom_supplements(request.custom_supplements_text)

        if operation == "generate_from_uploads":
            from datetime import datetime, timedelta
            if not request.data_date:
                request.data_date = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
            if not shop_path:
                raise ValueError("generate_from_uploads requires an uploaded shop list Excel file")
            if request.managed_shops_text.strip() or request.custom_supplements_text.strip():
                raise ValueError("generate_from_uploads accepts shops and supplements only from uploaded Excel files")
            managed_shops = parse_shop_list(module, shop_path)
            operator_name = parse_operator_name(shop_path)
            report_file_name = f"{operator_name}的店铺数据日报.xlsx"
            if not managed_shops:
                raise ValueError("No shops were parsed from the uploaded shop list")
            supplements = {}
            if supplement_paths:
                supplements = parse_supplements(module, supplement_paths, request.data_date)
                if not supplements:
                    supplements = parse_supplements(module, supplement_paths, None)
            unexpected_shops = sorted(set(supplements) - set(managed_shops))
            if unexpected_shops:
                raise ValueError(f"补单店铺不在本次负责店铺列表中: {', '.join(unexpected_shops)}")
            report_request = request.model_copy(
                update={
                    "username": operator_name,
                    "upload_to_knowledge_base": True,
                    "send_to_feishu": False,
                }
            )
            with REPORT_LOCK:
                module.WEKNORA_CONFIG["output_file_name"] = report_file_name
                report = run_single_report(module, report_request, request.data_date, managed_shops, supplements)
            report["report_download_url"] = f"/v1/report/download?operator_name={quote(operator_name)}"
            return ReportPayload(
                operation=operation,
                message="上传名单及补单生成日报成功，最新日报已替换至知识库",
                operator_name=operator_name,
                report_file_name=report_file_name,
                supplement_file_count=len(supplement_paths),
                shops=managed_shops,
                supplements=supplements,
                **report,
            )

        if operation == "generate_report":
            report = run_single_report(module, request, request.data_date, managed_shops or None, supplements or None)
            return ReportPayload(operation=operation, message="报表生成成功", supplements=supplements, **report)

        if operation == "upload_shop_list":
            if not managed_shops:
                raise ValueError("upload_shop_list requires a shop list file or managed_shops_text")
            if not request.auto_generate_report:
                return ReportPayload(
                    success=True,
                    operation=operation,
                    message=f"店铺列表解析成功，共 {len(managed_shops)} 个店铺",
                    shop_count=len(managed_shops),
                    shops=managed_shops,
                )
            report = run_single_report(module, request, request.data_date, managed_shops, supplements or None)
            return ReportPayload(
                operation=operation,
                message="店铺列表解析并生成报表成功",
                shops=managed_shops,
                supplements=supplements,
                **report,
            )

        if operation == "upload_shop_list_with_supplement":
            if not managed_shops:
                raise ValueError("shop list input is required")
            supplements.update(parse_supplements(module, supplement_paths, request.data_date))
            if not request.auto_generate_report:
                return ReportPayload(
                    success=True,
                    operation=operation,
                    message="店铺列表及补单解析成功",
                    shop_count=len(managed_shops),
                    supplement_file_count=len(supplement_paths),
                    shops=managed_shops,
                    supplements=supplements,
                )
            report = run_single_report(module, request, request.data_date, managed_shops, supplements)
            return ReportPayload(
                operation=operation,
                message="店铺列表及补单处理成功",
                supplement_file_count=len(supplement_paths),
                shops=managed_shops,
                supplements=supplements,
                **report,
            )

        if not supplement_paths and not supplements:
            raise ValueError(f"{operation} requires supplement files or custom_supplements_text")

        if operation == "upload_supplement":
            supplement_uploads = archive_supplement_files(module, supplement_paths[:1])
            if supplement_paths:
                parsed_date, _ = parse_supplement_filename(supplement_paths[0])
                effective_date = request.data_date or parsed_date
                supplements.update(parse_supplements(module, [supplement_paths[0]], effective_date))
            else:
                effective_date = request.data_date
            if not request.auto_generate_report:
                return ReportPayload(
                    success=True,
                    operation=operation,
                    message="补单数据解析成功",
                    supplement_file_count=len(supplement_paths),
                    supplements=supplements,
                    supplement_uploads=supplement_uploads,
                    delivery_skipped=bool(supplement_paths) and not ENABLE_EXTERNAL_DELIVERY,
                )
            report = run_single_report(module, request, effective_date, managed_shops or None, supplements)
            return ReportPayload(
                operation=operation,
                message="补单处理并生成报表成功",
                supplement_file_count=len(supplement_paths),
                supplements=supplements,
                supplement_uploads=supplement_uploads,
                delivery_skipped=bool(supplement_paths) and not ENABLE_EXTERNAL_DELIVERY,
                **report,
            )

        supplement_uploads = archive_supplement_files(module, supplement_paths)
        grouped: Dict[str, List[Path]] = defaultdict(list)
        for path in supplement_paths:
            parsed_date, _ = parse_supplement_filename(path)
            if parsed_date:
                grouped[parsed_date].append(path)
        if request.data_date and supplements:
            grouped.setdefault(request.data_date, [])
        reports: List[Dict[str, Any]] = []
        all_supplements: Dict[str, Any] = {}
        for data_date, paths in sorted(grouped.items()):
            data = dict(supplements) if data_date == request.data_date else {}
            data.update(parse_supplements(module, paths, data_date))
            all_supplements[data_date] = data
            if request.auto_generate_report:
                reports.append(run_single_report(module, request, data_date, managed_shops or None, data))
        last_report = reports[-1] if reports else {}
        return ReportPayload(
            operation=operation,
            message=f"批量补单处理完成，共 {len(grouped)} 个数据日期",
            report_generated=bool(reports),
            supplement_file_count=len(supplement_paths),
            supplements=all_supplements,
            supplement_uploads=supplement_uploads,
            reports=reports,
            delivery_skipped=bool(supplement_paths) and not ENABLE_EXTERNAL_DELIVERY,
            success=True,
            data_date=last_report.get("data_date"),
            shop_count=last_report.get("shop_count", 0),
            output_file=last_report.get("output_file"),
            report_download_url=last_report.get("report_download_url"),
            shop_data=last_report.get("shop_data", {}),
        )
    except HTTPException:
        raise
    except (OSError, ValueError, RuntimeError) as exc:
        print(
            "[ERROR] workflow failure "
            + json.dumps({"operation": operation, "detail": str(exc)}, ensure_ascii=True),
            flush=True,
        )
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    finally:
        for path in temporary:
            try:
                path.unlink(missing_ok=True)
            except OSError:
                pass


@app.get("/v1/report/download")
async def download_report(operator_name: Optional[str] = None):
    module = load_legacy()
    if operator_name:
        operator_name = operator_name.strip()
        if not operator_name or "/" in operator_name or "\\" in operator_name:
            raise HTTPException(status_code=400, detail="Invalid operator_name.")
        if ALLOWED_OPERATORS and operator_name not in ALLOWED_OPERATORS:
            raise HTTPException(status_code=404, detail="Unknown operator_name.")
        report_name = f"{operator_name}的店铺数据日报.xlsx"
    else:
        report_name = module.WEKNORA_CONFIG["output_file_name"]
    report_path = Path(module.OUTPUT_DIR) / report_name
    if not report_path.exists():
        raise HTTPException(status_code=404, detail="Report file does not exist.")
    filename = report_path.name
    return FileResponse(
        str(report_path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("DAILY_BRIDGE_PORT", "8767")))
