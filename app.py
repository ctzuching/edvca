"""企業文件新舊版本差異分析助手 - Flask 主程式：路由、上傳驗證、工作管理、下載。"""
import json
import logging
import os
import re
import secrets
import shutil
import time
import uuid
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from flask import (Flask, abort, redirect, render_template, request, send_file,
                   send_from_directory, session, url_for)

load_dotenv()

import ai_analysis  # noqa: E402
import docdiff  # noqa: E402
import report_builder  # noqa: E402

BASE_DIR = Path(__file__).resolve().parent
JOB_ID_RE = re.compile(r"[0-9a-f]{32}")
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
for noisy in ("httpx", "httpcore", "google", "urllib3"):
    logging.getLogger(noisy).setLevel(logging.WARNING)  # 避免第三方套件記錄請求細節
logger = logging.getLogger("docdiff")


def _clean_filename(name: str) -> str:
    """僅供顯示：移除路徑與控制字元；從不用於伺服器檔案路徑。"""
    name = os.path.basename((name or "").replace("\\", "/"))
    name = re.sub(r"[\x00-\x1f\x7f]", "", name).strip()
    return name[:120] or "未命名.docx"


def create_app(test_config=None):
    app = Flask(__name__, static_folder="static", static_url_path="/static", template_folder="templates")
    max_mb = int(os.getenv("MAX_UPLOAD_MB", "10"))
    secret = os.getenv("FLASK_SECRET_KEY", "").strip()
    if not secret:
        secret = secrets.token_hex(32)
        logger.warning("未設定 FLASK_SECRET_KEY，已使用隨機金鑰；重啟後使用者將無法存取先前的分析結果。")
    app.config.update(
        SECRET_KEY=secret,
        MAX_UPLOAD_MB=max_mb,
        MAX_CONTENT_LENGTH=(2 * max_mb + 1) * 1024 * 1024,
        JOBS_DIR=str(Path(os.getenv("JOBS_DIR", BASE_DIR / "instance" / "jobs"))),
        JOB_TTL_MINUTES=int(os.getenv("JOB_TTL_MINUTES", "60")),
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE="Lax",
    )
    if test_config:
        app.config.update(test_config)
    state = {"last_cleanup": 0.0}

    def jobs_dir() -> Path:
        d = Path(app.config["JOBS_DIR"])
        d.mkdir(parents=True, exist_ok=True)
        return d

    def cleanup_jobs(force=False):
        now = time.time()
        if not force and now - state["last_cleanup"] < 300:
            return
        state["last_cleanup"] = now
        ttl = app.config["JOB_TTL_MINUTES"] * 60
        for d in jobs_dir().iterdir():
            try:
                if d.is_dir() and JOB_ID_RE.fullmatch(d.name) and now - d.stat().st_mtime > ttl:
                    shutil.rmtree(d, ignore_errors=True)
            except OSError:
                continue

    app.cleanup_jobs = cleanup_jobs

    def read_upload(field: str, label: str):
        f = request.files.get(field)
        if f is None or not f.filename:
            raise docdiff.DocumentError(f"請選擇{label}的 Word 文件（.docx）。")
        name = _clean_filename(f.filename)
        if not name.lower().endswith(".docx"):
            raise docdiff.DocumentError(f"{label}必須是 .docx 格式的 Word 文件（不支援 .doc 或其他格式）。")
        limit = app.config["MAX_UPLOAD_MB"] * 1024 * 1024
        data = f.read(limit + 1)
        if len(data) > limit:
            raise docdiff.DocumentError(f"{label}超過大小上限（{app.config['MAX_UPLOAD_MB']} MB）。")
        return name, data

    def load_job(job_id: str) -> dict:
        if not JOB_ID_RE.fullmatch(job_id or ""):
            abort(404)
        try:
            data = json.loads((jobs_dir() / job_id / "result.json").read_text("utf-8"))
        except (OSError, ValueError):
            abort(404)
        sid, owner = session.get("sid"), data.get("owner")
        if not sid or not owner or not secrets.compare_digest(str(sid), str(owner)):
            abort(404)  # 不洩漏該工作是否存在
        return data

    # ------------------------------------------------------------ 路由 ----
    @app.before_request
    def _periodic_cleanup():
        cleanup_jobs()

    @app.after_request
    def _headers(resp):
        resp.headers["X-Content-Type-Options"] = "nosniff"
        resp.headers["X-Frame-Options"] = "DENY"
        resp.headers["Content-Security-Policy"] = "default-src 'self'; frame-ancestors 'none'"
        if request.path.startswith(("/result/", "/download/", "/analyze")):
            resp.headers["Cache-Control"] = "no-store"
        return resp

    @app.get("/")
    def index():
        return send_from_directory(app.static_folder, "index.html")

    @app.post("/analyze")
    def analyze():
        try:
            old_name, old_bytes = read_upload("old_file", "舊版本")
            new_name, new_bytes = read_upload("new_file", "新版本")
            old_doc = docdiff.load_docx(old_bytes, "舊版本")
            new_doc = docdiff.load_docx(new_bytes, "新版本")
        except docdiff.DocumentError as e:
            return render_template("output.html", error=str(e)), 400
        except Exception as e:
            logger.error("上傳處理失敗：%s", type(e).__name__)
            return render_template("output.html", error="處理上傳檔案時發生錯誤，請確認檔案後再試。"), 400

        try:
            diff = docdiff.compare(old_doc["paragraphs"], new_doc["paragraphs"])
            ai = ai_analysis.analyze(diff["items"], diff["stats"])
            result = {
                "old_name": old_name, "new_name": new_name,
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "stats": diff["stats"], "items": diff["items"], "ai": ai,
                "limitations": (docdiff.STATIC_LIMITATIONS + diff["notes"]
                                + docdiff.feature_notes("舊版本", old_doc["features"])
                                + docdiff.feature_notes("新版本", new_doc["features"])),
            }
            sid = session.get("sid") or secrets.token_hex(16)
            session["sid"] = sid
            job_id = uuid.uuid4().hex
            job_path = jobs_dir() / job_id
            job_path.mkdir(parents=True)
            report_builder.build_report(result, job_path / "report.docx")
            (job_path / "result.json").write_text(
                json.dumps({"owner": sid, **result}, ensure_ascii=False), encoding="utf-8")
        except Exception as e:
            logger.error("分析流程失敗：%s", type(e).__name__)
            return render_template("output.html", error="分析過程發生未預期的錯誤，請稍後再試。"), 500

        logger.info("job=%s added=%d deleted=%d modified=%d ai=%s", job_id[:8],
                    diff["stats"]["added"], diff["stats"]["deleted"], diff["stats"]["modified"], ai["status"])
        return redirect(url_for("result", job_id=job_id), code=303)

    @app.get("/result/<job_id>")
    def result(job_id):
        data = load_job(job_id)
        data.pop("owner", None)
        return render_template("output.html", result=data, job_id=job_id)

    @app.get("/download/<job_id>")
    def download(job_id):
        load_job(job_id)  # 驗證識別碼與存取權限
        path = jobs_dir() / job_id / "report.docx"
        if not path.is_file():
            abort(404)
        return send_file(path, as_attachment=True, mimetype=DOCX_MIME,
                         download_name=f"企業文件差異分析報告_{job_id[:8]}.docx")

    # ---------------------------------------------------------- 錯誤頁 ----
    @app.errorhandler(404)
    def _404(_):
        return render_template("output.html", error="找不到頁面，或分析結果已過期／您沒有存取權限。"), 404

    @app.errorhandler(405)
    def _405(_):
        return render_template("output.html", error="不支援的請求方式。"), 405

    @app.errorhandler(413)
    def _413(_):
        return render_template(
            "output.html", error=f"上傳內容過大，每個檔案上限為 {app.config['MAX_UPLOAD_MB']} MB。"), 413

    @app.errorhandler(500)
    def _500(_):
        return render_template("output.html", error="系統發生錯誤，請稍後再試。"), 500

    cleanup_jobs(force=True)
    return app


app = create_app()

if __name__ == "__main__":
    app.run(host=os.getenv("HOST", "127.0.0.1"), port=int(os.getenv("PORT", "5000")), debug=False)
