"""Grade Report Web Application
Flask-based web interface with user authentication for generating student grade reports.
"""

import io
import os
import re
import shutil
import sqlite3
import tempfile
import time
import uuid
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from contextlib import contextmanager
from pathlib import Path
from urllib.parse import urlparse

from flask import (
    Flask,
    g,
    render_template,
    request,
    redirect,
    url_for,
    send_file,
    abort,
    flash,
    stream_with_context,
    Response,
)
from flask_login import (
    LoginManager,
    UserMixin,
    login_user,
    logout_user,
    login_required,
    current_user,
)
from flask_wtf.csrf import CSRFProtect
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from werkzeug.security import generate_password_hash, check_password_hash

from werkzeug.middleware.proxy_fix import ProxyFix

from generate_reports import load_grades, build_student_report

# ── App setup ──────────────────────────────────────────────────────────────────

_secret = os.environ.get("SECRET_KEY")
if not _secret:
    raise RuntimeError(
        "SECRET_KEY environment variable is required. "
        "Set it in your .env file or environment."
    )

app = Flask(__name__)
app.secret_key = _secret

# Trust exactly one proxy (Caddy) for X-Forwarded-For / X-Forwarded-Proto.
# This lets the rate-limiter see real client IPs and allows HTTPS detection.
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)

# Session cookie hardening (#2)
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = (
    os.environ.get("COOKIE_SECURE", "true").lower() == "true"
)
# WTF_CSRF_TIME_LIMIT: tokens valid for 1 hour (default)
app.config["WTF_CSRF_TIME_LIMIT"] = 3600

csrf = CSRFProtect(app)  # CSRF protection for all POST/PUT/PATCH/DELETE (#1)

# Rate limiter — in-memory, no Redis required (#4)
limiter = Limiter(
    get_remote_address,
    app=app,
    default_limits=[],
    storage_uri="memory://",
)

UPLOAD_ROOT = Path(
    os.environ.get("UPLOAD_ROOT", os.path.join(
        tempfile.gettempdir(), "reportgen"))
)
ALLOWED_EXTENSIONS = {".xlsx", ".xls"}
MAX_UPLOAD_MB = 10
SESSION_TTL_SECONDS = 24 * 3600   # auto-clean sessions older than 24 h

# Password must be ≥8 chars with at least one letter and one digit (#7)
_PASSWORD_RE = re.compile(r"^(?=.*[A-Za-z])(?=.*\d).{8,}$")

# Sanitise filenames to safe characters (#8)
_SAFE_FILENAME_RE = re.compile(r"[^\w\-.]")

# ── Database ───────────────────────────────────────────────────────────────────

DB_PATH = Path(os.environ.get("DB_PATH", "/data/users.db"))


@contextmanager
def _startup_db():
    """Context manager for use outside request context (startup only)."""
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def get_db() -> sqlite3.Connection:
    """Return a request-scoped DB connection (#12 — one connection per request)."""
    if "db" not in g:
        DB_PATH.parent.mkdir(parents=True, exist_ok=True)
        g.db = sqlite3.connect(str(DB_PATH))
        g.db.row_factory = sqlite3.Row
    return g.db


@app.teardown_appcontext
def close_db(e=None):
    db = g.pop("db", None)
    if db is not None:
        db.commit()
        db.close()


def _init_db() -> None:
    with _startup_db() as db:
        db.execute("""
            CREATE TABLE IF NOT EXISTS users (
                username      TEXT PRIMARY KEY,
                password_hash TEXT NOT NULL,
                school        TEXT NOT NULL DEFAULT '',
                full_name     TEXT NOT NULL DEFAULT '',
                email         TEXT NOT NULL DEFAULT '',
                is_admin      INTEGER NOT NULL DEFAULT 0
            )
        """)
        # Migrate: add columns introduced in later versions
        for col, defn in [
            ("full_name", "TEXT NOT NULL DEFAULT ''"),
            ("email",     "TEXT NOT NULL DEFAULT ''"),
        ]:
            try:
                db.execute(f"ALTER TABLE users ADD COLUMN {col} {defn}")
            except sqlite3.OperationalError:
                pass  # column already exists
        db.execute("""
            CREATE TABLE IF NOT EXISTS sessions (
                uid        TEXT PRIMARY KEY,
                username   TEXT NOT NULL,
                created_at REAL NOT NULL
            )
        """)


def _seed_admin() -> None:
    """Create the bootstrap admin account if no users exist yet."""
    username = os.environ.get("ADMIN_USERNAME", "admin")
    password = os.environ.get("ADMIN_PASSWORD", "changeme")
    school = os.environ.get("ADMIN_SCHOOL",   "Westfield Academy")
    with _startup_db() as db:
        exists = db.execute(
            "SELECT 1 FROM users WHERE username = ?", (username,)
        ).fetchone()
        if not exists:
            db.execute(
                "INSERT INTO users (username, password_hash, school, is_admin) "
                "VALUES (?, ?, ?, 1)",
                (username, generate_password_hash(password), school),
            )


_init_db()
_seed_admin()

# ── Authentication ─────────────────────────────────────────────────────────────

login_manager = LoginManager(app)
login_manager.login_view = "login"
login_manager.login_message = "Please log in to access this page."


class User(UserMixin):
    def __init__(self, username: str, password_hash: str, school: str,
                 full_name: str, email: str, is_admin: bool):
        self.id = username
        self.username = username
        self.password_hash = password_hash
        self.school = school
        self.full_name = full_name
        self.email = email
        self.is_admin = is_admin

    @classmethod
    def from_row(cls, row) -> "User":
        return cls(row["username"], row["password_hash"],
                   row["school"], row["full_name"], row["email"],
                   bool(row["is_admin"]))


def _db_get_user(username: str) -> "User | None":
    row = get_db().execute(
        "SELECT * FROM users WHERE username = ?", (username,)
    ).fetchone()
    return User.from_row(row) if row else None


def _db_all_users() -> list[dict]:
    rows = get_db().execute(
        "SELECT username, school, is_admin FROM users ORDER BY username"
    ).fetchall()
    return [dict(r) for r in rows]


@login_manager.user_loader
def load_user(user_id: str):
    return _db_get_user(user_id)


def _is_safe_redirect(target: str) -> bool:
    """Allow only relative paths on the same origin (no scheme, no netloc)."""
    parsed = urlparse(target)
    return (
        not parsed.scheme
        and not parsed.netloc
        and parsed.path.startswith("/")
        and not parsed.path.startswith("//")
    )


# ── Security headers (#3) ──────────────────────────────────────────────────────

@app.after_request
def set_security_headers(resp):
    resp.headers["X-Frame-Options"] = "DENY"
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    resp.headers["Content-Security-Policy"] = (
        "default-src 'self'; "
        "script-src 'self'; "
        "style-src 'self' 'unsafe-inline'; "
        "font-src 'self'; "
        "object-src 'none'; "
        "frame-ancestors 'none'"
    )
    return resp


# ── Helpers ────────────────────────────────────────────────────────────────────

def _session_dir(uid: str) -> Path:
    """Return the working directory for a session, creating it if needed."""
    d = UPLOAD_ROOT / uid
    d.mkdir(parents=True, exist_ok=True)
    return d


def _safe_uid(uid: str) -> str:
    """Validate that uid is a well-formed UUID (prevents path traversal)."""
    try:
        return str(uuid.UUID(uid))
    except ValueError:
        abort(400)


def _check_session_owner(uid: str) -> None:
    """Abort 403 if the current user does not own this session (#6)."""
    row = get_db().execute(
        "SELECT username FROM sessions WHERE uid = ?", (uid,)
    ).fetchone()
    if not row or row["username"] != current_user.username:
        abort(403)


def _validate_password(password: str) -> str | None:
    """Return an error string if the password fails complexity rules, else None (#7)."""
    if not _PASSWORD_RE.match(password):
        return "Password must be at least 8 characters with at least one letter and one digit."
    return None


def _safe_filename(name: str) -> str:
    """Sanitise a filename to only safe characters (#8)."""
    return _SAFE_FILENAME_RE.sub("_", name)


def _cleanup_old_sessions() -> None:
    """Delete session directories and DB rows older than SESSION_TTL_SECONDS (#6)."""
    cutoff = time.time() - SESSION_TTL_SECONDS
    db = get_db()
    old_rows = db.execute(
        "SELECT uid FROM sessions WHERE created_at < ?", (cutoff,)
    ).fetchall()
    for row in old_rows:
        session_path = UPLOAD_ROOT / row["uid"]
        if session_path.exists():
            shutil.rmtree(session_path, ignore_errors=True)
    if old_rows:
        db.execute("DELETE FROM sessions WHERE created_at < ?", (cutoff,))


def _cleanup_user_sessions(username: str) -> None:
    """Delete all session directories and DB rows belonging to a user."""
    db = get_db()
    rows = db.execute(
        "SELECT uid FROM sessions WHERE username = ?", (username,)
    ).fetchall()
    for row in rows:
        session_path = UPLOAD_ROOT / row["uid"]
        if session_path.exists():
            shutil.rmtree(session_path, ignore_errors=True)
    if rows:
        db.execute("DELETE FROM sessions WHERE username = ?", (username,))


# ── Auth routes ────────────────────────────────────────────────────────────────

@app.route("/login", methods=["GET", "POST"])
@limiter.limit("10 per minute")   # brute-force protection (#4)
def login():
    if current_user.is_authenticated:
        return redirect(url_for("index"))
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = _db_get_user(username)
        if user and check_password_hash(user.password_hash, password):
            login_user(user)
            next_page = request.args.get("next", "").strip()
            if next_page and _is_safe_redirect(next_page):
                return redirect(next_page)
            return redirect(url_for("index"))
        flash("Invalid username or password.")
    return render_template("login.html", next=request.args.get("next", ""))


@app.route("/logout")
@login_required
def logout():
    _cleanup_user_sessions(current_user.username)
    logout_user()
    return redirect(url_for("login"))


@app.route("/profile", methods=["GET", "POST"])
@login_required
def profile():
    if request.method == "POST":
        school = request.form.get("school", "").strip()
        full_name = request.form.get("full_name", "").strip()
        email = request.form.get("email", "").strip()
        new_password = request.form.get("new_password", "").strip()
        confirm = request.form.get("confirm_password", "").strip()

        updates = []
        errors = []

        if school and school != current_user.school:
            get_db().execute(
                "UPDATE users SET school = ? WHERE username = ?",
                (school, current_user.username),
            )
            current_user.school = school
            updates.append(f'School updated to "{school}"')

        if full_name != current_user.full_name:
            get_db().execute(
                "UPDATE users SET full_name = ? WHERE username = ?",
                (full_name, current_user.username),
            )
            current_user.full_name = full_name
            updates.append("Full name updated.")

        if email != current_user.email:
            _parts = email.split("@") if email else []
            if email and (
                len(email) > 254
                or len(_parts) != 2
                or not _parts[0]
                or not _parts[1]
                or "." not in _parts[1]
                or any(c in email for c in (" ", "\t", "\n", "\r"))
            ):
                errors.append("Please enter a valid school email address.")
            else:
                get_db().execute(
                    "UPDATE users SET email = ? WHERE username = ?",
                    (email, current_user.username),
                )
                current_user.email = email
                updates.append("Email updated.")

        if new_password or confirm:
            if new_password != confirm:
                errors.append("Passwords do not match.")
            else:
                err = _validate_password(new_password)
                if err:
                    errors.append(err)
                else:
                    new_hash = generate_password_hash(new_password)
                    get_db().execute(
                        "UPDATE users SET password_hash = ? WHERE username = ?",
                        (new_hash, current_user.username),
                    )
                    current_user.password_hash = new_hash
                    updates.append("Password updated.")

        for msg in errors:
            flash(msg)
        if updates:
            flash("Saved: " + "  ·  ".join(updates))

    return render_template("profile.html")


# ── Admin routes ───────────────────────────────────────────────────────────────

def _admin_required(f):
    """Decorator: requires the logged-in user to have is_admin == True."""
    from functools import wraps

    @wraps(f)
    @login_required
    def decorated(*args, **kwargs):
        if not current_user.is_admin:
            abort(403)
        return f(*args, **kwargs)
    return decorated


@app.route("/admin/users")
@_admin_required
def admin_users():
    users = _db_all_users()
    return render_template("admin_users.html", users=users)


@app.route("/admin/users/add", methods=["POST"])
@_admin_required
def admin_add_user():
    username = request.form.get("username", "").strip()
    password = request.form.get("password", "").strip()
    school = request.form.get("school",   "").strip()
    is_admin = bool(request.form.get("is_admin"))

    errors = []
    if not username:
        errors.append("Username is required.")
    elif not username.replace("_", "").replace("-", "").replace(".", "").isalnum():
        errors.append(
            "Username may only contain letters, numbers, hyphens, underscores and dots.")
    if not password:
        errors.append("Password is required.")
    else:
        err = _validate_password(password)
        if err:
            errors.append(err)

    if not errors:
        try:
            get_db().execute(
                "INSERT INTO users (username, password_hash, school, is_admin) "
                "VALUES (?, ?, ?, ?)",
                (username, generate_password_hash(password),
                 school, int(is_admin)),
            )
            flash(f'User "{username}" created successfully.')
        except sqlite3.IntegrityError:
            flash(f'Username "{username}" is already taken.')
    else:
        for msg in errors:
            flash(msg)

    return redirect(url_for("admin_users"))


@app.route("/admin/users/delete/<username>", methods=["POST"])
@_admin_required
def admin_delete_user(username: str):
    if username == current_user.username:
        flash("You cannot delete your own account.")
        return redirect(url_for("admin_users"))
    get_db().execute("DELETE FROM users WHERE username = ?", (username,))
    flash(f'User "{username}" deleted.')
    return redirect(url_for("admin_users"))


# ── App routes ─────────────────────────────────────────────────────────────────

@app.route("/", methods=["GET"])
@login_required
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
@login_required
def upload():
    # Clean up stale sessions before starting new one (#6)
    _cleanup_old_sessions()

    file = request.files.get("gradefile")
    if not file or not file.filename:
        flash("Please select an Excel file to upload.")
        return redirect(url_for("index"))

    ext = Path(file.filename).suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        flash("Only .xlsx or .xls files are accepted.")
        return redirect(url_for("index"))

    # Enforce size limit before reading fully
    file.seek(0, 2)
    size_mb = file.tell() / (1024 * 1024)
    file.seek(0)
    if size_mb > MAX_UPLOAD_MB:
        flash(f"File too large. Maximum allowed size is {MAX_UPLOAD_MB} MB.")
        return redirect(url_for("index"))

    uid = str(uuid.uuid4())
    session_dir = _session_dir(uid)
    xlsx_path = session_dir / "grades.xlsx"
    file.save(str(xlsx_path))

    try:
        students = load_grades(str(xlsx_path))
    except Exception as exc:
        flash(f"Could not read the spreadsheet: {exc}")
        return redirect(url_for("index"))

    if not students:
        flash("No student data found in the spreadsheet.")
        return redirect(url_for("index"))

    reports_dir = session_dir / "reports"
    reports_dir.mkdir(exist_ok=True)

    school_name = current_user.school
    teacher_name = current_user.full_name
    teacher_email = current_user.email

    # Build a (student, out_path) list; sanitise filenames (#8)
    tasks = []
    for student in students:
        raw = (
            f"{student['last']}_{student['first']}.pdf"
            .replace(" ", "_")
        )
        filename = _safe_filename(raw)
        tasks.append((student, str(reports_dir / filename), school_name,
                      teacher_name, teacher_email))

    # Generate PDFs in parallel (#9)
    errors = []

    def _gen(args):
        student, out_path, sname, tname, temail = args
        build_student_report(student, out_path, school_name=sname,
                             teacher_name=tname, teacher_email=temail)
        return student

    with ThreadPoolExecutor() as executor:
        futures = {executor.submit(_gen, t): t[0] for t in tasks}
        for future in as_completed(futures):
            student = futures[future]
            try:
                future.result()
            except Exception as exc:
                errors.append(
                    f"{student['first']} {student['last']}: {exc}"
                )

    for msg in errors:
        flash(f"Error generating report for {msg}")

    # Record session ownership (#6)
    get_db().execute(
        "INSERT OR REPLACE INTO sessions (uid, username, created_at) VALUES (?, ?, ?)",
        (uid, current_user.username, time.time()),
    )

    return redirect(url_for("results", uid=uid))


@app.route("/results/<uid>")
@login_required
def results(uid: str):
    uid = _safe_uid(uid)
    _check_session_owner(uid)
    reports_dir = _session_dir(uid) / "reports"
    if not reports_dir.exists():
        abort(404)

    rpts = []
    for pdf in sorted(reports_dir.glob("*.pdf")):
        stem = pdf.stem
        parts = stem.split("_", 1)
        if len(parts) == 2:
            display_name = f"{parts[1]} {parts[0]}"
        else:
            display_name = stem
        rpts.append({
            "filename": pdf.name,
            "name":     display_name,
        })

    if not rpts:
        abort(404)

    return render_template("results.html", uid=uid, reports=rpts)


@app.route("/view/<uid>/<path:filename>")
@login_required
def view_report(uid: str, filename: str):
    uid = _safe_uid(uid)
    _check_session_owner(uid)
    safe_name = Path(filename).name   # strip directory components
    if _SAFE_FILENAME_RE.search(safe_name):
        abort(400)
    reports_dir = _session_dir(uid) / "reports"
    pdf_path = reports_dir / safe_name
    if not pdf_path.resolve().is_relative_to(reports_dir.resolve()):
        abort(400)
    if not pdf_path.exists() or pdf_path.suffix.lower() != ".pdf":
        abort(404)
    return send_file(
        str(pdf_path),
        mimetype="application/pdf",
        conditional=True,   # ETag + 304 support (#14)
        max_age=3600,
    )


@app.route("/download/<uid>/<path:filename>")
@login_required
def download_report(uid: str, filename: str):
    uid = _safe_uid(uid)
    _check_session_owner(uid)
    safe_name = Path(filename).name
    if _SAFE_FILENAME_RE.search(safe_name):
        abort(400)
    reports_dir = _session_dir(uid) / "reports"
    pdf_path = reports_dir / safe_name
    if not pdf_path.resolve().is_relative_to(reports_dir.resolve()):
        abort(400)
    if not pdf_path.exists() or pdf_path.suffix.lower() != ".pdf":
        abort(404)
    return send_file(
        str(pdf_path),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=safe_name,
        conditional=True,
        max_age=3600,
    )


@app.route("/download-all/<uid>")
@login_required
def download_all(uid: str):
    uid = _safe_uid(uid)
    _check_session_owner(uid)
    reports_dir = _session_dir(uid) / "reports"
    pdfs = sorted(reports_dir.glob("*.pdf"))
    if not pdfs:
        abort(404)

    # Stream the ZIP to avoid holding the whole file in memory (#11)
    def _generate():
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for pdf in pdfs:
                zf.write(pdf, pdf.name)
        buf.seek(0)
        while True:
            chunk = buf.read(65536)
            if not chunk:
                break
            yield chunk

    return Response(
        stream_with_context(_generate()),
        mimetype="application/zip",
        headers={
            "Content-Disposition": "attachment; filename=grade_reports.zip"
        },
    )


if __name__ == "__main__":
    app.run(debug=False)
