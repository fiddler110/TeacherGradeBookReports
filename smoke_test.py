"""
Smoke test for Grade Report Generator.
Uses only stdlib (urllib) — no pip install needed.
CSRF-aware: extracts the token from <meta name="csrf-token"> on every page.
"""
import re
import urllib.request
import urllib.parse
import urllib.error
import http.cookiejar
import uuid
import os
import sys

BASE   = "http://localhost:8000"
GRADES = r"D:\Desktop\ReportGen\Grades.xlsx"

jar    = http.cookiejar.CookieJar()
opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))
opener.addheaders = [("User-Agent", "smoke-test/1.0")]

pass_count = 0
fail_count = 0

# ── CSRF token tracking ────────────────────────────────────────────────────────
_csrf_token: str = ""


def _extract_csrf(html: str) -> str:
    """Pull CSRF token from the meta tag or a hidden form field."""
    m = re.search(r'<meta\s+name="csrf-token"\s+content="([^"]+)"', html)
    if m:
        return m.group(1)
    m = re.search(r'name="csrf_token"\s+value="([^"]+)"', html)
    if m:
        return m.group(1)
    return ""


def _save_csrf(html: str) -> None:
    global _csrf_token
    t = _extract_csrf(html)
    if t:
        _csrf_token = t


def _with_csrf(data: dict) -> dict:
    """Return a copy of data with the current CSRF token injected."""
    d = dict(data)
    if _csrf_token:
        d["csrf_token"] = _csrf_token
    return d


# ── Helpers ────────────────────────────────────────────────────────────────────
def check(label, got, expect):
    global pass_count, fail_count
    ok = str(got) == str(expect)
    sym = "PASS" if ok else "FAIL"
    colour = "\033[32m" if ok else "\033[31m"
    reset  = "\033[0m"
    if ok:
        print(f"  {colour}{sym}{reset}  {label}")
        pass_count += 1
    else:
        print(f"  {colour}{sym}{reset}  {label}  (got {got!r}, expected {expect!r})")
        fail_count += 1


def req(method, path, data=None, content_type="application/x-www-form-urlencoded"):
    url  = BASE + path
    body = None
    if data is not None:
        if isinstance(data, bytes):
            body = data
        else:
            body = urllib.parse.urlencode(data).encode()
    r = urllib.request.Request(url, data=body, method=method)
    if body and content_type:
        r.add_header("Content-Type", content_type)
    try:
        resp    = opener.open(r)
        status  = resp.status
        headers = dict(resp.headers)
        html    = resp.read().decode("utf-8", errors="replace")
        _save_csrf(html)
        return status, headers, html
    except urllib.error.HTTPError as e:
        html = e.read().decode("utf-8", errors="replace")
        _save_csrf(html)
        return e.code, dict(e.headers), html
    except urllib.error.URLError as e:
        print(f"  CONNECTION ERROR: {e}")
        return 0, {}, ""


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        return None


_jar2      = jar
_no_follow = urllib.request.build_opener(
    urllib.request.HTTPCookieProcessor(_jar2),
    _NoRedirect(),
)
_no_follow.addheaders = [("User-Agent", "smoke-test/1.0")]


def req_nofollow(method, path, data=None,
                 content_type="application/x-www-form-urlencoded"):
    url  = BASE + path
    body = None
    if data is not None:
        if isinstance(data, bytes):
            body = data
        else:
            body = urllib.parse.urlencode(data).encode()
    r = urllib.request.Request(url, data=body, method=method)
    if body and content_type:
        r.add_header("Content-Type", content_type)
    try:
        resp    = _no_follow.open(r)
        status  = resp.status
        headers = dict(resp.headers)
        html    = resp.read().decode("utf-8", errors="replace")
        _save_csrf(html)
        return status, headers, html
    except urllib.error.HTTPError as e:
        html = e.read().decode("utf-8", errors="replace")
        _save_csrf(html)
        return e.code, dict(e.headers), html
    except urllib.error.URLError as e:
        print(f"  CONNECTION ERROR: {e}")
        return 0, {}, ""


def fresh_login_csrf() -> None:
    """GET /login to obtain a fresh CSRF token for the unauthenticated session."""
    req("GET", "/login")


# ── Tests ──────────────────────────────────────────────────────────────────────
print("\n=== SMOKE TEST ===")

# ── Auth ───────────────────────────────────────────────────────────────────────
print("\n[Auth]")

status, headers, body = req_nofollow("GET", "/")
check("GET / unauthenticated -> 302", status, 302)

# GET /login to obtain CSRF for the unauthenticated session
status, headers, body = req("GET", "/login")
check("GET /login -> 200", status, 200)

status, headers, body = req(
    "POST", "/login", _with_csrf({"username": "admin", "password": "wrongpass"}))
check("POST /login bad creds -> 200", status, 200)
check("Bad creds shows error message",
      "Invalid username or password" in body, True)

# Refresh CSRF after the 200 (the failed login response embeds a new token)
fresh_login_csrf()

status, headers, body = req_nofollow(
    "POST", "/login", _with_csrf({"username": "admin", "password": "SDMac14!"}))
check("POST /login good creds -> 302", status, 302)

# GET / to pick up the authenticated-session CSRF token
status, headers, body = req("GET", "/")
check("GET / authenticated -> 200", status, 200)

status, headers, body = req("GET", "/profile")
check("GET /profile -> 200", status, 200)
check("Profile shows school field", "school" in body.lower(), True)

# ── Admin: User Management ─────────────────────────────────────────────────────
print("\n[Admin – User Management]")

status, headers, body = req("GET", "/admin/users")
check("GET /admin/users -> 200", status, 200)

status, headers, body = req_nofollow(
    "POST", "/admin/users/add",
    _with_csrf({"username": "smokeuser", "password": "smokepass1",
                "school": "Smoke School"}))
check("POST /admin/users/add -> 302", status, 302)

status, headers, body = req("GET", "/admin/users")
check("New user appears in list", "smokeuser" in body, True)

# Self-delete blocked
status, headers, body = req_nofollow(
    "POST", "/admin/users/delete/admin", _with_csrf({}))
check("POST delete own account blocked -> 302", status, 302)
status, headers, body = req("GET", "/admin/users")
check("Admin still in list", "badge-admin" in body, True)

# Delete test user
status, headers, body = req_nofollow(
    "POST", "/admin/users/delete/smokeuser", _with_csrf({}))
check("POST delete smokeuser -> 302", status, 302)
status, headers, body = req("GET", "/admin/users")
check("Deleted user gone from list",
      "/admin/users/delete/smokeuser" not in body, True)

# Non-admin user can't access admin route
req_nofollow(
    "POST", "/admin/users/add",
    _with_csrf({"username": "noadmin", "password": "noadmin1",
                "school": ""}))

# Logout and log in as noadmin
req_nofollow("GET", "/logout")
fresh_login_csrf()   # new unauthenticated session → new CSRF
req_nofollow(
    "POST", "/login",
    _with_csrf({"username": "noadmin", "password": "noadmin1"}))

# GET any page to refresh CSRF for the noadmin session
req("GET", "/")

status, headers, body = req("GET", "/admin/users")
check("Non-admin GET /admin/users -> 403", status, 403)

# Re-login as admin
req_nofollow("GET", "/logout")
fresh_login_csrf()   # new unauthenticated session → new CSRF
req_nofollow(
    "POST", "/login",
    _with_csrf({"username": "admin", "password": "SDMac14!"}))

# GET admin page to refresh CSRF for admin session
req("GET", "/admin/users")

req_nofollow(
    "POST", "/admin/users/delete/noadmin", _with_csrf({}))

# ── Upload ─────────────────────────────────────────────────────────────────────
print("\n[Upload + PDF Generation]")

# GET index to ensure we have an up-to-date CSRF for the admin session
req("GET", "/")

with open(GRADES, "rb") as f:
    file_data = f.read()

boundary   = uuid.uuid4().hex
csrf_part  = (
    f"--{boundary}\r\n"
    f'Content-Disposition: form-data; name="csrf_token"\r\n'
    f"\r\n"
    f"{_csrf_token}\r\n"
).encode()
file_part  = (
    f"--{boundary}\r\n"
    f'Content-Disposition: form-data; name="gradefile"; filename="Grades.xlsx"\r\n'
    f"Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n"
    f"\r\n"
).encode() + file_data + f"\r\n--{boundary}--\r\n".encode()
body_parts = csrf_part + file_part

status, headers, body = req_nofollow(
    "POST", "/upload",
    data=body_parts,
    content_type=f"multipart/form-data; boundary={boundary}",
)
check("POST /upload -> 302", status, 302)

loc = headers.get("Location", "")
uid = loc.split("/results/")[-1] if "/results/" in loc else ""
print(f"       Session UID: {uid}")
check("Redirect has valid UID", bool(uid), True)

status, headers, body = req("GET", f"/results/{uid}")
check("GET /results/uid -> 200", status, 200)

pdf_names = []
for m in re.finditer(r'/view/[^"\']+?/([^/"\']+\.pdf)', body):
    pdf_names.append(m.group(1))
pdf_names = list(dict.fromkeys(pdf_names))   # dedupe
print(f"       PDFs found: {len(pdf_names)} — {pdf_names}")
check("4 PDFs generated (one per student)", len(pdf_names), 4)

# ── PDF Delivery ───────────────────────────────────────────────────────────────
print("\n[PDF Delivery]")

if pdf_names:
    pdf = pdf_names[0]
    status, headers, body = req("GET", f"/view/{uid}/{pdf}")
    check("GET /view/uid/pdf -> 200", status, 200)
    ct = headers.get("Content-Type", "")
    check("View Content-Type is application/pdf", "application/pdf" in ct, True)

    status, headers, body = req("GET", f"/download/{uid}/{pdf}")
    check("GET /download/uid/pdf -> 200", status, 200)
    cd = headers.get("Content-Disposition", "")
    check("Download has attachment disposition", "attachment" in cd, True)

status, headers, body = req("GET", f"/download-all/{uid}")
check("GET /download-all/uid -> 200", status, 200)
ct = headers.get("Content-Type", "")
check("ZIP Content-Type correct", "application/zip" in ct, True)

# ── Security ───────────────────────────────────────────────────────────────────
print("\n[Security]")

status, headers, body = req("GET", "/view/not-a-uuid/test.pdf")
check("Invalid UID -> 400", status, 400)

status, headers, body = req("GET", f"/view/{uid}/../../etc/passwd")
check("Path traversal attempt -> 404", status, 404)

# Session ownership: log in as noadmin and try to access admin's session
req("GET", "/admin/users")
req_nofollow(
    "POST", "/admin/users/add",
    _with_csrf({"username": "noadmin2", "password": "noadmin22",
                "school": ""}))
req_nofollow("GET", "/logout")
fresh_login_csrf()
req_nofollow(
    "POST", "/login",
    _with_csrf({"username": "noadmin2", "password": "noadmin22"}))
req("GET", "/")

status, headers, body = req("GET", f"/results/{uid}")
check("noadmin2 cannot view admin session -> 403", status, 403)

# Restore admin session
req_nofollow("GET", "/logout")
fresh_login_csrf()
req_nofollow(
    "POST", "/login",
    _with_csrf({"username": "admin", "password": "SDMac14!"}))
req("GET", "/admin/users")
req_nofollow("POST", "/admin/users/delete/noadmin2", _with_csrf({}))

# ── Logout ─────────────────────────────────────────────────────────────────────
print("\n[Logout]")

status, headers, body = req_nofollow("GET", "/logout")
check("GET /logout -> 302", status, 302)

status, headers, body = req_nofollow("GET", "/")
check("GET / after logout -> 302", status, 302)

# ── Summary ────────────────────────────────────────────────────────────────────
total  = pass_count + fail_count
colour = "\033[32m" if fail_count == 0 else "\033[33m"
print(f"\n=== RESULTS: {colour}{pass_count}/{total} passed\033[0m ===\n")
sys.exit(0 if fail_count == 0 else 1)
