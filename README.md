# Grade Report Generator

A self-hosted web application for generating professional per-student PDF grade reports from a standard Excel spreadsheet. Built for teachers and school administrators — upload a gradebook once and download a polished, print-ready PDF for every student in seconds.

---

## Table of Contents

- [Features](#features)
- [Screenshots / Report Layout](#report-layout)
- [Prerequisites](#prerequisites)
- [Quick Start (Docker)](#quick-start-docker)
- [Environment Variables](#environment-variables)
- [Spreadsheet Format](#spreadsheet-format)
- [Grade Scale](#grade-scale)
- [User Management](#user-management)
- [Profile Settings](#profile-settings)
- [Running Without Docker](#running-without-docker)
- [CLI Usage](#cli-usage)
- [Project Structure](#project-structure)
- [Security](#security)
- [Smoke Test](#smoke-test)
- [Troubleshooting](#troubleshooting)

---

## Features

- **Upload → Generate → Download** — drag-and-drop an `.xlsx` or `.xls` file and receive one PDF per student, or a single ZIP containing all reports.
- **Professional PDF layout** — school name in the header/footer, per-module assignment tables with colour-coded score bars, an overall average summary card, and a grade scale legend.
- **Colour-coded grading** — green (B− and above), yellow (C range), orange (D range), red (F / below 50 %).
- **Personalised contact footer** — each report closes with a contact line referencing the teacher's name and school email address.
- **Multi-user authentication** — login-protected, with admin and standard user roles.
- **Admin panel** — create, list, and delete user accounts without touching the database directly.
- **Rate limiting** — login endpoint is limited to 10 requests per minute to block brute-force attempts.
- **CSRF protection** — all state-changing requests are protected by Flask-WTF tokens.
- **Parallel PDF generation** — all reports for an upload are built concurrently via `ThreadPoolExecutor`.
- **Session ownership isolation** — users can only access their own generated reports.
- **Automatic temp file cleanup** — all uploaded spreadsheets and generated PDFs are deleted from the server the moment a user logs out. Any sessions not explicitly closed are swept after 24 hours.
- **Hardened Docker image** — multi-stage build, non-root user, read-only root filesystem, all Linux capabilities dropped.

---

## Report Layout

Each PDF contains:

| Section                   | Details                                                                       |
| ------------------------- | ----------------------------------------------------------------------------- |
| **Header banner**         | School name and report period on every page                                   |
| **Student identity card** | Full name and report date                                                     |
| **Overall summary**       | Average percentage, letter grade, and submitted/total assignment count        |
| **Grade scale legend**    | Colour-coded reference strip (A+–D−, F)                                       |
| **Per-module tables**     | Assignment name · Score % · Letter grade · Visual bar · Module average footer |
| **Closing note**          | Auto-generated contact line with teacher name and email                       |
| **Footer**                | Generation date, school name, page number                                     |

---

## Prerequisites

- [Docker Desktop](https://www.docker.com/products/docker-desktop/) (recommended), **or**
- Python 3.12+ with pip (for running locally without Docker)

---

## Quick Start (Docker)

### 1. Create an environment file

Copy the example below into a file named `.env` in the project root and fill in your values.

```env
# Initial admin account (created on first boot only)
ADMIN_USERNAME=admin
ADMIN_PASSWORD=YourStrongPassword1
ADMIN_SCHOOL=Westfield Academy

# Must be a long random string — used to sign session cookies
SECRET_KEY=replace-with-a-long-random-secret

# Set to false only when running over plain HTTP (e.g. localhost dev)
COOKIE_SECURE=true

# Optional: increase for higher upload concurrency
GUNICORN_WORKERS=2
```

> **Important:** never commit `.env` to source control. It is listed in `.dockerignore` and should be in `.gitignore`.

### 2. Build and start

```bash
docker compose up --build
```

The application will be available at **http://localhost:8000**.

### 3. First login

Use the `ADMIN_USERNAME` / `ADMIN_PASSWORD` values from your `.env` file. On a fresh volume the admin account is created automatically on startup.

### 4. Stopping

```bash
docker compose down
```

The SQLite database is persisted in the named volume `reportgen_grade-reports-data` and survives container restarts.

---

## Environment Variables

| Variable           | Default             | Description                                                |
| ------------------ | ------------------- | ---------------------------------------------------------- |
| `SECRET_KEY`       | _(required)_        | Flask session signing key. Use a long random string.       |
| `ADMIN_USERNAME`   | `admin`             | Username for the bootstrap admin account.                  |
| `ADMIN_PASSWORD`   | `changeme`          | Password for the bootstrap admin account. **Change this.** |
| `ADMIN_SCHOOL`     | `Westfield Academy` | School name pre-set on the admin account.                  |
| `COOKIE_SECURE`    | `true`              | Set `false` when running without HTTPS (local dev only).   |
| `GUNICORN_WORKERS` | `2`                 | Number of Gunicorn worker processes.                       |
| `DB_PATH`          | `/data/users.db`    | Path to the SQLite user database. Mount a volume here.     |
| `UPLOAD_ROOT`      | `/tmp/reportgen`    | Directory for temporary upload sessions.                   |

---

## Spreadsheet Format

The application reads `.xlsx` / `.xls` files with the following column layout:

| Col      | Content                                      |
| -------- | -------------------------------------------- |
| A        | Student ID                                   |
| B        | Last name                                    |
| C        | First name                                   |
| D        | _(ignored — typically blank or a separator)_ |
| E onward | One column per assignment                    |

**Row 1** must be a header row. Assignment column names must start with a module number followed by a period and a space, for example:

```
1.1 Discussion Post
1.2 Quiz
2.1 Lab Report
2.2 Midterm Essay
```

Scores must be decimal values between `0.0` and `1.0` (e.g. `0.87` = 87 %). A blank cell or a value of `0` is treated as **Not Submitted**.

Assignments that share the same leading number (e.g. all `1.*` columns) are grouped into the same module section in the PDF.

---

## Grade Scale

| Letter | Range      |
| ------ | ---------- |
| A+     | 95 – 100 % |
| A      | 87 – 94 %  |
| A−     | 80 – 86 %  |
| B+     | 77 – 79 %  |
| B      | 73 – 76 %  |
| B−     | 70 – 72 %  |
| C+     | 67 – 69 %  |
| C      | 63 – 66 %  |
| C−     | 60 – 62 %  |
| D+     | 57 – 59 %  |
| D      | 53 – 56 %  |
| D−     | 50 – 52 %  |
| F      | 0 – 49 %   |

Colour coding in reports and visual bars:

| Colour    | Range                 |
| --------- | --------------------- |
| 🟢 Green  | B− and above (≥ 70 %) |
| 🟡 Yellow | C range (60 – 69 %)   |
| 🟠 Orange | D range (50 – 59 %)   |
| 🔴 Red    | F (below 50 %)        |

---

## User Management

### Admin panel

Navigate to **Users** in the top navigation bar (admin accounts only). From there you can:

- **Add** a new user — provide a username, password, school name, and optionally grant admin privileges.
- **Delete** an existing user — you cannot delete your own account.

### Password requirements

Passwords must be **at least 8 characters** and contain at least one letter and one digit.

### Roles

| Role         | Capabilities                                                           |
| ------------ | ---------------------------------------------------------------------- |
| **Admin**    | Everything below, plus access to the Users panel                       |
| **Standard** | Upload spreadsheets, generate and download reports, manage own profile |

---

## Profile Settings

Each user can update their profile at **Profile → Settings**:

| Field            | Used for                                                |
| ---------------- | ------------------------------------------------------- |
| **Full Name**    | Appears as the teacher name in the PDF contact footer   |
| **School Email** | Appears as the contact email in the PDF contact footer  |
| **School Name**  | Printed in the header and footer of every generated PDF |
| **Password**     | Change your login password                              |

The contact footer in the generated PDF will read:

- **Name and email set:** _"Please contact [Student]'s teacher [Teacher Name] at [email] if you have any questions or concerns."_
- **Name only:** _"Please contact [Student]'s teacher [Teacher Name] if you have any questions or concerns."_
- **Neither set:** _"Please contact [Student]'s teacher if you have any questions."_

---

## Running Without Docker

### 1. Create and activate a virtual environment

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Set required environment variables

```bash
# Windows PowerShell
$env:SECRET_KEY = "dev-secret-not-for-production"
$env:COOKIE_SECURE = "false"
$env:DB_PATH = "users.db"

# macOS / Linux
export SECRET_KEY="dev-secret-not-for-production"
export COOKIE_SECURE="false"
export DB_PATH="users.db"
```

### 4. Run the development server

```bash
python app.py
```

The application starts on **http://localhost:5000** (Flask's default debug port).

> For a production-like local run use Gunicorn directly:
>
> ```bash
> gunicorn --bind 0.0.0.0:8000 --workers 2 app:app
> ```

---

## CLI Usage

Reports can also be generated directly from the command line without the web interface. Place your `Grades.xlsx` in the project root and run:

```bash
python generate_reports.py
```

PDFs are written to the `reports/` subdirectory. The school name defaults to `Westfield Academy`; teacher name and email are not set in CLI mode (the generic fallback footer is used).

---

## Project Structure

```
ReportGen/
├── app.py                  # Flask web application
├── generate_reports.py     # PDF generation engine (also a standalone CLI)
├── requirements.txt        # Python dependencies
├── Dockerfile              # Multi-stage container build
├── docker-compose.yml      # Compose service definition with security hardening
├── .dockerignore           # Files excluded from the Docker build context
├── smoke_test.py           # Automated smoke tests (stdlib only, no pip required)
├── smoke_test.ps1          # PowerShell wrapper for the smoke test
├── static/
│   ├── upload.js           # Front-end upload UX
│   └── vendor/             # Bootstrap 5 CSS/JS and Bootstrap Icons (bundled)
├── templates/
│   ├── index.html          # Upload page
│   ├── login.html          # Login page
│   ├── profile.html        # Profile / settings page
│   ├── results.html        # Generated reports list page
│   └── admin_users.html    # Admin user management page
└── reports/                # CLI output directory (created on first CLI run)
```

---

## Security

The following controls are implemented:

| Control                     | Implementation                                                                                                  |
| --------------------------- | --------------------------------------------------------------------------------------------------------------- |
| CSRF protection             | Flask-WTF tokens on all POST forms; token embedded in HTML meta tag for JS requests                             |
| Session cookie hardening    | `HttpOnly`, `SameSite=Lax`, `Secure` (configurable)                                                             |
| Brute-force protection      | Flask-Limiter — login endpoint capped at 10 requests/minute per IP                                              |
| Password complexity         | Minimum 8 characters, at least one letter and one digit                                                         |
| Open redirect prevention    | Redirect targets validated against the current host before following                                            |
| Path traversal prevention   | Upload session IDs validated as UUIDs; filenames sanitised to `[A-Za-z0-9_\-.]`                                 |
| Session ownership           | Users can only access reports from their own upload sessions                                                    |
| Temp file cleanup on logout | All upload session directories and generated PDFs are removed when the user logs out (`_cleanup_user_sessions`) |
| Stale session cleanup       | Sessions and temp files older than 24 hours are deleted on each new upload as a fallback                        |
| Security response headers   | `X-Frame-Options: DENY`, `X-Content-Type-Options: nosniff`, `Referrer-Policy`, `Content-Security-Policy`        |
| File upload validation      | Extension allowlist (`.xlsx`, `.xls`), 10 MB size cap enforced before reading                                   |
| Non-root container          | Process runs as `appuser` (uid 1001); `USER` directive in Dockerfile                                            |
| Read-only root filesystem   | `read_only: true` in Compose; `/tmp` provided via tmpfs                                                         |
| No new privileges           | `security_opt: no-new-privileges:true` in Compose                                                               |
| All capabilities dropped    | `cap_drop: ALL` in Compose                                                                                      |
| Multi-stage Docker build    | Build tools, pip, and dev headers are absent from the final runtime image                                       |

---

## Smoke Test

A self-contained smoke test exercises login, file upload, report generation, individual PDF download, ZIP download, and logout — with no external dependencies beyond the standard library.

**Prerequisites:** the application must be running on `http://localhost:8000` and a `Grades.xlsx` file must exist at `D:\Desktop\ReportGen\Grades.xlsx`.

```bash
# Python directly
python smoke_test.py

# PowerShell wrapper
.\smoke_test.ps1
```

---

## Troubleshooting

### `SECRET_KEY environment variable is required`

The application will not start without a `SECRET_KEY`. Set it in your `.env` file (Docker) or as an environment variable (local).

### `attempt to write a readonly database`

This occurs when the `/data` volume was previously owned by root and the container now runs as `appuser`. Remove the old volume and let Docker re-create it:

```bash
docker compose down
docker volume rm reportgen_grade-reports-data
docker compose up
```

### Uploaded file is rejected

- Ensure the file is `.xlsx` or `.xls` format.
- Confirm the file size is under 10 MB.
- Verify row 1 is a header row and that assignment columns start at column E (index 4).
- Scores must be numeric decimal values (e.g. `0.85` for 85 %) — not formatted as `85%` text.

### No students found in spreadsheet

Column A (Student ID) is used as the row sentinel. If that column is blank the row is skipped. Ensure every student row has a value in column A.

### Reports generate but PDFs are blank / corrupted

This is typically caused by a `reportlab` or `pillow` version mismatch. Rebuild the container image to pick up a clean dependency install:

```bash
docker compose up --build
```

### Login rate limit hit (`429 Too Many Requests`)

The login endpoint allows 10 attempts per minute per IP. Wait 60 seconds and try again.
