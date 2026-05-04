param($Base = "http://localhost:8000", $GradesFile = "D:\Desktop\ReportGen\Grades.xlsx")

$session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
$pass = 0; $fail = 0

function Check($label, $got, $expect) {
    if ($got -eq $expect) { Write-Host "  PASS  $label" -ForegroundColor Green; $script:pass++ }
    else { Write-Host "  FAIL  $label  (got '$got', expected '$expect')" -ForegroundColor Red; $script:fail++ }
}

function Req($method, $path, $body=$null, $ct=$null) {
    $p = @{
        Uri = "$Base$path"
        Method = $method
        WebSession = $session
        UseBasicParsing = $true
        MaximumRedirection = 0
    }
    if ($body) { $p.Body = $body }
    if ($ct)   { $p.ContentType = $ct }
    try {
        Invoke-WebRequest @p
    } catch [System.Net.WebException] {
        $_.Exception.Response
    }
}

function StatusOf($r) {
    if ($r -eq $null) { return '' }
    if ($r.StatusCode -is [int]) { return [string]$r.StatusCode }
    return [string]([int]$r.StatusCode)
}

Write-Host "`n=== SMOKE TEST ===" -ForegroundColor Cyan

# ── AUTH ──────────────────────────────────────────────────────────────────────
Write-Host "`n[Auth]"

$r = Req GET "/"
Check "GET / unauthenticated -> 302" (StatusOf $r) 302

$r = Req GET "/login"
Check "GET /login -> 200" (StatusOf $r) 200

$r = Req POST "/login" @{username="admin";password="wrongpassword"} "application/x-www-form-urlencoded"
Check "POST /login bad creds -> 200" (StatusOf $r) 200

$r = Req POST "/login" @{username="admin";password="changeme"} "application/x-www-form-urlencoded"
Check "POST /login good creds -> 302" (StatusOf $r) 302

$r = Req GET "/"
Check "GET / authenticated -> 200" (StatusOf $r) 200

$r = Req GET "/profile"
Check "GET /profile -> 200" (StatusOf $r) 200

# ── ADMIN USER MANAGEMENT ─────────────────────────────────────────────────────
Write-Host "`n[Admin – User Management]"

$r = Req GET "/admin/users"
Check "GET /admin/users -> 200" (StatusOf $r) 200

$r = Req POST "/admin/users/add" @{username="smokeuser";password="smokepass1";school="Smoke School"} "application/x-www-form-urlencoded"
Check "POST /admin/users/add -> 302" (StatusOf $r) 302

$r = Req GET "/admin/users"
Check "New user appears in list" ($r.Content -match "smokeuser") $true

# Block self-delete
$r = Req POST "/admin/users/delete/admin" $null $null
Check "POST delete own account blocked -> 302" (StatusOf $r) 302
$r = Req GET "/admin/users"
Check "Admin still in list after blocked delete" ($r.Content -match ">admin<") $true

# Delete the test user
$r = Req POST "/admin/users/delete/smokeuser" $null $null
Check "POST /admin/users/delete/smokeuser -> 302" (StatusOf $r) 302
$r = Req GET "/admin/users"
Check "Deleted user gone from list" ($r.Content -notmatch "smokeuser") $true

# Non-admin user can't access admin page
$r2 = Req POST "/admin/users/add" @{username="noadmin";password="password1";school=""} "application/x-www-form-urlencoded"
$r3 = Req GET "/logout"
$r4 = Req POST "/login" @{username="noadmin";password="password1"} "application/x-www-form-urlencoded"
$r5 = Req GET "/admin/users"
Check "Non-admin GET /admin/users -> 403" (StatusOf $r5) 403
# Re-login as admin
$r6 = Req GET "/logout"
$r7 = Req POST "/login" @{username="admin";password="changeme"} "application/x-www-form-urlencoded"
$r8 = Req POST "/admin/users/delete/noadmin" $null $null

# ── UPLOAD + PDF GENERATION ───────────────────────────────────────────────────
Write-Host "`n[Upload + PDF Generation]"

$bytes = [IO.File]::ReadAllBytes($GradesFile)
$bnd   = "Boundary$([guid]::NewGuid().ToString('N'))"
$pre   = [Text.Encoding]::UTF8.GetBytes("--$bnd`r`nContent-Disposition: form-data; name=`"gradefile`"; filename=`"Grades.xlsx`"`r`nContent-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`r`n`r`n")
$suf   = [Text.Encoding]::UTF8.GetBytes("`r`n--$bnd--`r`n")
$all   = $pre + $bytes + $suf

$r = Invoke-WebRequest -Uri "$Base/upload" -Method POST -WebSession $session -UseBasicParsing -MaximumRedirection 0 -ErrorAction SilentlyContinue -Body $all -ContentType "multipart/form-data; boundary=$bnd"
Check "POST /upload -> 302" (StatusOf $r) 302
$loc = if ($r.Headers -and $r.Headers.Location) { $r.Headers.Location } else { '' }
$uid = ($loc -split "/results/")[1]
Write-Host "       Session UID: $uid"

$r = Req GET "/results/$uid"
Check "GET /results/uid -> 200" (StatusOf $r) 200
$pdfCount = ([regex]::Matches($r.Content, '\.pdf')).Count
Write-Host "       PDFs linked on results page: $pdfCount (expect ≥4)"
Check "At least 4 PDFs on results page" ($pdfCount -ge 4) $true

$m = [regex]::Match($r.Content, '/view/[^""]*?/([^/""]+\.pdf)'); $pdf = $m.Groups[1].Value
Write-Host "       First PDF: $pdf"

# ── PDF DELIVERY ──────────────────────────────────────────────────────────────
Write-Host "`n[PDF Delivery]"

$r = Req GET "/view/$uid/$pdf"
Check "GET /view/uid/pdf -> 200" (StatusOf $r) 200
$ct = if ($r.Headers) { $r.Headers.'Content-Type' } else { '' }
Check "View Content-Type is application/pdf" ($ct -match "application/pdf") $true

$r = Req GET "/download/$uid/$pdf"
Check "GET /download/uid/pdf -> 200" (StatusOf $r) 200
$cd = if ($r.Headers) { $r.Headers.'Content-Disposition' } else { '' }
Check "Download Content-Disposition is attachment" ($cd -match "attachment") $true

$r = Req GET "/download-all/$uid"
Check "GET /download-all/uid -> 200" (StatusOf $r) 200
$ct2 = if ($r.Headers) { $r.Headers.'Content-Type' } else { '' }
Check "ZIP Content-Type correct" ($ct2 -match "application/zip") $true

# ── PATH TRAVERSAL GUARD ──────────────────────────────────────────────────────
Write-Host "`n[Security Checks]"

$r = Req GET "/view/not-a-uuid/test.pdf"
Check "Invalid UID -> 400" (StatusOf $r) 400

$r = Req GET "/view/$uid/../../etc/passwd"
Check "Path traversal attempt -> 404" (StatusOf $r) 404

# ── LOGOUT ────────────────────────────────────────────────────────────────────
Write-Host "`n[Logout]"

$r = Req GET "/logout"
Check "GET /logout -> 302" (StatusOf $r) 302

$r = Req GET "/"
Check "GET / after logout -> 302" (StatusOf $r) 302

# ── SUMMARY ───────────────────────────────────────────────────────────────────
Write-Host "`n=== RESULTS: $pass passed, $fail failed ===" -ForegroundColor $(if ($fail -eq 0) { "Green" } else { "Yellow" })
