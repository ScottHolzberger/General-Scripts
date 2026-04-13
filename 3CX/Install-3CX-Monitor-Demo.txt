<#
.SYNOPSIS
  Multi-3CX local demo installer (Docker Postgres + Grafana + Python Collector)

.FEATURES
  - Stored PBX secrets (DPAPI, Export-Clixml) auto-used if present
  - Always prompts to add PBXs unless -NonInteractive
  - -ListPbxs / -RemovePbx / -CompactIndices / -NonInteractive
  - Per-PBX alert email recipient stored at add time
  - Email alert when ActiveCalls >= LicensedCalls - 1
  - Retention rollups for active_calls:
      * 0-30 days: keep all
      * 31-60 days: max per hour
      * 61-120 days: max per day
      * 121+ days: max per week (Sunday start)
  - Grafana provisioning (BOM-free) + Postgres datasource default DB fix
    (Postgres datasource requires a default DB; Grafana v12 provisioning can require jsonData.database) 

.PARAMETERS
  -ResetDb          wipes pgdata
  -NoGrafana        don't start/provision grafana
  -ListPbxs         list stored PBXs and exit
  -RemovePbx <n>    remove PBXIndex n from stored secrets
  -CompactIndices   renumber PBXIndex to 1..N (recommended after removals)
  -NonInteractive   no prompts (requires stored secrets)
#>

[CmdletBinding()]
param(
  [string]$ProjectDir = "C:\temp\msp-3cx-monitor-demo",
  [switch]$ResetDb,
  [switch]$NoGrafana,
  [int]$ActiveCallsSeconds = 15,
  [int]$LicenseCheckHours = 24,
  [int]$RetentionRunEveryHours = 24,
  [int]$MaxCollectorFailures = 3,
  [int]$PostgresWaitSeconds = 240,
  [int]$PoolMaxSize = 10,

  [switch]$ListPbxs,
  [int]$RemovePbx = 0,
  [switch]$CompactIndices,
  [switch]$NonInteractive
)

$ErrorActionPreference = 'Stop'

function Info([string]$m){ Write-Host "[INFO] $m" -ForegroundColor Cyan }
function Warn([string]$m){ Write-Host "[WARN] $m" -ForegroundColor Yellow }
function Err ([string]$m){ Write-Host "[ERR ] $m" -ForegroundColor Red }
function IsNullOrWhiteSpace([string]$s){ return [string]::IsNullOrWhiteSpace($s) }

function HasCmd([string]$n){ return (Get-Command $n -ErrorAction SilentlyContinue) -ne $null }

function Invoke-CmdCapture([string]$CommandLine){
  $out = cmd.exe /c "$CommandLine 2>&1"
  return @{ ExitCode = $LASTEXITCODE; Output = ($out | Out-String) }
}

function Ensure-DockerOnline {
  try { docker info *> $null; if ($LASTEXITCODE -eq 0) { return } } catch {}
  throw "Docker engine not running. Start Docker Desktop then re-run."
}

function Ensure-Python {
  if (HasCmd "py") {
    $out = & py -3 --version 2>&1
    if ($LASTEXITCODE -eq 0 -and $out -match '^Python\s+\d+\.\d+') { return @{ kind="py"; cmd=@("py","-3") } }
  }
  if (HasCmd "python") {
    $out = & python --version 2>&1
    if ($LASTEXITCODE -eq 0 -and $out -match '^Python\s+\d+\.\d+') { return @{ kind="python"; cmd=@("python") } }
  }
  throw "Python 3.11+ not found."
}

# BOM-free writer (critical for Grafana provisioning on Windows PowerShell 5.1)
function Set-ContentUtf8NoBom {
  param(
    [Parameter(Mandatory=$true)][string]$Path,
    [Parameter(Mandatory=$true)][string]$Value
  )
  $enc = New-Object System.Text.UTF8Encoding($false)
  $dir = Split-Path -Parent $Path
  if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
  [System.IO.File]::WriteAllText($Path, $Value, $enc)
}

function SecureStringToPlain([SecureString]$sec){
  $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
  $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
  [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
  return $plain
}

function Ensure-Dirs([string]$Dir){
  New-Item -ItemType Directory -Force -Path $Dir | Out-Null
  New-Item -ItemType Directory -Force -Path (Join-Path $Dir "logs") | Out-Null
  New-Item -ItemType Directory -Force -Path (Join-Path $Dir "secrets") | Out-Null
  if (-not $NoGrafana) {
    New-Item -ItemType Directory -Force -Path (Join-Path $Dir "grafana\provisioning\datasources") | Out-Null
    New-Item -ItemType Directory -Force -Path (Join-Path $Dir "grafana\provisioning\dashboards") | Out-Null
    New-Item -ItemType Directory -Force -Path (Join-Path $Dir "grafana\dashboards") | Out-Null
  }
}

# -------------------- SECRETS --------------------
$PbxSecretsPath  = Join-Path $ProjectDir "secrets\pbx-secrets.clixml"
$SmtpSecretsPath = Join-Path $ProjectDir "secrets\smtp-secret.clixml"

function Load-Secrets([string]$Path){
  if (Test-Path $Path) {
    try { return @((Import-Clixml $Path)) } catch { return @() }
  }
  return @()
}

function Save-Secrets([string]$Path, $Secrets){
  @($Secrets) | Export-Clixml -Path $Path
}

function Next-PbxIndexFromSecrets($Secrets){
  $Secrets = @($Secrets)
  if ($Secrets.Count -eq 0) { return 1 }
  $max = ($Secrets | ForEach-Object { [int]$_.PBXIndex } | Measure-Object -Maximum).Maximum
  return ([int]$max + 1)
}

function Apply-PbxSecretsToEnv($Secrets){
  foreach ($s in (@($Secrets) | Sort-Object PBXIndex)) {
    $i = [int]$s.PBXIndex
    [System.Environment]::SetEnvironmentVariable("PBX${i}_USERNAME", $s.Username, "Process")
    [System.Environment]::SetEnvironmentVariable("PBX${i}_PASSWORD", (SecureStringToPlain $s.Password), "Process")
    [System.Environment]::SetEnvironmentVariable("PBX${i}_MFA",      $s.MFA, "Process")
    [System.Environment]::SetEnvironmentVariable("PBX${i}_ALERT_EMAIL", $s.AlertEmail, "Process")
  }
}

function Show-Pbxs($Secrets){
  if (-not $Secrets -or $Secrets.Count -eq 0) {
    Warn "No stored PBXs found at: $PbxSecretsPath"
    return
  }
  $Secrets | Sort-Object PBXIndex | Select-Object `
    PBXIndex, Name, Tenant, BaseUrl, VerifyTls, Username, AlertEmail, AddedAt |
    Format-Table -AutoSize
}

function Compact-Indices($Secrets){
  $ordered = @($Secrets | Sort-Object PBXIndex)
  $new = @()
  $idx = 1
  foreach ($s in $ordered) {
    $new += [PSCustomObject]@{
      PBXIndex   = $idx
      Name       = $s.Name
      Tenant     = $s.Tenant
      BaseUrl    = $s.BaseUrl
      VerifyTls  = $s.VerifyTls
      Username   = $s.Username
      Password   = $s.Password
      MFA        = $s.MFA
      AlertEmail = $s.AlertEmail
      AddedAt    = $s.AddedAt
    }
    $idx++
  }
  return $new
}

function Remove-PbxByIndex($Secrets, [int]$Index){
  $Secrets = @($Secrets)
  if ($Index -le 0) { throw "-RemovePbx must be > 0" }
  if (-not ($Secrets | Where-Object { [int]$_.PBXIndex -eq $Index })) {
    throw "PBXIndex $Index not found."
  }
  return @($Secrets | Where-Object { [int]$_.PBXIndex -ne $Index })
}

function Prompt-SmtpConfig {
  $smtp = (Load-Secrets $SmtpSecretsPath | Select-Object -First 1)

  if ($smtp -and $NonInteractive) { return $smtp }
  if (-not $smtp -and $NonInteractive) { return $null }

  if ($smtp) { Info "SMTP config found (stored)." } else { Warn "No SMTP config stored yet (alerts will log-only unless configured)." }

  $ans = Read-Host "Configure/Update SMTP now? (y/n) [n]"
  if (IsNullOrWhiteSpace $ans) { $ans = "n" }
  if ($ans.ToLower() -ne "y") { return $smtp }

  $host = Read-Host "SMTP Host"
  $port = Read-Host "SMTP Port [587]"
  if (IsNullOrWhiteSpace $port) { $port = "587" }
  $from = Read-Host "From address (alerts sender)"
  $useTls = Read-Host "Use STARTTLS? (y/n) [y]"
  if (IsNullOrWhiteSpace $useTls) { $useTls = "y" }
  $auth = Read-Host "SMTP Auth? (y/n) [n]"
  if (IsNullOrWhiteSpace $auth) { $auth = "n" }

  $smtpUser = $null
  $smtpPass = $null
  if ($auth.ToLower() -eq "y") {
    $smtpUser = Read-Host "SMTP Username"
    $smtpPass = Read-Host "SMTP Password (hidden; stored encrypted)" -AsSecureString
  }

  $obj = [PSCustomObject]@{
    SmtpHost     = $host
    SmtpPort     = [int]$port
    FromAddress  = $from
    UseStartTls  = ($useTls.ToLower() -eq "y")
    SmtpUser     = $smtpUser
    SmtpPassword = $smtpPass
    UpdatedAt    = (Get-Date)
  }

  Save-Secrets $SmtpSecretsPath @($obj)
  Info "Saved SMTP settings to $SmtpSecretsPath"
  return $obj
}

function Apply-SmtpToEnv($smtp){
  if (-not $smtp) { return }
  [System.Environment]::SetEnvironmentVariable("SMTP_HOST", $smtp.SmtpHost, "Process")
  [System.Environment]::SetEnvironmentVariable("SMTP_PORT", "$($smtp.SmtpPort)", "Process")
  [System.Environment]::SetEnvironmentVariable("SMTP_FROM", $smtp.FromAddress, "Process")
  [System.Environment]::SetEnvironmentVariable("SMTP_STARTTLS", "$($smtp.UseStartTls)".ToLower(), "Process")
  if ($smtp.SmtpUser) { [System.Environment]::SetEnvironmentVariable("SMTP_USERNAME", $smtp.SmtpUser, "Process") }
  if ($smtp.SmtpPassword) { [System.Environment]::SetEnvironmentVariable("SMTP_PASSWORD", (SecureStringToPlain $smtp.SmtpPassword), "Process") }
}

function Prompt-AddPbxLoop($secrets){
  $secrets = @($secrets)

  if ($secrets.Count -gt 0) {
    Info "Stored PBX secrets detected ($($secrets.Count)). Using them automatically."
    Apply-PbxSecretsToEnv $secrets
  } else {
    Warn "No stored PBX secrets found yet."
  }

  if ($NonInteractive) {
    Info "NonInteractive mode: skipping PBX prompts."
    return @($secrets)
  }

  while ($true) {
    $ans = Read-Host "Add another PBX entry? (y/n) [n]"
    if (IsNullOrWhiteSpace $ans) { $ans = "n" }
    if ($ans.ToLower() -ne "y") { break }

    $idx = Next-PbxIndexFromSecrets $secrets
    Info "Adding PBX #$idx"

    $name = Read-Host "  Name"
    if (IsNullOrWhiteSpace $name) { $name = "PBX-$idx" }

    $tenant = Read-Host "  Tenant"
    if (IsNullOrWhiteSpace $tenant) { $tenant = "Tenant-$idx" }

    $base = Read-Host "  Base URL (NO auto port fix)"
    $tls  = Read-Host "  Verify TLS cert? (y/n) [y]"
    if (IsNullOrWhiteSpace $tls) { $tls = "y" }
    $verifyTls = ($tls.ToLower() -eq "y")

    $username = Read-Host "  Login Username"
    $pwSec    = Read-Host "  Login Password (hidden; stored encrypted)" -AsSecureString
    $mfa      = Read-Host "  MFA/SecurityCode (optional). Enter for none"
    $alertEmail = Read-Host "  Alert recipient email for this PBX"

    $newItem = [PSCustomObject]@{
      PBXIndex   = $idx
      Name       = $name
      Tenant     = $tenant
      BaseUrl    = $base
      VerifyTls  = $verifyTls
      Username   = $username
      Password   = $pwSec
      MFA        = $mfa
      AlertEmail = $alertEmail
      AddedAt    = (Get-Date)
    }

    $secrets += ,$newItem
    Save-Secrets $PbxSecretsPath $secrets
    Apply-PbxSecretsToEnv $secrets
    Info "Saved PBX #$idx. Total stored: $($secrets.Count)"
  }

  return @($secrets)
}

function Build-PbxListFromSecrets($Secrets){
  $list = @()
  foreach ($s in (@($Secrets) | Sort-Object PBXIndex)) {
    $i = [int]$s.PBXIndex
    $list += [PSCustomObject]@{
      name         = $s.Name
      tenant       = $s.Tenant
      base_url     = $s.BaseUrl
      verify_tls   = $s.VerifyTls
      alert_email  = $s.AlertEmail
      username_env = "PBX${i}_USERNAME"
      password_env = "PBX${i}_PASSWORD"
      mfa_env      = "PBX${i}_MFA"
    }
  }
  return $list
}

# -------------------- FILE GENERATION --------------------
function Write-Files {
  param(
    [string]$Dir,
    [array]$PBXs
  )

  Ensure-Dirs $Dir
  Set-Location $Dir
  Info "Project directory: $Dir"

  if ($ResetDb -and (Test-Path "$Dir\pgdata")) {
    Warn "ResetDb: removing pgdata..."
    Remove-Item -Recurse -Force "$Dir\pgdata"
  }

  # docker-compose.yml
  $compose = if ($NoGrafana) {
@"
services:
  postgres:
    image: postgres:16
    environment:
      POSTGRES_USER: 3cxmon
      POSTGRES_PASSWORD: 3cxmon
      POSTGRES_DB: 3cxmon
    ports:
      - "5432:5432"
    volumes:
      - ./pgdata:/var/lib/postgresql/data
"@
  } else {
@"
services:
  postgres:
    image: postgres:16
    environment:
      POSTGRES_USER: 3cxmon
      POSTGRES_PASSWORD: 3cxmon
      POSTGRES_DB: 3cxmon
    ports:
      - "5432:5432"
    volumes:
      - ./pgdata:/var/lib/postgresql/data

  grafana:
    image: grafana/grafana:latest
    environment:
      GF_SECURITY_ADMIN_USER: admin
      GF_SECURITY_ADMIN_PASSWORD: admin
      GF_USERS_ALLOW_SIGN_UP: "false"
    ports:
      - "3000:3000"
    depends_on:
      - postgres
    volumes:
      - ./grafana/provisioning:/etc/grafana/provisioning
      - ./grafana/dashboards:/var/lib/grafana/dashboards
"@
  }
  $compose | Set-Content -Encoding ASCII -Path docker-compose.yml

  # schema.sql
@'
CREATE TABLE IF NOT EXISTS pbx_instances (
  id              BIGSERIAL PRIMARY KEY,
  name            TEXT NOT NULL,
  tenant          TEXT NOT NULL,
  base_url        TEXT NOT NULL UNIQUE,
  auth_mode       TEXT NOT NULL,
  xapi_client_id  TEXT,
  xapi_secret_ref TEXT,
  enabled         BOOLEAN NOT NULL DEFAULT TRUE,
  created_at      TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS metric_points (
  pbx_id      BIGINT NOT NULL REFERENCES pbx_instances(id) ON DELETE CASCADE,
  metric_name TEXT   NOT NULL,
  ts          TIMESTAMPTZ NOT NULL,
  value_num   NUMERIC,
  value_json  JSONB,
  PRIMARY KEY (pbx_id, metric_name, ts)
);

CREATE INDEX IF NOT EXISTS idx_metric_points_ts ON metric_points (ts);
'@ | Set-Content -Encoding ASCII -Path schema.sql

  # requirements.txt
@'
httpx==0.27.0
asyncpg==0.29.0
pyyaml==6.0.2
'@ | Set-Content -Encoding ASCII -Path requirements.txt

  # config.yaml
  $pbxYaml = ($PBXs | ForEach-Object {
@"
  - name: "$($_.name)"
    tenant: "$($_.tenant)"
    base_url: "$($_.base_url)"
    verify_tls: $([string]$_.verify_tls).ToLower()
    alert_email: "$($_.alert_email)"
    auth:
      mode: "webclient_login"
      username_env: "$($_.username_env)"
      password_env: "$($_.password_env)"
      mfa_env: "$($_.mfa_env)"
"@
  }) -join "`n"

  $smtpHost = [Environment]::GetEnvironmentVariable("SMTP_HOST","Process")
  $smtpPort = [Environment]::GetEnvironmentVariable("SMTP_PORT","Process")
  $smtpFrom = [Environment]::GetEnvironmentVariable("SMTP_FROM","Process")
  $smtpTls  = [Environment]::GetEnvironmentVariable("SMTP_STARTTLS","Process")

@"
database:
  dsn: "postgresql://3cxmon:3cxmon@localhost:5432/3cxmon"

polling:
  active_calls_seconds: $ActiveCallsSeconds
  license_check_hours: $LicenseCheckHours

alerts:
  enabled: true
  near_capacity_delta: 1
  cooldown_minutes: 30
  smtp_host: "$smtpHost"
  smtp_port: $smtpPort
  from_address: "$smtpFrom"
  use_starttls: $smtpTls
  smtp_username_env: "SMTP_USERNAME"
  smtp_password_env: "SMTP_PASSWORD"

retention:
  enabled: true
  run_every_hours: $RetentionRunEveryHours

pbxes:
$pbxYaml
"@ | Set-Content -Encoding ASCII -Path config.yaml

  # Grafana provisioning (BOM-free + default DB fix) 
  if (-not $NoGrafana) {
    $ds = @'
apiVersion: 1

datasources:
  - name: Postgres-3cxmon
    type: postgres
    uid: ds_3cxmon
    access: proxy
    url: postgres:5432
    database: 3cxmon
    user: 3cxmon
    secureJsonData:
      password: 3cxmon
    jsonData:
      database: 3cxmon
      sslmode: disable
      postgresVersion: 1600
      timescaledb: false
    isDefault: true
'@
    Set-ContentUtf8NoBom -Path "$Dir\grafana\provisioning\datasources\postgres.yml" -Value $ds

    $prov = @'
apiVersion: 1

providers:
  - name: "3cx-dashboards"
    folder: "3CX Monitoring"
    type: file
    options:
      path: /var/lib/grafana/dashboards
      foldersFromFilesStructure: false
'@
    Set-ContentUtf8NoBom -Path "$Dir\grafana\provisioning\dashboards\dashboards.yml" -Value $prov

    $dash = @'
{
  "id": null,
  "uid": "msp-3cx-multi",
  "title": "3CX Multi-PBX Monitoring (Calls + License)",
  "tags": ["3CX","MSP","Monitoring"],
  "timezone": "browser",
  "schemaVersion": 39,
  "version": 1,
  "refresh": "5s",
  "time": { "from": "now-6h", "to": "now" },
  "templating": {
    "list": [
      {
        "name": "pbx",
        "label": "PBX",
        "type": "query",
        "datasource": { "type": "postgres", "uid": "ds_3cxmon" },
        "refresh": 1,
        "multi": true,
        "includeAll": true,
        "query": "SELECT id AS __value, name AS __text FROM pbx_instances ORDER BY name;",
        "current": { "selected": true, "text": "All", "value": "$__all" }
      }
    ]
  },
  "panels": [
    {
      "type": "timeseries",
      "title": "Active Calls (15s samples)",
      "datasource": { "type": "postgres", "uid": "ds_3cxmon" },
      "gridPos": { "x": 0, "y": 0, "w": 16, "h": 10 },
      "targets": [
        {
          "refId": "A",
          "format": "time_series",
          "rawSql": "SELECT mp.ts AS \"time\", mp.value_num AS value, pi.name AS metric FROM metric_points mp JOIN pbx_instances pi ON pi.id = mp.pbx_id WHERE mp.metric_name='active_calls' AND pi.id IN (${pbx:csv}) AND $__timeFilter(mp.ts) ORDER BY mp.ts;"
        }
      ]
    },
    {
      "type": "timeseries",
      "title": "Licensed Calls (Max Concurrent Calls)",
      "datasource": { "type": "postgres", "uid": "ds_3cxmon" },
      "gridPos": { "x": 16, "y": 0, "w": 8, "h": 10 },
      "targets": [
        {
          "refId": "A",
          "format": "time_series",
          "rawSql": "SELECT mp.ts AS \"time\", mp.value_num AS value, pi.name AS metric FROM metric_points mp JOIN pbx_instances pi ON pi.id = mp.pbx_id WHERE mp.metric_name='licensed_calls' AND pi.id IN (${pbx:csv}) AND $__timeFilter(mp.ts) ORDER BY mp.ts;"
        }
      ]
    }
  ]
}
'@
    Set-ContentUtf8NoBom -Path "$Dir\grafana\dashboards\3cx-multi-pbx.json" -Value $dash
  }

  # Write collector.py (single-quoted here-string, complete)
  Set-Content -Encoding ASCII -Path (Join-Path $Dir "collector.py") -Value $collectorPy

  Info "Files written (overwritten)."
}

# -------------------- COLLECTOR PY TEMPLATE --------------------
# Stored as a PowerShell string; written to collector.py in Write-Files (no interpolation).
$collectorPy = @'
<PASTE COLLECTOR CONTENT HERE>
'@

# Replace placeholder with the full collector from the previous message in this conversation.
# To keep this message from doubling in size, we embed it programmatically below:
$collectorPy = $collectorPy -replace '<PASTE COLLECTOR CONTENT HERE>', @'
import asyncio
import os
import time
import json
import smtplib
from email.message import EmailMessage
from dataclasses import dataclass
from datetime import datetime, timezone

import yaml
import httpx
import asyncpg

POOL_MAX_SIZE = int(os.getenv("POOL_MAX_SIZE", "10"))

def utcnow():
    return datetime.now(timezone.utc)

@dataclass
class PBXAuth:
    mode: str
    username_env: str
    password_env: str
    mfa_env: str | None = None

@dataclass
class PBX:
    name: str
    tenant: str
    base_url: str
    verify_tls: bool
    alert_email: str
    auth: PBXAuth

class TokenCache:
    def __init__(self):
        self.token = {}
        self.expiry = {}
        self._locks = {}

    def _lock_for(self, key: str):
        if key not in self._locks:
            self._locks[key] = asyncio.Lock()
        return self._locks[key]

    async def get(self, pbx: PBX, client: httpx.AsyncClient) -> str:
        now = time.time()
        if pbx.base_url in self.token and now < self.expiry[pbx.base_url] - 10:
            return self.token[pbx.base_url]

        async with self._lock_for(pbx.base_url):
            now = time.time()
            if pbx.base_url in self.token and now < self.expiry[pbx.base_url] - 10:
                return self.token[pbx.base_url]

            if pbx.auth.mode != "webclient_login":
                raise RuntimeError(f"{pbx.name}: Unsupported auth mode '{pbx.auth.mode}'")

            username = os.getenv(pbx.auth.username_env, "")
            password = os.getenv(pbx.auth.password_env, "")
            mfa = os.getenv(pbx.auth.mfa_env, "") if pbx.auth.mfa_env else ""

            if not username or not password:
                raise RuntimeError(f"{pbx.name}: Missing env vars {pbx.auth.username_env}/{pbx.auth.password_env}")

            body = {"SecurityCode": mfa or "", "Username": username, "Password": password}

            r = await client.post(f"{pbx.base_url}/webclient/api/Login/GetAccessToken", json=body, timeout=30.0)
            r.raise_for_status()
            resp = r.json()

            token_container = resp.get("Token") or resp.get("token") or resp.get("access_token") or resp
            access_token = None
            expires_in = 55 * 60

            if isinstance(token_container, str):
                access_token = token_container
            elif isinstance(token_container, dict):
                access_token = token_container.get("access_token") or token_container.get("AccessToken")
                if token_container.get("expires_in"):
                    try:
                        expires_in = int(token_container["expires_in"])
                    except Exception:
                        pass

            if not access_token:
                raise RuntimeError(f"{pbx.name}: Login succeeded but no access_token. keys={list(resp.keys())}")

            self.token[pbx.base_url] = access_token
            self.expiry[pbx.base_url] = time.time() + int(expires_in)
            return access_token

def _to_json_text(value):
    if value is None:
        return None
    if isinstance(value, (str, bytes)):
        return value
    return json.dumps(value, ensure_ascii=False)

async def ensure_pbx_row(pool: asyncpg.Pool, pbx: PBX) -> int:
    async with pool.acquire() as conn:
        row = await conn.fetchrow(
            """
            INSERT INTO pbx_instances (name, tenant, base_url, auth_mode, xapi_client_id, xapi_secret_ref, enabled)
            VALUES ($1,$2,$3,'webclient_login',NULL,NULL,true)
            ON CONFLICT (base_url) DO UPDATE
              SET name=EXCLUDED.name, tenant=EXCLUDED.tenant, enabled=true
            RETURNING id
            """,
            pbx.name, pbx.tenant, pbx.base_url
        )
        return int(row["id"])

async def upsert_metric(pool: asyncpg.Pool, pbx_id: int, metric: str, ts: datetime, value_num=None, value_json=None):
    value_json_text = _to_json_text(value_json)
    async with pool.acquire() as conn:
        await conn.execute(
            """
            INSERT INTO metric_points (pbx_id, metric_name, ts, value_num, value_json)
            VALUES ($1,$2,$3,$4, $5::jsonb)
            ON CONFLICT (pbx_id, metric_name, ts) DO UPDATE
              SET value_num=EXCLUDED.value_num, value_json=EXCLUDED.value_json
            """,
            pbx_id, metric, ts, value_num, value_json_text
        )

async def get_latest_metric(pool: asyncpg.Pool, pbx_id: int, metric: str):
    async with pool.acquire() as conn:
        return await conn.fetchrow(
            """
            SELECT value_num, ts
            FROM metric_points
            WHERE pbx_id=$1 AND metric_name=$2
            ORDER BY ts DESC
            LIMIT 1
            """,
            pbx_id, metric
        )

def extract_licensed_calls(payload: dict):
    for k in ("ConcurrentCalls","concurrentCalls","MaxConcurrentCalls","MaxSimCalls","LicensedCalls","licensedCalls"):
        if k in payload:
            try:
                return int(payload[k])
            except Exception:
                pass
    return None

def smtp_send(to_addr: str, subject: str, body: str, alert_cfg: dict):
    host = alert_cfg.get("smtp_host") or os.getenv("SMTP_HOST","")
    port = int(alert_cfg.get("smtp_port") or os.getenv("SMTP_PORT","0") or 0)
    from_addr = alert_cfg.get("from_address") or os.getenv("SMTP_FROM","")
    use_starttls = str(alert_cfg.get("use_starttls") or os.getenv("SMTP_STARTTLS","true")).lower() == "true"

    user_env = alert_cfg.get("smtp_username_env") or "SMTP_USERNAME"
    pass_env = alert_cfg.get("smtp_password_env") or "SMTP_PASSWORD"
    username = os.getenv(user_env, "")
    password = os.getenv(pass_env, "")

    if not host or not port or not from_addr or not to_addr:
        return False, "SMTP not configured (host/port/from/to missing)"

    msg = EmailMessage()
    msg["From"] = from_addr
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.set_content(body)

    with smtplib.SMTP(host, port, timeout=30) as s:
        if use_starttls:
            s.starttls()
        if username and password:
            s.login(username, password)
        s.send_message(msg)

    return True, "sent"

async def poll_active_calls_once(pbx: PBX, pbx_id: int, tokens: TokenCache, pool: asyncpg.Pool,
                                alert_cfg: dict, alert_state: dict):
    async with httpx.AsyncClient(verify=pbx.verify_tls, timeout=10.0) as client:
        ts = utcnow()
        try:
            token = await tokens.get(pbx, client)
            headers = {"Authorization": f"Bearer {token}"}
            r = await client.get(f"{pbx.base_url}/xapi/v1/ActiveCalls", headers=headers)
            r.raise_for_status()
            j = r.json()
            count = j.get("@odata.count")
            if count is None:
                count = len(j.get("value", []) or [])

            active_calls = int(count)
            await upsert_metric(pool, pbx_id, "active_calls", ts, value_num=active_calls, value_json=j)
            print(f"[{ts.isoformat()}] {pbx.name} active_calls={active_calls}")

            if alert_cfg.get("enabled", False) and pbx.alert_email:
                delta = int(alert_cfg.get("near_capacity_delta", 1))
                cooldown = int(alert_cfg.get("cooldown_minutes", 30))
                lic = await get_latest_metric(pool, pbx_id, "licensed_calls")
                if lic and lic["value_num"] is not None:
                    licensed_calls = int(lic["value_num"])
                    threshold = max(0, licensed_calls - delta)

                    key = f"{pbx_id}:nearcap"
                    last = alert_state.get(key, {"state": "ok", "last_sent": None})

                    now = utcnow()
                    near = (active_calls >= threshold and licensed_calls > 0)

                    if near and last["state"] != "near":
                        if last["last_sent"] is None or (now - last["last_sent"]).total_seconds() >= cooldown * 60:
                            subject = f"ALERT: {pbx.name} near capacity"
                            body = (
                                f"PBX: {pbx.name}\n"
                                f"URL: {pbx.base_url}\n"
                                f"Alarm time (UTC): {now.isoformat()}\n\n"
                                f"Active calls: {active_calls}\n"
                                f"Licensed calls: {licensed_calls}\n"
                                f"Trigger: active_calls >= licensed_calls - {delta}\n"
                            )
                            ok, msg = smtp_send(pbx.alert_email, subject, body, alert_cfg)
                            print(f"[{now.isoformat()}] {pbx.name} ALERT -> {pbx.alert_email}: {msg}")
                            alert_state[key] = {"state": "near", "last_sent": now}
                        else:
                            alert_state[key] = {"state": "near", "last_sent": last["last_sent"]}
                    elif not near:
                        alert_state[key] = {"state": "ok", "last_sent": last["last_sent"]}

        except Exception as e:
            await upsert_metric(pool, pbx_id, "active_calls", ts, value_json={"error": str(e)})
            print(f"[{ts.isoformat()}] {pbx.name} active_calls ERROR: {e}")

async def poll_license_once(pbx: PBX, pbx_id: int, tokens: TokenCache, pool: asyncpg.Pool):
    async with httpx.AsyncClient(verify=pbx.verify_tls, timeout=20.0) as client:
        ts = utcnow()
        try:
            token = await tokens.get(pbx, client)
            headers = {"Authorization": f"Bearer {token}"}
            licensed = None
            details = {}
            r = await client.get(f"{pbx.base_url}/xapi/v1/SystemStatus", headers=headers)
            if r.status_code == 200:
                sysj = r.json()
                details["systemstatus"] = sysj
                licensed = extract_licensed_calls(sysj)
            if licensed is None:
                await upsert_metric(pool, pbx_id, "licensed_calls", ts, value_json={"error": "Could not determine licensed calls", "details": details})
                print(f"[{ts.isoformat()}] {pbx.name} licensed_calls ERROR: could not determine")
                return
            await upsert_metric(pool, pbx_id, "licensed_calls", ts, value_num=int(licensed), value_json=details)
            print(f"[{ts.isoformat()}] {pbx.name} licensed_calls={int(licensed)}")
        except Exception as e:
            await upsert_metric(pool, pbx_id, "licensed_calls", ts, value_json={"error": str(e)})
            print(f"[{ts.isoformat()}] {pbx.name} licensed_calls ERROR: {e}")

async def rollup_active_calls(pool: asyncpg.Pool):
    async with pool.acquire() as conn:
        await conn.execute("BEGIN;")
        try:
            # 31-60 days: hourly max
            await conn.execute("""
                WITH src AS (
                  SELECT pbx_id,
                         date_trunc('hour', ts) AS bucket_ts,
                         max(value_num) AS max_val
                  FROM metric_points
                  WHERE metric_name='active_calls'
                    AND ts < (now() - interval '30 days')
                    AND ts >= (now() - interval '60 days')
                  GROUP BY pbx_id, date_trunc('hour', ts)
                )
                INSERT INTO metric_points (pbx_id, metric_name, ts, value_num, value_json)
                SELECT pbx_id, 'active_calls', bucket_ts, max_val, NULL::jsonb
                FROM src
                ON CONFLICT (pbx_id, metric_name, ts) DO UPDATE
                  SET value_num=EXCLUDED.value_num, value_json=EXCLUDED.value_json;
            """)
            await conn.execute("""
                DELETE FROM metric_points
                WHERE metric_name='active_calls'
                  AND ts < (now() - interval '30 days')
                  AND ts >= (now() - interval '60 days')
                  AND ts <> date_trunc('hour', ts);
            """)

            # 61-120 days: daily max
            await conn.execute("""
                WITH src AS (
                  SELECT pbx_id,
                         date_trunc('day', ts) AS bucket_ts,
                         max(value_num) AS max_val
                  FROM metric_points
                  WHERE metric_name='active_calls'
                    AND ts < (now() - interval '60 days')
                    AND ts >= (now() - interval '120 days')
                  GROUP BY pbx_id, date_trunc('day', ts)
                )
                INSERT INTO metric_points (pbx_id, metric_name, ts, value_num, value_json)
                SELECT pbx_id, 'active_calls', bucket_ts, max_val, NULL::jsonb
                FROM src
                ON CONFLICT (pbx_id, metric_name, ts) DO UPDATE
                  SET value_num=EXCLUDED.value_num, value_json=EXCLUDED.value_json;
            """)
            await conn.execute("""
                DELETE FROM metric_points
                WHERE metric_name='active_calls'
                  AND ts < (now() - interval '60 days')
                  AND ts >= (now() - interval '120 days')
                  AND ts <> date_trunc('day', ts);
            """)

            # 121+ days: weekly max (Sunday start)
            await conn.execute("""
                WITH src AS (
                  SELECT pbx_id,
                         (date_trunc('day', ts) - (extract(dow from ts)::int * interval '1 day')) AS bucket_ts,
                         max(value_num) AS max_val
                  FROM metric_points
                  WHERE metric_name='active_calls'
                    AND ts < (now() - interval '120 days')
                  GROUP BY pbx_id, (date_trunc('day', ts) - (extract(dow from ts)::int * interval '1 day'))
                )
                INSERT INTO metric_points (pbx_id, metric_name, ts, value_num, value_json)
                SELECT pbx_id, 'active_calls', bucket_ts, max_val, NULL::jsonb
                FROM src
                ON CONFLICT (pbx_id, metric_name, ts) DO UPDATE
                  SET value_num=EXCLUDED.value_num, value_json=EXCLUDED.value_json;
            """)
            await conn.execute("""
                DELETE FROM metric_points
                WHERE metric_name='active_calls'
                  AND ts < (now() - interval '120 days')
                  AND ts <> (date_trunc('day', ts) - (extract(dow from ts)::int * interval '1 day'));
            """)

            await conn.execute("COMMIT;")
        except Exception:
            await conn.execute("ROLLBACK;")
            raise

async def retention_loop(pool: asyncpg.Pool, retention_cfg: dict):
    if not retention_cfg.get("enabled", True):
        return
    interval = max(1, int(retention_cfg.get("run_every_hours", 24))) * 3600
    await asyncio.sleep(10)
    while True:
        try:
            await rollup_active_calls(pool)
            print(f"[{utcnow().isoformat()}] retention: rollup complete")
        except Exception as e:
            print(f"[{utcnow().isoformat()}] retention ERROR: {e}")
        await asyncio.sleep(interval)

async def active_loop(pbxes, pbx_ids, tokens, pool, interval, alert_cfg):
    alert_state = {}
    while True:
        await asyncio.gather(*[
            poll_active_calls_once(pbx, pbx_ids[pbx.base_url], tokens, pool, alert_cfg, alert_state)
            for pbx in pbxes
        ])
        await asyncio.sleep(interval)

async def license_loop(pbxes, pbx_ids, tokens, pool, hours):
    interval = max(1, hours) * 3600
    while True:
        await asyncio.gather(*[
            poll_license_once(pbx, pbx_ids[pbx.base_url], tokens, pool)
            for pbx in pbxes
        ])
        await asyncio.sleep(interval)

async def main():
    cfg = yaml.safe_load(open("config.yaml", "r", encoding="utf-8"))
    dsn = cfg["database"]["dsn"]
    poll_seconds = int(cfg["polling"].get("active_calls_seconds", 15))
    license_hours = int(cfg["polling"].get("license_check_hours", 24))
    alert_cfg = cfg.get("alerts", {}) or {}
    retention_cfg = cfg.get("retention", {}) or {}

    pbxes = []
    for p in cfg.get("pbxes", []):
        auth_cfg = p.get("auth", {}) or {}
        pbxes.append(PBX(
            name=p["name"],
            tenant=p.get("tenant",""),
            base_url=p["base_url"].rstrip("/"),
            verify_tls=bool(p.get("verify_tls", True)),
            alert_email=p.get("alert_email","") or "",
            auth=PBXAuth(
                mode=auth_cfg.get("mode", "webclient_login"),
                username_env=auth_cfg["username_env"],
                password_env=auth_cfg["password_env"],
                mfa_env=auth_cfg.get("mfa_env")
            )
        ))

    if not pbxes:
        raise RuntimeError("No PBXs defined in config.yaml under 'pbxes:'")

    pool = await asyncpg.create_pool(dsn=dsn, min_size=1, max_size=POOL_MAX_SIZE)

    pbx_ids = {}
    for pbx in pbxes:
        pbx_ids[pbx.base_url] = await ensure_pbx_row(pool, pbx)

    tokens = TokenCache()

    await asyncio.gather(
        active_loop(pbxes, pbx_ids, tokens, pool, poll_seconds, alert_cfg),
        license_loop(pbxes, pbx_ids, tokens, pool, license_hours),
        retention_loop(pool, retention_cfg),
    )

if __name__ == "__main__":
    asyncio.run(main())
'@
# -------------------- COLLECTOR PY TEMPLATE END --------------------

function Start-Collector {
  param([hashtable]$Py)

  $env:POOL_MAX_SIZE = "$PoolMaxSize"

  Info "Creating Python venv + installing dependencies..."
  if (Test-Path ".venv") { Remove-Item -Recurse -Force ".venv" | Out-Null }

  if ($Py.kind -eq "python") { & $Py.cmd[0] -m venv .venv }
  else { & $Py.cmd[0] $Py.cmd[1] -m venv .venv }

  $pip = ".\.venv\Scripts\pip.exe"
  $python = ".\.venv\Scripts\python.exe"

  & $python -m pip install --upgrade pip *> $null
  & $pip install -r requirements.txt *> $null

  $failures = 0
  $logDir = Join-Path (Get-Location) "logs"
  $stamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
  $logFile = Join-Path $logDir "collector-$stamp.log"

  Info "Collector log: $logFile"
  Info "Starting collector (Ctrl+C to stop)..."

  while ($true) {
    Add-Content -Path $logFile -Value ("`r`n===== COLLECTOR START {0} (failure {1}/{2}) =====`r`n" -f (Get-Date), $failures, $MaxCollectorFailures)

    $prev = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    & $python .\collector.py 2>&1 | Tee-Object -FilePath $logFile -Append
    $code = $LASTEXITCODE
    $ErrorActionPreference = $prev

    if ($code -eq 0) { break }

    $failures++
    Err "Collector crashed (exit code $code). Failure $failures of $MaxCollectorFailures."
    if ($failures -ge $MaxCollectorFailures) {
      throw "Collector failed $failures times. Not restarting again. See: $logFile"
    }
    Warn "Restarting collector in 5 seconds..."
    Start-Sleep -Seconds 5
  }
}

# -------------------- MAIN --------------------
Ensure-DockerOnline
$py = Ensure-Python
$composeCmd = "docker compose"

Ensure-Dirs $ProjectDir
Set-Location $ProjectDir

$pbxSecrets = Load-Secrets $PbxSecretsPath

if ($ListPbxs) { Show-Pbxs $pbxSecrets; return }

if ($RemovePbx -gt 0) {
  $pbxSecrets = Remove-PbxByIndex $pbxSecrets $RemovePbx
  if ($CompactIndices) {
    $pbxSecrets = Compact-Indices $pbxSecrets
    Info "CompactIndices: renumbered to 1..$($pbxSecrets.Count)"
  }
  Save-Secrets $PbxSecretsPath $pbxSecrets
}

$smtp = Prompt-SmtpConfig
Apply-SmtpToEnv $smtp

$pbxSecrets = Prompt-AddPbxLoop $pbxSecrets
if ($NonInteractive -and $pbxSecrets.Count -eq 0) { throw "NonInteractive requires stored PBXs." }
if ($pbxSecrets.Count -eq 0) { throw "No PBXs configured." }

if ($CompactIndices) {
  $pbxSecrets = Compact-Indices $pbxSecrets
  Save-Secrets $PbxSecretsPath $pbxSecrets
  Info "CompactIndices applied: indices 1..$($pbxSecrets.Count)"
}

Apply-PbxSecretsToEnv $pbxSecrets

$pbxList = Build-PbxListFromSecrets $pbxSecrets
Write-Files -Dir $ProjectDir -PBXs $pbxList

# Compose project name
$projName = (Split-Path -Leaf $ProjectDir) -replace '[^a-zA-Z0-9_-]', ''
if (IsNullOrWhiteSpace $projName) { $projName = 'msp-3cx-monitor-demo' }
$env:COMPOSE_PROJECT_NAME = $projName

if ($ResetDb) {
  Warn "ResetDb: stopping containers and removing pgdata..."
  Invoke-CmdCapture "$composeCmd down --remove-orphans" | Out-Null
  if (Test-Path "$ProjectDir\pgdata") { Remove-Item -Recurse -Force "$ProjectDir\pgdata" }
}

Info "Starting Docker services..."
$r = Invoke-CmdCapture "$composeCmd up -d --remove-orphans"
if ($r.ExitCode -ne 0) { throw "Compose up failed:`n$($r.Output)" }

# Wait for Postgres
$pgId = (Invoke-CmdCapture "$composeCmd ps -q postgres").Output.Trim()
if (IsNullOrWhiteSpace $pgId) { throw "Could not resolve postgres container ID." }

Info "Waiting for Postgres to be ready and stable..."
$sw = [Diagnostics.Stopwatch]::StartNew()
$stable = 0
while ($sw.Elapsed.TotalSeconds -lt $PostgresWaitSeconds) {
  $q = Invoke-CmdCapture "docker exec $pgId pg_isready -U 3cxmon -d postgres"
  if ($q.ExitCode -eq 0 -and ($q.Output -match "accepting connections")) { $stable++; if ($stable -ge 3) { break } }
  else { $stable = 0 }
  Start-Sleep -Seconds 2
}
if ($stable -lt 3) { throw "Postgres did not become stable in time." }
Info "Postgres is stable."

# Ensure DB + schema
$check = Invoke-CmdCapture "docker exec $pgId psql -U 3cxmon -d postgres -tAc ""SELECT 1 FROM pg_database WHERE datname='3cxmon';"""
if ($check.ExitCode -ne 0) { throw "DB check failed:`n$($check.Output)" }
if ($check.Output.Trim() -ne "1") {
  $create = Invoke-CmdCapture "docker exec $pgId psql -U 3cxmon -d postgres -v ON_ERROR_STOP=1 -c ""CREATE DATABASE 3cxmon;"""
  if ($create.ExitCode -ne 0) { throw "DB create failed:`n$($create.Output)" }
}
$schemaPath = Join-Path (Get-Location) "schema.sql"
$cp = Invoke-CmdCapture "docker cp ""$schemaPath"" $pgId`:/tmp/schema.sql"
if ($cp.ExitCode -ne 0) { throw "docker cp failed:`n$($cp.Output)" }
$apply = Invoke-CmdCapture "docker exec $pgId psql -U 3cxmon -d 3cxmon -v ON_ERROR_STOP=1 -f /tmp/schema.sql"
if ($apply.ExitCode -ne 0) { throw "Schema apply failed:`n$($apply.Output)" }
Info "Schema applied."

if (-not $NoGrafana) {
  Info "Grafana: http://localhost:3000 (admin/admin)"
  Info "Dashboard: Dashboards -> 3CX Monitoring -> 3CX Multi-PBX Monitoring (Calls + License)"
}

Start-Collector -Py $py