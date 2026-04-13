[CmdletBinding()]
param(
  [string]$ProjectDir = "",
  [switch]$SkipInstalls,
  [switch]$NoGrafana,
  [switch]$ResetDb
)

$ErrorActionPreference = 'Stop'

function Info([string]$m){ Write-Host "[INFO] $m" -ForegroundColor Cyan }
function Warn([string]$m){ Write-Host "[WARN] $m" -ForegroundColor Yellow }
function Err ([string]$m){ Write-Host "[ERR ] $m" -ForegroundColor Red }

function HasCmd([string]$n){
  return (Get-Command $n -ErrorAction SilentlyContinue) -ne $null
}

function Test-PythonExec {
  param([string[]]$Cmd)
  try {
    $exe = $Cmd[0]
    $args = @()
    if ($Cmd.Count -gt 1) { $args = $Cmd[1..($Cmd.Count-1)] }
    $out = & $exe @args --version 2>&1
    return ($LASTEXITCODE -eq 0 -and $out -match '^Python\s+\d+\.\d+')
  } catch { return $false }
}

function Get-PythonLauncher {
  # Prefer py -3 (avoids Windows Store alias issues)
  if (HasCmd 'py' -and (Test-PythonExec -Cmd @('py','-3'))) { return @{ Kind='py'; Cmd=@('py','-3') } }
  if (HasCmd 'python' -and (Test-PythonExec -Cmd @('python'))) { return @{ Kind='python'; Cmd=@('python') } }
  return $null
}

function Get-ComposeCmd {
  try { docker compose version *> $null; return "docker compose" } catch {}
  if (HasCmd 'docker-compose') { return "docker-compose" }
  throw "Docker Compose not found. Ensure Docker Desktop is installed and running."
}

function Ensure-Prereqs {
  param([switch]$SkipInstalls)

  $missingWingetIds = @()
  if (-not (HasCmd 'docker')) { $missingWingetIds += 'Docker.DockerDesktop' }
  if (-not (HasCmd 'git'))    { $missingWingetIds += 'Git.Git' }
  if (-not (Get-PythonLauncher)) { $missingWingetIds += 'Python.Python.3.11' }

  if ($missingWingetIds.Count -eq 0) {
    Info "Prereqs look OK (docker, git, python)."
    return
  }

  Warn ("Missing prerequisites: " + ($missingWingetIds -join ', '))
  if ($SkipInstalls) { throw "Missing prerequisites and -SkipInstalls specified." }

  if (-not (HasCmd 'winget')) {
    throw "winget not found. Install prerequisites manually: Docker Desktop, Git, Python 3.11+"
  }

  foreach ($id in $missingWingetIds) {
    Info "Attempting install via winget: $id"
    winget install --id $id --silent --accept-package-agreements --accept-source-agreements | Out-Null
  }

  Warn "Install attempted. If commands still not found, CLOSE and reopen PowerShell (PATH refresh)."
}

function Test-DockerReady {
  try { docker info *> $null; return ($LASTEXITCODE -eq 0) } catch { return $false }
}

function Ensure-WSL2Running {
  try {
    wsl --status *> $null
    wsl -e sh -lc "echo WSL_OK" *> $null
  } catch {
    Warn "WSL2 check failed. If Docker Desktop can't start, verify WSL2 is installed and working."
  }
}

function Start-DockerDesktop {
  $possible = @(
    "$env:ProgramFiles\Docker\Docker\Docker Desktop.exe",
    "$env:LocalAppData\Programs\Docker\Docker\Docker Desktop.exe"
  ) | Where-Object { Test-Path $_ }

  if ($possible.Count -gt 0) {
    Info "Launching Docker Desktop..."
    Start-Process -FilePath $possible[0] | Out-Null
    return
  }

  throw "Docker Desktop executable not found. Start Docker Desktop manually."
}

function Ensure-DockerOnline {
  param([int]$TimeoutSeconds = 240)

  if (Test-DockerReady) {
    Info "Docker engine is already running."
    return
  }

  Warn "Docker engine is not running. Attempting to start Docker Desktop (WSL2)..."
  Ensure-WSL2Running
  Start-DockerDesktop

  $sw = [Diagnostics.Stopwatch]::StartNew()
  while ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
    Start-Sleep -Seconds 3
    if (Test-DockerReady) { Info "Docker engine is online."; return }
    Info ("Waiting for Docker engine... ({0}s / {1}s)" -f [int]$sw.Elapsed.TotalSeconds, $TimeoutSeconds)
  }

  throw "Timed out waiting for Docker. Open Docker Desktop and ensure it shows 'Engine running'."
}

function Resolve-ProjectDir {
  if (-not [string]::IsNullOrWhiteSpace($ProjectDir)) {
    return (Resolve-Path $ProjectDir).Path
  }
  if (Test-Path (Join-Path $PWD "docker-compose.yml")) {
    return (Resolve-Path $PWD).Path
  }
  return (Join-Path (Resolve-Path $PWD).Path "msp-3cx-monitor-demo")
}

function Prompt-Config {
  $base = Read-Host "Enter PBX base URL (e.g. https://pbx.example.com:5001)"
  $cid  = Read-Host "Enter XAPI Client ID (3CX Admin -> Integrations -> API)"
  $sec  = Read-Host "Enter XAPI Client Secret (hidden)" -AsSecureString
  $tls  = Read-Host "Verify TLS certificate? (y/n) [y]"

  if ([string]::IsNullOrWhiteSpace($tls)) { $tls = 'y' }
  $verifyTls = ($tls.ToLower() -eq 'y')

  $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
  $plainSecret = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
  [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)

  # Store secret in current process env only (do not write to disk)
  $env:PBX_DEV_XAPI_SECRET = $plainSecret

  return @{
    BaseUrl   = $base     # DO NOT MODIFY (no :5001 fix)
    ClientId  = $cid
    VerifyTls = $verifyTls
  }
}

function Invoke-CmdCapture {
  param([string]$CommandLine)
  $out = cmd.exe /c "$CommandLine 2>&1"
  return @{
    ExitCode = $LASTEXITCODE
    Output   = ($out | Out-String)
  }
}

function Write-ProjectFiles {
  param([string]$Dir, [hashtable]$Cfg)

  New-Item -ItemType Directory -Force -Path $Dir | Out-Null
  Set-Location $Dir

  Info "Project directory: $Dir"

  $compose = if ($NoGrafana) {
@'
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
'@
  } else {
@'
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
    ports:
      - "3000:3000"
    depends_on:
      - postgres
'@
  }

  # ALWAYS overwrite
  $compose | Set-Content -Encoding UTF8 -Path docker-compose.yml

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

CREATE INDEX IF NOT EXISTS idx_metric_points_ts
  ON metric_points (ts);
'@ | Set-Content -Encoding UTF8 -Path schema.sql

  @'
httpx==0.27.0
asyncpg==0.29.0
pyyaml==6.0.2
'@ | Set-Content -Encoding UTF8 -Path requirements.txt

  @"
database:
  dsn: "postgresql://3cxmon:3cxmon@localhost:5432/3cxmon"

polling:
  active_calls_seconds: 15
  license_check_hours: 24

pbxes:
  - name: "Dev-PBX"
    tenant: "Internal"
    base_url: "$($Cfg.BaseUrl)"
    xapi_client_id: "$($Cfg.ClientId)"
    xapi_client_secret_env: "PBX_DEV_XAPI_SECRET"
    verify_tls: $(if($Cfg.VerifyTls){'true'}else{'false'})
"@ | Set-Content -Encoding UTF8 -Path config.yaml

  # ALWAYS overwrite collector.py (asyncpg-version-safe JSONB)
  @'
import asyncio
import os
import time
import json
from dataclasses import dataclass
from datetime import datetime, timezone

import yaml
import httpx
import asyncpg

@dataclass
class PBX:
    name: str
    tenant: str
    base_url: str
    xapi_client_id: str
    xapi_client_secret_env: str
    verify_tls: bool = True

class TokenCache:
    def __init__(self):
        self.token = {}
        self.expiry = {}

    async def get(self, pbx: PBX, client: httpx.AsyncClient) -> str:
        now = time.time()
        if pbx.base_url in self.token and now < self.expiry[pbx.base_url] - 10:
            return self.token[pbx.base_url]

        secret = os.getenv(pbx.xapi_client_secret_env, "")
        if not secret:
            raise RuntimeError(f"Missing env var: {pbx.xapi_client_secret_env}")

        # Token endpoint: /connect/token
        r = await client.post(
            f"{pbx.base_url}/connect/token",
            data={
                "client_id": pbx.xapi_client_id,
                "client_secret": secret,
                "grant_type": "client_credentials",
            },
        )
        r.raise_for_status()
        j = r.json()
        self.token[pbx.base_url] = j["access_token"]
        self.expiry[pbx.base_url] = now + int(j.get("expires_in", 60))
        return self.token[pbx.base_url]

async def ensure_pbx_row(conn, pbx: PBX) -> int:
    row = await conn.fetchrow(
        """
        INSERT INTO pbx_instances (name, tenant, base_url, auth_mode, xapi_client_id, xapi_secret_ref, enabled)
        VALUES ($1,$2,$3,'xapi_client_credentials',$4,$5,true)
        ON CONFLICT (base_url) DO UPDATE
          SET name=EXCLUDED.name, tenant=EXCLUDED.tenant, xapi_client_id=EXCLUDED.xapi_client_id
        RETURNING id
        """,
        pbx.name, pbx.tenant, pbx.base_url, pbx.xapi_client_id, pbx.xapi_client_secret_env
    )
    return int(row["id"])

def _to_json_text(value):
    if value is None:
        return None
    if isinstance(value, (str, bytes)):
        return value
    return json.dumps(value, ensure_ascii=False)

async def upsert_metric(conn, pbx_id: int, metric: str, ts: datetime, value_num=None, value_json=None):
    value_json_text = _to_json_text(value_json)
    await conn.execute(
        """
        INSERT INTO metric_points (pbx_id, metric_name, ts, value_num, value_json)
        VALUES ($1,$2,$3,$4, $5::jsonb)
        ON CONFLICT (pbx_id, metric_name, ts) DO UPDATE
          SET value_num=EXCLUDED.value_num, value_json=EXCLUDED.value_json
        """,
        pbx_id, metric, ts, value_num, value_json_text
    )

async def fetch_active_calls(pbx: PBX, tokens: TokenCache, client: httpx.AsyncClient):
    token = await tokens.get(pbx, client)
    headers = {"Authorization": f"Bearer {token}"}
    r = await client.get(f"{pbx.base_url}/xapi/v1/ActiveCalls", headers=headers)
    r.raise_for_status()
    j = r.json()
    count = j.get("@odata.count")
    if count is None:
        count = len(j.get("value", []) or [])
    return int(count), j

async def main():
    cfg = yaml.safe_load(open("config.yaml", "r", encoding="utf-8"))
    dsn = cfg["database"]["dsn"]
    poll_15s = int(cfg["polling"]["active_calls_seconds"])
    p = cfg["pbxes"][0]

    pbx = PBX(
        name=p["name"],
        tenant=p["tenant"],
        base_url=p["base_url"].rstrip("/"),
        xapi_client_id=p["xapi_client_id"],
        xapi_client_secret_env=p["xapi_client_secret_env"],
        verify_tls=bool(p.get("verify_tls", True)),
    )

    conn = await asyncpg.connect(dsn=dsn)
    pbx_id = await ensure_pbx_row(conn, pbx)
    tokens = TokenCache()

    async with httpx.AsyncClient(verify=pbx.verify_tls, timeout=10.0) as client:
        while True:
            ts = datetime.now(timezone.utc)
            try:
                count, payload = await fetch_active_calls(pbx, tokens, client)
                await upsert_metric(conn, pbx_id, "active_calls", ts, value_num=count, value_json=payload)
                print(f"[{ts.isoformat()}] active_calls={count}")
            except Exception as e:
                await upsert_metric(conn, pbx_id, "active_calls", ts, value_json={"error": str(e)})
                print(f"[{ts.isoformat()}] ERROR: {e}")
            await asyncio.sleep(poll_15s)

if __name__ == "__main__":
    asyncio.run(main())
'@ | Set-Content -Encoding UTF8 -Path collector.py

  Info "Files written."
}

function Compose-DownAndReset {
  param([string]$ComposeCmd, [string]$Dir)

  if ($ResetDb) {
    Warn "ResetDb specified: stopping project containers and removing pgdata..."
    Invoke-CmdCapture "$ComposeCmd down --remove-orphans" | Out-Null

    $pgPath = Join-Path $Dir "pgdata"
    if (Test-Path $pgPath) {
      Remove-Item -Recurse -Force $pgPath
    }
  }
}

function Invoke-ComposeUp {
  param([string]$ComposeCmd)
  Info "Starting Docker services..."
  $r = Invoke-CmdCapture "$ComposeCmd up -d --remove-orphans"
  if ($r.ExitCode -ne 0) { throw "Compose up failed (exit $($r.ExitCode))`n$r" }
}

function Get-PostgresContainerId {
  param([string]$ComposeCmd)
  $r = Invoke-CmdCapture "$ComposeCmd ps -q postgres"
  $id = $r.Output.Trim()
  if ([string]::IsNullOrWhiteSpace($id)) { throw "Could not resolve Postgres container ID via compose." }
  return $id
}

function Wait-PostgresStable {
  param([string]$PgId, [string]$ComposeCmd, [int]$TimeoutSeconds = 240)

  Info "Waiting for Postgres to be ready and stable..."
  $sw = [Diagnostics.Stopwatch]::StartNew()
  $stable = 0

  while ($sw.Elapsed.TotalSeconds -lt $TimeoutSeconds) {
    $q = Invoke-CmdCapture "docker exec $PgId pg_isready -U 3cxmon -d postgres"
    if ($q.ExitCode -eq 0 -and ($q.Output -match "accepting connections")) {
      $stable++
      if ($stable -ge 3) {
        Info "Postgres is stable."
        return
      }
    } else {
      $stable = 0
    }
    Start-Sleep -Seconds 2
  }

  Warn "Postgres did not become stable. Dumping logs..."
  cmd.exe /c "$ComposeCmd logs --tail=200 postgres"
  throw "Postgres did not become stable within $TimeoutSeconds seconds."
}

function Ensure-DbAndApplySchema {
  param([string]$PgId, [string]$DbName = '3cxmon', [string]$User = '3cxmon')

  Info "Ensuring database '$DbName' exists..."

  $check = Invoke-CmdCapture "docker exec $PgId psql -U $User -d postgres -tAc ""SELECT 1 FROM pg_database WHERE datname='$DbName';"""
  if ($check.ExitCode -ne 0) { throw "DB existence check failed:`n$($check.Output)" }

  if ($check.Output.Trim() -ne "1") {
    Info "Database '$DbName' not found. Creating..."
    $create = Invoke-CmdCapture "docker exec $PgId psql -U $User -d postgres -v ON_ERROR_STOP=1 -c ""CREATE DATABASE $DbName;"""
    if ($create.ExitCode -ne 0) { throw "DB create failed:`n$($create.Output)" }
    Info "Database '$DbName' created."
  } else {
    Info "Database '$DbName' already exists."
  }

  $schemaPath = Join-Path (Get-Location) "schema.sql"
  if (-not (Test-Path $schemaPath)) { throw "schema.sql not found at $schemaPath" }

  Info "Copying schema.sql into container..."
  $cp = Invoke-CmdCapture "docker cp ""$schemaPath"" $PgId`:/tmp/schema.sql"
  if ($cp.ExitCode -ne 0) { throw "docker cp failed:`n$($cp.Output)" }

  Info "Applying schema.sql to '$DbName'..."
  $apply = Invoke-CmdCapture "docker exec $PgId psql -U $User -d $DbName -v ON_ERROR_STOP=1 -f /tmp/schema.sql"
  if ($apply.ExitCode -ne 0) { throw "Schema apply failed:`n$($apply.Output)" }

  Info "Schema applied."
}

function Start-Collector {
  $py = Get-PythonLauncher
  if (-not $py) { throw "Python not available. Ensure 'py -3 --version' or 'python --version' works, then re-run." }

  Info "Creating Python venv + installing dependencies..."
  if (Test-Path ".venv") { Remove-Item -Recurse -Force ".venv" | Out-Null }

  if ($py.Kind -eq 'python') { & $py.Cmd[0] -m venv .venv }
  else { & $py.Cmd[0] $py.Cmd[1] -m venv .venv }

  $pip = ".\.venv\Scripts\pip.exe"
  $python = ".\.venv\Scripts\python.exe"

  & $python -m pip install --upgrade pip *> $null
  & $pip install -r requirements.txt *> $null

  if (-not $NoGrafana) { Info "Grafana: http://localhost:3000 (admin/admin first login)" }

  Info "Starting collector (Ctrl+C to stop)..."
  $prev = $ErrorActionPreference
  $ErrorActionPreference = 'Continue'
  & $python .\collector.py 2>&1 | ForEach-Object { Write-Host $_ }
  $code = $LASTEXITCODE
  $ErrorActionPreference = $prev
  if ($code -ne 0) { throw "Collector exited with code $code" }
}

# -------------------- RUN --------------------
Ensure-Prereqs -SkipInstalls:$SkipInstalls
$cfg = Prompt-Config

$resolvedDir = Resolve-ProjectDir
Write-ProjectFiles -Dir $resolvedDir -Cfg $cfg

Ensure-DockerOnline -TimeoutSeconds 240

$composeCmd = Get-ComposeCmd
$projName = (Split-Path -Leaf $resolvedDir) -replace '[^a-zA-Z0-9_-]', ''
if ([string]::IsNullOrWhiteSpace($projName)) { $projName = 'msp-3cx-monitor-demo' }
$env:COMPOSE_PROJECT_NAME = $projName

Compose-DownAndReset -ComposeCmd $composeCmd -Dir $resolvedDir
Invoke-ComposeUp -ComposeCmd $composeCmd

$pgId = Get-PostgresContainerId -ComposeCmd $composeCmd
Wait-PostgresStable -PgId $pgId -ComposeCmd $composeCmd -TimeoutSeconds 240
Ensure-DbAndApplySchema -PgId $pgId -DbName '3cxmon' -User '3cxmon'
Start-Collector