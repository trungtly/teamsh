#!/usr/bin/env python3
"""
Refresh Teams token by launching Edge with an existing profile and
intercepting the OAuth token response via CDP.

On WSL: launches Windows Edge, queries CDP via PowerShell to intercept tokens.
On native: uses Playwright directly.

Usage:
    python3 refresh-token.py [OPTIONS]

Options:
    --edge-path PATH       Path to Edge executable (auto-detected on WSL)
    --profile-dir PATH     Edge User Data directory (auto-detected)
    --profile-name NAME    Edge profile name (default: Default)
    --config-dir PATH      Teamsh config directory (default: ~/.config/teamsh)
    --timeout SECONDS      Max wait for token intercept (default: 120)
    --debug-port PORT      CDP port (default: 9222)

Cron example (every 20 hours):
    0 */20 * * * python3 /path/to/refresh-token.py 2>> /tmp/teamsh-refresh.log
"""

import argparse
import json
import os
import subprocess
import sys
import time
from pathlib import Path

TEAMS_URL = "https://teams.cloud.microsoft"
TOKEN_ENDPOINT = "login.microsoftonline.com"
TOKEN_PATH = "/oauth2/v2.0/token"
POWERSHELL = "/mnt/c/Windows/System32/WindowsPowerShell/v1.0/powershell.exe"


def find_config_dir():
    import platform
    if platform.system() == "Linux":
        xdg = Path.home() / ".config"
    elif platform.system() == "Darwin":
        xdg = Path.home() / "Library" / "Application Support"
    else:
        xdg = Path.home() / ".config"
    return xdg / "teamsh"


def is_wsl():
    try:
        with open("/proc/version", "r") as f:
            return "microsoft" in f.read().lower()
    except FileNotFoundError:
        return False


def find_edge_path():
    candidates = [
        "/mnt/c/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
        "/mnt/c/Program Files/Microsoft/Edge/Application/msedge.exe",
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    return None


def find_windows_user():
    users_dir = Path("/mnt/c/Users")
    if not users_dir.exists():
        return None
    skip = {"All Users", "Default", "Default User", "Public", "defaultuser0", "desktop.ini"}
    for d in users_dir.iterdir():
        if d.is_dir() and d.name not in skip:
            edge_data = d / "AppData" / "Local" / "Microsoft" / "Edge" / "User Data"
            try:
                if edge_data.exists():
                    return d.name
            except PermissionError:
                continue
    return None


def find_edge_profile_dir():
    win_user = find_windows_user()
    if not win_user:
        return None
    return Path(f"/mnt/c/Users/{win_user}/AppData/Local/Microsoft/Edge/User Data")


def wsl_to_windows_path(wsl_path: str) -> str:
    try:
        result = subprocess.run(
            ["wslpath", "-w", wsl_path],
            capture_output=True, text=True, timeout=5,
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass
    if wsl_path.startswith("/mnt/c/"):
        return "C:\\" + wsl_path[7:].replace("/", "\\")
    return wsl_path


def launch_edge_wsl(edge_path, profile_dir, profile_name, debug_port):
    """Launch Windows Edge from WSL with remote debugging.

    Uses the real Edge profile (not a copy) so SSO cookies work.
    Requires Edge to be fully closed first.
    """
    win_profile_dir = wsl_to_windows_path(str(profile_dir))

    cmd = [
        edge_path,
        f"--remote-debugging-port={debug_port}",
        f"--user-data-dir={win_profile_dir}",
        f"--profile-directory={profile_name}",
        "--no-first-run",
        "--no-default-browser-check",
        TEAMS_URL,
    ]

    print(f"Launching Edge (profile: {profile_name}, CDP port: {debug_port})...")
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    return proc


def ensure_edge_closed():
    """Check if Edge is running and ask user to close it."""
    result = ps_run("(Get-Process msedge -ErrorAction SilentlyContinue).Count")
    if result and result.strip() != "0":
        print("Edge is currently running. It must be closed to use the real profile with CDP.")
        print("Close Edge now, or press Enter to force-close it.")
        try:
            input()
        except EOFError:
            pass
        # Force close
        ps_run("Stop-Process -Name msedge -Force -ErrorAction SilentlyContinue")
        time.sleep(2)


def ps_run(command, timeout=30):
    """Run a PowerShell command and return stdout."""
    try:
        result = subprocess.run(
            [POWERSHELL, "-NoProfile", "-Command", command],
            capture_output=True, text=True, timeout=timeout,
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except (subprocess.TimeoutExpired, FileNotFoundError):
        pass
    return None


def ps_cdp_query(port, path="/json/version"):
    """Query CDP endpoint via PowerShell."""
    raw = ps_run(f"(Invoke-WebRequest -Uri 'http://localhost:{port}{path}' -UseBasicParsing -TimeoutSec 5).Content")
    if raw:
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            pass
    return None


def wait_for_cdp_wsl(port, timeout=20):
    """Wait for CDP endpoint via PowerShell bridge."""
    print(f"Waiting for Edge CDP on port {port}...")
    deadline = time.time() + timeout
    while time.time() < deadline:
        data = ps_cdp_query(port)
        if data and data.get("webSocketDebuggerUrl"):
            print(f"CDP ready: {data.get('Browser', '?')}")
            return True
        time.sleep(1)
    return False


def poll_tokens_via_cdp(port, timeout, page_url=None):
    """Use CDP protocol via PowerShell to enable network monitoring and
    intercept OAuth token responses.

    This is the core WSL approach: since Playwright can't reach Windows
    localhost, we drive CDP entirely through PowerShell/Invoke-WebRequest.
    """

    # Step 1: Get list of pages/targets
    targets = ps_cdp_query(port, "/json")
    if not targets:
        print("Error: could not list CDP targets.", file=sys.stderr)
        return None, None

    # Find a Teams page target
    target_ws = None
    for t in targets:
        t_url = t.get("url", "")
        if "teams" in t_url.lower() and t.get("type") == "page":
            target_ws = t.get("webSocketDebuggerUrl")
            print(f"Found Teams tab: {t_url[:80]}")
            break

    if not target_ws:
        # Navigate the first page to Teams
        if targets and targets[0].get("type") == "page":
            target_ws = targets[0].get("webSocketDebuggerUrl")
            target_id = targets[0].get("id", "")
            print("No Teams tab found, navigating...")
            # Use CDP to navigate
            ps_run(
                f"Invoke-WebRequest -Uri 'http://localhost:{port}/json/activate/{target_id}' "
                f"-Method PUT -UseBasicParsing -TimeoutSec 5",
            )

    if not target_ws:
        print("Error: no suitable page target.", file=sys.stderr)
        return None, None

    # Step 2: Use a PowerShell script that connects via WebSocket to CDP,
    # enables Network domain, and watches for token responses.
    # This is the most reliable approach for WSL.

    ps_script = r"""
$ErrorActionPreference = 'Stop'
$wsUrl = '""" + target_ws.replace("'", "''") + r"""'
$timeout = """ + str(timeout) + r"""

# PowerShell 5.1 has System.Net.WebSockets
$ws = New-Object System.Net.WebSockets.ClientWebSocket
$ct = New-Object System.Threading.CancellationToken

try {
    $ws.ConnectAsync([uri]$wsUrl, $ct).Wait()
} catch {
    Write-Error "WebSocket connect failed: $_"
    exit 1
}

# Helper to send CDP command
function Send-CDP($id, $method, $params = @{}) {
    $msg = @{ id = $id; method = $method; params = $params } | ConvertTo-Json -Compress -Depth 10
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($msg)
    $segment = New-Object System.ArraySegment[byte] -ArgumentList @(,$bytes)
    $ws.SendAsync($segment, [System.Net.WebSockets.WebSocketMessageType]::Text, $true, $ct).Wait()
}

# Helper to receive message
function Recv-CDP() {
    $buf = New-Object byte[] 1048576  # 1MB buffer
    $segment = New-Object System.ArraySegment[byte] -ArgumentList @(,$buf)
    $result = $ws.ReceiveAsync($segment, $ct).Result
    if ($result.Count -gt 0) {
        return [System.Text.Encoding]::UTF8.GetString($buf, 0, $result.Count)
    }
    return $null
}

# Enable Network monitoring and response body retrieval
Send-CDP -id 1 -method 'Network.enable'

# Navigate/reload to trigger token refresh
Send-CDP -id 2 -method 'Page.reload'

$deadline = (Get-Date).AddSeconds($timeout)
$pendingRequestIds = @{}  # requestId -> url

while ((Get-Date) -lt $deadline) {
    try {
        $raw = Recv-CDP
        if (-not $raw) { continue }

        $msg = $raw | ConvertFrom-Json -ErrorAction SilentlyContinue
        if (-not $msg) { continue }

        # Track requests to token endpoint
        if ($msg.method -eq 'Network.requestWillBeSent') {
            $url = $msg.params.request.url
            if ($url -match 'login\.microsoftonline\.com' -and $url -match 'oauth2/v2\.0/token') {
                $rid = $msg.params.requestId
                $pendingRequestIds[$rid] = $url
            }
        }

        # When response is received for a tracked request, get the body
        if ($msg.method -eq 'Network.responseReceived') {
            $rid = $msg.params.requestId
            if ($pendingRequestIds.ContainsKey($rid)) {
                # Request the response body
                Send-CDP -id 100 -method 'Network.getResponseBody' -params @{ requestId = $rid }
            }
        }

        # Check response to our getResponseBody call
        if ($msg.id -eq 100 -and $msg.result) {
            $bodyText = $msg.result.body
            if ($bodyText) {
                $body = $bodyText | ConvertFrom-Json -ErrorAction SilentlyContinue
                if ($body -and $body.refresh_token) {
                    # Output as JSON for the Python script to parse
                    $output = @{
                        refresh_token = $body.refresh_token
                        url = ($pendingRequestIds.Values | Select-Object -First 1)
                    } | ConvertTo-Json -Compress
                    Write-Output $output
                    break
                }
            }
        }
    } catch {
        # Timeout or partial read, continue
        Start-Sleep -Milliseconds 100
    }
}

$ws.CloseAsync([System.Net.WebSockets.WebSocketCloseStatus]::NormalClosure, '', $ct).Wait()
"""

    print("Monitoring network via CDP (waiting for token refresh)...")
    raw = ps_run(ps_script, timeout=timeout + 30)

    if raw:
        # Parse each line looking for our JSON output
        for line in raw.splitlines():
            line = line.strip()
            if not line or not line.startswith("{"):
                continue
            try:
                data = json.loads(line)
                token = data.get("refresh_token")
                url = data.get("url", "")
                tenant = None
                if url:
                    try:
                        parts = url.split("/")
                        idx = parts.index("oauth2") - 1
                        if idx >= 0:
                            tenant = parts[idx]
                    except (ValueError, IndexError):
                        pass
                if token:
                    return token, tenant
            except json.JSONDecodeError:
                continue

    return None, None


def main():
    parser = argparse.ArgumentParser(description="Refresh Teams token via Edge browser")
    parser.add_argument("--edge-path", type=str, default=None,
                        help="Path to Edge executable (auto-detected on WSL)")
    parser.add_argument("--profile-dir", type=str, default=None,
                        help="Edge User Data directory (auto-detected)")
    parser.add_argument("--profile-name", type=str, default="Default",
                        help="Edge profile name (default: Default)")
    parser.add_argument("--config-dir", type=str, default=None,
                        help="Teamsh config directory (default: ~/.config/teamsh)")
    parser.add_argument("--timeout", type=int, default=120,
                        help="Max seconds to wait for token (default: 120)")
    parser.add_argument("--debug-port", type=int, default=9222,
                        help="CDP remote debugging port (default: 9222)")
    args = parser.parse_args()

    config_dir = Path(args.config_dir) if args.config_dir else find_config_dir()
    config_dir.mkdir(parents=True, exist_ok=True)

    refresh_token_path = config_dir / "refresh_token"
    tenant_id_path = config_dir / "tenant_id"

    edge_path = args.edge_path
    profile_dir = Path(args.profile_dir) if args.profile_dir else None

    on_wsl = is_wsl()
    edge_proc = None

    if on_wsl:
        if not edge_path:
            edge_path = find_edge_path()
        if not edge_path:
            print("Error: Edge not found. Specify --edge-path.", file=sys.stderr)
            sys.exit(1)

        if not profile_dir:
            profile_dir = find_edge_profile_dir()
        if not profile_dir:
            print("Error: Edge profile not found. Specify --profile-dir.", file=sys.stderr)
            sys.exit(1)

        ensure_edge_closed()
        edge_proc = launch_edge_wsl(edge_path, profile_dir, args.profile_name, args.debug_port)

        if not wait_for_cdp_wsl(args.debug_port):
            print("Error: Edge did not expose CDP endpoint.", file=sys.stderr)
            print("Close all Edge windows and try again, or use a different --debug-port.", file=sys.stderr)
            edge_proc.terminate()
            sys.exit(1)

        # Use CDP via PowerShell (bypasses WSL2 network isolation)
        captured_token, captured_tenant = poll_tokens_via_cdp(args.debug_port, args.timeout)

    else:
        # Native (macOS/Linux): use Playwright directly
        try:
            from playwright.sync_api import sync_playwright
        except ImportError:
            print("Error: playwright not installed. Run: pip install playwright && playwright install chromium",
                  file=sys.stderr)
            sys.exit(1)

        captured_token = None
        captured_tenant = None

        def handle_response(response):
            nonlocal captured_token, captured_tenant
            if captured_token:
                return
            url = response.url
            if TOKEN_ENDPOINT not in url or TOKEN_PATH not in url:
                return
            try:
                body = response.json()
            except Exception:
                return
            rt = body.get("refresh_token")
            if not rt:
                return
            captured_token = rt
            try:
                parts = url.split("/")
                idx = parts.index("oauth2") - 1
                if idx >= 0:
                    captured_tenant = parts[idx]
            except (ValueError, IndexError):
                pass
            print(f"Token captured! (tenant: {captured_tenant or 'unknown'})")

        with sync_playwright() as p:
            context = p.chromium.launch_persistent_context(
                str(profile_dir or config_dir / "browser"),
                channel="msedge",
                headless=False,
                args=["--disable-blink-features=AutomationControlled"],
            )
            teams_page = context.pages[0] if context.pages else context.new_page()
            teams_page.on("response", handle_response)
            teams_page.goto(TEAMS_URL, wait_until="domcontentloaded", timeout=30000)
            teams_page.reload(wait_until="domcontentloaded", timeout=30000)

            deadline = time.time() + args.timeout
            while not captured_token and time.time() < deadline:
                teams_page.wait_for_timeout(500)
            context.close()

    # Clean up Edge
    if edge_proc:
        edge_proc.terminate()
        try:
            edge_proc.wait(timeout=5)
        except subprocess.TimeoutExpired:
            edge_proc.kill()

    if not captured_token:
        print("Error: timed out waiting for token.", file=sys.stderr)
        print("Make sure you're logged into Teams in Edge.", file=sys.stderr)
        sys.exit(1)

    print(f"Token captured! (tenant: {captured_tenant or 'unknown'})")

    # Save
    refresh_token_path.write_text(captured_token)
    print(f"Saved refresh_token to {refresh_token_path}")

    if captured_tenant and captured_tenant != "organizations":
        tenant_id_path.write_text(captured_tenant)
        print(f"Saved tenant_id to {tenant_id_path}")

    # Verify
    try:
        result = subprocess.run(
            ["teamsh", "auth", "test"],
            capture_output=True, text=True, timeout=15,
        )
        if result.returncode == 0:
            print("Token verified OK!")
        else:
            print(f"Warning: token verify failed: {result.stderr.strip()}")
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass

    print("Done.")


if __name__ == "__main__":
    main()
