#!/usr/bin/env python3
"""
Refresh both Teams and Graph tokens using agent-browser (Playwright CLI).

First run:  ./refresh-token.py --headed   (log in manually, SSO cookies saved)
Later runs: ./refresh-token.py            (headless, auto-login via saved cookies)

How it works:
  Teams token:
    1. Opens Teams in Chromium with persistent profile
    2. SSO cookies auto-login (or manual login if --headed)
    3. Extracts refresh token from MSAL browser storage
    4. Saves to ~/.config/teamsh/refresh_token

  Graph token:
    5. Starts device code flow (HTTP POST to get user_code + device_code)
    6. Navigates browser to microsoft.com/devicelogin
    7. Auto-types user_code, clicks through approval
    8. Polls token endpoint until Graph refresh token is received
    9. Saves to ~/.config/teamsh/graph_refresh_token

Cron example (every 12 hours):
    0 */12 * * * /path/to/refresh-token.py 2>> /tmp/teamsh-refresh.log

Requirements: npx (Node.js), agent-browser (auto-installed via npx)
"""

import argparse
import json
import os
import re
import subprocess
import sys
import threading
import time
import urllib.request
import urllib.parse
from pathlib import Path

TEAMS_URL = "https://teams.cloud.microsoft"
TEAMS_CLIENT_ID = "5e3ce6c0-2b1f-4285-8d4b-75ee78787346"
GRAPH_CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
GRAPH_SCOPE = "https://graph.microsoft.com/.default openid profile offline_access"
DEVICE_LOGIN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/devicecode"
TOKEN_URL_TEMPLATE = "https://login.microsoftonline.com/{}/oauth2/v2.0/token"

CONFIG_DIR = Path.home() / ".config" / "teamsh"


def ab(*args, profile=None, headed=False, capture=True):
    """Run an agent-browser command."""
    cmd = ["npx", "agent-browser"]
    if profile:
        cmd += ["--profile", str(profile)]
    if headed:
        cmd += ["--headed"]
    cmd += list(args)

    if capture:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
        # Strip ANSI codes
        import re
        clean = re.sub(r'\x1b\[[0-9;]*m', '', result.stdout).strip()
        return clean
    else:
        subprocess.run(cmd, timeout=30)
        return None


def ab_eval(js, profile=None, headed=False):
    """Run JavaScript via agent-browser eval --stdin."""
    cmd = ["npx", "agent-browser"]
    if profile:
        cmd += ["--profile", str(profile)]
    if headed:
        cmd += ["--headed"]
    cmd += ["eval", "--stdin"]

    result = subprocess.run(cmd, input=js, capture_output=True, text=True, timeout=30)
    import re
    clean = re.sub(r'\x1b\[[0-9;]*m', '', result.stdout).strip()
    return clean


def extract_teams_token(profile, headed, timeout):
    """Extract Teams refresh token from MSAL browser storage."""
    print("=== Teams Token ===")
    print("Opening Teams...")
    ab("open", TEAMS_URL, profile=profile, headed=headed)

    # Wait for Teams to load past login page
    print(f"Waiting for Teams to load (timeout: {timeout}s)...")
    start = time.time()
    logged_login = False

    while time.time() - start < timeout:
        url = ab("get", "url", profile=profile)
        if not url:
            time.sleep(2)
            continue

        # Check if we're on Teams (not login page)
        if "teams.cloud.microsoft" in url and "login.microsoftonline.com" not in url:
            elapsed = int(time.time() - start)
            print(f"Teams loaded ({elapsed}s)")
            break

        # On login page
        if "login.microsoftonline.com" in url:
            if not headed:
                print("Error: redirected to login. Run with --headed to log in manually.")
                return None
            if not logged_login:
                print("  On login page - please log in in the browser window...")
                logged_login = True

        time.sleep(2)
    else:
        print("Error: timed out waiting for Teams to load.")
        return None

    # Let MSAL finish storing tokens
    print("Waiting for MSAL token storage...")
    time.sleep(5)

    # Extract refresh token from browser storage
    print("Extracting token from browser storage...")
    js = """(() => {
    const results = [];
    for (const store of [sessionStorage, localStorage]) {
        for (let i = 0; i < store.length; i++) {
            const key = store.key(i);
            if (!key || !key.toLowerCase().includes('refreshtoken')) continue;
            try {
                const val = JSON.parse(store.getItem(key));
                if (val && val.secret) {
                    results.push({
                        refresh_token: val.secret,
                        client_id: val.clientId || null,
                        tenant_id: val.realm || null,
                    });
                }
            } catch(e) {}
        }
    }
    return JSON.stringify(results);
})()"""

    raw = ab_eval(js, profile=profile, headed=headed)

    # Parse JSON from output — agent-browser double-serializes:
    # JS returns JSON.stringify(array) → agent-browser wraps in quotes → "[{\"key\":...}]"
    tokens = None
    # Try 1: parse as JSON string (unwraps one layer), then parse inner
    try:
        inner = json.loads(raw)
        if isinstance(inner, str):
            tokens = json.loads(inner)
        elif isinstance(inner, list):
            tokens = inner
    except (json.JSONDecodeError, TypeError):
        pass
    # Try 2: regex extract the array
    if not tokens:
        match = re.search(r'\[.*\]', raw, re.DOTALL)
        if match:
            try:
                tokens = json.loads(match.group())
            except json.JSONDecodeError:
                pass

    if not tokens:
        print(f"Error: no tokens found in browser storage.")
        print(f"  Raw output: {raw[:300]}")
        return None

    # Prefer the Teams SPA client token
    teams_token = None
    for t in tokens:
        if t.get("client_id") == TEAMS_CLIENT_ID:
            teams_token = t
            break
    if not teams_token:
        teams_token = tokens[0]

    refresh_token = teams_token["refresh_token"]
    tenant_id = teams_token.get("tenant_id")

    # Save
    (CONFIG_DIR / "refresh_token").write_text(refresh_token)
    print(f"Saved Teams refresh_token ({len(refresh_token)} chars)")

    if tenant_id and tenant_id not in ("organizations", "common"):
        (CONFIG_DIR / "tenant_id").write_text(tenant_id)
        print(f"Saved tenant_id: {tenant_id}")

    return True


def get_tenant_id():
    """Load tenant_id from config."""
    path = CONFIG_DIR / "tenant_id"
    if path.exists():
        tid = path.read_text().strip()
        if tid:
            return tid
    return "organizations"


def start_device_code_flow(tenant_id):
    """Start OAuth2 device code flow for Graph API."""
    data = urllib.parse.urlencode({
        "client_id": GRAPH_CLIENT_ID,
        "scope": GRAPH_SCOPE,
    }).encode()

    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/devicecode"
    req = urllib.request.Request(url, data=data, method="POST")
    with urllib.request.urlopen(req, timeout=15) as resp:
        return json.loads(resp.read())


def poll_device_code_token(tenant_id, device_code, interval, timeout):
    """Poll token endpoint until device code flow completes."""
    url = TOKEN_URL_TEMPLATE.format(tenant_id)
    deadline = time.time() + timeout

    while time.time() < deadline:
        time.sleep(interval)
        data = urllib.parse.urlencode({
            "client_id": GRAPH_CLIENT_ID,
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "device_code": device_code,
            "scope": GRAPH_SCOPE,
        }).encode()

        req = urllib.request.Request(url, data=data, method="POST")
        try:
            with urllib.request.urlopen(req, timeout=15) as resp:
                body = json.loads(resp.read())
                return body
        except urllib.error.HTTPError as e:
            body = json.loads(e.read())
            error = body.get("error", "")
            if error == "authorization_pending":
                continue
            elif error == "slow_down":
                interval += 5
                continue
            elif error == "expired_token":
                print("Error: device code expired.")
                return None
            else:
                desc = body.get("error_description", "")
                print(f"Error: {error} - {desc}")
                return None

    print("Error: timed out polling for Graph token.")
    return None


def click_by_snapshot(profile, label_keywords):
    """Use accessibility snapshot to find and click a button by keyword."""
    snap = ab("snapshot", "-i", profile=profile)
    if not snap:
        return False, ""

    # Look for refs like @e3 next to matching text
    import re
    for line in snap.splitlines():
        line_lower = line.lower()
        for kw in label_keywords:
            if kw in line_lower:
                ref_match = re.search(r'@e\d+', line)
                if ref_match:
                    ref = ref_match.group()
                    ab("click", ref, profile=profile)
                    return True, snap
    return False, snap


def extract_graph_token(profile, headed, timeout):
    """Get Graph refresh token via device code flow automated in the browser."""
    print("\n=== Graph Token ===")

    tenant_id = get_tenant_id()

    # Start device code flow
    print("Starting device code flow...")
    try:
        device = start_device_code_flow(tenant_id)
    except Exception as e:
        print(f"Error starting device code flow: {e}")
        return False

    user_code = device["user_code"]
    device_code = device["device_code"]
    interval = device.get("interval", 5)

    print(f"User code: {user_code}")

    # Start polling in a background thread so browser interaction isn't blocked
    token_result = [None]

    def poll_thread():
        token_result[0] = poll_device_code_token(tenant_id, device_code, interval, timeout)

    poller = threading.Thread(target=poll_thread, daemon=True)
    poller.start()

    # Navigate to device login
    login_url = device.get("verification_uri", "https://login.microsoft.com/device")
    print(f"Navigating to {login_url}...")
    ab("open", login_url, profile=profile, headed=headed)
    time.sleep(3)

    # Step 1: Enter the user code
    # Use snapshot to find the input field
    print(f"Entering code: {user_code}")
    snap = ab("snapshot", "-i", profile=profile)
    print(f"  Page: {snap[:200]}")

    # Try filling by common selectors, then fallback to snapshot ref
    filled = False
    for selector in ["input#otc", "input[name='otc']", "input"]:
        result = ab("fill", selector, user_code, profile=profile)
        # Check if fill worked (agent-browser doesn't error on bad selectors,
        # but returns empty or error text)
        if result is not None and "error" not in result.lower():
            filled = True
            break

    if not filled:
        # Try via snapshot ref
        for line in snap.splitlines():
            if "textbox" in line.lower() or "input" in line.lower():
                ref_match = re.search(r'@e\d+', line)
                if ref_match:
                    ab("fill", ref_match.group(), user_code, profile=profile)
                    filled = True
                    break

    time.sleep(1)

    # Step 2: Click Next
    print("Clicking Next...")
    ab("click", "input#idSIButton9", profile=profile)
    time.sleep(3)

    # Step 3: Handle approval pages
    # Microsoft shows multiple screens: account picker, consent, etc.
    # Keep clicking through until done or timed out
    print("Completing approval flow...")
    deadline = time.time() + 60  # 60s for the approval part

    while time.time() < deadline:
        # Check if poller already got the token
        if token_result[0] is not None:
            print("Token received!")
            break

        url = ab("get", "url", profile=profile)
        snap = ab("snapshot", "-i", profile=profile)
        snap_lower = (snap or "").lower()

        print(f"  URL: {(url or '')[:80]}")
        if snap:
            # Print first few lines of snapshot for debugging
            snap_lines = snap.strip().splitlines()[:5]
            for sl in snap_lines:
                print(f"    {sl}")

        # Success page
        if snap and ("you have signed in" in snap_lower or "you're all set" in snap_lower
                      or "you may close" in snap_lower):
            print("  Approval complete!")
            break

        # Error
        if snap and ("error" in snap_lower and "expired" in snap_lower):
            print(f"  Error on page: {snap[:200]}")
            break

        # Try clicking interactive elements on the page
        clicked = False
        if snap:
            for line in snap.splitlines():
                line_lower = line.lower()
                # Account picker: "Sign in with ... account"
                if "sign in with" in line_lower:
                    ref_match = re.search(r'\[ref=(e\d+)\]', line)
                    if ref_match:
                        ref = f"@{ref_match.group(1)}"
                        print(f"  Clicking account: {line.strip()}")
                        ab("click", ref, profile=profile)
                        clicked = True
                        break
                # Continue/Accept/Yes buttons
                for kw in ["continue", "accept", "yes"]:
                    if kw in line_lower and "button" in line_lower:
                        ref_match = re.search(r'\[ref=(e\d+)\]', line)
                        if ref_match:
                            ref = f"@{ref_match.group(1)}"
                            print(f"  Clicking: {line.strip()}")
                            ab("click", ref, profile=profile)
                            clicked = True
                            break
                if clicked:
                    break

        if clicked:
            time.sleep(3)
        else:
            time.sleep(3)

    # Wait for poller to finish (give it a bit more time)
    print("Waiting for token polling to complete...")
    poller.join(timeout=30)

    if not token_result[0] or "refresh_token" not in token_result[0]:
        print("Error: did not receive Graph refresh token.")
        return False

    # Save
    (CONFIG_DIR / "graph_refresh_token").write_text(token_result[0]["refresh_token"])
    print(f"Saved Graph refresh_token ({len(token_result[0]['refresh_token'])} chars)")

    return True


def verify_tokens():
    """Verify tokens work by refreshing them via teamsh."""
    print("\n=== Verification ===")
    try:
        result = subprocess.run(
            ["teamsh", "auth", "test"],
            capture_output=True, text=True, timeout=15,
        )
        if result.returncode == 0:
            print("Teams token: OK")
        else:
            print(f"Teams token: FAILED ({result.stderr.strip()})")
    except (FileNotFoundError, subprocess.TimeoutExpired):
        print("Teams token: skipped (teamsh not in PATH)")


def main():
    parser = argparse.ArgumentParser(
        description="Refresh Teams and Graph tokens via browser automation"
    )
    parser.add_argument("--headed", action="store_true",
                        help="Show browser window (required for first login)")
    parser.add_argument("--timeout", type=int, default=120,
                        help="Max wait in seconds per token (default: 120)")
    parser.add_argument("--teams-only", action="store_true",
                        help="Only refresh Teams token")
    parser.add_argument("--graph-only", action="store_true",
                        help="Only refresh Graph token")
    args = parser.parse_args()

    CONFIG_DIR.mkdir(parents=True, exist_ok=True)

    profile = CONFIG_DIR / "browser-profile"
    ok = True

    try:
        if not args.graph_only:
            if not extract_teams_token(profile, args.headed, args.timeout):
                print("Teams token refresh FAILED.")
                ok = False

        if not args.teams_only:
            if not extract_graph_token(profile, args.headed, args.timeout):
                print("Graph token refresh FAILED.")
                ok = False

        if ok:
            verify_tokens()
            print("\nDone. Both tokens refreshed.")
        else:
            print("\nDone with errors.")
    finally:
        # Close browser
        print("Closing browser...")
        try:
            ab("close", profile=profile)
        except Exception:
            pass

    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
