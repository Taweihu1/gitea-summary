"""
Launch Chrome on CDP port 9222 using a dedicated ChromeClaudeDP user-data-dir.
Called by run_summary.bat. Python subprocess correctly quotes arguments
with spaces in --user-data-dir regardless of the calling environment.
Skips relaunch if port 9222 is already open.
"""
import http.client, os, subprocess, time, sys

CHROME   = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
CDP_PORT = 9222
# Dedicated no-space directory avoids Chrome CDP bind failures that occur
# with the default "User Data" path (spaces + existing profile state).
UD_DIR = os.path.join(os.environ["USERPROFILE"], "ChromeClaudeDP")
LOCK_FILES = ["SingletonLock", "SingletonCookie", "SingletonSocket", "lockfile"]

# Always kill all Chrome processes and wait until fully gone
print("  Closing all Chrome windows...")
subprocess.run(["taskkill", "/IM", "chrome.exe", "/F"], capture_output=True)
for _ in range(30):
    r = subprocess.run(["tasklist", "/FI", "IMAGENAME eq chrome.exe"],
                       capture_output=True, text=True)
    if "chrome.exe" not in r.stdout:
        break
    time.sleep(1)
print("  All Chrome processes exited.")

# Clear singleton locks left by force-kill (prevent CDP bind failure)
for f in LOCK_FILES:
    path = os.path.join(UD_DIR, f)
    try:
        os.remove(path)
    except FileNotFoundError:
        pass
print("  Lock files cleared.")

# Launch Chrome — subprocess list form handles spaces in path correctly
print(f"  Launching Chrome (ChromeClaudeDP, port {CDP_PORT})...")
subprocess.Popen([
    CHROME,
    f"--remote-debugging-port={CDP_PORT}",
    f"--user-data-dir={UD_DIR}",
    "--start-minimized",
])
