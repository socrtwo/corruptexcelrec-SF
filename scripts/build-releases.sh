#!/usr/bin/env bash
# build-releases.sh — produce per-platform release bundles for the PWA.
# Usage:  bash scripts/build-releases.sh <version>
#
# Output: dist/corruptexcelrec-<platform>-v<version>.{zip,tar.gz}
#         dist/RELEASE_NOTES.md
#
# Each bundle contains:
#   - The full PWA (web/) so it works offline
#   - A platform-appropriate launcher that opens the app
#   - PLATFORM_INSTALL.md with install steps for that platform
#
# Native builds (Windows .exe) are produced by the workflow on a Windows runner
# and uploaded separately.

set -euo pipefail

VERSION="${1:-0.0.0}"
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
DIST="$ROOT/dist"
WEB="$ROOT/web"

rm -rf "$DIST"
mkdir -p "$DIST"

if [ ! -d "$WEB" ]; then
  echo "FATAL: web/ directory missing" >&2
  exit 1
fi

stage() {
  local name="$1"
  local stagedir="$DIST/_stage/$name"
  rm -rf "$stagedir"
  mkdir -p "$stagedir/app"
  cp -r "$WEB"/. "$stagedir/app/"
  echo "$stagedir"
}

pkg_zip() {
  local name="$1" stagedir="$2"
  ( cd "$DIST/_stage" && zip -qr "$DIST/$name.zip" "$(basename "$stagedir")" )
  echo "  ✓ $name.zip"
}

pkg_tgz() {
  local name="$1" stagedir="$2"
  tar -C "$DIST/_stage" -czf "$DIST/$name.tar.gz" "$(basename "$stagedir")"
  echo "  ✓ $name.tar.gz"
}

# ----------------------------------------------------------------------
# Windows
# ----------------------------------------------------------------------
echo "▶ Windows"
S=$(stage "corruptexcelrec-windows-v$VERSION")
cat > "$S/Launch Excel Recovery.bat" <<'BAT'
@echo off
REM Launches the PWA in the user's default browser.
REM For best experience, install the app via the browser's install button
REM (URL bar, or browser menu -> "Install S2 Recovery Tools for Microsoft Excel").
setlocal
set "HERE=%~dp0app\index.html"
start "" "%HERE%"
endlocal
BAT
cat > "$S/PLATFORM_INSTALL.md" <<EOF
# S2 Recovery Tools for Microsoft Excel — Windows

## Quick start
1. Unzip this archive anywhere.
2. Double-click **Launch Excel Recovery.bat**.
3. Use it.

## Install as a real Windows app (recommended)
1. Open https://socrtwo.github.io/corruptexcelrec-SF/ in **Microsoft Edge** or **Chrome**.
2. Click the **install** icon in the address bar (or menu → "Install Excel Recovery").
3. The app appears in your Start menu and runs in its own window — works offline.

## Native legacy build
The classic VB.NET WinForms binary (.exe) is published as a separate artifact
\`corruptexcelrec-windows-native.zip\` on the same release page. It uses
Microsoft Excel COM Interop to access more advanced recovery routines and is
**Windows-only**.

Version: $VERSION
EOF
pkg_zip "corruptexcelrec-windows-v$VERSION" "$S"

# ----------------------------------------------------------------------
# macOS
# ----------------------------------------------------------------------
echo "▶ macOS"
S=$(stage "corruptexcelrec-macos-v$VERSION")
cat > "$S/Launch Excel Recovery.command" <<'CMD'
#!/usr/bin/env bash
HERE="$(cd "$(dirname "$0")" && pwd)"
open "$HERE/app/index.html"
CMD
chmod +x "$S/Launch Excel Recovery.command"
cat > "$S/PLATFORM_INSTALL.md" <<EOF
# S2 Recovery Tools for Microsoft Excel — macOS

## Quick start
1. Unzip this archive.
2. Double-click **Launch Excel Recovery.command**.
   (Right-click → Open the first time, to bypass Gatekeeper.)

## Install as a Mac app (recommended)
- **Safari 17+**: open https://socrtwo.github.io/corruptexcelrec-SF/, then **File → Add to Dock**.
- **Chrome / Edge / Arc**: open the URL, click the install button in the address bar.

The installed app runs in its own window, supports offline use, and shows up in
Spotlight, the Dock, and Launchpad like any native macOS application.

Version: $VERSION
EOF
pkg_zip "corruptexcelrec-macos-v$VERSION" "$S"

# ----------------------------------------------------------------------
# Linux
# ----------------------------------------------------------------------
echo "▶ Linux"
S=$(stage "corruptexcelrec-linux-v$VERSION")
cat > "$S/launch-excel-recovery.sh" <<'SH'
#!/usr/bin/env bash
HERE="$(cd "$(dirname "$0")" && pwd)"
URL="file://$HERE/app/index.html"
if   command -v xdg-open >/dev/null 2>&1; then xdg-open "$URL"
elif command -v gio      >/dev/null 2>&1; then gio open "$URL"
elif command -v firefox  >/dev/null 2>&1; then firefox  "$URL"
elif command -v chromium >/dev/null 2>&1; then chromium "$URL"
else echo "Open this URL in your browser: $URL"; fi
SH
chmod +x "$S/launch-excel-recovery.sh"
cat > "$S/excel-recovery.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=S2 Recovery Tools for Microsoft Excel
Comment=Recover and repair corrupt .xls / .xlsx files
Exec=bash -c "\$(dirname %k)/launch-excel-recovery.sh"
Icon=spreadsheet
Terminal=false
Categories=Office;Spreadsheet;Utility;
EOF
cat > "$S/PLATFORM_INSTALL.md" <<EOF
# S2 Recovery Tools for Microsoft Excel — Linux

## Quick start
1. Extract: \`tar -xzf corruptexcelrec-linux-v$VERSION.tar.gz\`
2. Run: \`./corruptexcelrec-linux-v$VERSION/launch-excel-recovery.sh\`

## Install a desktop entry
\`\`\`bash
DEST="\$HOME/.local/share/excel-recovery"
mkdir -p "\$DEST"
cp -r corruptexcelrec-linux-v$VERSION/* "\$DEST/"
sed -i "s|Exec=.*|Exec=bash \$DEST/launch-excel-recovery.sh|" "\$DEST/excel-recovery.desktop"
desktop-file-install --dir="\$HOME/.local/share/applications" "\$DEST/excel-recovery.desktop"
\`\`\`

## Install as a real Linux app (recommended)
Open https://socrtwo.github.io/corruptexcelrec-SF/ in **Chrome / Chromium /
Brave / Edge** and click the install button in the address bar. The PWA
integrates with your application launcher and works offline.

Version: $VERSION
EOF
pkg_tgz "corruptexcelrec-linux-v$VERSION" "$S"

# ----------------------------------------------------------------------
# ChromeOS
# ----------------------------------------------------------------------
echo "▶ ChromeOS"
S=$(stage "corruptexcelrec-chromeos-v$VERSION")
cat > "$S/PLATFORM_INSTALL.md" <<EOF
# S2 Recovery Tools for Microsoft Excel — ChromeOS

ChromeOS treats PWAs as first-class apps — no APK, no extension required.

## Install (recommended)
1. Open https://socrtwo.github.io/corruptexcelrec-SF/ in Chrome.
2. Click ⋮ → **Install Excel Recovery** (or the install icon in the address bar).
3. The app appears in your launcher (search "Excel Recovery") and works offline.

## Run from this bundle (no internet)
1. Unzip the archive into Files → My files.
2. Open **app/index.html** in Chrome (right-click → Open with → Chrome).

## Power users
The included \`app/\` directory is the full PWA. To self-host on a Chromebook
running Linux apps, drop \`app/\` into any static web server (e.g. \`python3
-m http.server\`) and visit it from Chrome.

Version: $VERSION
EOF
pkg_zip "corruptexcelrec-chromeos-v$VERSION" "$S"

# ----------------------------------------------------------------------
# Android
# ----------------------------------------------------------------------
echo "▶ Android"
S=$(stage "corruptexcelrec-android-v$VERSION")
cat > "$S/PLATFORM_INSTALL.md" <<EOF
# S2 Recovery Tools for Microsoft Excel — Android

## Install (recommended)
1. Open https://socrtwo.github.io/corruptexcelrec-SF/ in **Chrome** for Android.
2. Tap ⋮ → **Install app** (or "Add to Home screen").
3. The app installs like any other Android app: launcher icon, splash screen,
   no browser chrome, offline support, and a "Share" target so any spreadsheet
   sent from another app opens directly in Excel Recovery.

## Build your own APK (optional, advanced)
The PWA can be wrapped as a real Android Package via Google's
[PWABuilder](https://www.pwabuilder.com/):

1. Visit https://www.pwabuilder.com/
2. Enter the URL: \`https://socrtwo.github.io/corruptexcelrec-SF/\`
3. Click **Package for stores → Android → Generate Package**.
4. PWABuilder produces a signed APK / AAB you can sideload or upload to Play.

This bundle ships the unwrapped PWA in \`app/\` for reference and offline use.

Version: $VERSION
EOF
pkg_zip "corruptexcelrec-android-v$VERSION" "$S"

# ----------------------------------------------------------------------
# iOS / iPadOS
# ----------------------------------------------------------------------
echo "▶ iOS"
S=$(stage "corruptexcelrec-ios-v$VERSION")
cat > "$S/PLATFORM_INSTALL.md" <<EOF
# S2 Recovery Tools for Microsoft Excel — iOS / iPadOS

iOS does not allow third-party app stores or sideloading without Xcode + an
Apple Developer account. The PWA is the recommended distribution channel.

## Install (recommended)
1. Open https://socrtwo.github.io/corruptexcelrec-SF/ in **Safari** (this
   does not work in Chrome on iOS — Apple restricts PWA install to Safari).
2. Tap the **Share** button.
3. Tap **Add to Home Screen**.
4. The app appears on your Home Screen, runs in standalone mode, and works
   offline (cached service worker).

## Build a native IPA (optional, advanced)
- Use [Capacitor](https://capacitorjs.com/) to wrap the \`app/\` directory in
  this bundle and produce a real iOS app via Xcode.
- An Apple Developer Program membership (\$99/yr) is required to sign and
  distribute outside TestFlight.

Version: $VERSION
EOF
pkg_zip "corruptexcelrec-ios-v$VERSION" "$S"

# ----------------------------------------------------------------------
# Web (hosted)
# ----------------------------------------------------------------------
echo "▶ Web"
S=$(stage "corruptexcelrec-web-v$VERSION")
cat > "$S/PLATFORM_INSTALL.md" <<EOF
# S2 Recovery Tools for Microsoft Excel — Web

## Use without installing
Open https://socrtwo.github.io/corruptexcelrec-SF/ in any modern browser.

## Self-host
Drop the contents of \`app/\` into any static web server (Apache, nginx,
GitHub Pages, Netlify, Cloudflare Pages, Vercel, S3 + CloudFront…).

\`\`\`bash
cd app
python3 -m http.server 8080
# visit http://localhost:8080
\`\`\`

The app is 100% client-side — no backend, no database, no uploads. Files
opened in the recovery tool never leave the user's device.

Version: $VERSION
EOF
pkg_zip "corruptexcelrec-web-v$VERSION" "$S"

# ----------------------------------------------------------------------
# Cleanup + release notes
# ----------------------------------------------------------------------
rm -rf "$DIST/_stage"

cat > "$DIST/RELEASE_NOTES.md" <<EOF
# S2 Recovery Tools for Microsoft Excel — v$VERSION

A modernized, **fully cross-platform** release of the classic Windows
Excel recovery tool. The same PWA codebase runs everywhere; native bundles
are also published where the platform supports them.

## 📦 Downloads

| Platform | Bundle | What's inside |
|----------|--------|---------------|
| 🪟 Windows           | \`corruptexcelrec-windows-v$VERSION.zip\`         | PWA + .bat launcher |
| 🪟 Windows (native)  | \`corruptexcelrec-windows-native.zip\`            | Classic VB.NET .exe + DLLs |
| 🍎 macOS             | \`corruptexcelrec-macos-v$VERSION.zip\`           | PWA + .command launcher |
| 🐧 Linux             | \`corruptexcelrec-linux-v$VERSION.tar.gz\`        | PWA + .sh launcher + .desktop |
| 🟢 ChromeOS          | \`corruptexcelrec-chromeos-v$VERSION.zip\`        | PWA + install instructions |
| 🤖 Android           | \`corruptexcelrec-android-v$VERSION.zip\`         | PWA + APK build instructions |
| 📱 iOS / iPadOS      | \`corruptexcelrec-ios-v$VERSION.zip\`             | PWA + iOS install instructions |
| 🌐 Web (hosted)      | \`corruptexcelrec-web-v$VERSION.zip\`             | PWA static site for self-hosting |

The hosted web app lives at <https://socrtwo.github.io/corruptexcelrec-SF/> —
on every supported platform you can install it directly from your browser
(no download required).

## ✨ What's new in v$VERSION

- Brand-new in-browser Excel recovery PWA — works offline, never uploads files.
- Six recovery strategies: Auto, Strict Read, Lenient XML Repair, ZIP Recovery,
  Salvage Cells, Convert to CSV.
- Saves recovered workbooks as **.xlsx**, **.xls**, **.ods**, or **.csv**.
- Installable on Windows, macOS, Linux, ChromeOS, Android, iOS, and the web.
- File-handler + share-target manifest entries — open .xlsx files directly
  with the app on supported platforms.
- Modernized GitHub Actions: separate native-Windows build and PWA validation
  jobs, plus a one-shot multi-platform release workflow.
- The legacy VB.NET WinForms application is preserved for Windows users who
  need its Excel COM Interop-powered recovery routines.

EOF

echo
echo "✅ Built bundles in $DIST:"
ls -lh "$DIST"
