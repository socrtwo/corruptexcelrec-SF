<!--MODERNIZED:v2-->
# S2 Recovery Tools for Microsoft Excel

> Recover and repair corrupt `.xls` and `.xlsx` files. Now cross-platform.

[![Live app](https://img.shields.io/badge/live-app-ff2e93?style=for-the-badge)](https://socrtwo.github.io/corruptexcelrec-SF/)
[![Releases](https://img.shields.io/github/v/release/socrtwo/corruptexcelrec-SF?style=for-the-badge&color=7c3aed)](https://github.com/socrtwo/corruptexcelrec-SF/releases)
[![License](https://img.shields.io/github/license/socrtwo/corruptexcelrec-SF?style=for-the-badge&color=22d3ee)](LICENSE)
[![Last commit](https://img.shields.io/github/last-commit/socrtwo/corruptexcelrec-SF?style=for-the-badge&color=34d399)](https://github.com/socrtwo/corruptexcelrec-SF/commits)

🌐 **Live:** https://socrtwo.github.io/corruptexcelrec-SF/
📦 **Downloads:** [Latest release](https://github.com/socrtwo/corruptexcelrec-SF/releases/latest)
📂 **Source:** [socrtwo/corruptexcelrec-SF](https://github.com/socrtwo/corruptexcelrec-SF)

---

This project ships **two** recovery tools that share a brand and lineage:

1. **PWA recovery tool** (new) — runs in the browser on every modern OS, works
   offline, never uploads files. The single codebase under `web/` powers all
   cross-platform releases.
2. **Native VB.NET WinForms tool** (classic) — Windows-only, uses Excel COM
   Interop and bundled CLI utilities for the deepest recovery passes. Source
   under `Excel Recovery/`.

## ✨ Features

- All Microsoft-recommended Excel recovery methods in one interface
- Six strategies: Auto, Strict Read, Lenient XML Repair, ZIP Recovery,
  Salvage Cells, Convert to CSV
- Saves recovered workbooks as `.xlsx`, `.xls`, `.ods`, or `.csv`
- 100% client-side: files never leave your device
- Installable as a native-feeling app on every major platform
- Offline-capable (service worker caches the entire app shell)
- Drag-and-drop, file-picker, paste, or "Open with…" entry points
- Native Windows build keeps the legacy Excel COM Interop recovery routines
  for the most stubborn corruption cases

## 📦 Install / Download

| Platform | How to install |
|----------|----------------|
| 🪟 **Windows** | [Download `.zip`](https://github.com/socrtwo/corruptexcelrec-SF/releases/latest) · or visit the live URL in Edge/Chrome and click the address-bar install button. The classic native `.exe` is also available as `corruptexcelrec-windows-native.zip`. |
| 🍎 **macOS** | [Download `.zip`](https://github.com/socrtwo/corruptexcelrec-SF/releases/latest) · or open the live URL in Safari → **File → Add to Dock**. |
| 🐧 **Linux** | [Download `.tar.gz`](https://github.com/socrtwo/corruptexcelrec-SF/releases/latest) · or install via Chrome/Chromium/Firefox address bar. |
| 🟢 **ChromeOS** | Open the live URL → ⋮ → **Install Excel Recovery**. Appears in launcher. |
| 🤖 **Android** | Open the live URL in Chrome → ⋮ → **Install app**. Or build a signed APK with [PWABuilder](https://www.pwabuilder.com/). |
| 📱 **iOS / iPadOS** | Open the live URL in Safari → **Share → Add to Home Screen**. |
| 🌐 **Web** | Just visit https://socrtwo.github.io/corruptexcelrec-SF/. |

Every release on the [Releases page](https://github.com/socrtwo/corruptexcelrec-SF/releases)
ships one bundle per platform, plus shared release notes describing what's new.

## 🛠 Building from source

### PWA (cross-platform)

The PWA is plain HTML / JS — no build step. To work on it locally:

```bash
cd web
python3 -m http.server 8080
# open http://localhost:8080
```

To produce all per-platform release bundles locally:

```bash
bash scripts/build-releases.sh 5.0.0
ls dist/
```

### Native Windows build

Requires Windows + Visual Studio 2019+ (Community edition works) and
.NET Framework 4.0+:

1. Open `Excel Recovery.sln` in Visual Studio.
2. Restore NuGet packages if prompted.
3. **Build → Build Solution** (`Ctrl+Shift+B`).
4. Find the compiled `.exe` in `Excel Recovery/bin/Release/`.

CI builds the native binary automatically — see
[`.github/workflows/build.yml`](.github/workflows/build.yml).

## 🚀 Cutting a release

Tag-driven, fully automated:

```bash
git tag v5.0.0
git push origin v5.0.0
```

The [`release.yml`](.github/workflows/release.yml) workflow:

1. Builds all 7 platform bundles via `scripts/build-releases.sh`.
2. Builds the native Windows `.exe` on a Windows runner.
3. Creates the GitHub Release and uploads every artifact.

## 🔒 Privacy

Files you open in the recovery tool are processed **entirely in your
browser** using JSZip and SheetJS. Nothing is uploaded to any server.
The PWA's only network traffic is fetching its own static assets (and even
those are cached offline after first load).

## 📜 SourceForge heritage

This project originated on **SourceForge** before being migrated to GitHub.

🔗 https://sourceforge.net/projects/corruptexcelrec/

The repository here at [`socrtwo/corruptexcelrec-SF`](https://github.com/socrtwo/corruptexcelrec-SF)
is the canonical, actively-maintained home. All future updates, issue
tracking, and releases happen on GitHub.

## 🤝 Contributing

Issues and pull requests are welcome at
https://github.com/socrtwo/corruptexcelrec-SF/issues.

```bash
git checkout -b my-feature
# hack hack hack
git push -u origin my-feature
# open a PR
```

## 📝 License

MIT — see [LICENSE](LICENSE).

---

*Maintained by [@socrtwo](https://github.com/socrtwo)*
