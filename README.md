# Form Alchemist

> Android-first PDF form inspection, detection, mapping, and assisted form automation.

![Platform](https://img.shields.io/badge/Platform-Android-3DDC84?logo=android&logoColor=white)
![Framework](https://img.shields.io/badge/Framework-Kivy-5C2D91?logo=python&logoColor=white)
![Language](https://img.shields.io/badge/Language-Python%203.10-3776AB?logo=python&logoColor=white)
![Status](https://img.shields.io/badge/Status-Active%20Development-0A84FF)
![Updated](https://img.shields.io/badge/Updated-2026--04--05-6f42c1)

Form Alchemist is a Kivy-based Android application for working with structured PDF documents on mobile. It is built for workflows where users need to open a PDF, inspect pages, detect likely fields or answer regions, review results manually, map data to selected targets, and export filled output.

The project combines Android packaging, PDF preview, detection-assisted review, mapping tools, and persistent project state into a single mobile-first app that can continue to evolve with real-world form workflows.

---

## Overview

Form Alchemist is designed for dense, form-heavy document tasks where a desktop-style workflow needs to be brought onto Android without losing precision or control.

Current development focuses on:

- PDF viewing and page-by-page navigation
- detection-assisted field and box discovery
- manual review and correction of detections
- data-to-field mapping workflows
- filled PDF export
- project/session persistence and revision tracking
- touch-friendly Android UI and packaging stability

---

## Key Features

- **Android-first experience** built with Kivy
- **PDF preview and navigation** for multi-page documents
- **Field / line / checkbox detection** for structured form workflows
- **Manual selection and review tools** for refining results
- **Mapping system** for assigning external data to selected targets
- **Filled PDF export** using overlay-based generation
- **Session persistence** for projects and in-progress work
- **Revision tracking** for saved detection states and edits
- **Learning-oriented storage** for reusable profiles and detection improvements
- **Buildozer packaging** for APK generation

---

## What It Can Be Used For

Form Alchemist is suited for workflows such as:

- reviewing structured PDF forms on Android
- detecting likely answer regions, boxes, or fill targets
- validating or refining detected form elements manually
- mapping records or spreadsheet-backed data into PDF fields
- testing mobile document workflows before wider deployment
- iterating on Android UX for document-heavy applications

---

## Tech Stack

- **Python 3.10**
- **Kivy**
- **KivyMD**
- **Buildozer** / **python-for-android**
- **OpenCV**
- **NumPy**
- **Pandas**
- **OpenPyXL**
- **PyPDF**
- **ReportLab**
- **Pillow**

---

## Repository Structure

```text
.
├── main.py
├── buildozer.spec
├── android_src/
├── android_res/
├── assets/
└── README.md
```

### `main.py`
Main application entry point.

Contains the UI flow, PDF preview logic, Android behavior, detection workflow, mapping tools, state persistence, and app interaction logic.

### `buildozer.spec`
Android build configuration.

Defines:

- visible app title
- package name and domain
- requirements
- Android SDK / NDK targets
- icons and presplash assets
- permissions and packaging options

### `android_src/`
Android-specific helper classes and platform-side integration code.

### `assets/`
Bundled app resources such as icons, presplash assets, and UI visuals.

---

## Current Capabilities

Based on the current codebase, the app includes support for:

- opening and validating PDF files
- Android-safe PDF rendering paths
- page rasterization for preview and detection
- line, field, and checkbox-oriented detection flows
- custom mapping assignment to detected regions
- session save/load behavior for project continuity
- revision snapshots for iterative edits
- learning/profile storage for future detection improvements
- export of filled PDF output

---


## Getting Started (Repository)

```bash
git clone <your-fork-or-repo-url>
cd APKXin
python -m venv .venv
source .venv/bin/activate
pip install -U pip buildozer
```

> Use Linux or WSL for Android builds. Buildozer is not officially supported natively on Windows.

## Quick Start

### Requirements

Buildozer projects are typically built on:

- **Linux**, or
- **WSL** on Windows

You will need a normal Android build environment with Java and the packages required by Buildozer and python-for-android.

### Install Buildozer

```bash
pip install --upgrade buildozer
```

### Build a Debug APK

```bash
buildozer android debug
```

### Build, Deploy, and Run

```bash
buildozer android debug deploy run
```

### Create a Release Build

```bash
buildozer android release
```

> Release packages still need proper signing before public distribution.

---

## Build Configuration Notes

Before building, review the important fields in `buildozer.spec`:

- `title`
- `package.name`
- `package.domain`
- `version`
- `requirements`
- `android.api`
- `android.minapi`
- `android.ndk`
- `android.permissions`
- `icon.filename`
- `presplash.filename`

---

## Renaming the App

To change the visible app name, update:

### In `buildozer.spec`

```ini
title = Form Alchemist
```

### In `main.py`
Update any hardcoded title constant or window-title assignment so the in-app branding matches the APK branding.

> Change `package.name` and `package.domain` only if you want a new Android application ID.

If your project uses Android helper classes under `android_src`, keep their package declarations aligned with any package-ID changes.

---

## Android Development Notes

This repository is actively tuned for Android behavior, so real-device testing matters.

Recommended validation areas:

- PDF loading reliability
- page rendering performance
- touch hitboxes and controls
- overlay visibility and selection accuracy
- memory stability on mobile devices
- startup and presplash transitions
- Android file access behavior

---

## Recommended Development Workflow

1. Update `main.py`
2. Keep `buildozer.spec` aligned with current dependencies and branding
3. Rebuild often when working on rendering, selection, mapping, file access, or UI behavior
4. Test on a physical Android device whenever possible
5. Use release tags and GitHub Releases for APK distribution

---

## Release Distribution on GitHub

For visitors and testers, the recommended distribution method is:

1. build the APK
2. create a GitHub Release
3. attach the APK as a release asset
4. publish release notes for each version

That gives your repository a clean public download page for each APK version.

---

## Project Status

**Active development**

Form Alchemist is an evolving Android application focused on practical PDF form interaction, mobile review workflows, and detection-assisted document handling. Features and UI behavior may continue to change as the project is refined.

---

## Roadmap

Potential future directions include:

- stronger multi-select and batch actions
- improved mobile gesture support
- more polished contextual HUD behavior
- more advanced detection tuning and learning reuse
- richer export and project management workflows
- continued Android stability improvements

---

## Contributing

If you are extending or maintaining the app:

- keep Android packaging constraints in mind
- validate behavior on-device, not just on desktop
- keep build settings synchronized with dependency changes
- avoid package-ID mismatches between `buildozer.spec` and Android helper code

---

## License

Add the license you want to use for the repository here.

Examples:

- `MIT`
- `Apache-2.0`
- `All rights reserved`

---

## Acknowledgments

Built with the Python and Android open-source ecosystem, including Kivy, Buildozer, python-for-android, and the PDF/image-processing libraries that make mobile form tooling possible.
