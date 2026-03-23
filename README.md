# Form Alchemist

> Android-first PDF form inspection, field detection, mapping, and assisted form automation built with Kivy.

[![Platform](https://img.shields.io/badge/platform-Android-3DDC84?logo=android&logoColor=white)](#building-for-android)
[![Framework](https://img.shields.io/badge/framework-Kivy-1E1E1E?logo=python&logoColor=white)](#tech-stack)
[![Language](https://img.shields.io/badge/language-Python%203.10-3776AB?logo=python&logoColor=white)](#tech-stack)
[![Status](https://img.shields.io/badge/status-active%20development-blue)](#project-status)

Form Alchemist is a mobile-focused Android application for working with structured PDF forms. It is designed for workflows where users need to open documents, inspect pages, detect answer regions or boxes, map data to fields, review results manually, and iterate on document-handling behavior directly on-device.

This repository includes the Kivy app source, Android packaging configuration, and platform helpers used to build the project into an APK with Buildozer.

---

## Highlights

- PDF viewing and page navigation
- Detection-assisted field / box workflows
- Manual review and mapping tools
- Touch-oriented Android UI
- Session, revision, and project-state handling
- JSON-backed persistence for app data and learned state
- Android packaging via Buildozer
- Expandable architecture for continued UI and detection improvements

---

## What the app is for

Form Alchemist is intended for document-heavy workflows where structured forms need to be reviewed and interacted with efficiently on Android. Depending on the version or branch, the app may be used for tasks such as:

- opening and previewing multi-page PDFs
- moving page-by-page during review
- detecting possible form fields, boxes, or answer regions
- selecting, reviewing, and tuning detections
- mapping external data to detected targets
- storing app state, revisions, and project-specific progress
- refining mobile UX for dense document workflows

The project is under active development, so specific controls and behavior may evolve over time.

---

## Project structure

```text
.
├── main.py
├── buildozer.spec
├── android_src/
├── assets/
└── ...
```

### `main.py`
The main Kivy application entry point.

This file contains the app UI, PDF interaction flow, detection-related behavior, state handling, Android-specific adaptations, and the broader form-review workflow.

### `buildozer.spec`
Android build configuration for Buildozer / python-for-android.

This file controls app packaging details such as:

- visible app title
- package name and domain
- Python requirements
- Android API / NDK targets
- permissions
- icons and presplash assets
- source inclusion rules

### `android_src/`
Android-side helper code and platform integration classes used by the app when needed.

### `assets/`
Bundled static resources such as:

- app icon
- presplash image
- other app visuals or packaged assets

---

## Tech stack

- **Python 3.10**
- **Kivy**
- **KivyMD**
- **Buildozer**
- **python-for-android**
- **NumPy**
- **Pandas**
- **OpenCV**
- **OpenPyXL**
- **PyPDF / PDF utilities**
- **ReportLab**
- **Pillow**

Your exact dependency set is defined by the current `buildozer.spec` and source files in the repository.

---

## Building for Android

### 1. Prepare your environment
Buildozer is typically used on Linux or WSL.

Install Buildozer and required system dependencies for your environment.

```bash
pip install --upgrade buildozer
```

If this is your first Android build machine, you may also need the usual Buildozer / python-for-android prerequisites such as Java, SDK/NDK tooling, and native build packages.

### 2. Review `buildozer.spec`
Before building, verify the important fields in `buildozer.spec`:

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

### 3. Build a debug APK

```bash
buildozer android debug
```

### 4. Build, deploy, and run on a connected device

```bash
buildozer android debug deploy run
```

### 5. Create a release build

```bash
buildozer android release
```

Release artifacts still need proper signing before distribution.

---

## Renaming the app

To change the **visible app name**, update:

### `buildozer.spec`
```ini
title = Form Alchemist
```

### `main.py`
If your app defines a hardcoded title constant or window title, update that value as well so the branding stays consistent inside the app.

> Changing only the visible app name does **not** necessarily require changing `package.name` or `package.domain`.

Change `package.name` and `package.domain` only if you want a new Android application ID. If your project also contains Android-side Java helpers, keep their package paths aligned with any package-ID change.

---

## Development workflow

A practical iteration loop for this project is:

1. update or patch `main.py`
2. keep `buildozer.spec` synchronized with current dependencies and branding
3. rebuild often when working on:
   - PDF rendering
   - touch interactions
   - overlays and box selection
   - Android file access
   - detection tuning
   - startup behavior
4. test on a real Android device whenever possible

For Android apps with dense interactive UI, real-device testing is especially important for:

- touch hitboxes
- responsiveness
- page rendering performance
- memory stability
- startup transitions
- sidebar / overlay usability

---

## Notes on Android packaging

Android packaging for Python apps can be sensitive to dependency compatibility. When troubleshooting builds, commonly check:

- package compatibility with python-for-android
- Android API / NDK compatibility
- native dependency behavior
- file-access permissions and storage behavior
- rendering performance for PDFs and overlays
- device-specific memory pressure

If the project uses custom Android helper classes under `android_src`, make sure their package declarations remain aligned with your configured package name and domain.

---

## Current direction

The app is being shaped toward a more polished Android experience for document and form interaction. Current development generally centers on areas like:

- cleaner mobile UI hierarchy
- smoother preview and startup behavior
- stronger touch interaction support
- improved selection and mapping ergonomics
- better handling of detection overlays and review flows
- continued Android stability and packaging refinement

---

## Project status

**Active development**

Form Alchemist is not positioned here as a finished consumer release. It is an actively refined Android application and codebase intended for continued iteration, testing, and packaging improvements.

---

## Contributing

If you are maintaining or extending the project:

- keep changes aligned with Android packaging constraints
- validate UI behavior on-device, not only on desktop
- avoid introducing dependency changes without checking Buildozer compatibility
- document any platform-specific fixes in commits or release notes

---

## License

Add the license you want to use for this repository, for example:

- MIT License
- Apache-2.0
- All rights reserved

If you have not chosen one yet, update this section before public release.
