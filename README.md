# MediMap Pro Kivy App (Android)

This repository contains a **minimal Kivy application** designed to serve as the
foundation for an Android version of the Google Colab–based form‐filling tool.
It does *not* implement the full detection and PDF editing logic; instead it
provides a basic user interface that mirrors the notebook’s structure and
illustrates how the project can be ported to a mobile environment.

## Structure

* `main.py` – The Kivy application entry point.  It sets up the UI with a
  patient selection spinner, a page slider, controls for mapping CSV columns
  to form boxes, and a log output area.  You can extend this class to include
  your detection functions (`run_detection`, `process_doc`, etc.) and update
  the displayed PDF page accordingly.

* `buildozer.spec` – Configuration file for
  [Buildozer](https://buildozer.readthedocs.io/en/latest/).  Buildozer
  automates the process of turning a Kivy app into an Android APK or AAB.  The
  `requirements` line lists the Python packages your app depends on.  In this
  version we rely only on **pure‑Python dependencies** (`numpy`, `pandas`,
  `Pillow`, `pdfplumber`, `plyer`) so that the app can be built without
  custom recipes.  Libraries that require native code (e.g. `PyMuPDF`,
  `opencv-python`) are omitted because they have no official Android wheels
  and will cause build failures unless you provide your own recipes【894520788986263†L409-L419】【559456175720321†L60-L75】.

## Building the APK

1. **Install Buildozer and its prerequisites.** On a Linux machine (or WSL on
   Windows), run:

   ```bash
   pip install --upgrade buildozer
   buildozer init  # Already done in this project; generates buildozer.spec
   ```

   Buildozer will download and set up the Android SDK, NDK and other tools.

2. **Edit** `buildozer.spec` to suit your needs – e.g. change the package
   name, domain, version or permissions.

3. **Build and deploy**:

   ```bash
   buildozer android debug deploy run
   ```

   This command compiles the app, installs it on a connected device or
   emulator, and runs it automatically.  For release builds, use
   `buildozer android release` and sign the resulting AAB/APK as described in
   the Kivy documentation【401424508867218†L274-L399】.

## Notes on dependencies

The current `buildozer.spec` deliberately avoids including any libraries that
require native code.  This is because python‑for‑android cannot compile
packages like **OpenCV** or **PyMuPDF** without custom recipes, and the
upstream projects do not provide Android wheels【894520788986263†L409-L419】【559456175720321†L60-L75】.  Instead, you
should:

* Use **pdfplumber** or **PyPDF2** for reading PDFs.  These libraries are
  implemented in pure Python and work on Android.  To display pages
  graphically, convert them to images using the Android `PdfRenderer` API via
  [Pyjnius](https://pyjnius.readthedocs.io/en/latest/) or by using a server‑side
  service to render the PDF.

* Replace OpenCV functions with **Pillow** and **NumPy**.  Many simple
  image‑processing tasks (resizing, cropping, thresholding) can be
  accomplished with Pillow.  If your detection pipeline requires more
  advanced features, consider porting the heavy processing to a server or
  writing a custom python‑for‑android recipe.

If you choose to add native dependencies later, update the `requirements`
line accordingly and provide the necessary recipes as outlined in the
python‑for‑android documentation.

## Extending this example

To build a fully functional app:

1. **Load the CSV data** as you do in the notebook (using pandas or
   `csv`), populate `patient_spinner.values` with the list of display names
   (`df["_DISPLAY_NAME"]`), and store the row data for mapping.

2. **Read and render the PDF.**  On Android you can use the
   Android `PdfRenderer` API via [Pyjnius](https://pyjnius.readthedocs.io/en/latest/)
   or use a Python PDF library that works on Android.  Convert each page to a
   `Texture` and assign it to `self.image_widget.texture`.

3. **Port detection logic.**  Copy functions like `run_detection`,
   `find_answer_lines`, `looks_like_checkbox` etc. into this app.  Call them
   when the user changes detection parameters or selects a new page.  You may
   need to adapt OpenCV usage to ensure compatibility on Android.

4. **Implement mapping and saving.**  Maintain a data structure to store
   custom mappings and provide buttons to save/load the configuration as JSON
   using the Kivy file APIs.

Please refer to the Kivy and python‑for‑android documentation for detailed
guidance on packaging and debugging your application【401424508867218†L274-L399】【559456175720321†L60-L75】.