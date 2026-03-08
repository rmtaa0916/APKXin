# Buildozer specification file for MediMap Pro
#
# This spec file configures how Buildozer packages your Kivy application into
# an Android APK or AAB. See the Buildozer documentation for more details.

[app]
# (str) Title of your application
title = MediMap Pro

# (str) Package name
package.name = medimap

# (str) Package domain (must be a valid domain name)
package.domain = org.example

# (str) Version of your application
version = 0.1

# (str) Application source code directory
source.dir = .

# (str) The main .py file to use as the main entry point for your app
source.main = main.py

# (list) List of inclusions using pattern matching
# For example, *.png, assets/*
source.include_exts = py,png,jpg,kv,atlas

# (str) Presplash image used when loading the app
# presplash.filename = %(source.dir)s/data/presplash.png

# (str) Icon of the application
# icon.filename = %(source.dir)s/data/icon.png

# (str) Supported orientation (one of: portrait, landscape, all)
orientation = portrait

# (list) Permissions
android.permissions = WRITE_EXTERNAL_STORAGE, READ_EXTERNAL_STORAGE

# (list) Application requirements
# Specify all the Python packages and any extra dependencies your app uses.
# PyMuPDF currently has no official Android wheel; it will likely fail to
# compile. You may need to provide a custom recipe or use an alternative
# library such as pdfplumber.
# Use only pure‑Python dependencies when possible.  OpenCV and PyMuPDF
# depend on native code and do not have official Android builds; including
# them here will cause the build to fail unless you provide custom
# python‑for‑android recipes.  Instead, rely on pure‑Python packages such
# as pdfplumber and numpy and implement any image processing manually.
requirements = python3,kivy,numpy,pandas,Pillow,plyer,pdfplumber

# (str) Additional arguments to pass when installing requirements
#requirements_install_args =

# (str) Entry point for Kivy (if you use kivy == 1.11 or lower)
#application = main.py

[buildozer]
# (str) Directory where buildozer will execute its tasks
build_dir = build

# (str) Log level (0 = error only, 1 = normal, 2 = verbose, 3 = debug)
log_level = 1

# (bool) Warn if using root to run buildozer commands
warn_on_root = 1

# (str) Android NDK API level to use; Buildozer will automatically
# install the NDK.
android.ndk_api = 27

# (str) Android SDK API level to use
android.api = 27

# (str) Java compiler to use; auto by default
# android.java_toolchain = auto

# (bool) Copy your project to the /sdcard/ to test with Kivy Launcher
# (n.b. does not compile your app)
copy_to_sdcard = False

# (str) Path to the keystore to sign your app for release. Leave empty
# to use debug signing (not for Play Store).
# storefile = /home/user/.keystore
# storepass = password
# keyalias = mykey
# keypass = password