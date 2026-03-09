[app]
title = MediMap Pro
package.name = medimap
package.domain = org.example
version = 0.1

source.dir = .
source.main = main.py
source.include_exts = py,png,jpg,kv,atlas
source.exclude_dirs = build,.buildozer,.git,.github,__pycache__,venv
source.exclude_patterns = *.pyc,*.pyo

orientation = portrait

android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE
requirements = python3,kivy,Pillow,plyer,pdfplumber
android.archs = arm64-v8a

[buildozer]
log_level = 2
warn_on_root = 1
android.ndk_api = 21
android.api = 34
copy_to_sdcard = False
