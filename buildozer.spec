[app]
title = MediMap Pro
package.name = medimap
package.domain = org.example
version = 0.1

source.dir = .
source.main = main.py
source.include_exts = py,png,jpg,kv,atlas
orientation = portrait

android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE
requirements = python3,kivy,Pillow,plyer,pdfplumber
android.archs = arm64-v8a

[buildozer]
build_dir = build
log_level = 2
warn_on_root = 1
android.ndk_api = 21
android.api = 31
copy_to_sdcard = False
