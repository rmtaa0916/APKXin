[app]
title = MediMapPro
package.name = medimappro
package.domain = org.medimap

source.dir = .
source.include_exts = py,png,jpg,jpeg,kv,atlas,xlsx,pdf,json,csv
source.exclude_dirs = .git,__pycache__,bin,.buildozer,venv,.venv
version = 1.6.0

requirements = python3,kivy,kivymd,opencv,pandas,numpy,openpyxl,pypdf,typing_extensions,reportlab,pillow

orientation = portrait
fullscreen = 0

android.archs = arm64-v8a
android.api = 36
android.minapi = 24
android.ndk = 29
android.enable_androidx = True
android.accept_sdk_license = True
android.allow_backup = True

android.permissions = INTERNET

p4a.branch = develop

[buildozer]
log_level = 2
warn_on_root = 1
