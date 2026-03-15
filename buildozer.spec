[app]
title = MediMapPro
package.name = medimappro
package.domain = org.medimap
source.dir = .
source.include_exts = py,png,jpg,jpeg,kv,atlas,xlsx,pdf
version = 1.5.0

requirements = python3,kivy,pandas,numpy,openpyxl,pypdf,reportlab,pillow

orientation = portrait
fullscreen = 0

android.archs = arm64-v8a
android.allow_backup = True

# Safer defaults for newer Android packaging
android.api = 33
android.minapi = 21

# Storage permissions
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE,MANAGE_EXTERNAL_STORAGE

# Optional but often helpful
log_level = 2
warn_on_root = 1

[buildozer]
log_level = 2
warn_on_root = 1
