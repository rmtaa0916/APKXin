[app]
title = MediMapPro
package.name = medimappro
package.domain = org.medimap

source.dir = .
source.include_exts = py,png,jpg,jpeg,kv,atlas,xlsx,pdf,json,csv,java
version = 1.6.0

requirements = python3==3.10.11,hostpython3==3.10.11,kivy,kivymd,androidssystemfilechooser,opencv,pandas,numpy,openpyxl,pypdf,typing_extensions,reportlab,pillow

orientation = portrait
fullscreen = 0

android.archs = arm64-v8a
android.api = 33
android.minapi = 24
android.ndk = 25b
android.accept_sdk_license = True
android.allow_backup = True

android.permissions = INTERNET
android.add_src = android_src

[buildozer]
log_level = 2
warn_on_root = 1
