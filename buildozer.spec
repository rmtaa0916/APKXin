[app]
title = Form Alchemist
package.name = formalchemist
package.domain = org.formalchemist

source.dir = .
source.include_exts = py,png,jpg,jpeg,kv,atlas,xlsx,pdf,json,csv,java,xml,ttf,otf,txt
version = 1.6.0

requirements = python3==3.10.11,hostpython3==3.10.11,kivy,kivymd,androidssystemfilechooser,opencv,pandas,numpy,openpyxl,pypdf,typing_extensions,reportlab,pillow

orientation = portrait
fullscreen = 1

icon.filename = assets/icon.png
# Disabled on purpose: the native splash below replaces the old Buildozer presplash.
# presplash.filename = assets/presplash.png

android.archs = arm64-v8a
android.api = 33
android.minapi = 24
android.ndk = 25b
android.accept_sdk_license = True
android.allow_backup = True

android.permissions = INTERNET
android.add_src = android_src
# Copies the splash photo into Android drawable resources and includes our native resource XML.
android.add_resources = assets/presplash.png:drawable/presplash_native.png, android_res
android.entrypoint = org.formalchemist.formalchemist.FormAlchemistActivity
android.activity_class_name = org.formalchemist.formalchemist.FormAlchemistActivity
android.apptheme = "@style/FormAlchemistNativeSplashTheme"

[buildozer]
log_level = 2
warn_on_root = 1
