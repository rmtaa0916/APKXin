[app]
title = Form Alchemist
package.name = formalchemist
package.domain = org.formalchemist

source.dir = .
source.include_exts = py,png,jpg,jpeg,kv,atlas,xlsx,pdf,json,csv,java,ttf,otf,txt
version = 1.6.0

requirements = python3==3.10.11,hostpython3==3.10.11,kivy,kivymd,androidssystemfilechooser,opencv,openpyxl,pypdf,typing_extensions,reportlab,pillow

orientation = portrait
fullscreen = 0

icon.filename = assets/icon.png
# Route B: do NOT use Buildozer's presplash, because native startup is handled by MainActivity.
# Keep this disabled to avoid presplash resource collisions.
# presplash.filename = assets/presplash.png

android.archs = arm64-v8a
android.api = 33
android.minapi = 24
android.ndk = 25b
android.accept_sdk_license = True
android.allow_backup = True

android.permissions = INTERNET
android.add_src = android_src
# Native Android resources live here.
# Use android_res/drawable/presplash_native.png
# Do NOT add android_res/drawable/presplash.png or presplash.jpg
android.add_resources = android_res

# Custom native activity for startup/loading screen override.
android.entrypoint = org.formalchemist.formalchemist.MainActivity
android.activity_class_name = org.formalchemist.formalchemist.MainActivity

[buildozer]
log_level = 2
warn_on_root = 1
