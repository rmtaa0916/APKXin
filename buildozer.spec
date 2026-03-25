[app]
title = Form Alchemist
package.name = formalchemist
package.domain = org.formalchemist
source.dir = .
source.include_exts = py,png,jpg,jpeg,kv,atlas,xlsx,pdf,json,csv,java,ttf,otf,txt,xml
version = 1.6.0
requirements = python3==3.10.11,hostpython3==3.10.11,kivy,kivymd,androidstorage4kivy,opencv,pandas,numpy,openpyxl,pypdf,typing_extensions,reportlab,pillow
orientation = portrait
fullscreen = 1
icon.filename = assets/icon.png
# presplash.filename intentionally disabled; native splash handled by custom activity
# presplash.filename = assets/presplash.png

# Custom native splash activity. This is the minimum manifest-affecting setup.
android.entrypoint = org.formalchemist.formalchemist.FormAlchemistActivity
android.activity_class_name = org.formalchemist.formalchemist.FormAlchemistActivity

# Add Java source only; no custom Android resources/theme to reduce manifest/resource risk.
android.add_src = android_src

# keep AndroidX if already used elsewhere
android.enable_androidx = True
