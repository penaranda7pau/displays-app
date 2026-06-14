[app]
title = Inventario Displays
package.name = inventariodisplays
package.domain = org.cpfr
source.dir = .
source.include_exts = py,png,jpg,kv,atlas
version = 1.0
requirements = python3,kivy==2.3.0,requests
orientation = portrait
osx.python_version = 3
osx.kivy_version = 1.9.1
fullscreen = 0
android.permissions = INTERNET
android.api = 33
android.minapi = 21
android.sdk = 33
android.ndk = 25b
android.archs = arm64-v8a
android.allow_backup = True
[buildozer]
log_level = 2
warn_on_root = 1
