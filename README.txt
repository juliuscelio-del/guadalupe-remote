GUADALUPE NHS REMOTE - ANDROID APP
==================================

What the app does
  1. Connect to the computer in ONE of two ways:
       - Tap "Scan QR code" and scan the QR code shown by the computer app, or
       - Type the address shown under the QR code (for example 192.168.1.23)
         and tap "Connect with this address". The port (5001) is added for you.
  2. Type the PIN shown in the computer window.
  3. The 4 options appear (PowerPoint remote, View desktop, Mouse, Keyboard).
     Tick what you want, tap Continue, and use the remote.
  Android back button = Disconnect (go back to the scan screen).

You still run the computer app (PPTRemote-Windows or PPTRemote-Mac) first.
The phone and computer must be on the same Wi-Fi or hotspot.

BUILD THE APK (free, no Android Studio needed) - using GitHub
  1. Make a free account at github.com and click New repository
     (any name, e.g. guadalupe-remote). Keep it Private if you like.
  2. Click "uploading an existing file" and drag in EVERYTHING inside this
     folder (app, settings.gradle, build.gradle, gradle.properties).
     Commit.
  3. Add the build recipe: Add file > Create new file. For the file name type
        .github/workflows/build-apk.yml
     (typing the slashes makes the folders). Paste the contents of the
     build-apk.yml file from this folder. Commit.
  4. Open the "Actions" tab > "Build Android APK" > Run workflow.
  5. Wait about 3-5 minutes until it shows a green check. Open the run,
     scroll to "Artifacts", download GuadalupeNHS-Remote-APK, unzip it.
     Inside is app-debug.apk.

INSTALL ON THE PHONE
  1. Send app-debug.apk to the phone (USB, Google Drive, Messenger, etc.).
  2. Open it. If Android asks, allow "Install unknown apps" for that app.
     Play Protect may warn about an unknown app: choose Install anyway.
  3. Open "Guadalupe NHS Remote" and allow the Camera when asked.

Alternative: open this folder in Android Studio and press Build > Build APK.
