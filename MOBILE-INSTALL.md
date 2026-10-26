# Android and iPhone offline application

The mobile projects are native C# Android and iOS hosts containing the Chinese CFL interface, Three.js viewer and a shared offline calculation/drawing engine. They do not connect to a Windows computer or start Excel. All three DWGs, the complete original XLS workbook and a licensed Chinese PDF font are embedded. Projects remain XLS workbooks, including their database sheets and VBA streams; macros are preserved but never executed. Mobile forms edit input values and retain formulas. The Windows application remains a separate project.

## Android installer

The local package is `dist/CFL-Mobile/Android/CFL-SubmitTools-Android-arm64.apk`. It targets Android 8 or later on ARM64 phones. Copy it to the phone, open it, and allow installation from the file application when Android asks. Open **凯乐提交工具**, choose **新建项目** or **打开项目**, edit the parameter forms, calculate and generate DWG/PDF files. **导入 Excel** imports an existing XLS project; **导出 Excel / DWG / PDF** saves a copy through the phone's file picker.

The APK is development-signed with the local .NET Android build identity. Retain that private keystore for compatible local updates; do not put it in Git. Use a company release keystore before distributing a production release. Physical-device installation and airplane-mode checks are still required; compilation and desktop execution of the shared offline engine do not substitute for a phone test.

Version **1.2 (build 3)** updates both Android and iPhone completion handling. After successful calculation, the busy overlay closes and **生成 DWG＋PDF** becomes enabled. The app shows progress and provides **取消计算**. Cancellation and calculation errors leave the original project workbook unchanged; cancelled or failed calculations keep generation disabled. Calculation uses a temporary XLS copy and returns its computed state directly instead of reopening the workbook. The offline engine checks cancellation during workbook operations; automatic cancellation is requested after 60 seconds without progress or three minutes total and takes effect at the next workbook-operation checkpoint. It does not use Excel COM, a Windows server or a background Windows process. This update also removes duplicate formula evaluation and caches read-only pump database ranges within each calculation. A 30-pump shared-engine check completed in about 0.5 seconds on the development PC; actual phone times depend on the device. Calculation failures remain visible until the user starts another operation, including the exact missing database/model/baseplate.

Rebuild from Windows with `powershell -ExecutionPolicy Bypass -File tools/build-mobile-android.ps1`. The current project uses the installed .NET 9 / Android API 35 toolchain. Pass `-AndroidSdkDirectory` and `-JavaSdkDirectory` when your SDK/JDK locations differ. It uses the Android SDK directly, without requiring MAUI.

## iPhone installer on your Mac

The provided `dist/CFL-Mobile/iOS/CFL-SubmitTools-iPhone-Mac-Build-Kit.zip` contains the required source, unchanged business templates, logo, font licenses and signing script. It is a build kit, not an installable IPA. The signed device installer must be produced on your Mac using your Apple signing identity and provisioning profile. No Apple account passwords, private keys or provisioning profiles are included in the project or GitHub repository.

1. Extract the ZIP to a writable folder on your Mac, or clone the repository.
2. Install the .NET 9 SDK and its iOS workload. The locally checked SDK is 9.0.318 with iOS workload 26.0.9752; use matching Xcode (26.0) and select it using `xcode-select`. Check Microsoft's iOS release matrix when installing a different workload/Xcode pair. Run `dotnet workload install ios` if the iOS workload is missing.
3. In your Apple developer account/Xcode, register bundle identifier **com.paciguard.sdoctools**, your test iPhone and a development or Ad Hoc provisioning profile for that identifier. Install the matching certificate/private key and profile on your Mac. Keep these outside the repository. An App Store provisioning profile produces a distribution archive for Apple's distribution flow rather than direct installation on arbitrary phones.
4. In Terminal, move into the extracted project folder. List installed signing identities with `security find-identity -v -p codesigning`.
5. Set the two values for your own identity/profile and run:

   ```bash
   export CFL_CODESIGN_KEY='Apple Development: Your Name (TEAMID)'
   export CFL_CODESIGN_PROVISION='Your matching profile name or UUID'
   bash tools/build-mobile-ios.sh
   ```

6. The script creates **dist/CFL-Mobile/iOS/CFL-SubmitTools-iPhone.ipa**. Install it on an iPhone included in the provisioning profile using Xcode's Devices and Simulators window or Apple Configurator. Enable Developer Mode if your development-signed application requires it.
7. On the phone, test New/Open, all four parameter groups, calculation, file import/export, pinch zoom and DWG/PDF generation with airplane mode enabled.

iPhone C# API compilation has been checked on Windows using the installed iOS reference assemblies. Native linking, code signing and running on iOS require the Mac. The project enables the .NET interpreter and disables trimming because the migrated calculation rules and CAD/workbook libraries use dynamic calls and reflection. Do not enable Native AOT or aggressive trimming without additional compatibility work. App Store submission, store icons and production signing/distribution are separate release work.

## Checks and packaging

- `tests/OfflineSmoke`: edits/saves the original workbook without Excel, loads the parameter choices, calculates six equipment rows and generates six editable DWG/vector-PDF pairs. Each saved DWG is reopened and its attributes checked.
- `tests/OfflineCompare`: on Windows with desktop Excel, compares both the managed Windows and mobile results (546 cells each) against an independent Excel COM calculation, verifies saved print settings and checks that Windows calculation starts no Excel process (including dimensions, weights and SEC names).
- `tests/OfflineCalculationSmoke`: checks cancellation, byte-identical original XLS preservation after cancellation/missing-record failures, and successful calculation returning a state that enables drawing generation.
- `tests/MobileUiSmoke`: private loopback-only UI test host; `/phone` shows the real shared interface in a 390-pixel frame. It is a development test utility, not a mobile server dependency.
- `tools/package-mobile.ps1`: rebuilds the Android package and creates the iPhone Mac build kit under `dist/CFL-Mobile`.

Platform references: [.NET for Android](https://learn.microsoft.com/dotnet/android/), [.NET for iOS releases and Xcode requirements](https://github.com/dotnet/macios/releases), [Apple app distribution](https://developer.apple.com/documentation/xcode/distributing-your-app-for-beta-testing-and-releases).
