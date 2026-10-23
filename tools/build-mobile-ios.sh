#!/usr/bin/env bash
set -euo pipefail
project_root="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")/.." && pwd)"
if [[ "$(uname -s)" != Darwin ]]; then
  echo 'Build the iPhone installer on your Mac with Xcode.' >&2; exit 1
fi
: "${CFL_CODESIGN_KEY:?Set CFL_CODESIGN_KEY to the Apple signing identity installed in your Mac keychain.}"
: "${CFL_CODESIGN_PROVISION:?Set CFL_CODESIGN_PROVISION to the matching provisioning profile name or UUID.}"
export DOTNET_CLI_WORKLOAD_UPDATE_NOTIFY_DISABLE=true DOTNET_CLI_TELEMETRY_OPTOUT=1
command -v dotnet >/dev/null
xcode-select -p >/dev/null
destination="$project_root/dist/CFL-Mobile/iOS"
mkdir -p "$destination"
cd "$project_root/ZTools.Mobile.iOS"
sips -s format png -z 1024 1024 "$project_root/ZTools.Standalone/wwwroot/cfl-logo.jpg" \
  --out "$project_root/ZTools.Mobile.iOS/Assets.xcassets/AppIcon.appiconset/AppIcon.png" >/dev/null
dotnet publish "$project_root/ZTools.Mobile.iOS/ZTools.Mobile.iOS.csproj" \
  -c Release -r ios-arm64 --disable-build-servers \
  -p:ArchiveOnBuild=true -p:BuildIpa=true \
  "-p:CodesignKey=$CFL_CODESIGN_KEY" "-p:CodesignProvision=$CFL_CODESIGN_PROVISION" \
  "-p:IpaPackageDir=$destination/" -p:IpaPackageName=CFL-SubmitTools-iPhone.ipa \
  -p:UseSharedCompilation=false /nr:false /m:1
if [[ ! -f "$destination/CFL-SubmitTools-iPhone.ipa" ]]; then
  echo 'Build finished without the expected signed IPA. Inspect the signing/build output.' >&2; exit 1
fi
echo "iPhone installer: $destination/CFL-SubmitTools-iPhone.ipa"
