#!/usr/bin/env bash
# Build + đóng gói bộ cài V-FORMAT (chạy được trên macOS/Linux/Windows có .NET SDK 6+)
set -euo pipefail
cd "$(dirname "$0")"
dotnet test tests/VFormat.Core.Tests -v q
dotnet build src/VFormat.AddIn -c Release -v q
rm -rf dist && mkdir -p dist/VFormat
cp src/VFormat.AddIn/bin/Release/net48/VFormat.AddIn.dll dist/VFormat/
cp installer/install.ps1 installer/uninstall.ps1 installer/install.cmd installer/uninstall.cmd dist/VFormat/
cp README.md dist/VFormat/
(cd dist && zip -qr VFormat.zip VFormat)
echo "Đã tạo dist/VFormat.zip (add-in Word)"

# App độc lập: một tệp VFormat.exe (win-x64, kèm .NET, không cần cài gì)
dotnet publish src/VFormat.App -c Release -o dist/app -p:DebugType=none -v q
(cd dist/app && rm -f VFormat-App.zip && zip -q ../VFormat-App.zip VFormat.exe)
echo "Đã tạo dist/VFormat-App.zip (app kéo thả)"
