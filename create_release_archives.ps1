# SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
#
# SPDX-License-Identifier: MIT

$runtimes = @(
  "win-x64"
  "win-arm64"
  "osx-arm64"
  "linux-x64"
  "linux-arm64"
)

New-Item -ItemType Directory -Path releases -Force

foreach ($runtime in $runtimes) {
    trap { break }

    dotnet publish Plotance/Plotance.csproj -r $runtime
    Compress-Archive -Path "Plotance/bin/Release/*/$runtime/publish/*" -DestinationPath "releases/plotance-$runtime.zip" -Force
}
