# MIT License
#
# Copyright (c) 2025 iappyx
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
# 
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
# 
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.

if (!(Get-Module -ListAvailable ps2exe)) {
    Write-Host "Error: PS2EXE module is not installed." -ForegroundColor Red
    Write-Host "Install it using: Install-Module PS2EXE" -ForegroundColor Yellow
    exit
}

$routingPath = Join-Path $PSScriptRoot 'EditModeRouting.ps1'
$sourcePath = Join-Path $PSScriptRoot 'src.ps1'
$iconPath = Join-Path $PSScriptRoot 'Instrumenta.ico'
$outputPath = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\bin\windows\Instrumenta Keys.exe'))
$combinedPath = Join-Path $PSScriptRoot '.instrumenta-keys-build.ps1'

if (-not (Test-Path -LiteralPath $routingPath)) {
    throw "Missing Edit Mode routing source: $routingPath"
}

try {
    $combinedSource = @(
        Get-Content -LiteralPath $routingPath -Raw
        Get-Content -LiteralPath $sourcePath -Raw
    ) -join "`r`n"
    Set-Content -LiteralPath $combinedPath -Value $combinedSource -Encoding UTF8

    Invoke-PS2EXE `
        -inputFile $combinedPath `
        -outputFile $outputPath `
        -title "Instrumenta Keys" `
        -iconFile $iconPath
}
finally {
    if (Test-Path -LiteralPath $combinedPath) {
        Remove-Item -LiteralPath $combinedPath -Force
    }
}
