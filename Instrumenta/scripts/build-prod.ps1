[CmdletBinding()]
param(
    [switch]$ForceClosePowerPoint,
    [switch]$NoLaunchSmoke,
    [switch]$LaunchSmoke,
    [switch]$SyncVba,
    [switch]$UseExistingArtifact
)

& (Join-Path $PSScriptRoot 'instrumenta-build.ps1') -Mode prod -ForceClosePowerPoint:$ForceClosePowerPoint -NoLaunchSmoke:$NoLaunchSmoke -LaunchSmoke:$LaunchSmoke -SyncVba:$SyncVba -UseExistingArtifact:$UseExistingArtifact
