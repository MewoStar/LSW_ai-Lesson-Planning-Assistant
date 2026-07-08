# ============================================
# AI 备课助手 - 静默启动入口
# 调用 launcher.vbs 实现完全无窗口启动
# ============================================
$ErrorActionPreference = "Continue"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $scriptDir

Start-Process -FilePath "wscript.exe" -ArgumentList "//nologo", "`"$scriptDir\launcher.vbs`"" -WindowStyle Hidden
