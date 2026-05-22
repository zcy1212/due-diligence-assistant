# 尽调小助手 - 打包脚本
# 用于将skill打包成可分享的压缩包

param(
    [string]$Version = "1.0.0",
    [string]$OutputDir = ".."
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PackageName = "尽调小助手-v$Version"
$TempDir = Join-Path $ScriptDir "package_temp"
$ZipPath = Join-Path $OutputDir "$PackageName.zip"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  尽调小助手 - 打包工具" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 清理旧的临时目录
if (Test-Path $TempDir) {
    Write-Host "清理旧的临时目录..." -ForegroundColor Yellow
    Remove-Item $TempDir -Recurse -Force
}

# 创建临时目录
Write-Host "创建临时目录..." -ForegroundColor Green
New-Item -ItemType Directory -Path $TempDir | Out-Null
New-Item -ItemType Directory -Path "$TempDir\.claude\skills" | Out-Null
New-Item -ItemType Directory -Path "$TempDir\tools" | Out-Null
New-Item -ItemType Directory -Path "$TempDir\templates" | Out-Null
New-Item -ItemType Directory -Path "$TempDir\input" | Out-Null
New-Item -ItemType Directory -Path "$TempDir\output" | Out-Null
New-Item -ItemType Directory -Path "$TempDir\temp" | Out-Null

# 复制文件
Write-Host "复制文件..." -ForegroundColor Green

$filesToCopy = @(
    @{Source = ".claude\skills\尽调小助手.md"; Dest = ".claude\skills\"},
    @{Source = "tools\file_reader.py"; Dest = "tools\"},
    @{Source = "tools\excel_writer.py"; Dest = "tools\"},
    @{Source = "tools\report_writer.py"; Dest = "tools\"},
    @{Source = "templates\报告模板.txt"; Dest = "templates\"},
    @{Source = "templates\项目表示例.xlsx"; Dest = "templates\"},
    @{Source = "CLAUDE.md"; Dest = ""},
    @{Source = "README.md"; Dest = ""},
    @{Source = "INSTALL.md"; Dest = ""},
    @{Source = "requirements.txt"; Dest = ""},
    @{Source = "skill.json"; Dest = ""}
)

foreach ($file in $filesToCopy) {
    $sourcePath = Join-Path $ScriptDir $file.Source
    $destPath = Join-Path $TempDir $file.Dest
    
    if (Test-Path $sourcePath) {
        Copy-Item $sourcePath -Destination $destPath -Force
        Write-Host "  已复制: $($file.Source)" -ForegroundColor Gray
    } else {
        Write-Host "  警告: 文件不存在 - $($file.Source)" -ForegroundColor Yellow
    }
}

# 创建占位文件，确保空目录被打包
"# 在此目录放入待处理的尽调文件" | Out-File -FilePath "$TempDir\input\README.txt" -Encoding utf8
"# 生成的文件会保存在此目录" | Out-File -FilePath "$TempDir\output\README.txt" -Encoding utf8
"# 临时文件目录" | Out-File -FilePath "$TempDir\temp\README.txt" -Encoding utf8

Write-Host ""
Write-Host "创建压缩包..." -ForegroundColor Green

# 删除旧的压缩包
if (Test-Path $ZipPath) {
    Remove-Item $ZipPath -Force
}

# 创建压缩包
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($TempDir, $ZipPath)

# 清理临时目录
Write-Host "清理临时目录..." -ForegroundColor Yellow
Remove-Item $TempDir -Recurse -Force

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  打包完成!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "压缩包位置: $ZipPath" -ForegroundColor White
Write-Host "版本: v$Version" -ForegroundColor White
Write-Host ""
Write-Host "分享说明:" -ForegroundColor Yellow
Write-Host "1. 将 $PackageName.zip 发送给其他用户"
Write-Host "2. 用户解压后在Trae中打开目录"
Write-Host "3. 运行 'pip install -r requirements.txt' 安装依赖"
Write-Host "4. 输入 '/尽调小助手' 开始使用"
Write-Host ""
