# 电弧炉炼钢过程记录 - Windows PowerShell 启动脚本
# 右键点击 → 使用 PowerShell 运行

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Set-Location $PSScriptRoot

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  电弧炉炼钢过程记录 - 多人协作版" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 检查 Node.js
try {
    $nodeVer = node -v 2>&1
    Write-Host "[OK] Node.js 版本: $nodeVer" -ForegroundColor Green
} catch {
    Write-Host "[错误] 未检测到 Node.js！" -ForegroundColor Red
    Write-Host ""
    Write-Host "请先安装 Node.js："
    Write-Host "  1. 打开 https://nodejs.org/"
    Write-Host "  2. 下载左侧 LTS 版本（绿色大按钮）"
    Write-Host "  3. 安装时一路点下一步"
    Write-Host "  4. 装完后右键此脚本 → 使用 PowerShell 运行"
    Write-Host ""
    Read-Host "按回车退出"
    exit 1
}

# 首次安装依赖
if (-not (Test-Path "node_modules")) {
    Write-Host ""
    Write-Host "[首次运行] 正在安装依赖，请稍等..."
    npm install --registry=https://registry.npmmirror.com
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[错误] 依赖安装失败，请检查网络" -ForegroundColor Red
        Read-Host "按回车退出"
        exit 1
    }
    Write-Host "[OK] 依赖安装完成" -ForegroundColor Green
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  启动中..." -ForegroundColor Yellow
Write-Host "  浏览器打开: http://localhost:3000" -ForegroundColor White
Write-Host "  手机同 Wi-Fi: http://你的电脑IP:3000" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

node server.js
Read-Host "按回车退出"
