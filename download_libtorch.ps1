# LibTorch 下载脚本
$ProgressPreference = 'Continue'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$url = "https://download.pytorch.org/libtorch/cu121/libtorch-win-shared-with-deps-2.1.0%2Bcu121.zip"
$output = "libtorch_runtime.zip"

Write-Host "开始下载 LibTorch 运行时版本..."
Write-Host "URL: $url"
Write-Host "输出: $output"
Write-Host "文件大小约 2.3GB"
Write-Host ""

$start = Get-Date
try {
    Invoke-WebRequest -Uri $url -OutFile $output -UseBasicParsing -TimeoutSec 3600
    $end = Get-Date
    $duration = ($end - $start).TotalSeconds
    $size = (Get-Item $output).Length / 1GB
    Write-Host ""
    Write-Host "下载完成!"
    Write-Host "大小: $([math]::Round($size, 2)) GB"
    Write-Host "耗时: $([math]::Round($duration, 0)) 秒"
} catch {
    Write-Host "下载失败: $_"
}
