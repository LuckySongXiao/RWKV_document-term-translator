@echo off
chcp 65001 >nul
echo ========================================
echo LibTorch 运行时版本下载
echo ========================================
echo.
echo 正在下载 LibTorch CUDA 12.1 运行时版本...
echo 文件大小约 2.3GB
echo.
echo 开始时间: %time%
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; $ProgressPreference = 'SilentlyContinue'; Invoke-WebRequest -Uri 'https://download.pytorch.org/libtorch/cu121/libtorch-win-shared-with-deps-2.1.0%%2Bcu121.zip' -OutFile 'libtorch_runtime.zip' -UseBasicParsing -TimeoutSec 3600"

echo.
echo 结束时间: %time%
echo.

if exist libtorch_runtime.zip (
    echo 下载完成!
    for %%A in (libtorch_runtime.zip) do echo 文件大小: %%~zA bytes
) else (
    echo 下载失败!
)
echo.
pause
