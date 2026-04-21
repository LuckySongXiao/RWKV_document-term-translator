@echo off
chcp 65001 >nul
echo ========================================
echo LibTorch CPU 版本下载
echo ========================================
echo.
echo 正在下载 LibTorch CPU 版本...
echo 文件大小约 200MB
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://download.pytorch.org/libtorch/cpu/libtorch-win-shared-with-deps-2.1.0%%2Bcpu.zip' -OutFile 'libtorch_cpu.zip' -UseBasicParsing"

if exist libtorch_cpu.zip (
    echo.
    echo 下载完成! 正在解压...
    powershell -Command "Expand-Archive -Path 'libtorch_cpu.zip' -DestinationPath 'libtorch_cpu_extracted' -Force"
    echo.
    echo 复制 DLL 文件到 dist_native\bin\...
    xcopy /Y libtorch_cpu_extracted\libtorch\lib\*.dll dist_native\bin\
    echo.
    echo 清理临时文件...
    del libtorch_cpu.zip
    rmdir /s /q libtorch_cpu_extracted
    echo.
    echo 完成!
) else (
    echo 下载失败，请手动下载
)
echo.
pause
