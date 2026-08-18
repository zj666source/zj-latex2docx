@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ============================================
echo   zj公式转化工具 - 一键构建 exe
echo ============================================

echo [1/4] 安装依赖...
python -m pip install -r requirements.txt || goto :err

echo [2/4] 生成程序图标 gs.ico（make_icon.py）...
python make_icon.py || goto :err
if not exist "gs.ico" (
    echo [错误] 图标生成失败，构建中止。
    goto :err
)

echo [3/4] 获取 MML2OMML.XSL（微软 Office 组件）...
python fetch_xsl.py || goto :err
if not exist "MML2OMML.XSL" (
    echo [错误] 缺少 MML2OMML.XSL，构建中止。
    goto :err
)

rem 定位 latex2mathml 数据文件目录（unimathsymbols.txt 必须随 exe 打包）
for /f "delims=" %%i in ('python -c "import latex2mathml,os;print(os.path.dirname(latex2mathml.__file__).replace(os.sep,'/'))"') do set L2M=%%i
if "%L2M%"=="" (
    echo [错误] 未找到 latex2mathml 包，构建中止。
    goto :err
)
echo        latex2mathml 目录: %L2M%

echo [4/4] 构建两个 exe（主版 windowed + 诊断版 console）...
python -m PyInstaller --noconfirm --onefile --windowed --name "公式一键转换" ^
    --icon gs.ico ^
    --add-data "MML2OMML.XSL;." ^
    --add-data "%L2M%/unimathsymbols.txt;latex2mathml/" ^
    --hidden-import latex2mathml.converter --collect-submodules latex2mathml ^
    --hidden-import lxml.etree --collect-submodules lxml ^
    latex2docx.py || goto :err

python -m PyInstaller --noconfirm --onefile --console --name "公式一键转换_诊断版" ^
    --icon gs.ico ^
    --add-data "MML2OMML.XSL;." ^
    --add-data "%L2M%/unimathsymbols.txt;latex2mathml/" ^
    --hidden-import latex2mathml.converter --collect-submodules latex2mathml ^
    --hidden-import lxml.etree --collect-submodules lxml ^
    latex2docx.py || goto :err

echo.
echo [完成] 构建成功：
echo   dist\公式一键转换.exe          （主版，双击使用）
echo   dist\公式一键转换_诊断版.exe   （诊断版，报错可见）
pause
exit /b 0

:err
echo.
echo [失败] 构建出错，请根据上方日志排查。
pause
exit /b 1
