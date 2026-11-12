@echo off
rem Builds QuillCnv.cnv (64-bit, for 64-bit Word), CnvTest.exe and a single-file
rem QuillToDocx.exe into WordConverter\bin, then the MSI into Installer\bin\x64\Release.
rem Install-QuillConverter.ps1 installs from WordConverter\bin without the MSI.
setlocal
set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
for /f "usebackq delims=" %%i in (`"%VSWHERE%" -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath`) do set "VSDIR=%%i"
if not defined VSDIR (
    echo Visual Studio with the C++ x64 tools was not found.
    exit /b 1
)
call "%VSDIR%\VC\Auxiliary\Build\vcvars64.bat" >nul 2>&1 || exit /b 1

set "SRC=%~dp0"
set "ROOT=%~dp0..\"
set "OUT=%~dp0bin"
if not exist "%OUT%" mkdir "%OUT%"

rem The version lives in Directory.Build.props; stamp it into the converter DLL too.
for /f "usebackq delims=" %%v in (`powershell -NoProfile -Command "([xml](Get-Content '%ROOT%Directory.Build.props')).Project.PropertyGroup.Version"`) do set "VERSION=%%v"
if not defined VERSION (
    echo Could not read Version from Directory.Build.props.
    exit /b 1
)
set "VERCOMMA=%VERSION:.=,%,0"
> "%OUT%\version.h" echo #define QUILL_VERSION_COMMA %VERCOMMA%
>> "%OUT%\version.h" echo #define QUILL_VERSION_STR "%VERSION%"

rc /nologo /i "%OUT%" /fo "%OUT%\QuillCnv.res" "%SRC%QuillCnv.rc" || exit /b 1
cl /nologo /O2 /W4 /EHsc /MT /LD "%SRC%QuillCnv.cpp" "%OUT%\QuillCnv.res" /Fo"%OUT%\\" /Fe"%OUT%\QuillCnv.cnv" /link /DEF:"%SRC%QuillCnv.def" || exit /b 1
cl /nologo /O2 /W4 /EHsc /MT "%SRC%CnvTest.cpp" /Fo"%OUT%\\" /Fe"%OUT%\CnvTest.exe" || exit /b 1
dotnet publish "%ROOT%QuillToDocx\QuillToDocx.fsproj" -c Release -o "%OUT%" --nologo -v quiet || exit /b 1
dotnet build "%ROOT%Installer\Installer.wixproj" -c Release --nologo -v quiet || exit /b 1
for %%m in ("%ROOT%Installer\bin\x64\Release\QuillToDocx-%VERSION%-x64.msi") do echo Built %%~fm
