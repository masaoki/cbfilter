@echo off
setlocal

set "MSVC_ROOT=C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Tools\MSVC\14.44.35207"
set "SDK_VER=10.0.26100.0"
set "SDK_ROOT=C:\Program Files (x86)\Windows Kits\10"

set "INCLUDE=%MSVC_ROOT%\include;%SDK_ROOT%\Include\%SDK_VER%\ucrt;%SDK_ROOT%\Include\%SDK_VER%\um;%SDK_ROOT%\Include\%SDK_VER%\shared;%SDK_ROOT%\Include\%SDK_VER%\winrt;%SDK_ROOT%\Include\%SDK_VER%\cppwinrt"
set "LIB=%MSVC_ROOT%\lib\x64;%SDK_ROOT%\Lib\%SDK_VER%\ucrt\x64;%SDK_ROOT%\Lib\%SDK_VER%\um\x64"
set "PATH=%MSVC_ROOT%\bin\Hostx64\x64;%SDK_ROOT%\bin\%SDK_VER%\x64;%PATH%"

set "DEBUG_FLAGS=/DDEBUG=0 /O1"
if /I "%~1"=="debug" (
    echo Enabling DEBUG option
    set "DEBUG_FLAGS=/DDEBUG=1 /DEBUG"
)

rem Build resource
rc /nologo /fo cbfilter.res cbfilter.rc

rem Build with cl (C++20)
cl /nologo /EHsc /std:c++20 /utf-8 /I. /DUNICODE /D_UNICODE /DWIN32_LEAN_AND_MEAN /W4 %DEBUG_FLAGS% /Fe:cbfilter.exe ^
    src\main.cpp src\clipboard_processor.cpp cbfilter.res ^
    user32.lib gdi32.lib comctl32.lib shell32.lib winhttp.lib windowsapp.lib gdiplus.lib crypt32.lib ole32.lib
endlocal
