@echo off

rem This script register dependences for run the crud_vb6.exe application 
rem Source: https://github.com/manuel-chinchi/crud-vb6/tree/master/Scripts/dependences.bat

rem The .ocx and .dll files must be in the same directory as the .exe. This script must also 
rem be run from that directory.

echo Register dependences...

rem Check the System is 64 bits or 32 bits
if exist "%windir%\SysWOW64\regsvr32.exe" (
    echo Detect system 64 bits.
    
    rem CR8.5 dependences
    %windir%\SysWOW64\regsvr32.exe /s %~dp0crviewer.dll
    %windir%\SysWOW64\regsvr32.exe /s %~dp0craxdrt.dll
    rem %windir%\SysWOW64\regsvr32.exe /s %~dp0P2smon.dll
    %windir%\SysWOW64\regsvr32.exe /s %~dp0crxf_pdf.dll
    %windir%\SysWOW64\regsvr32.exe /s %~dp0crtslv.dll
    %windir%\SysWOW64\regsvr32.exe /s %~dp0EXPMOD.dll
    %windir%\SysWOW64\regsvr32.exe /s %~dp0u2ddisk.dll
    %windir%\SysWOW64\regsvr32.exe /s %~dp0u2fwordw.dll
    %windir%\SysWOW64\regsvr32.exe /s %~dp0u2fxls.dll
    
    rem VB6.0 dependences
    %windir%\SysWOW64\regsvr32.exe /s %~dp0mscomctl.ocx
    %windir%\SysWOW64\regsvr32.exe /s %~dp0comdlg32.ocx
) else (
    echo Detect system 32 bits.
    
    rem CR8.5 dependences
    %windir%\System32\regsvr32.exe /s %~dp0crviewer.dll
    %windir%\System32\regsvr32.exe /s %~dp0craxdrt.dll
    rem %windir%\System32\regsvr32.exe /s %~dp0P2smon.dll
    %windir%\System32\regsvr32.exe /s %~dp0crxf_pdf.dll
    %windir%\System32\regsvr32.exe /s %~dp0crtslv.dll
    %windir%\System32\regsvr32.exe /s %~dp0EXPMOD.dll
    %windir%\System32\regsvr32.exe /s %~dp0u2ddisk.dll
    %windir%\System32\regsvr32.exe /s %~dp0u2fwordw.dll
    %windir%\System32\regsvr32.exe /s %~dp0u2fxls.dll
    
    rem VB6.0 dependences
    %windir%\System32\regsvr32.exe /s %~dp0mscomctl.ocx
    %windir%\System32\regsvr32.exe /s %~dp0comdlg32.ocx
)
echo Complete registration.
pause
