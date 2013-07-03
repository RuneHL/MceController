mode 120,600
SET BUILD_TYPE=Release
if "%1" == "Debug" set BUILD_TYPE=Debug

REM Determine whether we are on an 32 or 64 bit machine
if "%PROCESSOR_ARCHITECTURE%"=="x86" if "%PROCESSOR_ARCHITEW6432%"=="" goto x86

set ProgramFilesPath=%ProgramFiles(x86)%

goto startInstall

:x86

set ProgramFilesPath=%ProgramFiles%

:startInstall

SET WIX_BUILD_LOCATION=C:\Program Files (x86)\WiX Toolset v3.7\bin
SET SRC_PATH=.
SET INTERMEDIATE_PATH=%SRC_PATH%\obj\%BUILD_TYPE%
REM SET OUTPUTNAME=%SRC_PATH%\bin\%BUILD_TYPE%\Media Center Controller.msi
SET OUTPUTNAME=%SRC_PATH%\Setup\%BUILD_TYPE%\Media Center Controller.msi

REM Cleanup leftover intermediate files

del /f /q "%INTERMEDIATE_PATH%\*.wixobj"
del /f /q "%OUTPUTNAME%"

REM Build the MSI for the setup package

pushd "%SRC_PATH%\Setup"

set wixext=-ext "%WIX_BUILD_LOCATION%\WixUIExtension.dll" -ext "%WIX_BUILD_LOCATION%\WixUtilExtension.dll" -ext "%WIX_BUILD_LOCATION%\WixFirewallExtension.dll"

"%WIX_BUILD_LOCATION%\candle.exe" "%SRC_PATH%\Setup.wxs" -dBuildType=%BUILD_TYPE% %wixext% -out "%INTERMEDIATE_PATH%\MCNC.wixobj"
"%WIX_BUILD_LOCATION%\light.exe" "%INTERMEDIATE_PATH%\MCNC.wixobj" %wixext% -cultures:en-US  -loc "Setup_en-us.wxl" -out "%OUTPUTNAME%"

popd
pause