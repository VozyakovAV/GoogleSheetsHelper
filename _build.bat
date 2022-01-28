@echo off
set sln="GoogleSheetsHelper\GoogleSheetsHelper.csproj"
set libMerge="lib\ilrepack.exe"
set pathBase="bin"
set pathBuild="build"
set fileDll="GoogleSheetsHelperFull.dll"

rem path to MSBuild.exe
set msbuild="C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe"
if not exist %msbuild% (
    set msbuild="C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe"
)

echo === Delete folder %pathBase% ===
rmdir /S /Q %pathBase%

echo === Build project ===
%msbuild% %sln% /t:Rebuild /p:Configuration=Debug /p:OutputPath="..\\%pathBase%\\%pathBuild%"

echo === Merge libs ===
set p=%pathBase%\\%pathBuild%
set d1=%p%\GoogleSheetsHelper.dll
set d2=%p%\Newtonsoft.Json.dll
set d3=%p%\Google.Apis.Core.dll %p%\Google.Apis.dll %p%\Google.Apis.Auth.dll %p%\Google.Apis.Auth.PlatformServices.dll %p%\Google.Apis.Sheets.v4.dll
%libMerge% /lib:"%pathBase%\\%pathBuild%" /out:"%pathBase%\\%fileDll%" %d1% %d2% %d3%

echo The end
pause
