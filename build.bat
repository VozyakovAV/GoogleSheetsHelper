@echo off
set libMerge="lib\ilrepack.exe"
set pathBase="bin"
set pathBuild="build"
set fileDll="GoogleSheetsHelperFull.dll"

echo === Delete folder %pathBase% ===
rmdir /S /Q %pathBase%

rem path to MSBuild.exe
dotnet build ExportToGoogle\ExportToGoogle.csproj -o "%pathBase%\\%pathBuild%"

echo === Build project ===
rem %msbuild% %sln% /t:Rebuild /p:Configuration=Debug /p:OutputPath="..\\%pathBase%\\%pathBuild%"
rem %msbuild% %sln% /t:Pack /p:Configuration=Debug /p:OutputPath="..\\%pathBase%\\%pathBuild%"

echo === Merge libs ===
set p=%pathBase%\\%pathBuild%
set d1=%p%\GoogleSheetsHelper.dll
set d2=%p%\Newtonsoft.Json.dll
set d3=%p%\Google.Apis.Core.dll %p%\Google.Apis.dll %p%\Google.Apis.Auth.dll %p%\Google.Apis.Auth.PlatformServices.dll %p%\Google.Apis.Sheets.v4.dll
rem %libMerge% /lib:"%pathBase%\\%pathBuild%" /out:"%pathBase%\\%fileDll%" %d1% %d2% %d3%

echo The end
pause
