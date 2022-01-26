@echo off
set sln="GoogleSheetsHelper\GoogleSheetsHelper.csproj"
set ilmerge="Libs\ilrepack.exe"
set pathBase="bin"
set pathBuild="build"
set fileDll="GoogleSheetsHelperFull.dll"

rem путь до MSBuild.exe
set msbuild="C:\Program Files (x86)\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe"
if not exist %msbuild% (
    set msbuild="C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe"
)

rem удаляем старые сборки
rmdir /S /Q %pathBase%

rem собираем проект в папку bin
%msbuild% %sln% /t:Rebuild /p:Configuration=Debug /p:OutputPath="..\\%pathBase%\\%pathBuild%"

rem мержим библиотеки в одну
set p=%pathBase%\\%pathBuild%
set d1=%p%\GoogleSheetsHelper.dll %p%\Newtonsoft.Json.dll
set d2=%p%\Google.Apis.Core.dll %p%\Google.Apis.dll %p%\Google.Apis.Auth.dll %p%\Google.Apis.Auth.PlatformServices.dll %p%\Google.Apis.Sheets.v4.dll
%ilmerge% /out:"%pathBase%\\%fileDll%" %d1% %d2%

rem xcopy "%pathBase%\%pathBuild%\TSLabCommon.dll" "%pathBase%" /Y
rem xcopy "%pathBase%\%pathBuild%\TSLabExtra.dll" "%pathBase%" /Y

rem echo F|xcopy /S /I /Q /Y /F "%pathBase%\%pathBuild%\TSLabCommon.dll" "%pathBase%\TSLabCommon.dll"
rem echo F|xcopy /S /I /Q /Y /F "%pathBase%\%pathBuild%\TSLabExtra.dll" "%pathBase%\TSLabExtra.dll"

pause