setlocal

set MSBuildPath="c:\Program Files\Microsoft Visual Studio\2022\Preview\Msbuild\Current\Bin\amd64\MSBuild.exe"
set OutFileName=ExcelDna.AddInManager.msi

del /q %OutFileName%

%MSBuildPath% ..\Source\AddInManager.sln /t:restore,build /p:Configuration=Release
@if errorlevel 1 goto end

xcopy ..\Source\ExcelDna.AddInManager\bin\Release\net6.0-windows\publish ..\Source\Installer\ExcelAddInDeploy\SourceFiles /I /Y

"c:\Windows\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe" ..\Source\Installer\Installer.sln /t:Rebuild /p:Configuration=Release
@if errorlevel 1 goto end

copy "..\Source\Installer\ExcelAddInDeploy\bin\Release\en-us\ExcelAddInDeploy.msi" %OutFileName%

:end
