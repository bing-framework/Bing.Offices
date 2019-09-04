@echo off
echo %1
cd ..

rem clear old packages
del output\* /q/f/s

rem build
dotnet build Bing.Offices.sln -c Release

rem pack
dotnet pack ./src/Bing.Offices.Abstractions/Bing.Offices.Abstractions.csproj
dotnet pack ./src/Bing.Offices.Core/Bing.Offices.Core.csproj
dotnet pack ./src/Bing.Offices.Npoi/Bing.Offices.Npoi.csproj

rem push
for %%i in (output\*.nupkg) do dotnet nuget push %%i -k %1 -s https://www.nuget.org/api/v2/package