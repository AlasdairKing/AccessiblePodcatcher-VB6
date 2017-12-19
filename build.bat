@echo off
path=%path%;C:\Program Files\Microsoft Visual Studio\VB98\;C:\Program Files\7-Zip;C:\Program Files\Microsoft SDKs\Windows\v6.0\Bin\;C:\Program Files\SoftwarePassport;C:\Program Files\WinZip Self-Extractor\;C:\Program Files\Crimson Editor

echo ==================
echo Doing code store 
7z a -tzip podcatcher.zip @filetypestozip.txt
echo.
echo ===================
echo Compiling to D:\Installers\Alasdair\files\Program Files\WebbIE
vb6 /m "AccessiblePodcatcher.vbp"  /outdir "D:\Installers\Alasdair\Files\Program Files\WebbIE"
echo.
echo Copying to Powerwraps
copy /Y "D:\Installers\Alasdair\files\Program Files\WebbIE\AccessiblePodcatcher.exe" "D:\Installers\Powerwraps\WebbIE"
echo.
pause