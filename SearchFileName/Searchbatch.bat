@echo off
cls
:start
echo ============================= Batch Search File ===========================================
echo ============================= Create BY : Theedanai Poomilamnao 03/10/2022 =============================
echo.

set required=false
set document=false
 
set /p "word=Wording that you want to find. : "
echo. 
set /p "fe=File Extension ( '*' if you want search all File Extension type). : "
echo.

if "%word%"=="" set required=true
if "%fe%"=="" set required=true

if "%required%"=="true" ( 
	set /p "end=Wording and File Extension are required field! Please try again next time."
	exit 
) 

if "%fe%"=="doc" set document=true
if "%fe%"=="docx" set document=true
if "%fe%"=="xls" set document=true
if "%fe%"=="xlsx" set document=true
if "%fe%"=="ppt" set document=true
if "%fe%"=="pptx" set document=true
if "%fe%"=="pps" set document=true
if "%fe%"=="ppsx" set document=true
if "%fe%"=="vsd" set document=true
if "%fe%"=="vsdx" set document=true
if "%fe%"=="odt" set document=true
if "%fe%"=="ods" set document=true
if "%fe%"=="pdf" set document=true

echo ----------------------------------- Result Search ----------------------------------------------
echo. 
if "%document%"=="true" ( 
	findstr /s /i /m %word% *.%fe% 
) else ( 
	findstr /s /i %word% *.%fe% 
)
echo.
set /p "end=----------------------------------- Finished! Press Any key to Finish Search ---------------------"
echo.
goto start