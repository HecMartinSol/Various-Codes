@ECHO off
REM
REM Muestra un Pop-up a los usuarios al incio de sesion 
REM Dependiendo de la fecha muestra el archivo elegido
REM
REM 17/12/2014
REM

REM -------------------------------------
REM GETS DAY, MONTH AND YEAR
REM -------------------------------------
set month=%Date:~3,2%
set day=%Date:~0,2%
set year=%Date:~6,4%
set arg=""

IF %month% equ 07 if %day% == 03 set arg=1.pps
IF %month% equ 07 if %day% == 04 set arg=1.pps

IF %month% equ 07 if %day% == 15 set arg=2.pps
IF %month% equ 07 if %day% == 17 set arg=2.pps

IF %month% equ 07 if %day% == 23 set arg=3.pps
IF %month% equ 07 if %day% == 31 set arg=3.pps



IF %month% equ 09 if %day% == 09 set arg=4.pps
IF %month% equ 09 if %day% == 11 set arg=4.pps

IF %month% equ 09 if %day% == 23 set arg=5.pps
IF %month% equ 09 if %day% == 26 set arg=5.pps



IF %month% equ 10 if %day% == 07 set arg=6.pps
IF %month% equ 10 if %day% == 09 set arg=6.pps



REM --------------------------TESTS--------------------------

	set day=99
	set month=99
	
REM	IF %month% equ 99 if %day% == 99 set arg=pc.gif
REM	IF %month% equ 99 if %day% == 99 set arg=pc.jpg
REM	IF %month% equ 99 if %day% == 99 set arg=pc.jpeg
REM	IF %month% equ 99 if %day% == 99 set arg=pc.bmp
REM	IF %month% equ 99 if %day% == 99 set arg=pc.png

REM	IF %month% equ 99 if %day% == 99 set arg=formulario.pdf

REM	IF %month% equ 99 if %day% == 99 set arg=fileTree@18-12-2014@11_20_06.txt

REM	IF %month% equ 99 if %day% == 99 set arg=1.pps

REM	IF %month% equ 99 if %day% == 99 set arg=presentacion.pptx
REM	IF %month% equ 99 if %day% == 99 set arg=uno.ppt
	
	IF %month% equ 99 if %day% == 99 set arg=documento.docx
REM --------------------------TESTS--------------------------

start filescript.vbs %arg%

exit
