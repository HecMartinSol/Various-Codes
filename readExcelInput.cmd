@ECHO off
SET id=-1
SET sheets=
SET numSheets=0

SET collumns=
SET numColumns=0



ECHO ENTER SHEET NAMES. PRESS ENTER TO END
ECHO /!\ SHEET NAMES COULDN'T CONTAIN SPACES
:LoopStart
		REM Do something
		SET /p id="  >"
		REM ECHO %id%
		
		IF NOT %id% == -1 SET sheets=%sheets% %id%
		IF NOT %id% == -1 SET /a numSheets=%numSheets%+1
		

		REM Break from the loop if a condition is met
		IF %id% == -1 GOTO LoopEnd
		SET id=-1
		REM Iterate through the loop once more if the condition wasn't met
		GOTO LoopStart
:LoopEnd
	REM You're out of the loop now
	ECHO -----------------------------------------


ECHO ENTER COLLUMN LETTERS. PRESS ENTER TO END
:LoopStart2
		REM Do something
		SET /p id="  >"
		REM ECHO %id%
		
		IF NOT %id% == -1 SET collumns=%collumns% %id%
		IF NOT %id% == -1 SET /a numCollumns=%numCollumns%+1
		
		REM ECHO %collumns%

		REM Break from the loop if a condition is met
		IF %id% == -1 GOTO LoopEnd2
		SET id=-1
		REM Iterate through the loop once more if the condition wasn't met
		GOTO LoopStart2
:LoopEnd2
	REM You're out of the loop now
	ECHO -----------------------------------------	
	IF NOT %numSheets% == 0 IF NOT %numCollumns% == 0 	start readExcelInput.vbs %sheets% 1 %collumns% 1
