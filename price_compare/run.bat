@echo off

rem refer: https://stackoverflow.com/questions/44577446/open-file-by-dragging-and-dropping-it-onto-batch-file

Title EzeeShip Price Compare

IF [%1] EQU [] (
	rem running with default parameters
	python compare.py
	rem python3 only for python2 pre-installed system, like Linux, Mac, ...
) ELSE (
	rem running with drag-and-drop file as input data file. even if not under the same folder
	cd /d %~dp0
	python compare.py -s "%~1"
)

PAUSE