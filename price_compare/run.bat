rem refer: https://stackoverflow.com/questions/44577446/open-file-by-dragging-and-dropping-it-onto-batch-file

@echo off
Title EzeeShip Price Compare

IF [%1] EQU [] (
	rem running with default parameters
	python3 compare.py
) ELSE (
	rem running with drag-and-drop file as input data file
	python3 compare.py -f "%~1"
)

PAUSE