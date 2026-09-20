@echo off
rem Double-click launcher. Korean text is left out on purpose: a .bat is read as cp949.
cd /d "%~dp0"
uv run broadcast-ppt gui || pause
