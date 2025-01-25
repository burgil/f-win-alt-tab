@echo off
title f-win-alt-tab by Burgil - 2025
color a
echo Running in the background and in a new window...
start /B "" C:\Windows\System32\conhost.exe cmd.exe /k "python f-win-alt-tab.py --hide && exit"
exit