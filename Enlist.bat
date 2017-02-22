@echo off
%debug%
sd.exe help >nul || (echo ERROR: Make sure SD.EXE is on your path & pause & goto :eof)
MD \CommandLineParser
CD \CommandLineParser
echo sdport=tkbgitsd01:4001>SD.INI
echo sdclient=%ComputerName%_CmdLineParsers>>SD.INI
echo *****************************************************************************
echo Change Your view to //depot/CmdLineParsers/... //%ComputerName%_CmdLineParsers/...
echo *****************************************************************************
SD.EXE client
echo After changing your view, hit any key to enlist
pause
sd.exe sync
echo Done! Thank you for contributing!
pause