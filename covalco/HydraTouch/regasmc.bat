@echo off
REM
REM Fill in your DLL directory on the cd line
REM Fill in your DLL name on the two RegAsm lines
REM
echo "Registering (thing)..."
cd "%programfiles(x86)%\Location\OfMy\Library"
echo "Changed directory. Trying 32 bit..."
%WINDIR%\Microsoft.Net\Framework\v4.0.30319\RegAsm.exe WinSCPnet.dll /codebase /tlb:WinSCPnet32.tlb
if ERRORLEVEL 100 (
	echo "Trying 64 bit..."
	%WINDIR%\Microsoft.Net\Framework64\v4.0.30319\RegAsm.exe WinSCPnet.dll /codebase /tlb:WinSCPnet32.tlb
)
echo "Finished registering"
echo:
echo %ERRORLEVEL%