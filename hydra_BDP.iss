[Setup]
AppName=HydraTouch
AppVerName=HydraTouchLite 4.0.1 BDP Artcon (10/02/26)
OutputDir=output\HydraTouchLite-BDP-4.0.1
AppPublisher=Grupo Epelsa
AppPublisherURL=http://www.grupoepelsa.com
AppSupportURL=http://www.grupoepelsa.com
AppUpdatesURL=http://www.grupoepelsa.com
DefaultDirName={pf}\HydraTouch
DefaultGroupName=Hydra
AlwaysShowComponentsList=true

[Types]
Name: hydratouch;    Description: HydraTouch;     flags: iscustom;

[Components]
Name: DCOM95;           Description: DCOM for Windows 95;                       Flags: disablenouninstallwarning restart fixed; MinVersion: 4.0,0;      OnlyBelowVersion: 4.1,0;        Types: hydratouch;
Name: DCOM98;           Description: DCOM for Windows 98;                       Flags: disablenouninstallwarning restart fixed; MinVersion: 4.1,0;                                      Types: hydratouch;
Name: MDAC25;           Description: Microsoft Data Access Components 2.5;      Flags: disablenouninstallwarning restart fixed; MinVersion: 4.0,4.0;    OnlyBelowVersion: 4.9,5.0;      Types: hydratouch;
Name: MDAC27;           Description: Microsoft Data Access Components 2.7;      Flags: disablenouninstallwarning restart fixed; MinVersion: 4.9,5.0;                                    Types: hydratouch;

[Tasks]
Name: "desktopicon";            Description: "Crear Icono en el Escritorio";    GroupDescription: "Additional icons:";  MinVersion: 4,4;        Flags: unchecked
Name: "quicklaunchicon";        Description: "Crear Icono de Acceso Rápido";    GroupDescription: "Additional icons:";  MinVersion: 4,4;        Flags: unchecked

[Files]
Source: hydraTouch.exe;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: tiqahTouch.exe;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: vacio.jpg;             DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: backgrounde.jpg;       DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: hydra-background-1024x768.jpg;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: f1.jpg;                DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: auxiliartouch.exe;     DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: zip.exe;          DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: unzip.exe;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: copiontouch.exe;       DestDir: "{app}"; CopyMode: alwaysoverwrite;
Source: C:\inno_appEPELSA\BDP\bal.dat.fia;       DestDir: "{app}\etc"; CopyMode: alwaysoverwrite;
Source: etc\artcon.dat.fia;       DestDir: "{app}\etc"; CopyMode: alwaysoverwrite;
Source: cmdlinetouch.cfg;       DestDir: "{app}"; CopyMode: alwaysoverwrite;
Source: C:\inno_appEPELSA\BDP\bdp.txt;           DestDir: "{app}"; CopyMode: alwaysoverwrite;
Source: bdpnot.txt;           DestDir: "{app}"; CopyMode: alwaysoverwrite;
Source: ctrlimpatouch.ocx;     DestDir: "{app}"; CopyMode: alwaysoverwrite; Flags : sharedfile regserver
Source: langtouch\langtouch.mdb;    DestDir: "{app}\langtouch";  CopyMode: alwaysoverwrite
Source: langtouch\1.ico;       DestDir: "{app}\langtouch";  CopyMode: alwaysoverwrite
Source: langtouch\2.ico;       DestDir: "{app}\langtouch";  CopyMode: alwaysoverwrite
Source: langtouch\3.ico;       DestDir: "{app}\langtouch";  CopyMode: alwaysoverwrite
Source: langtouch\4.ico;       DestDir: "{app}\langtouch";  CopyMode: alwaysoverwrite
Source: langtouch\5.ico;       DestDir: "{app}\langtouch";  CopyMode: alwaysoverwrite
Source: langtouch\6.ico;       DestDir: "{app}\langtouch";  CopyMode: alwaysoverwrite
Source: langtouch\7.ico;       DestDir: "{app}\langtouch";  CopyMode: alwaysoverwrite
Source: langtouch\8.ico;       DestDir: "{app}\langtouch";  CopyMode: alwaysoverwrite
Source: langtouch\9.ico;       DestDir: "{app}\langtouch";  CopyMode: alwaysoverwrite
Source: langtouch\langtouch.mdb;    DestDir: "{app}\lang";  CopyMode: alwaysoverwrite
Source: langtouch\1.ico;       DestDir: "{app}\lang";  CopyMode: alwaysoverwrite
Source: langtouch\2.ico;       DestDir: "{app}\lang";  CopyMode: alwaysoverwrite
Source: langtouch\3.ico;       DestDir: "{app}\lang";  CopyMode: alwaysoverwrite
Source: langtouch\4.ico;       DestDir: "{app}\lang";  CopyMode: alwaysoverwrite
Source: langtouch\5.ico;       DestDir: "{app}\lang";  CopyMode: alwaysoverwrite
Source: langtouch\6.ico;       DestDir: "{app}\lang";  CopyMode: alwaysoverwrite
Source: langtouch\7.ico;       DestDir: "{app}\lang";  CopyMode: alwaysoverwrite
Source: langtouch\8.ico;       DestDir: "{app}\lang";  CopyMode: alwaysoverwrite
Source: langtouch\9.ico;       DestDir: "{app}\lang";  CopyMode: alwaysoverwrite
Source: etc\*.fet;       DestDir: "{app}\etc";  CopyMode: alwaysoverwrite
Source: images\*.*;     destdir: "{app}\images"; copymode: alwaysoverwrite
Source: C:\inno_appEPELSA\3rdparty\inno.txt;        DestDir: "{app}\3rdparty";      CopyMode: alwaysoverwrite
Source: C:\inno_appEPELSA\3rdparty\zip_unzip.txt;   DestDir: "{app}\3rdparty";      CopyMode: alwaysoverwrite
Source: C:\inno_appEPELSA\3rdparty\xcopy.zip;       DestDir: "{app}\3rdparty";      CopyMode: alwaysoverwrite
Source: systray.ocx;     DestDir: "{app}"; CopyMode: alwaysoverwrite;         Flags : sharedfile regserver
source: unrar.dll;                            destdir: "{app}"; CopyMode: alwaysoverwrite;
Source: support\systray.ocx;     DestDir: "{sys}"; CopyMode: alwaysoverwrite;         Flags : sharedfile regserver
Source: support\RICHTX32.OCX;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\COMDLG32.OCX;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\DBGRID32.OCX;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSCHRT20.OCX;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSCOMCTL.OCX;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSCOMM32.OCX;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSMASK32.OCX;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSWINSCK.OCX;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\TABCTL32.OCX;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\RCHTXES.DLL;     DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\RICHED32.DLL;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: c:\inno_appepelsa\system\DAO350.DLL; DestDir: "{dao}"; CopyMode: alwaysskipifsameorolder; Flags: allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: c:\inno_appepelsa\system\DAO360.DLL; DestDir: "{dao}"; CopyMode: alwaysskipifsameorolder; Flags: allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\ASYCFILT.DLL;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: c:\inno_appEPELSA\system\COMCAT.DLL; DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\DBGRDES.DLL;     DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\DBRPRES.DLL;     DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\expsrv.dll;      DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\MSBIND.DLL;      DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSCH2ES.DLL;     DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\MSCMCES.DLL;     DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\MSCOMES.DLL;     DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\MSDBRPTR.DLL;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\msderun.dll;     DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSJET35.DLL;     DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSJINT35.DLL;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\MSJTER35.DLL;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\msjtes40.dll;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSJET40.DLL;     DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSJINT40.DLL;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\MSJTER40.DLL;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\MSMSKES.DLL;     DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\MSPRPES.DLL;     DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\MSRD2X35.DLL;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSRD2X40.DLL;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSRD3X40.DLL;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSREPL35.DLL;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\MSREPL40.DLL;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\MSSTDFMT.DLL;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSSTKPRP.DLL;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSVBVM60.DLL;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSXBSE35.DLL;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\odbcconf.dll;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\OLEAUT32.DLL;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\OLEPRO32.DLL;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\SCRRNES.DLL;     DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\SCRRUN.DLL;      DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: support\STDFTES.DLL;     DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\TABCTES.DLL;     DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\VB5DB.DLL;       DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\VB6ES.DLL;       DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\VB6STKIT.DLL;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\VBAJET32.DLL;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\WINSKES.DLL;     DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\RDOCURS.DLL;     DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\RDO20ES.DLL;     DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\STDOLE2.TLB;     DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\MSVCRT40.DLL;    DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile
Source: support\MSRDO20.DLL;     DestDir: "{sys}"; Flags: onlyifdoesntexist allowunsafefiles restartreplace uninsneveruninstall sharedfile regserver
Source: C:\inno_appEPELSA\lib\dcom95.EXE;         DestDir: {tmp}; CopyMode: alwaysoverwrite; MinVersion: 4.0,0;   Components: DCOM95;     OnlyBelowVersion: 4.1,0;
Source: C:\inno_appEPELSA\lib\dcom98.exe;         DestDir: {tmp}; CopyMode: alwaysoverwrite; MinVersion: 4.1,0;   Components: DCOM98;
Source: C:\inno_appEPELSA\lib\mdac_typ_25.exe;    DestDir: {tmp}; CopyMode: alwaysoverwrite; MinVersion: 4.0,4.0; Components: MDAC25;      OnlyBelowVersion: 4.9,5.0;
Source: C:\inno_appEPELSA\lib\mdac_typ_27.exe;    DestDir: {tmp}; CopyMode: alwaysoverwrite; MinVersion: 4.9,5.0; Components: MDAC27;

[Icons]
Name: "{group}\HydraTouch";                                                  Filename: "{app}\hydraTouch.exe"; WorkingDir: {app}
Name: "{userdesktop}\HydraTouch";                                            Filename: "{app}\hydraTouch.exe"; WorkingDir: {app};   MinVersion: 4,4;        Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\HydraTouch";   Filename: "{app}\hydraTouch.exe"; WorkingDir: {app};   MinVersion: 4,4;        Tasks: quicklaunchicon

[Dirs]
Name: "{app}\etc"
Name: "{app}\lang"
Name: "{app}\langtouch"
Name: "{app}\dirtmp"
Name: "{app}\images"
Name: "{app}\exporta"
Name: "{app}\backup"
Name: "{app}\images"
Name: "(app)\3rdparty"
Name: "{app}\asorden"

[Run]
Filename: {tmp}\dcom95.EXE;             Parameters: /r:n /q:u;                  WorkingDir: {tmp};      Flags: skipifdoesntexist;       Components: DCOM95;     MinVersion: 4.0,0;      OnlyBelowVersion: 4.1,0
Filename: {tmp}\dcom98.exe;             Parameters: /r:n /q:u;                  WorkingDir: {tmp};      Flags: skipifdoesntexist;       Components: DCOM98;     MinVersion: 4.1,0
Filename: {tmp}\mdac_typ_25.exe;        Parameters: "/Q /C:""setup /QNT""";     WorkingDir: {tmp};      Flags: skipifdoesntexist;       Components: MDAC25;     MinVersion: 4.0,4.0;    OnlyBelowVersion: 4.9,5.0
Filename: {tmp}\mdac_typ_27.exe;        Parameters: "/Q /C:""setup /QNT""";     WorkingDir: {tmp};      Flags: skipifdoesntexist;       Components: MDAC27;     MinVersion: 4.9,5.0;
Filename: "{app}\hydratouch.exe";            Description: "Launch Hydra";            Flags: nowait postinstall skipifsilent