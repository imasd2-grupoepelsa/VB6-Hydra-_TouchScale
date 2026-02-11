[Setup]
AppName=HydraTouch
AppVerName=HydraTouchLite 4.0.1 Patch Fornes (100226) I.M. TICF,RCC,DCF,TCF. TT AFG
OutputDir=output\stdTouchLite-4.0.1\fornes\parche
AppPublisher=Grupo Epelsa
AppPublisherURL=http://www.grupoepelsa.com
AppSupportURL=http://www.grupoepelsa.com
AppUpdatesURL=http://www.grupoepelsa.com
DefaultDirName={pf}\HydraTouch
DefaultGroupName=Hydra
AlwaysShowComponentsList=true

[Files]
Source: hydraTouch.exe;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: tiqahTouch.exe;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: vacio.jpg;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: backgrounde.jpg;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: hydra-background-1024x768_1.jpg;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: f1.jpg;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: auxiliartouch.exe;     DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: zip.exe;          DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: unzip.exe;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: copiontouch.exe;       DestDir: "{app}"; CopyMode: alwaysoverwrite;
Source: fornes\etc\*.*;        DestDir: "{app}\etc"; CopyMode: alwaysoverwrite
Source: fornes\cmdlinetouch.cfg;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: ctrlimpatouch.ocx;     DestDir: "{app}"; CopyMode: alwaysoverwrite; Flags : sharedfile regserver
Source: sidebar.ocx;           DestDir: "{app}"; CopyMode: alwaysoverwrite; Flags : sharedfile regserver
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