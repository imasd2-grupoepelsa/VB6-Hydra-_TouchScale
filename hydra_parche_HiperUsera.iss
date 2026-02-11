[Setup]
AppName=HydraTouch
AppVerName=HydraTouchLite 4.0.1 Patch HiperUsera (10022026)
OutputDir=output\stdTouchLiteHiperUsera401\parche
AppPublisher=Grupo Epelsa
AppPublisherURL=http://www.grupoepelsa.com
AppSupportURL=http://www.grupoepelsa.com
AppUpdatesURL=http://www.grupoepelsa.com
DefaultDirName={pf}\HydraTouch
DefaultGroupName=Hydra
AlwaysShowComponentsList=true

[Files]
Source: hiperusera\*.exe;      DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: hiperusera\*.txt;      DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: auxiliartouch.exe;     DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: copiontouch.exe;       DestDir: "{app}"; CopyMode: alwaysoverwrite;
Source: ctrlimpatouch.ocx;     DestDir: "{app}"; CopyMode: alwaysoverwrite; Flags : sharedfile regserver
Source: sidebar.ocx;           DestDir: "{app}"; CopyMode: alwaysoverwrite; Flags : sharedfile regserver
Source: support\hyperlbl.ocx;     DestDir: "{app}"; CopyMode: alwaysoverwrite;                          Flags: sharedfile regserver

