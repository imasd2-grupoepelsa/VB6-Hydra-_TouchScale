[Setup]
AppName=HydraTouch
AppVerName=HydraTouchLite 4.0.1 (10/02/2026)
OutputDir=output\stdTouchLite\parche_reducido
AppPublisher=Grupo Epelsa
AppPublisherURL=http://www.grupoepelsa.com
AppSupportURL=http://www.grupoepelsa.com
AppUpdatesURL=http://www.grupoepelsa.com
DefaultDirName={pf}\HydraTouch
DefaultGroupName=Hydra
AlwaysShowComponentsList=true

[Files]
Source: hydraTouch.exe;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: vacio.jpg;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: backgrounde.jpg;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: hydra-background-1024x768.jpg;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: f1.jpg;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: auxiliartouch.exe;     DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: zip.exe;          DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: unzip.exe;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: copiontouch.exe;       DestDir: "{app}"; CopyMode: alwaysoverwrite;
Source: ctrlimpatouch.ocx;     DestDir: "{app}"; CopyMode: alwaysoverwrite; Flags : sharedfile regserver