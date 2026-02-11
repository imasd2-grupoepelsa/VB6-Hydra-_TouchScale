[Setup]
AppName=HydraTouch
AppVerName=HydraTouchLite Parche BDP Artcon (10/02/2026)
OutputDir=output\BDP-4.0.1\parche
AppPublisher=Grupo Epelsa
AppPublisherURL=http://www.grupoepelsa.com
AppSupportURL=http://www.grupoepelsa.com
AppUpdatesURL=http://www.grupoepelsa.com
DefaultDirName={pf}\HydraTouch\bdp\parche
DefaultGroupName=Hydra
AlwaysShowComponentsList=true

[Files]
Source: hydraTouch.exe;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: tiqahTouch.exe;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: vacio.jpg;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: backgrounde.jpg;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: hydra-background-1024x768.jpg;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: f1.jpg;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: auxiliartouch.exe;     DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: zip.exe;          DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: unzip.exe;        DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: copiontouch.exe;       DestDir: "{app}"; CopyMode: alwaysoverwrite;
Source: support\bdp\bal.dat.fia;       DestDir: "{app}\etc"; CopyMode: alwaysoverwrite;
Source: etc\artcon.dat.fia;       DestDir: "{app}\etc"; CopyMode: alwaysoverwrite;
Source: cmdlinetouch.cfg;       DestDir: "{app}"; CopyMode: alwaysoverwrite;
Source: support\bdp\bdp.txt;           DestDir: "{app}"; CopyMode: alwaysoverwrite;
Source: bdpnot.txt;           DestDir: "{app}"; CopyMode: alwaysoverwrite;
Source: cmdlinetouch.cfg;           DestDir: "{app}"; CopyMode: alwaysoverwrite;
Source: ctrlimpatouch.ocx;     DestDir: "{app}"; CopyMode: alwaysoverwrite; Flags : sharedfile regserver