[Setup]
AppName=Hydra+
AppVerName=HydraTouch 4.0.1 patch MT-Rio
OutputDir=output\mettler\parche
AppPublisher=Grupo Epelsa
AppPublisherURL=http://www.grupoepelsa.com
AppSupportURL=http://www.grupoepelsa.com
AppUpdatesURL=http://www.grupoepelsa.com
DefaultDirName={pf}\Hydra
DefaultGroupName=Hydra
AlwaysShowComponentsList=true
PrivilegesRequired=admin

[Types]
Name: hydra; Description: Hydra; flags: iscustom;

[Tasks]
Name: "desktopicon";            Description: "Crear Icono en el Escritorio"; GroupDescription: "Additional icons:"; MinVersion: 4,4;    Flags: unchecked
Name: "quicklaunchicon";        Description: "Crear Icono de Acceso Rápido"; GroupDescription: "Additional icons:"; MinVersion: 4,4;    Flags: unchecked

[Files]
Source: support\inno.txt;        DestDir: "{app}\3rdparty";      CopyMode: alwaysoverwrite
Source: support\zip_unzip.txt;   DestDir: "{app}\3rdparty";      CopyMode: alwaysoverwrite
Source: support\xcopy.zip;       DestDir: "{app}\3rdparty";      CopyMode: alwaysoverwrite
Source: support\mettler\in.txt.fia;       DestDir: "{app}\etc";           CopyMode: alwaysoverwrite
Source: support\mettler\touch\cmdlinetouch.cfg;      DestDir: "{app}";               CopyMode: alwaysoverwrite
Source: support\mettler\tiendamt.txt;     DestDir: "{app}";      CopyMode: alwaysoverwrite
Source: support\mettler\touch\tiqahtouch.ini;        DestDir: "{app}";      CopyMode: alwaysoverwrite
Source: support\mettler\tramos.dat.fia;       DestDir: "{app}\etc";  CopyMode: alwaysoverwrite