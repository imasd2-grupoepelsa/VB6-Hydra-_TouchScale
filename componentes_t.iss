[Setup]
AppName=comp
AppVerName=comp
OutputDir=output\comp
DefaultDirName={pf}\comp

[Types]
Name: comp;    Description: comp;     flags: iscustom;

[Components]
Name: MDAC25;           Description: Microsoft Data Access Components 2.5;      Flags: disablenouninstallwarning restart fixed; MinVersion: 4.0,4.0;    OnlyBelowVersion: 4.9,5.0;      Types: comp;
Name: MDAC27;           Description: Microsoft Data Access Components 2.7;      Flags: disablenouninstallwarning restart fixed; MinVersion: 4.9,5.0;                                    Types: comp;

[Files]
source: support\unrar.dll;                            destdir: "{app}"; CopyMode: alwaysoverwrite;
Source: support\systray.ocx;     DestDir: "{sys}"; CopyMode: alwaysoverwrite;         Flags : sharedfile regserver
Source: support\RICHTX32.OCX;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\COMDLG32.OCX;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\DBGRID32.OCX;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSCHRT20.OCX;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSCOMCTL.OCX;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSCOMM32.OCX;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSMASK32.OCX;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSWINSCK.OCX;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\TABCTL32.OCX;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\RCHTXES.DLL;     DestDir: "{app}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\RICHED32.DLL;    DestDir: "{app}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\DAO350.DLL;      DestDir: "{dao}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\DAO360.DLL;      DestDir: "{dao}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\ASYCFILT.DLL;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\COMCAT.DLL;      DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\DBGRDES.DLL;     DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\DBRPRES.DLL;     DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\expsrv.dll;      DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\MSBIND.DLL;      DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSCH2ES.DLL;     DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\MSCMCES.DLL;     DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\MSCOMES.DLL;     DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\MSDBRPTR.DLL;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\msderun.dll;     DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSJET35.DLL;     DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSJINT35.DLL;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\MSJTER35.DLL;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\msjtes40.dll;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSJET40.DLL;     DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSJINT40.DLL;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\MSJTER40.DLL;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\MSMSKES.DLL;     DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\MSPRPES.DLL;     DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\MSRD2X35.DLL;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSRD2X40.DLL;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSRD3X40.DLL;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSREPL35.DLL;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\MSREPL40.DLL;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\MSSTDFMT.DLL;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSSTKPRP.DLL;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSVBVM60.DLL;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\MSXBSE35.DLL;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\odbcconf.dll;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\OLEAUT32.DLL;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\OLEPRO32.DLL;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\SCRRNES.DLL;     DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\SCRRUN.DLL;      DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\STDFTES.DLL;     DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\TABCTES.DLL;     DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\VB5DB.DLL;       DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\VB6ES.DLL;       DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\VB6STKIT.DLL;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\VBAJET32.DLL;    DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\WINSKES.DLL;     DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile
Source: support\RDOCURS.DLL;     DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\RDO20ES.DLL;     DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: support\STDOLE2.TLB;     DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: support\MSVBVM60.DLL; DestDir: "{app}\support"; CopyMode: alwaysoverwrite; Flags: sharedfile regserver allowunsafefiles
Source: support\OLEAUT32.DLL; DestDir: "{app}\support"; CopyMode: alwaysoverwrite; Flags: sharedfile regserver allowunsafefiles
Source: support\OLEPRO32.DLL; DestDir: "{app}\support"; CopyMode: alwaysoverwrite; Flags: sharedfile regserver allowunsafefiles
Source: support\ASYCFILT.DLL; DestDir: "{app}\support"; CopyMode: alwaysoverwrite
Source: support\STDOLE2.TLB;  DestDir: "{app}\support"; CopyMode: alwaysoverwrite
Source: support\COMCAT.DLL;   DestDir: "{app}\support"; CopyMode: alwaysoverwrite; Flags: sharedfile regserver allowunsafefiles
Source: support\VB6STKIT.DLL; DestDir: "{app}\support"; CopyMode: alwaysoverwrite
Source: support\VB6ES.DLL;    DestDir: "{app}\support"; CopyMode: alwaysoverwrite
Source: support\MSWINSCK.OCX; DestDir: "{app}\support"; CopyMode: alwaysoverwrite; Flags: sharedfile regserver
Source: support\WINSKES.DLL;  DestDir: "{app}\support"; CopyMode: alwaysoverwrite
Source: support\MSCOMM32.OCX; DestDir: "{app}\support"; CopyMode: alwaysoverwrite; Flags: sharedfile regserver
Source: support\MSCOMES.DLL;  DestDir: "{app}\support"; CopyMode: alwaysoverwrite
Source: support\MSSTKPRP.DLL; DestDir: "{app}\support"; CopyMode: alwaysoverwrite; Flags: sharedfile regserver
Source: support\MSPRPES.DLL;  DestDir: "{app}\support"; CopyMode: alwaysoverwrite
Source: support\MSCOMCTL.OCX; DestDir: "{app}\support"; CopyMode: alwaysoverwrite; Flags: sharedfile regserver
Source: support\MSCMCES.DLL;  DestDir: "{app}\support"; CopyMode: alwaysoverwrite
Source: support\MSVCRT40.DLL; DestDir: "{app}\support"; CopyMode: alwaysoverwrite
Source: support\RDO20ES.DLL;  DestDir: "{app}\support"; CopyMode: alwaysoverwrite; Flags: sharedfile
Source: support\RDOCURS.DLL;  DestDir: "{app}\support"; CopyMode: alwaysoverwrite; Flags: sharedfile
Source: support\MSRDO20.DLL;  DestDir: "{app}\support"; CopyMode: alwaysoverwrite; Flags: sharedfile regserver
Source: support\VB5DB.DLL;    DestDir: "{app}\support"; CopyMode: alwaysoverwrite; Flags: sharedfile
Source: support\MSREPL35.DLL; DestDir: "{app}\support"; CopyMode: alwaysoverwrite
Source: support\MSRD2X35.DLL;  DestDir: "{app}\support"; CopyMode: alwaysoverwrite; Flags: sharedfile regserver
Source: support\expsrv.dll; DestDir: "{app}\support"; CopyMode: alwaysoverwrite
Source: support\vbajet32.dll; DestDir: "{app}\support"; CopyMode: alwaysoverwrite
Source: support\msjint35.dll; DestDir: "{app}\support"; CopyMode: alwaysoverwrite
Source: support\MSJTER35.DLL; DestDir: "{app}\support"; CopyMode: alwaysoverwrite
Source: support\MSJET35.DLL;  DestDir: "{app}\support"; CopyMode: alwaysoverwrite; Flags: sharedfile regserver
Source: C:\inno_appEPELSA\lib\mdac_typ_25.exe;    DestDir: {tmp}; CopyMode: alwaysoverwrite; MinVersion: 4.0,4.0; Components: MDAC25;      OnlyBelowVersion: 4.9,5.0;
Source: C:\inno_appEPELSA\lib\mdac_typ_27.exe;    DestDir: {tmp}; CopyMode: alwaysoverwrite; MinVersion: 4.9,5.0; Components: MDAC27;
Source: mysql-connector-odbc-5.1.11-win32.msi;    DestDir: "{app}"; CopyMode: alwaysoverwrite

