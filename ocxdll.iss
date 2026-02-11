[Setup]
AppName=DLL/OCX Comunicaciones Sistemas Grupo Epelsa
AppVerName=DLL Versión 4.0.1, OCX Versión 4.0.1
AppPublisher=Grupo Epelsa
AppPublisherURL=http://www.grupoepelsa.com
AppSupportURL=http://www.grupoepelsa.com
AppUpdatesURL=http://www.grupoepelsa.com
DefaultDirName={pf}\EPDLLOCX

[Types]
Name: DLL ; Description: Utilidad Comunicación;

[Files]
Source: support\EpelDLL.dll; DestDir: {app}; CopyMode: alwaysoverwrite; Flags: sharedfile regserver
Source: support\EpelOCX.ocx; DestDir: {app}; CopyMode: alwaysoverwrite; Flags: sharedfile regserver
Source: support\epelcom_es.pdf; DestDir: {app}; CopyMode: alwaysoverwrite
Source: support\c9sec.cfg; DestDir: {app}; CopyMode: alwaysoverwrite
Source: support\MSVBVM60.DLL; DestDir: {app}\support; CopyMode: alwaysoverwrite; Flags: sharedfile regserver allowunsafefiles
Source: support\OLEAUT32.DLL; DestDir: {app}\support; CopyMode: alwaysoverwrite; Flags: sharedfile regserver allowunsafefiles
Source: support\OLEPRO32.DLL; DestDir: {app}\support; CopyMode: alwaysoverwrite; Flags: sharedfile regserver allowunsafefiles
Source: support\ASYCFILT.DLL; DestDir: {app}\support; CopyMode: alwaysoverwrite
Source: support\STDOLE2.TLB;  DestDir: {app}\support; CopyMode: alwaysoverwrite
Source: support\COMCAT.DLL;   DestDir: {app}\support; CopyMode: alwaysoverwrite; Flags: sharedfile regserver allowunsafefiles
Source: support\VB6STKIT.DLL; DestDir: {app}\support; CopyMode: alwaysoverwrite
Source: support\VB6ES.DLL;    DestDir: {app}\support; CopyMode: alwaysoverwrite
Source: support\MSWINSCK.OCX; DestDir: {app}\support; CopyMode: alwaysoverwrite; Flags: sharedfile regserver
Source: support\WINSKES.DLL;  DestDir: {app}\support; CopyMode: alwaysoverwrite
Source: support\MSCOMM32.OCX; DestDir: {app}\support; CopyMode: alwaysoverwrite; Flags: sharedfile regserver
Source: support\MSCOMES.DLL;  DestDir: {app}\support; CopyMode: alwaysoverwrite
Source: support\MSSTKPRP.DLL; DestDir: {app}\support; CopyMode: alwaysoverwrite; Flags: sharedfile regserver
Source: support\MSPRPES.DLL;  DestDir: {app}\support; CopyMode: alwaysoverwrite
Source: support\MSCOMCTL.OCX; DestDir: {app}\support; CopyMode: alwaysoverwrite; Flags: sharedfile regserver
Source: support\MSCMCES.DLL;  DestDir: {app}\support; CopyMode: alwaysoverwrite
Source: support\MSVCRT40.DLL; DestDir: {app}\support; CopyMode: alwaysoverwrite
Source: support\RDO20ES.DLL;  DestDir: {app}\support; CopyMode: alwaysoverwrite; Flags: sharedfile
Source: support\RDOCURS.DLL;  DestDir: {app}\support; CopyMode: alwaysoverwrite; Flags: sharedfile
Source: support\MSRDO20.DLL;  DestDir: {app}\support; CopyMode: alwaysoverwrite; Flags: sharedfile regserver
Source: support\VB5DB.DLL;    DestDir: {app}\support; CopyMode: alwaysoverwrite; Flags: sharedfile
Source: support\MSREPL35.DLL; DestDir: {app}\support; CopyMode: alwaysoverwrite
Source: support\MSRD2X35.DLL;  DestDir: {app}\support; CopyMode: alwaysoverwrite; Flags: sharedfile regserver
Source: support\expsrv.dll; DestDir: {app}\support; CopyMode: alwaysoverwrite
Source: support\vbajet32.dll; DestDir: {app}\support; CopyMode: alwaysoverwrite
Source: support\msjint35.dll; DestDir: {app}\support; CopyMode: alwaysoverwrite
Source: support\MSJTER35.DLL; DestDir: {app}\support; CopyMode: alwaysoverwrite
Source: support\MSJET35.DLL;  DestDir: {app}\support; CopyMode: alwaysoverwrite; Flags: sharedfile regserver
Source: support\DAO350.DLL;  DestDir: {app}\support; CopyMode: alwaysoverwrite; Flags: sharedfile regserver