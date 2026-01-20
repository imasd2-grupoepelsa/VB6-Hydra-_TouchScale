VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Attribute VB_HelpID = 2601
Private mRootName As String
Private lpVolumeNameBuffer As String
Private nVolumeNameSize As Long
Private lpVolumeSerialNumber As Long
Private lpMaximumComponentLength As Long
Private lpFileSystemFlags As Long
Private lpFileSystemNameBuffer As String
Private nFileSystemNameSize As Long

'\\ API call declarations...
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" _
(ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, _
ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, _
lpMaximumComponentLength As Long, lpFileSystemFlags As Long, _
ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Private Declare Function SetVolumeLabel Lib "kernel32" _
              Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, _
                                       ByVal lpVolumeName As String) As Long


Public Enum FileSystemFlags
     FILE_CASE_SENSITIVE_SEARCH = &H1
     FILE_CASE_PRESERVED_NAMES = &H2
     FILE_UNICODE_ON_DISK = &H4
     FILE_PERSISTENT_ACLS = &H8
     FILE_FILE_COMPRESSION = &H10
     FILE_VOLUME_QUOTAS = &H20
     FILE_SUPPORTS_SPARSE_FILES = &H40
     FILE_SUPPORTS_REPARSE_POINTS = &H80
     FILE_SUPPORTS_REMOTE_STORAGE = &H100
     FILE_VOLUME_IS_COMPRESSED = &H8000
     FILE_SUPPORTS_OBJECT_IDS = &H10000
     FILE_SUPPORTS_ENCRYPTION = &H20000
End Enum

Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property

Public Property Get CasePreserved() As Boolean
Attribute CasePreserved.VB_HelpID = 2602

Call RefreshVolumeInfo
CasePreserved = (lpFileSystemFlags And FILE_CASE_PRESERVED_NAMES)
    
End Property

Public Property Get CaseSensitive() As Boolean
Attribute CaseSensitive.VB_HelpID = 2603

Call RefreshVolumeInfo
CaseSensitive = (lpFileSystemFlags And FILE_CASE_SENSITIVE_SEARCH)

End Property


Public Property Get Compressed() As Boolean
Attribute Compressed.VB_HelpID = 2604

Call RefreshVolumeInfo
Compressed = (lpFileSystemFlags And FILE_VOLUME_IS_COMPRESSED)

End Property

Public Property Get Encryption() As Boolean
Attribute Encryption.VB_HelpID = 2605

Call RefreshVolumeInfo
Encryption = (lpFileSystemFlags And FILE_SUPPORTS_ENCRYPTION)

End Property

Public Property Get FileCompression() As Boolean
Attribute FileCompression.VB_HelpID = 2606

Call RefreshVolumeInfo
FileCompression = (lpFileSystemFlags And FILE_FILE_COMPRESSION)

End Property

Public Property Get FileSystemName() As String
Attribute FileSystemName.VB_HelpID = 2607

Call RefreshVolumeInfo
    If InStr(lpFileSystemNameBuffer, Chr$(0)) > 0 Then
        FileSystemName = Left$(lpFileSystemNameBuffer, InStr(lpFileSystemNameBuffer, Chr$(0)) - 1)
    End If

End Property

Public Property Get ObjectIds() As Boolean
Attribute ObjectIds.VB_HelpID = 2608

Call RefreshVolumeInfo
ObjectIds = (lpFileSystemFlags And FILE_SUPPORTS_OBJECT_IDS)

End Property

Public Property Get PersistantACLs() As Boolean
Attribute PersistantACLs.VB_HelpID = 2609

Call RefreshVolumeInfo
PersistantACLs = (lpFileSystemFlags And FILE_PERSISTENT_ACLS)

End Property

Private Sub RefreshVolumeInfo()

Dim lret As Long

    '\\ Initialise the buffers
    lpVolumeNameBuffer = String$(1024, 0)
    nVolumeNameSize = Len(lpVolumeNameBuffer)
    lpFileSystemNameBuffer = String$(1024, 0)
    nFileSystemNameSize = Len(lpFileSystemNameBuffer)
    lret = GetVolumeInformation(mRootName, lpVolumeNameBuffer, nVolumeNameSize, lpVolumeSerialNumber, lpMaximumComponentLength, lpFileSystemFlags, lpFileSystemNameBuffer, nFileSystemNameSize)
    If Err.LastDllError <> 0 Then
        ReportError Err.LastDllError, "ApiVolume:Name (Let)", GetLastSystemError
    End If

End Sub

Public Property Get RemoteStorage() As Boolean
Attribute RemoteStorage.VB_HelpID = 2610

Call RefreshVolumeInfo
RemoteStorage = (lpFileSystemFlags And FILE_SUPPORTS_REMOTE_STORAGE)

End Property

Public Property Get ReparsePoints() As Boolean
Attribute ReparsePoints.VB_HelpID = 2611

Call RefreshVolumeInfo
ReparsePoints = (lpFileSystemFlags And FILE_SUPPORTS_REPARSE_POINTS)

End Property

Public Property Get RootName() As String
Attribute RootName.VB_HelpID = 2612

    RootName = mRootName
    
End Property

Public Property Let RootName(ByVal newname As String)

Dim lret As Long

'\\ If the name changes get all the other volume info...
If newname <> mRootName Then
    mRootName = newname
End If

End Property

Public Property Get SerialNumber() As Long
Attribute SerialNumber.VB_HelpID = 2613

Call RefreshVolumeInfo
SerialNumber = lpVolumeSerialNumber
    
End Property

Public Property Get SparseFiles() As Boolean
Attribute SparseFiles.VB_HelpID = 2614

Call RefreshVolumeInfo
SparseFiles = (lpFileSystemFlags And FILE_SUPPORTS_SPARSE_FILES)

End Property

Public Property Get UnicodeFileTable() As Boolean
Attribute UnicodeFileTable.VB_HelpID = 2615

Call RefreshVolumeInfo
UnicodeFileTable = (lpFileSystemFlags And FILE_UNICODE_ON_DISK)

End Property

Public Property Let VolumeName(ByVal newname As String)
Attribute VolumeName.VB_HelpID = 2617

Dim lret As Long

lret = SetVolumeLabel(mRootName, newname)
If Err.LastDllError Then
    ReportError Err.LastDllError, "ApiVolume:Volumename (Let)", GetLastSystemError
End If

End Property

Public Property Get VolumeName() As String

    Call RefreshVolumeInfo
    If InStr(lpVolumeNameBuffer, Chr$(0)) > 0 Then
        VolumeName = Left$(lpVolumeNameBuffer, InStr(lpVolumeNameBuffer, Chr$(0)) - 1)
    End If
    
End Property


Public Property Get VolumeQuotas() As Boolean
Attribute VolumeQuotas.VB_HelpID = 2618

Call RefreshVolumeInfo
VolumeQuotas = (lpFileSystemFlags And FILE_VOLUME_QUOTAS)

End Property


