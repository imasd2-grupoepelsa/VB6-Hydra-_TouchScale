VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApiProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 3138
Attribute VB_Description="This class provides the properties, events and methods used to manipulate the individual window properties any given window."
Option Explicit


' ##MODULE_DESCRIPTION This class provides the properties, events and methods used _
to manipulate the individual window properties any given window.

Private Const CLSNAME = "ApiProperty"

'\\ Member variables
Private mHwnd As Long '\\ Window that has this property
Private mhProperty As Long '\\ Handle to the property value
Private msName As String '\\ Property name

Public Property Get Name() As String
Attribute Name.VB_HelpID = 3139

    Name = msName
    
End Property

Friend Property Let Name(ByVal newname As String)

    msName = newname
    
End Property
