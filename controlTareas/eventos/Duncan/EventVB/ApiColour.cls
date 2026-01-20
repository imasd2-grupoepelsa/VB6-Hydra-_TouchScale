VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApiColour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 2701
Attribute VB_Description = "This class is used to split convert colours between the (Red,Green,Blue) representation and the windows COLORREF representation."
Option Explicit

' ##MODULE_DESCRIPTION This class is used to split convert colours between the _
(Red,Green,Blue) representation and the windows COLORREF representation.


'\\ --[ApiColour]--------------------------------------------------
'\\ Utility class for making RGBQUADS from their respective colour
'\\ components
'\\ ---------------------------------------------------------------
'\\ (c) 2001 Merrion Computing Ltd

Private mColourRef As Long

Private bChanged As Boolean
Friend Property Get Changed() As Boolean
    Changed = bChanged
End Property
Friend Property Let Changed(ByVal newChanged As Boolean)
    bChanged = newChanged
End Property


Public Property Let Blue(ByVal newval As Byte)
Attribute Blue.VB_HelpID = 2708


    mColourRef = MakeLong(MakeInt(Red, Green), MakeInt(newval, 0))


End Property

Friend Property Let ColourRef(ByVal newref As Long)

    mColourRef = newref
    
End Property

Friend Property Get ColourRef() As Long

    ColourRef = mColourRef
    
End Property


Public Property Get Description() As String
Attribute Description.VB_HelpID = 2703

    Dim sdesc As String
    
    sdesc = "COLOUR:Red=" & Red & ",Green=" & Green & ",Blue=" & Blue
    
    Description = sdesc
    
End Property

Public Property Let Green(ByVal newval As Byte)
Attribute Green.VB_HelpID = 2707

    mColourRef = MakeLong(MakeInt(Red, newval), MakeInt(Blue, 0))
    
End Property


Public Property Let Red(ByVal newval As Byte)
Attribute Red.VB_HelpID = 2706

    mColourRef = MakeLong(MakeInt(newval, Green), MakeInt(Blue, 0))

End Property


Public Property Get Red() As Byte

    Red = LoByte(LoWord(mColourRef))
    
End Property

Public Property Get Green() As Byte

    Green = HiByte(LoWord(mColourRef))
    
End Property

Public Property Get Blue() As Byte

    Blue = LoByte(HiWord(mColourRef))
    
End Property


