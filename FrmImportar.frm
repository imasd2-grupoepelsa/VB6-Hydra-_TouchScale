VERSION 5.00
Begin VB.Form FrmImportar 
   Caption         =   "Importación"
   ClientHeight    =   1590
   ClientLeft      =   3105
   ClientTop       =   2610
   ClientWidth     =   2730
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   1590
   ScaleWidth      =   2730
End
Attribute VB_Name = "FrmImportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    FrmImportar.ControlImport1.FamFijas = FamiliasFijas
    FrmImportar.ControlImport1.SubFijas = SubSeccionesFijas
    ControlImport1.No_modWGH = False
    ControlImport1.Idioma = id
    If AhorraMas Then
        ControlImport1.Supermercado = "Ahorramas"
    End If
    If UsaEuro Then
        ControlImport1.Euros = True
    Else
        ControlImport1.Euros = False
    End If
    ControlImport1.NumeroDecimales = decimales
    ControlImport1.FactorEuro = 1
End Sub
