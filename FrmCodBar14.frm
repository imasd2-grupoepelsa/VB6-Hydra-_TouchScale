VERSION 5.00
Begin VB.Form FrmCodBar14 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Códigos de Barras"
   ClientHeight    =   6300
   ClientLeft      =   945
   ClientTop       =   810
   ClientWidth     =   9420
   LinkTopic       =   "Formcodbar14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   9420
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdPegar 
      BackColor       =   &H80000018&
      Caption         =   "&Pegar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3480
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   5880
      Width           =   1212
   End
   Begin VB.CommandButton CmdCopiar 
      BackColor       =   &H80000018&
      Caption         =   "&Copiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2160
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   5880
      Width           =   1212
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   8280
      TabIndex        =   14
      Top             =   5760
      Width           =   1092
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Borrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   6360
      TabIndex        =   13
      Top             =   5760
      Width           =   1092
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Añadir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   5160
      TabIndex        =   12
      Top             =   5760
      Width           =   1092
   End
   Begin VB.Frame Frame3 
      Height          =   5292
      Index           =   2
      Left            =   6000
      TabIndex        =   33
      Top             =   360
      Width           =   3372
      Begin VB.Label Label4 
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   4200
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "W: un digito de peso"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   28
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Z: un digito de precio."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   27
         Left            =   120
         TabIndex        =   42
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "I: un digito de importe."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   26
         Left            =   120
         TabIndex        =   41
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "T: un digito de tiquet."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   25
         Left            =   120
         TabIndex        =   40
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "P: un digito de numero de plu."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   24
         Left            =   120
         TabIndex        =   39
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "C: un digito del codigo."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   120
         TabIndex        =   38
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "V: un digito del numero del vendedor."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   22
         Left            =   120
         TabIndex        =   37
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Label Label3 
         Caption         =   "N: Equip."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   120
         TabIndex        =   36
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Y: un digito del numero de mostrador o sección."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   120
         TabIndex        =   35
         Top             =   3480
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "J: un digito de moneda nacional."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   120
         TabIndex        =   34
         Top             =   3840
         Width           =   2895
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Códigos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4452
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   5772
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Local"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3372
         Index           =   1
         Left            =   2880
         TabIndex        =   27
         Top             =   600
         Width           =   2772
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   9
            Left            =   480
            TabIndex        =   11
            Top             =   2880
            Width           =   2172
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   8
            Left            =   480
            TabIndex        =   10
            Top             =   2280
            Width           =   2172
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   7
            Left            =   480
            TabIndex        =   9
            Top             =   1680
            Width           =   2172
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   6
            Left            =   480
            TabIndex        =   8
            Top             =   1080
            Width           =   2172
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   5
            Left            =   480
            TabIndex        =   7
            Top             =   480
            Width           =   2172
         End
         Begin VB.Label Label3 
            Caption         =   "Modo Venta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   252
            Index           =   9
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1572
         End
         Begin VB.Label Label3 
            Caption         =   "Modo Super"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   252
            Index           =   8
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   1692
         End
         Begin VB.Label Label3 
            Caption         =   "Modo Mixto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   252
            Index           =   7
            Left            =   120
            TabIndex        =   30
            Top             =   1440
            Width           =   1692
         End
         Begin VB.Label Label3 
            Caption         =   "Envasado1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   252
            Index           =   6
            Left            =   120
            TabIndex        =   29
            Top             =   2040
            Width           =   1812
         End
         Begin VB.Label Label3 
            Caption         =   "Envasado2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   252
            Index           =   5
            Left            =   120
            TabIndex        =   28
            Top             =   2640
            Width           =   1932
         End
      End
      Begin VB.CheckBox ChkSimultaneo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Definir Ambos simultáneamente"
         ForeColor       =   &H80000008&
         Height          =   192
         Left            =   2880
         TabIndex        =   20
         Top             =   240
         Width           =   2652
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3372
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   2652
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   4
            Left            =   480
            TabIndex        =   6
            Top             =   2880
            Width           =   2052
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   3
            Left            =   480
            TabIndex        =   5
            Top             =   2280
            Width           =   2052
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   2
            Left            =   480
            TabIndex        =   4
            Top             =   1680
            Width           =   2052
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   1
            Left            =   480
            TabIndex        =   3
            Top             =   1080
            Width           =   2052
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   0
            Left            =   480
            TabIndex        =   2
            Top             =   480
            Width           =   2052
         End
         Begin VB.Label Label3 
            Caption         =   "Envasado2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   252
            Index           =   4
            Left            =   120
            TabIndex        =   26
            Top             =   2640
            Width           =   2052
         End
         Begin VB.Label Label3 
            Caption         =   "Envasado1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   252
            Index           =   3
            Left            =   120
            TabIndex        =   25
            Top             =   2040
            Width           =   1932
         End
         Begin VB.Label Label3 
            Caption         =   "Modo Mixto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   252
            Index           =   2
            Left            =   120
            TabIndex        =   24
            Top             =   1440
            Width           =   2052
         End
         Begin VB.Label Label3 
            Caption         =   "Modo Super"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   252
            Index           =   1
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Width           =   1932
         End
         Begin VB.Label Label3 
            Caption         =   "Modo Venta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1812
         End
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   972
      Left            =   1680
      TabIndex        =   15
      Top             =   120
      Width           =   4212
      Begin VB.ComboBox CmbNumero 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   2292
      End
      Begin VB.ComboBox CmbTipo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FrmCodBar14.frx":0000
         Left            =   2880
         List            =   "FrmCodBar14.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1212
      End
      Begin VB.Label LblNumero 
         Caption         =   "Mostrador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1452
      End
      Begin VB.Label Label2 
         Caption         =   "Mostrador o Equipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   2412
      End
   End
   Begin VB.Image Image1 
      Height          =   492
      Left            =   120
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   4692
   End
   Begin VB.Label LblInfo2 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   6000
      TabIndex        =   44
      Top             =   120
      Width           =   3372
   End
   Begin VB.Label LblInfo 
      BackColor       =   &H8000000E&
      Caption         =   "Nuevo Dato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   852
      Left            =   120
      TabIndex        =   21
      Top             =   240
      Width           =   1452
   End
End
Attribute VB_Name = "FrmCodBar14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Portapapeles(10) As String
Private Type TipoCodigos
    codigo() As Integer
    cantidad As Integer
End Type
Private Sub CambiarIdioma()
    FrmCodBar14.Caption = CargaCadena(725) & " (Ean14)"
    Frame2.Caption = CargaCadena(69)
    Label2.Caption = CargaCadena(720)
    Label3(0).Caption = CargaCadena(721)
    Label3(1).Caption = CargaCadena(722)
    Label3(2).Caption = CargaCadena(723)
    Label3(3).Caption = CargaCadena(724) & " 1"
    Label3(4).Caption = CargaCadena(724) & " 2"
    Label3(5).Caption = CargaCadena(724) & " 2"
    Label3(6).Caption = CargaCadena(724) & " 1"
    Label3(7).Caption = CargaCadena(723)
    Label3(8).Caption = CargaCadena(722)
    Label3(9).Caption = CargaCadena(721)
    LblNumero.Caption = CargaCadena(302)
    LblInfo.Caption = CargaCadena(524)
    CmdCopiar.Caption = CargaCadena(717)
    CmdPegar.Caption = CargaCadena(718)
    CmdAceptar.Caption = CargaCadena(550)
    CmdBorrar.Caption = CargaCadena(38)
    CmdSalir.Caption = CargaCadena(7)

    CmbTipo.AddItem CargaCadena(302)
    CmbTipo.AddItem CargaCadena(303)
    ChkSimultaneo.Caption = CargaCadena(726)
    Frame3(0).Caption = CargaCadena(332)
    ''''''''''''''
    Frame3(2).Caption = CargaCadena(334)


    'touchscale Barna
    'Z precio
    'M País
    'W   Peso o unidades(excepto modo Venta/Mixto)
    'F Fabricante
    '(I,i) Importe
    'B   Número de balanza
    'T   Número de tiquet
    'Y   Mostrador
    'P   Número de Plu
    '0..9    Dígitos fijos
    'C   Código de Plu
    'D   Peso /10  (Imprime peso sin último decimal)
    'V   Vendedor
    'K   Check-digit intermedio (posición 7)
    'O   Número de orden
    'J   Importe en moneda del país.(Independiente de la fase Euro)
    'A   Número de línea
    'X   Código de cliente
    'L   Número de lote
    'U   Unidades
    'G   Peso neto antes de merma(solo envasado)
    'Q   Importe sin IVA
    ''''''''''''''''''''''''''''''''''''''''
    '1385    W: un dígito de peso
    '1386    Z: un dígito de precio.
    '1387    I: un dígito de importe Total
    '1388    i: un dígito de importe operación
    '1389    T: un dígito de tique
    '1390    P: un dígito de numero de plu
    '1391    C: un dígito del codigo artículo
    '1392    V: un dígito del numero del vendedor
    '1393    N: un dígito número de equipo
    '1394    Y: un dígito del numero de mostrador o sección.
    '1395    J: un digito de Importe Total en moneda nacional.
    '1396    j: un digito de Importe operación en moneda nacional.
    ''''''''''''''''''''''''''''''''''''''''
    FrmCodBar14.Label3(28).Caption = CargaCadena(1385)
    FrmCodBar14.Label3(27).Caption = CargaCadena(1386)
    FrmCodBar14.Label3(26).Caption = CargaCadena(1387)

    FrmCodBar14.Label1.Caption = CargaCadena(1388)

    FrmCodBar14.Label3(25).Caption = CargaCadena(1389)
    FrmCodBar14.Label3(24).Caption = CargaCadena(1390)
    FrmCodBar14.Label3(23).Caption = CargaCadena(1391)
    FrmCodBar14.Label3(22).Caption = CargaCadena(1392)
    'FrmCodBar14.Label3(24).Caption = CargaCadena(1393)
    FrmCodBar14.Label3(15).Caption = CargaCadena(1394)
    FrmCodBar14.Label3(17).Caption = CargaCadena(1395)
    FrmCodBar14.Label3(14).Caption = CargaCadena(1393)
    FrmCodBar14.Label4.Caption = CargaCadena(1396)
    ''''''''''''''''''

End Sub
Private Function Desglosa_Codigos(LaCadena As String) As TipoCodigos
    Dim LosCodigos As TipoCodigos
    Dim Buffer As String
    Dim bucle As Long
    Dim cantidad As Integer
    cantidad = 0
    Buffer = ""
    For bucle = 1 To Len(LaCadena)
        If Mid(LaCadena, bucle, 1) = " " Or bucle = Len(LaCadena) Then
            If bucle = Len(LaCadena) Then Buffer = Buffer & Mid(LaCadena, bucle, 1)
            bucle = bucle + 2
            cantidad = cantidad + 1
            ReDim Preserve LosCodigos.codigo(cantidad + 1)
            LosCodigos.codigo(cantidad) = Buffer
            LosCodigos.cantidad = cantidad
            Buffer = ""
        Else
            Buffer = Buffer & Mid(LaCadena, bucle, 1)
        End If
    Next bucle
    Desglosa_Codigos = LosCodigos
End Function


Private Sub Refresca_Datos()

    Dim Registro As New RecordNet
    Dim bucle As Integer
    Dim posicion As Integer
    Dim eNCONTRADO As Boolean
    Dim ElCodigo As String
    For bucle = 0 To 9
        Text1(bucle).TexT = ""
        Text1(bucle).BackColor = vbWhite
    Next bucle
    LblInfo.BackColor = vbWhite
    LblInfo.Caption = CargaCadena(524)  '"Nuevo Dato"
    CmdAceptar.Caption = CargaCadena(550)  '"&Añadir"
    If (CmbNumero.TexT = "") And (CmbTipo.ListIndex < 2) Then Exit Sub
    eNCONTRADO = False

    Select Case CmbTipo.ListIndex
    Case 0
        bucle = 1
        ElCodigo = ""
        Do Until Mid(CmbNumero.TexT, bucle, 1) = " " Or bucle > Len(CmbNumero.TexT)
            ElCodigo = ElCodigo & Mid(CmbNumero.TexT, bucle, 1)
            bucle = bucle + 1
        Loop
        Registro.OpenRecordset _
                ("select * from gen_sam where baja=false and tipo='cb14' and destino=1 and numero=" & ElCodigo)

    Case 1
        Registro.OpenRecordset _
                ("select * from gen_sam where baja=false and tipo='cb14' and destino=2 and numero=" & CmbNumero.TexT)

    End Select
    With Registro
        If Not .EOF Then
            eNCONTRADO = True
            Text1(0).TexT = .Fields("txt01")    '("cb_vnt")
            Text1(1).TexT = .Fields("txt02")    '.Fields("cb_sup")
            Text1(2).TexT = .Fields("txt03")    '.Fields("cb_mix")
            Text1(3).TexT = .Fields("txt04")    '.Fields("cb_ev1")
            Text1(4).TexT = .Fields("txt05")    '.Fields("cb_ev2")
            Text1(5).TexT = .Fields("txt06")    '("cb_vnt")
            Text1(6).TexT = .Fields("txt07")    '.Fields("cb_sup")
            Text1(7).TexT = .Fields("txt08")    '.Fields("cb_mix")
            Text1(8).TexT = .Fields("txt09")    '.Fields("cb_ev1")
            Text1(9).TexT = .Fields("txt10")    '.Fields("cb_ev2")
        End If
    End With


    If eNCONTRADO Then
        LblInfo.BackColor = vbYellow
        LblInfo.Caption = CargaCadena(525)  '"Modificación"
        CmdAceptar.Caption = CargaCadena(36)  '"Modific&ar"
    End If
End Sub
Private Sub ChkSimultaneo_Click()
    Dim bucle As Integer
    If ChkSimultaneo.Value = vbChecked Then
        For bucle = 0 To 4
            Text1(bucle + 5).TexT = Text1(bucle).TexT
            Text1(bucle + 5).Enabled = False
        Next bucle
    Else
        For bucle = 5 To 9
            Text1(bucle).Enabled = True
        Next bucle
    End If
End Sub

Private Sub CmbNumero_Click()
    Ajusta_Visible
    Refresca_Datos
End Sub

Private Sub CmbTipo_Click()
    Dim bucle As Integer
    ChkSimultaneo.Visible = True
    For bucle = 1 To 9
        Text1(bucle).Visible = True
        Label3(bucle).Visible = True
    Next bucle
    If CmbTipo.ListIndex = 0 Then
        Aux_Refresca_Mostradores CmbNumero
        LblNumero.Caption = CargaCadena(302)  '"Mostrador"
    Else
        Aux_Refresca_equipos CmbNumero
        LblNumero.Caption = CargaCadena(303)  '"Equipo"
    End If
    Ajusta_Visible
    Refresca_Datos
End Sub
Private Sub Ajusta_Visible()
    Dim bucle As Integer
    Dim Deshabilitar As Boolean
    Dim SoyV8 As Boolean

    Dim MyNum As TipoCodigos
    Dim Registro As New RecordNet
    CmbNumero.Enabled = True
    If (CmbTipo.ListIndex = 1) And IsNumeric(CmbNumero.TexT) Then

        Registro.OpenRecordset ("select * from gen_sam where baja=false and tipo='cb14' and destino=2 and numero=" & CmbNumero.TexT)

    End If
    If CmbTipo.ListIndex = 0 Then
        If Trim(CmbNumero.TexT) <> "" Then
            MyNum = Desglosa_Codigos(CmbNumero.TexT)
            Registro.OpenRecordset ("select * from gen_sam where baja=false and tipo='cb14' and destino=1 and numero=" & MyNum.codigo(1))
        End If
    End If
    For bucle = 1 To 9
        Text1(bucle).Visible = True
        Label3(bucle).Visible = True
    Next bucle
    Label3(0).Caption = CargaCadena(721)
    Label3(3).Caption = CargaCadena(724) & " 1"
    Label3(4).Caption = CargaCadena(724) & " 2"
    Text1(3).Visible = True
    Label3(3).Visible = True
    Text1(4).Visible = True
    Label3(4).Visible = True
End Sub
Private Sub Cmdaceptar_Click()

    Dim Registro As New RecordNet
    Dim bucle As Integer
    Dim ElCodigo As TipoCodigos
    Dim MiCodigo As DB_CodigoBarras
    Dim nR As Integer

    '******************
    ' Condiciones de Error
    If CmbTipo.ListIndex < 2 Then
        If CmbNumero.TexT = "" Then
            If CmbTipo.TexT = CmbTipo.List(0) Then
                MsgBox CargaCadena(559)  '"Introduzca el número de mostrador"
            Else
                MsgBox CargaCadena(575)  '"Introduzca el número de equipo"
            End If
            Exit Sub
        End If
    End If
    For bucle = 0 To 9
        If Len(Text1(bucle).TexT) <> 13 And Len(Text1(bucle).TexT) <> 0 Then
            Text1(bucle).BackColor = vbYellow
            MsgBox "Hay un campo con menos de 13 caracteres"
            Text1(bucle).SetFocus
            Exit Sub
        End If
    Next bucle
    '********************

    Select Case CmbTipo.ListIndex
    Case 0
        ElCodigo = Desglosa_Codigos(CmbNumero.TexT)
        MiCodigo.Mostrador = ElCodigo.codigo(1)
        MiCodigo.Equipo = 0
    Case 1
        ElCodigo = Desglosa_Codigos(CmbNumero.TexT)
        Registro.OpenRecordset ("select codi_ident from equipos where borrado=false and numero_eqp=" & _
                                CmbNumero.TexT)
        MiCodigo.Mostrador = Registro.Fields("codi_ident")
        MiCodigo.Equipo = CmbNumero.TexT
    End Select
    For bucle = 0 To 9
        MiCodigo.Datos(bucle) = Text1(bucle).TexT
    Next bucle
    nR = Alta_Barras14(MiCodigo)
    Select Case nR
    Case 0
        LblInfo2.Caption = CargaCadena(527)  '"Datos Añadidos. Envíe Modificaciones"
    Case 1
        LblInfo2.Caption = CargaCadena(439)  '"Datos Modificados. Envíe Modificaciones"
    End Select
End Sub

Private Sub CmdBorrar_Click()

    Dim Registro As New RecordNet
    Dim Reg2 As New RecordNet
    Dim bucle As Integer
    Dim BucleGrande As Integer
    Dim ElCodigo As TipoCodigos
    Dim MiCodigo As DB_CodigoBarras
    If CmbNumero.TexT = "" Then Exit Sub
    bucle = 1
    ElCodigo = Desglosa_Codigos(CmbNumero.TexT)
    If CmbTipo.TexT = CmbTipo.List(0) Then
        MiCodigo.Equipo = 0
        MiCodigo.Mostrador = ElCodigo.codigo(1)
    Else

        Registro.OpenRecordset ("select codi_ident from equipos where numero_eqp=" & _
                                CmbNumero.TexT & " and borrado=false")
        MiCodigo.Mostrador = Registro.Fields("codi_ident")
        MiCodigo.Equipo = ElCodigo.codigo(1)

    End If
    If Baja_Barras14(MiCodigo) = 0 Then

        LblInfo2.Caption = CargaCadena(556)  '"Datos Eliminados.Envíe Modificaciones"

    End If
    If CmbTipo.TexT = CmbTipo.List(0) Then
        Aux_Refresca_Mostradores CmbNumero
    Else
        Aux_Refresca_equipos CmbNumero
    End If
    Refresca_Datos

End Sub

Private Sub CmdCopiar_Click()
    Dim bucle As Integer
    For bucle = 0 To 9
        Portapapeles(bucle) = Text1(bucle).TexT
    Next bucle
End Sub

Private Sub cmdpegar_click()
    Dim bucle As Integer
    For bucle = 0 To 9
        Text1(bucle).TexT = Portapapeles(bucle)
    Next bucle
End Sub
Private Sub CmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    CambiarIdioma
    Me.Icon = Form2.Icon
    Image1.Picture = Form2.Imagen.Picture
    CmbTipo.TexT = CmbTipo.List(0)
    Aux_Refresca_Mostradores CmbNumero
    If (NumeroSistemas = 1) And usaSC10 And TipoC9 Then
        Frame3(1).Enabled = False
        ChkSimultaneo.Enabled = False
        CmbTipo.Enabled = False
        Label2.Enabled = False
        Label3(5).Enabled = False
        Label3(6).Enabled = False
        Label3(7).Enabled = False
        Label3(8).Enabled = False
        Label3(9).Enabled = False
    End If
    '
    If Not UsaGamaAlta And ((Not usaSC10) Or (usaSC10 And TipoC9)) Then
        If UsaECO1 Or UsaECO2 Then CmbTipo.ListIndex = 0
        If UsaEcoPrint Then CmbTipo.ListIndex = 1
        If Not ((UsaECO1 Or UsaECO2) And UsaEcoPrint) Then
            CmbTipo.Enabled = False
        End If
    End If
    '
    If CmbTipo.ListIndex = 0 Or CmbTipo.ListIndex = 1 Then
        If CmbNumero.ListCount > 0 Then CmbNumero.ListIndex = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LockBase = False
    frmEpelsa.Enabled = True
    'If SoyMaster Or SoyReserva Then
    '        Me.Enabled = False
    '        MS_EnviaBaseRemota
    '        Me.Enabled = True
    'End If

End Sub

Private Sub Text1_Change(Index As Integer)
    If ChkSimultaneo.Value = vbChecked Then
        If Index < 5 Then
            Text1(Index + 5).TexT = Text1(Index).TexT

        End If
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim Matriz As String
    Dim bucle As Integer
    Dim Conforme As Boolean
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = 4 And Text1(5).Enabled = False Then
            CmdAceptar.SetFocus
            Exit Sub
        End If
        If Index = 9 Then
            CmdAceptar.SetFocus
            Exit Sub
        End If
        Text1(Index + 1).SetFocus
        Exit Sub
    End If
    If KeyAscii <> Asc("i") And KeyAscii <> Asc("j") Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    'If Dir(App.Path & "\digitonuba.txt") <> "" Then
    '    Matriz = "SWZITPCVALFMXDNYQJUKOGR"
    'Else
    'Matriz = "SWZITPCVALFMXDNYQJUKOG"
    'Matriz = "ZWIiTPCVOALGMFBYDKJjXUQ"
    Matriz = "ZWIiTPCVYJjN"
    'End If
    'Text1(Index).Locked = True
    Text1(Index).Locked = False
    If KeyAscii = 8 Then Exit Sub

    If Len(Text1(Index).TexT) >= 13 Then
        Text1(Index).Locked = True
        Exit Sub
    End If
    Conforme = False
    If IsNumeric(Chr(KeyAscii)) Then
        Conforme = True
    Else
        'If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32

        For bucle = 1 To Len(Matriz)
            If Mid(Matriz, bucle, 1) = Chr(KeyAscii) Then
                Conforme = True
            End If
        Next bucle
    End If
    If Not Conforme Then
        Text1(Index).Locked = True
    End If


End Sub

Private Sub Text1_LostFocus(Index As Integer)
    If Len(Text1(Index).TexT) <> 13 And Len(Text1(Index).TexT) <> 0 Then
        Text1(Index).BackColor = vbYellow
    Else
        Text1(Index).BackColor = vbWhite
    End If
End Sub

Private Function Alta_Barras14(MiCabecera As DB_CodigoBarras) As Integer
    Dim myRs As dao.Recordset
    Dim myB As dao.Database
    Dim Resp As Integer
    Dim Dest As Integer
    Dim numero As Integer
    Set myB = AbrirBase
    If MiCabecera.Equipo = 0 Then
        Set myRs = myB.OpenRecordset("select * from gen_sam where tipo='cb14' and destino=1 and numero=" & CStr(MiCabecera.Mostrador) & " and baja=false")
        Dest = 1
        numero = MiCabecera.Mostrador
    Else
        Set myRs = myB.OpenRecordset("select * from gen_sam where tipo='cb14' and destino=2 and numero=" & CStr(MiCabecera.Equipo) & " and baja=false")
        Dest = 2
        numero = MiCabecera.Equipo
    End If
    If myRs.EOF Then
        myRs.AddNew
        Resp = 0
    Else
        myRs.Edit
        Resp = 1
    End If
    myRs.Fields("tipo") = "cb14"
    myRs.Fields("destino") = Dest
    myRs.Fields("numero") = numero
    myRs.Fields("redlocal") = False
    myRs.Fields("baja") = False
    myRs.Fields("txt01") = Trim(Text1(0))
    myRs.Fields("txt02") = Trim(Text1(1))
    myRs.Fields("txt03") = Trim(Text1(2))
    myRs.Fields("txt04") = Trim(Text1(3))
    myRs.Fields("txt05") = Trim(Text1(4))
    myRs.Fields("txt06") = Trim(Text1(5))
    myRs.Fields("txt07") = Trim(Text1(6))
    myRs.Fields("txt08") = Trim(Text1(7))
    myRs.Fields("txt09") = Trim(Text1(8))
    myRs.Fields("txt10") = Trim(Text1(9))
    myRs.Fields("tran_txt") = " "
    myRs.Update
    myRs.Close
    Set myRs = Nothing
    CerrarBase myB
    Set myB = Nothing

    Alta_Barras14 = Resp
End Function

'Baja_Cabeceras
Private Function Baja_Barras14(MiCabecera As DB_CodigoBarras) As Integer
    Dim myRs As dao.Recordset
    Dim myB As dao.Database
    Dim Resp As Integer
    Dim Dest As Integer
    Dim numero As Integer
    Dim cm As String
    Set myB = AbrirBase
    If MiCabecera.Equipo = 0 Then
        Set myRs = myB.OpenRecordset("select * from gen_sam where tipo='cb14' and destino=1 and numero=" & CStr(MiCabecera.Mostrador) & " and baja=false")
        Dest = 1
        numero = MiCabecera.Mostrador
    Else
        Set myRs = myB.OpenRecordset("select * from gen_sam where tipo='cb14' and destino=2 and numero=" & CStr(MiCabecera.Equipo) & " and baja=false")
        Dest = 2
        numero = MiCabecera.Equipo
    End If
    If Not myRs.EOF Then
        myRs.Edit
        myRs.Fields("baja") = True
        myRs.Fields("tran_txt") = " "
        myRs.Update
    End If

    Baja_Barras14 = Resp
End Function

