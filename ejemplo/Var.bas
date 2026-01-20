Attribute VB_Name = "Var"
Option Explicit

Public MyConnObj As New ADODB.Connection

Public Type Tipo_Linea
    EsPesado As Boolean
    FactorPeso As Integer
    precio As Double
    cantidad As Double
    NumeroLinea As Byte
    CodigoArticulo As Long
    IVA As Integer
    TipoLinea As Byte
    DescuentoLinea As Double
    ImporteLinea As Double
    LineaCancelada As Boolean
    SehaModificado As Boolean
    ModificadoTemporal As Boolean
    FichaVacuno As Long
    ImporteBruto As Double
     descriptivo As String
End Type

Public Type tipo_cabecera
     Ntiquet As Long
     NVendedor As Long
     NMostrador As Long
     NCliente As Long
     Fecha As Variant
     hora As Variant
     NBalanza As Long
     ImporteTotal As Double
     NLineas As Long
     LineaInicial As Long
     LineaFinal As Long
     codigo As Long
     CadenaEstadoTiquet As String
     ImporteBruto As Double
     Impuestos As Double
     descuento As Double
     LineasCanceladas As Long
     tipoTiquet As Long
     EstadoTiquet As Long
     LineasMensaje As Long
     Incluir_en_Base As Boolean
     tipoPago As Long
     cantidadEntre As Double
     Cambio As Double
     lote As String
     lEncargo As Boolean
     documentoabono As Long
     Lines() As Tipo_Linea
End Type

