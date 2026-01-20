Attribute VB_Name = "configurator"
Public Function tratadato(tipo As String, valor As String, posicion As Integer) As Integer
' *******************************************************************
' 25-07-2001
' Esta función recibe un dato a introducir en la variable "articulo"
' que contiene los datos a introducir en la base de datos.
' necesita :
' Tipo : según el protocolo de comunicaciones (PLU,COD, etc)
' valor : valor a introducir
' posición : se utiliza en los datos TXT y TLT para indicar el número
' de línea de texto (TXT0, TXT1...TXT20) y el número de tipo de letra (TLT0..TLT10)
' Si el valor a introducir es correcto, devuelve TRUE, en caso contrario , FALSE
    Dim devuelve As Integer
    Dim bucle As Integer
    devuelve = 0
   ' For bucle = 0 To DimensionArticulo - 1
   '     If Propiedades(bucle).codigo = tipo Then
            'If (Propiedades(bucle).max >= Len(valor)) Or Propiedades(bucle).max = 0 Then
                On Error GoTo FinErr
                Select Case tipo
                  Case "COD"
                        articulo.cod = valor
                        articulo.m_cod = True
                        devuelve = 1
                  Case "PLU"
                        articulo.plu = valor
                        articulo.m_plu = True
                        devuelve = 1
                  Case "FAM"
                        articulo.fam = valor
                        articulo.m_fam = True
                        devuelve = 1
                  Case "SUB"
                        articulo.sub = valor
                        articulo.m_sub = True
                        devuelve = 1
                  Case "MOS"
                        articulo.sec = valor
                        articulo.m_sec = True
                        devuelve = 1
                  Case "PRC"
                        articulo.prc = Val(valor)
                        articulo.m_prc = True
                        devuelve = 1
                  Case "WGH"
                        articulo.wgh = valor
                        articulo.m_wgh = True
                        devuelve = 1
                  Case "CAD"
                        articulo.cad = valor
                        articulo.m_cad = True
                        devuelve = 1
                  Case "PRE"
                        articulo.pre = valor
                        articulo.m_pre = True
                        devuelve = 1
                  Case "TAR"
                        articulo.tar = valor
                        articulo.m_tar = True
                        devuelve = 1
                  Case "ETQ"
                        articulo.etq = valor
                        articulo.m_etq = True
                        devuelve = 1
                  Case "CBA"
                        articulo.cba = valor
                        articulo.m_cba = True
                        devuelve = 1
                        If (Val(Mid(valor, 1, 1)) = 3 Or Val(Mid(valor, 1, 1)) = 4) And Val(Mid(valor, 8, 5)) = 0 Then
                             If articulo.wgh = "0" And Val(Mid(valor, 1, 1)) = 4 Then
                                 valor = "4" & "CCCCCC" & "WWWWW"
                                 articulo.cba = valor
                             End If
                             If articulo.wgh = "1" And Val(Mid(valor, 1, 1)) = 3 Then
                                 valor = "3" & "CCCCCC" & "IIIII"
                                 articulo.cba = valor
                             End If
                             
                        End If
                  Case "TLT"
                        articulo.tlt(posicion) = valor
                        articulo.m_tlt(posicion) = True
                        devuelve = 1
                  Case "TXT"
                        articulo.txt(posicion) = valor
                        articulo.m_txt(posicion) = True
                        devuelve = 1
                  Case "TYP"
                        articulo.typ = valor
                        articulo.m_typ = True
                        devuelve = 1
                  Case "GPC"
                        articulo.gpc = Val(valor)
                        articulo.m_gpc = True
                        devuelve = 1
                  Case "TRM1"
                        articulo.trm1 = valor
                        articulo.m_trm1 = True
                        devuelve = 1
                  Case "TRM2"
                        articulo.trm2 = valor
                        articulo.m_trm2 = True
                        devuelve = 1
                  Case "PRC1"
                        articulo.prc1 = valor
                        articulo.m_prc1 = True
                        devuelve = 1
                  Case "PRC2"
                        articulo.prc2 = valor
                        articulo.m_prc2 = True
                        devuelve = 1
                  Case "IVA"
                        articulo.iva = valor
                        articulo.m_iva = True
                        devuelve = 1
                  Case "TEC"
                        articulo.tecla = valor
                        articulo.m_tecla = True
                        devuelve = 1
                  Case "JTEC"
                        articulo.JuegoTecla = valor
                        articulo.m_JuegoTecla = True
                        devuelve = 1
                  Case "MER"
                        articulo.Merma = valor
                        articulo.m_merma = True
                        devuelve = 1
                  Case "NPR"
                        articulo.Presec = valor
                        articulo.m_presec = True
                        devuelve = 1
                  Case "PRC100"
                        If Val(valor) = 0 Then
                            articulo.prc100g = False
                        Else
                            articulo.prc100g = True
                        End If
                        articulo.m_prc100g = True
                        devuelve = 1
                  Case "BAJA"
                        articulo.baja = valor
                        articulo.m_baja = True
                        devuelve = 1
                
                End Select
            'End If
    '    End If
    'Next bucle
    On Error GoTo 0
    tratadato = devuelve
    Exit Function
FinErr:
    tratadato = -1
End Function



Public Sub inicia()
' *******************************************************************
' Daniel Campos Fernández 25-07-2001
' Esta función inicializa los valores de la variable "Propiedades"
' que contiene los datos sobre artículos :
' .codigo = codigo de la propiedad según el protocolo de comunicaciones
' .descripción = una descripción informativa
' .max = longitud máxima admisible del campo
' *******************************************************************
    Propiedades(0).codigo = "COD"
    Propiedades(0).descripcion = "Código"
    Propiedades(0).max = 6
    
    Propiedades(1).codigo = "PLU"
    Propiedades(1).descripcion = "Plu"
    Propiedades(1).max = 4
    
    Propiedades(2).codigo = "MOS"
    Propiedades(2).descripcion = "Mostrador"
    Propiedades(2).max = 2
    
    Propiedades(3).codigo = "SUB"
    Propiedades(3).descripcion = "Subsección"
    Propiedades(3).max = 3
    
    Propiedades(4).codigo = "FAM"
    Propiedades(4).descripcion = "Familia"
    Propiedades(4).max = 4
    
    Propiedades(5).codigo = "PRC"
    Propiedades(5).descripcion = "Precio"
    Propiedades(5).max = 7
    
    Propiedades(6).codigo = "WGH"
    Propiedades(6).descripcion = "Tipo de venta (pesado/no pesado)"
    Propiedades(6).max = 1
    
    Propiedades(7).codigo = "CAD"
    Propiedades(7).descripcion = "Caducidad"
    Propiedades(7).max = 3
    
    Propiedades(8).codigo = "PRE"
    Propiedades(8).descripcion = "Consumo preferente"
    
    Propiedades(9).codigo = "TAR"
    Propiedades(9).descripcion = "Tara Asociada"
    Propiedades(9).max = 5
    
    Propiedades(10).codigo = "ETQ"
    Propiedades(10).descripcion = "Número de etiqueta"
    Propiedades(10).max = 2
    
    Propiedades(11).codigo = "CBA"
    Propiedades(11).descripcion = "Código de barras del artículo"
    Propiedades(11).max = 12
    
    Propiedades(12).codigo = "TLT"
    Propiedades(12).descripcion = "Tipo de letra"
    Propiedades(12).max = 1
    
    Propiedades(13).codigo = "TXT"
    Propiedades(13).descripcion = "Texto"
    Propiedades(13).max = 25
    
    Propiedades(14).codigo = "TYP"
    Propiedades(14).descripcion = "Tipo de descriptivo"
    
    Propiedades(15).codigo = "GPC"
    Propiedades(15).descripcion = "Grupo de Conservación"
    Propiedades(15).max = 2
    
    Propiedades(16).codigo = "ETQ"
    Propiedades(16).descripcion = "Tipo de Etiqueta"
    Propiedades(16).max = 2
    
    Propiedades(17).codigo = "TRM1"
    Propiedades(17).descripcion = "Tramo 1"
    Propiedades(17).max = 9
       
    Propiedades(18).codigo = "PRC1"
    Propiedades(18).descripcion = "Precio del Tramo 1"
    Propiedades(18).max = 9
       
    Propiedades(19).codigo = "TRM2"
    Propiedades(19).descripcion = "Tramo 2"
    Propiedades(19).max = 9
    
    Propiedades(20).codigo = "PRC2"
    Propiedades(20).descripcion = "Precio del Tramo 2"
    Propiedades(20).max = 9
    
    Propiedades(21).codigo = "TEC"
    Propiedades(21).descripcion = "Tecla en la sección"
    Propiedades(21).max = 3
    
    Propiedades(22).codigo = "JTEC"
    Propiedades(22).descripcion = "Juego de Tecla en la sección"
    Propiedades(22).max = 2
    
    Propiedades(23).codigo = "NPR"
    Propiedades(23).descripcion = "Nº Paquete Preseleccionado"
    Propiedades(23).max = 2
    
    Propiedades(24).codigo = "MER"
    Propiedades(24).descripcion = "Merma en tanto por ciento"
    Propiedades(24).max = 2
    
    Propiedades(25).codigo = "PRC100"
    Propiedades(25).descripcion = "Flag Precio por 100 gramos (EcoLabel)"
    Propiedades(25).max = 1
    
    Propiedades(26).codigo = "BAJA"
    Propiedades(26).descripcion = "Distinto de 0 indica baja Artículo"
    Propiedades(26).max = 2
    
End Sub
