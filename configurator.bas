Attribute VB_Name = "configurator"
Public Property Get tratadato(tipo As String, valor As String, posicion As Integer) As Boolean
' *******************************************************************
' Daniel Campos Fernández 25-07-2001
' Esta función recibe un dato a introducir en la variable "articulo"
' que contiene los datos a introducir en la base de datos.
' necesita :
' Tipo : según el protocolo de comunicaciones (PLU,COD, etc)
' valor : valor a introducir
' posición : se utiliza en los datos TXT y TLT para indicar el número
' de línea de texto (TXT0, TXT1...TXT20) y el número de tipo de letra (TLT0..TLT10)
' Si el valor a introducir es correcto, devuelve TRUE, en caso contrario , FALSE
    Dim devuelve As Boolean
    Dim bucle As Integer
    devuelve = False
    For bucle = 0 To DimensionArticulo - 1
        If Propiedades(bucle).codigo = tipo Then
            If (Propiedades(bucle).max >= Len(valor)) Or Propiedades(bucle).max = 0 Then
                devuelve = True
                Select Case tipo
                  Case "COD"
                        arTIculo.cod = valor
                  Case "PLU"
                        arTIculo.plu = valor
                  Case "FAM"
                        arTIculo.fam = valor
                  Case "SUB"
                        arTIculo.sub = valor
                  Case "SEC"
                        arTIculo.sec = valor
                  Case "PRC"
                        arTIculo.prc = Val(valor)
                  Case "WGH"
                        arTIculo.wgh = valor
                  Case "CAD"
                        arTIculo.cad = valor
                  Case "PRE"
                        arTIculo.pre = valor
                  Case "TAR"
                        arTIculo.tar = valor
                  Case "ETQ"
                        arTIculo.etq = valor
                  Case "CBA"
                        arTIculo.cba = valor
                  Case "TLT"
                        arTIculo.tlt(posicion) = valor
                  Case "TXT"
                        arTIculo.txt(posicion) = valor
                  Case "TYP"
                        arTIculo.typ = valor
                  
                End Select
                        
            End If
        End If
    Next bucle
    tratadato = devuelve
                
End Property



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
    
    Propiedades(2).codigo = "SEC"
    Propiedades(2).descripcion = "Sección"
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
    
    
    
       
    
End Sub
