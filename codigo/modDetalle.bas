Attribute VB_Name = "modDetalle"
Option Explicit

' DETALLE_OTROS documenta la columna I: mes, codigo, fecha, tipo, cantidad,
' documento y observacion. "Otro" exige explicacion. Si I>0 exige detalle
' cuya suma coincida con I por codigo y mes. Se compara con I: no se vuelve
' a contabilizar. El importe se calcula a costo promedio O y se concilia con
' J; el residuo de centavos se asigna a la ultima linea de orden estable
' (columna Folio).
' La salida tipo destruccion exige autorizacion o acta (Reg. art. 64 par. 4).
' Capacidad: 50.000 lineas (filas 7 a 50006).

Public Const DET_MES As Long = 1
Public Const DET_CODIGO As Long = 2
Public Const DET_FECHA As Long = 3
Public Const DET_TIPO As Long = 4
Public Const DET_CANTIDAD As Long = 5
Public Const DET_DOCUMENTO As Long = 6
Public Const DET_OBSERVACION As Long = 7
Public Const DET_AUTORIZACION As Long = 8
Public Const DET_FOLIO As Long = 9
Public Const DET_IMPORTE As Long = 10
Public Const DET_BASE As Long = 11

Public Function HojaDetalle() As Worksheet
    Set HojaDetalle = Libro().Sheets(HOJA_DETALLE)
End Function

Public Sub AgregarDetalle()
    Dim ws As Worksheet
    Dim f As Long
    Dim mes As Long
    Dim codigo As String
    Dim tipo As String
    Dim cantidad As Double
    Dim observacion As String
    Dim autorizacion As String
    Dim texto As String
    Dim protegido As Boolean
    Dim numero As Long
    Dim detalle As String
    On Error GoTo fallo
    If Reentrada Then Exit Sub
    Entrar
    mes = MesAbierto()
    If mes = 0 Then
        Salir
        Err.Raise vbObjectError + 540, "AgregarDetalle", _
            "Solo se documentan salidas de otros en un mes abierto."
    End If
    Set ws = HojaDetalle()
    protegido = (ws.ProtectContents Or ws.ProtectDrawingObjects)
    If protegido Then ws.Unprotect Password:=CLAVE_PROTECCION
    f = FILA_DET_1
    Do While Not EsVacio(ws.Cells(f, DET_MES).Value2)
        f = f + 1
    Loop
    If f > FILA_DET_FIN Then
        Salir
        Err.Raise vbObjectError + 545, "AgregarDetalle", _
            "Se alcanzo el limite de " & (FILA_DET_FIN - FILA_DET_1 + 1) & " lineas de detalle."
    End If

    codigo = modUI.CapturarTexto("Codigo del material (debe existir en CATALOGO).", _
        "Detalle de salidas otros")
    If codigo = "" Or FilaEnCatalogo(codigo) = 0 Then
        Salir
        Err.Raise vbObjectError + 546, "AgregarDetalle", _
            "El codigo debe existir en CATALOGO. No se registro nada."
    End If
    tipo = LCase$(modUI.CapturarTexto("Tipo (retiro, autoconsumo, merma, destruccion, otro).", _
        "Tipo de salida", "retiro"))
    If InStr("|" & TIPOS_OTROS & "|", "|" & tipo & "|") = 0 Then
        Salir
        Err.Raise vbObjectError + 547, "AgregarDetalle", _
            "Tipo no permitido. Use uno de: " & Replace(TIPOS_OTROS, "|", ", ") & "."
    End If
    texto = modUI.CapturarTexto("Cantidad de unidades (numerica, mayor o igual a 0).", "Cantidad", "0")
    If Not EsNumeroReal(texto) Then
        Salir
        Err.Raise vbObjectError + 548, "AgregarDetalle", _
            "La cantidad debe ser numerica. No se registro nada."
    End If
    cantidad = CDbl(texto)
    If cantidad < 0 Then
        Salir
        Err.Raise vbObjectError + 548, "AgregarDetalle", _
            "La cantidad no puede ser negativa. No se registro nada."
    End If
    observacion = modUI.CapturarTexto("Observacion (obligatoria si el tipo es otro).", "Observacion", "")
    If tipo = "otro" And observacion = "" Then
        Salir
        Err.Raise vbObjectError + 549, "AgregarDetalle", _
            "El tipo 'otro' exige explicacion en Observacion. No se registro nada."
    End If
    If tipo = "destruccion" Then
        autorizacion = modUI.CapturarTexto( _
            "Destruccion: numero de autorizacion o acta (Reg. art. 64 par. cuarto, " & _
            "D. 2.507, G.O.E. 5.662 del 24-09-2003). Sin ese dato el cierre queda bloqueado.", _
            "Autorizacion o acta", "")
    End If

    ws.Cells(f, DET_MES).Value = mes
    ws.Cells(f, DET_CODIGO).Value = codigo
    ws.Cells(f, DET_FECHA).Value = FechaDelMes(mes)
    ws.Cells(f, DET_FECHA).NumberFormat = "dd/mm/yyyy"
    ws.Cells(f, DET_TIPO).Value = tipo
    ws.Cells(f, DET_CANTIDAD).Value = cantidad
    ws.Cells(f, DET_DOCUMENTO).Value = modUI.CapturarTexto("Documento de soporte.", "Documento", "")
    ws.Cells(f, DET_OBSERVACION).Value = observacion
    If tipo = "destruccion" Then ws.Cells(f, DET_AUTORIZACION).Value = autorizacion
    ws.Cells(f, DET_FOLIO).Value = f
    ws.Cells(f, DET_FOLIO).NumberFormat = FMT_ENTERO
    ws.Cells(f, DET_IMPORTE).Value = 0
    ws.Cells(f, DET_IMPORTE).NumberFormat = FMT_MONEDA
    If protegido Then ws.Protect Password:=CLAVE_PROTECCION, UserInterfaceOnly:=True

    detalle = modDetalle.ValidarDetalleDeMes(mes)
    If detalle <> "" Then
        ws.Unprotect Password:=CLAVE_PROTECCION
        ws.Range(ws.Cells(f, DET_MES), ws.Cells(f, DET_BASE)).ClearContents
        ws.Protect Password:=CLAVE_PROTECCION, UserInterfaceOnly:=True
        Salir
        Err.Raise vbObjectError + 550, "AgregarDetalle", _
            "La linea no concilio con la columna I del mes; se retiro:" & vbCrLf & detalle
    End If
    modBitacora.Registrar "Detalle de salidas otros", codigo, "OK", _
        "Tipo " & tipo & ", cantidad " & Format$(cantidad, FMT_CANT), "", "", HojaMes(mes), ""
    Salir
    Exit Sub
fallo:
    numero = Err.Number
    detalle = Err.Description
    Err.Clear
    On Error Resume Next
    If Not ws Is Nothing Then ws.Protect Password:=CLAVE_PROTECCION, UserInterfaceOnly:=True
    On Error GoTo 0
    Salir
    Err.Raise numero, "modDetalle.AgregarDetalle", detalle
End Sub

Private Function FechaDelMes(ByVal mes As Long) As Date
    FechaDelMes = DateSerial(AnioDelLibro(), mes, 1)
End Function

' Suma de la columna I declarada en el mes, por codigo.
Public Function SumaDeclarada(ByVal indice As Long, ByVal codigo As String) As Double
    Dim hoja As Worksheet
    Dim fila As Long
    Set hoja = Libro().Sheets(HojaMes(indice))
    fila = FilaEnMes(indice, codigo)
    If fila = 0 Then Exit Function
    SumaDeclarada = modCalculo.ValorNumerico(hoja, fila, COL_CANT_OTR)
End Function

Public Function FilaEnMes(ByVal indice As Long, ByVal codigo As String) As Long
    Dim ws As Worksheet
    Dim f As Long
    Dim clave As String
    clave = NormalizaClave(codigo)
    If clave = "" Then Exit Function
    Set ws = Libro().Sheets(HojaMes(indice))
    For f = FILA_1 To FILA_FIN
        If NormalizaClave(ws.Cells(f, COL_CODIGO).Value2) = clave Then
            FilaEnMes = f
            Exit Function
        End If
    Next f
    FilaEnMes = 0
End Function

' Conciliacion: la suma del detalle debe coincidir con la columna I del mes.
Public Function ValidarDetalleDeMes(ByVal indice As Long) As String
    Dim problemas As String
    Dim ws As Worksheet
    Dim f As Long
    Dim clave As String
    Dim tipo As String
    Dim detalle As Double
    Dim codigos As Object
    Dim codigo As String
    Dim mes As Long
    Dim hoja As Worksheet
    Dim f2 As Long
    Dim declarado As Double
    Set ws = HojaDetalle()
    Set codigos = CreateObject("Scripting.Dictionary")

    For f = FILA_DET_1 To FILA_DET_FIN
        If Not EsVacio(ws.Cells(f, DET_MES).Value2) Then
            If Not EsNumeroReal(ws.Cells(f, DET_MES).Value2) Then
                problemas = problemas & "- DETALLE_OTROS fila " & f & ": el mes no es numerico." & vbCrLf
                GoTo siguienteFila
            End If
            mes = CLng(ws.Cells(f, DET_MES).Value2)
            codigo = CStr(ws.Cells(f, DET_CODIGO).Value2)
            clave = NormalizaClave(codigo)
            If clave = "" Then
                problemas = problemas & "- DETALLE_OTROS fila " & f & ": codigo vacio." & vbCrLf
            ElseIf FilaEnCatalogo(codigo) = 0 Then
                problemas = problemas & "- DETALLE_OTROS fila " & f & ": el codigo " & codigo & _
                    " no existe en CATALOGO." & vbCrLf
            End If
            If mes < 1 Or mes > 12 Then
                problemas = problemas & "- DETALLE_OTROS fila " & f & _
                    ": mes fuera de rango." & vbCrLf
            Else
                If mes <> indice And modEstados.EstadoMes(mes) <> EST_CERRADO Then
                    problemas = problemas & "- DETALLE_OTROS fila " & f & _
                        ": el mes indicado (" & mes & ") no corresponde al mes en curso ni esta cerrado." & vbCrLf
                End If
            End If
            If Not EsFecha(ws.Cells(f, DET_FECHA).Value2) Then
                problemas = problemas & "- DETALLE_OTROS fila " & f & ": fecha invalida." & vbCrLf
            ElseIf mes >= 1 And mes <= 12 Then
                If Month(ws.Cells(f, DET_FECHA).Value2) <> mes Or _
                   Year(ws.Cells(f, DET_FECHA).Value2) <> AnioDelLibro() Then
                    problemas = problemas & "- DETALLE_OTROS fila " & f & _
                        ": la fecha no pertenece al mes indicado." & vbCrLf
                End If
            End If
            tipo = LCase$(NormalizaClave(ws.Cells(f, DET_TIPO).Value2))
            If InStr("|" & TIPOS_OTROS & "|", "|" & tipo & "|") = 0 Then
                problemas = problemas & "- DETALLE_OTROS fila " & f & ": tipo '" & _
                    ws.Cells(f, DET_TIPO).Value2 & "' no permitido (" & Replace(TIPOS_OTROS, "|", ", ") & ")." & vbCrLf
            End If
            If tipo = "otro" And EsVacio(ws.Cells(f, DET_OBSERVACION).Value2) Then
                problemas = problemas & "- DETALLE_OTROS fila " & f & _
                    ": el tipo 'otro' exige explicacion en Observacion." & vbCrLf
            End If
            If tipo = "destruccion" And EsVacio(ws.Cells(f, DET_AUTORIZACION).Value2) Then
                problemas = problemas & "- DETALLE_OTROS fila " & f & _
                    ": destruccion sin numero de autorizacion o acta (Reg. art. 64 par. cuarto)." & vbCrLf
            End If
            If Not EsNumeroReal(ws.Cells(f, DET_CANTIDAD).Value2) Or _
               CDbl(ws.Cells(f, DET_CANTIDAD).Value2) < 0 Then
                problemas = problemas & "- DETALLE_OTROS fila " & f & _
                    ": cantidad no numerica o negativa." & vbCrLf
            Else
                detalle = CDbl(ws.Cells(f, DET_CANTIDAD).Value2)
                clave = mes & "|" & NormalizaClave(codigo)
                If codigos.Exists(clave) Then
                    codigos(clave) = codigos(clave) + detalle
                Else
                    codigos.Add clave, detalle
                End If
            End If
        End If
siguienteFila:
    Next f

    ' Cada codigo con I>0 necesita detalle, y el detalle no se contabiliza otra vez.
    Set hoja = Libro().Sheets(HojaMes(indice))
    For f2 = FILA_1 To FILA_FIN
        If Trim$(CStr(hoja.Cells(f2, COL_CODIGO).Value2)) <> "" Then
            declarado = modCalculo.ValorNumerico(hoja, f2, COL_CANT_OTR)
            clave = indice & "|" & NormalizaClave(hoja.Cells(f2, COL_CODIGO).Value2)
            detalle = 0
            If codigos.Exists(clave) Then detalle = codigos(clave)
            If declarado > 0 And detalle <= 0 Then
                problemas = problemas & "- " & hoja.Cells(f2, COL_CODIGO).Value2 & _
                    ": I>0 sin detalle en DETALLE_OTROS." & vbCrLf
            ElseIf declarado = 0 And detalle > 0 Then
                problemas = problemas & "- " & hoja.Cells(f2, COL_CODIGO).Value2 & _
                    ": DETALLE_OTROS registra " & Format$(detalle, FMT_CANT) & _
                    " y la columna I esta en cero." & vbCrLf
            ElseIf declarado > 0 And Abs(Cuantiza4(declarado) - Cuantiza4(detalle)) > 0.00005 Then
                problemas = problemas & "- " & hoja.Cells(f2, COL_CODIGO).Value2 & _
                    ": el detalle (" & Format$(detalle, FMT_CANT) & ") no cuadra con I (" & _
                    Format$(declarado, FMT_CANT) & ")." & vbCrLf
            End If
        End If
    Next f2
    ValidarDetalleDeMes = problemas
End Function

' Reparto del importe de "otros" entre las lineas del detalle, con el residuo
' de centavos en la ultima linea de orden estable (folio).
Public Sub ConciliarImporteOtros(ByVal indice As Long, ByVal codigo As String)
    Dim ws As Worksheet
    Dim hoja As Worksheet
    Dim f As Long
    Dim fUltima As Long
    Dim filaMes As Long
    Dim total As Double
    Dim acumulado As Double
    Dim cp As Double
    Dim importe As Double
    Dim protegido As Boolean
    Set ws = HojaDetalle()
    Set hoja = Libro().Sheets(HojaMes(indice))
    filaMes = FilaEnMes(indice, codigo)
    If filaMes = 0 Then Exit Sub
    cp = modCalculo.ValorNumerico(hoja, filaMes, COL_COSTO_PROM)
    total = modCalculo.ValorNumerico(hoja, filaMes, COL_VAL_OTR)
    For f = FILA_DET_1 To FILA_DET_FIN
        If EsNumeroReal(ws.Cells(f, DET_MES).Value2) Then
            If CLng(ws.Cells(f, DET_MES).Value2) = indice And _
               NormalizaClave(ws.Cells(f, DET_CODIGO).Value2) = NormalizaClave(codigo) Then
                fUltima = f
            End If
        End If
    Next f
    If fUltima = 0 Then Exit Sub
    protegido = (ws.ProtectContents Or ws.ProtectDrawingObjects)
    If protegido Then ws.Unprotect Password:=CLAVE_PROTECCION
    For f = FILA_DET_1 To fUltima
        If EsNumeroReal(ws.Cells(f, DET_MES).Value2) Then
            If CLng(ws.Cells(f, DET_MES).Value2) = indice And _
               NormalizaClave(ws.Cells(f, DET_CODIGO).Value2) = NormalizaClave(codigo) Then
                If f = fUltima Then
                    importe = Redondear2(total - acumulado)
                Else
                    importe = Redondear2(cp * CDbl(ws.Cells(f, DET_CANTIDAD).Value2))
                    acumulado = Redondear2(acumulado + importe)
                End If
                ws.Cells(f, DET_IMPORTE).Value = importe
                ws.Cells(f, DET_IMPORTE).NumberFormat = FMT_MONEDA
                ws.Cells(f, DET_BASE).Value = "costo promedio O"
            End If
        End If
    Next f
    If protegido Then ws.Protect Password:=CLAVE_PROTECCION, UserInterfaceOnly:=True
End Sub

' Reaplica los formatos numericos de DETALLE_OTROS: cantidad con la mascara
' elegida en CONFIGURACION e importe siempre a 2 decimales. Se llama con la
' hoja sin proteger (Workbook_Open / Workbook_BeforeSave).
Public Sub AplicarFormatos()
    Dim ws As Worksheet
    Set ws = Libro().Sheets(HOJA_DETALLE)
    ws.Range(ws.Cells(FILA_DET_1, DET_CANTIDAD), _
             ws.Cells(FILA_DET_FIN, DET_CANTIDAD)).NumberFormat = MascaraCant()
    ws.Range(ws.Cells(FILA_DET_1, DET_IMPORTE), _
             ws.Cells(FILA_DET_FIN, DET_IMPORTE)).NumberFormat = FMT_MONEDA
End Sub