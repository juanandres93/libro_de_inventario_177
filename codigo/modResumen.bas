Attribute VB_Name = "modResumen"
Option Explicit

' RESUMEN_ANUAL por codigo y unidad: inicial de enero, entradas sumadas,
' ventas y otras salidas sumadas y el final del ULTIMO mes cerrado (no la
' suma de saldos finales). Conciliacion: inicial + entradas - salidas = final.
' Separa lo definitivo (meses cerrados) de lo provisional (mes abierto).

Private Const RES_CODIGO As Long = 1
Private Const RES_DESCRIPCION As Long = 2
Private Const RES_UNIDAD As Long = 3
Private Const RES_INI_CANT As Long = 4
Private Const RES_INI_BS As Long = 5
Private Const RES_ENT_CANT As Long = 6
Private Const RES_ENT_BS As Long = 7
Private Const RES_VTA_CANT As Long = 8
Private Const RES_VTA_BS As Long = 9
Private Const RES_OTR_CANT As Long = 10
Private Const RES_OTR_BS As Long = 11
Private Const RES_FIN_CANT As Long = 12
Private Const RES_FIN_BS As Long = 13
Private Const RES_CHECK As Long = 14
Private Const RES_NOTA As Long = 15

Public Sub ConstruirResumen()
    Dim ws As Worksheet
    Dim fila As Long
    Dim numero As Long
    Dim detalle As String
    Dim paso As String
    On Error GoTo fallo
    paso = "reentrada"
    If Reentrada Then Exit Sub
    Entrar
    modEstadoExcel.CapturarEstado
    modEstadoExcel.SinPantalla
    On Error GoTo fallo
    paso = "abrir hoja"
    Set ws = Libro().Sheets(HOJA_RESUMEN)
    paso = "limpiar"
    ws.Unprotect Password:=CLAVE_PROTECCION
    ws.Cells.ClearContents
    ws.Cells.ClearFormats
    paso = "encabezados"
    EscribirEncabezados ws
    fila = 7
    paso = "bloque definitivo"
    fila = EscribirBloque(ws, fila, False)         ' definitivo: meses cerrados
    fila = fila + 1
    paso = "bloque provisional"
    fila = EscribirBloque(ws, fila, True)          ' provisional: incluye el abierto
    paso = "totales"
    EscribirTotalesResumen ws, fila
    ws.Protect Password:=CLAVE_PROTECCION, UserInterfaceOnly:=True
    modEstadoExcel.RestaurarEstado
    Salir
    Exit Sub
fallo:
    numero = Err.Number
    detalle = Err.Description
    Err.Clear
    On Error Resume Next
    If Not ws Is Nothing Then ws.Protect Password:=CLAVE_PROTECCION, UserInterfaceOnly:=True
    On Error GoTo 0
    modEstadoExcel.RestaurarEstado
    Salir
    Err.Raise numero, "modResumen.ConstruirResumen@" & paso, detalle
End Sub

Private Sub EscribirEncabezados(ByVal ws As Worksheet)
    With ws
        .Range("A1").Value = APP_NOMBRE
        .Range("A2").Value = "Resumen anual por codigo y unidad - " & _
            Libro().Sheets(HOJA_CONFIG).Range(CFG_ENTIDAD).Value2 & " (" & _
            Libro().Sheets(HOJA_CONFIG).Range(CFG_RIF).Value2 & ")"
        .Range("A3").Value = "Metodo: " & APP_METODO & " | Version " & APP_VERSION & _
            " | " & APP_URL
        .Range("A4").Value = "Corte: " & PeriodoActual() & " | Meses cerrados: " & _
            CerradosComoTexto() & " | Mes abierto: " & TextoMesAbierto()
        .Range("A5").Value = "AVISO " & AvisoArt177()
        .Range("A5").WrapText = False
    End With
End Sub

Private Function CerradosComoTexto() As String
    Dim i As Long
    Dim texto As String
    For i = 1 To 12
        If modEstados.EstadoMes(i) = EST_CERRADO Then
            If texto <> "" Then texto = texto & ", "
            texto = texto & i
        End If
    Next i
    If texto = "" Then texto = "ninguno"
    CerradosComoTexto = texto
End Function

' Devuelve la siguiente fila libre.
Private Function EscribirBloque(ByVal ws As Worksheet, ByVal fila As Long, _
                                ByVal provisional As Boolean) As Long
    Dim cat As Long
    Dim ultimaCat As Long
    Dim i As Long
    Dim limite As Long
    Dim ultimoCerrado As Long
    Dim f As Long
    Dim fm As Long
    Dim codigo As String
    Dim hoja As Worksheet
    Dim incC As Double, incB As Double, entC As Double, entB As Double
    Dim vtaC As Double, vtaB As Double, otrC As Double, otrB As Double
    Dim finC As Double, finB As Double
    Dim salida As Long

    ultimoCerrado = modEstados.MesCerradoHasta()
    If provisional Then
        limite = modEstados.MesAbierto()
        If limite = 0 Then limite = ultimoCerrado
    Else
        limite = ultimoCerrado
    End If

    ultimaCat = UltimaFilaUsada(modCatalogo.HojaCatalogo(), 1)
    If ultimaCat < FILA_1 Then
        EscribirBloque = fila + 2
        Exit Function
    End If

    If provisional Then
        ws.Cells(fila, 1).Value = "BLOQUE PROVISIONAL: incluye el mes abierto; no es definitivo"
    Else
        ws.Cells(fila, 1).Value = "BLOQUE DEFINITIVO: solo meses cerrados (1 a " & _
            ultimoCerrado & ")"
    End If
    ws.Cells(fila, 1).Font.Bold = True
    fila = fila + 1

    EncabezadoBloque ws, fila
    fila = fila + 1
    salida = fila

    For cat = FILA_1 To ultimaCat
        codigo = Trim$(CStr(modCatalogo.HojaCatalogo().Cells(cat, COL_CODIGO).Value2))
        f = salida + cat - FILA_1
        If codigo <> "" Then
            incC = 0: incB = 0
            If limite >= 1 Then
                Set hoja = Libro().Sheets(HojaMes(1))
                fm = FilaEnHoja(1, codigo)
                If fm > 0 Then
                    incC = modCalculo.ValorNumerico(hoja, fm, COL_CANT_INI)
                    incB = modCalculo.ValorNumerico(hoja, fm, COL_VAL_INI)
                End If
            End If
            entC = 0: entB = 0: vtaC = 0: vtaB = 0: otrC = 0: otrB = 0
            For i = 1 To limite
                Set hoja = Libro().Sheets(HojaMes(i))
                fm = FilaEnHoja(i, codigo)
                If fm > 0 Then
                    entC = entC + modCalculo.ValorNumerico(hoja, fm, COL_CANT_ENT)
                    entB = entB + modCalculo.ValorNumerico(hoja, fm, COL_VAL_ENT)
                    vtaC = vtaC + modCalculo.ValorNumerico(hoja, fm, COL_CANT_VTA)
                    vtaB = vtaB + modCalculo.ValorNumerico(hoja, fm, COL_VAL_VTA)
                    otrC = otrC + modCalculo.ValorNumerico(hoja, fm, COL_CANT_OTR)
                    otrB = otrB + modCalculo.ValorNumerico(hoja, fm, COL_VAL_OTR)
                End If
            Next i
            finC = 0: finB = 0
            If ultimoCerrado > 0 Then
                finC = modInstantanea.SaldoInstantaneaCantidad(HojaMes(ultimoCerrado), codigo)
                finB = modInstantanea.SaldoInstantanea(HojaMes(ultimoCerrado), codigo)
            End If
            ws.Cells(f, RES_CODIGO).Value = codigo
            ws.Cells(f, RES_DESCRIPCION).Value = _
                CStr(modCatalogo.HojaCatalogo().Cells(cat, COL_DESCRIPCION).Value2)
            ws.Cells(f, RES_UNIDAD).Value = _
                CStr(modCatalogo.HojaCatalogo().Cells(cat, 3).Value2)
            ws.Cells(f, RES_INI_CANT).Value = Cuantiza4(incC)
            ws.Cells(f, RES_INI_BS).Value = Redondear2(incB)
            ws.Cells(f, RES_ENT_CANT).Value = Cuantiza4(entC)
            ws.Cells(f, RES_ENT_BS).Value = Redondear2(entB)
            ws.Cells(f, RES_VTA_CANT).Value = Cuantiza4(vtaC)
            ws.Cells(f, RES_VTA_BS).Value = Redondear2(vtaB)
            ws.Cells(f, RES_OTR_CANT).Value = Cuantiza4(otrC)
            ws.Cells(f, RES_OTR_BS).Value = Redondear2(otrB)
            ws.Cells(f, RES_FIN_CANT).Value = Cuantiza4(finC)
            ws.Cells(f, RES_FIN_BS).Value = Redondear2(finB)
            If ultimoCerrado > 0 Then
                If Abs(Centavos(incB + entB - vtaB - otrB) - Centavos(finB)) <= 0.005 Then
                    ws.Cells(f, RES_CHECK).Value = "OK"
                Else
                    ws.Cells(f, RES_CHECK).Value = "ERROR de conciliacion"
                End If
                If limite < 12 Then
                    ws.Cells(f, RES_NOTA).Value = "Meses sin cerrar: " & (12 - limite)
                End If
            Else
                ws.Cells(f, RES_CHECK).Value = "sin cortes cerrados"
            End If
        End If
    Next cat
    ws.Range(ws.Cells(salida, RES_INI_BS), ws.Cells(salida + ultimaCat - FILA_1, RES_FIN_BS)) _
        .NumberFormat = FMT_MONEDA
    ws.Range(ws.Cells(salida, RES_INI_CANT), ws.Cells(salida + ultimaCat - FILA_1, RES_ENT_CANT)) _
        .NumberFormat = FMT_CANT
    ws.Range(ws.Cells(salida, RES_VTA_CANT), ws.Cells(salida + ultimaCat - FILA_1, RES_FIN_CANT)) _
        .NumberFormat = FMT_CANT

    EscribirBloque = salida + ultimaCat - FILA_1 + 1
End Function

' Reaplica los formatos numericos del resumen: importes en Bs. siempre a 2
' decimales y cantidades con la mascara elegida en CONFIGURACION. Se llama con
' la hoja sin proteger (Workbook_Open / Workbook_BeforeSave desprotegen antes).
Public Sub AplicarFormatos()
    Dim ws As Worksheet
    Dim ultima As Long
    Dim mascara As String
    Dim c As Variant
    Set ws = Libro().Sheets(HOJA_RESUMEN)
    ultima = ws.Cells(ws.Rows.Count, RES_CODIGO).End(xlUp).Row
    If ultima < FILA_1 Then Exit Sub
    mascara = MascaraCant()
    For Each c In Array(RES_INI_CANT, RES_ENT_CANT, RES_VTA_CANT, RES_OTR_CANT, RES_FIN_CANT)
        ws.Range(ws.Cells(FILA_1, c), ws.Cells(ultima, c)).NumberFormat = mascara
    Next c
    For Each c In Array(RES_INI_BS, RES_ENT_BS, RES_VTA_BS, RES_OTR_BS, RES_FIN_BS)
        ws.Range(ws.Cells(FILA_1, c), ws.Cells(ultima, c)).NumberFormat = FMT_MONEDA
    Next c
End Sub

Private Sub EncabezadoBloque(ByVal ws As Worksheet, ByVal fila As Long)
    ws.Cells(fila, RES_CODIGO).Value = "Codigo"
    ws.Cells(fila, RES_DESCRIPCION).Value = "Descripcion del material"
    ws.Cells(fila, RES_UNIDAD).Value = "Unidad"
    ws.Cells(fila, RES_INI_CANT).Value = "Inicial enero (Cant)"
    ws.Cells(fila, RES_INI_BS).Value = "Inicial enero (Bs)"
    ws.Cells(fila, RES_ENT_CANT).Value = "Entradas (Cant)"
    ws.Cells(fila, RES_ENT_BS).Value = "Entradas (Bs)"
    ws.Cells(fila, RES_VTA_CANT).Value = "Ventas (Cant)"
    ws.Cells(fila, RES_VTA_BS).Value = "Ventas (Bs)"
    ws.Cells(fila, RES_OTR_CANT).Value = "Otros (Cant)"
    ws.Cells(fila, RES_OTR_BS).Value = "Otros (Bs)"
    ws.Cells(fila, RES_FIN_CANT).Value = "Final ultimo mes cerrado (Cant)"
    ws.Cells(fila, RES_FIN_BS).Value = "Final ultimo mes cerrado (Bs)"
    ws.Cells(fila, RES_CHECK).Value = "Conciliacion inicial+entradas-salidas=final"
    ws.Cells(fila, RES_NOTA).Value = "Situacion"
    ws.Range(ws.Cells(fila, RES_CODIGO), ws.Cells(fila, RES_NOTA)).Font.Bold = True
End Sub

Private Function FilaEnHoja(ByVal indice As Long, ByVal codigo As String) As Long
    Dim ws As Worksheet
    Dim f As Long
    Dim clave As String
    clave = NormalizaClave(codigo)
    Set ws = Libro().Sheets(HojaMes(indice))
    For f = FILA_1 To FILA_FIN
        If NormalizaClave(ws.Cells(f, COL_CODIGO).Value2) = clave Then
            FilaEnHoja = f
            Exit Function
        End If
    Next f
    FilaEnHoja = 0
End Function

' Los Bs. se totalizan; las cantidades NO se suman entre unidades distintas.
Private Sub EscribirTotalesResumen(ByVal ws As Worksheet, ByVal fila As Long)
    Dim unidades As Object
    Dim cat As Long
    Dim ultimaCat As Long
    Dim unidad As String
    Dim total As Double
    Dim clave As Variant
    Dim claveBuscada As String

    ultimaCat = UltimaFilaUsada(modCatalogo.HojaCatalogo(), 1)
    Set unidades = CreateObject("Scripting.Dictionary")
    For cat = FILA_1 To ultimaCat
        unidad = CStr(modCatalogo.HojaCatalogo().Cells(cat, 3).Value2)
        claveBuscada = NormalizaClave(modCatalogo.HojaCatalogo().Cells(cat, COL_CODIGO).Value2)
        If claveBuscada <> "" Then
            total = BuscarValor(ws, fila, claveBuscada, RES_INI_BS) + _
                    BuscarValor(ws, fila, claveBuscada, RES_ENT_BS) - _
                    BuscarValor(ws, fila, claveBuscada, RES_VTA_BS) - _
                    BuscarValor(ws, fila, claveBuscada, RES_OTR_BS)
            If unidades.Exists(unidad) Then
                unidades(unidad) = unidades(unidad) + total
            Else
                unidades.Add unidad, total
            End If
        End If
    Next cat

    fila = fila + 2
    ws.Cells(fila, 1).Value = "TOTALES Bs. por unidad de medida (bloque definitivo)"
    ws.Cells(fila, 1).Font.Bold = True
    fila = fila + 1
    ws.Cells(fila, 1).Value = "Unidad"
    ws.Cells(fila, 2).Value = "Total Bs."
    ws.Cells(fila, 3).Value = "Las cantidades no se suman entre unidades distintas"
    ws.Cells(fila, 1).Font.Bold = True
    ws.Cells(fila, 2).Font.Bold = True
    For Each clave In unidades.Keys
        fila = fila + 1
        ws.Cells(fila, 1).Value = clave
        ws.Cells(fila, 2).Value = Redondear2(unidades(clave))
        ws.Cells(fila, 2).NumberFormat = FMT_MONEDA
    Next clave
End Sub

Private Function BuscarValor(ByVal ws As Worksheet, ByVal desde As Long, _
                             ByVal clave As String, ByVal columna As Long) As Double
    Dim f As Long
    For f = desde To ws.Rows.Count
        If EsVacio(ws.Cells(f, RES_CODIGO).Value2) Then Exit Function
        If NormalizaClave(ws.Cells(f, RES_CODIGO).Value2) = clave Then
            BuscarValor = ValorNumericoCelda(ws.Cells(f, columna))
            Exit Function
        End If
    Next f
End Function

Private Function ValorNumericoCelda(ByVal celda As Range) As Double
    If EsNumeroReal(celda.Value2) Then ValorNumericoCelda = CDbl(celda.Value2)
End Function