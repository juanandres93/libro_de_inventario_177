Attribute VB_Name = "modCatalogo"
Option Explicit

' CATALOGO: unica fuente maestra. Las relaciones se hacen por codigo
' normalizado, nunca por posicion de fila.

Private Const CAT_CODIGO As Long = 1
Private Const CAT_DESCRIPCION As Long = 2
Private Const CAT_UNIDAD As Long = 3
Private Const CAT_ACTIVO As Long = 4
Private Const CAT_FECHA_ALTA As Long = 5
Private Const CAT_FECHA_BAJA As Long = 6

Public Function HojaCatalogo() As Worksheet
    Set HojaCatalogo = Libro().Sheets(HOJA_CATALOGO)
End Function

' Los codigos de material son texto. En una celda de formato General Excel
' convierte "0007" en el numero 7 al escribir o al pegar, y los ceros
' iniciales desaparecen: la clave dejaria de ser estable. Se fija el formato
' Texto en las tres columnas de codigo (catalogo, mes y detalle).
Public Sub AsegurarFormatoCodigos()
    Dim ws As Worksheet
    Dim i As Long
    On Error Resume Next
    Set ws = HojaCatalogo()
    If Not ws Is Nothing Then
        ws.Range(ws.Cells(FILA_1, CAT_CODIGO), ws.Cells(FILA_FIN, CAT_CODIGO)).NumberFormat = "@"
    End If
    For i = 1 To 12
        Set ws = Libro().Sheets(HojaMes(i))
        If Not ws Is Nothing Then
            ws.Range(ws.Cells(FILA_1, COL_CODIGO), ws.Cells(FILA_FIN, COL_CODIGO)).NumberFormat = "@"
        End If
    Next i
    Set ws = Libro().Sheets(HOJA_DETALLE)
    If Not ws Is Nothing Then
        ws.Range(ws.Cells(FILA_1, DET_CODIGO), ws.Cells(FILA_FIN, DET_CODIGO)).NumberFormat = "@"
    End If
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' Validacion del maestro completo (una sola pasada: sin O(n^2))
' ---------------------------------------------------------------------------
Public Function ValidarCatalogo() As String
    Dim problemas As String
    Dim ws As Worksheet
    Dim f As Long
    Dim codigo As String
    Dim clave As String
    Dim vistos As Object
    Dim pendientes As String
    Set ws = HojaCatalogo()
    Set vistos = CreateObject("Scripting.Dictionary")
    For f = FILA_1 To FILA_FIN
        codigo = CStr(ws.Cells(f, CAT_CODIGO).Value2)
        If Trim$(codigo) <> "" Then
            pendientes = ""
            If InStr(codigo, vbLf) > 0 Or InStr(codigo, vbCr) > 0 Then
                problemas = problemas & "- Fila " & f & ": el codigo contiene saltos de linea." & vbCrLf
            End If
            If Len(LimpiarTexto(codigo)) = 0 Then
                problemas = problemas & "- Fila " & f & ": codigo vacio o solo espacios." & vbCrLf
            End If
            If Len(LimpiarTexto(codigo)) > 50 Then
                problemas = problemas & "- Fila " & f & ": codigo de mas de 50 caracteres." & vbCrLf
            End If
            clave = NormalizaClave(codigo)
            If clave <> "" Then
                If vistos.Exists(clave) Then
                    problemas = problemas & "- Fila " & f & ": codigo duplicado (normalizado) " & _
                        Trim$(codigo) & "; ya aparece en la fila " & vistos(clave) & "." & vbCrLf
                Else
                    vistos.Add clave, f
                End If
            End If
            If EsVacio(ws.Cells(f, CAT_DESCRIPCION).Value2) Then
                problemas = problemas & "- Fila " & f & ": descripcion vacia." & vbCrLf
            ElseIf Len(LimpiarTexto(ws.Cells(f, CAT_DESCRIPCION).Value2)) > 250 Then
                problemas = problemas & "- Fila " & f & ": descripcion de mas de 250 caracteres." & vbCrLf
            End If
            If EsVacio(ws.Cells(f, CAT_UNIDAD).Value2) Then
                problemas = problemas & "- Fila " & f & ": unidad de medida vacia." & vbCrLf
            End If
            If NormalizaClave(ws.Cells(f, CAT_ACTIVO).Value2) <> "SI" And _
               NormalizaClave(ws.Cells(f, CAT_ACTIVO).Value2) <> "NO" Then
                problemas = problemas & "- Fila " & f & ": Activo debe ser Si o No." & vbCrLf
            End If
            If Not EsVacio(ws.Cells(f, CAT_FECHA_ALTA).Value2) Then
                If Not EsFecha(ws.Cells(f, CAT_FECHA_ALTA).Value2) Then
                    problemas = problemas & "- Fila " & f & ": fecha de alta no es fecha." & vbCrLf
                End If
            End If
            If Not EsVacio(ws.Cells(f, CAT_FECHA_BAJA).Value2) Then
                If Not EsFecha(ws.Cells(f, CAT_FECHA_BAJA).Value2) Then
                    problemas = problemas & "- Fila " & f & ": fecha de baja no es fecha." & vbCrLf
                End If
            End If
            If Trim$(CStr(ws.Cells(f, CAT_UNIDAD).Value2)) <> "" And _
               NormalizaClave(ws.Cells(f, CAT_ACTIVO).Value2) = "NO" And _
               Not EsVacio(ws.Cells(f, CAT_FECHA_BAJA).Value2) Then
                ' inactivo con fecha de baja: control adicional
                If Not EsFecha(ws.Cells(f, CAT_FECHA_BAJA).Value2) Then
                    problemas = problemas & "- Fila " & f & ": baja sin fecha valida." & vbCrLf
                End If
            End If
        ElseIf Trim$(pendientes) = "" And Trim$(CStr(ws.Cells(f, CAT_CODIGO).Value2)) = "" Then
            ' se detectan los saltos al final, en una sola pasada
        End If
    Next f
    problemas = problemas & RevisarSaltosMaestro(ws)
    ValidarCatalogo = problemas
End Function

' Detecta filas totally vacias entre dos filas con datos (una sola pasada).
Private Function RevisarSaltosMaestro(ByVal ws As Worksheet) As String
    Dim f As Long
    Dim ultimoConDato As Long
    Dim texto As String
    ultimoConDato = 0
    For f = FILA_1 To FILA_FIN
        If Trim$(CStr(ws.Cells(f, CAT_CODIGO).Value2)) <> "" Then
            If ultimoConDato > 0 And f > ultimoConDato + 1 Then
                texto = texto & "- Filas " & (ultimoConDato + 1) & " a " & (f - 1) & _
                    ": fila vacia dentro del maestro." & vbCrLf
            End If
            ultimoConDato = f
        End If
    Next f
    RevisarSaltosMaestro = texto
End Function

' ---------------------------------------------------------------------------
' Alta y modificacion con controles
' ---------------------------------------------------------------------------
Public Function PuedeUsarMaterial(ByVal codigo As Variant) As Boolean
    Dim fila As Long
    fila = FilaEnCatalogo(codigo)
    If fila = 0 Then
        PuedeUsarMaterial = False
        Exit Function
    End If
    PuedeUsarMaterial = (NormalizaClave(HojaCatalogo().Cells(fila, CAT_ACTIVO).Value2) = "SI") Or _
                        SaldoPendiente(CStr(codigo))
End Function

' Un material inactivo con saldo sigue apareciendo hasta agotar existencias.
' Se revisa solo el ultimo mes cerrado (o el ultimo mes con datos), no los
' doce meses completos.
Public Function SaldoPendiente(ByVal codigo As String) As Boolean
    Dim clave As String
    Dim ultimo As Long
    Dim i As Long
    Dim hoja As Worksheet
    Dim f As Long
    clave = NormalizaClave(codigo)
    If clave = "" Then Exit Function
    ultimo = modEstados.MesCerradoHasta()
    If ultimo < 1 Then ultimo = UltimoMesConMovimiento()
    If ultimo < 1 Then Exit Function
    For i = ultimo To 1 Step -1
        Set hoja = Libro().Sheets(HojaMes(i))
        For f = FILA_1 To FILA_FIN
            If NormalizaClave(hoja.Cells(f, COL_CODIGO).Value2) = clave Then
                If EsNumeroReal(hoja.Cells(f, COL_CANT_FIN).Value2) Then
                    If Abs(CDbl(hoja.Cells(f, COL_CANT_FIN).Value2)) > 0.00005 Then
                        SaldoPendiente = True
                        Exit Function
                    End If
                End If
                Exit For
            End If
        Next f
    Next i
End Function

Private Function UltimoMesConMovimiento() As Long
    Dim i As Long
    Dim hoja As Worksheet
    Dim f As Long
    For i = 12 To 1 Step -1
        Set hoja = Libro().Sheets(HojaMes(i))
        For f = FILA_1 To FILA_FIN
            If Trim$(CStr(hoja.Cells(f, COL_CODIGO).Value2)) <> "" Then
                If EsNumeroReal(hoja.Cells(f, COL_CANT_ENT).Value2) Then
                    If CDbl(hoja.Cells(f, COL_CANT_ENT).Value2) <> 0 Then
                        UltimoMesConMovimiento = i
                        Exit Function
                    End If
                End If
                Exit For
            End If
        Next f
    Next i
End Function

Public Sub AltaModificarMaterial()
    Dim ws As Worksheet
    Dim codigo As String
    Dim fila As Long
    Dim unidadVieja As String
    Dim unidadNueva As String
    Dim problemas As String
    Dim original As Variant
    Dim existed As Boolean
    Dim protegido As Boolean
    Dim numero As Long
    Dim detalle As String
    On Error GoTo fallo
    If Reentrada Then Exit Sub
    Entrar
    Set ws = HojaCatalogo()
    codigo = modUI.CapturarTexto("Codigo del material (texto, hasta 50 caracteres). " & _
                            "Dejar vacio para cancelar.", "Alta o modificacion de material")
    If codigo = "" Then
        Salir
        Exit Sub
    End If
    fila = FilaEnCatalogo(codigo)
    existed = (fila > 0)
    If existed Then
        unidadVieja = Trim$(CStr(ws.Cells(fila, CAT_UNIDAD).Value2))
        original = ws.Range(ws.Cells(fila, CAT_CODIGO), ws.Cells(fila, CAT_FECHA_BAJA)).Value
    Else
        fila = FILA_1
        Do While Not EsVacio(ws.Cells(fila, CAT_CODIGO).Value2)
            fila = fila + 1
        Loop
        If fila > FILA_FIN Then
            Salir
            Err.Raise vbObjectError + 520, "AltaModificarMaterial", _
                "Se alcanzo el limite tecnico de " & LIMITE_MATERIALES & " materiales."
        End If
        original = Empty
    End If
    unidadNueva = modUI.CapturarTexto("Unidad de medida (obligatoria).", "Unidad", unidadVieja)
    If existed And unidadVieja <> "" And UCase$(unidadVieja) <> UCase$(unidadNueva) Then
        If TieneMovimientos(codigo) Then
            Salir
            Err.Raise vbObjectError + 521, "AltaModificarMaterial", _
                "No se cambia la unidad de un material con movimientos (" & codigo & _
                "): use un codigo nuevo o una migracion expresamente aprobada."
        End If
    End If

    protegido = (ws.ProtectContents Or ws.ProtectDrawingObjects)
    If protegido Then ws.Unprotect Password:=CLAVE_PROTECCION
    ws.Cells(fila, CAT_CODIGO).Value = codigo
    ws.Cells(fila, CAT_DESCRIPCION).Value = modUI.CapturarTexto( _
        "Descripcion del material (obligatoria, hasta 250).", "Descripcion", _
        CStr(ws.Cells(fila, CAT_DESCRIPCION).Value2))
    ws.Cells(fila, CAT_UNIDAD).Value = unidadNueva
    ws.Cells(fila, CAT_ACTIVO).Value = modUI.CapturarTexto("Activo (Si/No).", "Activo", _
        IIf(NormalizaClave(ws.Cells(fila, CAT_ACTIVO).Value2) = "NO", "No", "Si"))
    If EsVacio(ws.Cells(fila, CAT_FECHA_ALTA).Value2) Then ws.Cells(fila, CAT_FECHA_ALTA).Value = Date

    problemas = ValidarCatalogo()
    If problemas <> "" Then
        ws.Range(ws.Cells(fila, CAT_CODIGO), ws.Cells(fila, CAT_FECHA_BAJA)).ClearContents
        If existed Then ws.Range(ws.Cells(fila, CAT_CODIGO), ws.Cells(fila, CAT_FECHA_BAJA)).Value = original
        If protegido Then ws.Protect Password:=CLAVE_PROTECCION, UserInterfaceOnly:=True
        InvalidarIndice
        Salir
        Err.Raise vbObjectError + 522, "AltaModificarMaterial", _
            "El maestro no es valido, se revirtio el cambio:" & vbCrLf & problemas
    End If
    If protegido Then ws.Protect Password:=CLAVE_PROTECCION, UserInterfaceOnly:=True
    InvalidarIndice
    modBitacora.RegistrarCambio IIf(existed, "Modificacion de material", "Alta de material"), _
        codigo, unidadVieja, unidadNueva, "Control de catalogo"
    Salir
    Exit Sub
fallo:
    numero = Err.Number
    detalle = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then
        ws.Unprotect Password:=CLAVE_PROTECCION
        If fila > 0 Then
            ws.Range(ws.Cells(fila, CAT_CODIGO), ws.Cells(fila, CAT_FECHA_BAJA)).ClearContents
            If existed Then
                ws.Range(ws.Cells(fila, CAT_CODIGO), ws.Cells(fila, CAT_FECHA_BAJA)).Value = original
            End If
        End If
        ws.Protect Password:=CLAVE_PROTECCION, UserInterfaceOnly:=True
    End If
    Err.Clear
    On Error GoTo 0
    InvalidarIndice
    Salir
    Err.Raise numero, "modCatalogo.AltaModificarMaterial", _
        "No se pudo registrar el material; se revirtieron los cambios. " & detalle
End Sub

Public Function TieneMovimientos(ByVal codigo As String) As Boolean
    Dim clave As String
    Dim hoja As Worksheet
    Dim f As Long
    Dim i As Long
    clave = NormalizaClave(codigo)
    If clave = "" Then Exit Function
    For i = 1 To 12
        Set hoja = Libro().Sheets(HojaMes(i))
        For f = FILA_1 To FILA_FIN
            If NormalizaClave(hoja.Cells(f, COL_CODIGO).Value2) = clave Then
                If EsNumeroReal(hoja.Cells(f, COL_CANT_ENT).Value2) Then
                    If CDbl(hoja.Cells(f, COL_CANT_ENT).Value2) <> 0 Then
                        TieneMovimientos = True
                        Exit Function
                    End If
                End If
                If EsNumeroReal(hoja.Cells(f, COL_VAL_ENT).Value2) Then
                    If CDbl(hoja.Cells(f, COL_VAL_ENT).Value2) <> 0 Then
                        TieneMovimientos = True
                        Exit Function
                    End If
                End If
                If EsNumeroReal(hoja.Cells(f, COL_CANT_VTA).Value2) Then
                    If CDbl(hoja.Cells(f, COL_CANT_VTA).Value2) <> 0 Then
                        TieneMovimientos = True
                        Exit Function
                    End If
                End If
                If EsNumeroReal(hoja.Cells(f, COL_CANT_OTR).Value2) Then
                    If CDbl(hoja.Cells(f, COL_CANT_OTR).Value2) <> 0 Then
                        TieneMovimientos = True
                        Exit Function
                    End If
                End If
                Exit For
            End If
        Next f
    Next i
End Function

' ---------------------------------------------------------------------------
' Bloque pegado: se valida antes de aceptar. Si hay un error, el evento de
' hoja revierte TODO el cambio (no hay carga parcial).
' ---------------------------------------------------------------------------
Public Function ValidarBloquePegado(ByVal ws As Worksheet, ByVal rango As Range) As String
    Dim problemas As String
    Dim celda As Range
    Dim permitidas As Boolean
    Dim valores As Variant
    Dim f As Long, c As Long
    Dim primeraFila As Long, primeraCol As Long
    Dim esArreglo As Boolean
    permitidas = EsHojaMes(ws.Name)
    primeraFila = rango.Row
    primeraCol = rango.Column
    valores = rango.Value2
    esArreglo = IsArray(valores)
    ' Value2 devuelve el resultado de una formula, no la formula: un =1+1
    ' pareceria un 2 y pasaria el control numerico. En columnas de entrada
    ' solo se admiten valores fijos.
    For Each celda In rango
        If celda.HasFormula Then
            If permitidas And (celda.Column = COL_CANT_ENT Or celda.Column = COL_VAL_ENT Or _
                            celda.Column = COL_CANT_VTA Or celda.Column = COL_CANT_OTR) Then
                problemas = problemas & "Fila " & celda.Row & ", columna " & _
                    ColumnaLetra(celda.Column) & _
                    ": la celda contiene una formula; no es un numero fijo." & vbCrLf
            End If
        End If
    Next celda
    If Not esArreglo Then
        For Each celda In rango
            problemas = problemas & RevisarValorPegado(ws, celda.Row, celda.Column, celda.Value2, permitidas)
        Next celda
    ElseIf UBound(valores, 1) - LBound(valores, 1) = 0 And _
           UBound(valores, 2) - LBound(valores, 2) = 0 Then
        For Each celda In rango
            problemas = problemas & RevisarValorPegado(ws, celda.Row, celda.Column, celda.Value2, permitidas)
        Next celda
    Else
        For f = LBound(valores, 1) To UBound(valores, 1)
            For c = LBound(valores, 2) To UBound(valores, 2)
                problemas = problemas & RevisarValorPegado(ws, _
                    primeraFila + f - LBound(valores, 1), _
                    primeraCol + c - LBound(valores, 2), _
                    valores(f, c), permitidas)
            Next c
        Next f
    End If
    ValidarBloquePegado = problemas
End Function

Private Function RevisarValorPegado(ByVal ws As Worksheet, ByVal f As Long, ByVal c As Long, _
                                    ByVal valor As Variant, ByVal mensual As Boolean) As String
    Dim texto As String
    If f < FILA_1 Or f > FILA_FIN Then
        RevisarValorPegado = "Fila " & f & ": fuera del rango tecnico de materiales." & vbCrLf
        Exit Function
    End If
    If mensual And (c = COL_CANT_ENT Or c = COL_VAL_ENT Or c = COL_CANT_VTA Or c = COL_CANT_OTR) Then
        If Not EsVacio(valor) Then
            If Not EsNumeroReal(valor) Then
                texto = texto & "Fila " & f & ", columna " & ColumnaLetra(c) & _
                    ": " & TypeName(valor) & " no es un numero." & vbCrLf
            ElseIf CDbl(valor) < 0 Then
                texto = texto & "Fila " & f & ", columna " & ColumnaLetra(c) & _
                    ": valor negativo." & vbCrLf
            End If
        End If
    ElseIf mensual And c = COL_CODIGO Then
        If Not EsVacio(valor) Then
            If InStr(CStr(valor), vbLf) > 0 Or InStr(CStr(valor), vbCr) > 0 Then
                texto = texto & "Fila " & f & ": el codigo no puede traer saltos de linea." & vbCrLf
            End If
            If FilaEnCatalogo(valor) = 0 And Trim$(CStr(valor)) <> "" Then
                texto = texto & "Fila " & f & ": codigo " & Trim$(CStr(valor)) & _
                    " no existe en CATALOGO." & vbCrLf
            End If
        End If
    Else
        texto = texto & "Fila " & f & ", columna " & ColumnaLetra(c) & _
            ": celda bloqueada, no admite pegado." & vbCrLf
    End If
    RevisarValorPegado = texto
End Function

Public Function ColumnaLetra(ByVal columna As Long) As String
    Dim n As Long
    n = columna
    ColumnaLetra = Chr$(65 + ((n - 1) Mod 26))
End Function