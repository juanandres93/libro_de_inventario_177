Attribute VB_Name = "modImportacion"
Option Explicit

' Importacion de diciembre del anio anterior (solo enero del anio en curso).
' Reglas: misma entidad y RIF, misma moneda, esquema compatible, diciembre
' CERRADO y validado. Se abre en solo lectura, sin actualizar vinculos y con
' las macros de la fuente deshabilitadas; despues se restaura el estado
' previo (AutomationSecurity, EnableEvents, Calculation, separadores).
' El saldo anterior se copia como VALORES. Un codigo ausente del catalogo no
' se descarta: exige incorporacion o conciliacion. Una fuente invalida nunca
' se sustituye automaticamente por cero.

Private mPendientes As Object

Public Sub ImportarDiciembre()
    Dim ruta As String
    Dim origen As Workbook
    Dim seguridadPrevia As Long
    Dim eventosPrevio As Boolean
    Dim texto As String
    On Error GoTo fallo

    If Reentrada Then Exit Sub
    Entrar
    If AnioDelLibro() < 2001 Then
        Salir
        Err.Raise vbObjectError + 560, "ImportarDiciembre", "Anio de trabajo no valido."
    End If
    If modEstados.EstadoMes(1) <> EST_APERTURA_PEND And modEstados.MesAbierto() <> 0 Then
        Salir
        Err.Raise vbObjectError + 561, "ImportarDiciembre", _
            "La importacion de diciembre solo aplica a enero con APERTURA PENDIENTE."
    End If
    If modEstados.MesAbierto() <> 0 Then
        Salir
        Err.Raise vbObjectError + 562, "ImportarDiciembre", _
            "Ya hay un mes abierto: " & HojaMes(modEstados.MesAbierto()) & "."
    End If
    If modUI.PruebaSilencioso Then
        ' En modo de pruebas no hay dialogo: la seleccion se hace por codigo.
        ruta = ""
    Else
        ruta = Application.GetOpenFilename("Libros de inventarios,*.xlsm", 1, _
            "Seleccione el libro " & (AnioDelLibro() - 1) & "-DIC.xlsm")
    End If
    If Trim$(ruta) = "" Then
        Salir
        Exit Sub
    End If
    If InStr(1, ruta, Libro().FullName, vbTextCompare) > 0 Then
        Salir
        Err.Raise vbObjectError + 563, "ImportarDiciembre", _
            "La fuente no puede ser el propio libro."
    End If

    seguridadPrevia = Application.AutomationSecurity
    eventosPrevio = Application.EnableEvents
    Application.AutomationSecurity = 3          ' msoAutomationSecurityForceDisable
    Application.EnableEvents = False
    On Error Resume Next
    Set origen = Application.Workbooks.Open(ruta, 0, True, 5, "", "", True, 2, "", False, _
        False, 0, False, True, 0)
    On Error GoTo fallo
    If origen Is Nothing Then
        RestaurarSeguridad seguridadPrevia, eventosPrevio
        Salir
        Err.Raise vbObjectError + 564, "ImportarDiciembre", "No se pudo abrir la fuente."
    End If

    texto = ValidarFuente(origen)
    If Left$(texto, 1) = "!" Then
        origen.Close False
        RestaurarSeguridad seguridadPrevia, eventosPrevio
        Salir
        Err.Raise vbObjectError + 565, "ImportarDiciembre", Mid$(texto, 2)
    End If

    texto = Conciliar(origen)
    origen.Close False
    RestaurarSeguridad seguridadPrevia, eventosPrevio
    If Left$(texto, 1) = "!" Then
        Salir
        Err.Raise vbObjectError + 566, "ImportarDiciembre", Mid$(texto, 2)
    End If
    modUI.Aviso texto, vbInformation, APP_NOMBRE
    Salir
    Exit Sub
fallo:
    On Error Resume Next
    If Not origen Is Nothing Then origen.Close False
    Application.AutomationSecurity = seguridadPrevia
    Application.EnableEvents = eventosPrevio
    On Error GoTo 0
    modEstadoExcel.PropagarError "ImportarDiciembre", "modImportacion"
    Salir
End Sub

Private Sub RestaurarSeguridad(ByVal seguridad As Long, ByVal eventos As Boolean)
    On Error Resume Next
    Application.AutomationSecurity = seguridad
    Application.EnableEvents = eventos
    On Error GoTo 0
End Sub

' Devuelve "" si la fuente es valida; "!motivo" si no lo es.
Public Function ValidarFuente(ByVal origen As Workbook) As String
    Dim problemas As String
    Dim anioOrigen As Long
    On Error GoTo fallo
    If Not ExisteHoja(origen, HOJA_CONFIG) Then
        ValidarFuente = "!La fuente no tiene CONFIGURACION."
        Exit Function
    End If
    If UCase$(Trim$(CStr(origen.Sheets(HOJA_CONFIG).Range(CFG_ENTIDAD).Value2))) <> _
       UCase$(Trim$(CStr(Libro().Sheets(HOJA_CONFIG).Range(CFG_ENTIDAD).Value2))) Then
        problemas = problemas & "La entidad no coincide." & vbCrLf
    End If
    If NormalizaClave(origen.Sheets(HOJA_CONFIG).Range(CFG_RIF).Value2) <> _
       NormalizaClave(Libro().Sheets(HOJA_CONFIG).Range(CFG_RIF).Value2) Then
        problemas = problemas & "El RIF no coincide." & vbCrLf
    End If
    If Trim$(CStr(origen.Sheets(HOJA_CONFIG).Range(CFG_MONEDA).Value2)) <> _
       Trim$(CStr(Libro().Sheets(HOJA_CONFIG).Range(CFG_MONEDA).Value2)) Then
        problemas = problemas & "La moneda no coincide." & vbCrLf
    End If
    anioOrigen = CLng(origen.Sheets(HOJA_CONFIG).Range(CFG_ANIO).Value2)
    If anioOrigen <> AnioDelLibro() - 1 Then
        problemas = problemas & "El anio de la fuente (" & anioOrigen & ") no es " & _
            (AnioDelLibro() - 1) & "." & vbCrLf
    End If
    If not TieneHojaDeMes(origen, 12) Then
        problemas = problemas & "La fuente no tiene la hoja de diciembre." & vbCrLf
    End If
    If Not ExistenceDiciembreCerrado(origen) Then
        problemas = problemas & "El diciembre de la fuente no figura CERRADO con evidencia." & vbCrLf
    End If
    If problemas <> "" Then
        ValidarFuente = "!" & problemas
    End If
    Exit Function
fallo:
    ValidarFuente = "!No se pudo leer la fuente: " & Err.Description
End Function

Private Function ExisteHoja(ByVal libro As Workbook, ByVal nombre As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = libro.Sheets(nombre)
    ExisteHoja = (Err.Number = 0) And Not ws Is Nothing
    Err.Clear
    On Error GoTo 0
End Function

Public Function TieneHojaDeMes(ByVal libro As Workbook, ByVal mes As Long) As Boolean
    Dim nombres As Variant
    nombres = NombresMes()
    TieneHojaDeMes = ExisteHoja(libro, CStr(libro.Sheets(HOJA_CONFIG).Range(CFG_ANIO).Value2) & _
        "-" & nombres(mes - 1))
End Function

Public Function ExistenceDiciembreCerrado(ByVal origen As Workbook) As Boolean
    Dim ws As Worksheet
    Dim nombres As Variant
    Dim fila As Long
    nombres = NombresMes()
    On Error GoTo sinEvidencia
    Set ws = origen.Sheets(HOJA_ESTADOS)
    fila = modEstados.FilaEstado(12)
    ExistenceDiciembreCerrado = (Trim$(CStr(ws.Cells(fila, 2).Value2)) = EST_CERRADO) And _
                                (Trim$(CStr(ws.Cells(fila, 7).Value2)) <> "") And _
                                (Not EsVacio(ws.Cells(fila, 4).Value2))
    Exit Function
sinEvidencia:
    ExistenceDiciembreCerrado = False
End Function

' Conciliacion por codigo antes de confirmar: nunca se descarta ni se pone
' cero lo que no esta en la fuente o en el catalogo.
Private Function Conciliar(ByVal origen As Workbook) As String
    Dim wsOrigen As Worksheet
    Dim f As Long
    Dim codigo As String
    Dim saldos As Object
    Dim faltantes As String
    Dim nuevas As String
    Dim linea As String
    Dim nombres As Variant
    Dim totalBs As Double
    Dim pendientes As Long

    nombres = NombresMes()
    Set wsOrigen = origen.Sheets(CStr(origen.Sheets(HOJA_CONFIG).Range(CFG_ANIO).Value2) & _
        "-" & nombres(11))
    Set saldos = CreateObject("Scripting.Dictionary")
    For f = FILA_1 To FILA_FIN
        codigo = Trim$(CStr(wsOrigen.Cells(f, COL_CODIGO).Value2))
        If codigo <> "" Then
            If IsNumeric(wsOrigen.Cells(f, COL_CANT_FIN).Value2) Or _
               IsNumeric(wsOrigen.Cells(f, COL_VAL_FIN).Value2) Then
                saldos(NormalizaClave(codigo)) = Array(wsOrigen.Cells(f, COL_CODIGO).Value2, _
                    CDbl(wsOrigen.Cells(f, COL_CANT_FIN).Value2), _
                    CDbl(wsOrigen.Cells(f, COL_VAL_FIN).Value2), _
                    CDbl(wsOrigen.Cells(f, COL_COSTO_FIN).Value2))
                totalBs = totalBs + CDbl(wsOrigen.Cells(f, COL_VAL_FIN).Value2)
            End If
        End If
    Next f

    linea = "CONCILIACION DE IMPORTACION - diciembre " & (AnioDelLibro() - 1) & vbCrLf & vbCrLf
    linea = linea & "Materiales en la fuente: " & saldos.Count & vbCrLf
    linea = linea & "Total Bs. del cierre de diciembre: " & Format$(totalBs, FMT_MONEDA) & vbCrLf & vbCrLf

    ' codigos de la fuente ausentes del catalogo actual
    Dim clave As Variant
    For Each clave In saldos.Keys
        If FilaEnCatalogo(saldos(clave)(0)) = 0 Then
            faltantes = faltantes & saldos(clave)(0) & " (" & _
                Format$(saldos(clave)(2), FMT_MONEDA) & ")" & vbCrLf
        End If
    Next clave
    ' materiales del catalogo sin antecedente en la fuente
    For f = FILA_1 To FILA_FIN
        codigo = Trim$(CStr(modCatalogo.HojaCatalogo().Cells(f, 1).Value2))
        If codigo <> "" And Not saldos.Exists(NormalizaClave(codigo)) Then
            nuevas = nuevas & codigo & vbCrLf
        End If
    Next f

    If faltantes <> "" Then
        linea = linea & "PENDIENTE: saldos sin codigo en el catalogo actual (no se descartan):" & _
            vbCrLf & faltantes & vbCrLf
    End If
    If nuevas <> "" Then
        linea = linea & "PENDIENTE: materiales del catalogo sin antecedente (alta nueva o " & _
            "apertura inicial con soporte):" & vbCrLf & nuevas & vbCrLf
    End If
    pendientes = CountLines(faltantes) + CountLines(nuevas)
    If pendientes = 0 Then
        linea = linea & "Sin pendientes. Confirme para copiar K/L a C/D de enero como VALORES." & vbCrLf
        Set mPendientes = saldos
    Else
        linea = linea & "NO se confirma: hay " & pendientes & " pendientes por conciliar."
    End If
    modBitacora.Registrar "Importacion de diciembre", "origen " & origen.Name, _
        IIf(pendientes = 0, "CONCILIADO", "PENDIENTES"), _
        "Fuente " & (AnioDelLibro() - 1) & "-DIC", saldos.Count, totalBs, "", ""
    Conciliar = linea
End Function

Private Function CountLines(ByVal texto As String) As Long
    Dim partes As Variant
    If Trim$(texto) = "" Then Exit Function
    partes = Split(Replace(texto, vbCrLf, "|"), "|")
    CountLines = UBound(partes) + 1
End Function

' Confirmacion: copia el saldo de diciembre a C/D de enero como VALORES.
Public Sub ConfirmarImportacionEnero()
    Dim destino As Worksheet
    Dim f As Long
    Dim codigo As String
    Dim fila As Long
    Dim clave As Variant
    Dim ids As String
    On Error GoTo fallo
    If mPendientes Is Nothing Then
        Err.Raise vbObjectError + 567, "ConfirmarImportacionEnero", _
            "No hay conciliacion de importacion confirmada en esta sesion."
    End If
    If modEstados.MesAbierto() <> 0 Then
        Err.Raise vbObjectError + 568, "ConfirmarImportacionEnero", _
            "Ya hay un mes abierto."
    End If
    modEstadoExcel.CapturarEstado
    modEstadoExcel.SinPantalla
    Set destino = Libro().Sheets(HojaMes(1))
    Application.Calculate
    For Each clave In mPendientes.Keys
        codigo = mPendientes(clave)(0)
        fila = FilaEnHoja(1, codigo)
        If fila = 0 Then
            fila = FILA_1
            Do While Not EsVacio(destino.Cells(fila, COL_CODIGO).Value2)
                fila = fila + 1
            Loop
            If fila > FILA_FIN Then
                Err.Raise vbObjectError + 569, "ConfirmarImportacionEnero", _
                    "No caben los materiales importados en el rango tecnico."
            End If
            destino.Cells(fila, COL_CODIGO).Value = codigo
            destino.Cells(fila, COL_DESCRIPCION).Value = modCatalogo.HojaCatalogo().Cells( _
                FilaEnCatalogo(codigo), 2).Value2
            destino.Cells(fila, COL_UNIDAD).Value = modCatalogo.HojaCatalogo().Cells( _
                FilaEnCatalogo(codigo), 3).Value2
        End If
        destino.Cells(fila, COL_CANT_INI).Value = mPendientes(clave)(1)
        destino.Cells(fila, COL_VAL_INI).Value = mPendientes(clave)(2)
        destino.Cells(fila, COL_CANT_ENT).Value = 0
        destino.Cells(fila, COL_VAL_ENT).Value = 0
        destino.Cells(fila, COL_CANT_VTA).Value = 0
        destino.Cells(fila, COL_CANT_OTR).Value = 0
    Next clave
    modCalculo.AplicarFormulas destino, FILA_1, FILA_FIN
    modCalculo.FormatearCostos destino, FILA_1, FILA_FIN
    modCalculo.EscribirTotales destino
    destino.Range(destino.Cells(FILA_1, COL_CANT_INI), destino.Cells(FILA_FIN, COL_VAL_INI)) _
        .Value = destino.Range(destino.Cells(FILA_1, COL_CANT_INI), destino.Cells(FILA_FIN, COL_VAL_INI)).Value
    modEstados.FijarEstadoMes 1, EST_ABIERTO, Now, "", _
        Libro().Sheets(HOJA_CONFIG).Range(CFG_RESPONSABLE).Value2, "", "", _
        "Apertura de enero por importacion de diciembre"
    modEstados.DesbloquearINPUT destino
    Set mPendientes = Nothing
    modBitacora.Registrar "Confirmacion de importacion", HojaMes(1), "OK", _
        "Saldo de diciembre copiado como valores en C/D", "", "", "", ""
    modEstadoExcel.RestaurarEstado
    Exit Sub
fallo:
    modEstados.FijarEstadoMes 1, EST_INCIDENCIA, "", "", _
        Libro().Sheets(HOJA_CONFIG).Range(CFG_RESPONSABLE).Value2, "", "", _
        "Fallo en confirmacion de importacion: " & Err.Description
    modEstadoExcel.RestaurarEstado
    modEstadoExcel.PropagarError "ConfirmarImportacionEnero", "modImportacion"
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