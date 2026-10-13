Attribute VB_Name = "modEstados"
Option Explicit

' Estados: NO ABIERTO, APERTURA PENDIENTE, ABIERTO, CERRADO, INCIDENCIA.
' Solo un mes abierto a la vez. Secuencia obligatoria: no se abre marzo sin
' febrero cerrado, aunque febrero no tenga movimientos. Los botones son
' idempotentes: repetirlos no duplica ni sobrescribe aperturas confirmadas.

Public Function HojaEstados() As Worksheet
    Set HojaEstados = Libro().Sheets(HOJA_ESTADOS)
End Function

Public Function FilaEstado(ByVal indice As Long) As Long
    ' _ESTADOS: encabezados en la fila 3, meses en 4..15
    FilaEstado = 3 + indice
End Function

Public Function FechaAperturaMes(ByVal indice As Long) As Variant
    On Error GoTo sinFecha
    FechaAperturaMes = HojaEstados().Cells(FilaEstado(indice), 3).Value2
    Exit Function
sinFecha:
    FechaAperturaMes = ""
End Function

Public Function FechaCierreMes(ByVal indice As Long) As Variant
    On Error GoTo sinFecha
    FechaCierreMes = HojaEstados().Cells(FilaEstado(indice), 4).Value2
    Exit Function
sinFecha:
    FechaCierreMes = ""
End Function

Public Function EstadoMes(ByVal indice As Long) As String
    On Error GoTo sinEstado
    EstadoMes = Trim$(CStr(HojaEstados().Cells(FilaEstado(indice), 2).Value2))
    If EstadoMes = "" Then EstadoMes = EST_NO_ABIERTO
    Exit Function
sinEstado:
    EstadoMes = EST_NO_ABIERTO
End Function

Public Sub FijarEstadoMes(ByVal indice As Long, ByVal estado As String, _
                          ByVal apertura As Variant, ByVal cierre As Variant, _
                          ByVal responsable As String, ByVal identificador As String, _
                          ByVal huella As String, ByVal observacion As String)
    Dim ws As Worksheet
    Dim protegido As Boolean
    Set ws = HojaEstados()
    protegido = (ws.ProtectContents Or ws.ProtectDrawingObjects)
    If protegido Then ws.Unprotect Password:=CLAVE_PROTECCION
    With ws
        .Cells(FilaEstado(indice), 1).Value = HojaMes(indice)
        .Cells(FilaEstado(indice), 2).Value = estado
        If Not EsVacio(apertura) Then .Cells(FilaEstado(indice), 3).Value = apertura
        If Not EsVacio(cierre) Then .Cells(FilaEstado(indice), 4).Value = cierre
        .Cells(FilaEstado(indice), 5).Value = responsable
        .Cells(FilaEstado(indice), 6).Value = APP_VERSION
        .Cells(FilaEstado(indice), 7).Value = identificador
        .Cells(FilaEstado(indice), 8).Value = huella
        .Cells(FilaEstado(indice), 9).Value = observacion
        .Cells(FilaEstado(indice), 3).NumberFormat = "yyyy-mm-dd hh:mm:ss"
        .Cells(FilaEstado(indice), 4).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    End With
    If protegido Then ws.Protect Password:=CLAVE_PROTECCION, UserInterfaceOnly:=True, _
        DrawingObjects:=True, Contents:=True, Scenarios:=False
End Sub

' Unico mes abierto: devuelve el indice o 0.
Public Function MesAbierto() As Long
    Dim i As Long
    For i = 1 To 12
        If EstadoMes(i) = EST_ABIERTO Or EstadoMes(i) = EST_INCIDENCIA Then
            MesAbierto = i
            Exit Function
        End If
    Next i
    MesAbierto = 0
End Function

Public Function MesCerradoHasta() As Long
    Dim i As Long
    For i = 1 To 12
        If EstadoMes(i) <> EST_CERRADO Then Exit For
        MesCerradoHasta = i
    Next i
End Function

Public Function PuedeCerrar(ByVal indice As Long) As String
    Dim problemas As String
    If EstadoMes(indice) <> EST_ABIERTO Then
        problemas = problemas & "- El mes " & HojaMes(indice) & " no esta ABIERTO (estado: " & _
            EstadoMes(indice) & ")." & vbCrLf
    End If
    If indice > 1 Then
        If EstadoMes(indice - 1) <> EST_CERRADO Then
            problemas = problemas & "- El mes anterior " & HojaMes(indice - 1) & _
                " no esta CERRADO." & vbCrLf
        End If
    End If
    problemas = problemas & modValidacion.ValidarMes(indice)
    PuedeCerrar = problemas
End Function

' ---------------------------------------------------------------------------
' Apertura
' ---------------------------------------------------------------------------
Public Function AbrirMes(ByVal indice As Long) As String
    Dim problemas As String
    Dim numero As Long
    Dim detalle As String
    If Not AnioConfigurado() Then
        AbrirMes = "- El libro no esta configurado: capture la configuracion " & _
            "(anio) antes de abrir meses." & vbCrLf
        Exit Function
    End If
    If MesAbierto() <> 0 And MesAbierto() <> indice Then
        problemas = problemas & "- Ya hay un mes abierto: " & HojaMes(MesAbierto()) & _
            ". Solo puede haber uno." & vbCrLf
    End If
    If EstadoMes(indice) = EST_ABIERTO Then
        AbrirMes = problemas
        Exit Function
    End If
    If EstadoMes(indice) = EST_CERRADO Then
        problemas = problemas & "- El mes " & HojaMes(indice) & _
            " esta CERRADO y no hay boton de reapertura en esta version." & vbCrLf
    End If
    If indice > 1 Then
        If EstadoMes(indice - 1) <> EST_CERRADO Then
            problemas = problemas & "- No se puede abrir " & HojaMes(indice) & _
                " sin cerrar " & HojaMes(indice - 1) & "." & vbCrLf
        End If
    End If
    If indice = 1 And EstadoMes(1) <> EST_APERTURA_PEND Then
        problemas = problemas & "- Enero requiere apertura pendiente confirmada o " & _
            "importacion de diciembre del anio anterior." & vbCrLf
    End If
    If indice > 1 Then
        problemas = problemas & RevisarAperturaDestino(indice)
    End If
    If problemas <> "" Then
        AbrirMes = problemas
        Exit Function
    End If

    modEstadoExcel.CapturarEstado
    modEstadoExcel.SinPantalla
    On Error GoTo fallo
    If indice > 1 Then
        If CopiarMesAnterior(indice) Then
            FijarEstadoMes indice, EST_ABIERTO, Now, "", _
                Libro().Sheets(HOJA_CONFIG).Range(CFG_RESPONSABLE).Value2, _
                "", "", "Apertura desde " & HojaMes(indice - 1)
            DesbloquearINPUT Libro().Sheets(HojaMes(indice))
            modBitacora.Registrar "Apertura de mes", HojaMes(indice), "OK", _
                "Cierre verificado del mes anterior", "", "", HojaMes(indice), ""
            problemas = ""
        Else
            problemas = problemas & "- No se pudo copiar el maestro y el saldo de " & _
                HojaMes(indice - 1) & "." & vbCrLf
        End If
    Else
        FijarEstadoMes indice, EST_ABIERTO, Now, "", _
            Libro().Sheets(HOJA_CONFIG).Range(CFG_RESPONSABLE).Value2, _
            "", "", "Apertura de enero"
        DesbloquearINPUT Libro().Sheets(HojaMes(indice))
        modBitacora.Registrar "Apertura de mes", HojaMes(indice), "OK", _
            "Enero", "", "", HojaMes(indice), ""
        problemas = ""
    End If
    AbrirMes = problemas
    GoTo restaura
fallo:
    numero = Err.Number
    detalle = Err.Description
    Err.Clear
    On Error Resume Next
    FijarEstadoMes indice, EST_INCIDENCIA, "", "", _
        Libro().Sheets(HOJA_CONFIG).Range(CFG_RESPONSABLE).Value2, _
        "", "", "Fallo en apertura: " & detalle
    On Error GoTo 0
    AbrirMes = "- Fallo en la apertura de " & HojaMes(indice) & ": " & detalle & vbCrLf
restaura:
    modEstadoExcel.RestaurarEstado
    If numero <> 0 Then
        Err.Raise numero, "modEstados.AbrirMes", detalle
    End If
End Function

' Copia el MAESTRO (catalogo) y el saldo de cierre del mes anterior por CODIGO;
' C/D quedan como VALORES. Un material inactivo con saldo si entra.
Private Function CopiarMesAnterior(ByVal indice As Long) As Boolean
    Dim origen As Worksheet
    Dim destino As Worksheet
    Dim cat As Worksheet
    Dim f As Long
    Dim codigo As String
    Dim problemas As String
    Dim filaSaldo As Long
    Set origen = Libro().Sheets(HojaMes(indice - 1))
    Set destino = Libro().Sheets(HojaMes(indice))
    Set cat = Libro().Sheets(HOJA_CATALOGO)
    If TieneDatosDebajo(destino) Then
        CopiarMesAnterior = False
        Exit Function
    End If
    On Error GoTo fallo
    destino.Unprotect Password:=CLAVE_PROTECCION
    destino.Cells.Locked = True
    For f = FILA_1 To FILA_FIN
        codigo = Trim$(CStr(cat.Cells(f, COL_CODIGO).Value2))
        If codigo <> "" Then
            If PuedeUsarMaterial(codigo) Then
                destino.Cells(f, COL_CODIGO).Value = cat.Cells(f, COL_CODIGO).Value2
                destino.Cells(f, COL_DESCRIPCION).Value = cat.Cells(f, COL_DESCRIPCION).Value2
                destino.Cells(f, COL_UNIDAD).Value = cat.Cells(f, COL_UNIDAD).Value2
                filaSaldo = FilaDeCodigoEn(origen, codigo)
                If filaSaldo > 0 Then
                    destino.Cells(f, COL_CANT_INI).Value = modCalculo.ValorNumerico(origen, filaSaldo, COL_CANT_FIN)
                    destino.Cells(f, COL_VAL_INI).Value = modCalculo.ValorNumerico(origen, filaSaldo, COL_VAL_FIN)
                End If
            End If
        End If
    Next f
    For f = FILA_1 To FILA_FIN
        codigo = Trim$(CStr(origen.Cells(f, COL_CODIGO).Value2))
        If codigo <> "" And FilaEnCatalogo(codigo) = 0 Then
            problemas = problemas & "- Material " & codigo & " de " & HojaMes(indice - 1) & _
                " no esta en el catalogo: no se puede abrir " & HojaMes(indice) & "." & vbCrLf
        End If
        If codigo <> "" And Not PuedeUsarMaterial(codigo) Then
            If modCalculo.ValorNumerico(origen, f, COL_CANT_FIN) <> 0 Or _
               modCalculo.ValorNumerico(origen, f, COL_VAL_FIN) <> 0 Then
                destino.Cells(f, COL_CODIGO).Value = origen.Cells(f, COL_CODIGO).Value2
                destino.Cells(f, COL_DESCRIPCION).Value = origen.Cells(f, COL_DESCRIPCION).Value2
                destino.Cells(f, COL_UNIDAD).Value = origen.Cells(f, COL_UNIDAD).Value2
                destino.Cells(f, COL_CANT_INI).Value = modCalculo.ValorNumerico(origen, f, COL_CANT_FIN)
                destino.Cells(f, COL_VAL_INI).Value = modCalculo.ValorNumerico(origen, f, COL_VAL_FIN)
            End If
        End If
    Next f
    If problemas <> "" Then
        DestinoEnBlanco destino
        CopiarMesAnterior = False
        Exit Function
    End If
    modCalculo.AplicarFormulas destino, FILA_1, FILA_FIN
    modCalculo.FormatearCostos destino, FILA_1, FILA_FIN
    modCalculo.EscribirTotales destino
    CopiarMesAnterior = True
    Exit Function
fallo:
    DestinoEnBlanco destino
    CopiarMesAnterior = False
End Function

Private Function FilaDeCodigoEn(ByVal ws As Worksheet, ByVal codigo As String) As Long
    Dim f As Long
    Dim clave As String
    clave = NormalizaClave(codigo)
    For f = FILA_1 To FILA_FIN
        If NormalizaClave(ws.Cells(f, COL_CODIGO).Value2) = clave Then
            FilaDeCodigoEn = f
            Exit Function
        End If
    Next f
    FilaDeCodigoEn = 0
End Function

Private Sub DestinoEnBlanco(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:=CLAVE_PROTECCION
    ws.Range(ws.Cells(FILA_1, COL_CODIGO), ws.Cells(FILA_FIN, COL_ESTADO)).ClearContents
    ws.Cells.Locked = True
    On Error GoTo 0
End Sub

Private Function TieneDatosDebajo(ByVal ws As Worksheet) As Boolean
    Dim f As Long
    For f = FILA_1 To FILA_FIN
        If Trim$(CStr(ws.Cells(f, COL_CODIGO).Value2)) <> "" Then
            TieneDatosDebajo = True
            Exit Function
        End If
        If EsNumeroReal(ws.Cells(f, COL_CANT_ENT).Value2) Then
            If CDbl(ws.Cells(f, COL_CANT_ENT).Value2) <> 0 Then
                TieneDatosDebajo = True
                Exit Function
            End If
        End If
    Next f
End Function

' Asistente de apertura de enero: habilita C/D solo en enero, con
' responsable, soporte y motivo; al confirmar quedan fijas y bloqueadas.
Public Sub ConfirmarAperturaEnero()
    Dim ws As Worksheet
    Dim f As Long
    Dim pendientes As String
    Dim idOperacion As String
    On Error GoTo fallo
    If Reentrada Then Exit Sub
    Entrar
    If MesAbierto() <> 0 Then
        Salir
        Err.Raise vbObjectError + 530, "ConfirmarAperturaEnero", _
            "Ya hay un mes abierto: " & HojaMes(MesAbierto()) & "."
    End If
    If EstadoMes(1) <> EST_APERTURA_PEND Then
        Salir
        Exit Sub
    End If
    Set ws = Libro().Sheets(HojaMes(1))
    pendientes = RevisarEneroPendientes(ws)
    If pendientes <> "" Then
        Salir
        Err.Raise vbObjectError + 531, "ConfirmarAperturaEnero", _
            "Revise antes de confirmar:" & vbCrLf & pendientes
    End If
    idOperacion = "APERTURA-ENERO-" & Libro().Sheets(HOJA_CONFIG).Range(CFG_ANIO).Value2
    If modBitacora.YaRegistrado(idOperacion) Then
        Salir
        Exit Sub
    End If
    ws.Protect Password:=CLAVE_PROTECCION, UserInterfaceOnly:=True
    FijarEstadoMes 1, EST_ABIERTO, Now, "", _
        Libro().Sheets(HOJA_CONFIG).Range(CFG_RESPONSABLE).Value2, "", "", _
        "Apertura inicial documentada"
    modBitacora.Registrar "Confirmacion de apertura inicial", HojaMes(1), "OK", _
        "Asistente de enero: saldo inicial capturado por el responsable", "", "", _
        HojaMes(1), idOperacion
    DesbloquearINPUT ws
    Salir
    Exit Sub
fallo:
    modEstadoExcel.PropagarError "ConfirmarAperturaEnero", "modEstados"
    Salir
End Sub

' Pendientes de enero: ninguna fila de material puede quedar con C o D vacio.
Private Function RevisarEneroPendientes(ByVal ws As Worksheet) As String
    Dim f As Long
    Dim texto As String
    For f = FILA_1 To FILA_FIN
        If Trim$(CStr(ws.Cells(f, COL_CODIGO).Value2)) <> "" Then
            If EsVacio(ws.Cells(f, COL_CANT_INI).Value2) Or _
               EsVacio(ws.Cells(f, COL_VAL_INI).Value2) Then
                texto = texto & "- " & ws.Cells(f, COL_CODIGO).Value2 & _
                    ": saldo inicial incompleto (C/D)." & vbCrLf
            End If
        End If
    Next f
    RevisarEneroPendientes = texto
End Function

' ---------------------------------------------------------------------------
' Cierre
' ---------------------------------------------------------------------------
Public Function CerrarMes(ByVal indice As Long, ByVal responsable As String, _
                         ByVal motivo As String) As String
Dim problemas As String
    Dim idOperacion As String
    Dim ws As Worksheet
    Dim identificador As String
    Dim huella As String
    Dim rutaCopia As String
    Dim numero As Long
    Dim detalle As String

    ' La evidencia registrada manda sobre el estado: si el cierre ya quedo
    ' asentado, repetirlo es idempotente y no se informa como mes no abierto.
    If modBitacora.YaRegistrado("CIERRE-" & HojaMes(indice)) Then
        CerrarMes = "- El cierre de " & HojaMes(indice) & " ya estaba registrado."
        Exit Function
    End If
    problemas = PuedeCerrar(indice)
    If problemas <> "" Then
        CerrarMes = problemas
        Exit Function
    End If
    modEstadoExcel.CapturarEstado
    modEstadoExcel.SinPantalla
    On Error GoTo fallo
    Set ws = Libro().Sheets(HojaMes(indice))
    Application.Calculate
    problemas = modValidacion.ValidarMes(indice)
    If problemas <> "" Then
        CerrarMes = problemas
        GoTo restaura
    End If

    ' Respaldo verificable antes de cambiar el estado
    rutaCopia = modRespaldo.CrearRespaldo("pre-cierre-" & HojaMes(indice))
    If rutaCopia = "" Then
        CerrarMes = "- No se pudo crear el respaldo verificable. El cierre se detiene."
        GoTo restaura
    End If
    huella = SHA256DeArchivo(rutaCopia)

    identificador = "SNAP-" & HojaMes(indice) & "-" & SelloTiempo()
    modInstantanea.GrabarInstantanea indice, identificador
    problemas = modInstantanea.CuadraInstantanea(indice, identificador)
    If problemas <> "" Then
        Err.Raise vbObjectError + 533, "CerrarMes", _
            "La instantanea no concilio con la hoja:" & vbCrLf & problemas
    End If
    BloquearMesTodo ws

    FijarEstadoMes indice, EST_CERRADO, FechaAperturaMes(indice), Now, responsable, _
        identificador, huella, motivo
    modBitacora.Registrar "Cierre de mes", HojaMes(indice), "OK", _
        "Totales, validaciones, detalle e instantanea verificados", "", "", _
        HojaMes(indice), "CIERRE-" & HojaMes(indice)
    Libro().Save
    modRespaldo.CrearRespaldo "post-cierre-" & HojaMes(indice)
    CerrarMes = ""
    GoTo restaura
fallo:
    numero = Err.Number
    detalle = Err.Description
    Err.Clear
    On Error Resume Next
    FijarEstadoMes indice, EST_INCIDENCIA, "", "", responsable, "", "", _
        "Fallo en cierre: " & detalle
    On Error GoTo 0
    CerrarMes = "- Fallo en el cierre de " & HojaMes(indice) & ": " & detalle & vbCrLf
restaura:
    modEstadoExcel.RestaurarEstado
    If numero <> 0 Then
        Err.Raise numero, "modEstados.CerrarMes", detalle
    End If
End Function

' ---------------------------------------------------------------------------
' Bloqueos
' ---------------------------------------------------------------------------
Public Sub DesbloquearINPUT(ByVal hoja As Worksheet)
    On Error Resume Next
    hoja.Unprotect Password:=CLAVE_PROTECCION
    hoja.Cells.Locked = True
    hoja.Range(hoja.Cells(FILA_1, COL_CANT_ENT), hoja.Cells(FILA_FIN, COL_CANT_ENT)).Locked = False
    hoja.Range(hoja.Cells(FILA_1, COL_VAL_ENT), hoja.Cells(FILA_FIN, COL_VAL_ENT)).Locked = False
    hoja.Range(hoja.Cells(FILA_1, COL_CANT_VTA), hoja.Cells(FILA_FIN, COL_CANT_VTA)).Locked = False
    hoja.Range(hoja.Cells(FILA_1, COL_CANT_OTR), hoja.Cells(FILA_FIN, COL_CANT_OTR)).Locked = False
    hoja.Protect Password:=CLAVE_PROTECCION, UserInterfaceOnly:=True, _
        DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowFiltering:=True
    On Error GoTo 0
End Sub

Public Sub BloquearMesTodo(ByVal hoja As Worksheet)
    On Error Resume Next
    hoja.Unprotect Password:=CLAVE_PROTECCION
    hoja.Cells.Locked = True
    hoja.Protect Password:=CLAVE_PROTECCION, UserInterfaceOnly:=True, _
        DrawingObjects:=False, Contents:=True, Scenarios:=False, AllowFiltering:=True
    On Error GoTo 0
End Sub

' Al guardar siempre en modo seguro: todas las celdas bloqueadas y sin
' UserInterfaceOnly, para que un archivo abierto con macros deshabilitadas no
' permita entradas. Al abrir con macros se reactiva solo el mes abierto.
' Quita la proteccion de todas las hojas para poder reaplicar formatos y
' fijar el formato Texto de las columnas de codigo. El llamador debe cerrar
' con ProtegerTodoParaGuardar; si algo falla a mitad de camino, ese cierre
' se vuelve a intentar en el manejador de error.
Public Sub DesprotegerTodoParaEdicion()
    Dim ws As Worksheet
    For Each ws In Libro().Worksheets
        If ws.ProtectContents Or ws.ProtectDrawingObjects Then
            ws.Unprotect Password:=CLAVE_PROTECCION
        End If
    Next ws
End Sub

Public Sub ProtegerTodoParaGuardar()
    Dim ws As Worksheet
    On Error GoTo fallo
    For Each ws In Libro().Worksheets
        ws.Unprotect Password:=CLAVE_PROTECCION
        ws.Cells.Locked = True
        If ws.Name = HOJA_ESTADOS Or ws.Name = HOJA_SNAP Then
            ws.Protect Password:=CLAVE_PROTECCION, UserInterfaceOnly:=False, _
                DrawingObjects:=True, Contents:=True, Scenarios:=False
        Else
            ws.Protect Password:=CLAVE_PROTECCION, UserInterfaceOnly:=False, _
                DrawingObjects:=True, Contents:=True, Scenarios:=False, AllowFiltering:=True
        End If
    Next ws
    Exit Sub
fallo:
    Err.Raise Err.Number, "modEstados.ProtegerTodoParaGuardar", Err.Description
End Sub


' C/D vacio no es cero: una fila con codigo y saldo inicial vacio bloquea la
' apertura del mes (art. de validacion: vacio o ausencia no equivale a cero).
Private Function RevisarAperturaDestino(ByVal indice As Long) As String
    Dim ws As Worksheet
    Dim f As Long
    Dim texto As String
    Set ws = Libro().Sheets(HojaMes(indice))
    For f = FILA_1 To FILA_FIN
        If Trim$(CStr(ws.Cells(f, COL_CODIGO).Value2)) <> "" Then
            If EsVacio(ws.Cells(f, COL_CANT_INI).Value2) Or _
               EsVacio(ws.Cells(f, COL_VAL_INI).Value2) Then
                texto = texto & "- Fila " & f & ": " & ws.Cells(f, COL_CODIGO).Value2 & _
                    " tiene saldo inicial vacio (C/D); vacio no es cero." & vbCrLf
            End If
        End If
    Next f
    RevisarAperturaDestino = texto
End Function
