Attribute VB_Name = "modConfiguracion"
Option Explicit

' CONFIGURACION: identidad del archivo, ejercicio, metodo y estado.
' Se captura al crear el archivo y queda bloqueada despues.

Public Sub InicializarConfiguracion(ByVal anio As Long)
    Dim ws As Worksheet
    Set ws = Libro().Sheets(HOJA_CONFIG)
    ws.Unprotect Password:=CLAVE_PROTECCION
    ws.Range(CFG_ENTIDAD).Value = ""
    ws.Range(CFG_RIF).Value = ""
    ws.Range(CFG_DOMICILIO).Value = ""
    ws.Range(CFG_RESPONSABLE).Value = ""
    ws.Range(CFG_ANIO).Value = anio
    ws.Range(CFG_MONEDA).Value = APP_MONEDA
    ws.Range(CFG_VERSION).Value = APP_VERSION
    ws.Range(CFG_METODO).Value = APP_METODO
    ws.Range(CFG_MASCARA_CANT).Value = FMT_CANT
    ws.Range(CFG_ESTADO).Value = "SIN CAPTURAR"
    ws.Range(CFG_LIMITE).Value = "Materiales " & FILA_1 & " a " & FILA_FIN & _
        " (" & LIMITE_MATERIALES & "); totalizadores en filas " & FILA_TOTAL & _
        " y " & FILA_TOTAL_FILT
    ws.Range(CFG_AUTORIZACION).Value = "PENDIENTE DE CONFIRMAR POR EL RESPONSABLE"
    ws.Range(CFG_COSTO_CERO).Value = "NO ACEPTADO"
    ws.Range(CFG_CLAVE).Value = "Clave de proteccion local: " & CLAVE_PROTECCION & _
        " (control de accidentes, no seguridad de acceso)"
    ws.Protect Password:=CLAVE_PROTECCION, UserInterfaceOnly:=True
End Sub

' Captura guiada de la identidad del archivo. Solo habilita la hoja mientras
' dura la captura y vuelve a congelarla al terminar.
Public Sub CapturarDatosIniciales()
    Dim ws As Worksheet
    Dim entidad As String, rif As String, domicilio As String, responsable As String
    Dim autorizacion As String, costoCero As String
    Dim problemas As String
    On Error GoTo fallo
    Set ws = Libro().Sheets(HOJA_CONFIG)
    entidad = modUI.CapturarTexto("Entidad o razon social (obligatoria).", "Configuracion")
    If entidad = "" Then Exit Sub
    rif = modUI.CapturarTexto("RIF (obligatorio).", "Configuracion")
    If rif = "" Then Exit Sub
    domicilio = modUI.CapturarTexto("Domicilio fiscal.", "Configuracion")
    responsable = modUI.CapturarTexto("Responsable del libro (obligatorio).", "Configuracion")
    If responsable = "" Then Exit Sub
    If modUI.ConfirmarAccion("Art. 177 del Reglamento (D. 2.507, G.O.E. 5.662 del 24-09-2003): el " & _
        "registro por medios magneticos requiere autorizacion de la Administracion " & _
        "Tributaria. Este libro no la otorga ni la sustituye." & vbCrLf & vbCrLf & _
        "Confirma que la autorizacion fue verificada por el responsable?", _
        vbQuestion + vbYesNo, APP_NOMBRE) = vbYes Then
        autorizacion = "SI - verificada por el responsable el " & Format$(Date, "dd/mm/yyyy")
    Else
        autorizacion = "NO VERIFICADA - PENDIENTE"
    End If
    If modUI.ConfirmarAccion("Acepta explicitamente valued materials con costo promedio cero? " & _
        "(sin esto, esas filas bloquean el cierre)", vbQuestion + vbYesNo, APP_NOMBRE) = vbYes Then
        costoCero = "ACEPTADO por " & responsable & " el " & Format$(Date, "dd/mm/yyyy")
    Else
        costoCero = "NO ACEPTADO"
    End If
    ws.Unprotect Password:=CLAVE_PROTECCION
    ws.Range(CFG_ENTIDAD).Value = entidad
    ws.Range(CFG_RIF).Value = rif
    ws.Range(CFG_DOMICILIO).Value = domicilio
    ws.Range(CFG_RESPONSABLE).Value = responsable
    ws.Range(CFG_AUTORIZACION).Value = autorizacion
    ws.Range(CFG_COSTO_CERO).Value = costoCero
    ws.Range(CFG_ESTADO).Value = "CAPTURADO"
    problemas = ValidarConfiguracion()
    ws.Protect Password:=CLAVE_PROTECCION, UserInterfaceOnly:=True
    If problemas <> "" Then
        Err.Raise vbObjectError + 525, "CapturarDatosIniciales", _
            "La configuracion sigue incompleta:" & vbCrLf & problemas
    End If
    modBitacora.Registrar "Captura de configuracion", HOJA_CONFIG, "OK", _
        "Identidad del ejercicio capturada", "", "", "", ""
    Exit Sub
fallo:
    On Error Resume Next
    ws.Protect Password:=CLAVE_PROTECCION, UserInterfaceOnly:=True
    On Error GoTo 0
    Err.Raise Err.Number, "modConfiguracion.CapturarDatosIniciales", Err.Description
End Sub

' Devuelve una cadena con los problemas de configuracion; vacia si esta bien.
Public Function ValidarConfiguracion() As String
    Dim ws As Worksheet
    Dim problemas As String
    Dim anio As Variant
    Dim mascara As String
    Set ws = Libro().Sheets(HOJA_CONFIG)
    If EsVacio(ws.Range(CFG_ENTIDAD).Value2) Then problemas = problemas & _
        "- Entidad vacia." & vbCrLf
    If EsVacio(ws.Range(CFG_RIF).Value2) Then problemas = problemas & _
        "- RIF vacio." & vbCrLf
    If EsVacio(ws.Range(CFG_RESPONSABLE).Value2) Then problemas = problemas & _
        "- Responsable vacio." & vbCrLf
    anio = ws.Range(CFG_ANIO).Value2
    If EsVacio(anio) Or Not EsNumeroReal(anio) Then
        problemas = problemas & "- Anio no numerico." & vbCrLf
    ElseIf CLng(anio) < 2000 Or CLng(anio) > 2100 Then
        problemas = problemas & "- Anio fuera del rango tecnico 2000-2100." & vbCrLf
    End If
    If UCase$(Trim$(CStr(ws.Range(CFG_METODO).Value2))) <> APP_METODO Then
        problemas = problemas & "- Metodo distinto de " & APP_METODO & "." & vbCrLf
    End If
    mascara = Trim$(CStr(ws.Range(CFG_MASCARA_CANT).Value2))
    If mascara <> FMT_CANT And mascara <> MASCARA_CANT_SIN_DEC Then
        problemas = problemas & "- Mascara de cantidades invalida: use " & FMT_CANT & _
            " (2 decimales) o " & MASCARA_CANT_SIN_DEC & " (sin decimales)." & vbCrLf
    End If
    ValidarConfiguracion = problemas
End Function

' El archivo solo es operable si la configuracion esta completa y el mes
' abierto es coherente con el ejercicio capturado.
Public Function ConfiguracionOperativa() As Boolean
    ConfiguracionOperativa = (ValidarConfiguracion() = "") And _
                            (IndiceDeHoja(HojaMes(1)) > 0)
End Function

Public Sub CongelarConfiguracion()
    Dim ws As Worksheet
    Set ws = Libro().Sheets(HOJA_CONFIG)
    ws.Protect Password:=CLAVE_PROTECCION, UserInterfaceOnly:=True, _
        DrawingObjects:=False, Contents:=True, Scenarios:=False
End Sub

Public Sub DescongelarConfiguracion()
    Libro().Sheets(HOJA_CONFIG).Unprotect Password:=CLAVE_PROTECCION
End Sub

Public Function EstadoArchivo() As String
    On Error GoTo sinEstado
    EstadoArchivo = CStr(Libro().Sheets(HOJA_CONFIG).Range(CFG_ESTADO).Value2)
    Exit Function
sinEstado:
    EstadoArchivo = "desconocido"
End Function

Public Sub ActualizarEstadoArchivo(ByVal estado As String)
    Dim numero As Long
    Dim detalle As String
    On Error GoTo fallo
    DescongelarConfiguracion
    Libro().Sheets(HOJA_CONFIG).Range(CFG_ESTADO).Value = estado
    CongelarConfiguracion
    Exit Sub
fallo:
    numero = Err.Number
    detalle = Err.Description
    Err.Clear
    On Error Resume Next
    CongelarConfiguracion
    On Error GoTo 0
    Err.Raise numero, "modConfiguracion.ActualizarEstadoArchivo", detalle
End Sub

Public Function AceptacionCostoCero() As Boolean
    On Error GoTo sinAceptacion
    Dim texto As String
    texto = UCase$(CStr(Libro().Sheets(HOJA_CONFIG).Range(CFG_COSTO_CERO).Value2))
    ' "NO ACEPTADO" tambien contiene la palabra ACEPTADO: si solo se buscara
    ' esa subcadena, la negacion pasaria por aceptacion documentada.
    AceptacionCostoCero = (InStr(1, texto, "ACEPTADO") > 0) And _
                          (InStr(1, texto, "NO ACEPTADO") = 0)
    Exit Function
sinAceptacion:
    AceptacionCostoCero = False
End Function

Public Function AutorizacionConfirmada() As Boolean
    On Error GoTo sinAut
    Dim texto As String
    texto = CStr(Libro().Sheets(HOJA_CONFIG).Range(CFG_AUTORIZACION).Value2)
    AutorizacionConfirmada = (InStr(1, texto, "SI", vbTextCompare) > 0) And _
                             (InStr(1, texto, "PENDIENTE", vbTextCompare) = 0)
    Exit Function
sinAut:
    AutorizacionConfirmada = False
End Function