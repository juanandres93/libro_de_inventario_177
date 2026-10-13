Attribute VB_Name = "modUI"
Option Explicit

' Botones visibles (FORMULARIOS) y punto de entrada unico. Cada boton pide
' confirmacion, muestra periodo, totales y advertencias antes de escribir, y
' repite sin duplicar efectos.

' Estado del modo de pruebas: tiene que vivir en el bloque de declaraciones
' del modulo (antes del primer procedimiento) para que el compilador lo vea.
Public PruebaSilencioso As Boolean
Public UltimoAviso As String

Public Sub BtnAyudaMenu()
    On Error GoTo fallo
    MenuAyuda
    Exit Sub
fallo:
    Aviso Err.Description, vbExclamation, APP_NOMBRE
End Sub

Public Sub MenuAyuda()
    Aviso APP_NOMBRE & " " & APP_VERSION & vbCrLf & _
        "Metodo: " & APP_METODO & vbCrLf & _
        "Ejercicio: " & HojaMes(1) & vbCrLf & _
        "Estado: " & modConfiguracion.EstadoArchivo() & vbCrLf & _
        "Mes abierto: " & TextoMesAbierto() & vbCrLf & _
        "Ultimo mes cerrado: " & TextoUltimoCerrado() & vbCrLf & vbCrLf & _
        AvisoArt177() & vbCrLf & vbCrLf & _
        "Fuente del metodo: " & APP_URL, vbInformation, APP_NOMBRE
End Sub

Public Function TextoMesAbierto() As String
    Dim indice As Long
    indice = MesAbierto()
    If indice = 0 Then
        TextoMesAbierto = "ninguno"
    Else
        TextoMesAbierto = HojaMes(indice)
    End If
End Function

Public Function TextoUltimoCerrado() As String
    Dim indice As Long
    indice = MesCerradoHasta()
    If indice = 0 Then
        TextoUltimoCerrado = "ninguno"
    Else
        TextoUltimoCerrado = HojaMes(indice)
    End If
End Function

' ---------------------------------------------------------------------------
' Apertura
' ---------------------------------------------------------------------------
Public Sub BtnConfirmarAperturaEnero()
    If MesAbierto() <> 0 Then
        Aviso "Ya hay un mes abierto: " & TextoMesAbierto() & ".", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    If EstadoMes(1) <> EST_APERTURA_PEND Then
        Aviso "Enero no esta en APERTURA PENDIENTE.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    On Error GoTo fallo
    modEstados.ConfirmarAperturaEnero
    Aviso "Enero quedó abierto y documentado.", vbInformation, APP_NOMBRE
    Exit Sub
fallo:
    Aviso Err.Description, vbExclamation, APP_NOMBRE
End Sub

Public Sub BtnAbrirMesSiguiente()
    Dim indice As Long
    Dim problemas As String
    Dim siguiente As Long
    On Error GoTo fallo
    indice = MesAbierto()
    If indice <> 0 Then
        Aviso "Cierre primero " & HojaMes(indice) & ". Solo puede haber un mes abierto.", _
            vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    siguiente = MesCerradoHasta() + 1
    If siguiente > 12 Then
        Aviso "El ejercicio esta completo: diciembre esta CERRADO.", vbInformation, APP_NOMBRE
        Exit Sub
    End If
    problemas = modEstados.AbrirMes(siguiente)
    If problemas <> "" Then
        Aviso "No se puede abrir " & HojaMes(siguiente) & ":" & vbCrLf & problemas, _
            vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    Aviso HojaMes(siguiente) & " quedó abierto con el maestro y el saldo de cierre del mes anterior." & _
        vbCrLf & "Capture entradas, ventas y salidas en E, F, G e I.", vbInformation, APP_NOMBRE
    Exit Sub
fallo:
    Aviso Err.Description, vbExclamation, APP_NOMBRE
End Sub

' ---------------------------------------------------------------------------
' Cierre
' ---------------------------------------------------------------------------
Public Sub BtnCerrarMesActual()
    Dim indice As Long
    Dim problemas As String
    Dim ws As Worksheet
    Dim responsable As String
    Dim mensaje As String
    Dim respuesta As VbMsgBoxResult
    On Error GoTo fallo
    indice = MesAbierto()
    If indice = 0 Then
        Aviso "No hay ningun mes abierto.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    Set ws = Libro().Sheets(HojaMes(indice))
    Application.Calculate
    problemas = PuedeCerrar(indice)
    mensaje = "Periodo: " & HojaMes(indice) & vbCrLf & _
        "Estado: " & EstadoMes(indice) & vbCrLf & vbCrLf & _
        "Totales del mes (Bs.):" & vbCrLf & _
        "  Saldo inicial: " & Format$(ws.Cells(FILA_TOTAL, COL_VAL_INI).Value2, FMT_MONEDA) & vbCrLf & _
        "  Entradas: " & Format$(ws.Cells(FILA_TOTAL, COL_VAL_ENT).Value2, FMT_MONEDA) & vbCrLf & _
        "  Ventas: " & Format$(ws.Cells(FILA_TOTAL, COL_VAL_VTA).Value2, FMT_MONEDA) & vbCrLf & _
        "  Otros: " & Format$(ws.Cells(FILA_TOTAL, COL_VAL_OTR).Value2, FMT_MONEDA) & vbCrLf & _
        "  Saldo final: " & Format$(ws.Cells(FILA_TOTAL, COL_VAL_FIN).Value2, FMT_MONEDA) & vbCrLf & vbCrLf
    If problemas <> "" Then
        Aviso mensaje & "NO SE PUEDE CERRAR. Hallazgos pendientes:" & vbCrLf & problemas, _
            vbCritical, APP_NOMBRE
        Exit Sub
    End If
    If Not modConfiguracion.AutorizacionConfirmada() Then
        Aviso mensaje & "Falta confirmar la autorizacion del art. 177 en CONFIGURACION.", _
            vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    respuesta = ConfirmarAccion(mensaje & _
        "Se guardara una instantanea, un respaldo verificable y el estado pasara a CERRADO." & _
        vbCrLf & "¿Confirma el cierre?", vbQuestion + vbYesNo, APP_NOMBRE)
    If respuesta <> vbYes Then Exit Sub
    responsable = CapturarTexto("Responsable que cierra " & HojaMes(indice) & ":", _
        "Cierre de mes", CStr(Libro().Sheets(HOJA_CONFIG).Range(CFG_RESPONSABLE).Value2))
    If responsable = "" Then Exit Sub
    problemas = CerrarMes(indice, responsable, "Cierre manual desde el boton")
    If problemas <> "" Then
        Aviso problemas, vbCritical, APP_NOMBRE
    Else
        Aviso HojaMes(indice) & " quedó CERRADO con evidencia en BITACORA y respaldo local.", _
            vbInformation, APP_NOMBRE
    End If
    Exit Sub
fallo:
    Aviso Err.Description, vbCritical, APP_NOMBRE
End Sub

' ---------------------------------------------------------------------------
' Validacion, catalogo, detalle, resumen, respaldo y nuevo anio
' ---------------------------------------------------------------------------
Public Sub BtnValidarMes()
    On Error GoTo fallo
    If MesAbierto() = 0 Then
        Aviso "No hay ningun mes abierto.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    modValidacion.ReportarErrores MesAbierto()
    Exit Sub
fallo:
    Aviso Err.Description, vbExclamation, APP_NOMBRE
End Sub

Public Sub BtnAltaMaterial()
    On Error GoTo fallo
    modCatalogo.AltaModificarMaterial
    Exit Sub
fallo:
    Aviso Err.Description, vbExclamation, APP_NOMBRE
End Sub

Public Sub BtnDetalleOtros()
    On Error GoTo fallo
    modDetalle.AgregarDetalle
    Exit Sub
fallo:
    Aviso Err.Description, vbExclamation, APP_NOMBRE
End Sub

Public Sub BtnResumenAnual()
    On Error GoTo fallo
    modResumen.ConstruirResumen
    Aviso "Resumen anual construido. Corte: " & TextoUltimoCerrado() & " como definitive.", _
        vbInformation, APP_NOMBRE
    Exit Sub
fallo:
    Aviso Err.Description, vbExclamation, APP_NOMBRE
End Sub

Public Sub BtnRespaldo()
    Dim ruta As String
    On Error GoTo fallo
    If Len(Libro().FullName) = 0 Then
        Aviso "Guarde el archivo antes de respaldar.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    ruta = modRespaldo.CrearRespaldo("manual")
    If ruta = "" Then
        Aviso "No se pudo crear el respaldo verificable.", vbExclamation, APP_NOMBRE
    Else
        Aviso "Respaldo creado:" & vbCrLf & ruta & vbCrLf & vbCrLf & _
            "SHA-256: " & SHA256DeArchivo(ruta), vbInformation, APP_NOMBRE
    End If
    Exit Sub
fallo:
    Aviso Err.Description, vbExclamation, APP_NOMBRE
End Sub

Public Sub BtnImportarDiciembre()
    On Error GoTo fallo
    modImportacion.ImportarDiciembre
    Exit Sub
fallo:
    Aviso Err.Description, vbExclamation, APP_NOMBRE
End Sub

Public Sub BtnConfirmarImportacion()
    On Error GoTo fallo
    modImportacion.ConfirmarImportacionEnero
    Exit Sub
fallo:
    Aviso Err.Description, vbExclamation, APP_NOMBRE
End Sub

Public Sub BtnNuevoAnio()
    Dim ruta As String
    On Error GoTo fallo
    If ConfirmarAccion("Se creará un archivo NUEVO para el siguiente ejercicio con el maestro, el VBA " & _
        "y el saldo de diciembre. El archivo actual no se modifica." & vbCrLf & _
        "¿Continuar?", vbQuestion + vbYesNo, APP_NOMBRE) <> vbYes Then Exit Sub
    ruta = modNuevoAnio.CrearNuevoAnio()
    If ruta <> "" Then
        Aviso "Nuevo libro creado y verificado:" & vbCrLf & ruta, vbInformation, APP_NOMBRE
    End If
    Exit Sub
fallo:
    Aviso Err.Description, vbCritical, APP_NOMBRE
End Sub

Public Sub BtnValidarConfiguracion()
    Dim problemas As String
    On Error GoTo fallo
    problemas = modConfiguracion.ValidarConfiguracion()
    If problemas <> "" Then
        Aviso "CONFIGURACION incompleta:" & vbCrLf & problemas, vbExclamation, APP_NOMBRE
    Else
        problemas = modCatalogo.ValidarCatalogo()
        If problemas <> "" Then
            Aviso "CATALOGO con hallazgos:" & vbCrLf & problemas, vbExclamation, APP_NOMBRE
        Else
            Aviso "Configuracion y catalogo correctos.", vbInformation, APP_NOMBRE
        End If
    End If
    Exit Sub
fallo:
    Aviso Err.Description, vbExclamation, APP_NOMBRE
End Sub

Public Sub BtnCapturarConfiguracion()
    On Error GoTo fallo
    modConfiguracion.CapturarDatosIniciales
    Exit Sub
fallo:
    Aviso Err.Description, vbExclamation, APP_NOMBRE
End Sub

' ---------------------------------------------------------------------------
' Formatos numericos: mascara de cantidades elegida en CONFIGURACION e
' importes en Bs. siempre a 2 decimales. Se llama con las hojas sin proteger
' y el llamador vuelve a proteger justo despues (Workbook_Open y
' Workbook_BeforeSave). Los errores no se tragan: interrumpen la apertura o
' cancelan el guardado antes de dejar el libro a medias.
' ---------------------------------------------------------------------------
Public Sub AplicarFormatosNumericos()
    modCatalogo.AsegurarFormatoCodigos
    modCalculo.AplicarMascaraCantidades
    modResumen.AplicarFormatos
    modDetalle.AplicarFormatos
    modInstantanea.AplicarFormatos
End Sub

' ---------------------------------------------------------------------------
' Aviso: unico punto de mensajes. En modo de pruebas no abre ventanas y
' registra el texto para que el autocontrol pueda revisarlo.
' ---------------------------------------------------------------------------

' Interruptor del modo de pruebas para la automatizacion (seccion 0.6.4):
' silencia Aviso y CapturarTexto, de modo que cada boton se puede ejecutar sin
' que la prueba quede esperando una ventana. Por defecto esta apagado y no
' altera el funcionamiento normal del libro.
Public Sub ModoPrueba(Optional ByVal activo As Boolean = True)
    PruebaSilencioso = activo
End Sub

' Captura de texto: en modo de pruebas devuelve el valor por defecto y no
' abre el dialogo, para que la automatizacion no quede esperando una ventana.
Public Function CapturarTexto(ByVal prompt As String, ByVal titulo As String, _
                              Optional ByVal defecto As String = "") As String
    If PruebaSilencioso Then
        CapturarTexto = Trim$(defecto)
    Else
        CapturarTexto = Trim$(InputBox(prompt, titulo, defecto))
    End If
End Function

Public Function ConfirmarAccion(ByVal texto As String, _
                                Optional ByVal botones As VbMsgBoxStyle = vbQuestion + vbYesNo, _
                                Optional ByVal titulo As String = "") As VbMsgBoxResult
    If titulo = "" Then titulo = APP_NOMBRE
    ConfirmarAccion = Aviso(texto, botones, titulo)
End Function

Public Function Aviso(ByVal texto As String, _
                      Optional ByVal botones As VbMsgBoxStyle = vbOKOnly, _
                      Optional ByVal titulo As String = "") As VbMsgBoxResult
    UltimoAviso = texto
    If PruebaSilencioso Then
        If botones = vbYesNo Or botones = vbQuestion + vbYesNo Or _
           botones = vbYesNoCancel Or botones = vbQuestion + vbYesNoCancel Then
            Aviso = vbNo
        Else
            Aviso = vbOK
        End If
        Exit Function
    End If
    Aviso = MsgBox(texto, botones, titulo)
End Function