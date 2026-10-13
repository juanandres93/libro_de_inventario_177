Attribute VB_Name = "modPruebas"
Option Explicit

' Autocontrol del libro: pruebas T01 a T28 de la especificacion.
' Se ejecuta SOLO sobre una copia de pruebas. Cada prueba devuelve una cadena
' vacia si pasa y el hallazgo si falla. El resultado se escribe en un archivo
' de informe para que quede evidencia fuera de la sesion.

Private mPasadas As Long
Private mFallidas As Long
Private mInforme As String
Private mPruebaActual As String
Private mDonde As String

' Marca el paso en curso: si un control termina en error, el mensaje dice en
' que punto fallo en lugar de devolver un ERROR sin contexto.
Private Sub Donde(ByVal etapa As String)
    mDonde = etapa
End Sub

' Deja constancia de que prueba esta corriendo: si la suite se cuelga, el
' archivo de progreso (fuera del libro) indica donde quedo.
Private Sub Marcar(ByVal id As String)
    Dim fso As Object
    Dim ts As Object
    mPruebaActual = id
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(Libro().Path & "\libinv_progreso.txt", 8, True)
    ts.Write MarcaTiempo() & " " & id & vbCrLf
    ts.Close
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' Orquestador
' ---------------------------------------------------------------------------
Public Sub Suite(Optional ByVal silencioso As Boolean = True)
    Dim r As String
    Dim evAntes As Boolean
    On Error Resume Next
    evAntes = Application.EnableEvents
    Err.Clear
    On Error GoTo interrumpida
    mPasadas = 0: mFallidas = 0: mInforme = ""
    modUI.PruebaSilencioso = silencioso
    Marcar "T01"
    Correr "T01", "Caso HT-0001 completo, cierre y apertura de febrero", EjecutarPrueba("T01")
    Marcar "T02"
    Correr "T02", "Duplicados por mayusculas/espacios, alta y pegado", EjecutarPrueba("T02")
    Marcar "T03"
    Correr "T03", "Codigos con ceros iniciales y catalogo reordenado", EjecutarPrueba("T03")
    Marcar "T04"
    Correr "T04", "Salidas mayores a disponibles", EjecutarPrueba("T04")
    Marcar "T05"
    Correr "T05", "Cantidad cero con valor no cero", EjecutarPrueba("T05")
    Marcar "T06"
    Correr "T06", "Stock cero valido, apertura vacia invalida, costo cero", EjecutarPrueba("T06")
    Marcar "T07"
    Correr "T07", "Agotamiento con venta y otros: residuo exacto", EjecutarPrueba("T07")
    Marcar "T08"
    Correr "T08", "Agotamiento solo ventas o solo otros", EjecutarPrueba("T08")
    Marcar "T09"
    Correr "T09", "Alta y baja con stock: no pierde existencias", EjecutarPrueba("T09")
    Marcar "T10"
    Correr "T10", "Descripcion nueva no reescribe meses cerrados", EjecutarPrueba("T10")
    Marcar "T11"
    Correr "T11", "Bloque pegado con negativo/texto/formula: rechazo total", EjecutarPrueba("T11")
    Marcar "T12"
    Correr "T12", "Formula rota o ausente: bloqueo de cierre", EjecutarPrueba("T12")
    Marcar "T13"
    Correr "T13", "Importacion de diciembre: fuente valida e invalidas", EjecutarPrueba("T13")
    Marcar "T14"
    Correr "T14", "Nuevo anio: saldo identico, origen intacto, VBA presente", EjecutarPrueba("T14")
    Marcar "T15"
    Correr "T15", "Repetir apertura/cierre/importacion: idempotente", EjecutarPrueba("T15")
    Marcar "T16"
    Correr "T16", "Respaldo fallido: no hay cierre falso", EjecutarPrueba("T16")
    Marcar "T17"
    Correr "T17", "Fallo a mitad de cierre: INCIDENCIA, no exito parcial", EjecutarPrueba("T17")
    Marcar "T18"
    Correr "T18", "Proteccion segura antes y despues de guardar", EjecutarPrueba("T18")
    Marcar "T19"
    Correr "T19", "Configuracion de separadores restaurada", EjecutarPrueba("T19")
    Marcar "T20"
    Correr "T20", "Resumen al corte: flujos sumados, stocks no sumados", EjecutarPrueba("T20")
    Marcar "T21"
    Correr "T21", "Unidades mixtas: cantidades no sumadas", EjecutarPrueba("T21")
    Marcar "T22"
    Correr "T22", "Detalle ausente o descuadrado: bloqueo", EjecutarPrueba("T22")
    Marcar "T23"
    Correr "T23", "Fracciones, centavos pequenos y error de redondeo", EjecutarPrueba("T23")
    Marcar "T24"
    Correr "T24", "Todos los botones y proyecto sin referencias rotas", EjecutarPrueba("T24")
    Marcar "T25"
    Correr "T25", "5.000 materiales: tiempos de apertura/validacion/cierre", EjecutarPrueba("T25")
    Marcar "T26"
    Correr "T26", "Exportacion del libro para revision visual", EjecutarPrueba("T26")
    Marcar "T27"
    Correr "T27", "Archivo limpio: sin enlaces externos ni marcadores", EjecutarPrueba("T27")
    Marcar "T28"
    Correr "T28", "Trazabilidad normativa: cita norma, Gaceta y fecha", EjecutarPrueba("T28")

    modUI.PruebaSilencioso = False
    mInforme = mInforme & vbCrLf & "RESUMEN: pasadas=" & mPasadas & " fallidas=" & mFallidas & vbCrLf
    EscribirInforme
    ' El dialogo solo aparece si la suite se lanza a mano; en la automatizacion
    ' el informe del disco es la evidencia y no se puede bloquear la sesion.
    If Not silencioso Then
        MsgBox "Autocontrol terminado: " & mPasadas & " pasadas, " & mFallidas & _
            " fallidas." & vbCrLf & "Informe: " & RutaInforme(), vbInformation, APP_NOMBRE
    End If
    On Error Resume Next
    Application.EnableEvents = evAntes
    On Error GoTo 0
    Exit Sub
interrumpida:
    mInforme = mInforme & vbCrLf & "SUITE INTERRUMPIDA en " & mPruebaActual & _
        ": " & Err_() & vbCrLf
    modUI.PruebaSilencioso = False
    On Error Resume Next
    EscribirInforme
    Application.EnableEvents = evAntes
    On Error GoTo 0
End Sub

Private Sub Correr(ByVal id As String, ByVal nombre As String, ByVal resultado As String)
    If resultado = "" Then
        mPasadas = mPasadas + 1
        mInforme = mInforme & id & " PASA: " & nombre & vbCrLf
    Else
        mFallidas = mFallidas + 1
        mInforme = mInforme & id & " FALLA: " & nombre & vbCrLf & "      " & _
            Replace(Trim$(resultado), vbCrLf, vbCrLf & "      ") & vbCrLf
    End If
End Sub

Public Function RutaInforme() As String
    RutaInforme = Libro().Path & "\INFORME_AUTOCONTROL_T01_T28.txt"
End Function

Private Sub EscribirInforme()
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(RutaInforme(), True, False)
    ts.Write "AdaptaPro - Libro de inventarios " & APP_VERSION & " - autocontrol" & vbCrLf
    ts.Write "Fecha: " & MarcaTiempo() & " | Archivo: " & Libro().Name & vbCrLf
    ts.Write "Metodo: " & APP_METODO & " | " & APP_URL & vbCrLf & vbCrLf
    ts.Write mInforme
    ts.Close
End Sub

' Diagnostico: ejecuta una sola pieza para aislar errores de compilacion
' (VBA compila cada procedimiento al primer llamado). Solo para pruebas.
Public Function Probe(ByVal id As String) As String
    ' Un error sin manejador abriria el dialogo del VBE y bloquearia la
    ' sesion de automatizacion; aqui se convierte en una cadena de resultado.
    Dim wLimp As Worksheet
    Dim tLimp As Single
    Dim tCal As Single
    Dim evAntes As Boolean
    On Error Resume Next
    evAntes = Application.EnableEvents
    Err.Clear
    On Error GoTo fallo
    Select Case UCase$(Trim$(id))
        Case "MARCAR"
            Marcar "DIAG"
            Probe = "OK"
        Case "CORRER"
            Correr "DIAG", "diagnostico", ""
            Probe = "OK"
        Case "INFORME"
            EscribirInforme
            Probe = "OK"
        Case "RUTA"
            Probe = RutaInforme()
        Case "SALIR"
            modGlobales.Salir
            Probe = "OK"
        Case "DESPROT"
            DesprotegerTodo
            Probe = "OK"
        Case "NORMALIZA"
            Probe = NormalizaClave("  a b  ")
        Case "LIMPIAR"
            Probe = LimpiarTexto(" x ")
        Case "ANIODEL"
            Probe = CStr(AnioDelLibro())
        Case "INDICECAT"
            Probe = CStr(IndiceCatalogo().Count)
        Case "FILAENCAT"
            Probe = CStr(FilaEnCatalogo("M-0001"))
        Case "ULTIMAFILA"
            Probe = CStr(UltimaFilaUsada(modCatalogo.HojaCatalogo(), 1))
        Case "INVALIDAR2"
            modGlobales.InvalidarIndice
            Probe = "OK"
        Case "INVALIDAR"
            InvalidarIndice
            Probe = "OK"
        Case "CLRT"
            DesprotegerTodo
            Libro().Sheets(HOJA_CONFIG).Range(CFG_ANIO).Value = AnioDePrueba()
            Set wLimp = Libro().Sheets(HojaMes(1))
            tLimp = Timer
            wLimp.Range(wLimp.Cells(FILA_1, COL_CODIGO), _
                        wLimp.Cells(FILA_FIN, COL_ESTADO)).ClearContents
            Probe = "clear " & Format$(Timer - tLimp, "0.00") & "s"
        Case "CLRTMAN"
            DesprotegerTodo
            Libro().Sheets(HOJA_CONFIG).Range(CFG_ANIO).Value = AnioDePrueba()
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Set wLimp = Libro().Sheets(HojaMes(1))
            tLimp = Timer
            wLimp.Range(wLimp.Cells(FILA_1, COL_CODIGO), _
                        wLimp.Cells(FILA_FIN, COL_ESTADO)).ClearContents
            Probe = "clearman " & Format$(Timer - tLimp, "0.00") & "s"
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
        Case "CLRT3"
            DesprotegerTodo
            Dim wCfg As Worksheet
            Set wCfg = Libro().Sheets(HOJA_CONFIG)
            wCfg.Range(CFG_ANIO).Value = AnioDePrueba()
            wCfg.Range(CFG_ENTIDAD).Value = "EMPRESA DE PRUEBAS C.A."
            Set wLimp = Libro().Sheets(HOJA_CATALOGO)
            wLimp.Range(wLimp.Cells(FILA_1, 1), wLimp.Cells(FILA_FIN, 6)).ClearContents
            wLimp.Cells(FILA_1, 1).Value = "M-0001"
            wLimp.Cells(FILA_1, 2).Value = "Material uno"
            wLimp.Cells(FILA_1, 3).Value = "UND"
            wLimp.Cells(FILA_1, 4).Value = "Si"
            wLimp.Cells(FILA_1 + 1, 1).Value = "M-0002"
            wLimp.Cells(FILA_1 + 1, 2).Value = "Material dos"
            wLimp.Cells(FILA_1 + 1, 3).Value = "KG"
            wLimp.Cells(FILA_1 + 1, 4).Value = "Si"
            Set wLimp = Libro().Sheets(HojaMes(1))
            tLimp = Timer
            wLimp.Range(wLimp.Cells(FILA_1, COL_CODIGO), _
                        wLimp.Cells(FILA_FIN, COL_ESTADO)).ClearContents
            Probe = "clear3 " & Format$(Timer - tLimp, "0.00") & "s"
        Case "CLRT3M"
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            DesprotegerTodo
            Set wCfg = Libro().Sheets(HOJA_CONFIG)
            wCfg.Range(CFG_ANIO).Value = AnioDePrueba()
            wCfg.Range(CFG_ENTIDAD).Value = "EMPRESA DE PRUEBAS C.A."
            Set wLimp = Libro().Sheets(HOJA_CATALOGO)
            wLimp.Range(wLimp.Cells(FILA_1, 1), wLimp.Cells(FILA_FIN, 6)).ClearContents
            wLimp.Cells(FILA_1, 1).Value = "M-0001"
            wLimp.Cells(FILA_1, 2).Value = "Material uno"
            wLimp.Cells(FILA_1, 3).Value = "UND"
            wLimp.Cells(FILA_1, 4).Value = "Si"
            Set wLimp = Libro().Sheets(HojaMes(1))
            tLimp = Timer
            wLimp.Range(wLimp.Cells(FILA_1, COL_CODIGO), _
                        wLimp.Cells(FILA_FIN, COL_ESTADO)).ClearContents
            Probe = "clear3m " & Format$(Timer - tLimp, "0.00") & "s"
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
        Case "CALCTIEMPO"
            tCal = Timer
            Application.Calculate
            Probe = "calc " & Format$(Timer - tCal, "0.00") & "s"
        Case "PREPARAR"
            Preparar
            Probe = "OK"
        Case "RELENAR"
            modCalculo.RellenarFormulas Libro().Sheets(HojaMes(1))
            Probe = "OK"
        Case "TOTALES"
            modCalculo.EscribirTotales Libro().Sheets(HojaMes(1))
            Probe = "OK"
        Case "ESCRIBIRFILA"
            EscribirFila 1, "M-0001", 10, 1000, 5, 600, 4, 1
            Probe = "OK"
        Case "FILAENMES"
            Probe = CStr(FilaEnMesPublica(1, "M-0001"))
        Case "VFN"
            Probe = CStr(V(Libro().Sheets(HojaMes(1)), FILA_1, COL_CANT_FIN))
        Case "HALLAZGO"
            Probe = Hallazgo(True, "x")
        Case "ACUMULAR"
            Probe = Acumular("a", "b")
        Case "ABRIR"
            Probe = AbrirEnero()
        Case "CERRAR"
            Probe = Cerrar(1)
        Case "ANIO"
            Probe = CStr(AnioDePrueba())
        Case "SUITE"
            Suite True
            Probe = "OK"
        Case "EJEC"
            ' Compila el despachador y, con el, todas las pruebas T01-T28.
            Probe = EjecutarPrueba("NADA")
        Case "T01": Probe = T01()
        Case "T01DIAG": Probe = T01Diag()
        Case "DBG": Probe = DbgDiag(Trim$(Mid$(id, 4)))
        Case "DBGCAT": Probe = DbgDiag("CAT")
        Case "DBGT04": Probe = DbgDiag("T04")
        Case "DBGT02": Probe = DbgDiag("T02")
        Case "DBGT02A"
            Preparar
            Set wLimp = modCatalogo.HojaCatalogo()
            wLimp.Cells(FILA_1 + 3, 1).Value = "  m-0001  "
            wLimp.Cells(FILA_1 + 3, 2).Value = "Copia con espacios"
            wLimp.Cells(FILA_1 + 3, 3).Value = "UND"
            wLimp.Cells(FILA_1 + 3, 4).Value = "Si"
            InvalidarIndice
            Probe = "cat=[" & modCatalogo.ValidarCatalogo() & "]"
        Case "DBGDET"
            Preparar
            AgregarLineaDetalle 1, "M-0001", 15, "retiro", 1
            Set wLimp = modDetalle.HojaDetalle()
            Probe = "tipo=" & TypeName(wLimp.Cells(7, DET_FECHA).Value2) & _
                " val=[" & CStr(wLimp.Cells(7, DET_FECHA).Value2) & "]" & _
                " fmt=[" & wLimp.Cells(7, DET_FECHA).NumberFormat & "]" & _
                " isdate=" & IsDate(wLimp.Cells(7, DET_FECHA).Value2) & _
                " anio=" & AnioDePrueba() & _
                " val=[" & modDetalle.ValidarDetalleDeMes(1) & "]"
        Case "DBGCONT"
            Probe = "c1=" & Contiene("ERROR: salidas", "ERROR") & _
                " h=[" & Hallazgo(Not Contiene("ERROR: salidas", "ERROR"), "FALLA") & "]" & _
                " c2=" & Contiene("codigo duplicado (normalizado) m-0001", "duplicado")
        Case "DBGT02B"
            Preparar
            Set wLimp = modCatalogo.HojaCatalogo()
            wLimp.Cells(FILA_1 + 3, 1).Value = "  m-0001  "
            wLimp.Cells(FILA_1 + 3, 2).Value = "Copia con espacios"
            wLimp.Cells(FILA_1 + 3, 3).Value = "UND"
            wLimp.Cells(FILA_1 + 3, 4).Value = "Si"
            InvalidarIndice
            Probe = "c=" & Contiene(modCatalogo.ValidarCatalogo(), "duplicado") & _
                " h=[" & Hallazgo(Not Contiene(modCatalogo.ValidarCatalogo(), "duplicado"), "FALLA") & "]"
        Case "T09B": Probe = DbgDiag("T09B")
        Case "T23B": Probe = DbgDiag("T23B")
        Case "T27B": Probe = DbgDiag("T27B")
        Case "T02": Probe = T02()
        Case "T03": Probe = T03()
        Case "T04": Probe = T04()
        Case "T05": Probe = T05()
        Case "T06": Probe = T06()
        Case "T07": Probe = T07()
        Case "T08": Probe = T08()
        Case "T09": Probe = T09()
        Case "T10": Probe = T10()
        Case "T11": Probe = T11()
        Case "T12": Probe = T12()
        Case "T13": Probe = T13()
        Case "T14": Probe = T14()
        Case "T15": Probe = T15()
        Case "T16": Probe = T16()
        Case "T17": Probe = T17()
        Case "T18": Probe = T18()
        Case "T18DIAG": Probe = T18Diag()
        Case "T19": Probe = T19()
        Case "T20": Probe = T20()
        Case "T21": Probe = T21()
        Case "T22": Probe = T22()
        Case "T23": Probe = T23()
        Case "T24": Probe = T24()
        Case "T25": Probe = T25()
        Case "T26": Probe = T26()
        Case "T27": Probe = T27()
        Case "T28": Probe = T28()
        Case Else
            Probe = "ID DESCONOCIDO"
    End Select
    GoTo restaura
fallo:
    Probe = "ERROR " & Err.Number & ": " & Err.Description
restaura:
    On Error Resume Next
    Application.EnableEvents = evAntes
    On Error GoTo 0
End Function

' Ejecuta una prueba aislada: un fallo imprevisto se reporta como hallazgo de
' esa prueba y no interrumpe ni dialoga con la sesion.
Private Function EjecutarPrueba(ByVal id As String) As String
    On Error GoTo fallo
    Donde "inicio"
    Select Case id
        Case "T01": EjecutarPrueba = T01()
        Case "T02": EjecutarPrueba = T02()
        Case "T03": EjecutarPrueba = T03()
        Case "T04": EjecutarPrueba = T04()
        Case "T05": EjecutarPrueba = T05()
        Case "T06": EjecutarPrueba = T06()
        Case "T07": EjecutarPrueba = T07()
        Case "T08": EjecutarPrueba = T08()
        Case "T09": EjecutarPrueba = T09()
        Case "T10": EjecutarPrueba = T10()
        Case "T11": EjecutarPrueba = T11()
        Case "T12": EjecutarPrueba = T12()
        Case "T13": EjecutarPrueba = T13()
        Case "T14": EjecutarPrueba = T14()
        Case "T15": EjecutarPrueba = T15()
        Case "T16": EjecutarPrueba = T16()
        Case "T17": EjecutarPrueba = T17()
        Case "T18": EjecutarPrueba = T18()
        Case "T19": EjecutarPrueba = T19()
        Case "T20": EjecutarPrueba = T20()
        Case "T21": EjecutarPrueba = T21()
        Case "T22": EjecutarPrueba = T22()
        Case "T23": EjecutarPrueba = T23()
        Case "T24": EjecutarPrueba = T24()
        Case "T25": EjecutarPrueba = T25()
        Case "T26": EjecutarPrueba = T26()
        Case "T27": EjecutarPrueba = T27()
        Case "T28": EjecutarPrueba = T28()
        Case Else: EjecutarPrueba = "ID DESCONOCIDO: " & id
    End Select
    Exit Function
fallo:
    EjecutarPrueba = "ERROR " & Err.Number & ": " & Err.Description
End Function

' ---------------------------------------------------------------------------
' Utilidades de preparacion
' ---------------------------------------------------------------------------
Private Function Err_() As String
    Dim origen As String
    origen = Trim$(CStr(Err.Source))
    Err_ = "ERROR " & Err.Number & ": " & Err.Description & _
        IIf(Len(origen) > 0, " [" & origen & "]", "") & _
        IIf(Len(mDonde) > 0, " [" & mPruebaActual & " @ " & mDonde & "]", "")
End Function

Private Sub DesprotegerTodo()
    Dim ws As Worksheet
    For Each ws In Libro().Worksheets
        If ws.ProtectContents Then ws.Unprotect Password:=CLAVE_PROTECCION
    Next ws
End Sub

' Estado base conocido: configuracion valida, catalogo de 3 materiales,
' meses limpios, enero en APERTURA PENDIENTE.
Private Sub Preparar()
    Dim ws As Worksheet
    Dim f As Long
    Dim i As Long
    Dim tLimp As Single
    Dim textoPrep As String
    On Error Resume Next
    ' Los eventos de hoja se apagan durante la preparacion y el cuerpo de la
    ' prueba: las escrituras son programaticas y Workbook_SheetChangeEvent
    ' solo debe intervenir ante ediciones del usuario. Probe y Suite
    ' restauran el estado previo al terminar.
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    modUI.PruebaSilencioso = True
    On Error GoTo 0
    ' Las escrituras masivas de esta preparacion son programaticas: se protegen
    ' con Reentrada para que Workbook_SheetChange no valide 5.000 x 17 celdas.
    ' Sin esta guarda el manejador construye un informe enorme y la sesion se
    ' cuelga en el primer ClearContents de un mes.
    Entrar
    On Error GoTo falloPreparar

    Paso "inicio calc=" & CStr(Application.Calculation)
    DesprotegerTodo
    modCatalogo.AsegurarFormatoCodigos
    Paso "desproteger"

    ' configuracion
    Set ws = Libro().Sheets(HOJA_CONFIG)
    ws.Range(CFG_ENTIDAD).Value = "EMPRESA DE PRUEBAS C.A."
    ws.Range(CFG_RIF).Value = "J-12345678-9"
    ws.Range(CFG_DOMICILIO).Value = "AV. PRUEBA, EDIFICIO 1, CARACAS"
    ws.Range(CFG_RESPONSABLE).Value = "RESPONSABLE DEL LIBRO"
    ws.Range(CFG_ANIO).Value = AnioDePrueba()
    ws.Range(CFG_MONEDA).Value = APP_MONEDA
    ws.Range(CFG_VERSION).Value = APP_VERSION
    ws.Range(CFG_METODO).Value = APP_METODO
    ws.Range(CFG_MASCARA_CANT).Value = FMT_CANT
    ws.Range(CFG_ESTADO).Value = "CAPTURADO"
    ws.Range(CFG_AUTORIZACION).Value = "SI - verificada por el responsable el 01/01/2026"
    ws.Range(CFG_COSTO_CERO).Value = "NO ACEPTADO"
    Paso "config"

    ' catalogo
    Set ws = Libro().Sheets(HOJA_CATALOGO)
    ws.Range(ws.Cells(FILA_1, 1), ws.Cells(FILA_FIN, 6)).ClearContents
    ws.Cells(FILA_1, 1).Value = "M-0001"
    ws.Cells(FILA_1, 2).Value = "Material uno"
    ws.Cells(FILA_1, 3).Value = "UND"
    ws.Cells(FILA_1, 4).Value = "Si"
    ws.Cells(FILA_1, 5).Value = DateSerial(AnioDePrueba(), 1, 1)
    ws.Cells(FILA_1 + 1, 1).Value = "M-0002"
    ws.Cells(FILA_1 + 1, 2).Value = "Material dos"
    ws.Cells(FILA_1 + 1, 3).Value = "KG"
    ws.Cells(FILA_1 + 1, 4).Value = "Si"
    ws.Cells(FILA_1 + 2, 1).Value = "M-0003"
    ws.Cells(FILA_1 + 2, 2).Value = "Material tres"
    ws.Cells(FILA_1 + 2, 3).Value = "LT"
    ws.Cells(FILA_1 + 2, 4).Value = "Si"
    InvalidarIndice
    Paso "catalogo"

    ' meses
    For i = 1 To 12
        Set ws = Libro().Sheets(HojaMes(i))
        Paso "mes " & i & " hoja"
        tLimp = Timer
        ws.Range(ws.Cells(FILA_1, COL_CODIGO), ws.Cells(FILA_FIN, COL_ESTADO)).ClearContents
        Paso "mes " & i & " limpio dt=" & Format$(Timer - tLimp, "0.0") & "s"
        modCalculo.RellenarFormulas ws
        Paso "mes " & i & " formulas"
        modCalculo.EscribirTotales ws
        Paso "mes " & i & " totales"
    Next i
    Paso "meses ok"

    ' detalle, instantaneas, bitacora
    Libro().Sheets(HOJA_DETALLE).Range( _
        Libro().Sheets(HOJA_DETALLE).Cells(FILA_DET_1, 1), _
        Libro().Sheets(HOJA_DETALLE).Cells(FILA_DET_FIN, 11)).ClearContents
    Set ws = Libro().Sheets(HOJA_SNAP)
    ws.Range(ws.Cells(2, 1), ws.Cells(ws.Rows.Count, 9)).ClearContents
    Set ws = Libro().Sheets(HOJA_BITACORA)
    ws.Range(ws.Cells(FILA_1, 1), ws.Cells(FILA_FIN, 12)).ClearContents
    Paso "detalle"

    ' estados
    Set ws = Libro().Sheets(HOJA_ESTADOS)
    For i = 1 To 12
        ws.Cells(modEstados.FilaEstado(i), 1).Value = HojaMes(i)
        ws.Cells(modEstados.FilaEstado(i), 2).Value = EST_NO_ABIERTO
        ws.Cells(modEstados.FilaEstado(i), 3).ClearContents
        ws.Cells(modEstados.FilaEstado(i), 4).ClearContents
        ws.Cells(modEstados.FilaEstado(i), 5).ClearContents
        ws.Cells(modEstados.FilaEstado(i), 7).ClearContents
        ws.Cells(modEstados.FilaEstado(i), 8).ClearContents
        ws.Cells(modEstados.FilaEstado(i), 9).ClearContents
    Next i
    ws.Cells(modEstados.FilaEstado(1), 2).Value = EST_APERTURA_PEND
    Paso "estados"

    Paso "calculando"
    Application.Calculate
    Paso "fin calculo"
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Salir
    Exit Sub
falloPreparar:
    textoPrep = Err_()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Salir
    Err.Raise vbObjectError + 590, "modPruebas.Preparar", textoPrep
End Sub

Private Sub Paso(ByVal t As String)
    Dim k As Integer
    Dim fs As Object
    Dim t0 As Single
    On Error Resume Next
    For k = 1 To 40
        Err.Clear
        Set fs = CreateObject("Scripting.FileSystemObject")
        fs.OpenTextFile("C:\libinv177\construccion\out\prep.txt", 8, True).WriteLine _
            Format$(Now, "hh:nn:ss") & " " & t
        If Err.Number = 0 Then Exit Sub
        t0 = Timer
        Do While Timer < t0 + 0.05
            DoEvents
        Loop
    Next k
End Sub

Private Function AnioDePrueba() As Long
    AnioDePrueba = Year(Date)
    If AnioDePrueba < 2001 Then AnioDePrueba = 2026
End Function

' Escribe una fila de movimiento de un mes (sin tocar las columnas calculadas).
Private Sub EscribirFila(ByVal indice As Long, ByVal codigo As String, _
                         ByVal cIni As Variant, ByVal vIni As Variant, _
                         ByVal cEnt As Variant, ByVal vEnt As Variant, _
                         ByVal cVta As Variant, ByVal cOtr As Variant)
    Dim ws As Worksheet
    Dim f As Long
    Set ws = Libro().Sheets(HojaMes(indice))
    If ws.ProtectContents Then ws.Unprotect Password:=CLAVE_PROTECCION
    f = FilaEnMesPublica(indice, codigo)
    If f = 0 Then
        f = FILA_1
        Do While Trim$(CStr(ws.Cells(f, COL_CODIGO).Value2)) <> ""
            f = f + 1
        Loop
        ws.Cells(f, COL_CODIGO).Value = codigo
        ws.Cells(f, COL_DESCRIPCION).Value = CStr( _
            modCatalogo.HojaCatalogo().Cells(FilaEnCatalogo(codigo), COL_DESCRIPCION).Value2)
        ws.Cells(f, COL_UNIDAD).Value = CStr( _
            modCatalogo.HojaCatalogo().Cells(FilaEnCatalogo(codigo), 3).Value2)
    End If
    If Not EsVacio(cIni) Then ws.Cells(f, COL_CANT_INI).Value = cIni
    If Not EsVacio(vIni) Then ws.Cells(f, COL_VAL_INI).Value = vIni
    If Not EsVacio(cEnt) Then ws.Cells(f, COL_CANT_ENT).Value = cEnt
    If Not EsVacio(vEnt) Then ws.Cells(f, COL_VAL_ENT).Value = vEnt
    If Not EsVacio(cVta) Then ws.Cells(f, COL_CANT_VTA).Value = cVta
    If Not EsVacio(cOtr) Then ws.Cells(f, COL_CANT_OTR).Value = cOtr
End Sub

Public Function FilaEnMesPublica(ByVal indice As Long, ByVal codigo As String) As Long
    Dim ws As Worksheet
    Dim f As Long
    Dim clave As String
    clave = NormalizaClave(codigo)
    If clave = "" Then Exit Function
    Set ws = Libro().Sheets(HojaMes(indice))
    For f = FILA_1 To FILA_FIN
        If NormalizaClave(ws.Cells(f, COL_CODIGO).Value2) = clave Then
            FilaEnMesPublica = f
            Exit Function
        End If
    Next f
End Function

Private Function V(ByVal ws As Worksheet, ByVal f As Long, ByVal c As Long) As Double
    If EsNumeroReal(ws.Cells(f, c).Value2) Then V = CDbl(ws.Cells(f, c).Value2)
End Function

Private Function Hallazgo(ByVal condicion As Boolean, ByVal texto As String) As String
    If Not condicion Then Hallazgo = texto Else Hallazgo = ""
End Function

Private Function Acumular(ByVal a As String, ByVal b As String) As String
    If a = "" Then
        Acumular = b
    ElseIf b = "" Then
        Acumular = a
    Else
        Acumular = a & vbCrLf & b
    End If
End Function

Private Function Contiene(ByVal texto As String, ByVal busqueda As String) As Boolean
    Contiene = (InStr(1, texto, busqueda, vbTextCompare) > 0)
End Function

' Marcador de trabajo (TODO/FIXME) usado como palabra propia: un simple InStr
' marcaria como TODO el interior de palabras como "Metodo".
Private Function TieneMarcador(ByVal texto As String) As Boolean
    TieneMarcador = MarcadorPalabra(texto, "TODO") Or MarcadorPalabra(texto, "FIXME")
End Function

Private Function MarcadorPalabra(ByVal texto As String, ByVal m As String) As Boolean
    Dim t As String
    Dim p As Long
    Dim antes As String
    Dim despues As String
    t = UCase$(texto)
    p = InStr(1, t, m)
    Do While p > 0
        If p > 1 Then
            antes = Mid$(t, p - 1, 1)
        Else
            antes = ""
        End If
        despues = Mid$(t, p + Len(m), 1)
        If Not EsLetraAscii(antes) And Not EsLetraAscii(despues) Then
            MarcadorPalabra = True
            Exit Function
        End If
        p = InStr(p + 1, t, m)
    Loop
End Function

Private Function EsLetraAscii(ByVal c As String) As Boolean
    If Len(c) = 0 Then Exit Function
    EsLetraAscii = (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z")
End Function

' Abre enero con los datos ya cargados.
Private Function AbrirEnero() As String
    Dim errNum As Long
    Dim errTxt As String
    On Error Resume Next
    Err.Clear
    modEstados.ConfirmarAperturaEnero
    errNum = Err.Number
    errTxt = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNum <> 0 Then
        AbrirEnero = "fallo la apertura de enero: " & errTxt
        Exit Function
    End If
    AbrirEnero = Hallazgo(modEstados.EstadoMes(1) = EST_ABIERTO, _
        "enero no quedo en estado ABIERTO (queda " & modEstados.EstadoMes(1) & ")")
End Function

Private Function Cerrar(ByVal indice As Long) As String
    Dim errNum As Long
    Dim errTxt As String
    Dim respuesta As String
    On Error Resume Next
    Err.Clear
    respuesta = modEstados.CerrarMes(indice, "RESPONSABLE DEL LIBRO", "cierre de prueba")
    errNum = Err.Number
    errTxt = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNum <> 0 Then
        Cerrar = "el cierre fallo: " & errTxt
    Else
        Cerrar = respuesta
    End If
End Function

' ===========================================================================
' T01 - Caso HT-0001 completo, cierre y apertura de febrero
' ===========================================================================
Private Function T01() As String
    Dim res As String
    Dim ene As Worksheet, feb As Worksheet
    Dim f As Long, k As Long, l As Double, cc As Double
    On Error GoTo fallo
    Preparar
    EscribirFila 1, "M-0001", 10, 1000, 5, 600, 4, 1
    EscribirFila 1, "M-0002", 8, 800, 0, Empty, 3, 0
    ' M-0001 tiene I=1: sin detalle en DETALLE_OTROS el cierre se bloquea.
    AgregarLineaDetalle 1, "M-0001", 15, "retiro", 1
    Application.Calculate
    res = Acumular(res, AbrirEnero())
    res = Acumular(res, Cerrar(1))
    res = Acumular(res, Hallazgo(modEstados.EstadoMes(1) = EST_CERRADO, "enero no quedo CERRADO"))
    res = Acumular(res, Hallazgo(modEstados.AbrirMes(2) = "", "no se pudo abrir febrero"))
    Set ene = Libro().Sheets(HojaMes(1))
    Set feb = Libro().Sheets(HojaMes(2))
    f = FilaEnMesPublica(1, "M-0001")
    k = FilaEnMesPublica(2, "M-0001")
    If f = 0 Or k = 0 Then
        res = Acumular(res, "M-0001 no aparece en enero o febrero")
    Else
        cc = V(ene, f, COL_CANT_FIN)
        l = V(ene, f, COL_VAL_FIN)
        res = Acumular(res, Hallazgo(Abs(V(feb, k, COL_CANT_INI) - cc) <= 0.00005, _
            "febrero no tomo K de enero"))
        res = Acumular(res, Hallazgo(Abs(V(feb, k, COL_VAL_INI) - l) <= 0.005, _
            "febrero no tomo L de enero"))
        res = Acumular(res, Hallazgo(Abs(V(feb, k, COL_CANT_ENT)) <= 0.00005, _
            "febrero arrastro entradas en vez de dejarlas en cero"))
    End If
    res = Acumular(res, Hallazgo(modEstados.MesAbierto() = 2, "febrero no es el mes abierto"))
    res = Acumular(res, Hallazgo(modEstados.EstadoMes(1) = EST_CERRADO, "enero dejo de estar cerrado"))
    T01 = res
    Exit Function
fallo:
    T01 = Err_()
End Function

' Diagnostico: ejecuta T01 paso a paso dejando constancia en prep.txt para
' localizar la sentencia exacta que dispara un error de compilacion.
Private Function T01Diag() As String
    Dim res As String
    Dim txt As String
    On Error GoTo fallo
    Preparar
    Paso "dg preparar ok"
    modCalculo.EscribirFormulasFila Libro().Sheets(HojaMes(1)), FILA_1
    Paso "dg EscribirFormulasFila ok"
    modCalculo.AplicarFormulas Libro().Sheets(HojaMes(1)), FILA_1, FILA_1
    Paso "dg AplicarFormulas ok"
    txt = CStr(modCalculo.ValorNumerico(Libro().Sheets(HojaMes(1)), FILA_1, COL_VAL_INI))
    Paso "dg ValorNumerico ok " & txt
    txt = CStr(modCalculo.CuadraFila(Libro().Sheets(HojaMes(1)), FILA_1))
    Paso "dg CuadraFila ok " & txt
    txt = CStr(modCalculo.SumaColumna(Libro().Sheets(HojaMes(1)), COL_VAL_INI))
    Paso "dg SumaColumna ok " & txt
    txt = CStr(modCalculo.CuadraTotales(Libro().Sheets(HojaMes(1))))
    Paso "dg CuadraTotales ok " & txt
    modCalculo.AjustarDecimalesVisibles Libro().Sheets(HojaMes(1)), FILA_1
    Paso "dg AjustarDecimales ok"
    EscribirFila 1, "M-0001", 10, 1000, 5, 600, 4, 1
    EscribirFila 1, "M-0002", 8, 800, 0, Empty, 3, 0
    Application.Calculate
    Paso "dg filas ok"
    txt = AbrirEnero()
    Paso "dg abrirenero " & txt
    txt = Cerrar(1)
    Paso "dg cerrar " & txt
    res = Hallazgo(modEstados.EstadoMes(1) = EST_CERRADO, "enero no quedo CERRADO")
    txt = CStr(modEstados.AbrirMes(2))
    Paso "dg abrirfebrero " & txt
    T01Diag = res
    Exit Function
fallo:
    T01Diag = Err_()
End Function

' Diagnostico de los fallos de T02/T03/T04: devuelve el dato crudo.
Private Function DbgDiag(ByVal modo As String) As String
    Dim cat As Worksheet
    Dim ws As Worksheet
    Dim rango As Range
    Dim f As Long
    Dim a As Long, b As Long
    Dim i As Long
    Dim p As Long
    Dim texto As String
    On Error GoTo fallo
    Select Case UCase$(modo)
        Case "CAT"
            Preparar
            Set cat = modCatalogo.HojaCatalogo()
            cat.Cells(FILA_1, 1).Value = "0007"
            cat.Cells(FILA_1, 2).Value = "Con ceros iniciales"
            cat.Cells(FILA_1, 3).Value = "UND"
            cat.Cells(FILA_1, 4).Value = "Si"
            cat.Cells(FILA_1 + 1, 1).Value = "7"
            cat.Cells(FILA_1 + 1, 2).Value = "Sin ceros"
            cat.Cells(FILA_1 + 1, 3).Value = "UND"
            cat.Cells(FILA_1 + 1, 4).Value = "Si"
            InvalidarIndice
            a = FilaEnCatalogo("0007")
            b = FilaEnCatalogo("7")
            DbgDiag = "a=" & a & " b=" & b & " FILA_1=" & FILA_1 & _
                " norm7=[" & NormalizaClave("7") & "] norm0007=[" & NormalizaClave("0007") & _
                "] cat=[" & modCatalogo.ValidarCatalogo() & "]"
        Case "T04"
            Preparar
            EscribirFila 1, "M-0001", 5, 500, 0, Empty, 9, 0
            Application.Calculate
            Set ws = Libro().Sheets(HojaMes(1))
            f = FilaEnMesPublica(1, "M-0001")
            DbgDiag = "f=" & f & " K=" & V(ws, f, COL_CANT_FIN) & _
                " Q=[" & CStr(ws.Cells(f, COL_ESTADO).Value2) & "]" & _
                " Qf=[" & CStr(ws.Cells(f, COL_ESTADO).Formula) & "]" & _
                " puedeCerrar=[" & modEstados.PuedeCerrar(1) & "]"
        Case "T02"
            Preparar
            Set cat = modCatalogo.HojaCatalogo()
            Set ws = Libro().Sheets(HojaMes(1))
            If ws.ProtectContents Then ws.Unprotect Password:=CLAVE_PROTECCION
            Set rango = ws.Range(ws.Cells(FILA_1, COL_CODIGO), ws.Cells(FILA_1 + 2, COL_CODIGO))
            rango.Value2 = "NO-EXISTE"
            texto = "bloqueNoExiste=[" & modCatalogo.ValidarBloquePegado(ws, rango) & "]"
            rango.ClearContents
            Set rango = ws.Range(ws.Cells(FILA_1, COL_CANT_VTA), ws.Cells(FILA_1, COL_VAL_ENT))
            rango.Value2 = "texto"
            texto = texto & " || textoNum=[" & modCatalogo.ValidarBloquePegado(ws, rango) & "]"
            rango.ClearContents
            cat.Cells(FILA_1 + 3, 1).Value = "  m-0001  "
            cat.Cells(FILA_1 + 3, 2).Value = "Copia con espacios"
            cat.Cells(FILA_1 + 3, 3).Value = "UND"
            cat.Cells(FILA_1 + 3, 4).Value = "Si"
            InvalidarIndice
            texto = texto & " || cat=[" & modCatalogo.ValidarCatalogo() & "]"
            DbgDiag = texto
        Case "T03B"
            Preparar
            Set cat = modCatalogo.HojaCatalogo()
            cat.Cells(FILA_1, 1).Value = "0007"
            cat.Cells(FILA_1, 2).Value = "Con ceros iniciales"
            cat.Cells(FILA_1, 3).Value = "UND"
            cat.Cells(FILA_1, 4).Value = "Si"
            cat.Cells(FILA_1 + 1, 1).Value = "7"
            cat.Cells(FILA_1 + 1, 2).Value = "Sin ceros"
            cat.Cells(FILA_1 + 1, 3).Value = "UND"
            cat.Cells(FILA_1 + 1, 4).Value = "Si"
            For i = 0 To 2
                cat.Cells(FILA_1 + i, 1).Value = Choose(i + 1, "7", "M-0002", "0007")
                cat.Cells(FILA_1 + i, 2).Value = Choose(i + 1, "Sin ceros", "Material dos", "Con ceros iniciales")
                cat.Cells(FILA_1 + i, 3).Value = "UND"
                cat.Cells(FILA_1 + i, 4).Value = "Si"
    Next i
    InvalidarIndice
    Paso "T25 catalogo cargado"
            DbgDiag = "a=" & FilaEnCatalogo("0007") & " b=" & FilaEnCatalogo("M-0002") & _
                " c=" & FilaEnCatalogo("7") & _
                " r7=[" & CStr(cat.Cells(FILA_1, 1).Value2) & "]" & _
                " r8=[" & CStr(cat.Cells(FILA_1 + 1, 1).Value2) & "]" & _
                " r9=[" & CStr(cat.Cells(FILA_1 + 2, 1).Value2) & "]" & _
                " fmt=[" & CStr(cat.Cells(FILA_1, 1).NumberFormat) & "]" & _
                " ult=" & UltimaFilaUsada(cat, 1)
        Case "T09B"
            Preparar
            EscribirFila 1, "M-0002", 4, 400, 0, Empty, 1, 0
            Application.Calculate
            texto = "abrir=[" & AbrirEnero() & "] cerrar=[" & Cerrar(1) & "]"
            texto = texto & " estado1=" & modEstados.EstadoMes(1) & _
                " hasta=" & modEstados.MesCerradoHasta() & _
                " activo=[" & CStr(modCatalogo.HojaCatalogo().Cells(FILA_1 + 1, 4).Value2) & "]" & _
                " pend=" & modCatalogo.SaldoPendiente("M-0002") & _
                " K=" & V(Libro().Sheets(HojaMes(1)), FilaEnMesPublica(1, "M-0002"), COL_CANT_FIN)
            Set ws = Libro().Sheets(HojaMes(1))
            On Error Resume Next
            Err.Clear
            ws.Cells(FilaEnMesPublica(1, "M-0002"), COL_CANT_VTA).Value = 3
            texto = texto & " errVta=" & Err.Number & " " & Err.Description
            Err.Clear
            On Error GoTo fallo
            Application.Calculate
            texto = texto & " prot=" & CStr(ws.ProtectContents) & _
                " K2=" & V(Libro().Sheets(HojaMes(1)), FilaEnMesPublica(1, "M-0002"), COL_CANT_FIN) & _
                " pend2=" & modCatalogo.SaldoPendiente("M-0002")
            DbgDiag = texto
        Case "T23B"
            Preparar
            EscribirFila 2, "M-0002", 1, 0.01, 1, 0.01, 1, 0
            Application.Calculate
            Set ws = Libro().Sheets(HojaMes(2))
            f = FilaEnMesPublica(2, "M-0002")
            DbgDiag = "O=" & V(ws, f, COL_COSTO_PROM) & " K=" & V(ws, f, COL_CANT_FIN) & _
                " L=" & V(ws, f, COL_VAL_FIN) & _
                " val=[" & modValidacion.ValidarMes(2) & "]"
        Case "T27B"
            Preparar
            texto = ""
            For Each ws In Libro().Worksheets
                If ws.Name <> HOJA_DETALLE Then
                    texto = texto & CStr(ws.Range("A1").Value2) & " " & _
                        CStr(ws.Range("A2").Value2)
                End If
            Next ws
            p = InStr(1, texto, "TODO", vbTextCompare)
            If p > 0 Then
                DbgDiag = "len=" & Len(texto) & " p=" & p & " det=[" & HOJA_DETALLE & _
                    "] ctx=[" & Mid$(texto, IIf(p > 40, p - 40, 1), 90) & "]"
            Else
                DbgDiag = "len=" & Len(texto) & " p=0 det=[" & HOJA_DETALLE & "]"
            End If
        Case Else
            DbgDiag = "modo desconocido"
    End Select
    Exit Function
fallo:
    DbgDiag = Err_()
End Function

' ===========================================================================
' T02 - Duplicados por mayusculas/espacios, alta y pegado: rechazados integros
' ===========================================================================
Private Function T02() As String
    Dim res As String
    Dim ws As Worksheet
    Dim cat As Worksheet
    Dim rango As Range
    Dim errNum As Long
    On Error GoTo fallo
    Preparar
    Set cat = modCatalogo.HojaCatalogo()
    ' duplicado por mayusculas y espacios
    cat.Cells(FILA_1 + 3, 1).Value = "  m-0001  "
    cat.Cells(FILA_1 + 3, 2).Value = "Copia con espacios"
    cat.Cells(FILA_1 + 3, 3).Value = "UND"
    cat.Cells(FILA_1 + 3, 4).Value = "Si"
    InvalidarIndice
    res = Acumular(res, Hallazgo(Contiene(modCatalogo.ValidarCatalogo(), "duplicado"), _
        "el catalogo no detecto el duplicado por mayusculas/espacios"))
    ' alta repetida: la clave normalizada colisiona
    res = Acumular(res, Hallazgo(FilaEnCatalogo("M-0001") = _
        FilaEnCatalogo(" m-0001 "), _
        "las claves normalizadas no apuntan a la misma fila"))
    cat.Cells(FILA_1 + 3, 1).ClearContents
    cat.Cells(FILA_1 + 3, 2).ClearContents
    cat.Cells(FILA_1 + 3, 3).ClearContents
    cat.Cells(FILA_1 + 3, 4).ClearContents
    InvalidarIndice

    ' pegado de un codigo inexistente: se rechaza la accion completa
    Set ws = Libro().Sheets(HojaMes(1))
    If ws.ProtectContents Then ws.Unprotect Password:=CLAVE_PROTECCION
    Set rango = ws.Range(ws.Cells(FILA_1, COL_CODIGO), ws.Cells(FILA_1 + 2, COL_CODIGO))
    rango.Value2 = "NO-EXISTE"   ' 3 filas iguales
    res = Acumular(res, Hallazgo(Contiene(modCatalogo.ValidarBloquePegado(ws, rango), _
        "no existe en CATALOGO"), "el pegado con codigo inexistente no fue rechazado"))
    rango.ClearContents
    ' pegado de texto en columna numerica
    Set rango = ws.Range(ws.Cells(FILA_1, COL_CANT_VTA), ws.Cells(FILA_1, COL_VAL_ENT))
    rango.Value2 = "texto"
    res = Acumular(res, Hallazgo(Contiene(modCatalogo.ValidarBloquePegado(ws, rango), _
        "no es un numero"), "el texto en columna numerica no fue rechazado"))
    rango.ClearContents
    T02 = res
    Exit Function
fallo:
    T02 = Err_()
End Function

' ===========================================================================
' T03 - Codigos con ceros iniciales y catalogo reordenado
' ===========================================================================
Private Function T03() As String
    Dim res As String
    Dim cat As Worksheet
    Dim filaA As Long, filaB As Long
    Dim i As Long
    On Error GoTo fallo
    Preparar
    Set cat = modCatalogo.HojaCatalogo()
    cat.Cells(FILA_1, 1).Value = "0007"
    cat.Cells(FILA_1, 2).Value = "Con ceros iniciales"
    cat.Cells(FILA_1, 3).Value = "UND"
    cat.Cells(FILA_1, 4).Value = "Si"
    cat.Cells(FILA_1 + 1, 1).Value = "7"
    cat.Cells(FILA_1 + 1, 2).Value = "Sin ceros"
    cat.Cells(FILA_1 + 1, 3).Value = "UND"
    cat.Cells(FILA_1 + 1, 4).Value = "Si"
    InvalidarIndice
    filaA = FilaEnCatalogo("0007")
    filaB = FilaEnCatalogo("7")
    res = Acumular(res, Hallazgo(filaA <> 0, "el codigo 0007 no se encuentra"))
    res = Acumular(res, Hallazgo(filaB <> 0, "el codigo 7 no se encuentra"))
    res = Acumular(res, Hallazgo(filaA <> filaB, _
        "0007 y 7 colisionan: los ceros iniciales se pierden"))
    res = Acumular(res, Hallazgo(FilaEnCatalogo("0007 ") = filaA, _
        "un espacio final cambio la clave"))

    ' catalogo reordenado: se invierte el orden de las filas con datos
    For i = 0 To 2
        cat.Cells(FILA_1 + i, 1).Value = Choose(i + 1, "7", "M-0002", "0007")
        cat.Cells(FILA_1 + i, 2).Value = Choose(i + 1, "Sin ceros", "Material dos", "Con ceros iniciales")
        cat.Cells(FILA_1 + i, 3).Value = "UND"
        cat.Cells(FILA_1 + i, 4).Value = "Si"
    Next i
    InvalidarIndice
    res = Acumular(res, Hallazgo(FilaEnCatalogo("0007") = FILA_1 + 2, _
        "tras reordenar, 0007 dejo de apuntar a su fila"))
    res = Acumular(res, Hallazgo(FilaEnCatalogo("M-0002") = FILA_1 + 1, _
        "tras reordenar, M-0002 dejo de apuntar a su fila"))
    T03 = res
    Exit Function
fallo:
    T03 = Err_()
End Function

' ===========================================================================
' T04 - Salidas mayores a disponibles
' ===========================================================================
Private Function T04() As String
    Dim res As String
    Dim ws As Worksheet
    Dim f As Long
    Dim texto As String
    On Error GoTo fallo
    Preparar
    EscribirFila 1, "M-0001", 5, 500, 0, Empty, 9, 0    ' vende mas de lo disponible
    Application.Calculate
    Set ws = Libro().Sheets(HojaMes(1))
    f = FilaEnMesPublica(1, "M-0001")
    res = Acumular(res, Hallazgo(V(ws, f, COL_CANT_FIN) <= -0.00005, _
        "K no quedo negativo para evidentar el exceso"))
    res = Acumular(res, Hallazgo(Contiene(CStr(ws.Cells(f, COL_ESTADO).Value2), "ERROR"), _
        "la columna Q no marca ERROR (f=" & f & " Q=[" & _
        CStr(ws.Cells(f, COL_ESTADO).Value2) & "])"))
    res = Acumular(res, AbrirEnero())
    texto = modEstados.PuedeCerrar(1)
    res = Acumular(res, Hallazgo(texto <> "", "el cierre no quedo bloqueado con saldo negativo"))
    res = Acumular(res, Hallazgo(Contiene(texto, "salidas mayores"), _
        "no se explica el exceso de salidas"))
    res = Acumular(res, Hallazgo(Cerrar(1) <> "", "el cierre se ejecuto igual"))
    res = Acumular(res, Hallazgo(modEstados.EstadoMes(1) <> EST_CERRADO, _
        "el mes quedo CERRADO pese al exceso"))
    ' tampoco se abre febrero sin cerrar enero
    res = Acumular(res, Hallazgo(Contiene(modEstados.AbrirMes(2), "sin cerrar"), _
        "se abrio febrero sin cerrar enero"))
    T04 = res
    Exit Function
fallo:
    T04 = Err_()
End Function

' ===========================================================================
' T05 - Cantidad cero con valor distinto de cero
' ===========================================================================
Private Function T05() As String
    Dim res As String
    Dim ws As Worksheet
    Dim f As Long
    Dim texto As String
    On Error GoTo fallo
    Preparar
    EscribirFila 1, "M-0001", 0, 100, 0, Empty, 0, 0
    Application.Calculate
    Set ws = Libro().Sheets(HojaMes(1))
    f = FilaEnMesPublica(1, "M-0001")
    res = Acumular(res, Hallazgo(Contiene(CStr(ws.Cells(f, COL_ESTADO).Value2), "ERROR"), _
        "Q no marca ERROR con C=0 y D<>0 (f=" & f & " Q=[" & _
        CStr(ws.Cells(f, COL_ESTADO).Value2) & "])"))
    res = Acumular(res, Hallazgo(V(ws, f, COL_COSTO_PROM) = 0, _
        "O no quedo en cero (division oculta)"))
    res = Acumular(res, Hallazgo(Contiene(CStr(ws.Cells(f, COL_COSTO_PROM).Formula), "IF"), _
        "O no tiene proteccion por division entre cero"))
    res = Acumular(res, AbrirEnero())
    texto = modValidacion.ValidarMes(1)
    res = Acumular(res, Hallazgo(Contiene(texto, "C=0 con D distinto de 0"), _
        "la validacion no reporta C=0 con D<>0"))
    T05 = res
    Exit Function
fallo:
    T05 = Err_()
End Function

' ===========================================================================
' T06 - Stock cero valido; apertura vacia invalida; costo cero exigido
' ===========================================================================
Private Function T06() As String
    Dim res As String
    Dim ws As Worksheet
    Dim texto As String
    On Error GoTo fallo
    Preparar
    ' a) stock cero confirmado es valido
    EscribirFila 1, "M-0001", 10, 100, 0, Empty, 10, 0
    Application.Calculate
    res = Acumular(res, AbrirEnero())
    texto = modValidacion.ValidarMes(1)
    res = Acumular(res, Hallazgo(texto = "", "stock cero no es valido: " & texto))
    res = Acumular(res, Cerrar(1))
    res = Acumular(res, Hallazgo(modEstados.EstadoMes(1) = EST_CERRADO, "enero no cerro"))

    ' b) apertura vacia invalida (fila con codigo y C/D vacios)
    EscribirFila 2, "M-0002", Empty, Empty, Empty, Empty, Empty, Empty
    Application.Calculate
    texto = modValidacion.ValidarMes(2)
    res = Acumular(res, Hallazgo(Contiene(texto, "saldo inicial vacio"), _
        "no detecto apertura vacia en febrero"))
    res = Acumular(res, Hallazgo(Contiene(modEstados.AbrirMes(2), "saldo inicial"), _
        "febrero no fue bloqueado con una apertura vacia"))
    res = Acumular(res, Hallazgo(modEstados.EstadoMes(2) <> EST_ABIERTO, _
        "febrero quedo ABIERTO con una apertura vacia"))

    ' c) costo cero exige aceptacion documentada
    Preparar
    EscribirFila 1, "M-0001", 0, 0, 1, 0, 1, 0   ' QD>0 pero VD=0 -> O=0
    Application.Calculate
    res = Acumular(res, AbrirEnero())
    texto = modValidacion.ValidarMes(1)
    res = Acumular(res, Hallazgo(Contiene(texto, "costo promedio cero"), _
        "no exige aceptacion de costo cero"))
    res = Acumular(res, Hallazgo(modConfiguracion.AceptacionCostoCero() = False, _
        "la configuracion dice que el costo cero ya estaba aceptado"))
    Libro().Sheets(HOJA_CONFIG).Range(CFG_COSTO_CERO).Value = _
        "ACEPTADO por RESPONSABLE DEL LIBRO el 01/01/2026"
    Application.Calculate
    texto = modValidacion.ValidarMes(1)
    res = Acumular(res, Hallazgo(Not Contiene(texto, "costo promedio cero"), _
        "con aceptacion documentada sigue bloqueando"))
    T06 = res
    Exit Function
fallo:
    T06 = Err_()
End Function

' ===========================================================================
' T07 - Agotamiento C=3, D=1, E=F=0, G=1, I=2
' ===========================================================================
Private Function T07() As String
    Dim res As String
    Dim ws As Worksheet
    Dim f As Long
    Dim texto As String
    On Error GoTo fallo
    Preparar
    EscribirFila 1, "M-0001", 3, 1, 0, 0, 1, 2
    Application.Calculate
    Set ws = Libro().Sheets(HojaMes(1))
    f = FilaEnMesPublica(1, "M-0001")
    res = Acumular(res, Hallazgo(Abs(V(ws, f, COL_VAL_VTA) - 0.33) <= 0.005, _
        "H=" & V(ws, f, COL_VAL_VTA) & " y debe ser 0,33"))
    res = Acumular(res, Hallazgo(Abs(V(ws, f, COL_VAL_OTR) - 0.67) <= 0.005, _
        "J=" & V(ws, f, COL_VAL_OTR) & " y debe ser 0,67"))
    res = Acumular(res, Hallazgo(Abs(V(ws, f, COL_CANT_FIN)) <= 0.00005, "K no es 0"))
    res = Acumular(res, Hallazgo(Abs(V(ws, f, COL_VAL_FIN)) <= 0.005, "L no es 0"))
    res = Acumular(res, Hallazgo(Abs(V(ws, f, COL_VAL_VTA) + V(ws, f, COL_VAL_OTR) - 1) <= 0.005, _
        "H+J no reparte exactamente VD"))
    res = Acumular(res, AbrirEnero())

    ' detalle: I=2 exige dos lineas que sumen 2
    texto = modDetalle.ValidarDetalleDeMes(1)
    res = Acumular(res, Hallazgo(Contiene(texto, "I>0 sin detalle"), _
        "I=2 sin detalle no fue bloqueado"))
    AgregarLineaDetalle 1, "M-0001", 1, "retiro", 1
    texto = modDetalle.ValidarDetalleDeMes(1)
    res = Acumular(res, Hallazgo(Contiene(texto, "no cuadra con I"), _
        "detalle de 1 con I=2 no fue detectado"))
    AgregarLineaDetalle 1, "M-0001", 1, "retiro", 1
    texto = modDetalle.ValidarDetalleDeMes(1)
    res = Acumular(res, Hallazgo(texto = "", "detalle 1+1 con I=2 no concilio: " & texto))
    res = Acumular(res, Hallazgo(modDetalle.SumaDeclarada(1, "M-0001") = 2, _
        "SumaDeclarada no devuelve la columna I"))
    T07 = res
    Exit Function
fallo:
    T07 = Err_()
End Function

Private Sub AgregarLineaDetalle(ByVal mes As Long, ByVal codigo As String, _
                                ByVal dia As Long, ByVal tipo As String, _
                                ByVal cantidad As Double)
    Dim ws As Worksheet
    Dim f As Long
    Set ws = modDetalle.HojaDetalle()
    If ws.ProtectContents Then ws.Unprotect Password:=CLAVE_PROTECCION
    f = FILA_DET_1
    Do While Not EsVacio(ws.Cells(f, DET_MES).Value2)
        f = f + 1
    Loop
    ws.Cells(f, DET_MES).Value = mes
    ws.Cells(f, DET_CODIGO).Value = codigo
    ws.Cells(f, DET_FECHA).Value = DateSerial(AnioDePrueba(), mes, dia)
    ws.Cells(f, DET_TIPO).Value = tipo
    ws.Cells(f, DET_CANTIDAD).Value = cantidad
    ws.Cells(f, DET_DOCUMENTO).Value = "DOC-" & f
    ws.Cells(f, DET_OBSERVACION).Value = IIf(tipo = "otro", "explicacion de la salida", "")
    ws.Cells(f, DET_FOLIO).Value = f
End Sub

' ===========================================================================
' T08 - Agotamiento solo ventas o solo otros
' ===========================================================================
Private Function T08() As String
    Dim res As String
    Dim ws As Worksheet
    Dim f As Long
    On Error GoTo fallo
    ' a) solo ventas
    Preparar
    EscribirFila 1, "M-0001", 10, 100, 0, 0, 10, 0
    Application.Calculate
    Set ws = Libro().Sheets(HojaMes(1))
    f = FilaEnMesPublica(1, "M-0001")
    res = Acumular(res, Hallazgo(Abs(V(ws, f, COL_VAL_VTA) - 100) <= 0.005, _
        "solo ventas: H=" & V(ws, f, COL_VAL_VTA) & " y debe ser 100"))
    res = Acumular(res, Hallazgo(Abs(V(ws, f, COL_VAL_OTR)) <= 0.005, "J debe ser 0"))
    res = Acumular(res, Hallazgo(Abs(V(ws, f, COL_VAL_FIN)) <= 0.005, "L debe ser 0"))

    ' b) solo otros
    Preparar
    EscribirFila 1, "M-0001", 10, 100, 0, 0, 0, 10
    Application.Calculate
    Set ws = Libro().Sheets(HojaMes(1))
    f = FilaEnMesPublica(1, "M-0001")
    res = Acumular(res, Hallazgo(Abs(V(ws, f, COL_VAL_VTA)) <= 0.005, "H debe ser 0"))
    res = Acumular(res, Hallazgo(Abs(V(ws, f, COL_VAL_OTR) - 100) <= 0.005, _
        "solo otros: J=" & V(ws, f, COL_VAL_OTR) & " y debe ser 100"))
    res = Acumular(res, Hallazgo(Abs(V(ws, f, COL_VAL_FIN)) <= 0.005, "L debe ser 0"))
    res = Acumular(res, Hallazgo(Abs(V(ws, f, COL_VAL_VTA) + V(ws, f, COL_VAL_OTR) - 100) <= 0.005, _
        "la salida total no es VD"))
    T08 = res
    Exit Function
fallo:
    T08 = Err_()
End Function

' ===========================================================================
' T09 - Alta/baja con stock: no pierde stock ni reescribe historia
' ===========================================================================
Private Function T09() As String
    Dim res As String
    Dim cat As Worksheet
    Dim ws As Worksheet
    Dim f As Long
    Dim descCerrada As String
    On Error GoTo fallo
    Preparar
    EscribirFila 1, "M-0002", 4, 400, 0, Empty, 1, 0
    Application.Calculate
    res = Acumular(res, AbrirEnero())
    res = Acumular(res, Cerrar(1))
    ' con stock: aunque este inactivo, sigue utilizable
    Set cat = modCatalogo.HojaCatalogo()
    cat.Cells(FILA_1 + 1, 4).Value = "No"
    InvalidarIndice
    res = Acumular(res, Hallazgo(modCatalogo.PuedeUsarMaterial("M-0002"), _
        "inactivo con stock dejo de estar disponible"))
    ' en febrero debe arrastrarse el saldo
    res = Acumular(res, Hallazgo(modEstados.AbrirMes(2) = "", "no abrio febrero"))
    Set ws = Libro().Sheets(HojaMes(2))
    f = FilaEnMesPublica(2, "M-0002")
    res = Acumular(res, Hallazgo(f > 0, "M-0002 no llego a febrero"))
    If f > 0 Then
        res = Acumular(res, Hallazgo(Abs(V(ws, f, COL_CANT_INI) - 3) <= 0.00005, _
            "el saldo de cantidad no se arrastro (K de enero = 3)"))
        res = Acumular(res, Hallazgo(Abs(V(ws, f, COL_VAL_INI) - 300) <= 0.005, _
            "el saldo de valor no se arrastro (L de enero = 300)"))
    End If
    ' sin stock: ya no debe arrastrarse en meses nuevos
    Set ws = Libro().Sheets(HojaMes(1))
    ' C=4 con G=1 deja K=3: vender 4 (no 3) agota por completo la existencia.
    ws.Cells(FilaEnMesPublica(1, "M-0002"), COL_CANT_VTA).Value = 4
    Application.Calculate
    res = Acumular(res, Hallazgo(Not modCatalogo.PuedeUsarMaterial("M-0002"), _
        "inactivo sin stock sigue disponible"))
    ' la historia no se reescribe: la celda de enero conserva sus valores
    res = Acumular(res, Hallazgo(Abs(V(ws, FilaEnMesPublica(1, "M-0002"), COL_VAL_INI) - 400) <= 0.005, _
        "se modifico el saldo inicial de un mes cerrado"))
    T09 = res
    Exit Function
fallo:
    T09 = Err_()
End Function

' ===========================================================================
' T10 - Descripcion actualizada no cambia meses cerrados
' ===========================================================================
Private Function T10() As String
    Dim res As String
    Dim cat As Worksheet
    Dim ws As Worksheet
    Dim f As Long
    Dim descCerrada As String
    On Error GoTo fallo
    Preparar
    EscribirFila 1, "M-0001", 5, 500, 0, Empty, 1, 0
    Application.Calculate
    res = Acumular(res, AbrirEnero())
    res = Acumular(res, Cerrar(1))
    Set ws = Libro().Sheets(HojaMes(1))
    f = FilaEnMesPublica(1, "M-0001")
    descCerrada = CStr(ws.Cells(f, COL_DESCRIPCION).Value2)
    ' se actualiza la descripcion maestra
    Set cat = modCatalogo.HojaCatalogo()
    cat.Cells(FilaEnCatalogo("M-0001"), 2).Value = "Material uno DESCRIPCION NUEVA"
    InvalidarIndice
    res = Acumular(res, Hallazgo(CStr(ws.Cells(f, COL_DESCRIPCION).Value2) = descCerrada, _
        "el mes cerrado cambio de descripcion: " & CStr(ws.Cells(f, COL_DESCRIPCION).Value2)))
    res = Acumular(res, Hallazgo(descCerrada = "Material uno", _
        "la descripcion historica no era la original"))
    ' y el catalogo si quedo actualizado
    res = Acumular(res, Hallazgo(Contiene( _
        CStr(cat.Cells(FilaEnCatalogo("M-0001"), 2).Value2), "NUEVA"), _
        "el catalogo no guardo la descripcion nueva"))
    T10 = res
    Exit Function
fallo:
    T10 = Err_()
End Function

' ===========================================================================
' T11 - Bloque pegado con negativo / texto / formula: rechazo completo
' ===========================================================================
Private Function T11() As String
    Dim res As String
    Dim ws As Worksheet
    Dim rango As Range
    On Error GoTo fallo
    Preparar
    res = Acumular(res, AbrirEnero())
    Set ws = Libro().Sheets(HojaMes(1))
    If ws.ProtectContents Then ws.Unprotect Password:=CLAVE_PROTECCION

    Set rango = ws.Range(ws.Cells(FILA_1, COL_CANT_VTA), ws.Cells(FILA_1 + 4, COL_CANT_VTA))
    rango.Value = -1
    res = Acumular(res, Hallazgo(Contiene(modCatalogo.ValidarBloquePegado(ws, rango), _
        "valor negativo"), "el bloque con negativos no fue rechazado"))
    rango.ClearContents

    Set rango = ws.Range(ws.Cells(FILA_1, COL_VAL_ENT), ws.Cells(FILA_1 + 2, COL_VAL_ENT))
    rango.Value = "abc"
    res = Acumular(res, Hallazgo(Contiene(modCatalogo.ValidarBloquePegado(ws, rango), _
        "no es un numero"), "el bloque con texto no fue rechazado"))
    rango.ClearContents

    Set rango = ws.Range(ws.Cells(FILA_1, COL_CANT_ENT), ws.Cells(FILA_1 + 2, COL_CANT_ENT))
    rango.Formula = "=1+1"
    res = Acumular(res, Hallazgo(Contiene(modCatalogo.ValidarBloquePegado(ws, rango), _
        "no es un numero"), "el bloque con formula pegada no fue rechazado"))
    rango.ClearContents

    ' celda fuera del rango tecnico
    Set rango = ws.Range("A1")
    res = Acumular(res, Hallazgo(Contiene(modCatalogo.ValidarBloquePegado(ws, rango), _
        "fuera del rango tecnico"), "no rechazo una celda fuera del rango"))
    T11 = res
    Exit Function
fallo:
    T11 = Err_()
End Function

' ===========================================================================
' T12 - Formula rota o ausente
' ===========================================================================
Private Function T12() As String
    Dim res As String
    Dim ws As Worksheet
    Dim f As Long
    On Error GoTo fallo
    Preparar
    EscribirFila 1, "M-0001", 5, 500, 0, Empty, 1, 0
    Application.Calculate
    res = Acumular(res, AbrirEnero())
    Set ws = Libro().Sheets(HojaMes(1))
    f = FilaEnMesPublica(1, "M-0001")
    res = Acumular(res, Hallazgo(Not modValidacion.HayErrorDeFormula(ws, f), _
        "con formulas completas marco error"))
    ' se rompe la formula de L
    ws.Cells(f, COL_VAL_FIN).Value = 999
    Application.Calculate
    res = Acumular(res, Hallazgo(modValidacion.HayErrorDeFormula(ws, f), _
        "no detecto la formula rota"))
    res = Acumular(res, Hallazgo(Contiene(modValidacion.ValidarMes(1), "formula alterada"), _
        "el cierre no quedo bloqueado por formula alterada"))
    res = Acumular(res, Hallazgo(Cerrar(1) <> "", "cerro con la formula rota"))
    res = Acumular(res, Hallazgo(modEstados.EstadoMes(1) <> EST_CERRADO, _
        "quedo CERRADO con formula rota"))
    ' se repara y vuelve a cerrar
    modCalculo.EscribirFormulasFila ws, f
    Application.Calculate
    res = Acumular(res, Hallazgo(Cerrar(1) = "", "no cerro tras reparar la formula"))
    T12 = res
    Exit Function
fallo:
    T12 = Err_()
End Function

' ===========================================================================
' T13 - Importacion de diciembre: fuentes validas e invalidas
' ===========================================================================
Private Function T13() As String
    Dim res As String
    Dim wb As Workbook
    Dim ruta As String
    Dim prevSeg As Long
    On Error GoTo fallo
    Preparar
    Donde "crear fuente"
    ruta = Libro().Path & "\_fuente_diciembre.xlsm"
    prevSeg = Application.AutomationSecurity
    Application.AutomationSecurity = 3
    Set wb = CrearFuente(ruta, AnioDePrueba() - 1, "J-12345678-9", APP_MONEDA, True, True)
    Donde "fuente valida"
    res = Acumular(res, Hallazgo(modImportacion.ValidarFuente(wb) = "", _
        "la fuente valida fue rechazada: " & modImportacion.ValidarFuente(wb)))

    Donde "rif"
    ' otro RIF
    wb.Sheets(HOJA_CONFIG).Range(CFG_RIF).Value = "J-00000000-0"
    res = Acumular(res, Hallazgo(Contiene(modImportacion.ValidarFuente(wb), "RIF"), _
        "no detecto RIF distinto"))
    wb.Sheets(HOJA_CONFIG).Range(CFG_RIF).Value = "J-12345678-9"

    Donde "anio"
    ' otro anio
    wb.Sheets(HOJA_CONFIG).Range(CFG_ANIO).Value = AnioDePrueba() - 2
    res = Acumular(res, Hallazgo(Contiene(modImportacion.ValidarFuente(wb), "anio"), _
        "no detecto anio distinto"))
    wb.Sheets(HOJA_CONFIG).Range(CFG_ANIO).Value = AnioDePrueba() - 1

    Donde "moneda"
    ' otra moneda
    wb.Sheets(HOJA_CONFIG).Range(CFG_MONEDA).Value = "USD"
    res = Acumular(res, Hallazgo(Contiene(modImportacion.ValidarFuente(wb), "moneda"), _
        "no detecto moneda distinta"))
    wb.Sheets(HOJA_CONFIG).Range(CFG_MONEDA).Value = APP_MONEDA

    Donde "entidad"
    ' otra entidad
    wb.Sheets(HOJA_CONFIG).Range(CFG_ENTIDAD).Value = "OTRA ENTIDAD"
    res = Acumular(res, Hallazgo(Contiene(modImportacion.ValidarFuente(wb), "entidad"), _
        "no detecto entidad distinta"))
    wb.Sheets(HOJA_CONFIG).Range(CFG_ENTIDAD).Value = "EMPRESA DE PRUEBAS C.A."

    Donde "diciembre"
    ' diciembre no cerrado
    wb.Sheets(HOJA_ESTADOS).Cells(modEstados.FilaEstado(12), 2).Value = EST_NO_ABIERTO
    res = Acumular(res, Hallazgo(Left$(modImportacion.ValidarFuente(wb), 1) = "!", _
        "acepto un diciembre sin cerrar"))
    wb.Sheets(HOJA_ESTADOS).Cells(modEstados.FilaEstado(12), 2).Value = EST_CERRADO
    wb.Sheets(HOJA_ESTADOS).Cells(modEstados.FilaEstado(12), 4).Value = Now
    wb.Sheets(HOJA_ESTADOS).Cells(modEstados.FilaEstado(12), 7).Value = "SNAP-PRUEBA"
    res = Acumular(res, Hallazgo(modImportacion.ValidarFuente(wb) = "", _
        "la fuente con evidencia de cierre fue rechazada"))

    ' fuente invalida: NUNCA se sustituye por cero (no se ejecuta importacion)
    wb.Close False
    Application.AutomationSecurity = prevSeg
    Kill_ ruta
    T13 = res
    Exit Function
fallo:
    Application.AutomationSecurity = prevSeg
    T13 = Err_()
End Function

Private Function CrearFuente(ByVal ruta As String, ByVal anio As Long, ByVal rif As String, _
                             ByVal moneda As String, ByVal cerrado As Boolean, _
                             ByVal conSaldos As Boolean) As Workbook
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long
    Dim nombres As Variant
    Kill_ ruta
    nombres = NombresMes()
    Donde "crear libro"
    Set wb = Application.Workbooks.Add()
    ' Primero se crean todas las hojas y despues se nombran por indice: cada
    ' Worksheets.Add inserta al frente y desplaza los indices anteriores.
    Do While wb.Worksheets.Count < 15
        wb.Worksheets.Add
    Loop
    Donde "nombrar hojas"
    wb.Worksheets.Item(1).Name = HOJA_CONFIG
    wb.Worksheets.Item(2).Name = HOJA_CATALOGO
    For i = 1 To 12
        wb.Worksheets.Item(2 + i).Name = anio & "-" & nombres(i - 1)
    Next i
    wb.Worksheets.Item(15).Name = HOJA_ESTADOS

    Donde "cfg s=" & wb.Sheets.Count & " w=" & wb.Worksheets.Count & _
        " p1=[" & wb.Worksheets.Item(1).Name & "] const=[" & HOJA_CONFIG & "]"
    Set ws = wb.Sheets(HOJA_CONFIG)
    Donde "config rangos"
    ws.Range(CFG_ENTIDAD).Value = "EMPRESA DE PRUEBAS C.A."
    ws.Range(CFG_RIF).Value = rif
    ws.Range(CFG_ANIO).Value = anio
    ws.Range(CFG_MONEDA).Value = moneda

    Donde "catalogo fuente"
    Set ws = wb.Sheets(HOJA_CATALOGO)
    ws.Cells(FILA_1, 1).Value = "M-0001"
    ws.Cells(FILA_1, 2).Value = "Material uno"
    ws.Cells(FILA_1, 3).Value = "UND"

    If conSaldos Then
        Donde "saldos fuente"
        Set ws = wb.Sheets(anio & "-" & nombres(11))
        ws.Cells(FILA_1, COL_CODIGO).Value = "M-0001"
        ws.Cells(FILA_1, COL_DESCRIPCION).Value = "Material uno"
        ws.Cells(FILA_1, COL_UNIDAD).Value = "UND"
        ws.Cells(FILA_1, COL_CANT_FIN).Value = 7
        ws.Cells(FILA_1, COL_VAL_FIN).Value = 700
        ws.Cells(FILA_1, COL_COSTO_FIN).Value = 100
    End If

    Donde "estados fuente"
    Set ws = wb.Sheets(HOJA_ESTADOS)
    For i = 1 To 12
        ws.Cells(modEstados.FilaEstado(i), 1).Value = anio & "-" & nombres(i - 1)
        ws.Cells(modEstados.FilaEstado(i), 2).Value = IIf(cerrado, EST_CERRADO, EST_NO_ABIERTO)
        If cerrado Then
            ws.Cells(modEstados.FilaEstado(i), 4).Value = Now
            ws.Cells(modEstados.FilaEstado(i), 7).Value = "SNAP-" & i
        End If
    Next i
    If cerrado Then
        ws.Cells(modEstados.FilaEstado(1), 2).Value = EST_CERRADO
    End If
    wb.SaveAs ruta, 52
    Set CrearFuente = wb
End Function

Private Sub Kill_(ByParam As String)
    On Error Resume Next
    If Len(ByParam) > 0 Then Kill ByParam
    On Error GoTo 0
End Sub

' ===========================================================================
' T14 - Nuevo anio: mismo saldo, origen intacto, VBA presente
' ===========================================================================
Private Function T14() As String
    Dim res As String
    Dim ruta As String
    Dim copia As Workbook
    Dim prevSeg As Long
    Dim eneOrigen As Worksheet
    Dim f As Long, k As Long
    Dim cc As Double, vv As Double
    On Error GoTo fallo
    Preparar
    Donde "diciembre cargado"
    ' diciembre cargado y marcado CERRADO directamente (sin validar los 11 meses)
    EscribirFila 12, "M-0001", 3, 300, 0, Empty, 1, 0
    Application.Calculate
    DesprotegerTodo
    For f = 1 To 12
        Libro().Sheets(HOJA_ESTADOS).Cells(modEstados.FilaEstado(f), 2).Value = EST_CERRADO
    Next f
    Application.Calculate
    Set eneOrigen = Libro().Sheets(HojaMes(1))
    eneOrigen.Cells(FILA_1, COL_CODIGO).Value = "M-0001"
    eneOrigen.Cells(FILA_1, COL_DESCRIPCION).Value = "Material uno"
    eneOrigen.Cells(FILA_1, COL_UNIDAD).Value = "UND"
    eneOrigen.Cells(FILA_1, COL_CANT_INI).Value = 99
    eneOrigen.Cells(FILA_1, COL_VAL_INI).Value = 999

    Donde "copiar libro"
    ruta = Libro().Path & "\_nuevo_anio.xlsm"
    Kill_ ruta
    Libro().SaveCopyAs ruta
    prevSeg = Application.AutomationSecurity
    Application.AutomationSecurity = 3
    Set copia = Application.Workbooks.Open(ruta, 0, False, 5)
    Donde "transferir diciembre"
    modNuevoAnio.TransferirDiciembreANuevoAnio copia, AnioDePrueba()
    copia.Save

    res = Acumular(res, Hallazgo(CLng(copia.Sheets(HOJA_CONFIG).Range(CFG_ANIO).Value2) = _
        AnioDePrueba() + 1, "el anio del libro nuevo no es el siguiente"))
    f = FilaEnHojaDe(copia, AnioDePrueba() + 1, 1, "M-0001")
    ' El libro original conserva diciembre: la copia lo renombra al anio nuevo.
    k = FilaEnMesPublica(12, "M-0001")
    res = Acumular(res, Hallazgo(f > 0, "el material no paso al enero nuevo"))
    If f > 0 And k > 0 Then
        cc = V(copia.Sheets((AnioDePrueba() + 1) & "-ENE"), f, COL_CANT_INI)
        vv = V(copia.Sheets((AnioDePrueba() + 1) & "-ENE"), f, COL_VAL_INI)
        res = Acumular(res, Hallazgo(Abs(cc - V(Libro().Sheets(HojaMes(12)), k, COL_CANT_FIN)) <= 0.00005, _
            "enero nuevo no tomo la cantidad de cierre de diciembre"))
        res = Acumular(res, Hallazgo(Abs(vv - V(Libro().Sheets(HojaMes(12)), k, COL_VAL_FIN)) <= 0.005, _
            "enero nuevo no tomo el valor de cierre de diciembre"))
    End If
    ' origen intacto
    res = Acumular(res, Hallazgo(V(eneOrigen, FILA_1, COL_CANT_INI) = 99, _
        "el libro original cambio"))
    res = Acumular(res, Hallazgo(CLng(Libro().Sheets(HOJA_CONFIG).Range(CFG_ANIO).Value2) = AnioDePrueba(), _
        "el ejercicio original cambio"))
    ' VBA presente en la copia
    res = Acumular(res, Hallazgo(TieneVBA(copia), "la copia abrio sin proyecto VBA"))

    copia.Close False
    Application.AutomationSecurity = prevSeg
    Kill_ ruta
    T14 = res
    Exit Function
fallo:
    T14 = Err_()
    On Error Resume Next
    Application.AutomationSecurity = prevSeg
    If Not copia Is Nothing Then copia.Close False
    On Error GoTo 0
End Function

Private Function FilaEnHojaDe(ByVal wb As Workbook, ByVal anio As Long, ByVal mes As Long, _
                              ByVal codigo As String) As Long
    Dim ws As Worksheet
    Dim f As Long
    Dim clave As String
    Dim nombres As Variant
    nombres = NombresMes()
    clave = NormalizaClave(codigo)
    Set ws = wb.Sheets(anio & "-" & nombres(mes - 1))
    For f = FILA_1 To FILA_FIN
        If NormalizaClave(ws.Cells(f, COL_CODIGO).Value2) = clave Then
            FilaEnHojaDe = f
            Exit Function
        End If
    Next f
End Function

Private Function TieneVBA(ByVal wb As Workbook) As Boolean
    ' HasVBProject en lugar de VBProject: no exige "Acceso al modelo de objetos
    ' de VB" ni fuerza la carga del proyecto en el VBE.
    On Error GoTo sinProyecto
    TieneVBA = wb.HasVBProject
    Exit Function
sinProyecto:
    TieneVBA = False
End Function

' ===========================================================================
' T15 - Repetir apertura / cierre / importacion: sin duplicar ni sobrescribir
' ===========================================================================
Private Function T15() As String
    Dim res As String
    Dim ws As Worksheet
    Dim f As Long
    Dim idAntes As String
    On Error GoTo fallo
    Preparar
    EscribirFila 1, "M-0001", 5, 500, 2, 200, 1, 0
    Application.Calculate
    res = Acumular(res, AbrirEnero())
    ' segunda apertura del mismo mes: no cambia nada
    res = Acumular(res, Hallazgo(Trim$(modEstados.AbrirMes(1)) = "", _
        "reabrir el mes abierto marco error"))
    res = Acumular(res, Hallazgo(modEstados.EstadoMes(1) = EST_ABIERTO, _
        "el segundo intento cambio el estado"))
    res = Acumular(res, Cerrar(1))
    idAntes = CStr(Libro().Sheets(HOJA_ESTADOS).Cells(modEstados.FilaEstado(1), 7).Value2)
    ' segundo cierre: se rechaza sin tocar la evidencia
    res = Acumular(res, Hallazgo(Contiene(Cerrar(1), "ya estaba registrado"), _
        "el cierre repetido no fue detectado"))
    res = Acumular(res, Hallazgo(CStr(Libro().Sheets(HOJA_ESTADOS) _
        .Cells(modEstados.FilaEstado(1), 7).Value2) = idAntes, _
        "el cierre repetido sobrescribio la evidencia"))
    ' un solo cierre registrado en la bitacora
    Set ws = Libro().Sheets(HOJA_BITACORA)
    f = 0
    Dim g As Long
    For g = FILA_1 To modGlobales.UltimaFilaUsada(ws, 1)
        If CStr(ws.Cells(g, 6).Value2) = HojaMes(1) And _
           CStr(ws.Cells(g, 5).Value2) = "Cierre de mes" Then f = f + 1
    Next g
    res = Acumular(res, Hallazgo(f = 1, "hay " & f & " registros de cierre para enero"))
    ' apertura de febrero repetida
    res = Acumular(res, Hallazgo(modEstados.AbrirMes(2) = "", "no abrio febrero"))
    res = Acumular(res, Hallazgo(modEstados.AbrirMes(2) = "", "el segundo intento de febrero fallo"))
    res = Acumular(res, Hallazgo(modEstados.MesAbierto() = 2, _
        "hay mas de un mes abierto tras repetir"))
    T15 = res
    Exit Function
fallo:
    T15 = Err_()
End Function

' ===========================================================================
' T16 - Respaldo fallido: no hay cierre falso
' ===========================================================================
Private Function T16() As String
    Dim res As String
    Dim fso As Object
    Dim carpeta As String
    Dim ts As Object
    On Error GoTo fallo
    Preparar
    EscribirFila 1, "M-0001", 5, 500, 0, Empty, 1, 0
    Application.Calculate
    res = Acumular(res, AbrirEnero())

    ' se convierte la carpeta de respaldos en un archivo: no se puede escribir ahi
    carpeta = modRespaldo.CarpetaRespaldos()
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(carpeta) Then
        On Error Resume Next
        fso.DeleteFolder carpeta, True
        On Error GoTo fallo
    End If
    If Not fso.FileExists(carpeta) Then
        Set ts = fso.CreateTextFile(carpeta, True)
        ts.Write "bloqueo artificial de pruebas"
        ts.Close
    End If

    res = Acumular(res, Hallazgo(Contiene(Cerrar(1), "respaldo"), _
        "el cierre no informo el fallo del respaldo"))
    res = Acumular(res, Hallazgo(modEstados.EstadoMes(1) <> EST_CERRADO, _
        "quedo CERRADO sin respaldo verificable"))
    res = Acumular(res, Hallazgo(Not modBitacora.YaRegistrado("CIERRE-" & HojaMes(1)), _
        "queda un cierre registrado sin respaldo"))

    ' se restaura la carpeta
    On Error Resume Next
    fso.DeleteFile carpeta, True
    On Error GoTo fallo
    T16 = res
    Exit Function
fallo:
    T16 = Err_()
End Function

' ===========================================================================
' T17 - Fallo a mitad de transaccion: INCIDENCIA, no exito parcial
' ===========================================================================
Private Function T17() As String
    Dim res As String
    Dim ws As Worksheet
    Dim errNum As Long
    On Error GoTo fallo
    Preparar
    EscribirFila 1, "M-0001", 5, 500, 0, Empty, 1, 0
    Application.Calculate
    res = Acumular(res, AbrirEnero())

    ' se inutiliza _INSTANTANEAS con otra clave: la escritura fallara a mitad
    Set ws = Libro().Sheets(HOJA_SNAP)
    ws.Protect Password:="clave-de-otra-persona", Contents:=True

    On Error Resume Next
    Err.Clear
    modEstados.CerrarMes 1, "RESPONSABLE DEL LIBRO", "cierre con fallo simulado"
    errNum = Err.Number
    Err.Clear
    On Error GoTo fallo
    res = Acumular(res, Hallazgo(errNum <> 0, "el cierre no propago el fallo"))

    res = Acumular(res, Hallazgo(modEstados.EstadoMes(1) <> EST_CERRADO, _
        "marco CERRADO a pesar del fallo"))
    res = Acumular(res, Hallazgo(modEstados.EstadoMes(1) = EST_INCIDENCIA, _
        "no quedo en INCIDENCIA (queda " & modEstados.EstadoMes(1) & ")"))
    res = Acumular(res, Hallazgo(Not modBitacora.YaRegistrado("CIERRE-" & HojaMes(1)), _
        "dejo un cierre registrado pese al fallo"))

    ' se restaura la hoja para el resto de las pruebas
    ws.Unprotect Password:="clave-de-otra-persona"
    ws.Protect Password:=CLAVE_PROTECCION, Contents:=True
    T17 = res
    Exit Function
fallo:
    T17 = Err_()
End Function

' ===========================================================================
' T18 - Proteccion segura antes y despues de guardar
' ===========================================================================
Private Function T18() As String
    Dim res As String
    Dim ws As Worksheet
    Dim indice As Long
    On Error GoTo fallo
    Preparar
    EscribirFila 1, "M-0001", 5, 500, 2, 200, 1, 0
    Application.Calculate
    res = Acumular(res, AbrirEnero())

    ' modo seguro: todo bloqueado y sin UserInterfaceOnly
    modEstados.ProtegerTodoParaGuardar
    For Each ws In Libro().Worksheets
        res = Acumular(res, Hallazgo(ws.ProtectContents, "hoja " & ws.Name & " sin proteccion"))
        res = Acumular(res, Hallazgo(ws.Cells(FILA_1, COL_CANT_ENT).Locked, _
            "hoja " & ws.Name & ": celda de entrada desbloqueada en modo seguro"))
    Next ws

    ' al abrir con macros se reactiva solo el mes abierto
    indice = MesAbierto()
    If indice > 0 Then
        modEstados.DesbloquearINPUT Libro().Sheets(HojaMes(indice))
        res = Acumular(res, Hallazgo(Not Libro().Sheets(HojaMes(indice)) _
            .Cells(FILA_1, COL_CANT_ENT).Locked, "el mes abierto no reactivo sus entradas"))
        res = Acumular(res, Hallazgo(Libro().Sheets(HojaMes(indice)) _
            .Cells(FILA_1, COL_VAL_VTA).Locked, "la columna calculada quedo editable"))
        res = Acumular(res, Hallazgo(Libro().Sheets(HojaMes(2)) _
            .Cells(FILA_1, COL_CANT_ENT).Locked, "otro mes quedo editable"))
    End If
    T18 = res
    Exit Function
fallo:
    T18 = Err_()
End Function

' Diagnostico de T18: paso a paso para localizar el error de compilacion.
Private Function T18Diag() As String
    Dim res As String
    Dim ws As Worksheet
    Dim indice As Long
    On Error GoTo fallo
    Preparar
    Paso "t18 preparar ok"
    EscribirFila 1, "M-0001", 5, 500, 2, 200, 1, 0
    Application.Calculate
    res = AbrirEnero()
    Paso "t18 abrirenero " & res
    modEstados.ProtegerTodoParaGuardar
    Paso "t18 proteger ok"
    For Each ws In Libro().Worksheets
        res = Acumular(res, Hallazgo(ws.ProtectContents, "sin proteccion " & ws.Name))
        res = Acumular(res, Hallazgo(ws.Cells(FILA_1, COL_CANT_ENT).Locked, "entrada desbloqueada"))
    Next ws
    Paso "t18 bucle proteccion ok"
    indice = MesAbierto()
    Paso "t18 mesabierto " & indice
    If indice > 0 Then
        modEstados.DesbloquearINPUT Libro().Sheets(HojaMes(indice))
        Paso "t18 desbloquear ok"
        res = Acumular(res, Hallazgo(Not Libro().Sheets(HojaMes(indice)) _
            .Cells(FILA_1, COL_CANT_ENT).Locked, "el mes abierto no reactivo sus entradas"))
        Paso "t18 check entradas ok"
    End If
    T18Diag = res
    Exit Function
fallo:
    T18Diag = Err_()
End Function

' ===========================================================================
' T19 - Separadores: calculo correcto y estado restaurado
' ===========================================================================
Private Function T19() As String
    Dim res As String
    Dim ws As Worksheet
    Dim f As Long
    Dim sepDecAntes As String, sepMilAntes As String
    On Error GoTo fallo
    Preparar
    EscribirFila 1, "M-0001", 3, 1, 0, 0, 1, 2
    Application.Calculate

    modEstadoExcel.CapturarEstado
    sepDecAntes = Application.DecimalSeparator
    sepMilAntes = Application.ThousandsSeparator
    Application.UseSystemSeparators = False
    Application.DecimalSeparator = ","
    Application.ThousandsSeparator = "."
    Application.Calculate
    Set ws = Libro().Sheets(HojaMes(1))
    f = FilaEnMesPublica(1, "M-0001")
    res = Acumular(res, Hallazgo(Abs(V(ws, f, COL_VAL_VTA) - 0.33) <= 0.005, _
        "con coma decimal H salio " & V(ws, f, COL_VAL_VTA)))
    res = Acumular(res, Hallazgo(Abs(V(ws, f, COL_VAL_FIN)) <= 0.005, _
        "con coma decimal L no quedo en 0"))
    res = Acumular(res, Hallazgo(Not Contiene(modValidacion.ValidarMes(1), "redondeo"), _
        "con coma decimal aparece error de redondeo"))

    modEstadoExcel.RestaurarEstado
    Application.UseSystemSeparators = True
    res = Acumular(res, Hallazgo(Application.DecimalSeparator = sepDecAntes, _
        "no restauro el separador decimal (" & Application.DecimalSeparator & ")"))
    res = Acumular(res, Hallazgo(Application.ThousandsSeparator = sepMilAntes, _
        "no restauro el separador de miles"))
    T19 = res
    Exit Function
fallo:
    T19 = Err_()
    On Error Resume Next
    modEstadoExcel.RestaurarEstado
    Application.UseSystemSeparators = True
    On Error GoTo 0
End Function

' ===========================================================================
' T20 - Resumen al corte: flujos sumados, stocks no sumados
' ===========================================================================
Private Function T20() As String
    Dim res As String
    Dim ws As Worksheet
    Dim f As Long
    Dim dic As Long
    On Error GoTo fallo
    Preparar
    EscribirFila 1, "M-0001", 10, 1000, 5, 600, 4, 1
    ' la columna I exige detalle en DETALLE_OTROS para poder cerrar el mes
    AgregarLineaDetalle 1, "M-0001", 1, "retiro", 1
    Application.Calculate
    res = Acumular(res, AbrirEnero())
    res = Acumular(res, Cerrar(1))
    ' febrero con movimientos, aun abierto (provisional)
    res = Acumular(res, Hallazgo(modEstados.AbrirMes(2) = "", "no abrio febrero"))
    EscribirFila 2, "M-0001", Empty, Empty, 2, 250, 3, 0
    Application.Calculate

    Donde "construir resumen"
    modResumen.ConstruirResumen
    Set ws = Libro().Sheets(HOJA_RESUMEN)
    f = BuscarEnResumen(ws, "M-0001", 1)
    res = Acumular(res, Hallazgo(f > 0, "M-0001 no aparece en el bloque definitivo"))
    If f > 0 Then
        ' flujos sumados: entradas = 5 (solo enero, mes cerrado)
        res = Acumular(res, Hallazgo(Abs(V(ws, f, 7) - 600) <= 0.005, _
            "entradas del corte=" & V(ws, f, 7) & " y debe ser 600"))
        ' stock final = cierre de enero (NO suma enero + febrero).
        ' enero: 15 und / 1600 Bs, costo promedio 1600/15; ventas 4 y otros 1
        ' se valoran al costo promedio -> 1600 - 426,67 - 106,67 = 1066,66
        res = Acumular(res, Hallazgo(Abs(V(ws, f, 13) - 1066.66) <= 0.005, _
            "stock final=" & V(ws, f, 13) & " y debe ser el cierre de enero (1066,66)"))
        res = Acumular(res, Hallazgo(CStr(ws.Cells(f, 14).Value2) = "OK", _
            "la conciliacion no dio OK"))
    End If
    T20 = res
    Exit Function
fallo:
    T20 = Err_()
    On Error Resume Next
    On Error GoTo 0
End Function

Private Function BuscarEnResumen(ByVal ws As Worksheet, ByVal codigo As String, _
                                 ByVal bloque As Long) As Long
    Dim f As Long
    Dim clave As String
    clave = NormalizaClave(codigo)
    For f = 7 To 40000
        If EsVacio(ws.Cells(f, 1).Value2) Then Exit For
        If NormalizaClave(ws.Cells(f, 1).Value2) = clave Then
            BuscarEnResumen = f
            Exit Function
        End If
    Next f
End Function

' ===========================================================================
' T21 - Unidades mixtas: cantidades no sumadas
' ===========================================================================
Private Function T21() As String
    Dim res As String
    Dim ws As Worksheet
    Dim f As Long
    Dim tieneU As Boolean, tieneD As Boolean
    On Error GoTo fallo
    Preparar
    EscribirFila 1, "M-0001", 10, 1000, 0, Empty, 4, 0      ' UND
    EscribirFila 1, "M-0002", 6, 600, 0, Empty, 2, 0        ' KG
    Application.Calculate
    res = Acumular(res, AbrirEnero())
    res = Acumular(res, Cerrar(1))
    Donde "construir resumen"
    modResumen.ConstruirResumen
    Set ws = Libro().Sheets(HOJA_RESUMEN)
    ' hay un total por cada unidad
    f = BuscarTotalPorUnidad(ws, "UND")
    tieneU = (f > 0)
    f = BuscarTotalPorUnidad(ws, "KG")
    tieneD = (f > 0)
    res = Acumular(res, Hallazgo(tieneU And tieneD, _
        "el resumen no separa los totales por unidad de medida"))
    ' no existe un total de cantidades mezcladas
    res = Acumular(res, Hallazgo(Not Contiene(TextoRango(ws), "Total de cantidades"), _
        "hay un total de cantidades que mezcla unidades"))
    T21 = res
    Exit Function
fallo:
    T21 = Err_()
End Function

Private Function BuscarTotalPorUnidad(ByVal ws As Worksheet, ByVal unidad As String) As Long
    Dim f As Long
    Dim ultima As Long
    ' Se recorre hasta la ultima fila usada: el resumen deja filas en blanco
    ' entre bloques y un corte ahi esconderia el bloque de totales.
    On Error Resume Next
    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    On Error GoTo 0
    If ultima < 1 Then Exit Function
    For f = 1 To ultima
        If CStr(ws.Cells(f, 1).Value2) = unidad And _
           EsNumeroReal(ws.Cells(f, 2).Value2) Then
            BuscarTotalPorUnidad = f
            Exit Function
        End If
    Next f
End Function

Private Function TextoRango(ByVal ws As Worksheet) As String
    Dim c As Range
    On Error Resume Next
    For Each c In ws.Range("A1:O20000")
        If Not IsEmpty(c.Value2) Then
            If VarType(c.Value2) = vbString Then TextoRango = TextoRango & c.Value2 & "|"
        End If
    Next c
    On Error GoTo 0
End Function

' ===========================================================================
' T22 - Detalle ausente o descuadrado: bloqueo y sin doble contabilizacion
' ===========================================================================
Private Function T22() As String
    Dim res As String
    Dim ws As Worksheet
    Dim f As Long
    On Error GoTo fallo
    Preparar
    EscribirFila 1, "M-0001", 10, 1000, 0, Empty, 4, 3     ' I=3
    Application.Calculate
    res = Acumular(res, AbrirEnero())
    Set ws = Libro().Sheets(HojaMes(1))
    f = FilaEnMesPublica(1, "M-0001")

    res = Acumular(res, Hallazgo(Contiene(modValidacion.ValidarMes(1), "I>0 sin detalle"), _
        "I=3 sin detalle no bloqueo el cierre"))
    res = Acumular(res, Hallazgo(Cerrar(1) <> "", "cerro con detalle ausente"))
    res = Acumular(res, Hallazgo(modEstados.EstadoMes(1) <> EST_CERRADO, _
        "quedo CERRADO con detalle ausente"))

    ' sin doble contabilizacion: L sale solo de H y J (formula de la hoja)
    res = Acumular(res, Hallazgo(Contiene(CStr(ws.Cells(f, COL_VAL_FIN).Formula), _
        "H" & f & "-J" & f), "L no depende unicamente de H y J"))
    res = Acumular(res, Hallazgo(Abs(V(ws, f, COL_VAL_FIN) - _
        (V(ws, f, COL_VAL_INI) + V(ws, f, COL_VAL_ENT) - V(ws, f, COL_VAL_VTA) - _
         V(ws, f, COL_VAL_OTR))) <= 0.005, "L no cuadra con D+F-H-J"))

    ' detalle completo: concilia y se puede cerrar
    AgregarLineaDetalle 1, "M-0001", 5, "retiro", 1
    AgregarLineaDetalle 1, "M-0001", 6, "merma", 1
    AgregarLineaDetalle 1, "M-0001", 7, "autoconsumo", 1
    res = Acumular(res, Hallazgo(modValidacion.ValidarMes(1) = "", _
        "el detalle completo no concilio"))
    res = Acumular(res, Hallazgo(Cerrar(1) = "", "no cerro con el detalle completo"))
    ' el importe del detalle se reparte a costo promedio y suma J
    modDetalle.ConciliarImporteOtros 1, "M-0001"
    Set ws = Libro().Sheets(HOJA_DETALLE)
    res = Acumular(res, Hallazgo(Abs(ImporteDetalle(ws) - V(Libro().Sheets(HojaMes(1)), _
        FilaEnMesPublica(1, "M-0001"), COL_VAL_OTR)) <= 0.005, _
        "el importe repartido no coincide con J"))
    T22 = res
    Exit Function
fallo:
    T22 = Err_()
End Function

Private Function ImporteDetalle(ByVal ws As Worksheet) As Double
    Dim f As Long
    For f = FILA_DET_1 To FILA_DET_FIN
        If Not EsVacio(ws.Cells(f, DET_MES).Value2) Then
            ImporteDetalle = ImporteDetalle + CDbl(0 + ws.Cells(f, DET_IMPORTE).Value2)
        End If
    Next f
End Function

' ===========================================================================
' T23 - Cantidades fraccionarias, centavos pequenos y error de redondeo
' ===========================================================================
Private Function T23() As String
    Dim res As String
    Dim ws As Worksheet
    Dim f As Long
    Dim cm As Variant
    Dim okMoneda As Boolean
    On Error GoTo fallo
    Preparar
    ' fraccion con mas de 2 decimales
    EscribirFila 1, "M-0001", 1.234567, 10, 0, Empty, 0.5, 0
    Application.Calculate
    Set ws = Libro().Sheets(HojaMes(1))
    f = FilaEnMesPublica(1, "M-0001")
    res = Acumular(res, Hallazgo(Abs(V(ws, f, COL_CANT_INI) - 1.234567) < 0.0000005, _
        "se perdio precision en la cantidad inicial"))
    res = Acumular(res, Hallazgo(CStr(ws.Cells(f, COL_CANT_INI).NumberFormat) = MascaraCant(), _
        "la cantidad fraccionaria no usa la mascara de cantidades vigente"))
    res = Acumular(res, Hallazgo(Abs(V(ws, f, COL_CANT_FIN) - 0.734567) < 0.0000005, _
        "K perdio precision"))

    Donde "centavos"
    ' centavo pequeno: 0,01 en costo no debe descartarse
    EscribirFila 2, "M-0002", 1, 0.01, 1, 0.01, 1, 0
    Application.Calculate
    Set ws = Libro().Sheets(HojaMes(2))
    f = FilaEnMesPublica(2, "M-0002")
    res = Acumular(res, Hallazgo(Abs(V(ws, f, COL_COSTO_PROM) - 0.01) <= 0.000001, _
        "el costo promedio con centavos salio " & V(ws, f, COL_COSTO_PROM)))
    res = Acumular(res, Hallazgo(modValidacion.ValidarMes(2) = "", _
        "los centavos pequenos marcaron error"))

    Donde "redondeo"
    ' error de redondeo visible: importe de salida mayor que el disponible
    EscribirFila 3, "M-0003", 1, 0.01, 0, Empty, 1, 0
    Set ws = Libro().Sheets(HojaMes(3))
    f = FilaEnMesPublica(3, "M-0003")
    ' fuerza un valor de salida mayor que el valor disponible
    ws.Cells(f, COL_VAL_ENT).Value = 0.01
    ws.Cells(f, COL_CANT_VTA).Value = 1
    ws.Cells(f, COL_VAL_VTA).Formula = "=1.02"
    Application.Calculate
    res = Acumular(res, Hallazgo(V(ws, f, COL_VAL_FIN) < -0.00005, _
        "no quedo un valor final negativo para evidenciar el error"))
    res = Acumular(res, Hallazgo(Contiene(modValidacion.ValidarMes(3), "REDONDEO"), _
        "no reporto el error de redondeo"))

    Donde "mascara"
    ' Regla contable: todo importe en Bs. se presenta con exactamente 2 decimales.
    okMoneda = True
    For Each cm In Array(COL_VAL_INI, COL_VAL_ENT, COL_VAL_VTA, COL_VAL_OTR, COL_VAL_FIN)
        If CStr(ws.Cells(f, cm).NumberFormat) <> FMT_MONEDA Then okMoneda = False
    Next cm
    res = Acumular(res, Hallazgo(okMoneda, _
        "un importe en Bs. no quedo con 2 decimales"))

    ' Mascara numerica de cantidades: eleccion del usuario, solo presentacion.
    Libro().Sheets(HOJA_CONFIG).Range(CFG_MASCARA_CANT).Value = MASCARA_CANT_SIN_DEC
    res = Acumular(res, Hallazgo(MascaraCant() = MASCARA_CANT_SIN_DEC, _
        "no tomo la mascara sin decimales"))
    modUI.AplicarFormatosNumericos
    res = Acumular(res, Hallazgo(CStr(ws.Cells(f, COL_CANT_INI).NumberFormat) = MASCARA_CANT_SIN_DEC, _
        "no aplico la mascara sin decimales en los meses"))
    res = Acumular(res, Hallazgo(CStr(ws.Cells(f, COL_VAL_FIN).NumberFormat) = FMT_MONEDA, _
        "el importe en Bs. dejo de tener 2 decimales al cambiar la mascara"))
    res = Acumular(res, Hallazgo(Abs(V(ws, f, COL_CANT_INI) - 1) < 0.0000005, _
        "la mascara altero el valor guardado de la cantidad"))
    Libro().Sheets(HOJA_CONFIG).Range(CFG_MASCARA_CANT).Value = "mascara no valida"
    res = Acumular(res, Hallazgo(MascaraCant() = FMT_CANT, _
        "no volvio a la mascara por defecto ante un valor invalido"))
    Libro().Sheets(HOJA_CONFIG).Range(CFG_MASCARA_CANT).Value = FMT_CANT
    modUI.AplicarFormatosNumericos
    res = Acumular(res, Hallazgo(CStr(ws.Cells(f, COL_CANT_INI).NumberFormat) = FMT_CANT, _
        "no restauro la mascara por defecto"))

    T23 = res
    Exit Function
fallo:
    T23 = Err_()
End Function

' ===========================================================================
' T24 - Compilar y probar todos los botones
' ===========================================================================
Private Function T24() As String
    Dim res As String
    Dim botones As Variant
    Dim b As Variant
    Dim errNum As Long
    Dim errTxt As String
    On Error GoTo fallo
    Preparar
    modUI.PruebaSilencioso = True
    botones = Array("BtnAyudaMenu", "BtnValidarConfiguracion", "BtnValidarMes", _
                    "BtnAltaMaterial", "BtnDetalleOtros", "BtnResumenAnual", _
                    "BtnRespaldo", "BtnConfirmarAperturaEnero", "BtnAbrirMesSiguiente", _
                    "BtnCerrarMesActual", "BtnImportarDiciembre", "BtnConfirmarImportacion", _
                    "BtnNuevoAnio", "BtnCapturarConfiguracion")
    For Each b In botones
        Err.Clear
        On Error Resume Next
        Application.Run "'" & Libro().Name & "'!" & CStr(b)
        errNum = Err.Number
        errTxt = Err.Description
        On Error GoTo 0
        If errNum <> 0 Then
            res = Acumular(res, "el boton " & CStr(b) & " fallo: " & errTxt)
        End If
    Next b
    ' referencias: todo el libro usa vinculacion tardia (CreateObject), no
    ' referencias enlazadas que puedan quedar en MISSING.
    res = Acumular(res, Hallazgo(Not HayReferenciaRota(), _
        "se detecto una referencia enlazada rota"))
    ' la configuracion sigue valida tras los botones
    res = Acumular(res, Hallazgo(modConfiguracion.ValidarConfiguracion() = "", _
        "los botones rompieron la configuracion"))
    T24 = res
    Exit Function
fallo:
    T24 = Err_()
End Function

Private Function HayReferenciaRota() As Boolean
    ' Toda la interop con COM externo (FileSystemObject, ADODB, .NET) es tardia.
    ' Si alguna estuviera enlazada temprana y faltara, esta linea fallaria.
    On Error Resume Next
    CreateObject("Scripting.FileSystemObject")
    If Err.Number <> 0 Then HayReferenciaRota = True
    Err.Clear
    CreateObject("Scripting.Dictionary")
    If Err.Number <> 0 Then HayReferenciaRota = True
    Err.Clear
    On Error GoTo 0
End Function

' ===========================================================================
' T25 - 5.000 materiales: tiempos
' ===========================================================================
Private Function T25() As String
    Dim res As String
    Dim t0 As Double
    Dim tApertura As Double, tValidacion As Double, tCierre As Double
    Dim i As Long
    Dim ws As Worksheet
    Dim texto As String
    On Error GoTo fallo
    Preparar
    Set ws = modCatalogo.HojaCatalogo()
    If ws.ProtectContents Then ws.Unprotect Password:=CLAVE_PROTECCION
    For i = FILA_1 To FILA_FIN
        ws.Cells(i, 1).Value = "MAT-" & Format$(i - FILA_1 + 1, "00000")
        ws.Cells(i, 2).Value = "Material de prueba numero " & (i - FILA_1 + 1)
        ws.Cells(i, 3).Value = Choose(1 + ((i - FILA_1) Mod 3), "UND", "KG", "LT")
        ws.Cells(i, 4).Value = "Si"
    Next i
    InvalidarIndice

    Set ws = Libro().Sheets(HojaMes(1))
    If ws.ProtectContents Then ws.Unprotect Password:=CLAVE_PROTECCION
    For i = FILA_1 To FILA_FIN
        ws.Cells(i, COL_CODIGO).Value = "MAT-" & Format$(i - FILA_1 + 1, "00000")
        ws.Cells(i, COL_DESCRIPCION).Value = "Material de prueba numero " & (i - FILA_1 + 1)
        ws.Cells(i, COL_UNIDAD).Value = Choose(1 + ((i - FILA_1) Mod 3), "UND", "KG", "LT")
        ws.Cells(i, COL_CANT_INI).Value = 10 + ((i - FILA_1) Mod 7)
        ws.Cells(i, COL_VAL_INI).Value = (10 + ((i - FILA_1) Mod 7)) * 100
        ws.Cells(i, COL_CANT_ENT).Value = (i - FILA_1) Mod 3
        ws.Cells(i, COL_VAL_ENT).Value = ((i - FILA_1) Mod 3) * 100
        ws.Cells(i, COL_CANT_VTA).Value = (i - FILA_1) Mod 5
        ' I=0 en la carga masiva: T22 cubre la regla de DETALLE_OTROS y aqui
        ' solo se miden tiempos con 5.000 renglones.
        ws.Cells(i, COL_CANT_OTR).Value = 0
    Next i
    Application.Calculate
    Paso "T25 mes cargado y calculado"

    t0 = Timer
    modEstados.ConfirmarAperturaEnero
    tApertura = Timer - t0
    texto = ""
    Paso "T25 apertura " & Format$(tApertura, "0.0") & "s"

    t0 = Timer
    texto = modValidacion.ValidarMes(1)
    tValidacion = Timer - t0
    Paso "T25 validacion " & Format$(tValidacion, "0.0") & "s"

    t0 = Timer
    texto = modEstados.CerrarMes(1, "RESPONSABLE DEL LIBRO", "cierre con 5.000 materiales")
    tCierre = Timer - t0
    Paso "T25 cierre " & Format$(tCierre, "0.0") & "s"

    mInforme = mInforme & "      T25 tiempos: apertura=" & Format$(tApertura, "0.00") & _
        "s validacion=" & Format$(tValidacion, "0.00") & "s cierre=" & _
        Format$(tCierre, "0.00") & "s (5.000 materiales)" & vbCrLf

    res = Acumular(res, Hallazgo(modEstados.EstadoMes(1) = EST_CERRADO, _
        "no cerro con 5.000 materiales: " & texto))
    res = Acumular(res, Hallazgo(tApertura < 90, "apertura muy lenta: " & tApertura & "s"))
    res = Acumular(res, Hallazgo(tValidacion < 90, "validacion muy lenta: " & tValidacion & "s"))
    res = Acumular(res, Hallazgo(tCierre < 120, "cierre muy lento: " & tCierre & "s"))
    T25 = res
    Exit Function
fallo:
    T25 = Err_()
End Function

' ===========================================================================
' T26 - Exportacion para revision visual (capturas/PDF)
' ===========================================================================
Private Function T26() As String
    Dim res As String
    Dim ws As Worksheet
    Dim ruta As String
    Dim fso As Object
    On Error GoTo fallo
    Preparar
    EscribirFila 1, "M-0001", 10, 1000, 5, 600, 4, 1
    Application.Calculate
    res = Acumular(res, AbrirEnero())
    Set ws = Libro().Sheets(HojaMes(1))
    ruta = Libro().Path & "\PDF_REVISION_MES_ABIERTO.pdf"
    On Error Resume Next
    ws.ExportAsFixedFormat 0, ruta
    On Error GoTo fallo
    Set fso = CreateObject("Scripting.FileSystemObject")
    res = Acumular(res, Hallazgo(fso.FileExists(ruta), _
        "no se genero el PDF de revision visual"))
    ' legibilidad: encabezados, totales y costos visibles
    res = Acumular(res, Hallazgo(CStr(ws.Cells(FILA_ENCABEZADO, COL_COSTO_PROM).Value2) <> "", _
        "falta el encabezado del costo promedio"))
    res = Acumular(res, Hallazgo(CStr(ws.Cells(FILA_TOTAL, 1).Value2) <> "", _
        "falta la etiqueta del total independiente de filtros"))
    res = Acumular(res, Hallazgo(CStr(ws.Cells(FILA_TOTAL, COL_COSTO_PROM).Value2) = "no se totaliza", _
        "el costo promedio no esta marcado como no totalizable"))
    res = Acumular(res, Hallazgo(CStr(Libro().Sheets(HOJA_CONFIG).Range(CFG_METODO).Value2) = APP_METODO, _
        "falta el metodo en CONFIGURACION"))
    T26 = res
    Exit Function
fallo:
    T26 = Err_()
End Function

' ===========================================================================
' T27 - Archivo limpio: sin enlaces externos ni marcadores
' ===========================================================================
Private Function T27() As String
    Dim res As String
    Dim lnk As Variant
    Dim n As Name
    Dim texto As String
    Dim ws As Worksheet
    On Error GoTo fallo
    Preparar
    Donde "vinculos"
    ' sin vinculos a otros libros
    lnk = Empty
    On Error Resume Next
    lnk = Libro().LinkSources(1)   ' xlExcelLinks
    On Error GoTo fallo
    res = Acumular(res, Hallazgo(IsEmpty(lnk), _
        "el libro tiene enlaces externos a otros archivos"))
    Donde "nombres"
    ' sin nombres con referencia externa
    For Each n In Libro().Names
        If InStr(n.RefersTo, "[") > 0 Then
            res = Acumular(res, "nombre con referencia externa: " & n.Name)
        End If
    Next n
    Donde "marcadores"
    ' sin marcadores de trabajo
    For Each ws In Libro().Worksheets
        If ws.Name <> HOJA_DETALLE Then
            texto = texto & CStr(ws.Range("A1").Value2) & " " & CStr(ws.Range("A2").Value2)
        End If
    Next ws
    res = Acumular(res, Hallazgo(Not TieneMarcador(texto), _
        "quedaron marcadores de trabajo en el libro"))
    Donde "bitacora"
    ' la bitacora no guarda credenciales
    res = Acumular(res, Hallazgo(Not Contiene(TextoBitacora(), "clave=") And _
        Not Contiene(TextoBitacora(), "password"), _
        "la bitacora guarda credenciales"))
    T27 = res
    Exit Function
fallo:
    T27 = Err_()
End Function

Private Function TextoBitacora() As String
    Dim ws As Worksheet
    Dim f As Long
    Set ws = Libro().Sheets(HOJA_BITACORA)
    For f = FILA_1 To modGlobales.UltimaFilaUsada(ws, 1)
        TextoBitacora = TextoBitacora & CStr(ws.Cells(f, 6).Value2) & " " & _
            CStr(ws.Cells(f, 9).Value2) & "|"
    Next f
End Function

' ===========================================================================
' T28 - Trazabilidad normativa
' ===========================================================================
Private Function T28() As String
    Dim res As String
    Dim texto As String
    Dim ws As Worksheet
    On Error GoTo fallo
    Preparar
    Set ws = Libro().Sheets(HOJA_AYUDA)
    texto = ""
    Dim f As Long
    For f = 1 To 60
        texto = texto & CStr(ws.Cells(f, 1).Value2) & vbCrLf
    Next f
    Set ws = Libro().Sheets(HOJA_CONFIG)
    For f = 1 To 30
        texto = texto & CStr(ws.Cells(f, 1).Value2) & vbCrLf & _
            CStr(ws.Cells(f, 2).Value2) & vbCrLf
    Next f
    texto = texto & AvisoArt177()

    res = Acumular(res, Hallazgo(Contiene(texto, "Art. 177"), _
        "no cita el articulo 177"))
    res = Acumular(res, Hallazgo(Contiene(texto, "2.507"), _
        "no cita el decreto 2.507"))
    res = Acumular(res, Hallazgo(Contiene(texto, "G.O.E."), _
        "no cita la Gaceta Oficial Extraordinaria"))
    res = Acumular(res, Hallazgo(Contiene(texto, "5.662"), _
        "no cita el numero de Gaceta 5.662"))
    res = Acumular(res, Hallazgo(Contiene(texto, "24-09-2003") Or Contiene(texto, "24/09/2003"), _
        "no cita la fecha de la Gaceta"))
    res = Acumular(res, Hallazgo(Contiene(texto, "Administracion Tributaria"), _
        "no advierte que la autorizacion la da la Administracion Tributaria"))
    res = Acumular(res, Hallazgo(Contiene(texto, "88") And Contiene(texto, "89"), _
        "no cita los articulos 88 y 89 del LISLR sobre conservacion"))
    res = Acumular(res, Hallazgo(Contiene(texto, "art. 64"), _
        "no cita el articulo 64 del Reglamento para destrucciones"))
    res = Acumular(res, Hallazgo(Contiene(texto, APP_URL), "no incluye la URL del metodo"))
    T28 = res
    Exit Function
fallo:
    T28 = Err_()
End Function
