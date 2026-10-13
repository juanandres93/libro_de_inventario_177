VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

' Eventos del libro. Ninguna accion de usuario queda a medias: se valida y se
' revierte. Al abrir y al guardar se coloca el libro en modo seguro (todo
' bloqueado, sin UserInterfaceOnly) para que tambien funcione sin macros; al
' abrir CON macros se habilitan solo las celdas de entrada del mes abierto.

Private Sub Workbook_Open()
    Dim problemas As String
    Dim indice As Long
    Dim mensaje As String
    On Error GoTo fallo
    MarcarSesion True
    modEstados.DesprotegerTodoParaEdicion
    modUI.AplicarFormatosNumericos
    modEstados.ProtegerTodoParaGuardar
    indice = MesAbierto()
    If indice > 0 Then modEstados.DesbloquearINPUT Libro().Sheets(HojaMes(indice))
    problemas = modConfiguracion.ValidarConfiguracion()
    If problemas <> "" Then
        AvisoLibro "El archivo se abrio en modo protegido." & vbCrLf & vbCrLf & _
            "CONFIGURACION incompleta:" & vbCrLf & problemas & vbCrLf & vbCrLf & _
            "Use el boton 'Capturar configuracion' para completar la identidad del ejercicio.", _
            vbExclamation
        Exit Sub
    End If
    AvisoLibro "Libro de inventarios " & APP_VERSION & " - " & APP_METODO & vbCrLf & _
        "Ejercicio " & HojaMes(1) & " | " & modUI.TextoMesAbierto() & _
        " (abierto) | cerrado hasta " & modUI.TextoUltimoCerrado() & vbCrLf & vbCrLf & _
        AvisoArt177(), vbInformation
    Exit Sub
fallo:
    mensaje = Err.Description
    On Error Resume Next
    modEstados.ProtegerTodoParaGuardar
    On Error GoTo 0
    AvisoLibro "Apertura del libro: " & mensaje, vbExclamation
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim mensaje As String
    On Error GoTo fallo
    modEstados.DesprotegerTodoParaEdicion
    modUI.AplicarFormatosNumericos
    modEstados.ProtegerTodoParaGuardar
    Exit Sub
fallo:
    mensaje = Err.Description
    On Error Resume Next
    modEstados.ProtegerTodoParaGuardar
    On Error GoTo 0
    Cancel = True
    AvisoLibro "No se pudo preparar el libro antes de guardar: " & mensaje & _
        vbCrLf & "El guardado se cancela para no dejar el archivo editable sin macros.", vbCritical
End Sub

' Tras guardar, el mes abierto vuelve a admitir sus celdas de entrada.
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    Dim indice As Long
    On Error Resume Next
    If Success Then
        indice = MesAbierto()
        If indice > 0 Then modEstados.DesbloquearINPUT Libro().Sheets(HojaMes(indice))
    End If
    On Error GoTo 0
End Sub

' Control de entrada y pegado: si el bloque tiene texto en columnas numericas,
' negativos, codigos fuera del catalogo o duplicados, se revierte TODO el
' cambio (nunca queda una carga parcial).
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim problemas As String
    Dim revertido As Boolean
    Dim zona As String
    On Error GoTo fin
    If Reentrada Then Exit Sub
    zona = ""
    If EsHojaMes(Sh.Name) Then
        If Application.Intersect(Target, _
             Sh.Range(Sh.Cells(FILA_1, 1), Sh.Cells(FILA_TOTAL, 17))) Is Nothing Then Exit Sub
        zona = "mes"
    ElseIf Sh.Name = HOJA_CATALOGO Then
        If Application.Intersect(Target, _
             Sh.Range(Sh.Cells(FILA_1, 1), Sh.Cells(FILA_FIN, 6))) Is Nothing Then Exit Sub
        problemas = modCatalogo.ValidarCatalogo()
        zona = "catalogo"
    Else
        Exit Sub
    End If
    If zona = "mes" Then problemas = modCatalogo.ValidarBloquePegado(Sh, Target)
    If problemas = "" Then Exit Sub
    On Error Resume Next
    revertido = Target.Undo
    On Error GoTo fin
    modBitacora.Registrar "Cambio rechazado", Sh.Name & " " & Target.Address(False, False), _
        "BLOQUEADO", problemas, "", "", "", ""
    If Reentrada Then Exit Sub
    Entrar
    AvisoLibro "No se acepto el cambio en " & Sh.Name & " (" & Target.Address(False, False) & ")." & _
        vbCrLf & vbCrLf & problemas & vbCrLf & vbCrLf & _
        IIf(revertido, "El bloque se revirtio por completo.", _
            "Use Ctrl+Z para revertir el bloque completo."), vbExclamation
    Salir
fin:
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    MarcarSesion False
    On Error GoTo 0
End Sub

Private Sub AvisoLibro(ByVal texto As String, ByVal botones As VbMsgBoxStyle)
    On Error Resume Next
    Call modUI.Aviso(texto, botones, APP_NOMBRE)
    On Error GoTo 0
End Sub