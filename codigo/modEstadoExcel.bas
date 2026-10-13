Attribute VB_Name = "modEstadoExcel"
Option Explicit

' Captura y restaura el estado previo de Excel. Se usa en toda operacion
' critica: si algo falla, el estado vuelve a ser el que era.

Private mGuardado As Boolean
Private mPrevEventos As Boolean
Private mPrevCalculo As Long
Private mPrevPantalla As Boolean
Private mPrevAlertas As Boolean
Private mPrevSepDecimal As Variant
Private mPrevSepMiles As Variant

Public Sub CapturarEstado()
    If mGuardado Then Exit Sub
    mGuardado = True
    With Application
        mPrevEventos = .EnableEvents
        mPrevCalculo = .Calculation
        mPrevPantalla = .ScreenUpdating
        mPrevAlertas = .DisplayAlerts
    End With
    On Error Resume Next
    mPrevSepDecimal = Application.DecimalSeparator
    mPrevSepMiles = Application.ThousandsSeparator
    On Error GoTo 0
End Sub

Public Sub RestaurarEstado()
    If Not mGuardado Then Exit Sub
    mGuardado = False
    On Error Resume Next
    With Application
        .DisplayAlerts = mPrevAlertas
        .ScreenUpdating = mPrevPantalla
        .EnableEvents = mPrevEventos
        .Calculation = mPrevCalculo
        Application.DecimalSeparator = mPrevSepDecimal
        Application.ThousandsSeparator = mPrevSepMiles
    End With
    On Error GoTo 0
End Sub

Public Sub SinPantalla()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
End Sub

Public Sub ConPantalla(Optional ByVal recalcular As Boolean = True)
    Dim calc As Long
    On Error Resume Next
    If recalcular Then
        Application.Calculation = xlCalculationAutomatic
        Application.Calculate
    End If
    calc = Application.Calculation
    On Error GoTo 0
    RestaurarEstado
End Sub

' Registra un error con contexto y relanza, para que la operacion no parezca
' exitosa. No se usa On Error Resume Next general.
Public Sub PropagarError(ByVal contexto As String, ByVal modulo As String)
    Dim numero As Long
    Dim texto As String
    numero = Err.Number
    texto = Err.Description
    modBitacora.Registrar "ERROR", contexto & " (" & modulo & ")", _
        "Fallo", "numero=" & numero & " " & texto, "", "", "", ""
    Err.Raise numero, modulo & "." & contexto, texto
End Sub