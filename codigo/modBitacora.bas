Attribute VB_Name = "modBitacora"
Option Explicit

' BITACORA: trazabilidad modificable dentro de Excel. No es firma
' autenticada ni auditoria criptografica independiente, y no guarda secretos.

Private Const COL_ID As Long = 1
Private Const COL_FECHA As Long = 2
Private Const COL_ZONA As Long = 3
Private Const COL_USUARIO As Long = 4
Private Const COL_ACCION As Long = 5
Private Const COL_OBJETO As Long = 6
Private Const COL_RESULTADO As Long = 7
Private Const COL_VERSION As Long = 8
Private Const COL_MOTIVO As Long = 9
Private Const COL_PREVIO As Long = 10
Private Const COL_NUEVO As Long = 11
Private Const COL_PERIODO As Long = 12

Public Sub Registrar(ByVal accion As String, ByVal objeto As String, _
                     ByVal resultado As String, ByVal motivo As String, _
                     ByVal previo As Variant, ByVal nuevo As Variant, _
                     ByVal periodo As String, ByVal idOperacion As String)
    Dim ws As Worksheet
    Dim fila As Long
    Dim protegido As Boolean
    On Error GoTo sinBitacora
    Set ws = Libro().Sheets(HOJA_BITACORA)
    protegido = (ws.ProtectContents Or ws.ProtectDrawingObjects)
    If protegido Then ws.Unprotect Password:=CLAVE_PROTECCION
    fila = UltimaFilaUsada(ws, COL_ID) + 1
    If fila < FILA_1 Then fila = FILA_1
    ws.Cells(fila, COL_ID).Value = IIf(idOperacion = "", AhoraID(), idOperacion)
    ws.Cells(fila, COL_FECHA).Value = Now
    ws.Cells(fila, COL_FECHA).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Cells(fila, COL_ZONA).Value = ZonaHoraria()
    ws.Cells(fila, COL_USUARIO).Value = UsuarioOperativo()
    ws.Cells(fila, COL_ACCION).Value = accion
    ws.Cells(fila, COL_OBJETO).Value = objeto
    ws.Cells(fila, COL_RESULTADO).Value = resultado
    ws.Cells(fila, COL_VERSION).Value = APP_VERSION
    ws.Cells(fila, COL_MOTIVO).Value = motivo
    If Not EsVacio(previo) Then ws.Cells(fila, COL_PREVIO).Value = CStr(previo)
    If Not EsVacio(nuevo) Then ws.Cells(fila, COL_NUEVO).Value = CStr(nuevo)
    If Not EsVacio(periodo) Then ws.Cells(fila, COL_PERIODO).Value = periodo
    If protegido Then ws.Protect Password:=CLAVE_PROTECCION, UserInterfaceOnly:=True
    Exit Sub
sinBitacora:
    ' La bitacora nunca debe impedir la operacion: si falla, la operacion sigue
    ' y el error original (no esta) se propaga desde quien la invoco.
End Sub

Public Function AhoraID() As String
    AhoraID = SelloTiempo() & "-" & Format$(Int(Rnd() * 900) + 100, "000")
End Function

Public Function ZonaHoraria() As String
    ' VBA no expone la zona horaria del sistema; se registra la zona del
    ' equipo (variable TZ o nombre de zona de Windows) sin inventar datos.
    Dim sh As Object
    Dim zona As String
    On Error GoTo sinZona
    zona = Trim$(Environ$("TZ"))
    If zona = "" Then
        Set sh = CreateObject("WScript.Shell")
        zona = Trim$(CStr(sh.RegRead( _
            "HKLM\SYSTEM\CurrentControlSet\Control\TimeZoneInformation\TimeZoneKeyName")))
    End If
    If zona = "" Then zona = "no expuesta por el sistema"
    ZonaHoraria = zona
    Exit Function
sinZona:
    ZonaHoraria = "desconocida"
End Function

Public Function UsuarioOperativo() As String
    On Error GoTo sinUsuario
    Dim entorno As String
    entorno = Environ$("USERNAME")
    If entorno <> "" Then
        UsuarioOperativo = entorno & "@" & Environ$("COMPUTERNAME")
    Else
        UsuarioOperativo = Application.UserName
    End If
    Exit Function
sinUsuario:
    UsuarioOperativo = "desconocido"
End Function

' Operacion idempotente: si ya existe un evento de exito con el mismo ID,
' no se repite (botones que se pulsan dos veces).
Public Function YaRegistrado(ByVal idOperacion As String) As Boolean
    Dim ws As Worksheet
    Dim f As Long
    On Error GoTo no
    Set ws = Libro().Sheets(HOJA_BITACORA)
    For f = FILA_1 To UltimaFilaUsada(ws, COL_ID)
        If CStr(ws.Cells(f, COL_ID).Value2) = idOperacion Then
            YaRegistrado = True
            Exit Function
        End If
    Next f
no:
    YaRegistrado = False
End Function

Public Sub RegistrarCambio(ByVal accion As String, ByVal objeto As String, _
                           ByVal previo As Variant, ByVal nuevo As Variant, _
                           ByVal motivo As String)
    Registrar accion, objeto, "OK", motivo, previo, nuevo, PeriodoActual(), ""
End Sub