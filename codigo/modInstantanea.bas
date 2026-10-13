Attribute VB_Name = "modInstantanea"
Option Explicit

' _INSTANTANEAS: al cerrar un mes se guardan los VALORES de cierre, la copia
' del maestro y el detalle, con identificador, version y fecha. Un mes cerrado
' no depende de hojas vivas ni del catalogo actual.

Private Const SNAP_ID As Long = 1
Private Const SNAP_MES As Long = 2
Private Const SNAP_FECHA As Long = 3
Private Const SNAP_CODIGO As Long = 4
Private Const SNAP_DESCRIPCION As Long = 5
Private Const SNAP_UNIDAD As Long = 6
Private Const SNAP_CANT_FIN As Long = 7
Private Const SNAP_VAL_FIN As Long = 8
Private Const SNAP_COSTO_PROM As Long = 9

Public Function HojaInstantanea() As Worksheet
    Set HojaInstantanea = Libro().Sheets(HOJA_SNAP)
End Function

Public Sub GrabarInstantanea(ByVal indice As Long, ByVal identificador As String)
    Dim ws As Worksheet
    Dim mes As Worksheet
    Dim f As Long
    Dim inicio As Long
    Dim primerDato As Long
    Dim protegido As Boolean
    Dim mesTexto As String
    Set ws = HojaInstantanea()
    Set mes = Libro().Sheets(HojaMes(indice))
    mesTexto = HojaMes(indice)
    protegido = (ws.ProtectContents Or ws.ProtectDrawingObjects)
    If protegido Then ws.Unprotect Password:=CLAVE_PROTECCION
    inicio = ws.Cells(ws.Rows.Count, SNAP_ID).End(xlUp).Row + 1
    If inicio < 2 Then inicio = 2
    ws.Cells(inicio, SNAP_ID).Value = identificador
    ws.Cells(inicio, SNAP_MES).Value = mesTexto
    ws.Cells(inicio, SNAP_FECHA).Value = Now
    ws.Cells(inicio, SNAP_FECHA).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    inicio = inicio + 1
    primerDato = inicio
    For f = FILA_1 To FILA_FIN
        If Trim$(CStr(mes.Cells(f, COL_CODIGO).Value2)) <> "" Then
            ws.Cells(inicio, SNAP_ID).Value = identificador
            ws.Cells(inicio, SNAP_MES).Value = mesTexto
            ws.Cells(inicio, SNAP_FECHA).Value = Now
            ws.Cells(inicio, SNAP_CODIGO).Value = mes.Cells(f, COL_CODIGO).Value2
            ws.Cells(inicio, SNAP_DESCRIPCION).Value = mes.Cells(f, COL_DESCRIPCION).Value2
            ws.Cells(inicio, SNAP_UNIDAD).Value = mes.Cells(f, COL_UNIDAD).Value2
            ws.Cells(inicio, SNAP_CANT_FIN).Value = modCalculo.ValorNumerico(mes, f, COL_CANT_FIN)
            ws.Cells(inicio, SNAP_VAL_FIN).Value = modCalculo.ValorNumerico(mes, f, COL_VAL_FIN)
            ws.Cells(inicio, SNAP_COSTO_PROM).Value = modCalculo.ValorNumerico(mes, f, COL_COSTO_PROM)
            inicio = inicio + 1
        End If
    Next f

    ' El formato se aplica una sola vez al bloque recien escrito: por renglon
    ' eran tres asignaciones de NumberFormat y con 5.000 materiales pesaban.
    If inicio > primerDato Then
        ws.Range(ws.Cells(primerDato, SNAP_CANT_FIN), _
                 ws.Cells(inicio - 1, SNAP_CANT_FIN)).NumberFormat = MascaraCant()
        ws.Range(ws.Cells(primerDato, SNAP_VAL_FIN), _
                 ws.Cells(inicio - 1, SNAP_VAL_FIN)).NumberFormat = FMT_MONEDA
        ws.Range(ws.Cells(primerDato, SNAP_COSTO_PROM), _
                 ws.Cells(inicio - 1, SNAP_COSTO_PROM)).NumberFormat = FMT_COSTO
    End If

    ws.Cells(inicio, SNAP_ID).Value = identificador
    ws.Cells(inicio, SNAP_MES).Value = "fin de bloque " & mesTexto
    If protegido Then ws.Protect Password:=CLAVE_PROTECCION, UserInterfaceOnly:=True, _
        DrawingObjects:=True, Contents:=True, Scenarios:=False
End Sub

' Saldo de cierre de un material segun la instantanea (no la hoja viva).
Public Function SaldoInstantanea(ByVal mesTexto As String, ByVal codigo As String) As Double
    Dim ws As Worksheet
    Dim f As Long
    Dim clave As String
    Set ws = HojaInstantanea()
    clave = NormalizaClave(codigo)
    For f = 2 To UltimaFilaUsada(ws, SNAP_CODIGO)
        If CStr(ws.Cells(f, SNAP_MES).Value2) = mesTexto And _
           NormalizaClave(ws.Cells(f, SNAP_CODIGO).Value2) = clave Then
            SaldoInstantanea = CDbl(ws.Cells(f, SNAP_VAL_FIN).Value2)
            Exit Function
        End If
    Next f
End Function

Public Function SaldoInstantaneaCantidad(ByVal mesTexto As String, ByVal codigo As String) As Double
    Dim ws As Worksheet
    Dim f As Long
    Dim clave As String
    Set ws = HojaInstantanea()
    clave = NormalizaClave(codigo)
    For f = 2 To UltimaFilaUsada(ws, SNAP_CODIGO)
        If CStr(ws.Cells(f, SNAP_MES).Value2) = mesTexto And _
           NormalizaClave(ws.Cells(f, SNAP_CODIGO).Value2) = clave Then
            SaldoInstantaneaCantidad = CDbl(ws.Cells(f, SNAP_CANT_FIN).Value2)
            Exit Function
        End If
    Next f
End Function

' Conciliacion por codigo entre la hoja y la instantanea recien escrita.
' El indice de la instantanea se construye UNA sola vez: antes cada material
' recorria toda la hoja dos veces y con 5.000 renglones el cierre se disparaba.
Public Function CuadraInstantanea(ByVal indice As Long, ByVal identificador As String) As String
    Dim ws As Worksheet
    Dim mes As Worksheet
    Dim f As Long
    Dim ultima As Long
    Dim codigo As String
    Dim clave As String
    Dim texto As String
    Dim porId As Object
    Dim porMes As Object
    Dim filaSnap As Variant
    Dim mesTexto As String

    Set ws = HojaInstantanea()
    Set mes = Libro().Sheets(HojaMes(indice))
    mesTexto = HojaMes(indice)
    Set porId = CreateObject("Scripting.Dictionary")
    Set porMes = CreateObject("Scripting.Dictionary")

    ultima = UltimaFilaUsada(ws, SNAP_CODIGO)
    For f = 2 To ultima
        clave = NormalizaClave(CStr(ws.Cells(f, SNAP_CODIGO).Value2))
        If clave <> "" Then
            If CStr(ws.Cells(f, SNAP_ID).Value2) = identificador And _
               Not porId.Exists(clave) Then
                porId.Add clave, f
            End If
            If CStr(ws.Cells(f, SNAP_MES).Value2) = mesTexto And _
               Not porMes.Exists(clave) Then
                porMes.Add clave, f
            End If
        End If
    Next f

    For f = FILA_1 To FILA_FIN
        codigo = Trim$(CStr(mes.Cells(f, COL_CODIGO).Value2))
        If codigo <> "" Then
            clave = NormalizaClave(codigo)
            If Not porId.Exists(clave) Then
                texto = texto & "- " & codigo & ": no quedo en la instantanea." & vbCrLf
            Else
                If porMes.Exists(clave) Then
                    filaSnap = porMes(clave)
                Else
                    filaSnap = porId(clave)
                End If
                If Abs(Centavos(CDbl(ws.Cells(filaSnap, SNAP_VAL_FIN).Value2)) - _
                       Centavos(modCalculo.ValorNumerico(mes, f, COL_VAL_FIN))) > 0.005 Then
                    texto = texto & "- " & codigo & ": el saldo de la instantanea no coincide." & vbCrLf
                End If
            End If
        End If
    Next f
    CuadraInstantanea = texto
End Function

Private Function TieneInstantanea(ByVal ws As Worksheet, ByVal identificador As String, _
                                  ByVal codigo As String) As Boolean
    Dim f As Long
    For f = 2 To UltimaFilaUsada(ws, SNAP_CODIGO)
        If CStr(ws.Cells(f, SNAP_ID).Value2) = identificador And _
           NormalizaClave(ws.Cells(f, SNAP_CODIGO).Value2) = NormalizaClave(codigo) Then
            TieneInstantanea = True
            Exit Function
        End If
    Next f
End Function

' Reaplica los formatos numericos de _INSTANTANEAS: cantidad final con la
' mascara elegida en CONFIGURACION, saldo final a 2 decimales y costo final
' con al menos 6 decimales. Se llama con la hoja sin proteger.
Public Sub AplicarFormatos()
    Dim ws As Worksheet
    Dim ultima As Long
    Set ws = HojaInstantanea()
    ultima = ws.Cells(ws.Rows.Count, SNAP_ID).End(xlUp).Row
    If ultima < 2 Then Exit Sub
    ws.Range(ws.Cells(2, SNAP_CANT_FIN), ws.Cells(ultima, SNAP_CANT_FIN)) _
        .NumberFormat = MascaraCant()
    ws.Range(ws.Cells(2, SNAP_VAL_FIN), ws.Cells(ultima, SNAP_VAL_FIN)) _
        .NumberFormat = FMT_MONEDA
    ws.Range(ws.Cells(2, SNAP_COSTO_PROM), ws.Cells(ultima, SNAP_COSTO_PROM)) _
        .NumberFormat = FMT_COSTO
End Sub