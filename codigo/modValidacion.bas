Attribute VB_Name = "modValidacion"
Option Explicit

' Validación al ingresar, al pegar/importar, al abrir y antes del cierre.
' El cierre falla con cualquier error pendiente: no basta el color de la celda.

Public Function ValidarMes(ByVal indice As Long) As String
    Dim problemas As String
    Dim ws As Worksheet
    Dim f As Long
    Dim codigo As String
    Dim Q0 As Double, V0 As Double, QE As Double, VE As Double
    Dim QV As Double, QO As Double, QD As Double, VD As Double, QS As Double
    Dim QF As Double, K As Double, L As Double, CP As Double

    problemas = modConfiguracion.ValidarConfiguracion()
    Set ws = Libro().Sheets(HojaMes(indice))
    Application.Calculate

    For f = FILA_1 To FILA_FIN
        codigo = Trim$(CStr(ws.Cells(f, COL_CODIGO).Value2))
        If codigo <> "" Then
            If FilaEnCatalogo(codigo) = 0 Then
                problemas = problemas & "- Fila " & f & ": codigo " & codigo & _
                    " no existe en CATALOGO." & vbCrLf
            End If
            If EsVacio(ws.Cells(f, COL_UNIDAD).Value2) Then
                problemas = problemas & "- Fila " & f & ": unidad de medida vacia." & vbCrLf
            End If
            If EsVacio(ws.Cells(f, COL_CANT_INI).Value2) Or _
               EsVacio(ws.Cells(f, COL_VAL_INI).Value2) Then
                If indice > 1 Then
                    problemas = problemas & "- Fila " & f & ": saldo inicial vacio (C/D). " & _
                        "Cero es un saldo valido; vacio no." & vbCrLf
                End If
            End If
            Q0 = modCalculo.ValorNumerico(ws, f, COL_CANT_INI)
            V0 = modCalculo.ValorNumerico(ws, f, COL_VAL_INI)
            QE = modCalculo.ValorNumerico(ws, f, COL_CANT_ENT)
            VE = modCalculo.ValorNumerico(ws, f, COL_VAL_ENT)
            QV = modCalculo.ValorNumerico(ws, f, COL_CANT_VTA)
            QO = modCalculo.ValorNumerico(ws, f, COL_CANT_OTR)
            QD = Q0 + QE
            VD = V0 + VE
            QS = QV + QO
            QF = QD - QS
            K = modCalculo.ValorNumerico(ws, f, COL_CANT_FIN)
            L = modCalculo.ValorNumerico(ws, f, COL_VAL_FIN)
            CP = modCalculo.ValorNumerico(ws, f, COL_COSTO_PROM)

            If Q0 = 0 And V0 <> 0 Then
                problemas = problemas & "- Fila " & f & ": C=0 con D distinto de 0." & vbCrLf
            End If
            If QE = 0 And VE <> 0 Then
                problemas = problemas & "- Fila " & f & ": E=0 con F distinto de 0." & vbCrLf
            End If
            If QD = 0 And (VD <> 0 Or QS <> 0) Then
                problemas = problemas & "- Fila " & f & ": " & _
                    "QD=0 con VD o QS distintos de cero; no se puede valorar." & vbCrLf
            End If
            If K < -0.00005 Then
                problemas = problemas & "- Fila " & f & ": salidas mayores a lo disponible " & _
                    "(K=" & Format$(K, FMT_CANT) & "); no se reduce la salida ni se pone stock cero." & vbCrLf
            End If
            If L < -0.005 Then
                problemas = problemas & "- Fila " & f & ": ERROR DE REDONDEO, valor final negativo " & _
                    "(" & Format$(L, FMT_MONEDA) & ")." & vbCrLf
            End If
            If Not modCalculo.CuadraFila(ws, f) Then
                problemas = problemas & "- Fila " & f & ": D+F-H-J no es igual a L." & vbCrLf
            End If
            If (QV > 0 Or QO > 0) And CP = 0 And Not modConfiguracion.AceptacionCostoCero() Then
                problemas = problemas & "- Fila " & f & ": cantidad positiva con costo promedio cero; " & _
                    "requiere aceptacion documentada del responsable." & vbCrLf
            End If
            If HayErrorDeFormula(ws, f) Then
                problemas = problemas & "- Fila " & f & ": formula alterada o ausente en celdas calculadas." & vbCrLf
            End If
            If QF > 0 Then
                If Abs(Centavos(modCalculo.ValorNumerico(ws, f, COL_VAL_VTA)) - _
                       Centavos(Redondear2(QV * CP))) > 0.005 Then
                    problemas = problemas & "- Fila " & f & ": H no coincide con REDONDEAR(G*O;2)." & vbCrLf
                End If
                If Abs(Centavos(modCalculo.ValorNumerico(ws, f, COL_VAL_OTR)) - _
                       Centavos(Redondear2(QO * CP))) > 0.005 Then
                    problemas = problemas & "- Fila " & f & ": J no coincide con REDONDEAR(I*O;2)." & vbCrLf
                End If
            ElseIf Abs(QF) <= 0.00005 And QV + QO > 0 Then
                If Abs(Centavos(VD) - Centavos(modCalculo.ValorNumerico(ws, f, COL_VAL_VTA)) - _
                       Centavos(modCalculo.ValorNumerico(ws, f, COL_VAL_OTR))) > 0.005 Then
                    problemas = problemas & "- Fila " & f & ": con agotamiento, H+J debe ser exactamente VD." & vbCrLf
                End If
            End If
        End If
    Next f

    problemas = problemas & ValidarEntradaNumerica(ws)
    problemas = problemas & modDetalle.ValidarDetalleDeMes(indice)
    problemas = problemas & ValidarTotales(ws)
    ValidarMes = problemas
End Function

' Las celdas de entrada no admiten texto, negativos, errores ni formulas pegadas.
Private Function ValidarEntradaNumerica(ByVal ws As Worksheet) As String
    Dim f As Long
    Dim c As Variant
    Dim texto As String
    Dim celda As Range
    For Each c In Array(COL_CANT_ENT, COL_VAL_ENT, COL_CANT_VTA, COL_CANT_OTR)
        For f = FILA_1 To FILA_FIN
            Set celda = ws.Cells(f, c)
            If Not IsEmpty(celda.Value2) Then
                If Not EsNumeroReal(celda.Value2) Then
                    texto = texto & "- Fila " & f & ", columna " & modCatalogo.ColumnaLetra(c) & _
                        ": la entrada no es un numero." & vbCrLf
                ElseIf CDbl(celda.Value2) < 0 Then
                    texto = texto & "- Fila " & f & ", columna " & modCatalogo.ColumnaLetra(c) & _
                        ": entrada negativa." & vbCrLf
                End If
            End If
        Next f
    Next c
    ValidarEntradaNumerica = texto
End Function

' Control de fórmulas del mes abierto: las columnas calculadas deben
' contener la fórmula esperada. En meses cerrados se valida la instantánea.
Public Function HayErrorDeFormula(ByVal ws As Worksheet, ByVal f As Long) As Boolean
    Dim c As Variant
    For Each c In Array(COL_VAL_VTA, COL_VAL_OTR, COL_CANT_FIN, COL_VAL_FIN, _
                       COL_COSTO_INI, COL_COSTO_PROM, COL_COSTO_FIN, COL_ESTADO)
        If Not ws.Cells(f, c).HasFormula Then
            HayErrorDeFormula = True
            Exit Function
        End If
        If IsError(ws.Cells(f, c).Value2) Then
            HayErrorDeFormula = True
            Exit Function
        End If
    Next c
End Function

Public Function ValidarTotales(ByVal ws As Worksheet) As String
    Dim texto As String
    If Not modCalculo.CuadraTotales(ws) Then
        texto = texto & "- El totalizador de la fila " & FILA_TOTAL & _
            " no coincide con la suma de las filas." & vbCrLf
    End If
    ValidarTotales = texto
End Function

' Reporte de errores del mes abierto, para la columna Q y para la pantalla.
Public Sub ReportarErrores(ByVal indice As Long)
    Dim problemas As String
    problemas = ValidarMes(indice)
    If problemas = "" Then
        Aviso "Mes " & HojaMes(indice) & ": sin errores bloqueantes.", vbInformation, APP_NOMBRE
    Else
        Aviso "Mes " & HojaMes(indice) & ": hay hallazgos." & vbCrLf & vbCrLf & problemas, _
            vbExclamation, APP_NOMBRE
    End If
End Sub

' Boton Validar mes
Public Sub ValidarMesActual()
    Dim indice As Long
    indice = MesAbierto()
    If indice = 0 Then
        Aviso "No hay ningun mes abierto.", vbExclamation, APP_NOMBRE
        Exit Sub
    End If
    ReportarErrores indice
End Sub