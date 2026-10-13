Attribute VB_Name = "modCalculo"
Option Explicit

' Cálculo mensual: PROMEDIO PONDERADO PERIODICO MENSUAL.
'
'   QD = C + E        VD = D + F        QS = G + I        QF = QD - QS
'   CP = VD / QD   (columna O)
'   H  = REDONDEAR(G * CP ; 2)      J = REDONDEAR(I * CP ; 2)
'   K  = QD - QS                    L = VD - H - J
'   N  = D / C   si C > 0, si no 0
'   P  = L / K   si K > 0, si no 0
'
' Cuando QF = 0 la salida total vale exactamente VD: si hay salidas "otros",
' J = VD - H; si no las hay, H = VD y J = 0. El residuo de centavos se
' asigna a la ultima categoria no nula del orden ventas/otros.
' Cuando QF < 0 se calcula igual (saldo negativo visible), Q marca ERROR y
' el cierre queda bloqueado. No se usa IFERROR para ocultar defectos.

Public Sub AplicarFormulas(ByVal hoja As Worksheet, ByVal primera As Long, ByVal ultima As Long)
    Dim f As Long
    If ultima < primera Then Exit Sub
    For f = primera To ultima
        EscribirFormulasFila hoja, f
    Next f
End Sub

' Relleno por columna: una sola llamada por columna para los 5.000 renglones.
' Las referencias de fila son relativas, así que Excel las ajusta por renglon.
Public Sub RellenarFormulas(ByVal hoja As Worksheet)
    hoja.Range("H" & FILA_1 & ":H" & FILA_FIN).Formula = _
        "=IF(($C7+$E7-$G7-$I7)=0,IF($I7>0,ROUND($G7*$O7,2),$D7+$F7),ROUND($G7*$O7,2))"
    hoja.Range("J" & FILA_1 & ":J" & FILA_FIN).Formula = _
        "=IF(($C7+$E7-$G7-$I7)=0,IF($I7>0,$D7+$F7-ROUND($G7*$O7,2),0),ROUND($I7*$O7,2))"
    hoja.Range("K" & FILA_1 & ":K" & FILA_FIN).Formula = "=($C7+$E7)-$G7-$I7"
    hoja.Range("L" & FILA_1 & ":L" & FILA_FIN).Formula = "=($D7+$F7)-H7-J7"
    hoja.Range("N" & FILA_1 & ":N" & FILA_FIN).Formula = "=IF($C7>0,$D7/$C7,0)"
    hoja.Range("O" & FILA_1 & ":O" & FILA_FIN).Formula = "=IF(($C7+$E7)>0,($D7+$F7)/($C7+$E7),0)"
    hoja.Range("P" & FILA_1 & ":P" & FILA_FIN).Formula = "=IF(K7>0,L7/K7,0)"
    hoja.Range("Q" & FILA_1 & ":Q" & FILA_FIN).Formula = FormulaEstado(FILA_1)
    FormatearCostos hoja, FILA_1, FILA_FIN
End Sub

Public Sub EscribirFormulasFila(ByVal hoja As Worksheet, ByVal f As Long)
    Dim Q0 As String, V0 As String, QE As String, VE As String, VD As String
    Dim QV As String, QO As String, CP As String, QD As String, QF As String
    Q0 = "$C" & f
    V0 = "$D" & f
    QE = "$E" & f
    VE = "$F" & f
    QV = "$G" & f
    QO = "$I" & f
    CP = "$O" & f
    QD = "(" & Q0 & "+" & QE & ")"
    VD = "(" & V0 & "+" & VE & ")"
    QF = "(" & QD & "-" & QV & "-" & QO & ")"

    hoja.Cells(f, COL_VAL_VTA).Formula = _
        "=IF(" & QF & "=0,IF(" & QO & ">0,ROUND(" & QV & "*" & CP & ",2)," & VD & ")," & _
        "ROUND(" & QV & "*" & CP & ",2))"
    hoja.Cells(f, COL_VAL_OTR).Formula = _
        "=IF(" & QF & "=0,IF(" & QO & ">0," & VD & "-ROUND(" & QV & "*" & CP & ",2),0)," & _
        "ROUND(" & QO & "*" & CP & ",2))"
    hoja.Cells(f, COL_CANT_FIN).Formula = "=" & QD & "-" & QV & "-" & QO
    hoja.Cells(f, COL_VAL_FIN).Formula = "=" & VD & "-H" & f & "-J" & f
    hoja.Cells(f, COL_COSTO_INI).Formula = "=IF(" & Q0 & ">0," & V0 & "/" & Q0 & ",0)"
    hoja.Cells(f, COL_COSTO_PROM).Formula = "=IF(" & QD & ">0," & VD & "/" & QD & ",0)"
    hoja.Cells(f, COL_COSTO_FIN).Formula = "=IF(K" & f & ">0,L" & f & "/K" & f & ",0)"
    hoja.Cells(f, COL_ESTADO).Formula = FormulaEstado(f)
End Sub

Private Function FormulaEstado(ByVal f As Long) As String
    ' Q: estado de validación visible en la hoja.
    Dim A As String, C As String, D As String, E As String, G As String
    Dim I As String, K As String, L As String, O As String
    A = "$A" & f
    C = "$C" & f
    D = "$D" & f
    E = "$E" & f
    G = "$G" & f
    I = "$I" & f
    K = "$K" & f
    L = "$L" & f
    O = "$O" & f
    FormulaEstado = _
        "=IF(" & A & "="""",""""," & _
        "IF(AND(" & C & "=0," & D & "<>0),""ERROR: cantidad inicial 0 con valor distinto de 0""," & _
        "IF(AND(" & E & "=0,$F" & f & "<>0),""ERROR: entradas 0 con valor distinto de 0""," & _
        "IF(AND((" & C & "+" & E & ")=0,OR($D" & f & "+$F" & f & "<>0," & G & "+" & I & "<>0))," & _
        """ERROR: no se puede calcular el costo promedio""," & _
        "IF(" & K & "<0,""ERROR: salidas mayores a lo disponible""," & _
        "IF(" & L & "<0,""ERROR DE REDONDEO: revise importes de salida""," & _
        "IF(AND(OR(" & G & ">0," & I & ">0)," & O & "=0)," & _
        """REVISAR: cantidad positiva con costo promedio cero""," & _
        "IF(AND(" & L & "=0," & K & "=0),""OK: sin saldo"",""OK""))))))))"
End Function

' ---------------------------------------------------------------------------
' Totalizadores: los monetary son independientes de filtros; las cantidades
' solo se totalizan por unidad de medida, porque no es una magnitud valida
' sumar kilos, litros y unidades.
' ---------------------------------------------------------------------------
Public Sub EscribirTotales(ByVal hoja As Worksheet)
    Dim c As Variant
    Dim direccion As String
    hoja.Cells(FILA_TOTAL, 1).Value = "TOTAL Bs. (independiente de filtros)"
    hoja.Cells(FILA_TOTAL_FILT, 1).Value = "TOTAL Bs. solo filas visibles (SUBTOTAL 109)"
    For Each c In Array(COL_VAL_INI, COL_VAL_ENT, COL_VAL_VTA, COL_VAL_OTR, COL_VAL_FIN)
        direccion = hoja.Range(hoja.Cells(FILA_1, c), hoja.Cells(FILA_FIN, c)).Address
        hoja.Cells(FILA_TOTAL, c).Formula = "=SUM(" & direccion & ")"
        hoja.Cells(FILA_TOTAL_FILT, c).Formula = "=SUBTOTAL(109," & direccion & ")"
    Next c
    hoja.Cells(FILA_TOTAL, COL_UNIDAD).Value = "cantidades por unidad"
    hoja.Cells(FILA_TOTAL_FILT, COL_UNIDAD).Value = "cantidades por unidad"
    hoja.Cells(FILA_TOTAL, COL_COSTO_PROM).Value = "no se totaliza"
    hoja.Cells(FILA_TOTAL_FILT, COL_COSTO_PROM).Value = "no se totaliza"
End Sub

' ---------------------------------------------------------------------------
' Verificacion aritmetica independiente de los totales de la hoja
' ---------------------------------------------------------------------------
Public Function CuadraTotales(ByVal hoja As Worksheet) As Boolean
    Dim columnas As Variant
    Dim c As Variant
    Dim suma As Double
    Dim referencia As Double
    columnas = Array(COL_VAL_INI, COL_VAL_ENT, COL_VAL_VTA, COL_VAL_OTR, COL_VAL_FIN)
    CuadraTotales = True
    For Each c In columnas
        suma = 0
        referencia = 0
        On Error Resume Next
        suma = CDbl(hoja.Cells(FILA_TOTAL, c).Value2)
        On Error GoTo 0
        referencia = SumaColumna(hoja, c)
        If Abs(Centavos(suma) - Centavos(referencia)) > 0.01 Then CuadraTotales = False
    Next c
End Function

Public Function SumaColumna(ByVal hoja As Worksheet, ByVal columna As Long) As Double
    Dim f As Long
    Dim total As Double
    For f = FILA_1 To FILA_FIN
        If Not EsVacio(hoja.Cells(f, columna).Value2) Then
            If EsNumeroReal(hoja.Cells(f, columna).Value2) Then
                total = total + CDbl(hoja.Cells(f, columna).Value2)
            End If
        End If
    Next f
    SumaColumna = total
End Function

' D + F - H - J = L fila a fila: la identidad no depende de filtros.
Public Function CuadraFila(ByVal hoja As Worksheet, ByVal f As Long) As Boolean
    Dim ini As Double, ent As Double, vta As Double, otr As Double, fin As Double
    ini = ValorNumerico(hoja, f, COL_VAL_INI)
    ent = ValorNumerico(hoja, f, COL_VAL_ENT)
    vta = ValorNumerico(hoja, f, COL_VAL_VTA)
    otr = ValorNumerico(hoja, f, COL_VAL_OTR)
    fin = ValorNumerico(hoja, f, COL_VAL_FIN)
    CuadraFila = (Abs(Centavos(ini + ent - vta - otr) - Centavos(fin)) <= 0.005)
End Function

Public Function ValorNumerico(ByVal hoja As Worksheet, ByVal f As Long, ByVal columna As Long) As Double
    Dim v As Variant
    v = hoja.Cells(f, columna).Value2
    If EsNumeroReal(v) Then
        ValorNumerico = CDbl(v)
    Else
        ValorNumerico = 0
    End If
End Function

' Regla contable de presentacion:
'   - Importes en Bs. (D, F, H, J, L): SIEMPRE 2 decimales (FMT_MONEDA).
'   - Cantidades (C, E, G, I, K): la mascara elegida en CONFIGURACION.
'   - Costos unitarios (N, O, P): al menos 6 decimales visibles.
Public Sub FormatearCostos(ByVal hoja As Worksheet, ByVal primera As Long, ByVal ultima As Long)
    Dim c As Variant
    If ultima < primera Then Exit Sub
    ' Cantidades: C, E, G, I, K
    FormatearCantidades hoja, primera, ultima, MascaraCant()
    ' Importes en Bs.: D, F, H, J, L
    For Each c In Array(COL_VAL_INI, COL_VAL_ENT, COL_VAL_VTA, COL_VAL_OTR, COL_VAL_FIN)
        hoja.Range(hoja.Cells(primera, c), hoja.Cells(ultima, c)).NumberFormat = FMT_MONEDA
    Next c
    ' Costos unitarios: N, O, P
    hoja.Range(hoja.Cells(primera, COL_COSTO_INI), hoja.Cells(ultima, COL_COSTO_FIN)).NumberFormat = FMT_COSTO
End Sub

Private Sub FormatearCantidades(ByVal hoja As Worksheet, ByVal primera As Long, _
                                ByVal ultima As Long, ByVal mascara As String)
    Dim c As Variant
    If ultima < primera Then Exit Sub
    For Each c In Array(COL_CANT_INI, COL_CANT_ENT, COL_CANT_VTA, COL_CANT_OTR, COL_CANT_FIN)
        hoja.Range(hoja.Cells(primera, c), hoja.Cells(ultima, c)).NumberFormat = mascara
    Next c
End Sub

' Reaplica la mascara de cantidades elegida en CONFIGURACION a los 12 meses.
' Se llama con las hojas sin proteger: Workbook_Open y Workbook_BeforeSave
' desprotegen antes y modEstados.ProtegerTodoParaGuardar vuelve a proteger.
Public Sub AplicarMascaraCantidades()
    Dim ws As Worksheet
    Dim mascara As String
    mascara = MascaraCant()
    For Each ws In Libro().Worksheets
        If EsHojaMes(ws.Name) Then
            FormatearCantidades ws, FILA_1, FILA_TOTAL_FILT, mascara
        End If
    Next ws
End Sub

' Cantidades con mas precision que 2 decimales se muestran completas.
Public Sub AjustarDecimalesVisibles(ByVal hoja As Worksheet, ByVal f As Long)
    Dim mascara As String
    Dim c As Variant
    mascara = MascaraCant()
    ' Con la mascara sin decimales el usuario pidio no ver fracciones: no se
    ' contradice su eleccion. El valor guardado conserva la precision completa.
    If mascara = MASCARA_CANT_SIN_DEC Then Exit Sub
    For Each c In Array(COL_CANT_INI, COL_CANT_ENT, COL_CANT_VTA, COL_CANT_OTR, COL_CANT_FIN)
        With hoja.Cells(f, c)
            If EsNumeroReal(.Value2) Then
                If Abs(CDbl(.Value2) - Redondear2(CDbl(.Value2))) > 0.00005 Then
                    .NumberFormat = mascara
                End If
            End If
        End With
    Next c
End Sub