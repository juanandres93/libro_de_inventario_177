Attribute VB_Name = "modGlobales"
Option Explicit

' ============================================================================
' Libro de inventarios - constantes, direcciones y utilidades comunes.
' AdaptaPro - metodo PROMEDIO PONDERADO PERIODICO MENSUAL.
' No es promedio movil: la valoracion es mensual, con saldo inicial y total
' de entradas del mes (columna O).
' ============================================================================

Public Const APP_NOMBRE As String = "AdaptaPro - Libro de inventarios"
Public Const APP_URL As String = "https://adaptaproerp.com/blog_openconde/"
Public Const APP_VERSION As String = "1.1.0"
Public Const APP_METODO As String = "PROMEDIO PONDERADO PERIODICO MENSUAL"
Public Const APP_MONEDA As String = "Bs."
Public Const CLAVE_PROTECCION As String = "AdaptaPro2026"

' Aviso del articulo 177 del Reglamento (D. 2.507, G.O.E. 5.662, 24-09-2003)
' y deberes de conservacion de los articulos 88 y 89 del LISLR y del
' articulo 64 del Reglamento para las destrucciones.
' El texto completo mide mas de 255 caracteres, por lo que se divide en
' constantes (cada fragmento por debajo del limite de VBA) y el texto completo
' se arma en la funcion AvisoArt177(), usada en el libro, la ayuda y la bitacora.
Public Const AVISO_ART177_A As String = _
    "Art. 177 del Reglamento (D. 2.507, G.O.E. 5.662 del 24-09-2003): el " & _
    "registro por medios magneticos requiere autorizacion de la Administracion " & _
    "Tributaria. "
Public Const AVISO_ART177_B As String = _
    "Este libro no la otorga ni la sustituye: el responsable debe confirmar " & _
    "esa autorizacion antes de usarlo como registro fiscal. "
Public Const AVISO_ART177_C As String = _
    "Arts. 88 y 89 del LISLR: los comprobantes y demas documentos que amparan " & _
    "las operaciones se conservan durante el ejercicio fiscal en curso y los "
Public Const AVISO_ART177_D As String = _
    "tres siguientes; el libro de inventarios debe conservarse en las mismas " & _
    "condiciones. "
Public Const AVISO_ART177_E As String = _
    "Art. 64 (par. cuarto) del Reglamento para las destrucciones: la " & _
    "destruccion de existencias exige acta o autorizacion escrita, que debe " & _
    "registrarse en el detalle de otros movimientos de este libro."

' ---------------------------------------------------------------------------
' Estructura de filas y columnas
' ---------------------------------------------------------------------------
Public Const FILA_ENCABEZADO As Long = 6      ' encabezados de columna
Public Const FILA_1 As Long = 7               ' primer material
Public Const FILA_FIN As Long = 5006          ' limite tecnico: 5.000 materiales
Public Const FILA_TOTAL As Long = 5007        ' totales (independientes de filtro)
Public Const FILA_TOTAL_FILT As Long = 5008   ' total filtrado (SUBTOTAL, diferenciado)
Public Const LIMITE_MATERIALES As Long = 5000
Public Const FILA_DET_1 As Long = 7           ' primer renglon de DETALLE_OTROS
Public Const FILA_DET_FIN As Long = 50006     ' hasta 50.000 lineas de detalle

Public Const COL_CODIGO As Long = 1           ' A
Public Const COL_DESCRIPCION As Long = 2      ' B
Public Const COL_CANT_INI As Long = 3         ' C
Public Const COL_VAL_INI As Long = 4           ' D
Public Const COL_CANT_ENT As Long = 5         ' E
Public Const COL_VAL_ENT As Long = 6           ' F
Public Const COL_CANT_VTA As Long = 7         ' G
Public Const COL_VAL_VTA As Long = 8           ' H
Public Const COL_CANT_OTR As Long = 9         ' I
Public Const COL_VAL_OTR As Long = 10          ' J
Public Const COL_CANT_FIN As Long = 11        ' K
Public Const COL_VAL_FIN As Long = 12          ' L
Public Const COL_UNIDAD As Long = 13           ' M
Public Const COL_COSTO_INI As Long = 14        ' N
Public Const COL_COSTO_PROM As Long = 15       ' O
Public Const COL_COSTO_FIN As Long = 16        ' P
Public Const COL_ESTADO As Long = 17          ' Q

' Configuracion
Public Const HOJA_CONFIG As String = "CONFIGURACION"
Public Const HOJA_CATALOGO As String = "CATALOGO"
Public Const HOJA_RESUMEN As String = "RESUMEN_ANUAL"
Public Const HOJA_DETALLE As String = "DETALLE_OTROS"
Public Const HOJA_BITACORA As String = "BITACORA"
Public Const HOJA_AYUDA As String = "AYUDA"
Public Const HOJA_ESTADOS As String = "_ESTADOS"
Public Const HOJA_SNAP As String = "_INSTANTANEAS"

' Celdas de CONFIGURACION
Public Const CFG_ENTIDAD As String = "B3"
Public Const CFG_RIF As String = "B4"
Public Const CFG_DOMICILIO As String = "B5"
Public Const CFG_RESPONSABLE As String = "B6"
Public Const CFG_ANIO As String = "B7"
Public Const CFG_MONEDA As String = "B8"
Public Const CFG_VERSION As String = "B9"
Public Const CFG_METODO As String = "B10"
Public Const CFG_ESTADO As String = "B11"
Public Const CFG_LIMITE As String = "B12"
Public Const CFG_AUTORIZACION As String = "B13"
Public Const CFG_COSTO_CERO As String = "B14"
Public Const CFG_CLAVE As String = "B15"

' Estados de mes
Public Const EST_NO_ABIERTO As String = "NO ABIERTO"
Public Const EST_APERTURA_PEND As String = "APERTURA PENDIENTE"
Public Const EST_ABIERTO As String = "ABIERTO"
Public Const EST_CERRADO As String = "CERRADO"
Public Const EST_INCIDENCIA As String = "INCIDENCIA"

' Tipos de salida "otros" (detalle)
Public Const TIPOS_OTROS As String = "retiro|autoconsumo|merma|destruccion|otro"

' Formatos
Public Const FMT_CANT As String = "#,##0.00####"
Public Const FMT_MONEDA As String = "#,##0.00"
Public Const FMT_COSTO As String = "#,##0.000000"
Public Const FMT_ENTERO As String = "0"

' Mascara numerica de cantidades: la elige el usuario en CONFIGURACION (B16)
' y solo cambia la presentacion; el valor guardado y todos los calculos usan
' la precision completa (cantidades hasta 4 decimales).
Public Const CFG_MASCARA_CANT As String = "B16"
Public Const MASCARA_CANT_SIN_DEC As String = "#,##0"

' ---------------------------------------------------------------------------
' Estado de la sesion
' ---------------------------------------------------------------------------
Private mSesionIniciada As Boolean
Private mReentrada As Boolean

' Indice del catalogo (debe quedar en el bloque de declaraciones del modulo,
' antes del primer procedimiento).
Private mIndiceCatalogo As Object
Private mIndiceSello As String

Public Property Get SesionIniciada() As Boolean
    SesionIniciada = mSesionIniciada
End Property

Public Sub MarcarSesion(ByVal valor As Boolean)
    mSesionIniciada = valor
End Sub

Public Property Get Reentrada() As Boolean
    Reentrada = mReentrada
End Property

Public Sub Entrar()
    mReentrada = True
End Sub

Public Sub Salir()
    mReentrada = False
End Sub

' ---------------------------------------------------------------------------
' Utilidades de texto y de redondeo
' ---------------------------------------------------------------------------
Public Function LimpiarTexto(ByVal valor As Variant) As String
    Dim s As String
    If IsEmpty(valor) Or IsNull(valor) Then
        LimpiarTexto = ""
        Exit Function
    End If
    s = CStr(valor)
    s = Replace(s, Chr(160), " ")          ' espacio no separable
    s = Replace(s, ChrW(8203), "")          ' zero width (Chr solo admite 0-255)
    s = Trim$(s)                            ' externos
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    LimpiarTexto = s
End Function

Public Function NormalizaClave(ByVal valor As Variant) As String
    ' Clave de comparacion: mayusculas, sin espacios superfluous ni saltos.
    Dim s As String
    s = LimpiarTexto(valor)
    s = UCase$(s)
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    Do While InStr(s, " ") > 0
        s = Replace(s, " ", "")
    Loop
    NormalizaClave = s
End Function

Public Function EsVacio(ByVal valor As Variant) As Boolean
    If IsEmpty(valor) Then
        EsVacio = True
    ElseIf IsNull(valor) Then
        EsVacio = True
    ElseIf IsError(valor) Then
        EsVacio = False
    Else
        EsVacio = (Len(Trim$(CStr(valor))) = 0)
    End If
End Function

Public Function EsNumeroReal(ByVal valor As Variant) As Boolean
    If IsEmpty(valor) Or IsNull(valor) Or IsError(valor) Then
        EsNumeroReal = False
    Else
        EsNumeroReal = IsNumeric(valor) And Not IsDate(valor)
    End If
End Function

' Excel guarda las fechas como serial numerico: Value2 devuelve Double y
' IsDate(Double) devuelve False en VBA, por lo que una fecha de celda nunca
' pasaria el control. Se acepta el serial, el texto de fecha y la fecha nativa.
Public Function EsFecha(ByVal valor As Variant) As Boolean
    If IsEmpty(valor) Or IsNull(valor) Or IsError(valor) Then
        EsFecha = False
    ElseIf IsDate(valor) Then
        EsFecha = True
    ElseIf EsNumeroReal(valor) Then
        EsFecha = (CDbl(valor) >= 1 And CDbl(valor) <= 2958465)
    End If
End Function

' Redondeo de hoja (WorksheetFunction.Round). NO se usa VBA.Round porque
' aplica redondeo bancario; WorksheetFunction.Round coincide con ROUND().
Public Function Redondear2(ByVal valor As Double) As Double
    On Error GoTo sinRedondeo
    Redondear2 = Application.WorksheetFunction.Round(valor, 2)
    Exit Function
sinRedondeo:
    Redondear2 = valor
End Function

Public Function Redondear4(ByVal valor As Double) As Double
    On Error GoTo sinRedondeo
    Redondear4 = Application.WorksheetFunction.Round(valor, 4)
    Exit Function
sinRedondeo:
    Redondear4 = valor
End Function

' Centavos enteros: comparacion de dinero sin errores de igualdad binaria.
Public Function Centavos(ByVal valor As Double) As Double
    Centavos = Redondear2(valor * 100#) * 100#
End Function

' Cuatro decimales cuantizados, para comparar cantidades.
Public Function Cuantiza4(ByVal valor As Double) As Double
    Cuantiza4 = Redondear4(valor)
End Function

' Mascara numerica de cantidades: la elige el usuario en CONFIGURACION (B16) y
' solo cambia la presentacion; el valor guardado y todos los calculos usan la
' precision completa (cantidades hasta 4 decimales). Devuelve la mascara
' vigente; ante una celda vacia o con un valor que no sea una de las dos
' opciones permitidas, devuelve FMT_CANT.
Public Function MascaraCant() As String
    Dim texto As String
    MascaraCant = FMT_CANT
    On Error Resume Next
    texto = Trim$(CStr(Libro().Sheets(HOJA_CONFIG).Range(CFG_MASCARA_CANT).Value2))
    On Error GoTo 0
    If texto = MASCARA_CANT_SIN_DEC Then MascaraCant = MASCARA_CANT_SIN_DEC
End Function

' Texto completo del aviso normativo (art. 177 del Reglamento, art. 88/89 del
' LISLR y art. 64 del Reglamento para las destrucciones), armado a partir de
' las constantes AVISO_ART177_A..E porque VBA no admite una constante de mas
' de 255 caracteres.
Public Function AvisoArt177() As String
    AvisoArt177 = AVISO_ART177_A & AVISO_ART177_B & AVISO_ART177_C & _
                  AVISO_ART177_D & AVISO_ART177_E
End Function

' ---------------------------------------------------------------------------
' Fechas
' ---------------------------------------------------------------------------
Public Function PeriodoActual() As String
    PeriodoActual = Format$(Now, "yyyy-mm")
End Function

Public Function MarcaTiempo() As String
    MarcaTiempo = CStr(Year(Now)) & "-" & Dos(Month(Now)) & "-" & Dos(Day(Now)) & " " & _
                  Dos(Hour(Now)) & ":" & Dos(Minute(Now)) & ":" & Dos(Second(Now))
End Function

' Sello digitos para IDs y nombres de archivo (sin ambiguedad de formatos).
Public Function SelloTiempo() As String
    SelloTiempo = CStr(Year(Now)) & Dos(Month(Now)) & Dos(Day(Now)) & _
                  Dos(Hour(Now)) & Dos(Minute(Now)) & Dos(Second(Now))
End Function

Private Function Dos(ByVal n As Long) As String
    Dos = Format$(n, "00")
End Function

' ---------------------------------------------------------------------------
' Huella SHA-256 para respaldo y trazabilidad (arts. 88 y 89 LISLR).
' Usa .NET del propio Windows: sin descargas ni dependencias externas.
' ---------------------------------------------------------------------------
Public Function SHA256DeArchivo(ByVal ruta As String) As String
    On Error GoTo fallo
    Dim fso As Object
    Dim stream As Object
    Dim datos() As Byte
    Dim hash() As Byte
    Dim algoritmo As Object
    Dim i As Long
    Dim texto As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(ruta) Then
        SHA256DeArchivo = "no disponible"
        Exit Function
    End If
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1                        ' adTypeBinary
    stream.Open
    stream.LoadFromFile ruta
    datos = stream.Read
    stream.Close
    Set algoritmo = CreateObject("System.Security.Cryptography.SHA256Managed")
    hash = algoritmo.ComputeHash(datos)
    For i = LBound(hash) To UBound(hash)
        texto = texto & Right$("0" & Hex$(hash(i)), 2)
    Next i
    SHA256DeArchivo = LCase$(texto)
    Exit Function
fallo:
    SHA256DeArchivo = "no disponible"
End Function

' ---------------------------------------------------------------------------
' Nombres de hoja y acceso al libro (referencias calificadas: nada de
' ActiveSheet, ActiveWorkbook ni Select).
' ---------------------------------------------------------------------------
Public Function Libro() As Workbook
    Set Libro = ThisWorkbook
End Function

Public Function NombresMes() As Variant
    NombresMes = Array("ENE", "FEB", "MAR", "ABR", "MAY", "JUN", _
                       "JUL", "AGO", "SEP", "OCT", "NOV", "DIC")
End Function

Public Function HojaMes(ByVal indice As Long) As String
    Dim nombres As Variant
    Dim anio As String
    If indice < 1 Or indice > 12 Then
        Err.Raise vbObjectError + 513, "HojaMes", "mes fuera de rango"
    End If
    anio = Trim$(CStr(Libro().Sheets(HOJA_CONFIG).Range(CFG_ANIO).Value2))
    If anio = "" Then
        Err.Raise vbObjectError + 514, "HojaMes", _
            "libro sin configurar: capture la configuracion (anio) antes de operar"
    End If
    nombres = NombresMes()
    HojaMes = anio & "-" & nombres(indice - 1)
End Function

' El anio se captura al crear el archivo y se usa para nombrar las hojas
' AAAA-ENE ... AAAA-DIC. Sin el anio no hay nombre de hoja valido, por lo que
' las funciones de estado lo avisan con contexto en vez de fallar con un
' error de indice.
Public Function AnioConfigurado() As Boolean
    AnioConfigurado = Len(Trim$(CStr(Libro().Sheets(HOJA_CONFIG).Range(CFG_ANIO).Value2))) > 0
End Function

Public Function IndiceDeHoja(ByVal nombre As String) As Long
    Dim nombres As Variant, i As Long
    nombres = NombresMes()
    For i = 1 To 12
        If UCase$(Right$(nombre, 3)) = CStr(nombres(i - 1)) Then
            IndiceDeHoja = i
            Exit Function
        End If
    Next i
    IndiceDeHoja = 0
End Function

Public Function EsHojaMes(ByVal nombre As String) As Boolean
    EsHojaMes = (IndiceDeHoja(nombre) > 0)
End Function

Public Function AnioDelLibro() As Long
    On Error GoTo sinAnio
    AnioDelLibro = CLng(Libro().Sheets(HOJA_CONFIG).Range(CFG_ANIO).Value2)
    Exit Function
sinAnio:
    AnioDelLibro = 0
End Function

Public Function RutaArchivo() As String
    On Error GoTo sinRuta
    RutaArchivo = Libro().Path
    Exit Function
sinRuta:
    RutaArchivo = ""
End Function

' ---------------------------------------------------------------------------
' Busquedas por clave (nunca por posicion de fila)
' ---------------------------------------------------------------------------
' Indice de codigos del catalogo. Se reconstruye solo si cambia el contenido
' (evita recorrer 5.000 filas y normalizar 5.000 claves en cada validacion).
Public Function IndiceCatalogo() As Object
    Dim ws As Worksheet
    Dim f As Long
    Dim sello As String
    Dim clave As String
    If mIndiceCatalogo Is Nothing Then Set mIndiceCatalogo = CreateObject("Scripting.Dictionary")
    Set ws = Libro().Sheets(HOJA_CATALOGO)
    sello = CStr(UltimaFilaUsada(ws, COL_CODIGO)) & ":" & _
            CStr(Application.WorksheetFunction.CountA(ws.Range(ws.Cells(FILA_1, COL_CODIGO), _
            ws.Cells(FILA_FIN, COL_CODIGO))))
    If sello <> mIndiceSello Then
        mIndiceCatalogo.RemoveAll
        For f = FILA_1 To FILA_FIN
            clave = NormalizaClave(ws.Cells(f, COL_CODIGO).Value2)
            If clave <> "" Then
                If Not mIndiceCatalogo.Exists(clave) Then mIndiceCatalogo.Add clave, f
            End If
        Next f
        mIndiceSello = sello
    End If
    Set IndiceCatalogo = mIndiceCatalogo
End Function

Public Sub InvalidarIndice()
    Set mIndiceCatalogo = Nothing
    mIndiceSello = ""
End Sub

Public Function FilaEnCatalogo(ByVal codigo As Variant) As Long
    Dim clave As String
    Dim indice As Object
    clave = NormalizaClave(codigo)
    If clave = "" Then
        FilaEnCatalogo = 0
        Exit Function
    End If
    Set indice = IndiceCatalogo()
    If indice.Exists(clave) Then
        FilaEnCatalogo = indice(clave)
    Else
        FilaEnCatalogo = 0
    End If
End Function

Public Function UltimaFilaUsada(ByVal ws As Worksheet, ByVal columna As Long) As Long
    Dim f As Long
    For f = FILA_FIN To FILA_1 Step -1
        If Not EsVacio(ws.Cells(f, columna).Value2) Then
            UltimaFilaUsada = f
            Exit Function
        End If
    Next f
    UltimaFilaUsada = FILA_1 - 1
End Function
