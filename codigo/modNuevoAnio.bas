Attribute VB_Name = "modNuevoAnio"
Option Explicit

' Crear nuevo anio: solo con diciembre cerrado valido. Genera OTRO archivo,
' conserva el anterior intacto, copia estructura, VBA y catalogo, reinicia
' estados y movimientos, y transfiere diciembre a enero con conciliacion.

Public Function CrearNuevoAnio() As String
    Dim rutaNueva As String
    Dim copia As Workbook
    Dim seguridadPrevia As Long
    Dim anioViejo As Long
    On Error GoTo fallo

    If Reentrada Then Exit Function
    Entrar
    If modEstados.EstadoMes(12) <> EST_CERRADO Then
        Salir
        Err.Raise vbObjectError + 570, "CrearNuevoAnio", _
            "Diciembre no esta CERRADO: no se crea el nuevo anio."
    End If
    If modRespaldo.VerificarPersistencia() <> "" Then
        Salir
        Err.Raise vbObjectError + 571, "CrearNuevoAnio", _
            "El archivo actual no verifica persistencia: guarde y vuelva a intentar."
    End If
    anioViejo = AnioDelLibro()
    If modUI.PruebaSilencioso Then
        ' En modo de pruebas no hay dialogo de guardado.
        rutaNueva = ""
    Else
        rutaNueva = Application.GetSaveAsFilename( _
            Libro().Path & "\LIBRO_INVENTARIOS_" & (anioViejo + 1) & ".xlsm", _
            "Libros de inventarios,*.xlsm")
    End If
    If Trim$(rutaNueva) = "" Then
        Salir
        CrearNuevoAnio = ""
        Exit Function
    End If
    If Right$(rutaNueva, 5) <> ".xlsm" Then rutaNueva = rutaNueva & ".xlsm"
    If FileExists(rutaNueva) Then
        Salir
        Err.Raise vbObjectError + 572, "CrearNuevoAnio", _
            "La ruta ya existe. No se sobrescribe ningun archivo."
    End If
    If rutaNueva = Libro().FullName Then
        Salir
        Err.Raise vbObjectError + 573, "CrearNuevoAnio", _
            "El destino es el propio archivo de trabajo."
    End If

    Libro().Save
    Libro().SaveCopyAs rutaNueva
    If Not FileExists(rutaNueva) Then
        Salir
        Err.Raise vbObjectError + 574, "CrearNuevoAnio", "No se creo la copia."
    End If

    seguridadPrevia = Application.AutomationSecurity
    Application.AutomationSecurity = 3          ' macros de la copia deshabilitadas
    On Error Resume Next
    Set copia = Application.Workbooks.Open(rutaNueva, 0, False, 5)
    On Error GoTo fallo
    If copia Is Nothing Then
        Application.AutomationSecurity = seguridadPrevia
        Salir
        Err.Raise vbObjectError + 575, "CrearNuevoAnio", "No se pudo abrir la copia nueva."
    End If

    TransferirDiciembreANuevoAnio copia, anioViejo
    copia.Save
    copia.Close False
    Application.AutomationSecurity = seguridadPrevia

    ' Verificacion: se reabre el archivo nuevo y el anterior
    Dim verificacion As Workbook
    On Error Resume Next
    Set verificacion = Application.Workbooks.Open(rutaNueva, 0, True, 5)
    On Error GoTo fallo
    If verificacion Is Nothing Then
        Salir
        Err.Raise vbObjectError + 576, "CrearNuevoAnio", _
            "El archivo nuevo no se pudo reabrir para verificar."
    End If
    If Not TieneVBA(verificacion) Then
        verificacion.Close False
        Salir
        Err.Raise vbObjectError + 577, "CrearNuevoAnio", _
            "El archivo nuevo abrio sin proyecto VBA."
    End If
    verificacion.Close False
    If Not FileExists(Libro().FullName) Then
        Salir
        Err.Raise vbObjectError + 578, "CrearNuevoAnio", "El archivo original desaparecio."
    End If

    modBitacora.Registrar "Creacion de nuevo anio", rutaNueva, "OK", _
        "Diciembre " & anioViejo & " transferido a enero con conciliacion", _
        anioViejo, anioViejo + 1, "", ""
    Salir
    CrearNuevoAnio = rutaNueva
    Exit Function
fallo:
    On Error Resume Next
    If Not copia Is Nothing Then copia.Close False
    Application.AutomationSecurity = seguridadPrevia
    On Error GoTo 0
    modEstadoExcel.PropagarError "CrearNuevoAnio", "modNuevoAnio"
    Salir
End Function

Private Function FileExists(ByVal ruta As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(ruta)
End Function

Private Function TieneVBA(ByVal libro As Workbook) As Boolean
    ' No se consulta VBProject: eso exige "Acceso al modelo de objetos de VB"
    ' en el equipo del usuario y ademas fuerza la carga del proyecto en el VBE.
    ' HasVBProject no requiere ninguna opcion de confianza.
    On Error GoTo sinProyecto
    TieneVBA = libro.HasVBProject
    Exit Function
sinProyecto:
    TieneVBA = False
End Function

' Reinicia estados y movimientos del anio nuevo y pasa el saldo de
' diciembre a enero como valores, con conciliacion por codigo.
Public Sub TransferirDiciembreANuevoAnio(ByVal copia As Workbook, ByVal anioViejo As Long)
    Dim nombres As Variant
    Dim diciembre As Worksheet
    Dim enero As Worksheet
    Dim estados As Worksheet
    Dim i As Long
    Dim f As Long
    Dim codigo As String
    Dim faltantes As String
    Dim cat As Worksheet
    Dim filaCat As Long
    nombres = NombresMes()

    ' La copia es una fotografia del anio anterior: sus doce hojas de mes se
    ' renombran al anio nuevo. Diciembre se toma por referencia ANTES de
    ' renombrar y se vacia DESPUES de transferir, porque es el saldo de origen.
    Set diciembre = copia.Sheets(anioViejo & "-DIC")
    For i = 1 To 12
        copia.Sheets(anioViejo & "-" & nombres(i - 1)).Name = (anioViejo + 1) & "-" & nombres(i - 1)
    Next i
    Set enero = copia.Sheets((anioViejo + 1) & "-ENE")
    Set estados = copia.Sheets(HOJA_ESTADOS)
    Set cat = copia.Sheets(HOJA_CATALOGO)

    For i = 1 To 12
        estados.Cells(modEstados.FilaEstado(i), 1).Value = (anioViejo + 1) & "-" & nombres(i - 1)
        estados.Cells(modEstados.FilaEstado(i), 2).Value = EST_NO_ABIERTO
        estados.Cells(modEstados.FilaEstado(i), 3).ClearContents
        estados.Cells(modEstados.FilaEstado(i), 4).ClearContents
        estados.Cells(modEstados.FilaEstado(i), 5).ClearContents
        estados.Cells(modEstados.FilaEstado(i), 7).ClearContents
        estados.Cells(modEstados.FilaEstado(i), 8).ClearContents
        estados.Cells(modEstados.FilaEstado(i), 9).ClearContents
    Next i
    estados.Cells(modEstados.FilaEstado(1), 2).Value = EST_APERTURA_PEND
    copia.Sheets(HOJA_CONFIG).Range(CFG_ANIO).Value = anioViejo + 1

    ' Se vacian los once meses que no son la fuente del saldo.
    For i = 1 To 11
        With copia.Sheets((anioViejo + 1) & "-" & nombres(i - 1))
            .Unprotect Password:=CLAVE_PROTECCION
            .Range(.Cells(FILA_1, COL_CODIGO), .Cells(FILA_FIN, COL_ESTADO)).ClearContents
            .Cells.Locked = True
        End With
    Next i

    For f = FILA_1 To FILA_FIN
        codigo = Trim$(CStr(diciembre.Cells(f, COL_CODIGO).Value2))
        If codigo <> "" Then
            filaCat = BuscarFilaEnHoja(cat, codigo)
            If filaCat = 0 Then
                faltantes = faltantes & codigo & vbCrLf
            Else
                enero.Cells(f, COL_CODIGO).Value = codigo
                enero.Cells(f, COL_DESCRIPCION).Value = cat.Cells(filaCat, 2).Value2
                enero.Cells(f, COL_UNIDAD).Value = cat.Cells(filaCat, 3).Value2
                enero.Cells(f, COL_CANT_INI).Value = Val(diciembre.Cells(f, COL_CANT_FIN).Value2)
                enero.Cells(f, COL_VAL_INI).Value = Val(diciembre.Cells(f, COL_VAL_FIN).Value2)
                enero.Cells(f, COL_CANT_ENT).Value = 0
                enero.Cells(f, COL_VAL_ENT).Value = 0
                enero.Cells(f, COL_CANT_VTA).Value = 0
                enero.Cells(f, COL_CANT_OTR).Value = 0
            End If
        End If
    Next f

    ' Diciembre del anio nuevo arranca vacio: su saldo ya quedo en enero.
    With diciembre
        .Unprotect Password:=CLAVE_PROTECCION
        .Range(.Cells(FILA_1, COL_CODIGO), .Cells(FILA_FIN, COL_ESTADO)).ClearContents
        .Cells.Locked = True
    End With

    If faltantes <> "" Then
        estados.Cells(modEstados.FilaEstado(1), 9).Value = "Saldos de " & anioViejo & _
            "-DIC sin codigo en el catalogo: " & faltantes
    End If
End Sub

Private Function BuscarFilaEnHoja(ByVal ws As Worksheet, ByVal codigo As String) As Long
    Dim f As Long
    Dim clave As String
    clave = NormalizaClave(codigo)
    For f = FILA_1 To FILA_FIN
        If NormalizaClave(ws.Cells(f, 1).Value2) = clave Then
            BuscarFilaEnHoja = f
            Exit Function
        End If
    Next f
    BuscarFilaEnHoja = 0
End Function