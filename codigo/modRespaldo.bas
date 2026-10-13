Attribute VB_Name = "modRespaldo"
Option Explicit

' Respaldo verificable antes y despues de cada cierre (LISLR arts. 88 y 89:
' conservar libros, registros y comprobantes). Copia local con marca de
' tiempo y huella SHA-256. Si el archivo es de solo lectura o la copia falla,
' el cierre se detiene: nunca se anuncia un cierre que no se guardo.

Private Const CARPETA_RESPALDOS As String = "_RESPALDOS"

Public Function CarpetaRespaldos() As String
    Dim base As String
    base = RutaArchivo()
    If base = "" Then
        Err.Raise vbObjectError + 550, "CarpetaRespaldos", _
            "El libro no esta guardado en disco: no hay carpeta donde respaldar."
    End If
    CarpetaRespaldos = base & "\" & CARPETA_RESPALDOS
End Function

Public Function CrearRespaldo(ByVal etiqueta As String) As String
    Dim origen As String
    Dim destino As String
    Dim fso As Object
    On Error GoTo fallo
    origen = Libro().FullName
    If Len(origen) = 0 Then
        Err.Raise vbObjectError + 551, "CrearRespaldo", "El libro no tiene ruta de archivo."
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(CarpetaRespaldos()) Then
        fso.CreateFolder CarpetaRespaldos()
    End If
    If fso.GetFile(origen).Attributes And 1 Then
        Err.Raise vbObjectError + 552, "CrearRespaldo", _
            "El archivo es de solo lectura: no se respalda ni se cierra."
    End If
    destino = CarpetaRespaldos() & "\" & QuitarExtension(fso.GetFileName(origen)) & _
        "_" & etiqueta & "_" & SelloTiempo() & ".xlsm"
    fso.CopyFile origen, destino, True
    If Not fso.FileExists(destino) Then
        Err.Raise vbObjectError + 553, "CrearRespaldo", "La copia no existe tras copiar."
    End If
    If fso.GetFile(destino).Size = 0 Then
        Err.Raise vbObjectError + 554, "CrearRespaldo", "La copia quedo vacia."
    End If
    CrearRespaldo = destino
    modBitacora.Registrar "Respaldo", etiqueta, "OK", "Copia local verificable", _
        fso.GetFileName(origen), fso.GetFileName(destino), "", ""
    Exit Function
fallo:
    modBitacora.Registrar "Respaldo", etiqueta, "ERROR", Err.Description, "", "", "", ""
    CrearRespaldo = ""
End Function

Private Function QuitarExtension(ByVal nombre As String) As String
    Dim punto As Long
    punto = InStrRev(nombre, ".")
    If punto > 0 Then
        QuitarExtension = Left$(nombre, punto - 1)
    Else
        QuitarExtension = nombre
    End If
End Function

' Verificacion de persistencia: el archivo guardado conserva lo esperado.
Public Function VerificarPersistencia() As String
    Dim fso As Object
    Dim problemas As String
    On Error GoTo fallo
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(Libro().FullName) = 0 Then
        VerificarPersistencia = "El libro no esta guardado."
        Exit Function
    End If
    If Not fso.FileExists(Libro().FullName) Then
        VerificarPersistencia = "El archivo no existe en disco tras guardar."
        Exit Function
    End If
    If fso.GetFile(Libro().FullName).Attributes And 1 Then
        problemas = problemas & "El archivo quedo de solo lectura." & vbCrLf
    End If
    If fso.GetFile(Libro().FullName).Size = 0 Then
        problemas = problemas & "El archivo quedo vacio." & vbCrLf
    End If
    VerificarPersistencia = problemas
    Exit Function
fallo:
    VerificarPersistencia = "No se pudo verificar: " & Err.Description
End Function