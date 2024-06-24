Attribute VB_Name = "Sub_ConsultarHistoria"
'Subproseso para consultar la hisotira de cada molde

Sub LimpiarConsulta() ' limpiar la hoja consulta desde la fila 10
        Dim destinoWs As Worksheet
        Set destinoWs = ThisWorkbook.Sheets("consulta")
        destinoWs.Rows("10:" & destinoWs.Rows.Count).ClearContents
        destinoWs.Rows("10:" & destinoWs.Rows.Count).Interior.Color = RGB(221, 235, 247)
End Sub

Sub ConsultarHistoria()
    
    ' Definir variables
    Dim nombreMolde As String
    Dim rutaArchivo As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim respuesta As VbMsgBoxResult
    Dim excelApp As Excel.Application
    Dim destinoWb As Workbook
    Dim destinoWs As Worksheet
    
    ' Obtener nombre del molde y ruta del archivo
    nombreMolde = ThisWorkbook.Sheets("consulta").Range("consultaMolde").Value ' Obtener nombre de molde
    rutaArchivo = BuscarRutaArchivo(nombreMolde) ' Obtener ruta del archivo
    
    ' Confirmar si se encontro la ruta del archivo
    If rutaArchivo = "" Then
        MsgBox ("No se encontro el documento"), vbExclamation
        Exit Sub
    End If
    
    ' Confirmar consulta
    respuesta = MsgBox("Consultar historia de: " & nombreMolde & "?", vbQuestion + vbYesNo, "Confirmar Registro")
    If respuesta = vbNo Then
        Exit Sub
    End If
    
    ' Desactivar actualizacion de pantalla
    Application.ScreenUpdating = False
    
    ' Crear instancia y abir archivo en segundo plano
    Set excelApp = New Excel.Application
    Set destinoWb = Excel.Application.Workbooks.Open(rutaArchivo)
    
    'Verificar si el archiv se abreio
    If destinoWb Is Nothing Then
        MsgBox "No se pudo abrir el documento", vbExclamation
        Application.ScreenUpdating = True
        excelApp.Quit
        Set excelApp = Nothing
        Exit Sub
    End If
    
    ' Asignar referencias a archivos abiertos
    Set ws = destinoWb.Sheets("HISTORIA")
    Set tbl = ws.ListObjects("historia")
    
    'Verificar si la tabla hisotira existe
    If tbl Is Nothing Then
    MsgBox "No se encontro la tabla ""historia"" del molde: " & nombreMolde, vbExclamation
        destinoWb.Close False
        excelApp.Quit
        Set excelApp = Nothing
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    
    ' limpiar la hoja consulta
    Set destinoWs = ThisWorkbook.Sheets("consulta")
    LimpiarConsulta
    
    ' Copiamos la tabla hisotia y pegamos desde la celda A10 de la consula
    tbl.Range.Copy
    destinoWs.Range("A10").PasteSpecial
        Paste = xlPasteValues
    destinoWs.Range("A10").PasteSpecial
        Paste = xlPasteFormats
    destinoWs.Columns.AutoFit
    
    'Cerrar archivo sin guardar y eliminar referencias
    destinoWb.Close False
    excelApp.Quit
    Set excelApp = Nothing
    
    ' Activar actualizacion de pantalla
    Application.ScreenUpdating = True
    
    'Mensaje de conslta exitosa
    MsgBox "Consulta exitosa", vbInformation
End Sub
