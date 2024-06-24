Attribute VB_Name = "Sub_AgregarRegistro"
'Subproseso que escribe datos de la hoja "registro" en la hoja "HISTORIA" del documento seleccionado

Sub LimpiarRegistro()
    Dim respuesta As VbMsgBoxResult
    
    ' Confirmar realizar registro
    respuesta = MsgBox("¿Limpiar el registro?", vbQuestion + vbYesNo, "Confirmar Limpiar Registro")
    If respuesta = vbNo Then
        Exit Sub
    End If
    
    ' Limpiar el contenido de los rangos nombrados
    With ThisWorkbook.Sheets("registro")
        .Range("molde").ClearContents
        .Range("fecha").ClearContents
        .Range("estado").ClearContents
        .Range("mantenimiento").ClearContents
        .Range("nAnuladas").ClearContents
        .Range("anuladas").ClearContents
        .Range("novedad").Value = ""
    End With
End Sub

Sub AgregarRegistro()
    
    ' Definir Variables
    Dim nombreMolde As String
    Dim rutaArchivo As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim respuesta As VbMsgBoxResult
    Dim excelApp As Excel.Application
    Dim destinoWb As Workbook
    
    ' Obtener nombre del molde y ruta del archivo
    nombreMolde = ThisWorkbook.Sheets("registro").Range("molde").Value ' Obtener nombre de molde
    rutaArchivo = BuscarRutaArchivo(nombreMolde) ' Obtener ruta del archivo
    
    ' Confirmar realizar registro
    respuesta = MsgBox("¿Realizar registro de historia en molde " & nombreMolde & "?", vbQuestion + vbYesNo, "Confirmar Registro")
    
    ' Verificar confirmación
    If respuesta = vbNo Then
        Exit Sub
    End If
    
    ' Verificar si se encontró la ruta del archivo
    If rutaArchivo = "" Then
        MsgBox "No se encontró el documento", vbExclamation
        Exit Sub
    End If
    
    ' Desactivar actualizaciones de pantalla
    Application.ScreenUpdating = False
    
    ' Crear una nueva instancia y abrir archivo en segundo plano
    Set excelApp = New Excel.Application
    Set destinoWb = excelApp.Workbooks.Open(rutaArchivo)
    
    ' Verificar si el archivo se abrió correctamente
    If destinoWb Is Nothing Then
        MsgBox "No se pudo abrir el documento", vbExclamation
        Application.ScreenUpdating = True
        excelApp.Quit
        Set excelApp = Nothing
        Exit Sub
    End If
    
    ' Asignar referencias al archivo abierto
    Set ws = destinoWb.Sheets("HISTORIA")
    Set tbl = ws.ListObjects("historia")
    
    ' Verificar si la tabla "historia" existe
    If tbl Is Nothing Then
        MsgBox "No se encontró la tabla ""historia"" del molde " & nombreMolde, vbExclamation
        destinoWb.Close False
        excelApp.Quit
        Set excelApp = Nothing
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' Agregar datos en la tabla "historia"
    With tbl.ListRows.Add
        .Range(1, tbl.ListColumns("FECHA").Index).Value = ThisWorkbook.Sheets("registro").Range("fecha").Value
        .Range(1, tbl.ListColumns("NOVEDAD").Index).Value = ThisWorkbook.Sheets("registro").Range("novedad").Value
        .Range(1, tbl.ListColumns("ESTADO").Index).Value = ThisWorkbook.Sheets("registro").Range("estado").Value
        .Range(1, tbl.ListColumns("MANTENIMIENTO").Index).Value = ThisWorkbook.Sheets("registro").Range("mantenimiento").Value
        .Range(1, tbl.ListColumns("# CAVIDADES ANULADAS").Index).Value = ThisWorkbook.Sheets("registro").Range("nAnuladas").Value
        .Range(1, tbl.ListColumns("CAVIDADES ANULADAS").Index).Value = ThisWorkbook.Sheets("registro").Range("anuladas").Value
    End With
    
    ' Guardar cambios en el archivo y cerrarlo
    destinoWb.Save
    destinoWb.Close
    
    ' Cerrar la instancia de Excel en segundo plano y limpiar referencias
    excelApp.Quit
    Set excelApp = Nothing
    
    ' Activar actualizaciones de pantalla
    Application.ScreenUpdating = True
    
    ' Mostrar mensaje de registro exitoso
    MsgBox "Registro exitoso", vbInformation
    
    ' limpiar registro
    LimpiarRegistro
End Sub
