Attribute VB_Name = "Sub_AbrirDocumento"
'Sub que abrira el archivo correspondiente usando el boton "abrirDoc"

Sub AbrirDocumento()
    
    'Definir variables
    Dim nombreArchivo As String
    Dim rutaArchivo As String
    Dim celdaNombreArchivo As Range
    
    'Asignar valores a las variables
    Set celdaNombreArchivo = ThisWorkbook.Sheets("consulta").Range("consultaMolde") 'celda que contiene el nombre del molde
    nombreArchivo = celdaNombreArchivo.Value 'Asignar valor de la celda a la variable
    rutaArchivo = BuscarRutaArchivo(nombreArchivo) 'usar funcion para obtener la ruta corespondiente
    
    If rutaArchivo <> "" Then
        Workbooks.Open rutaArchivo
    Else
        MsgBox "El archivo no se encuentra en la tabla."
    End If
    
End Sub
