Attribute VB_Name = "fun_BuscarRutaArchivo"
'Funcion que buscara la ruta del archivo en la tabla listaMoldes segun el nombre del molde

Function BuscarRutaArchivo(nombreArchivo As String)

    'Definir Variables
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim columNombre As Integer
    Dim columRuta As Integer
    Dim i As Long
    Dim rutaArchivo As String
    
    'Asignar valores a las Variables
    Set ws = ThisWorkbook.Sheets("listaMoldes") 'hoja listaMoldes
    Set tbl = ws.ListObjects("listaMoldes") 'tabla listaMoldes
    columnaNombre = tbl.ListColumns("NOMBRE").Index 'Indice de la columna "NOMBRE"
    columnaRuta = tbl.ListColumns("RUTA").Index 'Indice de la columna "RUTA"
    rutaArchivo = "" 'Inicializar ruta como string vacio
    
    For i = 1 To tbl.ListRows.Count
        
        'Buscar nombre del molde
        If tbl.DataBodyRange(i, columnaNombre).Value = nombreArchivo Then
            'Si nombre del archivo coincide, tomar la ruta del archivo
            rutaArchivo = tbl.DataBodyRange(i, columnaRuta).Value
            Exit For
        End If
    Next i
    
    'Retornar la ruta del archivo
    BuscarRutaArchivo = rutaArchivo

End Function
