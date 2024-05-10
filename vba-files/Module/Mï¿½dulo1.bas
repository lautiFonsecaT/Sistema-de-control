Attribute VB_Name = "M�dulo1"
Public Sub borrarCalendario(celda As Range)

    If Not isEmpty(celda.value) Then
        celda.value = ""
    End If
    
End Sub
Public Function verificarHoja(hoja As String)

    Dim hojaV As Worksheet
    On Error Resume Next
    Set hojaV = ThisWorkbook.sheets(hoja)
    On Error GoTo 0


    If Not hojaV Is Nothing Then
    verificarHoja = True
    Else: verificarHoja = False
    End If

End Function
Public Function buscarFecha(hoja As String, id As Long, fecha As Date, lastRow As Variant) As Boolean
    
    Dim mes As Worksheet
    Set mes = ThisWorkbook.sheets(hoja)
    Dim fila As Long
    Dim resultado As Variant
    Dim ingreso As Boolean
    Dim egreso As Boolean
    Dim columna As Range
    
    buscarFecha = False
      For fila = 1 To lastRow
          If mes.cells(fila, "A") = id Then
              If mes.cells(fila, "C") = fecha Then
                  buscarFecha = True
              End If
          End If
      Next fila


End Function

Public Function analisis()
End Function

