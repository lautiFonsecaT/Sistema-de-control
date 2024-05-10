VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub cmdAnalisis_Click()

  dim recuadroAnalisis as Range
  set recuadroAnalisis = ThisWorkbook.Names("recuadroAnalisis").RefersToRange

  If recuadroAnalisis.hidden = false Then
    recuadroAnalisis.EntireColumn.hidden = true
  else: recuadroAnalisis.EntireColumn.hidden = false
  End If
  
 
End Sub

'Este evento escucha las celdas establecidas y borra los datos del calendario cuando ocurre un cambio
Private Sub Worksheet_Change(ByVal Target As Range)


  Dim celdaMes As Range
  Dim celdaAño As Range
  
  
  ' Define la celda que quieres monitorear
  Set celdaMes = ThisWorkbook.sheets("INDEX").cells(10, "N")
  Set celdaAño = ThisWorkbook.sheets("INDEX").cells(10, "R")
  
  ' Verifica si la celda afectada está dentro de la celda que quieres monitorear
  If Not Intersect(Target, celdaMes) Is Nothing Or Not Intersect(Target, celdaAño) Is Nothing Then
    Dim matriz As Range
    Dim fila As Integer, columna As Integer
    Set matriz = ThisWorkbook.Names("calendar_index").RefersToRange
    For fila = 1 To 12
      For columna = 1 To 7
        If IsDate(matriz.cells(fila, columna).value) Then

        borrarCalendario (matriz.cells(fila + 1, columna))

        End If
      Next columna
    Next fila
  End If
End Sub


Private Sub cmd_buscar_Click()
  
'---(Valores de celdas que contienen la fecha de ingreso)---
'Dim Fecha_1 As Variant
'Dim Fecha_2 As Variant

Dim matriz As Range
Dim fila As Integer, columna As Integer
Dim id As Range
Dim mes As String
Dim fechaCalendario As Date
Dim fechaMes As Date
Dim lastRow As Variant
Dim filaMes As Integer
Dim idMes As Long
Dim hojaMes As Worksheet
Dim accion As String


'Aqui se asigna el valor de la celda de ingresos que luego verificaremos
Set matriz = ThisWorkbook.Names("calendar_index").RefersToRange
Set id = ThisWorkbook.Names("id_index").RefersToRange
'Se asigna mes en relacion a la celda "MES" en la planilla
mes = ThisWorkbook.sheets("INDEX").cells(10, "N")
                    
If verificarHoja(mes) Then
Set hojaMes = ThisWorkbook.sheets(mes)
lastRow = hojaMes.cells(hojaMes.Rows.Count, "C").End(xlUp).Row
Else: MsgBox "Mes seleccionado no cargado", 48, "Error"
  Exit Sub
End If

For fila = 1 To 12
 For columna = 1 To 7
  
   If Not isEmpty(matriz.cells(fila, columna).value) And IsDate(matriz.cells(fila, columna).value) And Not isEmpty(id) Then
    fechaCalendario = matriz.cells(fila, columna).value
    borrarCalendario (matriz.cells(fila + 1, columna))

    If buscarFecha(mes, id.value, fechaCalendario, lastRow) Then
      matriz.cells(fila + 1, columna).value = "A"
    Else: matriz.cells(fila + 1, columna).value = "F"
    End If

    End If
 Next columna
Next fila
End Sub
