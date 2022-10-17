Attribute VB_Name = "Module1"
Sub BuildShapeKey()
Attribute BuildShapeKey.VB_Description = "Build the Shape Key field based on the Shape Image value"
Attribute BuildShapeKey.VB_ProcData.VB_Invoke_Func = " \n14"
'
' BuildShapeKey Macro
' Build the Shape Key field based on the Shape Image value
'
'
    Dim ws As Worksheet
    Dim wbk As Workbook
    
    'Dim wsCols As Long
    Dim wsRows As Long
    
    Dim wsShapeKeyCol As Integer
    wsShapeKeyCol = 3
    
    Dim wsShapeImageCol As Integer
    wsShapeImageCol = 4
    
    Dim nRow As Integer 'loop row counter
    Dim value As String
    
    Application.CommandBars("Help").Visible = False
    Application.Goto Reference:="BuildShapeKey"
    
    wsRows = GetMaxRows()
    
    For nRow = 2 To wsRows  'start at 2 to skip the header
        value = GetValue(nRow, wsShapeImageCol)
        If Len(value) > 0 Then
            ' lets set the Shape Key field for this row
            If SetValue(nRow, wsShapeKeyCol, value) = True Then
                MsgBox "Error setting the Shape Key value for this Row:" + nRow
            End If
        End If
    Next
End Sub
' get the cell value at the row, col position
Function GetValue(row As Integer, col As Integer)
    GetValue = ActiveSheet.Cells(row, col)
End Function
' set the cell value based on the Row, Col position.
' also validate the value set is same as value in cell after it has been updated
Function SetValue(row As Integer, col As Integer, data As String)
    ActiveSheet.Cells(row, col).value = (data & ":" & row)
    value = GetValue(row, col)
    If value <> data Then
        SetValue = True
    End If
    SetValue = False
End Function
'Get Maximum Rows
Function GetMaxRows()
    Dim lngMaxRows As Long
     
    'Find the last non-blank cell in column A...
    'Or, you may use any column from A to H, that has data
    With Sheet1
        lngMaxRows = .Cells(.Rows.Count, "A").End(xlUp).row
    End With
    ' MsgBox "The count of number of rows with data: " & lngMaxRows
    GetMaxRows = lngMaxRows
End Function
'Maximum Columns
Function GetMaxColumns()
    Dim lngMaxCols As Long
     
    'Find the last non-blank cell in row 1...
    With Sheet1
        lngMaxCols = .Cells(1, .Columns.Count).End(xlToLeft).Column
    End With
    ' MsgBox "The Column Count with data: " & lngMaxCols
    GetMaxColumns = lngMaxCols
End Function


