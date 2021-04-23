Sub Macro_hide_rows_columns()

    Dim last_row As Long
    Dim last_column As Integer
    
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    last_column = Cells(1, Columns.Count).End(xlToLeft).Column

    ' hide rows
    Range(Cells(last_row + 1, 1), Cells(Rows.Count, 1)).EntireRow.Hidden = True
    
    ' hide columns
    Range(Cells(1, last_column + 1), Cells(1, Columns.Count)).EntireColumn.Hidden = True
     
End Sub

Sub Macro_conditional_formatting()
    
    Dim i As Integer
    Dim j As Integer
    For i = 2 To (Cells(1, Columns.Count).End(xlToLeft).Column)
        For j = 2 To 4
            ' cells selection
            Cells(j, i).Select
        
            ' delete any previous conditional rules
            Selection.FormatConditions.Delete
        
            If Not IsEmpty(Cells(5, i)) Then
                ' add less than condition
                Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                Formula1:=Cells(5, i)
                
                ' add bigger than condition
                Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
                Formula1:=Cells(6, i)
                
                        
                ' add formatting styles for the condition one above
                With Selection.FormatConditions(1).Font
                .Color = -16383844
                .TintAndShade = 0
                End With
                With Selection.FormatConditions(1).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 13551615
                .TintAndShade = 0
                End With
                Selection.FormatConditions(1).StopIfTrue = False
                
                ' add formatting styles for the condition two above
                With Selection.FormatConditions(2).Font
                .Color = -16383844
                .TintAndShade = 0
                End With
                With Selection.FormatConditions(2).Interior
                .PatternColorIndex = xlAutomatic
                .Color = 13551615
                .TintAndShade = 0
                End With
                Selection.FormatConditions(2).StopIfTrue = False
                
            End If
        Next j
    Next i
End Sub

