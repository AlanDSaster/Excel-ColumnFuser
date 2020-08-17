Attribute VB_Name = "ColumnFusion"
Option Explicit

'This module will fuse all selected columns into the left-most column
'It will find the first non-empty cell in each row and retain its value

Sub FuseSelection()
    
    FuseColumnValues Selection
    
End Sub

Sub FuseColumnValues(tr As Range)
    
    Set tr = Intersect(tr, tr.Parent.UsedRange)
    
    Dim resultColumn As Range
    Set resultColumn = tr.Columns(1)
    
    Dim rCol As Range
    Dim rRow As Range
    
    
    For Each rRow In tr.Rows
        
        resultColumn.Cells(rRow.Row).Value = GetFirstNonBlankValue(rRow)
        
    Next rRow
    
    DeleteSecondaryColumns tr
    
End Sub

Function GetFirstNonBlankValue(r As Range) As String
    
    Dim c As Range
    Dim s As String
    
    For Each c In r.Cells
    
        s = c.Value
        
        If Len(Trim(s)) > 0 Then
            GetFirstNonBlankValue = s
            Exit For
        End If
        
    Next c
    
End Function

Sub DeleteSecondaryColumns(r As Range)
    
    Dim subrange As Range
    
    If r.Columns.Count > 1 Then
        
        Set subrange = Range(r.Cells(1, 2), r.Cells(r.Cells.Count)).EntireColumn
        Debug.Print "Deleting " & subrange.Address
        subrange.EntireColumn.Delete
        
    End If
    
End Sub
