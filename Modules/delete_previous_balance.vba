Option Explicit
Option Private Module

Sub DeletePreviousBalance(Year, Month, Entity)
    
    Dim id As String
    
    id = Year & Month & Entity
        
    Sheets("DB").Activate
    ActiveSheet.ListObjects("DB").Range.AutoFilter Field:=12, Criteria1:=id
    Range("DB").Select
    Selection.EntireRow.Delete
    ActiveSheet.ListObjects("DB").Range.AutoFilter Field:=12

End Sub

