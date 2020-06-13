Option Explicit
Option Private Module
Sub SortChartOfAccounts()
    
    Sheets("COA").Activate
    
    With ActiveSheet.ListObjects("COA").sort
        .SortFields.Clear
        .SortFields.Add2 _
            Key:=Range("COA[[#All],[Compte]]"), _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        .Apply
    End With
    
End Sub
Sub SortDB()

End Sub


