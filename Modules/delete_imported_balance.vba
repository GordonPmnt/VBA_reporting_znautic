Sub DeleteImportedBalance()

    Sheets("IMPORT").Activate

    Rows("3:3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    Range("A3").Select

End Sub
