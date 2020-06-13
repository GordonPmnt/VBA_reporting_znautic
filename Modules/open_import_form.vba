Option Explicit

Sub OpenImportForm()

    Sheets("IMPORT").Activate
    
    If Range("A3") = "" Then
        MsgBox "Vérifiez que la balance est correctement collée ci-dessous et qu'elle démarre à la cellule 'A3'", vbCritical
        End
    Else
        ImportForm.Show
    End If
    
End Sub

