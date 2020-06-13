Option Explicit
Option Private Module
Sub CreateAccount()

    Dim NewRow As ListRow

    Sheets("COA").Activate
    Set NewRow = ActiveSheet.ListObjects("COA").ListRows.Add
    
    With NewRow
        .Range(1) = NewAccountForm.TextBoxNewAccount.Value
        .Range(2) = NewAccountForm.TextBoxDesription.Value
        .Range(3) = DefineBFR()
        .Range(4) = DefineAL()
        .Range(5).Formula = "=IF(RIGHT([@Compte],1)=""9"",""Y"",""N"")"
        .Range(6) = NewAccountForm.ComboBoxType.Value
    End With

    MsgBox "Compte " & NewAccountForm.TextBoxNewAccount.Value & " créé.", vbOKOnly

End Sub
Function DefineBFR() As String
    If NewAccountForm.OptionButtonE Then
        DefineBFR = "BFR E"
    Else
        DefineBFR = "BFR HE"
    End If
End Function
Function DefineAL() As String
    If NewAccountForm.OptionButtonAsset Then
        DefineAL = "Asset"
    Else
        DefineAL = "Liability"
    End If
End Function

