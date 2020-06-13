Option Explicit
Private Sub UserForm_Initialize()

'1. Initialize list of types
    
    With ComboBoxType
        .AddItem "Shareholder 's Equity"
        .AddItem "Cash"
        .AddItem "Net Inventory"
        .AddItem "Net Property And Equipment"
        .AddItem "Non Current Assets"
        .AddItem "Non Current Liabilities"
        .AddItem "Other Payables"
        .AddItem "Other Receivables"
        .AddItem "Prepaid Expenses"
        .AddItem "Prepaid Incomes"
        .AddItem "Shareholder 's Equity"
        .AddItem "Trade Payables"
        .AddItem "Trade Receivables"
    End With

End Sub

Private Sub UserForm_Terminate()
    
    MsgBox "Vous avez interrompu la procédure d'import !", vbCritical
    End
    
End Sub
Private Sub CommandButtonNewAccount_Click()

    If NewAccountForm.ComboBoxType.Value = "" Then
        MsgBox "Veuillez complèter tous les champs avant de continuer !", vbExclamation
    Else
        Call CreateAccount
        NewAccountForm.Hide
    End If

End Sub

