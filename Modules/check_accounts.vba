Option Explicit
Option Private Module
Sub CheckAccounts()

    Dim ExistingAccounts As Range
    Dim ImportAccounts As Range
    Dim ImportAccount As Range
    Dim Match As Range


' 1. Define value of ExistingAccounts

    Sheets("COA").Activate
    Range("COA[Compte]").Select
    Set ExistingAccounts = Selection


' 2. Define value of ImportAccounts

    Sheets("IMPORT").Activate
    Range("A3", Range("A3").End(xlDown)).Select
    Set ImportAccounts = Selection
    
    
' 3. Define params of new accounts
    
    For Each ImportAccount In ImportAccounts
        With ExistingAccounts
            Set Match = .Find(ImportAccount, LookIn:=xlValues)
            If Not Match Is Nothing Then
                'Do nothing
            Else
                With NewAccountForm
                    .TextBoxNewAccount.Value = ImportAccount
                    .TextBoxDesription.Value = ImportAccount.Offset(0, 1)
                    .OptionButtonE.Value = True
                    .OptionButtonAsset.Value = True
                    .Show
                End With
            End If
        End With
    Next ImportAccount
    
End Sub
