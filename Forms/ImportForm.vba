Option Explicit
Private Sub UserForm_Initialize()

'1. Initialize list of years and months
    
    With ComboBoxYear
        .AddItem "2018"
        .AddItem "2019"
        .AddItem "2020"
        .AddItem "2021"
        .AddItem "2022"
        .AddItem "2023"
        .AddItem "2024"
        .AddItem "2025"
        .Value = Format(Date, "yyyy")
    End With
    With ComboBoxMonth
        .AddItem "01"
        .AddItem "02"
        .AddItem "03"
        .AddItem "04"
        .AddItem "05"
        .AddItem "06"
        .AddItem "07"
        .AddItem "08"
        .AddItem "09"
        .AddItem "10"
        .AddItem "11"
        .AddItem "12"
        .Value = Format(Date, "mm")
    End With
    
' 2. Initialize entities

    With ComboBoxEntity
        .AddItem "Americas"
        .AddItem "France"
        .AddItem "Tunisia"
    End With

End Sub
Private Sub CommandButtonImport_Click()
    
    Dim Response As Variant
    
    If ImportForm.ComboBoxEntity.Value = "" Then
        MsgBox "Veuillez complèter tous les champs avant de continuer !", vbExclamation
    Else
        Response = MsgBox( _
            "Vous êtes sur le point d'importer une balance avec les paramètres suivants:" & _
            Chr(10) & " " & Chr(10) & _
            "Année : " & ComboBoxYear.Value & Chr(10) & _
            "Mois : " & ComboBoxMonth.Value & Chr(10) & _
            "Pays : " & ComboBoxEntity.Value & Chr(10) & " " & Chr(10) & _
            "Cliquez sur OK pour continuer ou sur ANNULER pour modifier ces paramètres", _
            vbOKCancel, _
            "Résumé import" _
        )
    End If
    
    If Response = vbOK Then
        Call CheckAccounts
        Call SortChartOfAccounts
        Call ImportBalance
        'Call SortDB
        Sheets("IMPORT").Select
        Unload Me
    End If
    
End Sub
