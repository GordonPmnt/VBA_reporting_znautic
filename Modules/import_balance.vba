Option Explicit
Option Private Module
Sub ImportBalance()

    Dim Year As String
    Dim Month As String
    Dim Entity As String
    Dim Response As Variant
    Dim ImportAccounts As Range
    Dim ImportAccount As Range
    Dim NewRow As ListRow
    
    Year = ImportForm.ComboBoxYear.Value
    Month = ImportForm.ComboBoxMonth.Value
    Entity = ImportForm.ComboBoxEntity.Value

' 1. Check whether import with given params already exists

    If ImportAlreadyExists(Year, Month, Entity) Then
        Response = MsgBox( _
            "Une balance ayant ces paramètres existe déjà. Cliquez sur OK pour poursuivre et écraser les données existantes.", _
            vbOKCancel _
            )
        If Response = vbOK Then
            Call DeletePreviousBalance(Year, Month, Entity)
        Else
            Sheets("IMPORT").Activate
            End
        End If
    End If


' 2. Executes import

    Sheets("IMPORT").Activate
    Range("A3", Range("A3").End(xlDown)).Select
    Set ImportAccounts = Selection
    
    Sheets("DB").Activate
    Columns("A:B").Select
    Selection.NumberFormat = "@"
    
    For Each ImportAccount In ImportAccounts
        Set NewRow = ActiveSheet.ListObjects("DB").ListRows.Add
        With NewRow
            .Range(1) = Year
            .Range(2) = Format(Month, "00")
            .Range(3) = Entity
            .Range(4) = ImportAccount
            .Range(5) = ImportAccount.Offset(0, 1)
            .Range(6) = ImportAccount.Offset(0, 2)
            .Range(7) = ImportAccount.Offset(0, 3)
            .Range(8) = ImportAccount.Offset(0, 4)
            .Range(9).Formula = "=VLOOKUP([@Compte],COA,3,FALSE)"
            .Range(10).Formula = "=VLOOKUP([@Compte],COA,5,FALSE)"
            .Range(11).Formula = "=VLOOKUP([@Compte],COA,4,FALSE)"
            .Range(12).Formula = "=[@Année]&[@Mois]&[@Pays]"
            .Range(13).Formula = "=VLOOKUP([@Compte],COA,6,FALSE)"
        End With
        Next ImportAccount

    Response = MsgBox( _
        "Import terminé. Souhaitez-vous réinitialiser la feuille d'import ?", _
        vbOKCancel _
        )
    If Response = vbOK Then
        Call DeleteImportedBalance
    End If

End Sub
Function ImportAlreadyExists(Year, Month, Entity) As Boolean
    
    Dim ids As Range
    Dim id As String
    Dim Match As Range
        
    id = Year & Month & Entity
    
    Sheets("DB").Activate
    
    Range("DB[ID]").Select
    Set ids = Selection
    Set Match = ids.Find(id, LookIn:=xlValues)
            
    If Not Match Is Nothing Then
        ImportAlreadyExists = True
    Else
        ImportAlreadyExists = False
    End If

End Function
