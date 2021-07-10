Option Explicit
 
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

Dim LastRow As Integer
LastRow = wsTable1.Range("A1").CurrentRegion.Rows.Count

    If Target.Column = 7 And Target.Row <> 1 And Target.Row <= LastRow Then
        wsData.Range("O1") = Target.Address
        Categories.Show
    End If

End Sub

