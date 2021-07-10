VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsTable1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit
 
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

Dim LastRow As Integer
LastRow = wsTable1.Range("A1").CurrentRegion.Rows.Count

    If Target.Column = 7 And Target.Row <> 1 And Target.Row <= LastRow Then
        wsData.Range("O1") = Target.Address
        Categories.Show
    End If

End Sub

