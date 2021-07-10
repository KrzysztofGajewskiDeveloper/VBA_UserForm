VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Categories 
   Caption         =   "Categories"
   ClientHeight    =   13410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   24030
   OleObjectBlob   =   "UserForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Categories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Group_A_Change()
     
    If Group_A.Value = True Then
    
        With lbNumbers
            .Visible = True
            .Value = 1
        End With
        
        With tbGroupA1
            .Visible = True
            .SetFocus
        End With
        
        RiskA1.Visible = True
        RegionA1.Visible = True
        
    Else
    
        lbNumbers.Visible = False
        tbGroupA1.Visible = False
        tbGroupA2.Visible = False
        tbGroupA3.Visible = False
        tbGroupA4.Visible = False
        tbGroupA5.Visible = False
    
        RiskA1.Visible = False
        RiskA2.Visible = False
        RiskA3.Visible = False
        RiskA4.Visible = False
        RiskA5.Visible = False
        
        RegionA1.Visible = False
        RegionA2.Visible = False
        RegionA3.Visible = False
        RegionA4.Visible = False
        RegionA5.Visible = False
        
    End If

End Sub

Private Sub Group_B_Change()
     
    If Group_B.Value = True Then
      
        With lbNumbers2
            .Visible = True
            .Value = 1
        End With
        
        With tbGroupB1
            .Visible = True
            .SetFocus
        End With
        
        RiskB1.Visible = True
        RegionB1.Visible = True
        
    Else
    
        lbNumbers2.Visible = False
        tbGroupB1.Visible = False
        tbGroupB2.Visible = False
        tbGroupB3.Visible = False
        tbGroupB4.Visible = False
        tbGroupB5.Visible = False
        
        RiskB1.Visible = False
        RiskB2.Visible = False
        RiskB3.Visible = False
        RiskB4.Visible = False
        RiskB5.Visible = False
    
        RegionB1.Visible = False
        RegionB2.Visible = False
        RegionB3.Visible = False
        RegionB4.Visible = False
        RegionB5.Visible = False
        
    End If


End Sub

Private Sub Group_C_Change()
    
    If Group_C.Value = True Then
      
        With lbNumbers3
            .Visible = True
            .Value = 1
        End With
        
        With tbGroupC1
            .Visible = True
            .SetFocus
        End With
        
        RiskC1.Visible = True
        RegionC1.Visible = True
        
    Else
    
        lbNumbers3.Visible = False
        tbGroupC1.Visible = False
        tbGroupC2.Visible = False
        tbGroupC3.Visible = False
        tbGroupC4.Visible = False
        tbGroupC5.Visible = False
        
        RiskC1.Visible = False
        RiskC2.Visible = False
        RiskC3.Visible = False
        RiskC4.Visible = False
        RiskC5.Visible = False
        
        RegionC1.Visible = False
        RegionC2.Visible = False
        RegionC3.Visible = False
        RegionC4.Visible = False
        RegionC5.Visible = False
        
    End If

End Sub

Private Sub Group_D_Change()

    If Group_D.Value = True Then
      
        With lbNumbers4
            .Visible = True
            .Value = 1
        End With
        
        With tbGroupD1
            .Visible = True
            .SetFocus
        End With
        
        RiskD1.Visible = True
        RegionD1.Visible = True
        
    Else
    
        lbNumbers4.Visible = False
        tbGroupD1.Visible = False
        tbGroupD2.Visible = False
        tbGroupD3.Visible = False
        tbGroupD4.Visible = False
        tbGroupD5.Visible = False
        
        RiskD1.Visible = False
        RiskD2.Visible = False
        RiskD3.Visible = False
        RiskD4.Visible = False
        RiskD5.Visible = False
        
        RegionD1.Visible = False
        RegionD2.Visible = False
        RegionD3.Visible = False
        RegionD4.Visible = False
        RegionD5.Visible = False
        
    End If

End Sub

Private Sub Group_E_Change()

    If Group_E.Value = True Then
      
        With lbNumbers5
            .Visible = True
            .Value = 1
        End With
        
        With tbGroupE1
            .Visible = True
            .SetFocus
        End With
        
        RiskE1.Visible = True
        RegionE1.Visible = True
        
    Else
    
        lbNumbers5.Visible = False
        tbGroupE1.Visible = False
        tbGroupE2.Visible = False
        tbGroupE3.Visible = False
        tbGroupE4.Visible = False
        tbGroupE5.Visible = False
        
        RiskE1.Visible = False
        RiskE2.Visible = False
        RiskE3.Visible = False
        RiskE4.Visible = False
        RiskE5.Visible = False
        
        RegionE1.Visible = False
        RegionE2.Visible = False
        RegionE3.Visible = False
        RegionE4.Visible = False
        RegionE5.Visible = False
        
    End If

End Sub

Private Sub lbNumbers_Click()

    Select Case lbNumbers.Value
    
        Case 1
             tbGroupA1.Visible = True
             tbGroupA2.Visible = False
             tbGroupA3.Visible = False
             tbGroupA4.Visible = False
             tbGroupA5.Visible = False
             tbGroupA1.SetFocus
    
             RegionA1.Visible = True
             RegionA2.Visible = False
             RegionA3.Visible = False
             RegionA4.Visible = False
             RegionA5.Visible = False
             
             RiskA1.Visible = True
             RiskA2.Visible = False
             RiskA3.Visible = False
             RiskA4.Visible = False
             RiskA5.Visible = False
    
        Case 2
             tbGroupA1.Visible = True
             tbGroupA2.Visible = True
             tbGroupA3.Visible = False
             tbGroupA4.Visible = False
             tbGroupA5.Visible = False
             tbGroupA2.SetFocus
    
             RegionA1.Visible = True
             RegionA2.Visible = True
             RegionA3.Visible = False
             RegionA4.Visible = False
             RegionA5.Visible = False
             
             RiskA1.Visible = True
             RiskA2.Visible = True
             RiskA3.Visible = False
             RiskA4.Visible = False
             RiskA5.Visible = False
    
        Case 3
             tbGroupA1.Visible = True
             tbGroupA2.Visible = True
             tbGroupA3.Visible = True
             tbGroupA4.Visible = False
             tbGroupA5.Visible = False
             tbGroupA3.SetFocus
    
             RegionA1.Visible = True
             RegionA2.Visible = True
             RegionA3.Visible = True
             RegionA4.Visible = False
             RegionA5.Visible = False
             
             RiskA1.Visible = True
             RiskA2.Visible = True
             RiskA3.Visible = True
             RiskA4.Visible = False
             RiskA5.Visible = False
             
        Case 4
             tbGroupA1.Visible = True
             tbGroupA2.Visible = True
             tbGroupA3.Visible = True
             tbGroupA4.Visible = True
             tbGroupA5.Visible = False
             tbGroupA4.SetFocus
    
             RegionA1.Visible = True
             RegionA2.Visible = True
             RegionA3.Visible = True
             RegionA4.Visible = True
             RegionA5.Visible = False
             
             RiskA1.Visible = True
             RiskA2.Visible = True
             RiskA3.Visible = True
             RiskA4.Visible = True
             RiskA5.Visible = False
             
        Case 5
             tbGroupA1.Visible = True
             tbGroupA2.Visible = True
             tbGroupA3.Visible = True
             tbGroupA4.Visible = True
             tbGroupA5.Visible = True
             tbGroupA5.SetFocus
    
             RegionA1.Visible = True
             RegionA2.Visible = True
             RegionA3.Visible = True
             RegionA4.Visible = True
             RegionA5.Visible = True
             
             RiskA1.Visible = True
             RiskA2.Visible = True
             RiskA3.Visible = True
             RiskA4.Visible = True
             RiskA5.Visible = True
    
    End Select

End Sub

Private Sub lbNumbers2_Click()

    Select Case lbNumbers2.Value
    
        Case 1
            tbGroupB1.Visible = True
            tbGroupB2.Visible = False
            tbGroupB3.Visible = False
            tbGroupB4.Visible = False
            tbGroupB5.Visible = False
            tbGroupB1.SetFocus
            
            RiskB1.Visible = True
            RiskB2.Visible = False
            RiskB3.Visible = False
            RiskB4.Visible = False
            RiskB5.Visible = False
            
            RegionB1.Visible = True
            RegionB2.Visible = False
            RegionB3.Visible = False
            RegionB4.Visible = False
            RegionB5.Visible = False
    
        Case 2
            tbGroupB1.Visible = True
            tbGroupB2.Visible = True
            tbGroupB3.Visible = False
            tbGroupB4.Visible = False
            tbGroupB5.Visible = False
            tbGroupB2.SetFocus
            
            RiskB1.Visible = True
            RiskB2.Visible = True
            RiskB3.Visible = False
            RiskB4.Visible = False
            RiskB5.Visible = False
    
            RegionB1.Visible = True
            RegionB2.Visible = True
            RegionB3.Visible = False
            RegionB4.Visible = False
            RegionB5.Visible = False
             
        Case 3
            tbGroupB1.Visible = True
            tbGroupB2.Visible = True
            tbGroupB3.Visible = True
            tbGroupB4.Visible = False
            tbGroupB5.Visible = False
            tbGroupB3.SetFocus
            
            RiskB1.Visible = True
            RiskB2.Visible = True
            RiskB3.Visible = True
            RiskB4.Visible = False
            RiskB5.Visible = False
            
            RegionB1.Visible = True
            RegionB2.Visible = True
            RegionB3.Visible = True
            RegionB4.Visible = False
            RegionB5.Visible = False
            
        Case 4
            tbGroupB1.Visible = True
            tbGroupB2.Visible = True
            tbGroupB3.Visible = True
            tbGroupB4.Visible = True
            tbGroupB5.Visible = False
            tbGroupB4.SetFocus
    
            RiskB1.Visible = True
            RiskB2.Visible = True
            RiskB3.Visible = True
            RiskB4.Visible = True
            RiskB5.Visible = False
            
            RegionB1.Visible = True
            RegionB2.Visible = True
            RegionB3.Visible = True
            RegionB4.Visible = True
            RegionB5.Visible = False
            
        Case 5
            tbGroupB1.Visible = True
            tbGroupB2.Visible = True
            tbGroupB3.Visible = True
            tbGroupB4.Visible = True
            tbGroupB5.Visible = True
            tbGroupB5.SetFocus
            
            RiskB1.Visible = True
            RiskB2.Visible = True
            RiskB3.Visible = True
            RiskB4.Visible = True
            RiskB5.Visible = True
            
            RegionB1.Visible = True
            RegionB2.Visible = True
            RegionB3.Visible = True
            RegionB4.Visible = True
            RegionB5.Visible = True
     
    End Select

End Sub

Private Sub lbNumbers3_Click()

    Select Case lbNumbers3.Value
    
        Case 1
            tbGroupC1.Visible = True
            tbGroupC2.Visible = False
            tbGroupC3.Visible = False
            tbGroupC4.Visible = False
            tbGroupC5.Visible = False
            tbGroupC1.SetFocus
            
            RiskC1.Visible = True
            RiskC2.Visible = False
            RiskC3.Visible = False
            RiskC4.Visible = False
            RiskC5.Visible = False
    
            RegionC1.Visible = True
            RegionC2.Visible = False
            RegionC3.Visible = False
            RegionC4.Visible = False
            RegionC5.Visible = False
    
        Case 2
        
            tbGroupC1.Visible = True
            tbGroupC2.Visible = True
            tbGroupC3.Visible = False
            tbGroupC4.Visible = False
            tbGroupC5.Visible = False
            tbGroupC2.SetFocus
            
            RiskC1.Visible = True
            RiskC2.Visible = True
            RiskC3.Visible = False
            RiskC4.Visible = False
            RiskC5.Visible = False
            
            RegionC1.Visible = True
            RegionC2.Visible = True
            RegionC3.Visible = False
            RegionC4.Visible = False
            RegionC5.Visible = False
             
        Case 3
            tbGroupC1.Visible = True
            tbGroupC2.Visible = True
            tbGroupC3.Visible = True
            tbGroupC4.Visible = False
            tbGroupC5.Visible = False
            tbGroupC3.SetFocus
            
            RiskC1.Visible = True
            RiskC2.Visible = True
            RiskC3.Visible = True
            RiskC4.Visible = False
            RiskC5.Visible = False
            
            RegionC1.Visible = True
            RegionC2.Visible = True
            RegionC3.Visible = True
            RegionC4.Visible = False
            RegionC5.Visible = False
    
        Case 4
            tbGroupC1.Visible = True
            tbGroupC2.Visible = True
            tbGroupC3.Visible = True
            tbGroupC4.Visible = True
            tbGroupC5.Visible = False
            tbGroupC4.SetFocus
            
            RiskC1.Visible = True
            RiskC2.Visible = True
            RiskC3.Visible = True
            RiskC4.Visible = True
            RiskC5.Visible = False
            
            RegionC1.Visible = True
            RegionC2.Visible = True
            RegionC3.Visible = True
            RegionC4.Visible = True
            RegionC5.Visible = False
            
        Case 5
            tbGroupC1.Visible = True
            tbGroupC2.Visible = True
            tbGroupC3.Visible = True
            tbGroupC4.Visible = True
            tbGroupC5.Visible = True
            tbGroupC5.SetFocus
            
            RiskC1.Visible = True
            RiskC2.Visible = True
            RiskC3.Visible = True
            RiskC4.Visible = True
            RiskC5.Visible = True
            
            RegionC1.Visible = True
            RegionC2.Visible = True
            RegionC3.Visible = True
            RegionC4.Visible = True
            RegionC5.Visible = True
    
    End Select

End Sub

Private Sub lbNumbers4_Click()

    Select Case lbNumbers4.Value
    
        Case 1
            tbGroupD1.Visible = True
            tbGroupD2.Visible = False
            tbGroupD3.Visible = False
            tbGroupD4.Visible = False
            tbGroupD5.Visible = False
            tbGroupD1.SetFocus
            
            RiskD1.Visible = True
            RiskD2.Visible = False
            RiskD3.Visible = False
            RiskD4.Visible = False
            RiskD5.Visible = False
            
            RegionD1.Visible = True
            RegionD2.Visible = False
            RegionD3.Visible = False
            RegionD4.Visible = False
            RegionD5.Visible = False
            
        Case 2
            tbGroupD1.Visible = True
            tbGroupD2.Visible = True
            tbGroupD3.Visible = False
            tbGroupD4.Visible = False
            tbGroupD5.Visible = False
            tbGroupD2.SetFocus
             
            RiskD1.Visible = True
            RiskD2.Visible = True
            RiskD3.Visible = False
            RiskD4.Visible = False
            RiskD5.Visible = False
            
            RegionD1.Visible = True
            RegionD2.Visible = True
            RegionD3.Visible = False
            RegionD4.Visible = False
            RegionD5.Visible = False
            
        Case 3
            tbGroupD1.Visible = True
            tbGroupD2.Visible = True
            tbGroupD3.Visible = True
            tbGroupD4.Visible = False
            tbGroupD5.Visible = False
            tbGroupD3.SetFocus
            
            RiskD1.Visible = True
            RiskD2.Visible = True
            RiskD3.Visible = True
            RiskD4.Visible = False
            RiskD5.Visible = False
            
            RegionD1.Visible = True
            RegionD2.Visible = True
            RegionD3.Visible = True
            RegionD4.Visible = False
            RegionD5.Visible = False
            
    
        Case 4
            tbGroupD1.Visible = True
            tbGroupD2.Visible = True
            tbGroupD3.Visible = True
            tbGroupD4.Visible = True
            tbGroupD5.Visible = False
            tbGroupD4.SetFocus
            
            RiskD1.Visible = True
            RiskD2.Visible = True
            RiskD3.Visible = True
            RiskD4.Visible = True
            RiskD5.Visible = False
            
            RegionD1.Visible = True
            RegionD2.Visible = True
            RegionD3.Visible = True
            RegionD4.Visible = True
            RegionD5.Visible = False
    
        Case 5
            tbGroupD1.Visible = True
            tbGroupD2.Visible = True
            tbGroupD3.Visible = True
            tbGroupD4.Visible = True
            tbGroupD5.Visible = True
            tbGroupD5.SetFocus
            
            RiskD1.Visible = True
            RiskD2.Visible = True
            RiskD3.Visible = True
            RiskD4.Visible = True
            RiskD5.Visible = True
            
            RegionD1.Visible = True
            RegionD2.Visible = True
            RegionD3.Visible = True
            RegionD4.Visible = True
            RegionD5.Visible = True
           
    End Select

End Sub

Private Sub lbNumbers5_Click()

    Select Case lbNumbers5.Value
    
        Case 1
            tbGroupE1.Visible = True
            tbGroupE2.Visible = False
            tbGroupE3.Visible = False
            tbGroupE4.Visible = False
            tbGroupE5.Visible = False
            tbGroupE1.SetFocus
            
            RiskE1.Visible = True
            RiskE2.Visible = False
            RiskE3.Visible = False
            RiskE4.Visible = False
            RiskE5.Visible = False
            
            RegionE1.Visible = True
            RegionE2.Visible = False
            RegionE3.Visible = False
            RegionE4.Visible = False
            RegionE5.Visible = False
    
        Case 2
            tbGroupE1.Visible = True
            tbGroupE2.Visible = True
            tbGroupE3.Visible = False
            tbGroupE4.Visible = False
            tbGroupE5.Visible = False
            tbGroupE2.SetFocus
             
            RiskE1.Visible = True
            RiskE2.Visible = True
            RiskE3.Visible = False
            RiskE4.Visible = False
            RiskE5.Visible = False
            
            RegionE1.Visible = True
            RegionE2.Visible = True
            RegionE3.Visible = False
            RegionE4.Visible = False
            RegionE5.Visible = False
            
        Case 3
            tbGroupE1.Visible = True
            tbGroupE2.Visible = True
            tbGroupE3.Visible = True
            tbGroupE4.Visible = False
            tbGroupE5.Visible = False
            tbGroupE3.SetFocus
    
            RiskE1.Visible = True
            RiskE2.Visible = True
            RiskE3.Visible = True
            RiskE4.Visible = False
            RiskE5.Visible = False
            
            RegionE1.Visible = True
            RegionE2.Visible = True
            RegionE3.Visible = True
            RegionE4.Visible = False
            RegionE5.Visible = False
                    
        Case 4
            tbGroupE1.Visible = True
            tbGroupE2.Visible = True
            tbGroupE3.Visible = True
            tbGroupE4.Visible = True
            tbGroupE5.Visible = False
            tbGroupE4.SetFocus
    
            RiskE1.Visible = True
            RiskE2.Visible = True
            RiskE3.Visible = True
            RiskE4.Visible = True
            RiskE5.Visible = False
    
            
            RegionE1.Visible = True
            RegionE2.Visible = True
            RegionE3.Visible = True
            RegionE4.Visible = True
            RegionE5.Visible = False
        
        Case 5
            tbGroupE1.Visible = True
            tbGroupE2.Visible = True
            tbGroupE3.Visible = True
            tbGroupE4.Visible = True
            tbGroupE5.Visible = True
            tbGroupE5.SetFocus
    
            RiskE1.Visible = True
            RiskE2.Visible = True
            RiskE3.Visible = True
            RiskE4.Visible = True
            RiskE5.Visible = True
            
            RegionE1.Visible = True
            RegionE2.Visible = True
            RegionE3.Visible = True
            RegionE4.Visible = True
            RegionE5.Visible = True
            
    End Select

End Sub

Private Sub Save_Button_Click()

Dim output As String

    '-------------------------------------------------Group A
    If tbGroupA1.Visible = True Then
        output = output & "Group A:" & vbCrLf & "1. " & tbGroupA1.Text & " (" & RiskA1 & ", " & RegionA1 & ")" & vbCrLf
    End If
    
    If tbGroupA2.Visible = True Then
        output = output & "2. " & tbGroupA2.Text & " (" & RiskA2 & ", " & RegionA2 & ")" & vbCrLf
    End If
    
    If tbGroupA3.Visible = True Then
        output = output & "3. " & tbGroupA3.Text & " (" & RiskA3 & ", " & RegionA3 & ")" & vbCrLf
    End If
    
    If tbGroupA4.Visible = True Then
        output = output & "4. " & tbGroupA4.Text & " (" & RiskA4 & ", " & RegionA4 & ")" & vbCrLf
    End If
    
    If tbGroupA5.Visible = True Then
        output = output & "5. " & tbGroupA5.Text & " (" & RiskA5 & ", " & RegionA5 & ")" & vbCrLf
    End If
    
    '-------------------------------------------------Group B
    
    If tbGroupB1.Visible = True Then
        output = output & vbCrLf & "Group B:" & vbCrLf & "1. " & tbGroupB1.Text & " (" & RegionB1 & ", " & RegionB1 & ")" & vbCrLf
    End If
    
    If tbGroupB2.Visible = True Then
        output = output & "2. " & tbGroupB2.Text & " (" & RegionB2 & ", " & RegionB2 & ")" & vbCrLf
    End If
    
    If tbGroupB3.Visible = True Then
        output = output & "3. " & tbGroupB3.Text & " (" & RegionB3 & ", " & RegionB3 & ")" & vbCrLf
    End If
    
    If tbGroupB4.Visible = True Then
        output = output & "4. " & tbGroupB4.Text & " (" & RegionB4 & ", " & RegionB4 & ")" & vbCrLf
    End If
    
    If tbGroupB5.Visible = True Then
        output = output & "5. " & tbGroupB5.Text & " (" & RegionB5 & ", " & RegionB5 & ")" & vbCrLf
    End If
    
    
    '-------------------------------------------------Group C
    
    
    If tbGroupC1.Visible = True Then
        output = output & vbCrLf & "Group C:" & vbCrLf & "1. " & tbGroupC1.Text & " (" & RiskB1 & ", " & RegionC1 & ")" & vbCrLf
    End If
    
    If tbGroupC2.Visible = True Then
        output = output & "2. " & tbGroupC2.Text & " (" & RiskB2 & ", " & RegionC2 & ")" & vbCrLf
    End If
    
    If tbGroupC3.Visible = True Then
        output = output & "3. " & tbGroupC3.Text & " (" & RiskB3 & ", " & RegionC3 & ")" & vbCrLf
    End If
    
    If tbGroupC4.Visible = True Then
        output = output & "4. " & tbGroupC4.Text & " (" & RiskB4 & ", " & RegionC4 & ")" & vbCrLf
    End If
    
    If tbGroupC5.Visible = True Then
        output = output & "5. " & tbGroupC5.Text & " (" & RiskB5 & ", " & RegionC5 & ")" & vbCrLf
    End If
    
    
    '-------------------------------------------------Group D
    
    If tbGroupD1.Visible = True Then
        output = output & vbCrLf & "Group D:" & vbCrLf & "1. " & tbGroupD1.Text & " (" & RiskD1 & ", " & RegionD1 & ")" & vbCrLf
    End If
    
    If tbGroupD2.Visible = True Then
        output = output & "2. " & tbGroupD2.Text & " (" & RiskD2 & ", " & RegionD2 & ")" & vbCrLf
    End If
    
    If tbGroupD3.Visible = True Then
        output = output & "3. " & tbGroupD3.Text & " (" & RiskD3 & ", " & RegionD3 & ")" & vbCrLf
    End If
    
    If tbGroupD4.Visible = True Then
        output = output & "4. " & tbGroupD4.Text & " (" & RiskD4 & ", " & RegionD4 & ")" & vbCrLf
    End If
    
    If tbGroupD5.Visible = True Then
        output = output & "5. " & tbGroupD5.Text & " (" & RiskD5 & ", " & RegionD5 & ")" & vbCrLf
    End If
    
    
    '-------------------------------------------------Group E
    
    If tbGroupE1.Visible = True Then
        output = output & vbCrLf & "Group E:" & vbCrLf & "1. " & tbGroupE1.Text & " (" & RiskE1 & ", " & RegionE1 & ")" & vbCrLf
    End If
    
    If tbGroupE2.Visible = True Then
        output = output & "2. " & tbGroupE2.Text & " (" & RiskE2 & ", " & RegionE2 & ")" & vbCrLf
    End If
    
    If tbGroupE3.Visible = True Then
        output = output & "3. " & tbGroupE3.Text & " (" & RiskE3 & ", " & RegionE3 & ")" & vbCrLf
    End If
    
    If tbGroupE4.Visible = True Then
        output = output & "4. " & tbGroupE4.Text & " (" & RiskE4 & ", " & RegionE4 & ")" & vbCrLf
    End If
    
    If tbGroupE5.Visible = True Then
        output = output & "5. " & tbGroupE5.Text & " (" & RiskE5 & ", " & RegionE5 & ")" & vbCrLf
    End If

        Dim A As String
        A = wsData.Range("O1")
        Range(A) = output
        wsData.Range("O1").Clear
        
    Unload Me

End Sub
Private Sub Cancel_Button_Click()
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
