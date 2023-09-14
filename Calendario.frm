VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   4275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3600
   OleObjectBlob   =   "Calendario.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fecha As String, etiqueta As String
Public Mes As Integer

Private Sub CommandButton1_Click()
    If Me.TextBox1.Text = "AM" Then
        Me.TextBox1.Text = "PM"
    Else
        Me.TextBox1.Text = "AM"
    End If
End Sub

Private Sub CommandButton2_Click()
    Mes = Mes + 1
    Call Control_Calendario(DateSerial(LabelYear.Caption, Mes, 1))
End Sub

Private Sub CommandButton3_Click()
    Mes = Mes - 1
    Call Control_Calendario(DateSerial(LabelYear.Caption, Mes, 1))
End Sub


Private Sub CommandButton4_Click()
    Dim i As Integer
    For i = 0 To 34
        If (Me.Controls("d" & i).borderStyle = fmBorderStyleSingle) Then
            UserForm1.Controls(UserForm1.labelNombre).Text = Format(Me.Controls("d" & i).Caption, "00") & "/" & Format(Mes, "00") & "/" & LabelYear.Caption & " " & TextBoxHo.Value & ":" & TextBoxMin.Value & ":" & TextBoxSeg.Value & " " & TextBox1.Text
            Unload Me
        End If
    Next i
End Sub

Private Sub d0_Click()
    Call Limpiar_Marca
    d0.borderStyle = fmBorderStyleSingle
End Sub

Private Sub d1_Click()
    Call Limpiar_Marca
    d1.borderStyle = fmBorderStyleSingle
End Sub

Private Sub d2_Click()
    Call Limpiar_Marca
    d2.borderStyle = fmBorderStyleSingle
End Sub

Private Sub d3_Click()
    Call Limpiar_Marca
    d3.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d4_Click()
    Call Limpiar_Marca
    d4.borderStyle = fmBorderStyleSingle
End Sub

Private Sub d5_Click()
    Call Limpiar_Marca
    d5.borderStyle = fmBorderStyleSingle
End Sub

Private Sub d6_Click()
    Call Limpiar_Marca
    d6.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d7_Click()
    Call Limpiar_Marca
    d7.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d8_Click()
    Call Limpiar_Marca
    d8.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d9_Click()
    Call Limpiar_Marca
    d9.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d10_Click()
    Call Limpiar_Marca
    d10.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d11_Click()
    Call Limpiar_Marca
    d11.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d12_Click()
    Call Limpiar_Marca
    d12.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d13_Click()
    Call Limpiar_Marca
    d13.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d14_Click()
    Call Limpiar_Marca
    d14.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d15_Click()
    Call Limpiar_Marca
    d15.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d16_Click()
    Call Limpiar_Marca
    d16.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d17_Click()
    Call Limpiar_Marca
    d17.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d18_Click()
    Call Limpiar_Marca
    d18.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d19_Click()
    Call Limpiar_Marca
    d19.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d20_Click()
    Call Limpiar_Marca
    d20.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d21_Click()
    Call Limpiar_Marca
    d21.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d22_Click()
    Call Limpiar_Marca
    d22.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d23_Click()
    Call Limpiar_Marca
    d23.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d24_Click()
    Call Limpiar_Marca
    d24.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d25_Click()
    Call Limpiar_Marca
    d25.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d26_Click()
    Call Limpiar_Marca
    d26.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d27_Click()
    Call Limpiar_Marca
    d27.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d28_Click()
    Call Limpiar_Marca
    d28.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d29_Click()
    Call Limpiar_Marca
    d29.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d30_Click()
    Call Limpiar_Marca
    d30.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d31_Click()
    Call Limpiar_Marca
    d31.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d32_Click()
    Call Limpiar_Marca
    d32.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d33_Click()
    Call Limpiar_Marca
    d33.borderStyle = fmBorderStyleSingle
End Sub
Private Sub d34_Click()
    Call Limpiar_Marca
    d34.borderStyle = fmBorderStyleSingle
End Sub

Private Sub FrameCal_Click()

End Sub

Private Sub SpinButton2_SpinDown()
    Me.LabelYear.Caption = Me.LabelYear.Caption - 1
    Call Control_Calendario(DateSerial(LabelYear.Caption, Mes, 1))
End Sub

Private Sub SpinButton2_SpinUp()
    Me.LabelYear.Caption = Me.LabelYear.Caption + 1
    Call Control_Calendario(DateSerial(LabelYear.Caption, Mes, 1))
End Sub

Private Sub SpinButtonHor_SpinDown()
    If Me.TextBoxHo.Value > 0 Then
        If Me.TextBoxHo.Value = 12 Then
            Call CommandButton1_Click
        End If
        Me.TextBoxHo.Value = Format(Me.TextBoxHo.Value - 1, "00")
    Else
        Me.TextBoxHo.Value = 12
    End If
End Sub

Private Sub SpinButtonHor_SpinUp()
    If Me.TextBoxHo.Value < 12 Then
        If Me.TextBoxHo.Value = 11 Then
            Call CommandButton1_Click
        End If
        Me.TextBoxHo.Value = Format(Me.TextBoxHo.Value + 1, "00")
    Else
        Me.TextBoxHo.Value = Format(0, "00")
    End If
End Sub

Private Sub SpinButtonMin_SpinDown()
    If Me.TextBoxMin.Value > 0 Then
        Me.TextBoxMin.Value = Format(Me.TextBoxMin.Value - 1, "00")
    Else
        Me.TextBoxMin.Value = 59
        Call SpinButtonHor_SpinDown
    End If
End Sub

Private Sub SpinButtonMin_SpinUp()
    If Me.TextBoxMin.Value < 59 Then
        Me.TextBoxMin.Value = Format(Me.TextBoxMin.Value + 1, "00")
    Else
        Me.TextBoxMin.Value = Format(0, "00")
        Call SpinButtonHor_SpinUp
    End If
End Sub

Private Sub SpinButtonSeg_SpinDown()
    If Me.TextBoxSeg.Value > 0 Then
        Me.TextBoxSeg.Value = Format(Me.TextBoxSeg.Value - 1, "00")
    Else
        Me.TextBoxSeg.Value = 59
        Call SpinButtonMin_SpinDown
    End If
End Sub

Private Sub SpinButtonSeg_SpinUp()
    If Me.TextBoxSeg.Value < 59 Then
        Me.TextBoxSeg.Value = Format(Me.TextBoxSeg.Value + 1, "00")
    Else
        Me.TextBoxSeg.Value = Format(0, "00")
        Call SpinButtonMin_SpinUp
    End If
End Sub


Private Sub UserForm_Initialize()
    Me.Caption = UserForm1.calendarioNombre
    etiqueta = UserForm1.labelNombre
    Call Control_Calendario(Now)
    LabelYear.Caption = Year(Now)
End Sub
