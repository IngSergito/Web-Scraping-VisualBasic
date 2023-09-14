VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Reporte de Tickets"
   ClientHeight    =   4740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8940.001
   OleObjectBlob   =   "FormularioPrincipal.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public calendarioNombre As String, labelNombre As String

Private Sub CommandButton1_Click()
    Me.CheckBox1.Value = True
    Me.CheckBox2.Value = True
    Me.CheckBox3.Value = True
    Me.CheckBox4.Value = True
    Me.CheckBox5.Value = True
    Me.CheckBox6.Value = True
    Me.CheckBox0.Value = True
    
End Sub

Private Sub CommandButton2_Click()
    Me.CheckBox0.Value = False
    Me.CheckBox1.Value = False
    Me.CheckBox2.Value = False
    Me.CheckBox3.Value = False
    Me.CheckBox4.Value = False
    Me.CheckBox5.Value = False
    Me.CheckBox6.Value = False
End Sub

Private Sub RefEdit1_BeforeDragOver(Cancel As Boolean, ByVal Data As MSForms.DataObject, ByVal X As stdole.OLE_XPOS_CONTAINER, ByVal Y As stdole.OLE_YPOS_CONTAINER, ByVal DragState As MSForms.fmDragState, Effect As MSForms.fmDropEffect, ByVal Shift As Integer)

End Sub

Private Sub CommandButton3_Click()
    calendarioNombre = "Fecha Inicio Desde"
    labelNombre = "TextBoxFID"
    UserForm2.Show
End Sub


Private Sub CommandButton4_Click()
    calendarioNombre = "Fecha Inicio Hasta"
    labelNombre = "TextBoxFIH"
    UserForm2.Show
End Sub

Private Sub Inspeccion_Click()
    Dim i As Integer, estadosCodigoSi As String, estadosCodigoNo As String, estados As String, fechaID As String, fechaIH As String
    estadosCodigoSi = ""
    estadosCodigoNo = ""

    For i = 0 To 6
        If Me.Frame1.Controls(i).Value Then
            estadosCodigoSi = estadosCodigoSi & "&ticketSearch%3AestadoField%3Aestado=" & CStr(i) & "%3A" & CStr(i + 115)
        Else
            estadosCodigoNo = estadosCodigoNo & "&ticketSearch%3AestadoField%3Aestado=" & CStr(i) & "%3A" & CStr(i + 115)
        End If
    Next i
    
    If estadosCodigoSi = "" Or estadosCodigoNo = "" Then
        If estadosCodigoNo = "" Then
            estados = estadosCodigoSi
        Else
            estados = estadosCodigoNo
        End If
        estados = estados & "&ticketSearch%3AestadoField%3Aestado=%3A"
    Else
        estados = estadosCodigoNo & "&ticketSearch%3AestadoField%3Aestado=%3A" & estadosCodigoSi
    End If
    
    fechaID = Replace(Replace(Replace(TextBoxFID.Text, "/", "%2F"), ":", "%3A"), " ", "%20")
    fechaIH = Replace(Replace(Replace(TextBoxFIH.Text, "/", "%2F"), ":", "%3A"), " ", "%20")
    
    Call Webtickets(userTextBox.Text, passTextBox.Text, estados, fechaID, fechaIH, Hoja3.ListObjects(1))
    
    Unload Me
End Sub

