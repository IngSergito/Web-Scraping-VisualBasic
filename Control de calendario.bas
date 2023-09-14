Attribute VB_Name = "Módulo2"
Sub Control_Calendario(fecha As Date)
    Dim primerDia As Integer, diaMesAnterior As Integer, diasMesActual As Integer, i As Integer
    primerDia = Weekday(DateSerial(Year(fecha), Month(fecha), 1), vbMonday)
    diaMesAnterior = Day(DateSerial(Year(fecha), Month(fecha), 1) - 1)

    If Month(fecha) = 12 Then
        diasMesActual = Day(DateSerial(Year(fecha) + 1, 1, 1) - 1)
    Else
        diasMesActual = Day(DateSerial(Year(fecha), Month(fecha) + 1, 1) - 1)
    End If
    
    With UserForm2.FrameCal
        For i = primerDia - 1 To 0 Step -1
            .Controls("d" & i).Caption = diaMesAnterior
            diaMesAnterior = diaMesAnterior - 1
        Next i

        For i = primerDia To 34
            If (i + 1 - primerDia) <= diasMesActual Then
                .Controls("d" & i).Caption = i + 1 - primerDia
            Else
                .Controls("d" & i).Caption = i + 1 - primerDia - diasMesActual
            End If
        Next i
        
        .Caption = "Calendario " & MonthName(Month(fecha)) & " " & Year(fecha)
    End With
    UserForm2.Mes = Month(fecha)
End Sub

Sub Limpiar_Marca()
    Dim i As Integer
    For i = 0 To 34
        UserForm2.Controls("d" & i).borderStyle = fmBorderStyleNone
    Next i
End Sub
