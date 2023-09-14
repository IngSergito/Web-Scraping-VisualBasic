Function Convertirduracion(duration As String) As Double
    Dim texto As String
    Dim numeros() As String
    
    ' Definir el texto a analizar
    'duration = "0 hours 1 minutes 14 seconds"
    
    ' Dividir el texto en un arreglo de cadenas
    numeros = Split(duration, " ")
    
    ' Sumar los números
    Convertirduracion = CDbl(numeros(0)) / 24 + CDbl(numeros(2)) / 1440 + CDbl(numeros(4)) / 86400
    
End Function

Sub CrearEstilos(documentx As Word.Document)
    documentx.Styles(wdStyleNormal).Font.Name = ("Arial")
    For i = 1 To 3
        'Verificar si el estilo ya existe antes de agregarlo
        Dim newStyle As Object
        On Error Resume Next
        Set newStyle = documentx.Styles("Título " & i)
        On Error GoTo 0
        
        'Agregar un nuevo estilo de título solo si no existe previamente
        If newStyle Is Nothing Then
            Set newStyle = documentx.Styles.Add(Name:=("Título " & CStr(i)), Type:=wdStyleTypeParagraph)
         End If
         
        'Establecer las propiedades del estilo
        With newStyle
            .Font.Name = "Arial"
            .Font.Size = 24 - 4 * (i - 1)
            .Font.Bold = False
            .Font.Underline = wdUnderlineSingle
            If i = 1 Then
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
            Else
                .ParagraphFormat.Alignment = wdAlignParagraphLeft
            End If
        End With
    Next i
End Sub
 
Sub Alarmasmovil()
    'Configurar ventana de selección de archivo
    Dim ventana As FileDialog
    Set ventana = Application.FileDialog(msoFileDialogFilePicker)
    ventana.Title = "Seleccione alarmas de la red móvil"
    ventana.InitialFileName = ThisWorkbook.Path
    ventana.ButtonName = "Analizar"
    
    'Abrir ventana de selección de archivo
    If ventana.Show Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Dim archivo_alarmas As Workbook
        Set archivo_alarmas = Workbooks.Open(ventana.SelectedItems(1))
        Dim rango_alarmas As Range
        Set rango_alarmas = archivo_alarmas.Sheets(1).Cells
        Dim hoja_tablas As Worksheet
        Set hoja_tablas = ThisWorkbook.Sheets("Tablas de datos")
        
        hoja_tablas.Activate
        
        'Encontrar filas y columnas
        Dim Fila As Long, fila_ultima As Long, severidadCol As Integer, nameCol As Integer, duraCol As Integer, equipCol As Integer
                
        With rango_alarmas
            'Quitar valores a la fila anterior a la tabla
            On Error Resume Next
            .Find("Severity", LookAt:=xlWhole).Offset(-1, 0).EntireRow.ClearContents
            On Error GoTo 0
            
            Fila = .Find("Alarm ID", LookAt:=xlWhole).Row + 1
            severidadCol = .Find("Severity", LookAt:=xlWhole).Column
            nameCol = .Find("Name", LookAt:=xlWhole).Column
            duraCol = .Find("Alarm Duration", LookAt:=xlWhole).Column
            equipCol = .Find("Alarm Source", LookAt:=xlWhole).Column
            fila_ultima = .Find("Severity", LookAt:=xlWhole).CurrentRegion.Rows.Count + Fila - 2
        End With
        
        'Borrar datos anteriores
        Dim Objeto As ListObject
        Dim fila_tabla As ListRow
        
        For Each Objeto In ThisWorkbook.Sheets("Tablas de datos").ListObjects
            Debug.Print (Objeto.Comment)
            If Objeto.Name = "Severidad" Then
                hoja_tablas.Range("B3:B6").Value = 0
            Else
                For i = Objeto.ListRows.Count To 1 Step -1
                    Objeto.ListRows(i).Delete
                Next i
            End If
        Next Objeto
            
        'Analizar las alarmas
        Dim proceso As Integer
        Dim ralarma_encontrada As Range
        Dim fila_nueva As ListRow
        Dim equipo_existente As Boolean 'Para manejar las últimas tablas
        Dim alarma_existente As Boolean 'Para manejar las últimas tablas
        Dim posicion As Integer
        
        For i = Fila To fila_ultima
            proceso = (i - Fila) * 100 / fila_ultima
            Application.StatusBar = "Analizando - " & proceso & "%"
            
            'Actualizar tabla general
            hoja_tablas.Cells.Find(rango_alarmas.Cells(i, severidadCol)).Offset(0, 1).Value = hoja_tablas.Cells.Find(rango_alarmas.Cells(i, severidadCol)).Offset(0, 1).Value + 1
            
            'Actualizar tablas de resumen por severidad
            Set ralarma_encontrada = hoja_tablas.ListObjects("Alarms" & rango_alarmas.Cells(i, severidadCol).Value).Range.Find(rango_alarmas.Cells(i, nameCol), LookAt:=xlWhole)
            
            If ralarma_encontrada Is Nothing Then
                Set fila_nueva = hoja_tablas.ListObjects("Alarms" & rango_alarmas.Cells(i, severidadCol).Value).ListRows.Add(1, True)
                fila_nueva.Range.Columns(1).Value = rango_alarmas.Cells(i, nameCol).Value
                fila_nueva.Range.Columns(2).Value = 1
                fila_nueva.Range.Columns(3).Value = Convertirduracion(rango_alarmas.Cells(i, duraCol).Value)
                fila_nueva.Range.Columns(3).NumberFormat = "[h]:mm:ss"
            Else
                ralarma_encontrada.Offset(0, 1).Value = ralarma_encontrada.Offset(0, 1).Value + 1
                ralarma_encontrada.Offset(0, 2).Value = ralarma_encontrada.Offset(0, 2).Value + Convertirduracion(rango_alarmas.Cells(i, duraCol).Value)
            End If
            
            'Actualizar tablas de equipo por severidad
            equipo_existente = False
            alarma_existente = False
            Dim numerotabla As Long
            numerotabla = hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows.Count
            'Algoritmo para insertar alfabéticamente las alarmas
            
            For j = 1 To hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows.Count
                ' Si se encuentra la primera posición en la que debe ir alfabéticamente
                If rango_alarmas.Cells(i, nameCol).Value < hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows(j).Range.Columns(1).Value And Left(hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows(j).Range.Columns(1).Value, 3) <> "   " Then
                    'Se agrega la alarma
                    Set fila_nueva = hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows.Add(j, True)
                    fila_nueva.Range.Font.Bold = True
                    fila_nueva.Range.Columns(1).Value = rango_alarmas.Cells(i, nameCol).Value
                    fila_nueva.Range.Columns(2).Value = 1
                    fila_nueva.Range.Columns(3).Value = Convertirduracion(rango_alarmas.Cells(i, duraCol).Value)
                    fila_nueva.Range.Columns(3).NumberFormat = "[h]:mm:ss"
                    If Not (ThisWorkbook.Sheets("Correlación alarma-observación").ListObjects("Correlación").Range.Find(rango_alarmas.Cells(i, nameCol).Value, LookAt:=xlWhole) Is Nothing) Then
                        fila_nueva.Range.Columns(4).Value = ThisWorkbook.Sheets("Correlación alarma-observación").ListObjects("Correlación").Range.Find(rango_alarmas.Cells(i, nameCol).Value, LookAt:=xlWhole).Offset(0, 1)
                    End If
                    alarma_existente = True
                    'Se agrega el equipo
                    Set fila_nueva = hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows.Add(j + 1, True)
                    fila_nueva.Range.Font.Bold = False
                    fila_nueva.Range.Columns(1).Value = "   " & rango_alarmas.Cells(i, equipCol).Value
                    fila_nueva.Range.Columns(2).Value = 1
                    fila_nueva.Range.Columns(3).Value = Convertirduracion(rango_alarmas.Cells(i, duraCol).Value)
                    fila_nueva.Range.Columns(3).NumberFormat = "[h]:mm:ss"
                    equipo_existente = True
                    'Se sale del for
                    Exit For
                ' Si se encuentra la alarma en la tabla
                ElseIf rango_alarmas.Cells(i, nameCol).Value = hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows(j).Range.Columns(1).Value Then
                    hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows(j).Range.Columns(2).Value = hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows(j).Range.Columns(2).Value + 1
                    hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows(j).Range.Columns(3).Value = hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows(j).Range.Columns(3).Value + Convertirduracion(rango_alarmas.Cells(i, duraCol).Value)
                    ' Se buscará el equipo luego de esta alarma hasta la siguiente alarma
                    alarma_existente = True
                    posicion = j + 1
                    
                    'Bucle de búsqueda de equipo
                    Do While Left(hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows(posicion).Range.Columns(1).Value, 3) = "   "
                        'Se encontró el equipo
                        If hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows(posicion).Range.Columns(1).Value = "   " & rango_alarmas.Cells(i, equipCol).Value Then
                            equipo_existente = True
                            hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows(posicion).Range.Columns(2).Value = hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows(posicion).Range.Columns(2).Value + 1
                            hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows(posicion).Range.Columns(3).Value = hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows(posicion).Range.Columns(3).Value + Convertirduracion(rango_alarmas.Cells(i, duraCol).Value)
                            Exit Do
                        'Se encontró la posición en la que se ubicará alfbéticamente
                        ElseIf hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows(posicion).Range.Columns(1).Value > "   " & rango_alarmas.Cells(i, equipCol).Value Then
                            equipo_existente = True
                            Set fila_nueva = hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows.Add(posicion + 1, True)
                            fila_nueva.Range.Font.Bold = False
                            fila_nueva.Range.Columns(1).Value = "   " & rango_alarmas.Cells(i, equipCol).Value
                            fila_nueva.Range.Columns(2).Value = 1
                            fila_nueva.Range.Columns(3).Value = Convertirduracion(rango_alarmas.Cells(i, duraCol).Value)
                            fila_nueva.Range.Columns(3).NumberFormat = "[h]:mm:ss"
                            Exit Do
                        End If
                        posicion = posicion + 1
                        If posicion > hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows.Count Then
                            Exit Do
                        End If
                    Loop
                    Exit For
                End If
            Next j
            
            ' Si no se encuentra nada
            If alarma_existente = False Then
                Set fila_nueva = hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows.Add(j, True)
                fila_nueva.Range.Font.Bold = True
                fila_nueva.Range.Columns(1).Value = rango_alarmas.Cells(i, nameCol).Value
                fila_nueva.Range.Columns(2).Value = 1
                fila_nueva.Range.Columns(3).Value = Convertirduracion(rango_alarmas.Cells(i, duraCol).Value)
                fila_nueva.Range.Columns(3).NumberFormat = "[h]:mm:ss"
                If Not (ThisWorkbook.Sheets("Correlaci�n alarma-observaci�n").ListObjects("Correlaci�n").Range.Find(rango_alarmas.Cells(i, nameCol).Value, LookAt:=xlWhole) Is Nothing) Then
                    fila_nueva.Range.Columns(4).Value = ThisWorkbook.Sheets("Correlaci�n alarma-observaci�n").ListObjects("Correlaci�n").Range.Find(rango_alarmas.Cells(i, nameCol).Value, LookAt:=xlWhole).Offset(0, 1).Value
                End If
                Set fila_nueva = hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows.Add(j + 1, True)
                fila_nueva.Range.Font.Bold = False
                fila_nueva.Range.Columns(1).Value = "   " & rango_alarmas.Cells(i, equipCol).Value
                fila_nueva.Range.Columns(2).Value = 1
                fila_nueva.Range.Columns(3).Value = Convertirduracion(rango_alarmas.Cells(i, duraCol).Value)
                fila_nueva.Range.Columns(3).NumberFormat = "[h]:mm:ss"
            ElseIf equipo_existente = False Then
                Set fila_nueva = hoja_tablas.ListObjects("AlarmEquip" & rango_alarmas.Cells(i, severidadCol).Value).ListRows.Add(posicion, True)
                fila_nueva.Range.Font.Bold = False
                fila_nueva.Range.Columns(1).Value = "   " & rango_alarmas.Cells(i, equipCol).Value
                fila_nueva.Range.Columns(2).Value = 1
                fila_nueva.Range.Columns(3).Value = Convertirduracion(rango_alarmas.Cells(i, duraCol).Value)
                fila_nueva.Range.Columns(3).NumberFormat = "[h]:mm:ss"
            End If
        Next i
        
        For Each Objeto In ThisWorkbook.Sheets("Tablas de datos").ListObjects
            If Objeto.Name <> "Severidad" Then
                Objeto.Range.EntireColumn.AutoFit
            End If
        Next Objeto
        Application.StatusBar = False
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
    Else
        Exit Sub
    End If
    
    archivo_alarmas.Close (False)
End Sub

Sub GenerarWord()
    Dim Documento As Object
    Dim objWord As Object
    
    ' Crea un objeto que representa la aplicaci�n de Word
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = True
    ' Crea un nuevo documento
    Set Documento = objWord.Documents.Add
    
    Call CrearEstilos(Documento)
    
    objWord.Selection.Font.Name = "Arial"
    ' Cambiar las propiedades del documento
    With Documento
        .PageSetup.Orientation = wdOrientLandscape
        .BuiltinDocumentProperties("Title").Value = "Diagn�stico Nivel 2"
        .BuiltinDocumentProperties("Author").Value = "Equipos de Acceso Inal�mbrico"
        .BuiltinDocumentProperties("Subject").Value = "Departamento de Acceso"
        .BuiltinDocumentProperties("Company").Value = "ETECSA"
        .BuiltinDocumentProperties("Comments").Value = "Realizado con VBA, por Ing. Sergio Rosales Mojena para ETECSA (contacte +53 54206843; rosalessergioc4t@gmail.com)"
    End With
    
    ' Insertar el t�tulo, asunto y autor en la primera p�gina
    With objWord.Selection
        ' Insertar el t�tulo
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Font.Color = 6575178
        .Font.Size = 36
        .Font.Bold = True
        .TypeText (Documento.BuiltinDocumentProperties("Title").Value)
        .TypeParagraph
        
        ' Insertar el asunto
        .Font.Color = 13998939
        .Font.Size = 20
        .Font.Bold = True
        .TypeText (Documento.BuiltinDocumentProperties("Subject").Value)
        .TypeParagraph
        
        
        ' Insertar el autor
        .Font.Color = wdColorGray50
        .Font.Size = 16
        .Font.Bold = True
        .TypeText (Documento.BuiltinDocumentProperties("Author").Value)
        .TypeParagraph
        
        ' Insertar imagen
        ThisWorkbook.Sheets("Tablas de datos").Shapes.Range(Array("ImagenP")).Select
        Selection.Copy
        .Paste
        .TypeParagraph
        
        ' Insertar Frecuencia
        .TypeText ("Frecuencia: Mensual")
        .InsertBreak Type:=wdPageBreak

        'Insertar informaci�n
        .ParagraphFormat.SpaceAfter = 0
        .ParagraphFormat.SpaceBefore = 0
        .Style = Documento.Styles("T�tulo 1")
        .TypeText ("Cap�tulo 1 NE: BSC, RNC, BTS, NodosB y eNodosB")
        .TypeParagraph
        .Style = Documento.Styles("T�tulo 2")
        .TypeText ("1.1 Desempe�o (an�lisis de alarmas)")
        .TypeParagraph
        .Font.Bold = False
        .Font.Size = 11
        .Font.Underline = wdUnderlineNone
        .TypeText ("Resumen mensual del an�lisis del desempe�o semanal y acciones acometidas.")
        .TypeParagraph
        .Style = Documento.Styles("T�tulo 3")
        .TypeText ("1.1.2 Huawei")
        .TypeParagraph
        
    End With
    
    ' Copiar tablas
    Dim Tabla As ListObject
    
    For Each Tabla In ThisWorkbook.Sheets("Tablas de datos").ListObjects
        
        objWord.Selection.Font.Bold = True
        objWord.Selection.TypeText (Tabla.Comment)
        objWord.Selection.TypeParagraph
        Tabla.Range.Select
        objWord.Selection.ParagraphFormat.SpaceAfter = 0
        Selection.Copy
        objWord.Selection.PasteExcelTable False, False, False
    Next Tabla
    
    Application.CutCopyMode = False
    ThisWorkbook.Sheets("Tablas de datos").Range("A1").Select
    
    ' Preparar resto del documento
    With objWord.Selection
        .InsertBreak Type:=wdPageBreak
        .Style = Documento.Styles("T�tulo 2")
        .TypeText ("1.2 Tickets de Fallas ocurridas en el per�odo")
        .TypeParagraph
        .TypeParagraph
        .Style = Documento.Styles("T�tulo 3")
        .TypeText ("1.2.2 Huawei")
        .TypeParagraph
        .TypeParagraph
        .Font.Name = "Arial"
        .TypeText ("Al cierre del mes de __________ en estado pendiente un total de __ tickets, ___________ que el mes anterior, ")
        .TypeText ("entre ellos la mayor causa se encuentra banco de bater�as defectuoso, tec cooler, licencia, deterioro de la calidad del servicio y de fan.")
        .TypeParagraph
        .TypeParagraph
        .TypeText ("insertar tabla de tickets")
        .InsertBreak Type:=wdPageBreak
        .Style = Documento.Styles("T�tulo 2")
        .TypeText ("1.4 Plan de medidas")
    End With
    
    'Recorrer todas las tablas del documento
    Dim tbl As Table
    For Each tbl In Documento.Tables
        'Establecer el espaciado inferior de la tabla en 0 puntos
        tbl.Select
        With objWord.Selection.ParagraphFormat
            'Establecer el espaciado
            .SpaceAfter = 0
        End With
    Next tbl
    
    ' Libera la memoria asignada
    Set Documento = Nothing
    Set objWord = Nothing
    
    MsgBox "El documento de Word ha sido creado correctamente."

End Sub
