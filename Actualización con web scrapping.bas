Sub Actualizar_Libro(Libro As Workbook)
    Dim hoja As Worksheet, tablaDin As PivotTable
    'Actualizar todas las tablas din�micas del libro
    For Each hoja In Libro.Worksheets
        For Each tablaDin In hoja.PivotTables
            tablaDin.RefreshTable
        Next tablaDin
    Next hoja
End Sub

Function HTTPPayload(headers() As Object) As String
    Dim header As Object, requestPayload As String
    
    requestPayload = headers(1).Name & "=" & headers(1).Value
    
    For i = 2 To headers.Count
        requestPayload = "&" & header.Name & "=" & header.Value
    Next header
    
    HTTPPayload = requestPayload
End Function


Sub Actualizar_tabla_tickets(rutaArchivoTickets As String, tablaTickets As ListObject)
    Dim libroTickets As Workbook, hoja As Worksheet, nuevosTickets As ListObject, columna As ListColumn
    Dim ticketsEncontrados As Boolean, encabezado As Range
    Application.DisplayAlerts = False
    'Abrir nuevo libro
    Set libroTickets = Workbooks.Open(rutaArchivoTickets)
    'Buscar tabla de tickets
    'Recorrer todas las hojas
    For Each hoja In libroTickets.Worksheets
        ticketsEncontrados = False
        'Buscar el encabezado de la primera columna en la hoja
        Set encabezado = hoja.Cells.Find(What:=tablaTickets.ListColumns(1).Name, LookAt:=xlWhole)
        
        'Si se encuentra el encabezado en la hoja y la cantidad de columnas de su regi�n es mayor que 1
        If Not (encabezado Is Nothing) And encabezado.CurrentRegion.Columns.Count > 1 Then
            For Each columna In tablaTickets.ListColumns
                If columna.Name <> encabezado.Value Then
                
                    Exit For
                End If
                If tablaTickets.ListColumns(tablaTickets.ListColumns.Count) = columna Then
                    ticketsEncontrados = True
                    Exit For
                Else
                    Set encabezado = encabezado.Offset(0, 1)
                End If
            Next columna
            
                Exit For
        Else
        End If
    Next hoja
    
    If Not (ticketsEncontrados) Then
        MsgBox "No se un listado de tickets compatible al formato registrado"
    Else
        Application.ScreenUpdating = False
        Set nuevosTickets = encabezado.Parent.ListObjects.Add(xlSrcRange, encabezado.CurrentRegion, , xlYes)
        Dim cant As Integer
        cant = tablaTickets.ListRows.Count
        For i = 1 To cant
            tablaTickets.ListRows(1).Delete
        Next i
        For i = 1 To nuevosTickets.ListRows.Count
            tablaTickets.ListRows.Add (i)
            nuevosTickets.ListRows(i).Range.Copy
            tablaTickets.ListRows(i).Range.PasteSpecial (xlPasteValues)
        Next i
        Call Actualizar_Libro(ThisWorkbook)
    End If
    libroTickets.Close (False)
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub Webtickets(urlWeb As String, urlListadoT As String, user As String, password As String, estados As String, fechaInicioDesde As String, fechaInicioHasta As String, listaTickets As ListObject)

    Dim XMLPage As New MSXML2.XMLHTTP60
    Dim htmldoc As New MSHTML.HTMLDocument
    Dim htmlim As MSHTML.IHTMLElement
    Dim htmlims As MSHTML.IHTMLElementCollection
    Dim URL As String, login As String
    Dim fileStream As Object
    Dim Username As String, buscar As String, mesInicioDesde As String, mesInicioHasta As String, fechaInicioDesdeCodif As String, fechaInicioHastaCodif As String
    Dim excelTicket As Workbook, cid As String
    Dim ruta As String
    
    
    ' Rellenar campos redundantes de petición HTTP
    mesInicioDesde = Mid(fechaInicioDesde, 6, 9)
    mesInicioHasta = Mid(fechaInicioHasta, 6, 9)
    fechaInicioDesdeCodif = Replace(fechaInicioDesde, "%20", "+")
    fechaInicioHastaCodif = Replace(fechaInicioHasta, "%20", "+")
    
    'Ir a la página de autenticación
    URL = urlWeb
    XMLPage.Open "Get", URL, False
    XMLPage.send
    
    ' Crea un nuevo objeto HTML para analizar el contenido de la página de inicio de sesión
    Set htmldoc = CreateObject("HTMLFile")
    htmldoc.body.innerHTML = XMLPage.responseText
    
    ' Obtén los elementos del formulario de inicio de sesión
    Dim java As Object
    
    Set java = htmldoc.getElementById("javax.faces.ViewState")
    ' Rellena los campos del formulario con tus credenciales
    login = "loginForm=loginForm&loginForm%3Aj_id22%3Ausername=" & user & "&loginForm%3Aj_id36%3Apassword=" & password & "&loginForm%3Asubmit=Entrar&javax.faces.ViewState=" & java.Value
    ' Envía el formulario de inicio de sesión
    With XMLPage
        .Open "POST", URL, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send login
    End With
    
    ' Ver listado de tickets
    With XMLPage
        .Open "GET", urlListadoT, False
        .send
    End With
        
    ' Guardar elementos de la página de identificación de vista
    htmldoc.body.innerHTML = XMLPage.responseText
    Set htmlim = htmldoc.getElementById("j_id20:perfilField")
    
    cid = Mid(htmlim.href, InStr(1, htmlim.href, "cid=", vbTextCompare) + 4, 5)
    Set java = htmldoc.getElementById("javax.faces.ViewState")
    
    buscar = "AJAXREQUEST=_viewRoot&ticketSearch=ticketSearch&ticketSearch%3Aj_id31=true&ticketSearch%3Aj_id66%3Asuggestionnumero_selection=&ticketSearch%3Aj_id66%3Anumero="
    buscar = buscar & "&ticketSearch%3Aj_id82%3Asuggestionnumero_selection=&ticketSearch%3Aj_id82%3Atitulo="
    buscar = buscar & "&ticketSearch%3AfechaInicioDesdeField%3AfechaInicioDesdeInputDate=" & fechaInicioDesde & "&ticketSearch%3AfechaInicioDesdeField%3AfechaInicioDesdeInputCurrentDate=" & mesInicioDesde & "&ticketSearch%3AfechaInicioHastaField%3AfechaInicioHastaInputDate=" & fechaInicioHasta & "&ticketSearch%3AfechaInicioHastaField%3AfechaInicioHastaInputCurrentDate=" & mesInicioHasta & "&ticketSearch%3AfechaFinDesdeField%3AfechaFinDesdeInputDate=&ticketSearch%3AfechaFinDesdeField%3AfechaFinDesdeInputCurrentDate=08%2F2023&ticketSearch%3AfechaFinHastaField%3AfechaFinHastaInputDate=&ticketSearch%3AfechaFinHastaField%3AfechaFinHastaInputCurrentDate=08%2F2023"
    buscar = buscar & "&ticketSearch%3AtipoProblemaField%3AtipoProblema=org.jboss.seam.ui.NoSelectionConverter.noSelectionValue"
    buscar = buscar & estados
    buscar = buscar & "&ticketSearch%3AespecialidadField%3Aespecialidad=org.jboss.seam.ui.NoSelectionConverter.noSelectionValue&ticketSearch%3AespecialidadCierreField%3AespecialidadCierre=org.jboss.seam.ui.NoSelectionConverter.noSelectionValue&ticketSearch%3Aj_id209%3Aprovincia=org.jboss.seam.ui.NoSelectionConverter.noSelectionValue&ticketSearch%3AcentroField%3Acentro=org.jboss.seam.ui.NoSelectionConverter.noSelectionValue&ticketSearch%3Aj_id254%3Asitio=org.jboss.seam.ui.NoSelectionConverter.noSelectionValue&ticketSearch%3Aj_id269%3Aprioridad=org.jboss.seam.ui.NoSelectionConverter.noSelectionValue&ticketSearch%3AelementoField%3Aelemento=org.jboss.seam.ui.NoSelectionConverter.noSelectionValue&ticketSearch%3Aj_id298%3Acategoria=org.jboss.seam.ui.NoSelectionConverter.noSelectionValue&ticketSearch%3AoperadorField%3AsuggestionIdoperador_selection=&ticketSearch%3AoperadorField%3Aoperador="
    buscar = buscar & "&ticketSearch%3AticketoperadorField%3AsuggestionIdoperador_selection=&ticketSearch%3AticketoperadorField%3Aticketoperador="
    buscar = buscar & "&ticketSearch%3AusuarioAbreField%3AusuarioAbre=%3A&ticketSearch%3AusuarioResponsableField%3AusuarioResponsable=%3A&javax.faces.ViewState=" & java.Value & "&ajaxSingle=ticketSearch%3Asearch&ticketSearch%3Aj_id543=ticketSearch%3Aj_id543&"
    
    'Solicitar búsqueda
    XMLPage.Open "POST", urlListadoT, False
    XMLPage.setRequestHeader "Referer", urlListadoT
    XMLPage.send buscar
    
    buscar = "ticketSearch=ticketSearch&ticketSearch%3Aj_id31=true&ticketSearch%3Aj_id66%3Asuggestionnumero_selection=&ticketSearch%3Aj_id66%3Anumero="
    buscar = buscar & "&ticketSearch%3Aj_id82%3Asuggestionnumero_selection=&ticketSearch%3Aj_id82%3Atitulo="
    buscar = buscar & "&ticketSearch%3AfechaInicioDesdeField%3AfechaInicioDesdeInputDate=" & fechaInicioDesdeCodif & "&ticketSearch%3AfechaInicioDesdeField%3AfechaInicioDesdeInputCurrentDate=" & mesInicioDesde & "&ticketSearch%3AfechaInicioHastaField%3AfechaInicioHastaInputDate=" & fechaInicioHastaCodif & "&ticketSearch%3AfechaInicioHastaField%3AfechaInicioHastaInputCurrentDate=" & mesInicioHasta & "&ticketSearch%3AfechaFinDesdeField%3AfechaFinDesdeInputDate=&ticketSearch%3AfechaFinDesdeField%3AfechaFinDesdeInputCurrentDate=08%2F2023&ticketSearch%3AfechaFinHastaField%3AfechaFinHastaInputDate=&ticketSearch%3AfechaFinHastaField%3AfechaFinHastaInputCurrentDate=08%2F2023"
    buscar = buscar & "&ticketSearch%3AtipoProblemaField%3AtipoProblema=org.jboss.seam.ui.NoSelectionConverter.noSelectionValue"
    buscar = buscar & estados
    buscar = buscar & "&ticketSearch%3AespecialidadField%3Aespecialidad=org.jboss.seam.ui.NoSelectionConverter.noSelectionValue&ticketSearch%3AespecialidadCierreField%3AespecialidadCierre=org.jboss.seam.ui.NoSelectionConverter.noSelectionValue&ticketSearch%3Aj_id209%3Aprovincia=org.jboss.seam.ui.NoSelectionConverter.noSelectionValue&ticketSearch%3AcentroField%3Acentro=org.jboss.seam.ui.NoSelectionConverter.noSelectionValue&ticketSearch%3Aj_id254%3Asitio=org.jboss.seam.ui.NoSelectionConverter.noSelectionValue&ticketSearch%3Aj_id269%3Aprioridad=org.jboss.seam.ui.NoSelectionConverter.noSelectionValue&ticketSearch%3AelementoField%3Aelemento=org.jboss.seam.ui.NoSelectionConverter.noSelectionValue&ticketSearch%3Aj_id298%3Acategoria=org.jboss.seam.ui.NoSelectionConverter.noSelectionValue&ticketSearch%3AoperadorField%3AsuggestionIdoperador_selection=&ticketSearch%3AoperadorField%3Aoperador="
    buscar = buscar & "&ticketSearch%3AticketoperadorField%3AsuggestionIdoperador_selection=&ticketSearch%3AticketoperadorField%3Aticketoperador="
    buscar = buscar & "&ticketSearch%3AusuarioAbreField%3AusuarioAbre=%3A&ticketSearch%3AusuarioResponsableField%3AusuarioResponsable=%3A&ticketSearch%3Asearch=Buscar&javax.faces.ViewState=" & java.Value


    
    'Set htmlims = htmldoc.getElementById("richtoolBar_body")
    Set htmlim = htmldoc.getElementById("mytab1_lbl")
    
    Debug.Print htmlim.innerText
   
    'Solicitar búsqueda
    XMLPage.Open "POST", urlListadoT, False
    XMLPage.send buscar
    
    ' Crea un nuevo objeto HTML para analizar el contenido de la p�gina web
    Set htmldoc = CreateObject("HTMLFile")
    XMLPage.Open "GET", urlListadoT & "?operadorId=&ticketoperadorId=&numeroId=&tituloId=&fechaInicioDesdeId=" & fechaInicioDesdeCodif & "&fechaInicioHastaId=" & fechaInicioHastaCodif & "&colectionEstadoToString=&resultCantId=5&colectionProvinciarespToString=6&colectionProvinciaToString=6&cid=" & cid & "&conversationPropagation=join"
    
    'htmldoc.body.innerHTML = XMLPage.responseText
    XMLPage.send "cid=" & cid
    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Open
    fileStream.Type = 1
    fileStream.Write XMLPage.responseBody
    
    ' Seleccionando ruta de guardar documento
    ruta = ThisWorkbook.Path & "\Reportes de tickets"
    Call CrearCarpetaSiNoExiste(ruta)
    ruta = ruta & "\" & Year(Now)
    Call CrearCarpetaSiNoExiste(ruta)
    ruta = ruta & "\" & MonthName(Month(Now))
    Call CrearCarpetaSiNoExiste(ruta)
    ruta = ruta & "\" & Day(Now) & " a las " & Format(Now, "hh-mm-ss") & ".xls"
    
    fileStream.SaveToFile ruta, 2
    fileStream.Close
    Set fileStream = Nothing
    Set htmldoc = Nothing
    
    Call Actualizar_tabla_tickets(ruta, listaTickets)
End Sub

Function ParametrosHTTP(parametros() As Object) As String
    Dim cadenaParametros As String, i As Integer
    
    cadenaParametros = ""
    
    For i = LBound(parametros) To UBound(parametros)
        cadenaParametros = cadenaParametros + parametros(i).ID + "=" + parametros.Value
        If i < UBound(parametros) Then
            cadenaParametros = cadenaParametros + "&"
        End If
    Next i
End Function

Sub prueba()
    Dim grafico As ChartObject, serie As FullSeriesCollection, i As Integer
    
    For Each grafico In Hoja2.ChartObjects
        Debug.Print grafico.Chart.ChartTitle.Text
        Debug.Print grafico.Chart.Parent.Name
        Debug.Print grafico.Chart.FullSeriesCollection(1)
        
        For i = 1 To 4
            
        Next i
    Next grafico
End Sub

Sub CrearCarpetaSiNoExiste(ruta As String)
    If Dir(ruta, vbDirectory) = "" Then
        MkDir ruta
    End If
End Sub
