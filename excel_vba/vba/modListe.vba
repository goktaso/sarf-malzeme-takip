Option Explicit

Private UnitColumn As Long
Public QtyColumn As Long
Public CostColumn As Long
Public MonthColumn As Long
Private ReportColumn As Long
Public CountVariantColumn As Long
Public CountKey As String
Public SarfManualFiltered As Boolean
Public SarfManualMonth As Boolean
Public SarfFilterContext As String
Private DataColumns As Long
Private ReportSignature As String
Private ReportMonths As Long

Public Sub SarfRender(ByVal xml As Object)
    Dim ws As Worksheet, lo As ListObject, item As Object, node As Object
    Dim heads As Object, records As Object, cells() As Variant, numeric() As Boolean
    Dim previous() As Variant, filter As Filter, keep As Boolean, title As String, role As String
    Dim r As Long, c As Long, a As Long, rowCount As Long, search As String
    Dim operationQty As String, operationCost As String, target As Range
    Set ws = ThisWorkbook.Worksheets("Liste")
    SarfBusy = True
    On Error GoTo Failed
    title = xml.DocumentElement.getAttribute("title")
    keep = (CStr(ws.Range("A1").Value2) = title)
    If keep Then
        search = CStr(ws.Range("B3").Value2)
        If QtyColumn > 0 Then operationQty = CStr(ws.Cells(5, QtyColumn).Value2)
        If CostColumn > 0 Then operationCost = CStr(ws.Cells(5, CostColumn).Value2)
    End If
    If ws.ListObjects.Count > 0 Then
        Set lo = ws.ListObjects(1)
        If keep Then
            ReDim previous(1 To lo.ListColumns.Count, 1 To 4)
            On Error Resume Next
            For c = 1 To lo.ListColumns.Count
                Set filter = lo.AutoFilter.Filters(c)
                If filter.On Then
                    previous(c, 1) = filter.Criteria1
                    previous(c, 2) = filter.Operator
                    previous(c, 3) = filter.Criteria2
                    previous(c, 4) = lo.ListColumns(c).Name
                End If
            Next
            On Error GoTo Failed
        End If
        lo.Unlist
    End If
    ws.Cells.Clear
    ws.Range("A9:A" & Application.Max(9, ws.UsedRange.Rows.Count + ws.UsedRange.Row - 1)).EntireRow.Hidden = False
    ws.Columns("A:AZ").Hidden = False
    Set heads = xml.SelectNodes("/grid/column")
    Set records = xml.SelectNodes("/grid/row")
    DataColumns = heads.Length
    rowCount = records.Length
    UnitColumn = 0: QtyColumn = 0: CostColumn = 0: MonthColumn = 0: ReportColumn = 0: CountVariantColumn = 0
    CountKey = SarfClient.AttributeValue(xml.DocumentElement, "count_key")
    ReDim numeric(1 To DataColumns)
    ReDim cells(1 To rowCount + 1, 1 To DataColumns + 2)
    c = 0
    For Each node In heads
        c = c + 1
        cells(1, c) = node.Text
        role = node.getAttribute("role")
        Select Case role
            Case "unit": UnitColumn = c
            Case "qty": QtyColumn = c
            Case "cost": CostColumn = c
            Case "month": MonthColumn = c
            Case "report": ReportColumn = c
            Case "count-variant": CountVariantColumn = c
        End Select
    Next
    cells(1, DataColumns + 1) = "_Arama"
    cells(1, DataColumns + 2) = "_Gorunurluk"
    r = 1
    For Each item In records
        r = r + 1
        c = 0
        For Each node In item.SelectNodes("cell")
            c = c + 1
            If node.getAttribute("type") = "number" Then
                numeric(c) = True
                cells(r, c) = Val(node.Text)
            Else
                cells(r, c) = SarfLiteral(node.Text)
            End If
            cells(r, DataColumns + 1) = cells(r, DataColumns + 1) & " " & node.Text
        Next
    Next
    Set target = ws.Cells(8, 1).Resize(rowCount + 1, DataColumns + 2)
    target.NumberFormat = "@"
    For c = 1 To DataColumns
        If numeric(c) And rowCount > 0 Then ws.Cells(9, c).Resize(rowCount, 1).NumberFormat = "0.###"
    Next
    target.Value2 = cells
    Set lo = ws.ListObjects.Add(xlSrcRange, target, , xlYes)
    lo.Name = "tblSarf"
    lo.TableStyle = "TableStyleMedium4"
    If rowCount > 0 Then
        lo.ListColumns(DataColumns + 2).DataBodyRange.NumberFormat = "General"
        lo.ListColumns(DataColumns + 2).DataBodyRange.FormulaR1C1 = "=SUBTOTAL(103,RC1)"
    End If
    target.EntireColumn.AutoFit
    For c = 1 To DataColumns
        If ws.Columns(c).ColumnWidth < 16 Then ws.Columns(c).ColumnWidth = 16
        If ws.Columns(c).ColumnWidth > 44 Then ws.Columns(c).ColumnWidth = 44
    Next
    ws.Columns(DataColumns + 1).Hidden = True
    ws.Columns(DataColumns + 2).Hidden = True
    If ReportColumn > 0 Then ws.Columns(ReportColumn).Hidden = True
    If CountVariantColumn > 0 Then ws.Columns(CountVariantColumn).Hidden = True
    With ws.Range("A1")
        .Value2 = title
        .Font.Size = 22
        .Font.Bold = True
        .Font.Color = RGB(21, 77, 62)
    End With
    ws.Range("A3").Value2 = "Ürün ara"
    ws.Range("A3").Font.Bold = True
    ws.Range("B3").Value2 = search
    ws.Range("B3").Interior.Color = RGB(232, 246, 239)
    ws.Rows(3).RowHeight = 28
    ws.Shapes("SarfAddButton").Visible = msoFalse
    If SarfViewKey = "products" Or SarfViewKey = "stock" Then
        If Not SarfMenu.SelectSingleNode("/menu/action[@key='product']") Is Nothing Then ws.Shapes("SarfAddButton").Visible = msoTrue
    End If
    ws.Shapes("SarfExcelButton").Visible = (SarfMenu.DocumentElement.getAttribute("export") = "1")
    ws.Shapes("SarfPdfButton").Visible = ws.Shapes("SarfExcelButton").Visible
    ws.Shapes("SarfCsvButton").Visible = ws.Shapes("SarfExcelButton").Visible
    ws.Shapes("SarfCountButton").Visible = (CountKey <> "")
    ws.Shapes("SarfCountButton").Left = ws.Shapes("SarfSearchButton").Left + 85
    ws.Shapes("SarfSearchButton").Left = ws.Range("C3").Left + 4
    ws.Shapes("SarfAddButton").Left = ws.Shapes("SarfSearchButton").Left + 85
    If SarfViewKey = "stock" Or SarfViewKey = "movements" Then
        If QtyColumn > 0 Then CalculationControl ws, QtyColumn, operationQty
        If CostColumn > 0 Then CalculationControl ws, CostColumn, operationCost
    End If
    If keep And Not IsEmptyArray(previous) Then
        On Error Resume Next
        For a = LBound(previous, 1) To UBound(previous, 1)
            If Not IsEmpty(previous(a, 1)) Then
                c = lo.ListColumns(CStr(previous(a, 4))).Index
                If IsEmpty(previous(a, 3)) Then
                    lo.Range.AutoFilter Field:=c, Criteria1:=previous(a, 1), Operator:=previous(a, 2)
                Else
                    lo.Range.AutoFilter Field:=c, Criteria1:=previous(a, 1), Operator:=previous(a, 2), Criteria2:=previous(a, 3)
                End If
            End If
        Next
        On Error GoTo Failed
    End If
    ReportSignature = ""
    ReportMonths = 0
    SarfBusy = False
    SarfToplam
    Exit Sub
Failed:
    SarfBusy = False
    Err.Raise Err.Number, , Err.Description
End Sub

Private Function IsEmptyArray(ByRef items() As Variant) As Boolean
    Dim size As Long
    On Error Resume Next
    size = UBound(items)
    IsEmptyArray = (Err.Number <> 0)
    Err.Clear
End Function

Private Function SarfLiteral(ByVal value As String) As String
    If Len(value) = 0 Then Exit Function
    If InStr(1, "=+-@", Left$(value, 1), vbBinaryCompare) > 0 Then value = "'" & value
    SarfLiteral = value
End Function

Private Sub CalculationControl(ByVal ws As Worksheet, ByVal col As Long, ByVal previous As String)
    Dim separator As String
    separator = Application.International(xlListSeparator)
    With ws.Cells(5, col)
        .Validation.Delete
        .Validation.Add xlValidateList, xlValidAlertStop, xlBetween, Join(Array("Toplam", "Ortalama", "Min", "Max"), separator)
        .Value2 = IIf(previous = "", "Toplam", previous)
        .Interior.Color = RGB(232, 246, 239)
    End With
    ws.Cells(6, col).Font.Bold = True
End Sub

Public Sub SarfAra()
    On Error GoTo Failed
    Dim ws As Worksheet, lo As ListObject, query As String
    Set ws = ThisWorkbook.Worksheets("Liste")
    If ws.ListObjects.Count = 0 Then Exit Sub
    Set lo = ws.ListObjects(1)
    query = Trim$(CStr(ws.Range("B3").Value2))
    query = Replace(Replace(Replace(query, "~", "~~"), "*", "~*"), "?", "~?")
    If query = "" Then
        lo.Range.AutoFilter Field:=DataColumns + 1
    Else
        lo.Range.AutoFilter Field:=DataColumns + 1, Criteria1:="=*" & query & "*"
    End If
    SarfToplam
    Exit Sub
Failed:
    SarfHata
End Sub

Public Sub SarfSayimMiktari()
    On Error GoTo Failed
    Dim ws As Worksheet, lo As ListObject, row As Range, variantId As String
    If CountKey = "" Or CountVariantColumn = 0 Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("Liste")
    Set lo = ws.ListObjects(1)
    If lo.DataBodyRange Is Nothing Then Exit Sub
    Set row = Intersect(ActiveCell.EntireRow, lo.DataBodyRange)
    If row Is Nothing Then Err.Raise vbObjectError + 713, , "Önce miktarı düzeltilecek ürün satırını seçin."
    variantId = CStr(row.Cells(1, CountVariantColumn).Value2)
    If variantId = "" Then Exit Sub
    frmSarfIslem.Hazirla "count-line", CountKey & ":" & variantId
    SarfShowForm frmSarfIslem, 125
    Exit Sub
Failed:
    SarfHata
End Sub

Private Sub Aggregate(ByVal totals As Object, ByVal unit As String, ByVal value As Double)
    Dim previous As Variant
    If totals.Exists(unit) Then
        previous = totals(unit)
        previous(0) = previous(0) + value
        previous(1) = previous(1) + 1
        If value < previous(2) Then previous(2) = value
        If value > previous(3) Then previous(3) = value
        totals(unit) = previous
    Else
        totals.Add unit, Array(value, 1, value, value)
    End If
End Sub

Private Function Formatted(ByVal totals As Object, ByVal operation As String) As String
    Dim key As Variant, entry As Variant, value As Double
    For Each key In totals.Keys
        entry = totals(key)
        Select Case operation
            Case "Ortalama": value = entry(0) / entry(1)
            Case "Min": value = entry(2)
            Case "Max": value = entry(3)
            Case Else: value = entry(0)
        End Select
        If Formatted <> "" Then Formatted = Formatted & " | "
        Formatted = Formatted & Format$(value, "#,##0.###") & " " & key
    Next
    If Formatted = "" Then Formatted = "-"
End Function

Private Function SummaryQty(ByVal xml As Object, ByVal key As String) As String
    Dim item As Object
    For Each item In xml.SelectNodes("/summary/" & key)
        If SummaryQty <> "" Then SummaryQty = SummaryQty & " | "
        SummaryQty = SummaryQty & Format$(Val(item.Text), "#,##0.###") & " " & item.getAttribute("unit")
    Next
    If SummaryQty = "" Then SummaryQty = "-"
End Function

Public Sub SarfToplam()
    If SarfBusy Or SarfViewKey = "" Then Exit Sub
    If SarfViewKey <> "report" And SarfViewKey <> "stock" And SarfViewKey <> "movements" Then Exit Sub
    Dim ws As Worksheet, lo As ListObject, row As ListRow, units As Object, values As Object, ids As Object
    Dim unit As String, k As String, filtered As Boolean, monthFilter As Boolean
    Dim signature As String, summary As Object, cursor As Filter, total As Double, extra As Object
    Set ws = ThisWorkbook.Worksheets("Liste")
    If ws.ListObjects.Count = 0 Then Exit Sub
    Set lo = ws.ListObjects(1)
    Set units = CreateObject("Scripting.Dictionary")
    Set values = CreateObject("Scripting.Dictionary")
    Set ids = CreateObject("Scripting.Dictionary")
    SarfBusy = True
    On Error GoTo Finished
    For Each row In lo.ListRows
        If Not row.Range.EntireRow.Hidden And Len(CStr(row.Range.Cells(1, 1).Value2)) > 0 Then
            If UnitColumn <> 0 Then unit = CStr(row.Range.Cells(1, UnitColumn).Value2)
            If QtyColumn <> 0 Then
                If IsNumeric(row.Range.Cells(1, QtyColumn).Value2) Then Aggregate units, unit, CDbl(row.Range.Cells(1, QtyColumn).Value2)
            End If
            If CostColumn <> 0 Then
                If IsNumeric(row.Range.Cells(1, CostColumn).Value2) Then
                    total = total + CDbl(row.Range.Cells(1, CostColumn).Value2)
                    Aggregate values, "TL", CDbl(row.Range.Cells(1, CostColumn).Value2)
                End If
            End If
            If ReportColumn <> 0 Then
                k = CStr(row.Range.Cells(1, ReportColumn).Value2)
                If Not ids.Exists(k) Then ids.Add k, True
            End If
        End If
    Next
    If SarfViewKey <> "report" Then
        If QtyColumn <> 0 Then ws.Cells(6, QtyColumn).Value2 = Formatted(units, CStr(ws.Cells(5, QtyColumn).Value2))
        If CostColumn <> 0 Then ws.Cells(6, CostColumn).Value2 = Formatted(values, CStr(ws.Cells(5, CostColumn).Value2))
        GoTo Finished
    End If
    filtered = SarfManualFiltered
    monthFilter = SarfManualMonth
    For Each cursor In lo.AutoFilter.Filters
        If cursor.On Then filtered = True
    Next
    If MonthColumn > 0 Then monthFilter = monthFilter Or lo.AutoFilter.Filters(MonthColumn).On
    signature = SarfViewQuery & SarfParam("keys", Join(ids.Keys, ",")) & SarfParam("filtered", IIf(filtered, "1", "0")) _
                & SarfParam("month_filter", IIf(monthFilter, "1", "0"))
    ws.Range("A5").Value2 = "Tüketim miktarı: " & Formatted(units, "Toplam")
    If CostColumn > 0 Then ws.Range("A7").Value2 = "Net tüketim maliyeti: " & Format$(total, "#,##0.00") & " TL"
    If signature <> ReportSignature Then
        Set summary = SarfClient.Xml("POST", "/summary", signature, "application/x-www-form-urlencoded")
        ReportSignature = signature
        ReportMonths = Val(summary.DocumentElement.getAttribute("months"))
        If CStr(summary.DocumentElement.getAttribute("day")) <> "" Then
            ws.Range("A4").Value2 = "Giriş miktarı: " & SummaryQty(summary, "received")
            ws.Range("A6").Value2 = "Son kalan bakiye (" & summary.DocumentElement.getAttribute("day") & "): " & SummaryQty(summary, "balance")
            Set extra = summary.SelectSingleNode("/summary/value")
            If Not extra Is Nothing Then ws.Range("E7").Value2 = "Depo giriş değeri: " & Format$(Val(extra.Text), "#,##0.00") & " TL"
        End If
    End If
    If CostColumn > 0 Then
        ws.Range("D5").Value2 = "Aylık ortalama (" & ReportMonths & " ay): " & _
            Format$(total / IIf(ReportMonths = 0, 1, ReportMonths), "#,##0.00") & " TL"
    End If
Finished:
    SarfBusy = False
End Sub

Public Sub SarfExcel(): SarfExport "xlsx": End Sub
Public Sub SarfPDF(): SarfExport "pdf": End Sub
Public Sub SarfCSV(): SarfExport "csv": End Sub

Public Sub SarfExport(ByVal extension As String, Optional ByVal targetFile As String = "")
    Dim view As Worksheet, lo As ListObject, output As Workbook, sheet As Worksheet, row As ListRow
    Dim filename As Variant, c As Long, a As Long, r As Long, headings As Long, menu As Object, originalAlerts As Boolean
    originalAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    If SarfViewKey = "" Or SarfClient Is Nothing Then Exit Sub
    Set menu = SarfClient.Xml("GET", "/menu")
    If menu.DocumentElement.getAttribute("export") <> "1" Then Err.Raise vbObjectError + 711, , "Çıktı yetkiniz yok."
    Set view = ThisWorkbook.Worksheets("Liste")
    If view.ListObjects.Count = 0 Then Exit Sub
    Set lo = view.ListObjects(1)
    filename = targetFile
    If targetFile = "" Then filename = Application.GetSaveAsFilename("Sarf-" & SarfViewKey & "." & extension, _
        UCase$(extension) & " (*." & extension & "), *." & extension)
    If VarType(filename) = vbBoolean Then Exit Sub
    If extension = "csv" Then
        WriteCsv lo, CStr(filename)
        Exit Sub
    ElseIf extension = "pdf" Then
        WritePdf view, lo, CStr(filename)
        Exit Sub
    End If
    Set output = Workbooks.Add(xlWBATWorksheet)
    Set sheet = output.Worksheets(1)
    headings = 1
    If extension <> "csv" Then
        sheet.Range("A1").Value2 = view.Range("A1").Value2
        sheet.Range("A1").Font.Size = 20
        sheet.Range("A1").Font.Bold = True
        If SarfViewKey = "report" Then
            sheet.Range("A2:A5").Value2 = view.Range("A4:A7").Value2
            sheet.Range("A6").Value2 = view.Range("E7").Value2
            sheet.Range("A7").Value2 = view.Range("D5").Value2
            headings = 9
        ElseIf SarfViewKey = "stock" Or SarfViewKey = "movements" Then
            For c = 1 To DataColumns
                If view.Cells(6, c).Value2 <> "" Then
                    sheet.Cells(3, 1).Value2 = sheet.Cells(3, 1).Value2 & "  " & lo.ListColumns(c).Name & _
                        " " & view.Cells(5, c).Value2 & ": " & view.Cells(6, c).Value2
                End If
            Next
            headings = 5
        Else
            headings = 3
        End If
        sheet.Cells(headings - 1, 1).Value2 = FilterDescription(lo)
    End If
    a = 0
    For c = 1 To DataColumns
        If c <> ReportColumn And c <> CountVariantColumn Then
            a = a + 1
            sheet.Cells(headings, a).Value2 = lo.ListColumns(c).Name
        End If
    Next
    r = headings
    For Each row In lo.ListRows
        If Not row.Range.EntireRow.Hidden And Len(CStr(row.Range.Cells(1, 1).Value2)) > 0 Then
            r = r + 1
            a = 0
            For c = 1 To DataColumns
                If c <> ReportColumn And c <> CountVariantColumn Then
                    a = a + 1
                    If VarType(row.Range.Cells(1, c).Value2) = vbString Then
                        sheet.Cells(r, a).NumberFormat = "@"
                        sheet.Cells(r, a).Value2 = SarfLiteral(CStr(row.Range.Cells(1, c).Value2))
                    Else
                        sheet.Cells(r, a).Value2 = row.Range.Cells(1, c).Value2
                    End If
                End If
            Next
        End If
    Next
    If a > 0 Then
        With sheet.Cells(headings, 1).Resize(1, a)
            .Font.Bold = True
            .Interior.Color = RGB(22, 76, 64)
            .Font.Color = RGB(255, 255, 255)
        End With
        sheet.Cells(headings, 1).Resize(r - headings + 1, a).EntireColumn.AutoFit
        For c = 1 To a
            If sheet.Columns(c).ColumnWidth > 45 Then sheet.Columns(c).ColumnWidth = 45
        Next
    End If
    Application.DisplayAlerts = False
    Select Case extension
        Case "xlsx": output.SaveAs CStr(filename), 51
    End Select
    output.Close False
    Application.DisplayAlerts = originalAlerts
    Exit Sub
Failed:
    Application.DisplayAlerts = originalAlerts
    If Not output Is Nothing Then output.Close False
    If targetFile <> "" Then Err.Raise Err.Number, , Err.Description
    SarfHata
End Sub

Private Sub WritePdf(ByVal view As Worksheet, ByVal lo As ListObject, ByVal filename As String)
    Dim doc As Object, root As Object, section As Object, stripe As ListRow, entry As Object, c As Long, item As Variant
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.LoadXML "<export/>"
    Set root = doc.DocumentElement
    root.setAttribute "name", SarfViewKey
    root.setAttribute "context", FilterDescription(lo)
    If SarfViewKey = "report" Or SarfViewKey = "stock" Or SarfViewKey = "movements" Then
        Set section = AddNode(doc, root, "section")
        section.setAttribute "title", "Özet"
        AddNode doc, section, "column", "Gösterge"
        AddNode doc, section, "column", "Değer"
        If SarfViewKey = "report" Then
            For Each item In Array("A4", "A5", "A6", "A7", "D5", "E7")
                If view.Range(item).Value2 <> "" Then
                    Set entry = AddNode(doc, section, "row")
                    AddNode doc, entry, "cell", CStr(Split(CStr(view.Range(item).Value2), ": ", 2)(0))
                    AddNode doc, entry, "cell", Mid$(CStr(view.Range(item).Value2), InStr(CStr(view.Range(item).Value2), ": ") + 2)
                End If
            Next
        Else
            For c = 1 To DataColumns
                If view.Cells(6, c).Value2 <> "" Then
                    Set entry = AddNode(doc, section, "row")
                    AddNode doc, entry, "cell", lo.ListColumns(c).Name & " · " & view.Cells(5, c).Value2
                    AddNode doc, entry, "cell", CStr(view.Cells(6, c).Value2)
                End If
            Next
        End If
    End If
    Set section = AddNode(doc, root, "section")
    section.setAttribute "title", "Filtreli liste"
    For c = 1 To DataColumns
        If c <> ReportColumn And c <> CountVariantColumn Then AddNode doc, section, "column", lo.ListColumns(c).Name
    Next
    For Each stripe In lo.ListRows
        If Not stripe.Range.EntireRow.Hidden And Len(CStr(stripe.Range.Cells(1, 1).Value2)) > 0 Then
            Set entry = AddNode(doc, section, "row")
            For c = 1 To DataColumns
                If c <> ReportColumn And c <> CountVariantColumn Then
                    AddNode doc, entry, "cell", CStr(stripe.Range.Cells(1, c).Value2)
                End If
            Next
        End If
    Next
    SarfClient.Download "/pdf", filename, doc.XML, "application/xml; charset=utf-8"
End Sub

Private Function FilterDescription(ByVal lo As ListObject) As String
    Dim c As Long, criteria As Variant, value As String
    For c = 1 To lo.ListColumns.Count
        If lo.AutoFilter.Filters(c).On Then
            value = ""
            On Error Resume Next
            criteria = Empty
            criteria = lo.AutoFilter.Filters(c).Criteria1
            If IsArray(criteria) Then value = Join(criteria, ", ") Else value = CStr(criteria)
            On Error GoTo 0
            If FilterDescription <> "" Then FilterDescription = FilterDescription & " | "
            FilterDescription = FilterDescription & lo.ListColumns(c).Name & ": " & value
        End If
    Next
    If SarfFilterContext <> "" Then FilterDescription = SarfFilterContext
    If FilterDescription <> "" Then FilterDescription = "Filtreler: " & Left$(FilterDescription, 1400)
End Function

Private Function AddNode(ByVal doc As Object, ByVal parent As Object, ByVal tag As String, _
                         Optional ByVal value As String = "") As Object
    Dim child As Object
    Set child = doc.createElement(tag)
    child.Text = value
    parent.appendChild child
    Set AddNode = child
End Function

Private Sub WriteCsv(ByVal lo As ListObject, ByVal filename As String)
    Dim stream As Object, stripe As ListRow, fields() As String, c As Long, index As Long, value As Variant
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    ReDim fields(0 To DataColumns - 1)
    For c = 1 To DataColumns
        If c <> ReportColumn And c <> CountVariantColumn Then
            fields(index) = CsvCell(lo.ListColumns(c).Name)
            index = index + 1
        End If
    Next
    ReDim Preserve fields(0 To index - 1)
    stream.WriteText Join(fields, ";"), 1
    For Each stripe In lo.ListRows
        If Not stripe.Range.EntireRow.Hidden And Len(CStr(stripe.Range.Cells(1, 1).Value2)) > 0 Then
            index = 0
            For c = 1 To DataColumns
                If c <> ReportColumn And c <> CountVariantColumn Then
                    value = stripe.Range.Cells(1, c).Value2
                    If VarType(value) = vbDouble Then
                        fields(index) = Replace(Trim$(Str$(value)), ".", ",")
                    Else
                        fields(index) = CsvCell(SarfLiteral(CStr(value)))
                    End If
                    index = index + 1
                End If
            Next
            stream.WriteText Join(fields, ";"), 1
        End If
    Next
    stream.SaveToFile filename, 2
    stream.Close
End Sub

Private Function CsvCell(ByVal value As String) As String
    Dim quote As String
    quote = Chr$(34)
    CsvCell = quote & Replace(value, quote, quote & quote) & quote
End Function
