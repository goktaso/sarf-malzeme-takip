Option Explicit

Public SwitchUserRequested As Boolean

Private currentSection As String
Private changing As Boolean
Private table As ListObject
Private matrix As Variant
Private visibleColumns As Collection
Private filteredRows As Collection
Private filters As Object
Private filterSets As Object
Private Const GridRows As Long = 32
Private rowHeight As Long
Private gridBottom As Single
Private filterColumnId As Long
Private filterChoices As Object
Private filterMarked As Object
Private filterUpdating As Boolean
Private Const GridColumns As Long = 16
Private pageRows As Long
Private stockScrollUpdating As Boolean
Private gridAreaWidth As Single
Private heads(1 To GridColumns) As cSarfGridCell
Private columnWidths(1 To GridColumns) As Single
Private cells(1 To GridRows, 1 To GridColumns) As cSarfGridCell
Private firstRow As Long
Private selectedRow As Long

Public Sub Hazirla(ByVal menu As Object)
    cmdSwitchUser.Visible = False
    Dim day As String
    Set filterSets = CreateObject("Scripting.Dictionary")
    lblTitle.Font.Size = 18
    lblTitle.Font.Bold = True
    lblBrand.Font.Size = 13
    lblBrand.Font.Bold = True
    BuildHeaders
    pageRows = 5
    gridBottom = fraGrid.Top + fraGrid.Height
    gridAreaWidth = fraGrid.Width
    StylePage
    lblSession.Caption = menu.DocumentElement.getAttribute("account")
    day = menu.DocumentElement.getAttribute("today")
    txtStart.Text = Left$(day, 7) & "-01"
    txtEnd.Text = day
    cmdStock.Visible = Allowed("view", "stock")
    cmdMoves.Visible = Allowed("view", "movements")
    cmdProducts.Visible = Allowed("view", "products")
    cmdPeople.Visible = Allowed("view", "departments") Or Allowed("view", "people")
    cmdReports.Visible = Allowed("view", "report")
    cmdCounts.Visible = Allowed("view", "counts")
    cmdImport.Visible = IsAdmin Or cmdCounts.Visible
    cmdAdmin.Visible = IsAdmin
    cmdPeriod.Visible = IsAdmin
    Dim button As Variant
    For Each button In Array("cmdStock", "cmdMoves", "cmdProducts", "cmdPeople", "cmdReports", "cmdCounts", "cmdImport", "cmdAdmin", "cmdPeriod")
        If Me.Controls(button).Visible Then
            Navigate Mid$(button, 4)
            Exit Sub
        End If
    Next
    Navigate ""
End Sub

Private Function IsAdmin() As Boolean
    IsAdmin = SarfMenu.DocumentElement.getAttribute("admin") = "1"
End Function

Private Function Allowed(ByVal kind As String, ByVal key As String) As Boolean
    Allowed = Not (SarfMenu.SelectSingleNode("/menu/" & kind & "[@key='" & key & "']") Is Nothing)
End Function

Public Sub Navigate(ByVal section As String)
    On Error GoTo Failed
    Dim keys As String, actions As String, nav As Variant, control As Object
    changing = True
    currentSection = section
    For Each nav In Array("Stock", "Moves", "Products", "People", "Reports", "Counts", "Import", "Admin", "Period")
        Set control = Me.Controls("cmd" & nav)
        control.BackColor = IIf(nav = section, RGB(36, 85, 71), RGB(17, 47, 41))
    Next
    lblTitle.Caption = "Sarf Malzeme Takip"
    lblSummary.Visible = False
    lblExtra.Visible = False
    If section <> "" Then lblTitle.Caption = Me.Controls("cmd" & section).Caption
    Select Case section
        Case "Stock": keys = "stock"
        Case "Moves": keys = "movements": actions = "receipt person-issue department-issue return"
        Case "Products": keys = "products groups history": actions = "group group-edit product-edit variant"
        Case "People": keys = "departments people history": actions = "department department-edit person person-edit"
        Case "Reports": keys = "report"
        Case "Counts": keys = "counts count-lines": actions = "count-delete count-refresh count-lock count-reject count-unlock"
        Case "Admin": keys = "users audit": actions = "user-save"
        Case "Period": keys = "periods": actions = "period-open period-close"
    End Select
    FillOptions cboView, "view", keys
    FillOptions cboAction, "action", actions
    RefreshCounts
    If section = "Counts" And cboView.ListCount > 1 And cboCount.ListCount > 0 Then cboView.ListIndex = 1
    txtSearch.Text = ""
    lblSearch.Caption = IIf(section = "Products", "Ürün ara", "Listede ara")
    changing = False
    LoadPage
    Exit Sub
Failed:
    changing = False
    SarfHata
End Sub

Private Sub FillOptions(ByVal box As Object, ByVal kind As String, ByVal keys As String)
    Dim key As Variant, entry As Object
    box.Clear
    If Len(keys) > 0 Then
        For Each key In Split(keys)
            Set entry = SarfMenu.SelectSingleNode("/menu/" & kind & "[@key='" & key & "']")
            If Not entry Is Nothing Then SarfOptionAdd box, entry.getAttribute("label"), CStr(key)
        Next
    End If
    If box.ListCount > 0 Then box.ListIndex = 0
End Sub

Private Sub RefreshCounts()
    Dim entry As Object, previous As String, key As String, i As Long
    previous = Split(SarfOption(cboCount) & ":", ":")(0)
    cboCount.Clear
    If currentSection = "Import" And IsAdmin Then SarfOptionAdd cboCount, "Ana veri şablonu", ""
    For Each entry In SarfMenu.SelectNodes("/menu/count")
        SarfOptionAdd cboCount, entry.getAttribute("label"), entry.getAttribute("key")
    Next
    If cboCount.ListCount > 0 Then cboCount.ListIndex = 0
    For i = 0 To cboCount.ListCount - 1
        key = Split(cboCount.List(i, 1) & ":", ":")(0)
        If key = previous Then cboCount.ListIndex = i
    Next
End Sub

Private Sub LoadPage()
    On Error GoTo Failed
    Dim key As String, query As String, countId As String, listPage As Boolean, dates As Boolean, c As Long, scopedKey As String
    Dim ws As Worksheet, item As Variant, column As ListColumn, countEntry As Object
    CloseColumnFilter
    Set table = Nothing
    key = SarfOption(cboView)
    ArrangePage key
    listPage = key <> ""
    dates = key = "movements" Or key = "report" Or key = "people"
    changing = True
    For Each item In Array("fraGrid", "txtSearch", "lblSearch", "cmdClear", "lblRows", "cmdRowsBack", "cmdRowsForward")
        Me.Controls(item).Visible = listPage
    Next
    scrStock.Visible = False
    cmdRowsBack.Visible = listPage And key <> "stock"
    cmdRowsForward.Visible = cmdRowsBack.Visible
    For Each item In Array("cmdExcel", "cmdPdf", "cmdCsv")
        Me.Controls(item).Visible = listPage And SarfMenu.DocumentElement.getAttribute("export") = "1"
    Next
    For Each item In Array("txtStart", "txtEnd", "lblDates", "cmdDates")
        Me.Controls(item).Visible = dates
    Next
    txtStart.Visible = dates And key <> "people"
    For Each item In Array("cboCount", "lblCounts")
        Me.Controls(item).Visible = currentSection = "Counts" Or currentSection = "Import"
    Next
    lblCounts.Caption = IIf(currentSection = "Counts", "Kayıtlı sayımlar", "Aktarım türü")
    For Each item In Array("cmdTemplate", "cmdUpload")
        Me.Controls(item).Visible = currentSection = "Import" And cboCount.ListCount > 0
    Next
    cmdCountNew.Visible = currentSection = "Counts" And Allowed("action", "count-new")
    cmdCountTemplate.Visible = currentSection = "Counts"
    cmdCountUpload.Visible = currentSection = "Counts"
    cmdCountManual.Visible = currentSection = "Counts" And Allowed("action", "count-line")
    cmdCountPost.Visible = currentSection = "Counts" And Allowed("action", "count-approve")
    cmdCountOpen.Visible = key = "counts"
    If currentSection = "Counts" Then
        Set countEntry = SarfMenu.SelectSingleNode("/menu/count[@key='" & SarfOption(cboCount) & "']")
        cmdCountTemplate.Enabled = Not countEntry Is Nothing
        cmdCountOpen.Enabled = Not countEntry Is Nothing
        cmdCountUpload.Enabled = False
        cmdCountManual.Enabled = False
        cmdCountPost.Enabled = False
        If Not countEntry Is Nothing Then
            cmdCountUpload.Enabled = (countEntry.getAttribute("editable") = "1")
            cmdCountManual.Enabled = cmdCountUpload.Enabled
            cmdCountPost.Enabled = (countEntry.getAttribute("postable") = "1")
        End If
    End If
    cboView.Visible = cboView.ListCount > 1
    cboAction.Visible = cboAction.ListCount > 0
    If currentSection = "Counts" Then cboAction.Visible = cboAction.Visible And key = "count-lines"
    cmdAction.Visible = cboAction.Visible
    cmdProduct.Visible = currentSection = "Products" And Allowed("action", "product")
    cmdCountEdit.Visible = False
    cboQty.Visible = False
    cboCost.Visible = False
    lblSummary.Caption = ""
    lblExtra.Caption = ""
    lblSummary.Visible = False
    lblExtra.Visible = False
    lblRows.Caption = ""
    ClearGrid
    For c = 1 To GridColumns
        heads(c).Header.Visible = False
    Next
    If currentSection = "Import" Then
        lblExtra.Visible = True
        lblSummary.Caption = ""
        lblExtra.Caption = "Şablonu doldurduktan sonra Excel içe aktar düğmesiyle yükleyin."
    End If
    If Not listPage Then GoTo Finished
    If key = "movements" Or key = "report" Then
        query = SarfParam("from", txtStart.Text) & SarfParam("to", txtEnd.Text)
    ElseIf key = "people" Then
        query = SarfParam("to", txtEnd.Text)
    ElseIf key = "history" Then
        query = SarfParam("scope", IIf(currentSection = "Products", "products", "people"))
    ElseIf key = "count-lines" Then
        countId = Split(SarfOption(cboCount) & ":", ":")(0)
        If countId = "" Then
            lblSummary.Visible = True
            lblSummary.Caption = "Sayım satırlarını görmek için bir sayım seçin."
            GoTo Finished
        End If
        query = SarfParam("id", countId)
    End If
    scopedKey = currentSection & ":" & key & ":" & countId
    If Not filterSets.Exists(scopedKey) Then filterSets.Add scopedKey, CreateObject("Scripting.Dictionary")
    Set filters = filterSets(scopedKey)
    SarfManualFiltered = False
    SarfManualMonth = False
    SarfFilterContext = ""
    SarfListe key, query
    Set ws = ThisWorkbook.Worksheets("Liste")
    Set table = ws.ListObjects(1)
    Set visibleColumns = New Collection
    For Each column In table.ListColumns
        If Left$(column.Name, 1) <> "_" Then visibleColumns.Add column.Index
    Next
    matrix = Empty
    If Not table.DataBodyRange Is Nothing Then matrix = table.DataBodyRange.Value2
    cmdCountEdit.Visible = (CountKey <> "")
    SetupCalculation cboQty, QtyColumn, "Miktar"
    SetupCalculation cboCost, CostColumn, "Değer"
    RenderRows
Finished:
    changing = False
    Exit Sub
Failed:
    changing = False
    SarfHata
End Sub

Private Sub SetupCalculation(ByVal box As Object, ByVal col As Long, ByVal label As String)
    box.Visible = (col > 0 And (SarfViewKey = "stock" Or SarfViewKey = "movements"))
    If Not box.Visible Then Exit Sub
    Dim op As Variant, current As String
    current = CStr(table.Parent.Cells(5, col).Value2)
    box.Clear
    For Each op In Array("Toplam", "Ortalama", "Min", "Max")
        SarfOptionAdd box, label & " / " & op, CStr(op)
        If op = current Then box.ListIndex = box.ListCount - 1
    Next
    If box.ListIndex < 0 Then box.ListIndex = 0
End Sub

Private Function Matches(ByVal row As Long, Optional ByVal ignore As Long = 0) As Boolean
    Dim col As Variant, selected As Object, needle As String
    needle = Trim$(txtSearch.Text)
    If needle <> "" Then
        If InStr(1, CStr(matrix(row, table.ListColumns.Count - 1)), needle, vbTextCompare) = 0 Then Exit Function
    End If
    For Each col In filters.Keys
        If CLng(col) <> ignore Then
            Set selected = filters(col)
            If Not selected.Exists(CStr(matrix(row, CLng(col)))) Then Exit Function
        End If
    Next
    Matches = True
End Function

Private Sub RenderRows()
    If table Is Nothing Then Exit Sub
    Dim row As Long, runStart As Long, lastRow As Long, col As Variant, selected As Object, visible As Boolean
    Dim priorBusy As Boolean
    priorBusy = SarfBusy
    On Error GoTo Failed
    SarfBusy = True
    Set filteredRows = New Collection
    firstRow = 1
    selectedRow = 0
    SarfManualFiltered = (filters.Count > 0 Or Trim$(txtSearch.Text) <> "")
    SarfManualMonth = filters.Exists(CStr(MonthColumn))
    SarfFilterContext = Trim$(txtSearch.Text)
    If SarfFilterContext <> "" Then SarfFilterContext = "Arama: " & SarfFilterContext
    For Each col In filters.Keys
        Set selected = filters(col)
        If SarfFilterContext <> "" Then SarfFilterContext = SarfFilterContext & " | "
        SarfFilterContext = SarfFilterContext & table.ListColumns(CLng(col)).Name & ": " & Join(selected.Keys, ", ")
    Next
    If IsArray(matrix) Then
        lastRow = UBound(matrix, 1)
        table.DataBodyRange.EntireRow.Hidden = False
        For row = 1 To lastRow + 1
            visible = True
            If row <= lastRow Then visible = Matches(row)
            If visible Then
                If runStart > 0 Then
                    table.DataBodyRange.Rows(runStart).Resize(row - runStart).EntireRow.Hidden = True
                    runStart = 0
                End If
                If row <= lastRow Then filteredRows.Add row
            ElseIf runStart = 0 Then
                runStart = row
            End If
        Next
    End If
    SarfBusy = priorBusy
    SarfToplam
    RenderColumns
    RenderSummary

    Exit Sub
Failed:
    SarfBusy = priorBusy
    SarfHata
End Sub

Private Sub BuildHeaders()
    Dim col As Long, binding As cSarfGridCell
    For col = 1 To GridColumns
        Set binding = New cSarfGridCell
        Set binding.Header = Me.Controls.Add("Forms.CommandButton.1", "columnHeader" & col, False)
        Set binding.Owner = Me
        binding.ColumnSlot = col
        With binding.Header
            .Top = 250
            .Height = 39
            .Font.Name = "Aptos"
            .Font.Size = 9
            .WordWrap = True
            .ForeColor = RGB(23, 62, 49)
        End With
        Set heads(col) = binding
    Next
End Sub

Public Sub FitLayout()
    Dim zoom As Long, rightEdge As Single, bottomEdge As Single
    zoom = Me.Zoom
    Me.Zoom = 100
    rightEdge = Me.InsideWidth * 100# / zoom - 18
    bottomEdge = Me.InsideHeight * 100# / zoom - 13
    Me.Zoom = zoom
    gridAreaWidth = rightEdge - fraGrid.Left
    fraGrid.Width = gridAreaWidth
    gridBottom = bottomEdge - 33
    lblSide.Height = bottomEdge + 13
    cmdLogout.Top = bottomEdge - cmdLogout.Height
    cmdSwitchUser.Top = cmdLogout.Top - cmdSwitchUser.Height - 8
    lblRows.Top = bottomEdge - lblRows.Height
    cmdRowsForward.Move rightEdge - cmdRowsForward.Width, bottomEdge - cmdRowsForward.Height
    cmdRowsBack.Move cmdRowsForward.Left - cmdRowsBack.Width - 8, cmdRowsForward.Top
    lblTitle.Width = fraGrid.Width
    lblSession.Width = fraGrid.Width
    ArrangePage SarfOption(cboView)
    ClearGrid
    If Not table Is Nothing Then RenderColumns
End Sub

Private Sub ArrangePage(ByVal key As String)
    Dim summaryWidth As Single, contentLeft As Single
    contentLeft = fraGrid.Left
    If key = "stock" Then
        lblSearch.Move contentLeft, 76, 68, 18
        txtSearch.Move contentLeft + 71, 70, 185, 24
        cmdClear.Move 456, 70, 60, 24
        cmdExcel.Move 524, 70, 46, 24
        cmdPdf.Move 578, 70, 46, 24
        cmdCsv.Move 632, 70, 46, 24
        cboQty.Move 686, 70, 84, 24
        cboCost.Move 778, 70, 84, 24
        summaryWidth = (gridAreaWidth - 14) / 2
        lblSummary.Move contentLeft, 105, summaryWidth, 25
        lblExtra.Move contentLeft + 14 + summaryWidth, 105, summaryWidth, 25
        fraGrid.Top = 171
        rowHeight = 26
    Else
        lblSearch.Move 192, 101, 240, 15
        txtSearch.Move 192, 118, 236, 26
        cmdClear.Move 550, 118, 84, 26
        cmdExcel.Move 666, 118, 60, 26
        cmdPdf.Move 733, 118, 60, 26
        cmdCsv.Move 800, 118, 60, 26
        cboQty.Move 583, 151, 133, 24
        cboCost.Move 725, 151, 135, 24
        lblSummary.Move contentLeft, 185, gridAreaWidth, 33
        lblExtra.Move contentLeft, 219, gridAreaWidth, 23
        fraGrid.Top = 290
        rowHeight = 34
    End If
    cboQty.ColumnWidths = CStr(cboQty.Width - 15) & " pt;0 pt"
    cboCost.ColumnWidths = CStr(cboCost.Width - 15) & " pt;0 pt"
    fraGrid.Height = gridBottom - fraGrid.Top
    scrStock.Move fraGrid.Left + gridAreaWidth - scrStock.Width, fraGrid.Top, scrStock.Width, fraGrid.Height
    pageRows = Application.Min(GridRows, Application.Max(1, Fix((fraGrid.Height - 3) / rowHeight)))
End Sub

Private Sub ClearGrid()
    Dim row As Long, col As Long
    For row = 1 To GridRows
        For col = 1 To GridColumns
            If Not cells(row, col) Is Nothing Then cells(row, col).Cell.Visible = False
        Next
    Next
    fraGrid.Caption = ""
End Sub

Private Sub RenderColumns()
    Dim shown As Long, slot As Long, column As Long, totalWeight As Single, x As Single, header As Object
    Dim weight(1 To GridColumns) As Single, name As String, gridWidth As Single
    shown = visibleColumns.Count
    ConfigureStockScroll
    fraGrid.Width = gridAreaWidth
    If scrStock.Visible Then
        fraGrid.Width = gridAreaWidth - scrStock.Width - 4
        scrStock.ZOrder 0
    End If
    gridWidth = fraGrid.Width - 2
    For slot = 1 To shown
        name = table.ListColumns(visibleColumns(slot)).Name
        Select Case name
            Case "Ürün", "Personel", "Teslim alan", "Windows hesabı", "Yetkiler"
                weight(slot) = 1.85
            Case "Tarih", "Alış tarihi", "Stok kodu", "İşlem"
                weight(slot) = 1.28
            Case "Açıklama", "Tedarikçi", "Belge numarası", "Kaynak fiş"
                weight(slot) = 0.83
            Case "No", "Fiş no", "Birim", "Renk", "Beden", "Ay", "Fark", "Rol", "Miktar"
                weight(slot) = 0.68
            Case Else
                weight(slot) = 1
        End Select
        totalWeight = totalWeight + weight(slot)
    Next
    x = fraGrid.Left + 1
    For slot = 1 To GridColumns
        Set header = heads(slot).Header
        header.Visible = slot <= shown
        If slot <= shown Then
            column = visibleColumns(slot)
            columnWidths(slot) = gridWidth * weight(slot) / totalWeight
            header.Left = x
            header.Height = IIf(SarfViewKey = "stock", 31, 39)
            header.Top = fraGrid.Top - header.Height - 1
            header.Width = columnWidths(slot) - 1
            header.Caption = GridHeading(table.ListColumns(column).Name)
            header.ControlTipText = table.ListColumns(column).Name & " - filtrele"
            header.BackColor = RGB(225, 237, 232)
            If filters.Exists(CStr(column)) Then header.BackColor = RGB(179, 221, 202)
            x = x + columnWidths(slot)
        End If
    Next
    RenderGridPage
End Sub

Private Sub ConfigureStockScroll()
    If SarfViewKey <> "stock" Then Exit Sub
    Dim lastStart As Long
    lastStart = Application.Max(0, filteredRows.Count - pageRows)
    firstRow = Application.Min(firstRow, lastStart + 1)
    stockScrollUpdating = True
    scrStock.Visible = lastStart > 0
    scrStock.Min = 0
    scrStock.Max = Application.Max(1, Application.Min(lastStart, 32767))
    scrStock.SmallChange = 1
    scrStock.LargeChange = pageRows
    scrStock.Value = 0
    If lastStart > 0 Then scrStock.Value = CLng((firstRow - 1) * CDbl(scrStock.Max) / lastStart)
    stockScrollUpdating = False
    cmdClear.Visible = SarfManualFiltered
End Sub

Private Sub scrStock_Change()
    ScrollStock
End Sub

Private Sub scrStock_Scroll()
    ScrollStock
End Sub

Private Sub ScrollStock()
    If stockScrollUpdating Or SarfViewKey <> "stock" Or table Is Nothing Then Exit Sub
    Dim lastStart As Long, targetRow As Long
    lastStart = Application.Max(0, filteredRows.Count - pageRows)
    targetRow = 1 + CLng(scrStock.Value * CDbl(lastStart) / scrStock.Max)
    If targetRow = firstRow Then Exit Sub
    firstRow = targetRow
    RenderGridPage
End Sub

Private Function GridHeading(ByVal title As String) As String
    Select Case title
        Case "Fiş no": GridHeading = "Fiş"
        Case "Belge numarası": GridHeading = "Belge"
        Case "Kaynak fiş": GridHeading = "Kaynak"
        Case "Stok değeri farkı": GridHeading = "Değer farkı"
        Case "Beden / numara": GridHeading = "Beden / No"
        Case Else: GridHeading = title
    End Select
End Function

Private Function GridCell(ByVal row As Long, ByVal col As Long) As Object
    Dim binding As cSarfGridCell
    If cells(row, col) Is Nothing Then
        Set binding = New cSarfGridCell
        Set binding.Cell = fraGrid.Controls.Add("Forms.Label.1", "cell" & row & "_" & col)
        Set binding.Owner = Me
        binding.RowSlot = row
        With binding.Cell
            .Font.Name = "Aptos"
            .Font.Size = 10
            .WordWrap = True
        End With
        Set cells(row, col) = binding
    End If
    cells(row, col).Cell.Top = 3 + (row - 1) * rowHeight
    cells(row, col).Cell.Height = rowHeight - 4
    Set GridCell = cells(row, col).Cell
End Function

Private Sub RenderGridPage()
    Dim row As Long, col As Long, shown As Long, position As Long, rowIndex As Long
    Dim fill As Long, x As Single, value As Variant, title As String
    If table Is Nothing Then Exit Sub
    shown = visibleColumns.Count
    ClearGrid
    For row = 1 To pageRows
        position = firstRow + row - 1
        If position > filteredRows.Count Then Exit For
        rowIndex = filteredRows(position)
        fill = IIf(position Mod 2 = 0, RGB(239, 246, 243), vbWhite)
        If position = selectedRow Then fill = RGB(189, 226, 211)
        x = 1
        For col = 1 To shown
            value = matrix(rowIndex, visibleColumns(col))
            title = table.ListColumns(visibleColumns(col)).Name
            With GridCell(row, col)
                .Left = x
                .Width = columnWidths(col) - 1
                .BackColor = fill
                .WordWrap = title <> "Tarih" And title <> "Alış tarihi" And title <> "Stok kodu" And VarType(value) <> vbDouble
                .Caption = CStr(value)
                .ControlTipText = CStr(value)
                .TextAlign = fmTextAlignLeft
                If VarType(value) = vbDouble Then .TextAlign = fmTextAlignRight
                .Visible = True
            End With
            x = x + columnWidths(col)
        Next
    Next
    If filteredRows.Count = 0 Then
        lblRows.Caption = "Eşleşen kayıt yok"
    Else
        lblRows.Caption = firstRow & "-" & Application.Min(firstRow + pageRows - 1, filteredRows.Count) _
                          & " / " & filteredRows.Count & " kayıt"
    End If
    cmdRowsBack.Enabled = firstRow > 1
    cmdRowsForward.Enabled = firstRow + pageRows <= filteredRows.Count
End Sub

Public Sub ShowGridRow(ByVal slot As Long)
    If table Is Nothing Or SarfViewKey = "counts" Then Exit Sub
    Dim position As Long, rowIndex As Long, col As Variant, details As String, value As String
    position = firstRow + slot - 1
    If position < 1 Or position > filteredRows.Count Then Exit Sub
    rowIndex = filteredRows(position)
    For Each col In visibleColumns
        value = CStr(matrix(rowIndex, col))
        If Len(value) = 0 Then value = ChrW(8212)
        details = details & table.ListColumns(col).Name & vbCrLf & value & vbCrLf & vbCrLf
    Next
    frmSarfIslem.HazirlaDetay details
    SarfShowForm frmSarfIslem, 125
    Unload frmSarfIslem
End Sub

Public Sub SelectGridRow(ByVal slot As Long)
    selectedRow = firstRow + slot - 1
    If SarfOption(cboView) = "counts" Then
        OpenCount CStr(matrix(filteredRows(selectedRow), 2))
    Else
        RenderGridPage
    End If
End Sub

Private Sub cmdRowsBack_Click()
    firstRow = Application.Max(1, firstRow - pageRows)
    RenderGridPage
End Sub

Private Sub cmdRowsForward_Click()
    firstRow = firstRow + pageRows
    RenderGridPage
End Sub

Private Sub StylePage()
    Dim item As Variant, control As Object
    lblTitle.ForeColor = RGB(17, 47, 41)
    lblSession.ForeColor = RGB(68, 90, 82)
    For Each item In Array("lblTitle", "lblSession", "lblSearch", "lblDates", "lblCounts", "lblRows")
        Me.Controls(item).BackStyle = fmBackStyleTransparent
        Me.Controls(item).Font.Name = "Aptos"
    Next
    For Each item In Array("cmdClear", "cmdExcel", "cmdPdf", "cmdCsv", "cmdRowsBack", "cmdRowsForward")
        Set control = Me.Controls(item)
        control.Font.Name = "Aptos"
        control.Font.Size = 10
        control.BackColor = RGB(232, 240, 236)
        control.ForeColor = RGB(23, 62, 49)
    Next
    fraGrid.BackColor = RGB(216, 230, 223)
    For Each item In Array("lblSummary", "lblExtra")
        Set control = Me.Controls(item)
        control.Font.Name = "Aptos"
        control.Font.Size = 10
        control.ForeColor = RGB(23, 62, 49)
        control.BackColor = RGB(229, 240, 234)
    Next
End Sub

Private Sub ReleaseGrid()
    Dim row As Long, col As Long
    For col = 1 To GridColumns
        If Not heads(col) Is Nothing Then Set heads(col).Owner = Nothing
        Set heads(col) = Nothing
    Next
    For row = 1 To GridRows
        For col = 1 To GridColumns
            If Not cells(row, col) Is Nothing Then Set cells(row, col).Owner = Nothing
            Set cells(row, col) = Nothing
        Next
    Next
End Sub

Private Sub RenderSummary()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Liste")
    If SarfViewKey = "report" Then
        lblSummary.Caption = ws.Range("A4").Value2 & "   |   " & ws.Range("A5").Value2 & vbCrLf & ws.Range("A6").Value2
        lblExtra.Caption = ws.Range("A7").Value2 & "    " & ws.Range("D5").Value2 & "    " & ws.Range("E7").Value2
    Else
        lblSummary.Caption = ""
        lblExtra.Caption = ""
        If cboQty.Visible Then lblSummary.Caption = cboQty.Text & ": " & ws.Cells(6, QtyColumn).Value2
        If cboCost.Visible Then lblExtra.Caption = cboCost.Text & ": " & ws.Cells(6, CostColumn).Value2
    End If
    If SarfViewKey = "counts" Then
        lblExtra.Visible = True
        lblExtra.Caption = "Bir sayım seçin veya yeni sayım oluşturun. Ürünleri açmak için sayım adına tıklayabilirsiniz."
    ElseIf SarfViewKey = "count-lines" Then
        lblExtra.Visible = True
        lblExtra.Caption = "Excel yükleme sayımı kaydeder. Stoğu güncellemek için Sayımı stoğa işle düğmesini kullanın."
        If Not cmdCountUpload.Enabled And Not cmdCountPost.Enabled Then
            lblExtra.Caption = "Bu sayımın ürünleri ve kayıtlı sonuçları görüntüleniyor."
        End If
    End If
    lblSummary.Visible = lblSummary.Caption <> ""
    lblExtra.Visible = lblExtra.Caption <> ""
End Sub

Public Sub FilterColumn(ByVal position As Long)
    On Error GoTo Failed
    If table Is Nothing Then Exit Sub
    If position > visibleColumns.Count Then Exit Sub
    Dim options As Object, current As Object, r As Long, col As Long, key As String
    col = visibleColumns(position)
    If fraColumnFilter.Visible And filterColumnId = col Then
        CloseColumnFilter
        Exit Sub
    End If
    Set options = CreateObject("Scripting.Dictionary")
    If IsArray(matrix) Then
        For r = 1 To UBound(matrix, 1)
            If Matches(r, col) Then
                key = CStr(matrix(r, col))
                options(key) = IIf(key = "", "(Boş)", key)
            End If
        Next
    End If
    If filters.Exists(CStr(col)) Then
        Set current = filters(CStr(col))
    End If
    filterColumnId = col
    Set filterChoices = options
    Set filterMarked = CreateObject("Scripting.Dictionary")
    If current Is Nothing Then Set current = options
    Dim value As Variant, header As Object
    For Each value In current.Keys
        filterMarked(value) = True
    Next
    lblFilterTitle.Caption = table.ListColumns(col).Name
    filterUpdating = True
    txtFind.Text = ""
    lstValues.ListStyle = 1
    lstValues.MultiSelect = 1
    RefreshFilterChoices
    Set header = heads(position).Header
    fraColumnFilter.Left = Application.Min(header.Left, fraGrid.Left + gridAreaWidth - fraColumnFilter.Width)
    fraColumnFilter.Top = header.Top + header.Height + 1
    fraColumnFilter.Visible = True
    fraColumnFilter.ZOrder 0
    If Me.Visible Then txtFind.SetFocus
    Exit Sub
Failed:
    CloseColumnFilter
    SarfHata
End Sub

Private Sub CloseColumnFilter()
    fraColumnFilter.Visible = False
    Set filterChoices = Nothing
    Set filterMarked = Nothing
    filterColumnId = 0
    filterUpdating = False
End Sub

Private Sub cmdFilterClose_Click()
    CloseColumnFilter
End Sub

Private Sub cmdReset_Click()
    ApplyColumnFilter Nothing
End Sub

Private Sub cmdApply_Click()
    Dim selection As Object
    If filterMarked Is Nothing Then Exit Sub
    If filterMarked.Count > 0 And Not AllFilterChoicesMarked() Then Set selection = filterMarked
    ApplyColumnFilter selection
End Sub

Private Sub ApplyColumnFilter(ByVal selection As Object)
    If filterColumnId = 0 Then Exit Sub
    Dim key As String
    key = CStr(filterColumnId)
    If filters.Exists(key) Then filters.Remove key
    If Not selection Is Nothing Then filters.Add key, selection
    CloseColumnFilter
    RenderRows
End Sub

Private Sub RefreshFilterChoices()
    Dim key As Variant
    filterUpdating = True
    lstValues.Clear
    For Each key In filterChoices.Keys
        If InStr(1, filterChoices(key), txtFind.Text, vbTextCompare) > 0 Or txtFind.Text = "" Then
            lstValues.AddItem filterChoices(key)
            lstValues.List(lstValues.ListCount - 1, 1) = key
            lstValues.Selected(lstValues.ListCount - 1) = filterMarked.Exists(key)
        End If
    Next
    chkAll.Value = filterChoices.Count > 0 And AllFilterChoicesMarked()
    filterUpdating = False
End Sub

Private Sub lstValues_Change()
    If filterUpdating Then Exit Sub
    Dim i As Long, key As String
    For i = 0 To lstValues.ListCount - 1
        key = lstValues.List(i, 1)
        If lstValues.Selected(i) Then
            filterMarked(key) = True
        ElseIf filterMarked.Exists(key) Then
            filterMarked.Remove key
        End If
    Next
    filterUpdating = True
    chkAll.Value = filterChoices.Count > 0 And AllFilterChoicesMarked()
    filterUpdating = False
End Sub

Private Sub txtFind_Change()
    If Not filterChoices Is Nothing Then RefreshFilterChoices
End Sub

Private Sub chkAll_Click()
    If filterUpdating Or filterChoices Is Nothing Then Exit Sub
    Dim key As Variant
    filterMarked.RemoveAll
    If chkAll.Value Then
        For Each key In filterChoices.Keys
            filterMarked(key) = True
        Next
    End If
    RefreshFilterChoices
End Sub

Private Function AllFilterChoicesMarked() As Boolean
    Dim key As Variant
    If filterMarked.Count <> filterChoices.Count Then Exit Function
    For Each key In filterChoices.Keys
        If Not filterMarked.Exists(key) Then Exit Function
    Next
    AllFilterChoicesMarked = True
End Function

Private Sub OpenAction(ByVal action As String, Optional ByVal record As String = "")
    On Error GoTo Failed
    If action = "" Then Exit Sub
    frmSarfIslem.Hazirla action, record
    SarfShowForm frmSarfIslem, 125
    Unload frmSarfIslem
    Reload
    Exit Sub
Failed:
    Unload frmSarfIslem
    SarfHata
End Sub

Private Sub Reload()
    Set SarfMenu = SarfClient.Xml("GET", "/menu")
    changing = True
    RefreshCounts
    changing = False
    LoadPage
End Sub

Private Sub cmdAction_Click()
    Dim key As String, record As String
    key = SarfOption(cboAction)
    If Left$(key, 6) = "count-" And key <> "count-new" Then record = SarfOption(cboCount)
    OpenAction key, record
End Sub

Private Sub OpenCount(ByVal countId As String)
    Dim i As Long
    For i = 0 To cboCount.ListCount - 1
        If Split(cboCount.List(i, 1), ":")(0) = countId Then
            changing = True
            cboCount.ListIndex = i
            cboView.ListIndex = 1
            changing = False
            LoadPage
            Exit Sub
        End If
    Next
End Sub

Private Sub cmdCountNew_Click()
    On Error GoTo Failed
    Dim countId As String, entryMode As String
    frmSarfIslem.Hazirla "count-new"
    SarfShowForm frmSarfIslem, 125
    countId = frmSarfIslem.SavedCountId
    entryMode = frmSarfIslem.CountEntryMode
    Unload frmSarfIslem
    If countId = "" Then Exit Sub
    Reload
    OpenCount countId
    If entryMode = "manual" Then
        cmdCountManual_Click
    Else
        FileAction "upload"
    End If
    Exit Sub
Failed:
    Unload frmSarfIslem
    SarfHata
End Sub

Private Sub cmdCountManual_Click()
    OpenAction "count-line", SarfOption(cboCount)
End Sub

Private Sub cmdCountOpen_Click()
    OpenCount Split(SarfOption(cboCount) & ":", ":")(0)
End Sub

Private Sub cmdCountTemplate_Click()
    FileAction "template"
End Sub

Private Sub cmdCountUpload_Click()
    FileAction "upload"
End Sub

Private Sub cmdCountPost_Click()
    On Error GoTo Failed
    Dim response As Object, payload As String
    If MsgBox("Kayıtlı sayım farkları stoğa Sayım düzeltmesi olarak işlenecek. Devam edilsin mi?", _
              vbQuestion Or vbYesNo Or vbDefaultButton2, "Sayımı stoğa işle") <> vbYes Then Exit Sub
    payload = "action=count-approve" & SarfParam("count_key", SarfOption(cboCount))
    Set response = SarfClient.Xml("POST", "/apply", payload, "application/x-www-form-urlencoded")
    Reload
    MsgBox "Sayım farkları stoğa işlendi.", vbInformation, "Sayım"
    Exit Sub
Failed:
    SarfHata
End Sub

Private Sub cmdCountEdit_Click()
    If selectedRow = 0 Then
        MsgBox "Önce sayım satırını seçin.", vbInformation
        Exit Sub
    End If
    OpenAction "count-line", CountKey & ":" & CStr(matrix(filteredRows(selectedRow), CountVariantColumn))
End Sub

Private Sub cmdProduct_Click()
    OpenAction "product"
End Sub

Private Sub FileAction(ByVal action As String)
    On Error GoTo Failed
    SarfDosya action, SarfOption(cboCount)
    If action = "upload" Then Reload
    Exit Sub
Failed:
    SarfHata
End Sub

Private Sub cmdTemplate_Click()
    FileAction "template"
End Sub

Private Sub cmdUpload_Click()
    FileAction "upload"
End Sub

Private Sub cboCount_Change()
    If Not changing And currentSection = "Counts" Then
        OpenCount Split(SarfOption(cboCount) & ":", ":")(0)
    End If
End Sub

Private Sub cboView_Change()
    If Not changing Then LoadPage
End Sub

Private Sub cmdDates_Click()
    LoadPage
End Sub

Private Sub txtSearch_Change()
    If Not changing Then RenderRows
End Sub

Private Sub cmdClear_Click()
    If table Is Nothing Then Exit Sub
    filters.RemoveAll
    changing = True
    txtSearch.Text = ""
    changing = False
    RenderRows
End Sub

Private Sub cboQty_Change()
    If changing Or QtyColumn = 0 Then Exit Sub
    table.Parent.Cells(5, QtyColumn).Value2 = SarfOption(cboQty)
    SarfToplam
    RenderSummary
End Sub

Private Sub cboCost_Change()
    If changing Or CostColumn = 0 Then Exit Sub
    table.Parent.Cells(5, CostColumn).Value2 = SarfOption(cboCost)
    SarfToplam
    RenderSummary
End Sub

Private Sub cmdExcel_Click()
    If Not table Is Nothing Then SarfExcel
End Sub

Private Sub cmdPdf_Click()
    If Not table Is Nothing Then SarfPDF
End Sub

Private Sub cmdCsv_Click()
    If Not table Is Nothing Then SarfCSV
End Sub

Private Sub cmdLogout_Click()
    Me.Hide
End Sub

Private Sub cmdSwitchUser_Click()
    SwitchUserRequested = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode <> 0 Then ReleaseGrid
    If CloseMode = 0 Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub cmdStock_Click()
    Navigate "Stock"
End Sub

Private Sub cmdMoves_Click()
    Navigate "Moves"
End Sub

Private Sub cmdProducts_Click()
    Navigate "Products"
End Sub

Private Sub cmdPeople_Click()
    Navigate "People"
End Sub

Private Sub cmdReports_Click()
    Navigate "Reports"
End Sub

Private Sub cmdCounts_Click()
    Navigate "Counts"
End Sub

Private Sub cmdImport_Click()
    Navigate "Import"
End Sub

Private Sub cmdAdmin_Click()
    Navigate "Admin"
End Sub

Private Sub cmdPeriod_Click()
    Navigate "Period"
End Sub
