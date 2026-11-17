param(
    [string]$OutputPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'outputs/excel-vba/Sarf-Malzeme-Takip-Sutun-Filtresi.xlsm'),
    [string]$Endpoint = 'http://127.0.0.1:8768'
)

$ErrorActionPreference = 'Stop'

function Add-FormControl {
    param($Design, [string]$Type, [string]$Name, [string]$Caption, [int]$X, [int]$Y, [int]$Width, [int]$Height)
    $control = $Design.Controls.Add("Forms.$Type.1", $Name, $true)
    $control.Left = $X
    $control.Top = $Y
    $control.Width = $Width
    $control.Height = $Height
    if ($Type -in 'Label','CommandButton','CheckBox') { $control.Caption = $Caption }
    if ($Type -eq 'ComboBox') {
        $control.ColumnCount = 2
        $control.ColumnWidths = "$($Width - 16) pt;0 pt"
        $control.Style = 2
    }
    if ($Type -eq 'CommandButton') {
        $control.BackColor = 6013974
        $control.ForeColor = 16777215
    }
    return $control
}

function Add-WorksheetButton {
    param($Sheet, [string]$Name, [string]$Caption, [string]$Action, [double]$X, [double]$Y, [int]$Width)
    $shape = $Sheet.Shapes.AddShape(5, $X, $Y, $Width, 24)
    $shape.Name = $Name
    $shape.OnAction = $Action
    $shape.Fill.ForeColor.RGB = 4213782
    $shape.Line.Visible = 0
    $shape.TextFrame.Characters().Text = $Caption
    $shape.TextFrame.Characters().Font.Color = 16777215
    $shape.TextFrame.Characters().Font.Bold = $true
    $shape.TextFrame.HorizontalAlignment = -4108
    $shape.TextFrame.VerticalAlignment = -4108
    $shape.Placement = 3
}

function Add-Code {
    param($Project, [int]$Type, [string]$Name)
    $part = $Project.VBComponents.Add($Type)
    $part.Name = $Name
    $part.CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $PSScriptRoot "vba/$Name.vba")))
    return $part
}

$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    $book = $excel.Workbooks.Add(-4167)
    $homeSheet = $book.Worksheets.Item(1)
    $homeSheet.Name = 'Ana'
    $list = $book.Worksheets.Add([Type]::Missing, $homeSheet)
    $list.Name = 'Liste'
    $project = $book.VBProject
    $project.Name = 'SarfMalzemeTakip'
    $null = Add-Code $project 1 'modSarf'
    $null = Add-Code $project 1 'modSarfScreen'
    $null = Add-Code $project 1 'modListe'
    $null = Add-Code $project 2 'cSarfClient'
    $null = Add-Code $project 2 'cSarfLocalReply'
    $null = Add-Code $project 2 'cSarfGridCell'
    $loginForm = Add-Code $project 3 'frmSarfGiris'
    $loginForm.Properties.Item('Caption').Value = 'Sarf Malzeme Takip'
    $loginForm.Properties.Item('Width').Value = 407
    $loginForm.Properties.Item('Height').Value = 326
    $loginForm.Properties.Item('BackColor').Value = 16777215
    $login = $loginForm.Designer
    $null = Add-FormControl $login Label lblBanner '' 0 0 408 75
    $null = Add-FormControl $login Label lblTitle 'Sarf Malzeme Takip' 25 17 345 28
    $null = Add-FormControl $login Label lblSubtitle 'Windows ile giriş' 26 48 340 20
    $null = Add-FormControl $login Label lblAccount 'Windows hesabı' 26 94 348 18
    $null = Add-FormControl $login TextBox txtAccount '' 26 115 348 27
    $null = Add-FormControl $login Label lblPassword 'Windows parolası' 26 156 348 18
    $null = Add-FormControl $login TextBox txtPassword '' 26 177 348 27
    $null = Add-FormControl $login Label lblMessage '' 26 213 348 31
    $null = Add-FormControl $login CommandButton cmdCancel 'Kapat' 171 253 91 32
    $null = Add-FormControl $login CommandButton cmdLogin 'Giriş yap' 271 253 103 32
    $dashboard = Add-Code $project 3 'frmSarfAna'
    $dashboard.Properties.Item('Caption').Value = 'Sarf Malzeme Takip'
    $dashboard.Properties.Item('Width').Value = 902
    $dashboard.Properties.Item('Height').Value = 553
    $dashboard.Properties.Item('BackColor').Value = 16185075
    $design = $dashboard.Designer
    $side = Add-FormControl $design Label lblSide '' 0 0 142 535
    $side.BackColor = 2699025
    $brand = Add-FormControl $design Label lblBrand 'Sarf Malzeme Takip' 10 24 122 28
    $brand.BackColor = 2699025
    $brand.ForeColor = 11920554
    $button = Add-FormControl $design CommandButton cmdStock 'Stok durumu' 8 64 126 32
    $button.BackColor = 2699025
    $button = Add-FormControl $design CommandButton cmdMoves 'Stok hareketleri' 8 106 126 32
    $button.BackColor = 2699025
    $button = Add-FormControl $design CommandButton cmdProducts 'Ürün kartları' 8 148 126 32
    $button.BackColor = 2699025
    $button = Add-FormControl $design CommandButton cmdPeople 'Departman ve personel' 8 190 126 32
    $button.BackColor = 2699025
    $button = Add-FormControl $design CommandButton cmdReports 'Tüketim raporları' 8 232 126 32
    $button.BackColor = 2699025
    $button = Add-FormControl $design CommandButton cmdCounts 'Sayım' 8 274 126 32
    $button.BackColor = 2699025
    $button = Add-FormControl $design CommandButton cmdImport 'Excel içe aktar' 8 316 126 32
    $button.BackColor = 2699025
    $button = Add-FormControl $design CommandButton cmdAdmin 'Kullanıcılar ve yetkiler' 8 358 126 32
    $button.BackColor = 2699025
    $button = Add-FormControl $design CommandButton cmdPeriod 'Dönem kilidi' 8 400 126 32
    $button.BackColor = 2699025
    $button = Add-FormControl $design CommandButton cmdSwitchUser 'Kullanıcı değiştir' 8 447 126 32
    $button.BackColor = 2699025
    $button = Add-FormControl $design CommandButton cmdLogout 'Kapat' 8 487 126 32
    $button.BackColor = 2699025
    $null = Add-FormControl $design Label lblTitle '' 160 15 707 26
    $null = Add-FormControl $design Label lblSession '' 160 46 707 15
    $null = Add-FormControl $design ComboBox cboView '' 192 71 168 24
    $null = Add-FormControl $design ComboBox cboAction '' 371 71 236 24
    $null = Add-FormControl $design CommandButton cmdAction 'İşlemi aç' 614 71 92 24
    $null = Add-FormControl $design Label lblSearch 'Listede ara' 192 101 240 15
    $null = Add-FormControl $design TextBox txtSearch '' 192 118 236 26
    $null = Add-FormControl $design CommandButton cmdProduct '+ Ürün kartı' 438 118 104 26
    $null = Add-FormControl $design CommandButton cmdClear 'Temizle' 550 118 84 26
    $null = Add-FormControl $design CommandButton cmdExcel 'Excel' 666 118 60 26
    $null = Add-FormControl $design CommandButton cmdPdf 'PDF' 733 118 60 26
    $null = Add-FormControl $design CommandButton cmdCsv 'CSV' 800 118 60 26
    $null = Add-FormControl $design Label lblDates 'Tarih (YYYY-AA-GG)' 192 155 106 15
    $null = Add-FormControl $design TextBox txtStart '' 303 150 95 22
    $null = Add-FormControl $design TextBox txtEnd '' 404 150 95 22
    $null = Add-FormControl $design CommandButton cmdDates 'Uygula' 505 150 66 22
    $null = Add-FormControl $design Label lblCounts 'Sayım / aktarım' 192 156 105 16
    $null = Add-FormControl $design ComboBox cboCount '' 303 151 255 24
    $null = Add-FormControl $design CommandButton cmdCountEdit 'Miktarı düzelt' 566 151 113 24
    $null = Add-FormControl $design CommandButton cmdCountOpen 'Ürünleri aç' 566 151 113 24
    $null = Add-FormControl $design CommandButton cmdCountNew 'Yeni sayım' 729 71 131 24
    $null = Add-FormControl $design CommandButton cmdCountTemplate 'Sayım şablonunu indir' 192 185 154 26
    $null = Add-FormControl $design CommandButton cmdCountUpload 'Excel sonuçlarını yükle' 358 185 161 26
    $null = Add-FormControl $design CommandButton cmdCountPost 'Sayımı stoğa işle' 531 185 146 26
    $null = Add-FormControl $design CommandButton cmdCountManual 'Elle giriş' 687 185 173 26
    $null = Add-FormControl $design CommandButton cmdTemplate 'Şablonu indir' 303 189 123 26
    $null = Add-FormControl $design CommandButton cmdUpload 'Excel içe aktar' 434 189 123 26
    $null = Add-FormControl $design ComboBox cboQty '' 583 151 133 24
    $null = Add-FormControl $design ComboBox cboCost '' 725 151 135 24
    $null = Add-FormControl $design Label lblSummary '' 192 185 666 33
    $null = Add-FormControl $design Label lblExtra '' 192 219 666 23
    $grid = Add-FormControl $design Frame fraGrid '' 160 290 707 197
    $grid.Caption = ''
    $null = Add-FormControl $design ScrollBar scrStock '' 849 290 18 197
    $null = Add-FormControl $design CommandButton cmdRowsBack 'Önceki satırlar' 396 493 110 25
    $null = Add-FormControl $design CommandButton cmdRowsForward 'Sonraki satırlar' 514 493 116 25
    $null = Add-FormControl $design Label lblRows '' 160 497 225 20
    $popup = Add-FormControl $design Frame fraColumnFilter '' 160 172 257 205
    $popup.Caption = ''
    $popup.Visible = $false
    $menuDesign = $popup
    $null = Add-FormControl $menuDesign Label lblFilterTitle '' 9 7 208 18
    $null = Add-FormControl $menuDesign CommandButton cmdFilterClose 'Kapat' 206 5 42 20
    $null = Add-FormControl $menuDesign TextBox txtFind '' 9 29 239 21
    $null = Add-FormControl $menuDesign CheckBox chkAll 'Tümünü seç' 9 54 239 18
    $null = Add-FormControl $menuDesign ListBox lstValues '' 9 75 239 94
    $null = Add-FormControl $menuDesign CommandButton cmdReset 'Filtreyi kaldır' 9 174 99 23
    $null = Add-FormControl $menuDesign CommandButton cmdApply 'Uygula' 170 174 78 23
    $popup.Controls.Item('lstValues').ColumnCount = 2
    $popup.Controls.Item('lstValues').ColumnWidths = '214 pt;0 pt'
    $operation = Add-Code $project 3 'frmSarfIslem'
    $operation.Properties.Item('Caption').Value = 'Sarf Malzeme İşlemi'
    $operation.Properties.Item('Width').Value = 566
    $operation.Properties.Item('Height').Value = 412
    $designer = $operation.Designer
    $null = Add-FormControl $designer Label lblTitle '' 15 12 532 27
    $frame = Add-FormControl $designer Frame fraFields '' 14 48 530 278
    $frame.Caption = ''
    $frame.ScrollBars = 2
    $frame.KeepScrollBarsVisible = 2
    $null = Add-FormControl $designer CommandButton cmdLoad 'Kaydı yükle' 16 344 118 28
    $null = Add-FormControl $designer CommandButton cmdManual 'Elle giriş' 145 344 130 28
    $null = Add-FormControl $designer CommandButton cmdCancel 'Vazgeç' 301 344 105 28
    $null = Add-FormControl $designer CommandButton cmdSave 'Kaydet' 419 344 109 28
    $project.VBComponents.Item($list.CodeName).CodeModule.AddFromString(
        [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'vba/ListeEvents.vba')))
    $project.VBComponents.Item($book.CodeName).CodeModule.AddFromString(
        [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'vba/WorkbookEvents.vba')))

    $homeSheet.Columns.Item('A').ColumnWidth = 4
    $homeSheet.Columns.Item('B').ColumnWidth = 52
    $homeSheet.Columns.Item('C').ColumnWidth = 25
    $homeSheet.Range('B2').Value2 = 'Sarf Malzeme Takip'
    $homeSheet.Range('B2').Font.Size = 26
    $homeSheet.Range('B2').Font.Bold = $true
    $homeSheet.Range('B2').Font.Color = 4213782
    $homeSheet.Range('B4').Value2 = 'Satın Alma Deposu · Excel masaüstü uygulaması'
    $homeSheet.Range('B6').Value2 = 'Windows servis adresi'
    $homeSheet.Range('B6').Font.Bold = $true
    $homeSheet.Range('B7').Value2 = $Endpoint
    $homeSheet.Range('B7').Interior.Color = 15792864
    $homeSheet.Range('B8').Value2 = 'Açık oturum'
    $homeSheet.Range('B8').Font.Bold = $true
    Add-WorksheetButton $homeSheet SarfLoginButton 'Giriş ekranını aç' 'SarfOturum' $homeSheet.Range('B11').Left $homeSheet.Range('B11').Top 152
    Add-WorksheetButton $homeSheet SarfLogoutButton 'Oturumu kapat' 'SarfCikis' ($homeSheet.Range('B11').Left + 164) $homeSheet.Range('B11').Top 138
    $homeSheet.Range('B14').Value2 = 'Windows ile girişten sonra listeleri ve işlem formlarını açabilirsiniz.'
    $homeSheet.Range('B15').Value2 = 'Makrolar etkin olduğunda giriş ekranı otomatik açılır.'
    $homeSheet.Range('B17').Value2 = 'Kurulum ve çalıştırma bilgileri excel_vba klasöründedir.'

    $list.Range('A1').Value2 = 'Listeyi Windows oturumundan açın'
    $list.Range('A1').Font.Size = 18
    $list.Rows.Item(3).RowHeight = 28
    Add-WorksheetButton $list SarfMenuButton 'İşlemler' 'SarfOturum' 345 3 84
    Add-WorksheetButton $list SarfRefreshButton 'Yenile' 'SarfYenile' 435 3 73
    Add-WorksheetButton $list SarfExcelButton 'Excel indir' 'SarfExcel' 514 3 82
    Add-WorksheetButton $list SarfPdfButton 'PDF indir' 'SarfPDF' 602 3 78
    Add-WorksheetButton $list SarfCsvButton 'CSV indir' 'SarfCSV' 686 3 76
    Add-WorksheetButton $list SarfSearchButton 'Ara' 'SarfAra' 243 $list.Range('C3').Top 74
    Add-WorksheetButton $list SarfAddButton '+ Ürün kartı' 'SarfYeniUrun' 325 $list.Range('C3').Top 106
    Add-WorksheetButton $list SarfCountButton 'Miktarı düzelt' 'SarfSayimMiktari' 325 $list.Range('C3').Top 118
    $list.Shapes.Item('SarfAddButton').Visible = 0
    $list.Shapes.Item('SarfCountButton').Visible = 0
    foreach ($sheet in @($list, $homeSheet)) {
        $sheet.Activate()
        $excel.ActiveWindow.DisplayGridlines = $false
    }
    $homeSheet.Activate()
    $list.Visible = 2
    $homeSheet.Range('B2').Select()
    $fullPath = [IO.Path]::GetFullPath($OutputPath)
    [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($fullPath)) | Out-Null
    $book.SaveAs($fullPath, 52)
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Sarf-Malzeme-Takip.vbs') -Destination (Join-Path ([IO.Path]::GetDirectoryName($fullPath)) 'Sarf-Malzeme-Takip.vbs')
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Sarf-Malzeme-Takip.ps1') -Destination (Join-Path ([IO.Path]::GetDirectoryName($fullPath)) 'Sarf-Malzeme-Takip.ps1')
    Write-Output $fullPath
}
finally {
    if ($book) { $book.Close($false) }
    if ($excel) { $excel.Quit() }
}

