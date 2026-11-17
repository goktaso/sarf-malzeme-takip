param([switch]$ShowExcel)

$ErrorActionPreference = 'Stop'
$workbookPath = Join-Path $PSScriptRoot 'Sarf-Malzeme-Takip-Sutun-Filtresi.xlsm'
$project = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$startup = Join-Path $project 'excel_vba/Start-LocalService.ps1'
$localData = Join-Path $project 'data/ExcelVbaPilot'
$startupLog = Join-Path $localData 'desktop-startup.log'
$excel = $null
$book = $null
$job = $null
try {
    Set-Content -LiteralPath $startupLog -Value 'Starting local application'
    if (!(Test-Path -LiteralPath $workbookPath)) { throw "Uygulama dosyasi bulunamadi. $workbookPath" }
    $windowsPowerShell = Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe'
    & $windowsPowerShell -NoProfile -NonInteractive -File $startup
    if ($LASTEXITCODE -ne 0) { throw 'Yerel uygulama servisi baslatilamadi.' }
    Add-Type -AssemblyName System.Security
    $request = [System.Net.WebRequest]::Create('http://127.0.0.1:8768/local-proof')
    $request.Proxy = $null
    $request.Timeout = 2000
    $response = $request.GetResponse()
    try {
        $reader = New-Object IO.StreamReader($response.GetResponseStream())
        $encrypted = [Convert]::FromBase64String($reader.ReadToEnd())
    } finally {
        $response.Dispose()
    }
    $plain = [Security.Cryptography.ProtectedData]::Unprotect($encrypted, $null, 'CurrentUser')
    $proof = [Text.Encoding]::ASCII.GetString($plain)
    [Array]::Clear($plain, 0, $plain.Length)

    $request = [System.Net.WebRequest]::Create('http://127.0.0.1:8768/local-login')
    $request.Proxy = $null
    $request.Timeout = 3000
    $request.Method = 'POST'
    $request.ContentType = 'text/plain'
    $body = [Text.Encoding]::ASCII.GetBytes($proof)
    $request.ContentLength = $body.Length
    $stream = $request.GetRequestStream()
    try { $stream.Write($body, 0, $body.Length) } finally {
        $stream.Dispose()
        [Array]::Clear($body, 0, $body.Length)
        $proof = $null
    }
    $response = $request.GetResponse()
    try {
        $document = New-Object System.Xml.XmlDocument
        $document.Load($response.GetResponseStream())
        $session = $document.DocumentElement.GetAttribute('session')
    } finally { $response.Dispose() }
    if (!$session) { throw 'Yerel Windows oturumu acilamadi.' }
    $bridge = 'sarf-session-' + [Guid]::NewGuid().ToString('N')
    $worker = Join-Path $project 'excel_vba/LocalRequestWorker.ps1'
    $job = Start-Job -FilePath $worker -ArgumentList $bridge,$session
    Add-Content -LiteralPath $startupLog -Value 'Local Windows session opened'

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.AutomationSecurity = 2
    $excel.EnableEvents = $false
    $book = $excel.Workbooks.Open($workbookPath, 0, $true)
    $excel.EnableEvents = $true
    if ($ShowExcel) { $excel.Visible = $true }
    Add-Content -LiteralPath $startupLog -Value 'Workbook opened; starting main form'
    $null = $excel.Run("'" + $book.Name + "'!SarfBaslaticiOturumu", $bridge)
} catch {
    Add-Content -LiteralPath $startupLog -Value $_.Exception.Message -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.Windows.Forms
    $null = [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Sarf Malzeme Takip')
} finally {
    $proof = $null
    $session = $null
    if ($excel) {
        $excel.DisplayAlerts = $false
        if ($book) { $book.Close($false) }
        $excel.Quit()
    }
    if ($job) { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue }
}
