$ErrorActionPreference = 'Stop'
$project = Split-Path $PSScriptRoot -Parent
$data = Join-Path $project 'data/ExcelVbaPilot'

function Test-LocalService {
    try {
        $request = [System.Net.WebRequest]::Create('http://127.0.0.1:8768/health')
        $request.Proxy = $null
        $request.Timeout = 1000
        $request.ReadWriteTimeout = 1000
        $reply = $request.GetResponse()
        try {
            $response = New-Object System.Xml.XmlDocument
            $response.Load($reply.GetResponseStream())
            return $response.DocumentElement.LocalName -eq 'service' -and $response.DocumentElement.GetAttribute('local_login') -eq '1'
        } finally {
            $reply.Dispose()
        }
    } catch {
        return $false
    }
}

if (Test-LocalService) { exit 0 }
if (!(Test-Path -LiteralPath (Join-Path $data 'records.json'))) {
    throw 'Yerel uygulama verisi bulunamadi.'
}
$runtime = Get-ChildItem 'C:/Program Files/WindowsApps/OpenAI.CodexPrimaryRuntime*/dependencies/python/python.exe' |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
$python = if ($runtime) { $runtime.FullName } else { (Get-Command python.exe).Source }

$log = Join-Path $data 'service-start-error.log'
$process = Start-Process -FilePath $python -WorkingDirectory $project -WindowStyle Hidden -PassThru `
    -ArgumentList '-m','excel_vba.service','--data','data/ExcelVbaPilot','--local-login' -RedirectStandardError $log
for ($attempt = 0; $attempt -lt 20; $attempt++) {
    if (Test-LocalService) { exit 0 }
    if ($process.HasExited) { break }
    Start-Sleep -Milliseconds 200
}
if (Test-LocalService) { exit 0 }
throw "Yerel uygulama servisi baslatilamadi. Gunluk: $log"
