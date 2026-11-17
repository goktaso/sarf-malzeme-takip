param([string]$PipeName, [string]$Session)

$ErrorActionPreference = 'Stop'
$identity = [Security.Principal.WindowsIdentity]::GetCurrent().User
$security = New-Object IO.Pipes.PipeSecurity
$security.SetAccessRuleProtection($true, $false)
$security.AddAccessRule((New-Object IO.Pipes.PipeAccessRule($identity, 'FullControl', 'Allow')))

while ($true) {
    $pipe = New-Object IO.Pipes.NamedPipeServerStream($PipeName, 'InOut', 1, 'Byte', 'None', 0, 0, $security)
    try {
    $pipe.WaitForConnection()
    $inputReader = New-Object IO.BinaryReader($pipe)
    $outputWriter = New-Object IO.BinaryWriter($pipe)
    $output = New-Object System.Xml.XmlDocument
    $root = $output.AppendChild($output.CreateElement('response'))
    try {
        $reader = New-Object System.Xml.XmlDocument
        $reader.XmlResolver = $null
        $length = $inputReader.ReadInt32()
        if ($length -lt 1 -or $length -gt 8388608) { throw 'Gecersiz yerel istek boyutu.' }
        $reader.LoadXml([Text.Encoding]::UTF8.GetString($inputReader.ReadBytes($length)))
        $message = $reader.DocumentElement
        $resource = $message.GetAttribute('resource')
        $method = $message.GetAttribute('method')
        if ($method -notin 'GET','POST' -or $resource -notmatch '^/[a-z-]+(?:\?[^\r\n]*)?$') {
            throw 'Gecersiz yerel istek.'
        }
        $http = [System.Net.WebRequest]::Create('http://127.0.0.1:8768' + $resource)
        $http.Proxy = $null
        $http.AllowAutoRedirect = $false
        $http.Timeout = 25000
        $http.ReadWriteTimeout = 25000
        $http.Method = $method
        $http.Headers['X-Sarf-Session'] = $Session
        if ($method -eq 'POST') {
            $bytes = [Convert]::FromBase64String($message.InnerText)
            $http.ContentLength = $bytes.Length
            $mime = $message.GetAttribute('mime')
            if ($mime) { $http.ContentType = $mime }
            if ($bytes.Length) {
                $send = $http.GetRequestStream()
                try { $send.Write($bytes, 0, $bytes.Length) } finally { $send.Dispose() }
            }
        }
        try { $response = $http.GetResponse() } catch [System.Net.WebException] {
            if (!$_.Exception.Response) { throw }
            $response = $_.Exception.Response
        }
        try {
            $memory = New-Object IO.MemoryStream
            try {
                $response.GetResponseStream().CopyTo($memory)
                $root.InnerText = [Convert]::ToBase64String($memory.ToArray())
            } finally { $memory.Dispose() }
            $root.SetAttribute('status', [string][int]$response.StatusCode)
        } finally { $response.Dispose() }
    } catch {
        $root.SetAttribute('error', $_.Exception.Message)
    }
    $encoded = [Text.Encoding]::UTF8.GetBytes($output.OuterXml)
    $outputWriter.Write([int]$encoded.Length)
    $outputWriter.Write([byte[]]$encoded)
    $outputWriter.Flush()
    } catch [IO.IOException] {
        # The client closed its workbook before reading the response.
    } finally {
        $pipe.Dispose()
    }
}
