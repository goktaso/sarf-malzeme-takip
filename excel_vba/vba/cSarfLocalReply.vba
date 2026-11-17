Option Explicit

Public Status As Long
Public ResponseBody As Variant

Public Property Get ResponseText() As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write ResponseBody
    stream.Position = 0
    stream.Type = 2
    stream.Charset = "utf-8"
    ResponseText = stream.ReadText
    stream.Close
End Property
