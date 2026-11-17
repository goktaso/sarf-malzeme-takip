Option Explicit

Private Type ScreenRect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type MonitorInfo
    Size As Long
    Bounds As ScreenRect
    Work As ScreenRect
    Flags As Long
End Type

Private Declare PtrSafe Function MonitorFromWindow Lib "user32" (ByVal windowHandle As LongPtr, ByVal flags As Long) As LongPtr
Private Declare PtrSafe Function GetMonitorInfo Lib "user32" Alias "GetMonitorInfoW" (ByVal monitor As LongPtr, ByRef info As MonitorInfo) As Long
Private Declare PtrSafe Function GetDpiForWindow Lib "user32" (ByVal windowHandle As LongPtr) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowW" (ByVal className As LongPtr, ByVal caption As LongPtr) As LongPtr
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal windowHandle As LongPtr, ByVal command As Long) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal windowHandle As LongPtr) As Long
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)

Public Sub SarfShowDesktop(ByVal dialog As frmSarfAna)
    Dim initialCaption As String, uniqueCaption As String, windowHandle As LongPtr
    SarfFitForm dialog, 150
    initialCaption = dialog.Caption
    uniqueCaption = initialCaption & " #" & GetCurrentProcessId()
    dialog.Caption = uniqueCaption
    dialog.Show vbModeless
    windowHandle = FindWindow(0, StrPtr(uniqueCaption))
    dialog.Caption = initialCaption
    If windowHandle <> 0 Then
        ShowWindow windowHandle, 5
        SetForegroundWindow windowHandle
    End If
    Do While dialog.Visible
        DoEvents
        Sleep 20
    Loop
End Sub

Public Sub SarfShowForm(ByVal dialog As Object, Optional ByVal maxZoom As Long = 100)
    SarfFitForm dialog, maxZoom
    dialog.Show
End Sub

Public Sub SarfFitForm(ByVal dialog As Object, ByVal maxZoom As Long)
    Dim screen As MonitorInfo, windowHandle As LongPtr, monitor As LongPtr, dpi As Long
    Dim desktopWidth As Double, desktopHeight As Double, pointsPerPixel As Double
    windowHandle = Application.hWnd
    monitor = MonitorFromWindow(windowHandle, 2)
    screen.Size = LenB(screen)
    If monitor = 0 Then Exit Sub
    If GetMonitorInfo(monitor, screen) = 0 Then Exit Sub
    On Error Resume Next
    dpi = GetDpiForWindow(windowHandle)
    If dpi <= 96 And (screen.Flags And 1) = 1 Then
        dpi = CLng(CreateObject("WScript.Shell").RegRead("HKCU\Control Panel\Desktop\WindowMetrics\AppliedDPI"))
    End If
    On Error GoTo 0
    If dpi = 0 Then dpi = 96
    pointsPerPixel = 72# / dpi
    desktopWidth = (screen.Work.Right - screen.Work.Left) * pointsPerPixel
    desktopHeight = (screen.Work.Bottom - screen.Work.Top) * pointsPerPixel
    SarfSizeForm dialog, desktopWidth, desktopHeight, maxZoom
    If TypeOf dialog Is frmSarfAna Then
        dialog.Width = desktopWidth - 24
        dialog.Height = desktopHeight - 24
        dialog.FitLayout
    End If
    dialog.StartUpPosition = 0
    dialog.Left = screen.Work.Left * pointsPerPixel + (desktopWidth - dialog.Width) / 2
    dialog.Top = screen.Work.Top * pointsPerPixel + (desktopHeight - dialog.Height) / 2
End Sub

Public Sub SarfSizeForm(ByVal dialog As Object, ByVal desktopWidth As Double, ByVal desktopHeight As Double, ByVal maxZoom As Long)
    Const margin As Long = 24
    Dim insideWidth As Double, insideHeight As Double, chromeWidth As Double, chromeHeight As Double
    Dim zoom As Long, previousScale As Double
    previousScale = dialog.Zoom / 100#
    dialog.Zoom = 100
    insideWidth = dialog.InsideWidth / previousScale
    insideHeight = dialog.InsideHeight / previousScale
    chromeWidth = dialog.Width - dialog.InsideWidth
    chromeHeight = dialog.Height - dialog.InsideHeight
    zoom = Fix(Application.Min(maxZoom, _
                (desktopWidth - margin - chromeWidth) * 100# / insideWidth, _
                (desktopHeight - margin - chromeHeight) * 100# / insideHeight))
    zoom = Application.Max(10, zoom)
    dialog.Width = insideWidth * zoom / 100# + chromeWidth
    dialog.Height = insideHeight * zoom / 100# + chromeHeight
    dialog.Zoom = zoom
End Sub
