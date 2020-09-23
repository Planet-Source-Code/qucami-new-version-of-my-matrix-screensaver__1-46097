Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Type udtBox
    x As Long
    y As Long
    xCurr As Long
    yCurr As Long
    xMax As Long
    yMax As Long
    Text As String
    LastTickCount As Long
    BoxOpenFor As Long
    BoxOpened As Long
    TimeTilNextBox As Long
End Type
Public Type CharTrail
    x As Integer
    y As Integer
    TailLength As Integer
    Speed As Long
    LastTickCount As Long
End Type
Public Type CircTrail
    x As Integer
    y As Integer
    r As Long
    LastTickCount As Long
    NextCircle As Long
End Type
Public Type SS
    LastTickCount As Long
    NextSS As Long
End Type
Public Const Pi = 3.1415926535898
Public Const iCharHeight As Integer = 22
Public Const iCharWidth As Integer = 15
Function GimmeX(ByVal aIn As Single, lIn As Long) As Long
    GimmeX = Sin(aIn * (Pi / 180)) * lIn

End Function
Function GimmeY(ByVal aIn As Single, lIn As Long) As Long
    GimmeY = Cos(aIn * (Pi / 180)) * lIn
End Function
Sub InitRegistry()
Dim iMinCols As Integer
    iMinCols = (Screen.Width \ Screen.TwipsPerPixelX) \ iCharWidth
    If Len(GetSetting(App.Title, "Settings", "NumCols")) = 0 Then
        SaveSetting App.Title, "Settings", "NumCols", iMinCols
    End If
    If Len(GetSetting(App.Title, "Settings", "Circles")) = 0 Then
        SaveSetting App.Title, "Settings", "Circles", vbChecked
    End If
    If Len(GetSetting(App.Title, "Settings", "Twitch")) = 0 Then
        SaveSetting App.Title, "Settings", "Twitch", vbChecked
    End If
    If Len(GetSetting(App.Title, "Settings", "Time")) = 0 Then
        SaveSetting App.Title, "Settings", "Time", vbChecked
    End If
    If Len(GetSetting(App.Title, "Settings", "Zoom")) = 0 Then
        SaveSetting App.Title, "Settings", "Zoom", vbChecked
    End If
End Sub
Public Function AdjustBrightness(RGB_In As Long, ShiftPercentage As Integer, Optional GotoWhite As Boolean = False) As Long
Dim lColor As Long
Dim r As Single, G As Single, B As Single

    lColor = RGB_In
    r = lColor Mod &H100
    lColor = lColor \ &H100
    G = lColor Mod &H100
    lColor = lColor \ &H100
    B = lColor Mod &H100

    r = r + ((r / 100) * ShiftPercentage)
    G = G + ((G / 100) * ShiftPercentage)
    B = B + ((B / 100) * ShiftPercentage)
    
    If r > 255 Or G > 255 Or B > 255 Then
        If GotoWhite Then
            If r > 255 Then r = 255
            If G > 255 Then G = 255
            If B > 255 Then B = 255
            AdjustBrightness = RGB(r, G, B)
        Else
            AdjustBrightness = RGB_In
        End If
    ElseIf r < 0 Or G < 0 Or B < 0 Then
        AdjustBrightness = RGB_In
    Else
        AdjustBrightness = RGB(r, G, B)
    End If
End Function
