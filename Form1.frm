VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00001000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13665
   LinkTopic       =   "Form1"
   ScaleHeight     =   708
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   911
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMagnify 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   9420
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox picTime 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   7260
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   313
      TabIndex        =   2
      Top             =   9180
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.PictureBox picTextS 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   3060
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   473
      TabIndex        =   1
      Top             =   3600
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   3720
      Top             =   4440
   End
   Begin VB.PictureBox picText 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   60
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   473
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   7095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MC As String = "abcdefghijklABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const xOffset As Integer = 4
Const yOffset As Integer = 4


Dim iNumCols As Integer
Dim iNumRows As Integer
Dim iNumChars As Integer
Dim udtTails() As CharTrail
Dim udtCircle As CircTrail
Dim TimeBox As udtBox
Dim ScrScr As SS
Dim iNumTails As Integer
Dim lStartCol As Long
Dim lCol As Long
Dim bStop As Boolean
Dim bDrawCircles As Boolean
Dim bTwitch As Boolean
Dim lNumerics() As Long
Dim bZoom As Boolean
Dim bTime As Boolean


Private Sub Form_Click()
bStop = True
End Sub
Sub DisplayTime()

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        Static X0 As Integer, Y0 As Integer
'-----------------------------------------------------------------
    If (RunMode = RM_NORMAL) Then           ' Determine screen saver mode
        If ((X0 = 0) And (Y0 = 0)) Or _
           ((Abs(X0 - x) < 5) And (Abs(Y0 - y) < 5)) Then ' small mouse movement...
            X0 = x                          ' Save current x coordinate
            Y0 = y                          ' Save current y coordinate
            Exit Sub                        ' Exit
        End If
    
        Unload Me
        End ' Large mouse movement (terminate screensaver)
    End If
End Sub
Sub ScrollScreen()
Dim lH As Integer
Dim lW As Integer
Dim lRC As Long
Dim y As Integer
Dim z As Integer
Dim x As Long
Const iScrollPixels As Integer = 2 'iCharWidth
Const iRowHeight As Integer = 2 'iCharHeight
   
    ScrScr.LastTickCount = GetTickCount
    ScrScr.NextSS = Int(Rnd * 4000) + 4000
   
    lH = picText.Height
    lW = picText.Width
    If picTextS.Width <> picText.Width Then
        picTextS.Move 0, 0, picText.Width, picText.Height
    End If
    For z = 0 To 2
       For y = 0 To lH Step iRowHeight
        If y Mod iRowHeight * 2 = 0 Then
           'save column
           BitBlt picTextS.hdc, 0, y, iScrollPixels, iRowHeight, _
                  picText.hdc, lW - iScrollPixels, y, vbSrcCopy
           'move bulk of screen right
           BitBlt picText.hdc, iScrollPixels, y, lW - iScrollPixels, iRowHeight, _
                  picText.hdc, 0, y, vbSrcCopy
           'move saved column to left of new screen
           BitBlt picText.hdc, 0, y, iScrollPixels, iRowHeight, _
                  picTextS.hdc, 0, y, vbSrcCopy
        Else
           'save column
           BitBlt picTextS.hdc, 0, y, iScrollPixels, iRowHeight, _
                  picText.hdc, 0, y, vbSrcCopy
           'move bulk of screen right
           BitBlt picText.hdc, 0, y, lW - iScrollPixels, iRowHeight, _
                  picText.hdc, iScrollPixels, y, vbSrcCopy
           'move saved column to left of new screen
           BitBlt picText.hdc, lW - iScrollPixels, y, iScrollPixels, iRowHeight, _
                  picTextS.hdc, 0, y, vbSrcCopy
        End If

       Next
        BitBlt Me.hdc, 0, 0, picText.Width, picText.Height, picText.hdc, 0, 0, vbSrcCopy
    Next
    
End Sub

Private Sub Form_Load()
    bStop = False
    iNumChars = Len(MC)
    Randomize Timer

    SetupFont

    lStartCol = RGB(10, 20, 10)

    picText.Move 0, 0, Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY

    iNumCols = picText.Width \ iCharWidth
    iNumRows = picText.Height \ iCharHeight
    
    If Len(GetSetting(App.Title, "Settings", "NumCols")) > 0 Then
        If CInt(GetSetting(App.Title, "Settings", "NumCols")) > iNumCols Then
            iNumTails = CInt(GetSetting(App.Title, "Settings", "NumCols"))
        Else
            iNumTails = iNumCols
        End If
    Else
        iNumTails = iNumCols
    End If
    
    bDrawCircles = IIf(GetSetting(App.Title, "Settings", "Circles") = vbChecked, True, False)
    bTwitch = IIf(GetSetting(App.Title, "Settings", "Twitch") = vbChecked, True, False)
    bZoom = IIf(GetSetting(App.Title, "Settings", "Zoom") = vbChecked, True, False)
    bTime = IIf(GetSetting(App.Title, "Settings", "Time") = vbChecked, True, False)
    SetUpTails
    If bTime Then SetupBox
    If bTwitch Then setupSS
    If bDrawCircles Then SetUpCircle
    
    If (RunMode = RM_NORMAL) Then ShowCursor 0
    InitDeskDC DeskDC, DeskBmp, DispRec
    Timer1.Enabled = True
End Sub
Sub DrawBox()
    With TimeBox
        .Text = Format(Now, "yyyy/mm/dd hh:mm:ss")
        If .BoxOpened = 0 Then 'opening box
            'x axis first
            If .xCurr < (.xMax + 1) Then
                DrawRectangle picTime, 1, 1, .xCurr, .yCurr, RGB(51, 102, 51)
                .xCurr = .xCurr + 4
            ElseIf .yCurr < (.yMax + 1) Then
                DrawRectangle picTime, 1, 1, .xCurr, .yCurr, RGB(51, 102, 51)
                .xCurr = .xMax + 1
                .yCurr = .yCurr + 2
            Else
                .yCurr = .yMax + 1
                .BoxOpened = GetTickCount
                .LastTickCount = GetTickCount
            End If
        ElseIf .BoxOpened > 0 And _
               .xCurr >= (.xMax + 1) And .yCurr >= (.yMax + 1) And _
               (.BoxOpenFor + .LastTickCount) > GetTickCount Then 'leave box open
            DrawRectangle picTime, 1, 1, .xCurr, .yCurr, RGB(51, 102, 51)
            picTime.CurrentX = 5
            picTime.CurrentY = 3
            picTime.Print Format(Now, "yyyy/mm/dd hh:mm:ss")
        ElseIf .yCurr > 2 Then 'box closing
            .yCurr = .yCurr - 2
            DrawRectangle picTime, 1, 1, .xCurr, .yCurr, RGB(51, 102, 51)
        Else
            .BoxOpened = 0
            .LastTickCount = GetTickCount
            .xCurr = 1
            .yCurr = 1
            .x = Int(Rnd * (picText.Width - .xMax))
            .y = Int(Rnd * (picText.Height - .yMax))
        End If
               
    End With
End Sub
Sub DrawRectangle(picIn As PictureBox, x1 As Long, y1 As Long, x2 As Long, y2 As Long, lColor As Long)
Dim hRPen As Long
    hRPen = CreatePen(0, 1, lColor)
    DeleteObject SelectObject(picIn.hdc, hRPen)
    Rectangle picIn.hdc, x1, y1, x2, y2
    DeleteObject hRPen
    DoEvents
End Sub
Sub SetupBox()
    picTime.Font = "Courier New"
    picTime.FontSize = 16
    picTime.FontBold = True
    picTime.ForeColor = RGB(102, 204, 102)
    With TimeBox
        .Text = Format(Now, "yyyy/mm/dd hh:mm:ss")
        .xMax = picTime.TextWidth(.Text) + 7
        .yMax = picTime.TextHeight(.Text) + 2
        .x = Int(Rnd * (picText.Width - .xMax - 170))
        .y = Int(Rnd * (picText.Height - .yMax - 170))
        .xCurr = 1
        .yCurr = 1
        .LastTickCount = GetTickCount
        .TimeTilNextBox = Int(Rnd * 4000) + 4000
        .BoxOpenFor = 20000
        .BoxOpened = 0
    End With
End Sub
Sub DrawScreen()
Dim iX As Integer
Dim xMag As Long, yMag As Long
    If bTwitch Then
        If GetTickCount > ScrScr.LastTickCount + ScrScr.NextSS Then
            ScrollScreen
        End If
    End If
    picText.Cls
    For iX = 0 To iNumTails - 1
        DrawTail iX
    Next
    If bDrawCircles Then
        If GetTickCount > (udtCircle.LastTickCount + udtCircle.NextCircle) Then
            DrawCircle
        End If
    End If
    If bTime Then
        If GetTickCount > (TimeBox.LastTickCount + TimeBox.TimeTilNextBox) _
        Or TimeBox.BoxOpened > 0 Then
            picTime.Cls
            DrawBox
            BitBlt picText.hdc, TimeBox.x, TimeBox.y, TimeBox.xCurr + 1, TimeBox.yCurr + 1, picTime.hdc, 0, 0, vbSrcCopy
        End If
    End If
    If bZoom Then
        For iX = 20 To picText.Height Step (picMagnify.Height + 20)
            xMag = Int(Rnd * (picText.Width - 220))
            yMag = Int(Rnd * (picText.Height - 220))
            StretchBlt picMagnify.hdc, 0, 0, picMagnify.Width, picMagnify.Height, _
                       picText.hdc, xMag, yMag, 50, 50, _
                       vbSrcCopy
            DrawRectangle picText, xMag - 1, yMag - 1, xMag + 51, yMag + 51, RGB(20, 40, 20)
            DrawRectangle picMagnify, 0, 0, picMagnify.Width, picMagnify.Height, vbBlack
            DrawRectangle picMagnify, 1, 1, picMagnify.Width - 1, picMagnify.Height - 1, RGB(51, 102, 51)
            DrawRectangle picMagnify, 2, 2, picMagnify.Width - 2, picMagnify.Height - 2, vbBlack
            BitBlt picText.hdc, picText.Width - picMagnify.Width - 20, iX, picMagnify.Width, picMagnify.Height, picMagnify.hdc, 0, 0, vbSrcCopy
        Next
    End If
    
    BitBlt Me.hdc, 0, 0, picText.Width, picText.Height, picText.hdc, 0, 0, vbSrcCopy
End Sub
Private Sub Form_Unload(Cancel As Integer)
'-----------------------------------------------------------------
    Dim Idx As Integer                          ' Array index
'-----------------------------------------------------------------
    ' [* YOU MUST TURN OFF THE TIMER BEFORE DESTROYING THE SPRITE OBJECT *]
    Timer1.Enabled = False                     ' [* YOU MAY DEADLOCK!!! *]
'   Set gSpriteCollection = Nothing             ' Not sure if this would work...

    DelDeskDC DeskDC                            ' Cleanup the DeskDC (Memleak will occure if not done)
    
    If (RunMode = RM_NORMAL) Then ShowCursor -1 ' Show MousePointer
    Screen.MousePointer = vbDefault             ' Reset MousePointer
    End
'-----------------------------------------------------------------
End Sub
Sub DrawCircle()
Dim xPos As Integer
Dim yPos As Integer
Dim sAngle As Single
Dim iColor As Integer
Dim iNumCircles As Integer

    iColor = 255 - (udtCircle.r * 7)
    If (iColor \ 2) < 2 Then iColor = 2
    For iNumCircles = 0 To 1
    For sAngle = 0 To 360 Step 6
        xPos = (GimmeX(sAngle, udtCircle.r - (iNumCircles * 2)) + udtCircle.x) * (iCharWidth * 1.2)
        yPos = (GimmeY(sAngle, udtCircle.r - (iNumCircles * 2)) + udtCircle.y) * iCharHeight
        picText.ForeColor = RGB(iColor \ 2, iColor, iColor \ 2)
        picText.CurrentX = xPos
        picText.CurrentY = yPos
        If xPos >= 0 And xPos < picText.Width And yPos >= 0 And yPos < picText.Height Then
            picText.Print Mid$(MC, Int(Rnd * iNumChars) + 1, 1)
        End If
    Next
    Next
    With udtCircle
        .r = .r + 2
        If iColor = 2 Then
            .LastTickCount = GetTickCount
            .NextCircle = Int(Rnd * 4000) + 3000
            .r = 4
            .x = Int(Rnd * iNumCols)
            .y = Int(Rnd * iNumRows)
        End If
    End With
End Sub
Sub DrawTail(iTail As Integer)
Dim yPosition As Integer
Dim iPercentageInc As Integer
Dim lCol As Long
    With udtTails(iTail)
        iPercentageInc = (.TailLength \ 2) + 3
        lCol = lStartCol
        '.x = .x + Int(Rnd * 3) - 1
        For yPosition = 0 To .TailLength
            picText.CurrentX = xOffset + (.x * iCharWidth)
            picText.CurrentY = yOffset + ((.y + yPosition) * iCharHeight)
            picText.ForeColor = lCol
            lCol = AdjustBrightness(lCol, iPercentageInc, True)
            'iPercentageInc = iPercentageInc * (1 + (Rnd * 1))
            picText.Print Mid$(MC, 1 + (Rnd * iNumChars), 1)
            .LastTickCount = .LastTickCount + 1
            If .LastTickCount > .Speed Then
                .LastTickCount = 0
                .y = .y + 1
                If .y > iNumRows Then
                    .y = .TailLength * -1
                End If
            End If
        Next
    End With
End Sub
Sub SetupFont()
    With picText
        .Font = "Matrix"
        .FontSize = 14
    End With
    
    'do color shift later
    picText.ForeColor = vbGreen
End Sub
Sub SetUpCircle()
    With udtCircle
        .LastTickCount = GetTickCount
        .x = Int(Rnd * iNumCols)
        .y = Int(Rnd * iNumRows)
        .r = 4
        .NextCircle = Int(Rnd * 4000) + 3000
    End With
End Sub
Sub setupSS()
    ScrScr.LastTickCount = GetTickCount
    ScrScr.NextSS = Int(Rnd * 4000 + 4000)
End Sub
Sub SetUpTails()
Dim x As Integer
    ReDim udtTails(iNumTails)
    For x = 0 To iNumTails - 1
        With udtTails(x)
            .LastTickCount = GetTickCount
            .Speed = (Rnd * 50) + 20
            .TailLength = Int(Rnd * 25) + 4
            .x = x 'Rnd * iNumCols
            If .x >= iNumCols Then
                .x = Rnd * iNumCols
            End If
            .y = ((Rnd * (iNumRows * 1.5)) + .TailLength) * -1
        End With
    Next
End Sub

Private Sub Timer1_Timer()
    DrawScreen
    If bStop Then
        Timer1.Enabled = True
        Unload Me
        End
    End If
End Sub
