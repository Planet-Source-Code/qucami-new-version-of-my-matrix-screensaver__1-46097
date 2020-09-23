VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Matrix Screensaver"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkTime 
      Alignment       =   1  'Right Justify
      Caption         =   "Show Time"
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   2280
      Width           =   2115
   End
   Begin VB.CheckBox chkZoom 
      Alignment       =   1  'Right Justify
      Caption         =   "Magnification boxes"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   1980
      Width           =   2115
   End
   Begin VB.CheckBox chkTwitch 
      Alignment       =   1  'Right Justify
      Caption         =   "Code Interference haze"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   1680
      Width           =   2115
   End
   Begin VB.CheckBox chkCircles 
      Alignment       =   1  'Right Justify
      Caption         =   "Draw Code Circles"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   1380
      Width           =   2115
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3600
      TabIndex        =   4
      Top             =   3060
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Top             =   3060
      Width           =   975
   End
   Begin VB.TextBox txtNumCodeCols 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2100
      TabIndex        =   1
      Text            =   "1"
      Top             =   900
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "Number of code columns"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   960
      Width           =   1875
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form2.frx":0000
      Height          =   615
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   4275
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim iMinCols As Integer
    iMinCols = (Screen.Width \ Screen.TwipsPerPixelX) \ iCharWidth
    If Not IsNumeric(txtNumCodeCols) Then
        MsgBox "Numerics for number of code columns!", vbExclamation, "Matrix Screensaver"
        Exit Sub
    End If
    If CInt(txtNumCodeCols) < iMinCols Then
        MsgBox "Number of columns too low. Setting to minimum.", vbExclamation, "Matrix Screensaver"
        txtNumCodeCols = iMinCols
        Exit Sub
    End If
    
    SaveSetting App.Title, "Settings", "NumCols", txtNumCodeCols
    SaveSetting App.Title, "Settings", "Circles", chkCircles.Value
    SaveSetting App.Title, "Settings", "Twitch", chkTwitch.Value
    SaveSetting App.Title, "Settings", "Time", chkTime.Value
    SaveSetting App.Title, "Settings", "Zoom", chkZoom.Value
    
    Unload Me
    End
End Sub

Private Sub Command2_Click()
    Unload Me
    End
End Sub

Private Sub Form_Load()
Dim iMinCols As Integer
    iMinCols = (Screen.Width \ Screen.TwipsPerPixelX) \ iCharWidth
    txtNumCodeCols = GetSetting(App.Title, "Settings", "NumCols")
    If CInt(txtNumCodeCols) < iMinCols Then
        txtNumCodeCols = iMinCols
        SaveSetting App.Title, "Settings", "NumCols", iMinCols
    End If
    chkCircles.Value = GetSetting(App.Title, "Settings", "Circles")
    chkTwitch.Value = GetSetting(App.Title, "Settings", "Twitch")
    chkTime.Value = GetSetting(App.Title, "Settings", "Time")
    chkZoom.Value = GetSetting(App.Title, "Settings", "Zoom")
End Sub
