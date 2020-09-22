VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Wav 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "LMS NETWORKS MP3 Player"
   ClientHeight    =   3315
   ClientLeft      =   165
   ClientTop       =   195
   ClientWidth     =   6525
   Icon            =   "MP3.frx":0000
   ScaleHeight     =   3315
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   4800
      Top             =   120
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00000000&
      Caption         =   "Save Current position in song last played; like a Resume function on a CD player"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3000
      Width           =   6015
   End
   Begin RichTextLib.RichTextBox RT1 
      Height          =   135
      Left            =   6120
      TabIndex        =   20
      Top             =   360
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   238
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"MP3.frx":000C
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   2640
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin RichTextLib.RichTextBox RT 
      Height          =   135
      Left            =   6000
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   238
      _Version        =   393217
      TextRTF         =   $"MP3.frx":00BA
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   0
   End
   Begin VB.Timer Timer6 
      Interval        =   1000
      Left            =   360
      Top             =   0
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      TickFrequency   =   0
   End
   Begin VB.TextBox MMFile 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   840
      TabIndex        =   13
      Top             =   600
      Width           =   5655
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      ItemData        =   "MP3.frx":0168
      Left            =   3840
      List            =   "MP3.frx":016A
      TabIndex        =   11
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   0
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Reverse"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      ToolTipText     =   "Rewind current MP3"
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "F. Forward"
      Height          =   375
      Left            =   4440
      TabIndex        =   9
      ToolTipText     =   "Fast Forward current MP3"
      Top             =   1440
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   600
      Top             =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   0
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Repeat"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      ToolTipText     =   "Repeat MP3 forever"
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Pause"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      ToolTipText     =   "Pause current MP3"
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      ToolTipText     =   "Stop current MP3"
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Play"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Play currnet MP3"
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Browse.."
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      ToolTipText     =   "Browse for MP3 file..."
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label4 
      Height          =   135
      Left            =   5760
      TabIndex        =   21
      Top             =   360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   120
      X2              =   6360
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   120
      X2              =   6360
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label LFile 
      Height          =   135
      Left            =   5520
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "MP3 Player"
      BeginProperty Font 
         Name            =   "PlaLCD"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   1800
      TabIndex        =   15
      Top             =   120
      Width           =   2895
   End
   Begin VB.Line Line8 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   120
      X2              =   120
      Y1              =   960
      Y2              =   1320
   End
   Begin VB.Line Line7 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   6360
      X2              =   6360
      Y1              =   960
      Y2              =   1320
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "The last MP3 played before the program was closed:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Label MFile 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "FilePath:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Option for play:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin MediaPlayerCtl.MediaPlayer Player 
      Height          =   255
      Left            =   4920
      TabIndex        =   0
      Top             =   1920
      Width           =   1575
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   30
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   -2147483633
      DisplayForeColor=   65280
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   0   'False
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   0   'False
      EnableTracker   =   0   'False
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   -1  'True
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   -2147483633
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Pop-Up"
      Visible         =   0   'False
      Begin VB.Menu mnuPlay 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnuPause 
         Caption         =   "Pause"
      End
      Begin VB.Menu mnuBrowse 
         Caption         =   "Browse..."
      End
      Begin VB.Menu Slash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMin 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnuSlash 
         Caption         =   "-"
      End
      Begin VB.Menu Option 
         Caption         =   "Option"
         Begin VB.Menu Repeat 
            Caption         =   "Repeat"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu Slash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu Slash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Wav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub Check1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Wav.PopupMenu mnuPopUp
    Else
        DoEvents
    End If
End Sub

Private Sub Check2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Wav.PopupMenu mnuPopUp
    Else
        DoEvents
    End If
End Sub

Private Sub Command10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
      Timer3.Enabled = True
    End If
    
    If Button = 2 Then
        Wav.PopupMenu mnuPopUp
    Else
        DoEvents
    End If
End Sub

Private Sub Command10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer3.Enabled = False
End Sub

Private Sub Command11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
      Timer4.Enabled = True
    End If
    
    If Button = 2 Then
        Wav.PopupMenu mnuPopUp
    Else
        DoEvents
    End If
End Sub

Private Sub Command11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer4.Enabled = False
End Sub

Private Sub Command6_Click()
    
    On Error Resume Next
    CommonDialog2.Filter = "MP3 Files (*.mp3)|*.mp3"
    CommonDialog2.DialogTitle = "LMS NETWORKS MP3 Play"
    CommonDialog2.InitDir = "C:\Program Files\MusicMatch\Music"
    CommonDialog2.ShowOpen
    
    If CommonDialog2.CancelError = True Then
      Exit Sub
    End If
        
    If MMFile.Text <> CommonDialog2.FileName Then
      MMFile.Text = (CommonDialog2.FileName)
      Label4.Caption = ""
      Player.Open (MMFile.Text)
      List1.RemoveItem (0)
      LFile.Caption = MMFile.Text
      RT.Text = MMFile.Text
      Slider1.Min = 0
      MFile.Caption = "Status: Playing"
      Timer5.Enabled = True
    End If
    
End Sub

Private Sub Command6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Wav.PopupMenu mnuPopUp
    Else
        DoEvents
    End If
End Sub

Private Sub Command7_Click()
    On Error Resume Next
    If MMFile.Text <> "" Then
      Player.Play
      MFile.Caption = "Status: Playing"
    End If
End Sub

Private Sub Command7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Wav.PopupMenu mnuPopUp
    Else
        DoEvents
    End If
End Sub

Private Sub Command8_Click()
    On Error Resume Next
    If MMFile.Text <> "" Then
      Player.CurrentPosition = 0
      Player.Stop
      MFile.Caption = "Status: Stopped"
    End If
End Sub

Private Sub Command8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Wav.PopupMenu mnuPopUp
    Else
        DoEvents
    End If
End Sub

Private Sub Command9_Click()
    On Error Resume Next
    If MMFile.Text <> "" Then
      Player.Pause
      MFile.Caption = "Status: Paused"
    End If
End Sub

Private Sub Command9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Wav.PopupMenu mnuPopUp
    Else
        DoEvents
    End If
End Sub

Private Sub Form_Load()
    If DirectoryExists("C:\LMS NETWORKS") = True Then
      If FileExists("C:\LMS NETWORKS\Last.wt") = True Then
        RT.LoadFile "C:\LMS NETWORKS\Last.wt", rftRFt
        LFile.Caption = RT.Text
        List1.AddItem Dir(LFile.Caption)
      End If
      If FileExists("C:\LMS NETWORKS\Post.wt") = True Then
        RT1.LoadFile "C:\LMS NETWORKS\Post.wt", rftRFt
        Label4.Caption = RT1.Text
        Check2.Value = 1
      End If
    End If
    ProgressBar1.Value = 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Wav.PopupMenu mnuPopUp
    Else
        DoEvents
    End If
End Sub

Private Sub Form_Resize()
    If Wav.WindowState = 2 Then
    Wav.WindowState = 0
    End If

    If Wav.WindowState <> 1 Then
      If Wav.Height <> 3780 Then
        Wav.Height = 3780
      End If
    End If
    
    If Wav.WindowState <> 1 Then
      If Wav.Width <> 6645 Then
        Wav.Width = 6645
      End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If DirectoryExists("C:\LMS NETWORKS") = False Then
      MkDir "C:\LMS NETWORKS"
    End If
    
    If DirectoryExists("C:\LMS NETWORKS") = True Then
      RT.SaveFile "C:\LMS NETWORKS\Last.wt", rftRFt
      If Check2.Value = 1 Then
        RT1.Text = Round(Player.CurrentPosition, 1)
        RT1.SaveFile "C:\LMS NETWORKS\Post.wt", rftRFt
      End If
    
      If Check2.Value = 0 Then
        If FileExists("C:\LMS NETWORKS\Post.wt") = True Then
          Kill "C:\LMS NETWORKS\Post.wt"
        End If
      End If
    End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Wav.PopupMenu mnuPopUp
    Else
        DoEvents
    End If
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Wav.PopupMenu mnuPopUp
    Else
        DoEvents
    End If
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Wav.PopupMenu mnuPopUp
    Else
        DoEvents
    End If
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Wav.PopupMenu mnuPopUp
    Else
        DoEvents
    End If
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Wav.PopupMenu mnuPopUp
    Else
        DoEvents
    End If
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Wav.PopupMenu mnuPopUp
    Else
        DoEvents
    End If
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Wav.PopupMenu mnuPopUp
    Else
        DoEvents
    End If
End Sub

Private Sub List1_DblClick()
    Select Case List1.ListIndex
    Case 0
      If List1.Text <> "" Then
        CommonDialog2.FileName = LFile.Caption
        MMFile.Text = CommonDialog2.FileName
        Player.Open (LFile.Caption)
        MFile.Caption = "Status: Playing"
        Timer7.Enabled = True
      End If
    End Select
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Wav.PopupMenu mnuPopUp
    Else
        DoEvents
    End If
End Sub

Private Sub MFile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Wav.PopupMenu mnuPopUp
    Else
        DoEvents
    End If
End Sub

Private Sub mnuBrowse_Click()
    Command6_Click
End Sub

Private Sub mnuExit_Click()
    Load Files
    Files.Show
    Wav.Hide
    Unload Wav
End Sub

Private Sub mnuMin_Click()
    Wav.WindowState = 1
End Sub

Private Sub mnuPause_Click()
    Command9_Click
End Sub

Private Sub mnuPlay_Click()
    Command7_Click
End Sub

Private Sub mnuRefresh_Click()
    Wav.Refresh
End Sub

Private Sub mnuStop_Click()
    Command8_Click
End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Wav.PopupMenu mnuPopUp
    Else
        DoEvents
    End If
End Sub

Private Sub Repeat_Click()
    Repeat.Checked = Not Repeat.Checked
        
    If Repeat.Checked = True Then
      Check1.Value = 1
    Else
      Check1.Value = 0
    End If
End Sub

Private Sub Slider1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Wav.PopupMenu mnuPopUp
    Else
        DoEvents
    End If
End Sub

Private Sub Slider1_Scroll()
    Slider1.ToolTipText = ""
    Player.CurrentPosition = Slider1.Value
End Sub

Private Sub Timer1_Timer()
    If Check1.Value = 1 Then
    Player.PlayCount = 1000000
    Else
    Player.PlayCount = 1
    End If
    
    If MMFile.Text = "" Then
      If MFile.Caption <> "" Then
      MFile.Caption = MFile.Caption
      Else
      MFile.Caption = ""
      End If
    End If
    
End Sub

Private Sub Timer2_Timer()
    
    If CommonDialog2.FileName <> "" Then
      If MMFile.Text <> CommonDialog2.FileName Then
      MMFile.Text = CommonDialog2.FileName
      End If
    End If
    
    If Check1.Value = 1 Then
      Repeat.Checked = True
    Else
      Repeat.Checked = False
    End If
End Sub

Private Sub Timer3_Timer()
    Player.CurrentPosition = Player.CurrentPosition + 1
End Sub

Private Sub Timer4_Timer()
    If Player.CurrentPosition <= 1 Then
      Timer4.Enabled = False
    Else
      Player.CurrentPosition = Player.CurrentPosition - 1
    End If
        
End Sub

Private Sub Timer5_Timer()
    List1.AddItem Dir(MMFile.Text)
    Timer5.Enabled = False
End Sub

Private Sub Timer6_Timer()
    Dim Fall As Long
    
    On Error Resume Next
    If CommonDialog2.FileName <> "" Then
      Slider1.Max = Player.SelectionEnd
      ProgressBar1.Max = Player.SelectionEnd
      Fall = ProgressBar1.Value / ProgressBar1.Max * 100
      Label3.Caption = Round(Fall, 1) & "% Of song left"
      ProgressBar1.Value = Player.SelectionEnd - Player.CurrentPosition
    End If
    Slider1.Value = Player.CurrentPosition
        
    If Slider1.Value = Slider1.Max Then
      MFile.Caption = "Status: Stopped"
        If Check1.Value = 1 Then
          If Player.PlayState <> mpStopped Then
            MFile.Caption = "Status: Playing"
          End If
        End If
    End If
End Sub

Public Function FileExists(sFileName As String) As Boolean
    If Len(sFileName$) = 0 Then
        FileExists = False
        Exit Function
    End If
    If Len(Dir$(sFileName$)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

Public Function DirectoryExists(Path As String) As Boolean
    On Error Resume Next
    Dim IDE As Integer
    
    IDE = GetAttr(Path)
    
    If IDE = vbDirectory Then
      DirectoryExists = True
    Else
      DirectoryExists = False
    End If
End Function

Private Sub Timer7_Timer()
    If Label4.Caption <> "" Then
      Player.CurrentPosition = Label4.Caption
    End If
    Timer7.Enabled = False
End Sub
