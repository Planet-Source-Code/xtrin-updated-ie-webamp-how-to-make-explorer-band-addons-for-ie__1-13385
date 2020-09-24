VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form frmBand 
   BorderStyle     =   0  'None
   Caption         =   "WebAmp"
   ClientHeight    =   4035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdResume 
      Caption         =   "Resume"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton CmdPause 
      Caption         =   "Pause"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Timer tmrPosition 
      Left            =   840
      Top             =   3240
   End
   Begin MSComctlLib.Slider sldPosition 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CdbOpen 
      Left            =   1440
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmControls 
      Caption         =   " Controls "
      Height          =   3615
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   3255
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblName 
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   2655
      End
   End
   Begin MediaPlayerCtl.MediaPlayer mp3player 
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
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
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "frmBand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Const BORDER_X = 255
Private Const BORDER_Y = 127


Public Sub MenuHandler(lpcmi As Long)

Dim cmi As CMINVOKECOMMANDINFO
CopyMemory cmi, ByVal lpcmi, Len(cmi)

Select Case cmi.lpVerb
  Case 0
   cmdOpen = True
  Case 1
   cmdPlay = True
  Case 2
   cmdStop = True
End Select
End Sub

Private Sub cmdOpen_Click()
CdbOpen.Filter = " mp3 files (*.mp3) | *.mp3"
CdbOpen.ShowOpen
lblName.Caption = CdbOpen.FileName
End Sub

Private Sub CmdPause_Click()
On Error Resume Next
mp3player.Pause
End Sub

Private Sub cmdPlay_Click()
On Error Resume Next
Dim sPlay As String
sPlay = lblName.Caption
mp3player.FileName = sPlay
mp3player.Play
tmrPosition.Interval = 1
tmrPosition.Enabled = True

End Sub

Private Sub CmdResume_Click()
On Error Resume Next
mp3player.Play
mp3player.CurrentPosition = sldPosition.Value
End Sub

Private Sub cmdStop_Click()
On Error Resume Next
mp3player.Stop
sldPosition.Value = 0
tmrPosition.Enabled = False

End Sub

Private Sub sldPosition_Scroll()
On Error Resume Next
mp3player.CurrentPosition = sldPosition.Value
End Sub

Private Sub tmrPosition_Timer()
On Error Resume Next
sldPosition.Value = mp3player.CurrentPosition
sldPosition.Max = mp3player.Duration
End Sub

