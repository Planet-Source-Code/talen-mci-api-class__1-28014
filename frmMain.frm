VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MCI Sample"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Slider sldPlayrate 
      Height          =   255
      Left            =   5280
      TabIndex        =   59
      Top             =   3540
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   100
      SmallChange     =   10
      Max             =   2000
      SelStart        =   1000
      TickFrequency   =   200
      Value           =   1000
   End
   Begin MSComctlLib.Slider sldTracker 
      Height          =   255
      Left            =   2100
      TabIndex        =   47
      Top             =   3540
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   450
      _Version        =   393216
      TickStyle       =   3
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1140
      Top             =   3120
   End
   Begin VB.CheckBox chkBoth 
      Caption         =   "Both"
      Height          =   255
      Left            =   5940
      TabIndex        =   40
      Top             =   2880
      Value           =   1  'Checked
      Width           =   675
   End
   Begin VB.CheckBox chkRight 
      Caption         =   "Right"
      Height          =   255
      Left            =   5220
      TabIndex        =   39
      Top             =   2880
      Value           =   1  'Checked
      Width           =   675
   End
   Begin VB.CheckBox chkLeft 
      Caption         =   "Left"
      Height          =   255
      Left            =   4560
      TabIndex        =   38
      Top             =   2880
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.PictureBox picVideo 
      BackColor       =   &H00000000&
      Height          =   1875
      Left            =   2160
      ScaleHeight     =   1815
      ScaleWidth      =   2235
      TabIndex        =   23
      Top             =   1260
      Width           =   2295
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   435
      Left            =   1020
      TabIndex        =   22
      Top             =   1020
      Width           =   915
   End
   Begin VB.CommandButton cmdToEnd 
      Caption         =   ">>"
      Height          =   435
      Left            =   1020
      TabIndex        =   21
      Top             =   1500
      Width           =   915
   End
   Begin VB.CommandButton cmdToStart 
      Caption         =   "<<"
      Height          =   435
      Left            =   60
      TabIndex        =   20
      Top             =   1500
      Width           =   915
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   435
      Left            =   60
      TabIndex        =   19
      Top             =   1020
      Width           =   915
   End
   Begin MSComDlg.CommonDialog cdgOpen 
      Left            =   180
      Top             =   5580
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select file to play"
      FileName        =   "*.*"
      Filter          =   "All files (*.*)|*.*"
      MaxFileSize     =   9999
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Height          =   435
      Left            =   1020
      TabIndex        =   18
      Top             =   540
      Width           =   915
   End
   Begin VB.CommandButton cmdStretch 
      Caption         =   "Resize to fit (in pixels)"
      Height          =   315
      Left            =   4500
      TabIndex        =   15
      Top             =   420
      Width           =   2055
   End
   Begin VB.CommandButton cmdResize 
      Caption         =   "Resize (in pixels)"
      Height          =   315
      Left            =   4500
      TabIndex        =   14
      Top             =   60
      Width           =   2055
   End
   Begin VB.TextBox txtSize 
      Height          =   315
      Index           =   3
      Left            =   3720
      TabIndex        =   9
      Top             =   420
      Width           =   675
   End
   Begin VB.TextBox txtSize 
      Height          =   315
      Index           =   2
      Left            =   3720
      TabIndex        =   8
      Top             =   60
      Width           =   675
   End
   Begin VB.TextBox txtSize 
      Height          =   315
      Index           =   1
      Left            =   2460
      TabIndex        =   7
      Top             =   60
      Width           =   675
   End
   Begin VB.TextBox txtSize 
      Height          =   315
      Index           =   0
      Left            =   2460
      TabIndex        =   6
      Top             =   420
      Width           =   675
   End
   Begin VB.CheckBox chkFullScreen 
      Caption         =   "Fullscreen"
      Height          =   255
      Left            =   540
      TabIndex        =   5
      Top             =   2820
      Width           =   1395
   End
   Begin VB.TextBox txtTo 
      Height          =   315
      Left            =   840
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtFrom 
      Height          =   315
      Left            =   840
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Height          =   435
      Left            =   60
      TabIndex        =   2
      Top             =   540
      Width           =   915
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H00C0C0C0&
      Height          =   915
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   5520
      Width           =   6630
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load file..."
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1875
   End
   Begin MSComctlLib.Slider sldVolume 
      Height          =   1575
      Index           =   1
      Left            =   5220
      TabIndex        =   27
      Top             =   1260
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   2778
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   100
      SmallChange     =   10
      Min             =   -1000
      Max             =   0
      SelStart        =   -500
      TickStyle       =   3
      Value           =   -500
   End
   Begin MSComctlLib.Slider sldVolume 
      Height          =   1575
      Index           =   2
      Left            =   5940
      TabIndex        =   28
      Top             =   1260
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   2778
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   100
      SmallChange     =   10
      Min             =   -1000
      Max             =   0
      SelStart        =   -500
      TickStyle       =   3
      Value           =   -500
   End
   Begin MSComctlLib.Slider sldVolume 
      Height          =   1575
      Index           =   0
      Left            =   4560
      TabIndex        =   26
      Top             =   1260
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   2778
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   100
      SmallChange     =   10
      Min             =   -1000
      Max             =   0
      SelStart        =   -500
      TickStyle       =   3
      Value           =   -500
   End
   Begin VB.Label mailto 
      Caption         =   "mailto:another_reality@mail.ru"
      Height          =   195
      Left            =   120
      TabIndex        =   77
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Label lblAlbum 
      Height          =   195
      Left            =   5520
      TabIndex        =   76
      Top             =   5100
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      Height          =   195
      Left            =   5520
      TabIndex        =   75
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblArtist 
      Height          =   195
      Left            =   5520
      TabIndex        =   74
      Top             =   4500
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Album:"
      Height          =   195
      Index           =   11
      Left            =   4740
      TabIndex        =   73
      Top             =   5100
      Width           =   735
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Song title:"
      Height          =   195
      Index           =   10
      Left            =   4740
      TabIndex        =   72
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Artist:"
      Height          =   195
      Index           =   9
      Left            =   4740
      TabIndex        =   71
      Top             =   4500
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "ID3 Tag info:"
      Height          =   195
      Index           =   8
      Left            =   4620
      TabIndex        =   70
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblStatus 
      Caption         =   "none"
      Height          =   195
      Left            =   1260
      TabIndex        =   69
      Top             =   3540
      Width           =   915
   End
   Begin VB.Label lblCTime 
      Caption         =   "00:00:00"
      Height          =   195
      Left            =   1260
      TabIndex        =   68
      Top             =   3780
      Width           =   915
   End
   Begin VB.Label lblTTime 
      Caption         =   "00:00:00"
      Height          =   195
      Left            =   1260
      TabIndex        =   67
      Top             =   4020
      Width           =   915
   End
   Begin VB.Label lblCFrame 
      Caption         =   "00"
      Height          =   195
      Left            =   1260
      TabIndex        =   66
      Top             =   4260
      Width           =   915
   End
   Begin VB.Label lblTFrames 
      Caption         =   "00"
      Height          =   195
      Left            =   1260
      TabIndex        =   65
      Top             =   4500
      Width           =   915
   End
   Begin VB.Label lblFPS 
      Caption         =   "00"
      Height          =   195
      Left            =   1260
      TabIndex        =   64
      Top             =   4740
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Current status:"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   63
      Top             =   3540
      Width           =   1035
   End
   Begin VB.Label lblPlayrate 
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   6300
      TabIndex        =   62
      Top             =   3840
      Width           =   315
   End
   Begin VB.Label lblPlayrate 
      Caption         =   "50%"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   5820
      TabIndex        =   61
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label lblPlayrate 
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   5340
      TabIndex        =   60
      Top             =   3840
      Width           =   195
   End
   Begin VB.Label lbl 
      Caption         =   "Playrate:"
      Height          =   195
      Index           =   28
      Left            =   4620
      TabIndex        =   58
      Top             =   3540
      Width           =   615
   End
   Begin VB.Label lblCSh 
      Height          =   195
      Left            =   3060
      TabIndex        =   57
      Top             =   5220
      Width           =   1335
   End
   Begin VB.Label lblCSw 
      Height          =   195
      Left            =   3060
      TabIndex        =   56
      Top             =   4980
      Width           =   1335
   End
   Begin VB.Label lblBSh 
      Height          =   195
      Left            =   3060
      TabIndex        =   55
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lblBSw 
      Height          =   195
      Left            =   3060
      TabIndex        =   54
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Height"
      Height          =   195
      Index           =   27
      Left            =   2520
      TabIndex        =   53
      Top             =   5220
      Width           =   495
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Width:"
      Height          =   195
      Index           =   26
      Left            =   2520
      TabIndex        =   52
      Top             =   4980
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "Current size:"
      Height          =   255
      Index           =   23
      Left            =   2160
      TabIndex        =   51
      Top             =   4680
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Height:"
      Height          =   195
      Index           =   25
      Left            =   2520
      TabIndex        =   50
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Width:"
      Height          =   195
      Index           =   24
      Left            =   2520
      TabIndex        =   49
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "Best size:"
      Height          =   255
      Index           =   22
      Left            =   2160
      TabIndex        =   48
      Top             =   3900
      Width           =   675
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "FPS:"
      Height          =   255
      Index           =   21
      Left            =   120
      TabIndex        =   46
      Top             =   4740
      Width           =   1035
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Total frames:"
      Height          =   195
      Index           =   20
      Left            =   120
      TabIndex        =   45
      Top             =   4500
      Width           =   1035
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Current frame:"
      Height          =   195
      Index           =   19
      Left            =   120
      TabIndex        =   44
      Top             =   4260
      Width           =   1035
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Total time:"
      Height          =   195
      Index           =   18
      Left            =   120
      TabIndex        =   43
      Top             =   4020
      Width           =   1035
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "Current time:"
      Height          =   195
      Index           =   17
      Left            =   120
      TabIndex        =   42
      Top             =   3780
      Width           =   1035
   End
   Begin VB.Label lbl 
      Caption         =   "File statistics:"
      Height          =   195
      Index           =   16
      Left            =   60
      TabIndex        =   41
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label lblBoth 
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   6240
      TabIndex        =   37
      Top             =   2580
      Width           =   255
   End
   Begin VB.Label lblBoth 
      Caption         =   "50%"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   6240
      TabIndex        =   36
      Top             =   1980
      Width           =   315
   End
   Begin VB.Label lblBoth 
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   6180
      TabIndex        =   35
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label lblRight 
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   5520
      TabIndex        =   34
      Top             =   2580
      Width           =   255
   End
   Begin VB.Label lblRight 
      Caption         =   "50%"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   5520
      TabIndex        =   33
      Top             =   1980
      Width           =   315
   End
   Begin VB.Label lblRight 
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   5460
      TabIndex        =   32
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label lblLeft 
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   4800
      TabIndex        =   29
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label lblLeft 
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   4860
      TabIndex        =   31
      Top             =   2580
      Width           =   255
   End
   Begin VB.Label lblLeft 
      Caption         =   "50%"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   4860
      TabIndex        =   30
      Top             =   1980
      Width           =   315
   End
   Begin VB.Label lbl 
      Caption         =   "Volume:"
      Height          =   195
      Index           =   6
      Left            =   4560
      TabIndex        =   25
      Top             =   900
      Width           =   615
   End
   Begin VB.Label lblVideo 
      Caption         =   "Video here:"
      Height          =   195
      Left            =   2280
      TabIndex        =   24
      Top             =   900
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "To (ms):"
      Height          =   195
      Index           =   5
      Left            =   180
      TabIndex        =   17
      Top             =   2460
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "From (ms):"
      Height          =   195
      Index           =   4
      Left            =   60
      TabIndex        =   16
      Top             =   2100
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "Height:"
      Height          =   195
      Index           =   3
      Left            =   3180
      TabIndex        =   13
      Top             =   480
      Width           =   555
   End
   Begin VB.Label lbl 
      Caption         =   "Width:"
      Height          =   195
      Index           =   2
      Left            =   3180
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lbl 
      Caption         =   "Top:"
      Height          =   195
      Index           =   1
      Left            =   2040
      TabIndex        =   11
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lbl 
      Caption         =   "Left:"
      Height          =   195
      Index           =   0
      Left            =   2040
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mci As clsMCIApi
Attribute mci.VB_VarHelpID = -1
Private Const WS_CHILD = &H40000000
Private mp3tag As mp3tag_dummy

Private Sub chkBoth_Click()
Dim rc&
If chkBoth = vbChecked Then
    chkLeft = vbChecked
    chkRight = vbChecked
ElseIf chkBoth = vbUnchecked Then
    chkLeft = vbUnchecked
    chkRight = vbUnchecked
End If
'If rc <> 0 Then txtLog = txtLog + mci.ErrorDescription + vbCrLf
End Sub

Private Sub chkFullScreen_Click()
cmdPlay_Click
End Sub

Private Sub chkLeft_Click()
Dim rc&
If chkLeft = vbChecked Then
    rc = mci.cmdChannels("device1", 0, True)
    If chkRight = vbChecked Then chkBoth = vbChecked Else chkBoth = vbGrayed
Else
    rc = mci.cmdChannels("device1", 0, False)
    If chkRight = vbChecked Then chkBoth = vbGrayed Else chkBoth = vbUnchecked
End If
If rc <> 0 Then txtLog = txtLog + mci.ErrorDescription + vbCrLf


End Sub

Private Sub chkRight_Click()
Dim rc&
If chkRight = vbChecked Then
    rc = mci.cmdChannels("device1", 1, True)
    If chkLeft = vbChecked Then chkBoth = vbChecked Else chkBoth = vbGrayed
Else
    rc = mci.cmdChannels("device1", 1, False)
    If chkLeft = vbChecked Then chkBoth = vbGrayed Else chkBoth = vbUnchecked
End If
If rc <> 0 Then txtLog = txtLog + mci.ErrorDescription + vbCrLf
End Sub

Private Sub cmdClose_Click()
Dim rc&
rc = mci.CloseDevice("device1")
End Sub

Private Sub cmdLoad_Click()
Dim t$, rc&, i%, v%(0 To 2)
On Error GoTo errhandle

cdgOpen.ShowOpen
t = cdgOpen.filename
txtLog = txtLog + "Initializing device and loading file...": DoEvents
Do
    rc = mci.OpenDevice("MPEGVideo", "device1", t, picVideo.hwnd, CStr(WS_CHILD))
    If rc <> 0 Then
        If mci.ErrorNumber = 289 Then
            rc = mci.CloseDevice("device1")
        Else
            txtLog = vbCrLf + mci.ErrorDescription
            Exit Sub
        End If
    Else
        txtLog = txtLog + "success." + vbCrLf

        Exit Do
    End If
Loop

For i = 0 To 2
    rc = mci.cmdGetVolume("device1", i, v(i))
    sldVolume(i) = v(i) * -1
Next

lblArtist = "": lblAlbum = "": lblTitle = ""
If LCase(right(t, 4)) = ".mp3" Then
    mp3tag = getID3Tag(t)
    lblArtist = mp3tag.artist
    lblAlbum = mp3tag.Album
    lblTitle = mp3tag.SongTitle
End If

lblCFrame = "0"
lblCTime = convertTime(0)
lblTFrames = mci.TotalFrames
lblTTime = convertTime(mci.TotalTime / 1000)
lblFPS = mci.FramesPerSec
rc = mci.cmdTimeFormat("device1", "ms")
sldTracker.Min = 0: sldTracker.Max = mci.TotalTime / 1000
t = mci.cmdGetSize("device1")
lblBSw = left(t, InStr(t, " "))
lblCSw = left(t, InStr(t, " "))
t = right(t, Len(t) - InStr(t, " "))
lblBSh = t
lblCSh = t

Exit Sub
errhandle:
    If Err.Number = 32755 Then  'Cancel error
        Exit Sub
    Else
        On Error GoTo 0
        Err.Raise Err.Number
    End If
End Sub

Private Sub cmdPause_Click()
Dim rc&
rc = mci.cmdPause("device1")
If rc <> 0 Then txtLog = txtLog + mci.ErrorDescription + vbCrLf
End Sub

Private Sub cmdPlay_Click()
Dim f%, t%, fs As Boolean, rc&

f = -1: t = -1
If txtFrom <> "" Then f = Val(txtFrom) * 1000
If txtTo <> "" Then t = Val(txtTo) * 1000
If chkFullScreen.Value = vbChecked Then fs = True Else fs = False
frmMain.MousePointer = 11
rc = mci.cmdPlay("device1", f, t, fs)
frmMain.MousePointer = 0
If rc <> 0 Then txtLog = txtLog + mci.ErrorDescription + vbCrLf
Timer.Enabled = True

End Sub

Private Sub cmdResize_Click()
Dim i%, rc&, t$
For i = 0 To 3
    If txtSize(i) = "" Then Exit Sub
Next

rc = mci.cmdResize("device1", Val(txtSize(0)), Val(txtSize(1)), Val(txtSize(2)), Val(txtSize(3)))
If rc <> 0 Then txtLog = txtLog + mci.ErrorDescription + vbCrLf

t = mci.cmdGetSize("device1")
lblCSw = left(t, InStr(t, " "))
t = right(t, Len(t) - InStr(t, " "))
lblCSh = t

End Sub

Private Sub cmdResizeClient_Click()
Dim rc&
rc = mci.cmdResizeClient("device1", 0, 0, 100, 10)
If rc <> 0 Then txtLog = txtLog + mci.ErrorDescription + vbCrLf
End Sub

Private Sub cmdStop_Click()
Dim rc&
rc = mci.cmdStop("device1")
Timer.Enabled = False
End Sub

Private Sub cmdStretch_Click()
Dim x2%, y2%, rc&, t$

x2 = picVideo.Width / Screen.TwipsPerPixelX
y2 = picVideo.Height / Screen.TwipsPerPixelY

rc = mci.cmdResize("device1", 0, 0, x2, y2)
If rc <> 0 Then txtLog = txtLog + mci.ErrorDescription + vbCrLf

t = mci.cmdGetSize("device1")
lblCSw = left(t, InStr(t, " "))
t = right(t, Len(t) - InStr(t, " "))
lblCSh = t

End Sub

Private Sub cmdToEnd_Click()
Dim rc&
rc = mci.cmdToEnd("device1")
End Sub

Private Sub cmdToStart_Click()
Dim rc&
rc = mci.cmdToStart("device1")
sldTracker = 0
lblCTime = convertTime(0)
End Sub

Private Sub Command1_Click()
MsgBox mci.cmdGetSize("device1")
End Sub

Private Sub Form_Load()
Set mci = New clsMCIApi
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i%
For i = 0 To 2
    lblPlayrate(i).FontUnderline = False
    lblPlayrate(i).ForeColor = &H0
    lblLeft(i).FontUnderline = False
    lblLeft(i).ForeColor = &H0
    lblRight(i).FontUnderline = False
    lblRight(i).ForeColor = &H0
    lblBoth(i).FontUnderline = False
    lblBoth(i).ForeColor = &H0
Next
mailto.FontUnderline = False
mailto.ForeColor = &H0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim rc&

rc = mci.CloseDevice("device1")
If (rc <> 0) And (mci.ErrorNumber <> 263) Then MsgBox mci.ErrorDescription
Set mci = Nothing
End Sub

Private Sub lblBoth_Click(Index As Integer)
Select Case Index
    Case 0
        sldVolume(2) = 0
    Case 1
        sldVolume(2) = -500
    Case 2
        sldVolume(2) = -1000
End Select
sldVolume_Scroll 2
End Sub


Private Sub lblBoth_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblBoth(Index).FontUnderline = True
lblBoth(Index).ForeColor = &HFF0000
End Sub

Private Sub lblLeft_Click(Index As Integer)
Select Case Index
    Case 0
        sldVolume(0) = 0
    Case 1
        sldVolume(0) = -500
    Case 2
        sldVolume(0) = -1000
End Select
sldVolume_Scroll 0
End Sub

Private Sub lblLeft_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblLeft(Index).FontUnderline = True
lblLeft(Index).ForeColor = &HFF0000
End Sub

Private Sub lblPlayrate_Click(Index As Integer)
Select Case Index
    Case 0
        sldPlayrate = 10
    Case 1
        sldPlayrate = 1000
    Case 2
        sldPlayrate = 2000
End Select
sldPlayrate_Scroll
End Sub

Private Sub lblPlayrate_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblPlayrate(Index).FontUnderline = True
lblPlayrate(Index).ForeColor = &HFF0000
End Sub

Private Sub lblRight_Click(Index As Integer)
Select Case Index
    Case 0
        sldVolume(1) = 0
    Case 1
        sldVolume(1) = -500
    Case 2
        sldVolume(1) = -1000
End Select
sldVolume_Scroll 1
End Sub

Private Sub lblRight_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblRight(Index).FontUnderline = True
lblRight(Index).ForeColor = &HFF0000
End Sub

Private Sub mailto_Click()
Dim rc&
rc = ShellExecute(Me.hwnd, "open", mailto, "", "", 1)
End Sub

Private Sub mailto_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
mailto.FontUnderline = True
mailto.ForeColor = &HFF0000
End Sub

Private Sub sldPlayrate_Scroll()
Dim rc&
rc = mci.cmdRate("device1", sldPlayrate)
End Sub

Private Sub sldVolume_Scroll(Index As Integer)
Dim rc&, left%, right%
rc = mci.cmdVolume("device1", Index, Abs(sldVolume(Index).Value))
If rc <> 0 Then txtLog = txtLog + mci.ErrorDescription + vbCrLf

If Index = 2 Then
    rc = mci.cmdGetVolume("device1", 0, left)
    rc = mci.cmdGetVolume("device1", 1, right)
    sldVolume(0).Value = left * -1
    sldVolume(1).Value = right * -1
Else
    sldVolume(2) = (sldVolume(0) + sldVolume(1)) \ 2
End If
End Sub

Private Sub Slider1_Click()

End Sub

Private Sub Timer_Timer()
Dim curpos
curpos = mci.cmdGetPosition("device1")
lblCTime = convertTime(curpos / 1000)
lblCFrame = (curpos / 1000) * mci.FramesPerSec
sldTracker = curpos / 1000
lblStatus = mci.cmdGetStatus("device1")
If curpos = mci.TotalTime Then cmdStop_Click: cmdToStart_Click
End Sub

Private Sub txtLog_Change()
txtLog.SelStart = Len(txtLog)
End Sub

Public Function convertTime(ByVal convTime As Long) As String
Dim h As String, m As String, s As String
h = LTrim(Str(convTime \ 3600))
If Len(h) = 1 Then h = "0" + h
convTime = convTime Mod 3600
m = LTrim(Str(convTime \ 60))
If Len(m) = 1 Then m = "0" + m
convTime = convTime Mod 60
s = LTrim(Str(convTime))
If Len(s) = 1 Then s = "0" + s
convertTime = h + ":" + m + ":" + s
End Function

