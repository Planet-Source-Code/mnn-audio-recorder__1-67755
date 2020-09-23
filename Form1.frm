VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Audio Recorder"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   720
      ScaleHeight     =   690
      ScaleWidth      =   6735
      TabIndex        =   19
      Top             =   2280
      Width           =   6735
      Begin VB.PictureBox picPeakRight 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   3105
         TabIndex        =   21
         Tag             =   "3870"
         Top             =   360
         Width           =   3105
      End
      Begin VB.PictureBox picPeakLeft 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   3105
         TabIndex        =   20
         Tag             =   "3870"
         Top             =   0
         Width           =   3105
      End
   End
   Begin VB.Timer Timer 
      Interval        =   1
      Left            =   2880
      Top             =   0
   End
   Begin VB.Frame fraStatus 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   3960
      TabIndex        =   10
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton btnHook 
         Caption         =   "Hook"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton btnUnhook 
         Caption         =   "Unhook"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton btnStop 
         Caption         =   "Stop"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   26
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton btnStart 
         Cancel          =   -1  'True
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton btnPause 
         Caption         =   "Pause"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   24
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0:00:00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2850
         TabIndex        =   17
         Top             =   960
         Width           =   570
      End
      Begin VB.Label lblTimeRecorded 
         AutoSize        =   -1  'True
         Caption         =   "Time recorded:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label lblBytesWritten 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0 bytes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2880
         TabIndex        =   14
         Top             =   600
         Width           =   540
      End
      Begin VB.Label lblBytes 
         AutoSize        =   -1  'True
         Caption         =   "Bytes written:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label lblState 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "not recording"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   2280
         TabIndex        =   12
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label lblStateLabel 
         AutoSize        =   -1  'True
         Caption         =   "State:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   450
      End
   End
   Begin VB.Frame fraRecSets 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.CheckBox chkFile 
         Caption         =   "Don't write to file"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton btnHelp 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   9
         Top             =   1320
         Width           =   230
      End
      Begin VB.TextBox txtFile 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         TabIndex        =   8
         Text            =   "c:\rec_%time%.wav"
         Top             =   1320
         Width           =   2415
      End
      Begin VB.ComboBox cmbBits 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   2520
         List            =   "Form1.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox cmbFrequency 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":0015
         Left            =   2520
         List            =   "Form1.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox cmbChannels 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Form1.frx":004E
         Left            =   2520
         List            =   "Form1.frx":0058
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         Caption         =   "File:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1365
         Width           =   300
      End
      Begin VB.Label lblBits 
         AutoSize        =   -1  'True
         Caption         =   "Bits per sec.:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1005
         Width           =   945
      End
      Begin VB.Label lblFrequency 
         AutoSize        =   -1  'True
         Caption         =   "Frequency:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   645
         Width           =   825
      End
      Begin VB.Label lblChannels 
         AutoSize        =   -1  'True
         Caption         =   "Channels:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   285
         Width           =   720
      End
   End
   Begin VB.TextBox txtDebug 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   3000
      Width           =   7575
   End
   Begin VB.Label lblLeft 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblRight 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2640
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

Dim bRecording As Boolean
Dim bMonitoring As Boolean

Private Sub btnHelp_Click()
    MsgBox "Variables:" & vbCrLf & "======" & vbCrLf & _
        "%date%" & vbTab & "- current date in system format (e.g. 12/30/2006)" & vbCrLf & _
        "%time%" & vbTab & "- current time in system format (e.g. 2:30 PM)", vbInformation, Caption
End Sub

Private Sub btnHook_Click()
    Hook True
End Sub

Private Sub btnPause_Click()
    bRecording = Not bRecording
    
    If bRecording Then
        With lblState
            .Caption = "RECORDING"
            .ForeColor = &HFF&
            .FontBold = True
        End With
    Else
        With lblState
            .Caption = "PAUSED"
            .ForeColor = &H8000&
            .FontBold = True
        End With
    End If
    
    modWaveIn.Pause
End Sub

Private Sub btnStart_Click()
    Dim File As String
    Dim channels As Integer
    
    EnableRecording
    
    File = ParseFile(txtFile.Text)
    
    If cmbChannels.ListIndex = 0 Then
        channels = 2
    ElseIf cmbChannels.ListIndex = 1 Then
        channels = 1
    End If
    
    If chkFile.Value = vbChecked Then File = ""
    
    modWaveIn.PrepareRecording channels, _
                               CInt(cmbBits.Text), _
                               CLng(cmbFrequency.Text), _
                               CStr(File)
End Sub

Private Sub btnStop_Click()
    DisableRecording
    
    Hook False
    modWaveIn.StopRec
    Hook True
End Sub

Private Sub btnUnhook_Click()
    Hook False
End Sub

Private Sub Form_Load()
    Dim channels As Integer
    Dim File As String
    
    cmbChannels.Text = cmbChannels.List(0)
    cmbBits.Text = cmbBits.List(0)
    cmbFrequency.Text = cmbFrequency.List(1)
    
    With lblState
        .Caption = "not recording"
        .ForeColor = &HFF0000
        .FontBold = False
    End With
    
    modWaveIn.msg = Space(Len(modWaveIn.msg))
    
    Hook True
    
    If Tag = "1" Then
        
        channels = 1
        If cmbChannels.ListIndex = 0 Then channels = 2
        
        File = ParseFile(txtFile.Text)
        If chkFile.Value = vbChecked Then File = ""
        
        EnableRecording
        
        modWaveIn.PrepareRecording channels, _
                                   CInt(cmbBits.Text), _
                                   CLng(cmbFrequency.Text), _
                                   File
    End If
End Sub

Private Sub Form_Resize()
    txtDebug.Move 0, txtDebug.Top, ScaleWidth, ScaleHeight - txtDebug.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Hook False
    modWaveIn.StopRec
End Sub

Sub DebugIt(ByVal inString As String)
    txtDebug.Text = txtDebug.Text & inString & vbCrLf
End Sub

Private Sub Timer_Timer()
    Static sngLast As Single
    Static sLast As Single
    Dim lLeft As Long
    Dim lRight As Long
    
    If bMonitoring Then
        modWaveIn.GetCurPeak lLeft, lRight
        
        If (VBA.Timer - sLast) >= 0.05 Then
            picPeakLeft.Width = (lLeft * CLng(Picture1.ScaleWidth)) \ modWaveIn.PeakMax
            picPeakRight.Width = (lRight * CLng(Picture1.ScaleWidth)) \ modWaveIn.PeakMax
            
            lblLeft.Caption = CStr(Round((lLeft / modWaveIn.PeakMax), 3))
            lblRight.Caption = CStr(Round((lRight / modWaveIn.PeakMax), 3))
            
            sLast = VBA.Timer
        End If
    End If
    
    If bRecording Then
    
        If (VBA.Timer - sngLast) >= 0.5 Then
            If lblState.Caption = "RECORDING" Then
                lblState.Caption = ""
            Else
                lblState.Caption = "RECORDING"
            End If
            
            sngLast = VBA.Timer
        End If
        
        lblTime.Caption = FormatTime(modWaveIn.GetTime)
        
        'lSample = (modWaveIn.SampleFrequency * 1000) \ (
    End If
    
    If Trim(modWaveIn.msg) <> "" Then
        DebugIt Trim(modWaveIn.msg)
        modWaveIn.msg = ""
    End If
    
    lblBytesWritten.Caption = CStr(modWaveIn.BytesWritten) & " bytes"
End Sub

Function FormatTime(ByVal lIn As Long, Optional ByVal bInMS As Boolean) As String
    Dim sec As String
    Dim min As String
    Dim tim As String
    Dim ms As String
    Dim hour As String
    
    ms = CStr(lIn)
    sec = CStr(ms)
    If bInMS Then sec = CStr(CLng(sec) \ 1000)
    
    
    hour = format(CStr(CLng(sec) \ 60 \ 60), "0#")
    min = format(CStr(CLng(sec) \ 60 - (CLng(hour) * 60)), "0#")
    sec = format(CStr(CLng(sec) - CLng(min) * 60 - (CLng(hour) * 60 * 60)), "0#")
    'ms = ms - (((CInt(min) * 60) + CInt(sec)) * 1000)
    
    tim = hour & ":" & min & ":" & sec
    
    FormatTime = tim
End Function

Sub EnableRecording()
    bRecording = True
    
    bMonitoring = True
    
    btnPause.Enabled = True
    btnStop.Enabled = True
    btnStart.Enabled = False
    
    cmbChannels.Enabled = False
    cmbFrequency.Enabled = False
    cmbBits.Enabled = False
    txtFile.Enabled = False
    chkFile.Enabled = False
    
    With lblState
        .Caption = "RECORDING"
        .ForeColor = &HFF&
        .FontBold = True
    End With
End Sub

Sub DisableRecording()
    bRecording = False
    
    bMonitoring = False
    
    btnPause.Enabled = False
    btnStop.Enabled = False
    btnStart.Enabled = True
    
    cmbChannels.Enabled = True
    cmbFrequency.Enabled = True
    cmbBits.Enabled = True
    txtFile.Enabled = True
    chkFile.Enabled = True
    
    With lblState
        .Caption = "not recording"
        .ForeColor = &HFF0000
        .FontBold = False
    End With
End Sub

Function ParseFile(ByVal File As String) As String
    ParseFile = Replace(File, "%date%", Replace(CStr(Date), ".", "-"))
    ParseFile = Replace(File, "%time%", Replace(CStr(Time), ":", "_"))
End Function
