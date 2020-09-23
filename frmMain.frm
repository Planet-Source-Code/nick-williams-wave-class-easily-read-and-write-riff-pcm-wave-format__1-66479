VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wave Example"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraFeedback 
      Caption         =   "Feedback:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   18
      Top             =   4680
      Width           =   6855
      Begin VB.TextBox txtFeedback 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.CommandButton cmdPlot 
      Caption         =   "Plot Waveform"
      Height          =   375
      Left            =   2520
      TabIndex        =   17
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "Save Wave As..."
      Height          =   375
      Left            =   4800
      TabIndex        =   16
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame fraWaveform 
      Caption         =   "Waveform:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   6855
      Begin VB.PictureBox picWaveform 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   120
         ScaleHeight     =   2145
         ScaleWidth      =   6585
         TabIndex        =   15
         Top             =   240
         Width           =   6615
      End
   End
   Begin VB.Frame fraWaveInfo 
      Caption         =   "Wave Information:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6855
      Begin VB.Label lblFormat 
         AutoSize        =   -1  'True
         Caption         =   "?"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   6000
         TabIndex        =   13
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblWaveLength 
         AutoSize        =   -1  'True
         Caption         =   "?"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   6000
         TabIndex        =   12
         Top             =   600
         Width           =   90
      End
      Begin VB.Label lblSamples 
         AutoSize        =   -1  'True
         Caption         =   "?"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   3960
         TabIndex        =   11
         Top             =   600
         Width           =   90
      End
      Begin VB.Label lblBitsPerSample 
         AutoSize        =   -1  'True
         Caption         =   "?"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   3960
         TabIndex        =   10
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Audio Format:"
         Height          =   255
         Index           =   7
         Left            =   4680
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Length:"
         Height          =   255
         Index           =   5
         Left            =   4680
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Samples:"
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblSampleRate 
         AutoSize        =   -1  'True
         Caption         =   "?"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1560
         TabIndex        =   6
         Top             =   600
         Width           =   90
      End
      Begin VB.Label lblChannels 
         AutoSize        =   -1  'True
         Caption         =   "?"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Bits Per Sample:"
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Sample Rate:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Channels:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open Wave File..."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Timer As clsPerformanceTimer
Private WAVE As clsWave
Private m_bytWave() As Byte
'For plotting the waveform quickly.
'MoveToEx has been modified from normal so that we don't have to pass an empty
'PointApi UDT to it since we won't use that parameter, instead we can just specify zero
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Single) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
'Credits Donald@xbeat.net - http://www.xbeat.net/vbspeed/c_GetFile.htm
Public Function GetFileName(strFileName As String) As String
    '
    Dim lngPos As Long
    Dim lngPosPrev As Long
    '
    Do
        '
        lngPosPrev = lngPos
        lngPos = InStr(lngPos + 1, strFileName, "\", vbBinaryCompare)
        '
    Loop While lngPos > 0
    '
    If lngPosPrev > 0 Then
        '
        GetFileName = Mid$(strFileName, lngPosPrev + 1)
        '
    Else
        '
        GetFileName = strFileName
        '
    End If
    '
End Function


Private Sub cmdOpen_Click()
    '
    On Error GoTo ErrHandler
    '
    Dim strFileName As String
    'Show the open dialog
    Dialog.Filter = "Wav Audio Files (*.wav)|*.wav|All Files (*.*)|*.*"
    Dialog.DialogTitle = "Open Wave File..."
    Dialog.ShowOpen
    'An error would have result if we cancelled, this is an actual selection
    strFileName = Dialog.FileName
    'Let's see how long it takes to open the file and get the data
    Timer.StartTimer
        '
        If Not WAVE.OpenWave(strFileName) Then GoTo ErrHandler
        '
    Timer.StopTimer
    '
    txtFeedback.Text = vbNullString
    '
    PrintFeedback "It took " & Int(Timer.TimeElapsed(pvMicroSecond)) & " microseconds to open the wav and return " & WAVE.WaveDataLength & " bytes to a string."
    '
    Timer.StartTimer
        '
        m_bytWave = WAVE.GetData
        '
    Timer.StopTimer
    '
    PrintFeedback "It took " & Int(Timer.TimeElapsed(pvMicroSecond)) & " microseconds to convert those bytes into a byte array using strConv."
    '
    Timer.StartTimer
        '
        m_bytWave = WAVE.GetData
        '
    Timer.StopTimer
    '
    PrintFeedback "Since the byte string is now cached, it only took " & Int(Timer.TimeElapsed(pvMicroSecond)) & " microseconds to return the data a second time."
    '
    fraWaveInfo.Caption = "Wave Information: " & GetFileName(strFileName)
    '
    lblChannels = WAVE.WaveChannels
    lblSampleRate = WAVE.WaveSampleRate & " HZ"
    lblBitsPerSample = WAVE.WaveBitsPerSample
    lblFormat = WAVE.WaveAudioFormat
    lblSamples = WAVE.WaveSamples
    lblWaveLength = WAVE.WaveLength & " ms"
    '
Exit Sub
    '
ErrHandler:
    '
    If Err.Number = 32755 Then
        'Cancel was selected via common dialog
    Else
        '
        MsgBox Err.Description, , Err.Number & "@frmMain.cmdOpen.Click"
        '
    End If
    '
End Sub


Private Sub cmdPlot_Click()
    '
    Call PlotWaveform(m_bytWave)
    '
End Sub

Private Sub cmdSaveAs_Click()
    '
    On Error Resume Next
    '
    Dialog.DialogTitle = "Save Wave as..."
    Dialog.Filter = "Wav Audio Files (*.wav)|*.wav"
    Dialog.ShowSave
    '
    Timer.StartTimer
        '
        If Not WAVE.WriteWave(m_bytWave, Dialog.FileName, 1, 8000, 8) Then GoTo ErrHandler
        '
    Timer.StopTimer
    '
    PrintFeedback "It took " & Int(Timer.TimeElapsed(pvMicroSecond)) & " microseconds to write " & UBound(m_bytWave) + 1 & " bytes."
    '
Exit Sub
    '
ErrHandler:
    '
    If Err.Number = 32755 Then
        'Cancel was selected via common dialog
    Else
        '
        MsgBox Err.Description, , Err.Number & "@frmMain.cmdSaveAs.Click"
        '
    End If
    '
End Sub

Private Sub Form_Load()
    '
    Set Timer = New clsPerformanceTimer
    Set WAVE = New clsWave
    '
End Sub




Public Sub PlotWaveform(bytData() As Byte)
    'The only possible error would be clicking plot before a file is opened
    'this would cause the assignment to lngDataLength to raise err 9.  So this is ok.
    On Error Resume Next
    '
    Dim lngDataLength As Long
    Dim lngIndex As Long
    '
    If WAVE.WaveChannels > 1 Or WAVE.WaveBitsPerSample > 8 Then
        '
        MsgBox "This example only plots single channel, 8BPS WAVE formats.", vbInformation, "Sorry!"
        Exit Sub
        '
    End If
    '
    lngDataLength = UBound(bytData)
    '
    If lngDataLength < 1 Then
        'This is where we catch the only possible error
        MsgBox "Data length is too short to plot, try opening a file first.", vbInformation, "Plotting Error"
        Exit Sub
        '
    End If
    '
    picWaveform.Cls
    picWaveform.ScaleWidth = lngDataLength
    picWaveform.ScaleHeight = 255
    picWaveform.PSet (0, 127), &H8000000D
    'We start from one because we will actually start drawing from lngIndex - 1 (zero)
    'due to the PSet function (assming it works like the MoveToEx API).
    For lngIndex = 1 To lngDataLength - 1
        '
        picWaveform.Line -(lngIndex, m_bytWave(lngIndex)), &H8000000D
        '
    Next lngIndex
    '
End Sub

Public Sub PrintFeedback(strFeedback As String)
    '
    txtFeedback.Text = txtFeedback.Text & strFeedback & vbCrLf
    txtFeedback.SelStart = Len(txtFeedback.Text)
    '
End Sub

