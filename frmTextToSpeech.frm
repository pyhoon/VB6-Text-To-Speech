VERSION 5.00
Begin VB.Form frmTextToSpeech 
   Caption         =   "Text To Speech Application"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   8205
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNormal 
      Caption         =   "Normal"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   6
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdFaster 
      Caption         =   "Faster"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdSlower 
      Caption         =   "Slower"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Text"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop Read"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton cmdSpeak 
      Caption         =   "Read Text"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox txtText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7935
   End
End
Attribute VB_Name = "frmTextToSpeech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents Voice As SpeechLib.SpVoice
Attribute Voice.VB_VarHelpID = -1
Dim strTitle As String

Private Sub cmdClear_Click()
    txtText.Text = ""
End Sub

Private Sub cmdSlower_Click()
    If Voice.Rate > -10 Then Voice.Rate = Voice.Rate - 1
    Me.Caption = strTitle & " (Speed: " & Voice.Rate & ")"
End Sub

Private Sub cmdNormal_Click()
    Voice.Rate = 0 ' -1
    Me.Caption = strTitle & " (Speed: " & Voice.Rate & ")"
End Sub

Private Sub cmdFaster_Click()
    If Voice.Rate < 10 Then Voice.Rate = Voice.Rate + 1
    Me.Caption = strTitle & " (Speed: " & Voice.Rate & ")"
End Sub

Private Sub Form_Load()
    strTitle = "Text To Speech Application"
    Set Voice = New SpeechLib.SpVoice
    'Voice.Rate = -1
    Me.Caption = strTitle & " (Speed: " & Voice.Rate & ")"
End Sub

Private Sub cmdSpeak_Click()
    Dim strText As String
    
    strText = txtText.Text
    If strText = "" Then Exit Sub
    
    If cmdSpeak.Caption = "Pause" Then
        cmdSpeak.Caption = "Resume"
        Voice.Pause
    ElseIf cmdSpeak.Caption = "Resume" Then
        cmdSpeak.Caption = "Pause"
        Voice.Resume
    Else
        cmdSpeak.Caption = "Pause"
        Voice.Speak strText, SVSFlagsAsync
    End If
End Sub

Private Sub cmdStop_Click()
    If cmdSpeak.Caption = "Pause" Then
        Voice.Skip "Sentence", 999
    End If
End Sub

Private Sub Voice_EndStream(ByVal StreamNumber As Long, ByVal StreamPosition As Variant)
    'Debug.Print "End"
    cmdSpeak.Caption = "Speak"
    cmdStop.Enabled = False
End Sub

Private Sub Voice_StartStream(ByVal StreamNumber As Long, ByVal StreamPosition As Variant)
    'Debug.Print "Start"
    cmdStop.Enabled = True
End Sub
