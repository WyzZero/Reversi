VERSION 5.00
Begin VB.Form Options 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   2925
   ClientLeft      =   555
   ClientTop       =   5175
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CheckBox Mon 
      BackColor       =   &H00000000&
      Caption         =   "Music On"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Value           =   1  'Checked
      Width           =   4335
   End
   Begin VB.CheckBox SPM 
      BackColor       =   &H00000000&
      Caption         =   "Show Possible Moves"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Value           =   1  'Checked
      Width           =   4335
   End
   Begin VB.TextBox Name2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3960
      TabIndex        =   3
      Text            =   "Player2"
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Name1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3960
      TabIndex        =   2
      Text            =   "Player1"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Player 2's Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Player 1's Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Name1.Text = Othello.P1namelbl.Caption
Name2.Text = Othello.P2namelbl.Caption

If musicon = True Then
    Mon.Value = 1
Else
    Mon.Value = 0
End If

If moveson = True Then
    SPM.Value = 1
Else
    SPM.Value = 0
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Where
End Sub

Private Sub Mon_Click()
If Mon.Value = 1 Then
    musicon = True
    Othello.Music.Enabled = True
    song = Int(Rnd * 6)
    Musicsel
ElseIf Mon.Value = 0 Then
    musicon = False
    Othello.Music.Stop
End If
End Sub

Private Sub Name1_Change()
Othello.P1namelbl.Caption = Name1.Text
Log2.histn1.Caption = Name1.Text
End Sub

Private Sub Name2_Change()
Othello.P2namelbl.Caption = Name2.Text
Log2.histn2.Caption = Name2.Text
End Sub

Private Sub SPM_Click()
If SPM.Value = 1 Then
    moveson = True
ElseIf SPM.Value = 0 Then
    moveson = False
End If

End Sub
