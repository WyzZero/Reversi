VERSION 5.00
Begin VB.Form Log2 
   BorderStyle     =   0  'None
   Caption         =   "History"
   ClientHeight    =   6465
   ClientLeft      =   690
   ClientTop       =   3105
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Log2.frx":0000
   ScaleHeight     =   6465
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Tweak2 
      Caption         =   "Command1"
      Height          =   255
      Left            =   6960
      TabIndex        =   5
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Player2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4935
      Left            =   3780
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   935
      Width           =   2535
   End
   Begin VB.TextBox Player1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   4935
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   935
      Width           =   2535
   End
   Begin VB.ComboBox Stylebox 
      Height          =   315
      ItemData        =   "Log2.frx":12565
      Left            =   2880
      List            =   "Log2.frx":12572
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label histn1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   5880
      Width           =   2535
   End
   Begin VB.Label histn2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player 2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   615
      Left            =   4200
      TabIndex        =   0
      Top             =   5880
      Width           =   2535
   End
End
Attribute VB_Name = "Log2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Log1.picture = LoadPicture(App.Path & "\Graphx\Log1bg.gif")
Stylebox.ListIndex = 1
End Sub


Private Sub Player1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Tweak2.SetFocus
End Sub

Private Sub Player2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Tweak2.SetFocus
End Sub

Private Sub Stylebox_Click()
If Stylebox.ListIndex = 0 Then
    Log1.Show
    Log2.Hide
    Log1.Stylebox.ListIndex = 0
    Ssel = 1
ElseIf Stylebox.ListIndex = 2 Then
    Log2.Hide
    Ssel = 0
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
repos = True
OrigX = x
OrigY = y
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If repos = True Then
    moved = True
End If
Tweak2.SetFocus
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
repos = False
If moved = True Then
    NewX = x
    NewY = y
    Movedform2
End If
Tweak2.SetFocus
End Sub
