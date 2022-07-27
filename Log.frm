VERSION 5.00
Begin VB.Form Log1 
   BorderStyle     =   0  'None
   Caption         =   "Othello"
   ClientHeight    =   6465
   ClientLeft      =   510
   ClientTop       =   2595
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Log.frx":0000
   ScaleHeight     =   6465
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox History 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   935
      Width           =   2535
   End
   Begin VB.CommandButton Tweak2 
      Caption         =   "Command1"
      Height          =   195
      Left            =   3600
      TabIndex        =   1
      Top             =   6480
      Width           =   375
   End
   Begin VB.ComboBox Stylebox 
      Height          =   315
      ItemData        =   "Log.frx":ABAF
      Left            =   1200
      List            =   "Log.frx":ABBC
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   6000
      Width           =   1215
   End
End
Attribute VB_Name = "Log1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Log1.picture = LoadPicture(App.Path & "\Graphx\Log1bg.gif")
Stylebox.ListIndex = 0
Ssel = 1
End Sub

Private Sub Form_Resize()


If Log1.WindowState = 1 Then
    Othello.Visible = False
    If Ssel = 2 Then Log2.Visible = False
ElseIf Log1.WindowState = 0 Then
    If Ssel = 2 Then
        Log1.Visible = False
        Log2.Visible = True
    ElseIf Ssel = 0 Then Log1.Visible = False
    End If
    Othello.Visible = True
End If

End Sub



Private Sub History_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Tweak2.SetFocus
End Sub

Private Sub Stylebox_Click()
If Stylebox.ListIndex = 1 Then
    Log2.Show
    Log1.Hide
    Log2.Stylebox.ListIndex = 1
    Ssel = 2
ElseIf Stylebox.ListIndex = 2 Then
    Log1.Hide
    Ssel = 0
End If
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
repos = True
OrigX = X
OrigY = Y
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If repos = True Then
    moved = True
End If
Tweak2.SetFocus
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
repos = False
If moved = True Then
    NewX = X
    NewY = Y
    Movedform1
End If
Tweak2.SetFocus
End Sub
