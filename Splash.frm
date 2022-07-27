VERSION 5.00
Begin VB.Form Splash 
   BorderStyle     =   0  'None
   ClientHeight    =   2880
   ClientLeft      =   165
   ClientTop       =   2250
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2500
      Left            =   1560
      Top             =   1320
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
Othello.Show
Unload Me

End Sub

Private Sub Form_Load()
Splash.picture = LoadPicture(App.Path & "\graphx\Wyz.jpg")

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Othello.Show
Unload Me
End Sub

Private Sub Timer1_Timer()
Othello.Show
Unload Me

End Sub
