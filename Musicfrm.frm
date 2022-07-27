VERSION 5.00
Begin VB.Form Musicfrm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Music"
   ClientHeight    =   1215
   ClientLeft      =   510
   ClientTop       =   1215
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Musicfrm.frx":0000
   ScaleHeight     =   1215
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Tweak2 
      Caption         =   "Command1"
      Height          =   435
      Left            =   3460
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton Closebtn 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin VB.ComboBox Musicbox 
      Height          =   315
      ItemData        =   "Musicfrm.frx":2304
      Left            =   960
      List            =   "Musicfrm.frx":231A
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Musicfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------Used for Transparent Forms----------
Dim rgnBasic1 As New Region
Dim rgnExtended1 As New Region
Dim CurrentRgn1 As Long
Dim pic1(0 To 1) As New StdPicture
'----------Used for Transparent Forms----------


Private Sub Closebtn_Click()
Musicfrm.Visible = False

End Sub

Private Sub Form_Load()
'----------Used for Transparent Forms----------
    ' Load the image
    Set pic1(1) = LoadPicture(App.Path & "/Graphx/musicbg.gif", 0, 0, 0, 0)
    
    ' Scan the image
    Call rgnBasic1.ScanPicture(pic1(1))
    
    ' Offset the Shape to allow for the form header.
    Call rgnBasic1.OffsetHeader(Me)
        
    Me.picture = pic1(1) ' Set the Form Background
    Call rgnBasic1.ApplyRgn(Me.hWnd) ' Set the Form Shape
    CurrentRgn1 = rgnBasic1.hndRegion ' Set the Current Shape
'----------Used for Transparent Forms----------


End Sub

Private Sub Musicbox_Click()
If Musicbox.ListIndex <> song Then
    musicon = True
    song = Musicbox.ListIndex
    Musicsel
    Musicfrm.Visible = False
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
    Movedform3
End If
Tweak2.SetFocus
End Sub

