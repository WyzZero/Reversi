Attribute VB_Name = "Others"
Option Explicit


Sub EmptyBox()
For y = 0 To 7
    For x = 1 To 8
        If Othello.Tile(8 * y + x).Tag = 4 Then
            Othello.Tile(8 * y + x).Tag = 0
            Othello.Tile(8 * y + x) = LoadPicture(App.Path & "/Graphx/Tile.gif")
            grid(x, y) = 0
        End If
        If Othello.Tile(8 * y + x).Tag = Turn And grid(x, y) <> Turn Then
            grid(x, y) = Turn
            tx = x
            ty = y
        End If
    Next x
Next y
End Sub

Sub SkipC()
Skip = True
For y = 0 To 7
    For x = 1 To 8
        If Othello.Tile(8 * y + x).Tag = 4 Then Skip = False
    Next x
Next y

If Skip = True Then
    If Turn = 1 Then
        Turn = 2
        NT = 1
    ElseIf Turn = 2 Then
        Turn = 1
        NT = 2
    End If
    If Skips = 1 Then
        MsgBox "Game Over!"
        Exit Sub
    End If
    Skips = 1
    Where
Else
    Skips = 0
End If

End Sub

Sub Score()
RScore = 0
BScore = 0

For y = 0 To 7
    For x = 1 To 8
        If Othello.Tile(8 * y + x).Tag = 1 Then RScore = RScore + 1
        If Othello.Tile(8 * y + x).Tag = 2 Then BScore = BScore + 1
    Next x
Next y
Othello.Bluelbl.Caption = BScore
Othello.Redlbl.Caption = RScore

End Sub


Sub Movedform()
    Othello.Left = Othello.Left - (OrigX - NewX)
    Othello.Top = Othello.Top - (OrigY - NewY)
End Sub
Sub Movedform1()
    Log1.Left = Log1.Left - (OrigX - NewX)
    Log1.Top = Log1.Top - (OrigY - NewY)
End Sub
Sub Movedform2()
    Log2.Left = Log2.Left - (OrigX - NewX)
    Log2.Top = Log2.Top - (OrigY - NewY)
End Sub

Sub Movedform3()
    Musicfrm.Left = Musicfrm.Left - (OrigX - NewX)
    Musicfrm.Top = Musicfrm.Top - (OrigY - NewY)
End Sub

Sub Musicsel()


Select Case song
Case 0
    Othello.Music.FileName = (App.Path & "/Music/Between_Realities.mp3")
    song = 0
Case 1
    Othello.Music.FileName = (App.Path & "/Music/Floating_and_Dreaming.mp3")
    song = 1
Case 2
    Othello.Music.FileName = (App.Path & "/Music/Othello.mp3")
    song = 2
Case 3
    Othello.Music.FileName = (App.Path & "/Music/OthelloUK.mp3")
    song = 3
Case 4
    Othello.Music.FileName = (App.Path & "/Music/The_Stonecrest_Visitat.mp3")
    song = 4
Case 5
    Othello.Music.FileName = (App.Path & "/Music/Winds.mp3")
    song = 5
End Select

End Sub

