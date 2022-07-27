Attribute VB_Name = "Check"
Option Explicit

Sub Where()

If Turn = 1 Then
    Othello.Turnlbl.Caption = Othello.P1namelbl.Caption & "'s Turn!"
ElseIf Turn = 2 Then
    Othello.Turnlbl.Caption = Othello.P2namelbl.Caption & "'s Turn!"
End If

For y = 0 To 7
    For x = 1 To 8
        If grid(x, y) = Turn Then
            If y > 1 Then CheckN
            If y < 6 Then CheckS
            If x > 2 Then CheckW
            If x < 7 Then CheckE
            If y > 1 And x > 2 Then CheckNW
            If y > 1 And x < 7 Then CheckNE
            If y < 6 And x > 2 Then CheckSW
            If y < 6 And x < 7 Then CheckSE
        End If
    Next x
Next y
SkipC


End Sub

Sub CheckN()
c = 0
    
Do
    c = c + 1
    If grid(x, y - c) = NT And grid(x, y - (c + 1)) = 0 Then
        If moveson = True Then Othello.Tile(8 * (y - (c + 1)) + x) = LoadPicture(App.Path & "/Graphx/movetile.gif")
        If moveson = False Then Othello.Tile(8 * (y - (c + 1)) + x) = LoadPicture(App.Path & "/Graphx/tile.gif")
        Othello.Tile(8 * (y - (c + 1)) + x).Tag = 4
    End If
Loop Until grid(x, y - c) <> NT Or grid(x, y - (c + 1)) <> NT Or y - (c + 1) = 0
End Sub


Sub CheckS()
c = 0
    
Do
    c = c + 1
    If grid(x, y + c) = NT And grid(x, y + (c + 1)) = 0 Then
        If moveson = True Then Othello.Tile(8 * (y + (c + 1)) + x) = LoadPicture(App.Path & "/Graphx/movetile.gif")
        If moveson = False Then Othello.Tile(8 * (y + (c + 1)) + x) = LoadPicture(App.Path & "/Graphx/tile.gif")
        Othello.Tile(8 * (y + (c + 1)) + x).Tag = 4
    End If
Loop Until grid(x, y + c) <> NT Or grid(x, y + (c + 1)) <> NT Or y + (c + 1) = 7
End Sub

Sub CheckW()
c = 0

Do
    c = c + 1
    If grid(x - c, y) = NT And grid(x - (c + 1), y) = 0 Then
        If moveson = True Then Othello.Tile(8 * y + (x - (c + 1))) = LoadPicture(App.Path & "/Graphx/movetile.gif")
        If moveson = False Then Othello.Tile(8 * y + (x - (c + 1))) = LoadPicture(App.Path & "/Graphx/tile.gif")
        Othello.Tile(8 * y + (x - (c + 1))).Tag = 4
    End If
Loop Until grid(x - c, y) <> NT Or grid(x - (c + 1), y) <> NT Or x - (c + 1) = 1
End Sub

Sub CheckE()
c = 0
    
Do
    c = c + 1
    If grid(x + c, y) = NT And grid(x + (c + 1), y) = 0 Then
        If moveson = True Then Othello.Tile(8 * y + x + (c + 1)) = LoadPicture(App.Path & "/Graphx/movetile.gif")
        If moveson = False Then Othello.Tile(8 * y + x + (c + 1)) = LoadPicture(App.Path & "/Graphx/tile.gif")
        Othello.Tile(8 * y + x + (c + 1)).Tag = 4
    End If
Loop Until grid(x + c, y) <> NT Or grid(x + (c + 1), y) <> NT Or x + (c + 1) = 8
End Sub

Sub CheckNW()
c = 0
    
Do
    c = c + 1
    If grid(x - c, y - c) = NT And grid(x - (c + 1), y - (c + 1)) = 0 Then
        If moveson = True Then Othello.Tile(8 * (y - (c + 1)) + x - (c + 1)) = LoadPicture(App.Path & "/Graphx/movetile.gif")
        If moveson = False Then Othello.Tile(8 * (y - (c + 1)) + x - (c + 1)) = LoadPicture(App.Path & "/Graphx/tile.gif")
        Othello.Tile(8 * (y - (c + 1)) + x - (c + 1)).Tag = 4
    End If
Loop Until grid(x - c, y - c) <> NT Or grid(x - (c + 1), y - (c + 1)) <> NT Or y - (c + 1) = 0 Or x - (c + 1) = 1
End Sub

Sub CheckNE()
c = 0
    
Do
    c = c + 1
    If grid(x + c, y - c) = NT And grid(x + (c + 1), y - (c + 1)) = 0 Then
        If moveson = True Then Othello.Tile(8 * (y - (c + 1)) + x + (c + 1)) = LoadPicture(App.Path & "/Graphx/movetile.gif")
        If moveson = False Then Othello.Tile(8 * (y - (c + 1)) + x + (c + 1)) = LoadPicture(App.Path & "/Graphx/tile.gif")
        Othello.Tile(8 * (y - (c + 1)) + x + (c + 1)).Tag = 4
    End If
Loop Until grid(x + c, y - c) <> NT Or grid(x + (c + 1), y - (c + 1)) <> NT Or y - (c + 1) = 0 Or x + (c + 1) = 8
End Sub

Sub CheckSW()
c = 0
    
Do
    c = c + 1
    If grid(x - c, y + c) = NT And grid(x - (c + 1), y + (c + 1)) = 0 Then
        If moveson = True Then Othello.Tile(8 * (y + (c + 1)) + x - (c + 1)) = LoadPicture(App.Path & "/Graphx/movetile.gif")
        If moveson = False Then Othello.Tile(8 * (y + (c + 1)) + x - (c + 1)) = LoadPicture(App.Path & "/Graphx/tile.gif")
        Othello.Tile(8 * (y + (c + 1)) + x - (c + 1)).Tag = 4
    End If
Loop Until grid(x - c, y + c) <> NT Or grid(x - (c + 1), y + (c + 1)) <> NT Or y + (c + 1) = 7 Or x - (c + 1) = 1
End Sub

Sub CheckSE()
c = 0
    
Do
    c = c + 1
    If grid(x + c, y + c) = NT And grid(x + (c + 1), y + (c + 1)) = 0 Then
        If moveson = True Then Othello.Tile(8 * (y + (c + 1)) + x + (c + 1)) = LoadPicture(App.Path & "/Graphx/movetile.gif")
        If moveson = False Then Othello.Tile(8 * (y + (c + 1)) + x + (c + 1)) = LoadPicture(App.Path & "/Graphx/tile.gif")
        Othello.Tile(8 * (y + (c + 1)) + x + (c + 1)).Tag = 4
    End If
Loop Until grid(x + c, y + c) <> NT Or grid(x + (c + 1), y + (c + 1)) <> NT Or y + (c + 1) = 7 Or x + (c + 1) = 8
End Sub
