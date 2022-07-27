Attribute VB_Name = "Take"
Option Explicit


Sub TakeS()
c = 1
c2 = 0

If ty < 7 Then
    Do While (grid(tx, ty + c) = NT)
        If (grid(tx, ty + c) = NT) And ty + c = 7 Then
            c2 = 0
            GoTo OutS
        End If
        c = c + 1
        c2 = c2 + 1
    Loop
End If

OutS:

If c2 <> 0 Then
    If grid(tx, ty + c) = Turn Then
        c = 0
        Do
            c = c + 1
            If Turn = 1 Then
                Othello.Tile(8 * (ty + c) + tx) = LoadPicture(App.Path & "/Graphx/RedTile.gif")
                Othello.Tile(8 * (ty + c) + tx).Tag = 1
                grid(tx, ty + c) = 1
            End If
            If Turn = 2 Then
                Othello.Tile(8 * (ty + c) + tx) = LoadPicture(App.Path & "/Graphx/BlueTile.gif")
                Othello.Tile(8 * (ty + c) + tx).Tag = 2
                grid(tx, ty + c) = 2
            End If
        Loop Until (c = c2)
    End If
End If
End Sub

Sub TakeN()
c = 1
c2 = 0

If ty > 0 Then
     Do While (grid(tx, ty - c) = NT)
        If (grid(tx, ty - c) = NT) And ty - c = 0 Then GoTo OutN
        c = c + 1
        c2 = c2 + 1
    Loop
End If

If c2 <> 0 Then
    If grid(tx, ty - c) = Turn Then
        c = 0
        Do
            c = c + 1
            If Turn = 1 Then
                Othello.Tile(8 * (ty - c) + tx) = LoadPicture(App.Path & "/Graphx/RedTile.gif")
                Othello.Tile(8 * (ty - c) + tx).Tag = 1
                grid(tx, ty - c) = 1
            End If
            If Turn = 2 Then
                Othello.Tile(8 * (ty - c) + tx) = LoadPicture(App.Path & "/Graphx/BlueTile.gif")
                Othello.Tile(8 * (ty - c) + tx).Tag = 2
                grid(tx, ty - c) = 2
            End If
        Loop Until (c = c2)
    End If
End If

OutN:

End Sub

Sub TakeE()
c = 1
c2 = 0

If tx < 8 Then
    Do While (grid(tx + c, ty) = NT)
        If (grid(tx + c, ty) = NT) And tx + c = 8 Then GoTo OutE
        c = c + 1
        c2 = c2 + 1
    Loop
End If

If c2 <> 0 Then
    If grid(tx + c, ty) = Turn Then
        c = 0
        Do
            c = c + 1
            If Turn = 1 Then
                Othello.Tile(8 * ty + (tx + c)) = LoadPicture(App.Path & "/Graphx/RedTile.gif")
                Othello.Tile(8 * ty + (tx + c)).Tag = 1
                grid(tx + c, ty) = 1
            End If
            If Turn = 2 Then
                Othello.Tile(8 * ty + (tx + c)) = LoadPicture(App.Path & "/Graphx/BlueTile.gif")
                Othello.Tile(8 * ty + (tx + c)).Tag = 2
                grid(tx + c, ty) = 2
            End If
        Loop Until (c = c2)
    End If
End If

OutE:

End Sub

Sub TakeW()
c = 1
c2 = 0

If tx > 1 Then
    Do While (grid(tx - c, ty) = NT)
        If (grid(tx - c, ty) = NT) And tx - c = 1 Then GoTo OutW
        c = c + 1
        c2 = c2 + 1
    Loop
End If

If c2 <> 0 Then
    If grid(tx - c, ty) = Turn Then
        c = 0
        Do
            c = c + 1
            If Turn = 1 Then
                Othello.Tile(8 * ty + (tx - c)) = LoadPicture(App.Path & "/Graphx/RedTile.gif")
                Othello.Tile(8 * ty + (tx - c)).Tag = 1
                grid(tx - c, ty) = 1
            End If
            If Turn = 2 Then
                Othello.Tile(8 * ty + (tx - c)) = LoadPicture(App.Path & "/Graphx/BlueTile.gif")
                Othello.Tile(8 * ty + (tx - c)).Tag = 2
                grid(tx - c, ty) = 2
            End If
        Loop Until (c = c2)
    End If
End If

OutW:

End Sub

Sub TakeSE()
c = 1
c2 = 0

If ty < 7 And tx < 8 Then
    Do While (grid(tx + c, ty + c) = NT)
        If (grid(tx + c, ty + c) = NT) And (ty + c = 7 Or tx + c = 8) Then GoTo OutSE
        c = c + 1
        c2 = c2 + 1
    Loop
End If

If c2 <> 0 Then
    If grid(tx + c, ty + c) = Turn Then
        c = 0
        Do
            c = c + 1
            If Turn = 1 Then
                Othello.Tile(8 * (ty + c) + (tx + c)) = LoadPicture(App.Path & "/Graphx/RedTile.gif")
                Othello.Tile(8 * (ty + c) + (tx + c)).Tag = 1
                grid(tx + c, ty + c) = 1
            End If
            If Turn = 2 Then
                Othello.Tile(8 * (ty + c) + (tx + c)) = LoadPicture(App.Path & "/Graphx/BlueTile.gif")
                Othello.Tile(8 * (ty + c) + (tx + c)).Tag = 2
                grid(tx + c, ty + c) = 2
            End If
        Loop Until (c = c2)
    End If
End If

OutSE:

End Sub

Sub TakeSW()
c = 1
c2 = 0

If ty < 7 And tx > 1 Then
    Do While (grid(tx - c, ty + c) = NT)
        If (grid(tx - c, ty + c) = NT) And (ty + c = 7 Or tx - c = 1) Then GoTo OutSW
        c = c + 1
        c2 = c2 + 1
    Loop
End If

If c2 <> 0 Then
    If grid(tx - c, ty + c) = Turn Then
        c = 0
        Do
            c = c + 1
            If Turn = 1 Then
                Othello.Tile(8 * (ty + c) + (tx - c)) = LoadPicture(App.Path & "/Graphx/RedTile.gif")
                Othello.Tile(8 * (ty + c) + (tx - c)).Tag = 1
                grid(tx - c, ty + c) = 1
            End If
            If Turn = 2 Then
                Othello.Tile(8 * (ty + c) + (tx - c)) = LoadPicture(App.Path & "/Graphx/BlueTile.gif")
                Othello.Tile(8 * (ty + c) + (tx - c)).Tag = 2
                grid(tx - c, ty + c) = 2
            End If
        Loop Until (c = c2)
    End If
End If

OutSW:

End Sub


Sub TakeNE()
c = 1
c2 = 0

If ty > 0 And tx < 8 Then
    Do While (grid(tx + c, ty - c) = NT)
    If (grid(tx + c, ty - c) = NT) And (ty - c = 0 Or tx + c = 8) Then GoTo OutNE
        c = c + 1
        c2 = c2 + 1
    Loop
End If

If c2 <> 0 Then
    If grid(tx + c, ty - c) = Turn Then
        c = 0
        Do
            c = c + 1
            If Turn = 1 Then
                Othello.Tile(8 * (ty - c) + (tx + c)) = LoadPicture(App.Path & "/Graphx/RedTile.gif")
                Othello.Tile(8 * (ty - c) + (tx + c)).Tag = 1
                grid(tx + c, ty - c) = 1
            End If
            If Turn = 2 Then
                Othello.Tile(8 * (ty - c) + (tx + c)) = LoadPicture(App.Path & "/Graphx/BlueTile.gif")
                Othello.Tile(8 * (ty - c) + (tx + c)).Tag = 2
                grid(tx + c, ty - c) = 2
            End If
        Loop Until (c = c2)
    End If
End If

OutNE:

End Sub

Sub TakeNW()
c = 1
c2 = 0

If ty > 0 And tx > 1 Then
    Do While (grid(tx - c, ty - c) = NT)
    If (grid(tx - c, ty - c) = NT) And (ty - c = 0 Or tx - c = 1) Then GoTo OutNW
        c = c + 1
        c2 = c2 + 1
    Loop
End If

If c2 <> 0 Then
    If grid(tx - c, ty - c) = Turn Then
        c = 0
        Do
            c = c + 1
            If Turn = 1 Then
                Othello.Tile(8 * (ty - c) + (tx - c)) = LoadPicture(App.Path & "/Graphx/RedTile.gif")
                Othello.Tile(8 * (ty - c) + (tx - c)).Tag = 1
                grid(tx - c, ty - c) = 1
            End If
            If Turn = 2 Then
                Othello.Tile(8 * (ty - c) + (tx - c)) = LoadPicture(App.Path & "/Graphx/BlueTile.gif")
                Othello.Tile(8 * (ty - c) + (tx - c)).Tag = 2
                grid(tx - c, ty - c) = 2
            End If
        Loop Until (c = c2)
    End If
End If

OutNW:

End Sub
