Attribute VB_Name = "LogStyle"
Option Explicit
Dim posY As String
Global m As Integer 'Which move is it?

Sub hist()

For Y = 0 To 7
    For X = 1 To 8
        If (8 * Y + X) = pos Then
            FindY
            Log1.History.Text = Log1.History.Text & m & ". " & posY & X & "  (Player " & Turn & ")" & vbNewLine
            GoTo Outh
        End If
    Next X
Next Y

Outh:

End Sub

Sub hist2()

For Y = 0 To 7
    For X = 1 To 8
        If (8 * Y + X) = pos Then
            FindY
            If Turn = 1 Then Log2.Player1.Text = Log2.Player1.Text & m & ". " & posY & X & vbNewLine
            If Turn = 2 Then Log2.Player2.Text = Log2.Player2.Text & m & ". " & posY & X & vbNewLine
            GoTo Outh2
        End If
    Next X
Next Y

Outh2:

End Sub

Sub FindY()

If Y = 0 Then posY = "A"
If Y = 1 Then posY = "B"
If Y = 2 Then posY = "C"
If Y = 3 Then posY = "D"
If Y = 4 Then posY = "E"
If Y = 5 Then posY = "F"
If Y = 6 Then posY = "G"
If Y = 7 Then posY = "H"

End Sub
