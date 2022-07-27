Attribute VB_Name = "Variables"
Option Explicit

Global Ssel As Integer 'History Style in use (For minimization)
Global repos As Boolean

Global x As Integer 'Used for loops
Global y As Integer 'Used for loops

Global c As Integer 'Used for checks
Global c2 As Integer 'Used for checks

Global ty As Integer 'y = tile click
Global tx As Integer 'x = tile click

Global Skip As Boolean
Global Skips As Integer

Global RScore As Integer 'Red Score
Global BScore As Integer 'Blue Score

Global pos As Integer 'Clicked

'0 = Empty / 1 = Red / 2 = Blue
Global grid(1 To 8, 0 To 7) As Integer '(x,y)

Global Turn As Integer '1 = Red / 2 = Blue
Global NT As Integer '1 = Red / 2 = Blue

Global moved As Boolean
Global OrigX As Single
Global OrigY As Single
Global NewX As Single
Global NewY As Single

Global song As Integer

Global musicon As Boolean
Global moveson As Boolean

