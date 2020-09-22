Attribute VB_Name = "Module1"

'The point type
Public Type Point
    X As Integer
    Y As Integer
End Type

'the code type
Public Type Code
    c(4) As Boolean
End Type

'the line type
Public Type line
    p1 As Point
    p2 As Point
    c As Integer
End Type

Public Lines(100) As line      'All the lines
Public n As Integer            'the number of lines
'add line to the array
Function addLine(l As line)
    'MsgBox l.c
    Lines(n) = l
    n = n + 1
    Form1.List1.AddItem (Str$(n) + ". " + toString(l))
End Function
Function toString(l As line) As String
    toString = "(" + Str$(l.p1.X) + "," + Str$(l.p1.Y) + ") - (" + Str$(l.p2.X) + "," + Str$(l.p2.Y) + ")"
End Function

Function refreshDisplay(p As PictureBox)
    p.Cls
    For i = 1 To n
        drawLine Lines(i - 1), 0, p
    Next i
End Function

Function drawLine(l As line, c As Integer, p As PictureBox)
    If c = 7 Then
        p.Line (l.p1.X, l.p1.Y)-(l.p2.X, l.p2.Y), QBColor(c)
    Else
        p.Line (l.p1.X, l.p1.Y)-(l.p2.X, l.p2.Y), QBColor(l.c)
    End If
End Function

Function drawBox(l As line, p As PictureBox)
    p.Line (l.p1.X, l.p1.Y)-(l.p2.X, l.p2.Y), QBColor(0), BF
End Function



Function accept(c1 As Code, c2 As Code) As Boolean
    accept = True
    For i = 1 To 4
        If (c1.c(i) Or c2.c(i)) Then accept = False
    Next i
End Function

Function reject(c1 As Code, c2 As Code) As Boolean
    reject = False
    For i = 1 To 4
        If (c1.c(i) And c2.c(i)) Then reject = True
    Next i
End Function

Function isInside(c As Code) As Boolean
    If (c.c(1) And c.c(1) And c.c(1) And c.c(1)) Then
        isInside = False
    Else
        isInside = True
    End If
End Function

Function swapPts(p1 As Point, p2 As Point)
    Dim temp As Point
    temp = p1
    p1 = p2
    p2 = temp
End Function

Function swapCodes(c1 As Code, c2 As Code)
    Dim temp As Code
    temp = c1
    c1 = c2
    c2 = temp
End Function

Function getCode(p As Point, region As line) As Code
    If p.X < region.p1.X Then getCode.c(1) = True Else getCode.c(1) = False
    If p.X > region.p2.X Then getCode.c(2) = True Else getCode.c(2) = False
    If p.Y < region.p1.Y Then getCode.c(3) = True Else getCode.c(3) = False
    If p.Y > region.p2.Y Then getCode.c(4) = True Else getCode.c(4) = False
End Function

'-------------------------------------------------------------------------------'
'                                                                               '
'   Clipping routine for Cohen-SutherLand                                       '
'                                                                               '
'                                                                               '
'-------------------------------------------------------------------------------'


Function clipLine(l As line, r As line)
    Dim p1 As Point
    Dim p2 As Point
    Dim c1 As Code
    Dim c2 As Code
    Dim t As line
    Dim done As Boolean
    Dim draw As Boolean
    Dim m As Variant
    
    'MsgBox toString(l)
    fixRegion r
    
    p1 = l.p1
    p2 = l.p2
    
    done = False
    draw = False
    
    While done = False
        c1 = getCode(p1, r)
        c2 = getCode(p2, r)
        
        If accept(c1, c2) Then
            done = True
            draw = True
        ElseIf reject(c1, c2) Then
            done = True
        Else
            If isInside(c1) Then
                swapPts p1, p2
                swapCodes c1, c2
            End If
            
            m = (p2.Y - p1.Y) / (p2.X - p1.X)
            If c1.c(1) Then
                'crosses left
                p1.Y = p1.Y + (r.p1.X - p1.X) * m
                p1.X = r.p1.X
                
            ElseIf c1.c(2) Then
                'crosses right
                p1.Y = p1.Y + (r.p2.X - p1.X) * m
                p1.X = r.p2.X
                
            ElseIf c1.c(3) Then
                'crosses bottom
                p1.X = p1.X + (r.p1.Y - p1.Y) / m
                p1.Y = r.p1.Y
                
            ElseIf c1.c(4) Then
                'crosses bottom
                p1.X = p1.X + (r.p2.Y - p1.Y) / m
                p1.Y = r.p2.Y
            End If
        End If
    Wend
    
    t.p1 = p1
    t.p2 = p2
    t.c = l.c
    
    If draw Then
        drawLine t, 0, Form1.Picture1
        Form1.List2.AddItem (toString(t))
    End If
End Function

Function clipCohSuth(region As line, p As PictureBox)
Form1.Picture1.Cls
Form1.List2.Clear

p.DrawWidth = 1
p.DrawStyle = 2
p.Line (0, region.p1.Y)-(p.Width, region.p1.Y), QBColor(7)
p.Line (0, region.p2.Y)-(p.Width, region.p2.Y), QBColor(7)
p.Line (region.p1.X, 0)-(region.p1.X, p.Height), QBColor(7)
p.Line (region.p2.X, 0)-(region.p2.X, p.Height), QBColor(7)
p.DrawStyle = 1
For i = 0 To n
    drawLine Lines(i), 7, p
Next i
p.DrawStyle = 0
p.DrawWidth = 2

Form1.t1.Caption = Str$(GetTickCount)
For i = 0 To n
    'MsgBox toString(Lines(i))
    clipLine Lines(i), region
Next i
Form1.t2.Caption = Str$(GetTickCount)
End Function


' Fixes the region so that p1 is xmin and ymin and p2 is
' xmax and ymax. This is done because the user can select
' the clipping area with the mouse in any order

Function fixRegion(r As line)
If (r.p1.X > r.p2.X) Then
    temp = r.p1.X
    r.p1.X = r.p2.X
    r.p2.X = temp
End If

If (r.p1.Y > r.p2.Y) Then
    temp = r.p1.Y
    r.p1.Y = r.p2.Y
    r.p2.Y = temp
End If

End Function





'===============================================================================
'
'
'     Liang-Barsky Clipping
'
'===============================================================================



Function clipTest(p As Double, q As Double, u1 As Double, u2 As Double) As Boolean
    Dim r As Double
    clipTest = True
    
    If p < 0 Then
        r = q / p
        If r > u2 Then
            clipTest = False
        ElseIf r > u1 Then
            u1 = r
        End If
    Else
        If p > 0 Then
            r = q / p
            If r < u1 Then
                clipTest = False
            ElseIf r < u2 Then
                u2 = r
            End If
        ElseIf q < 0 Then
            clipTest = False
        End If
    End If

End Function

Function clipLine2(l As line, r As line, p As PictureBox)
    Dim u1 As Double
    Dim u2 As Double
    Dim dx As Double
    Dim dy As Double
    Dim p1 As Point
    Dim p2 As Point
    Dim t As line
    
    p1 = l.p1
    p2 = l.p2
    
    
    u1 = 0
    u2 = 1
    dx = p2.X - p1.X
    
    If clipTest(-1 * dx, (p1.X - r.p1.X), u1, u2) Then
        If clipTest(dx, (r.p2.X - p1.X), u1, u2) Then
            dy = p2.Y - p1.Y
            If clipTest(-1 * dy, (p1.Y - r.p1.Y), u1, u2) Then
                If clipTest(dy, (r.p2.Y - p1.Y), u1, u2) Then
                    If u2 < 1 Then
                        p2.X = p1.X + (u2 * dx)
                        p2.Y = p1.Y + (u2 * dy)
                    End If
                    If u1 > 0 Then
                        p1.X = p1.X + (u1 * dx)
                        p1.Y = p1.Y + (u1 * dy)
                    End If
                    
                    t.p1 = p1
                    t.p2 = p2
                    t.c = l.c
                    'draw line here
                    drawLine t, 0, p
                    Form1.List2.AddItem (toString(t))
                End If
            End If
        End If
    End If
    
    
    
ending:

                 

End Function

Function clipLiangBarsky(region As line, p As PictureBox)
Form1.Picture1.Cls
Form1.List2.Clear

p.DrawWidth = 1
p.DrawStyle = 2
p.Line (0, region.p1.Y)-(p.Width, region.p1.Y), QBColor(7)
p.Line (0, region.p2.Y)-(p.Width, region.p2.Y), QBColor(7)
p.Line (region.p1.X, 0)-(region.p1.X, p.Height), QBColor(7)
p.Line (region.p2.X, 0)-(region.p2.X, p.Height), QBColor(7)
p.DrawStyle = 1
For i = 0 To n
    drawLine Lines(i), 7, p
Next i
p.DrawStyle = 0
p.DrawWidth = 2

fixRegion region

For i = 0 To (n - 1)
    clipLine2 Lines(i), region, p
Next i

End Function
