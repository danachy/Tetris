Attribute VB_Name = "tetris"

Option Base 1
Type tempbox
    num As Integer
    X(4) As Integer
    Y(4) As Integer
    r As Integer
    g As Integer
    b As Integer
    rot As Integer
End Type

Public nBox As tempbox
Public cBox As tempbox
Public tmp As tempbox


Type tempold
     X As Integer
     Y As Integer
     bl As Boolean
     r As Integer
     g As Integer
     b As Integer
End Type

Public old(200) As tempold

Type tempclean
     X As Integer
     Y As Integer
     bl As Boolean
     r As Integer
     g As Integer
     b As Integer
End Type

Public clean(200) As tempclean

Public score, state, boxwidth, pos, fx, fy As Integer

Public Function rotate(ByRef box As tempbox, ByVal bstate As Integer) As Boolean
    With box
        tmp.num = box.num
        If .num = 1 Or .num = 3 Or .num = 4 Or .num = 5 Or .num = 6 Then
           If state = 1 Then
              tmp.X(1) = .X(1) + boxwidth
              tmp.Y(1) = .Y(1) - boxwidth
           End If
           If state = 2 Then
              tmp.X(1) = .X(1) + boxwidth
              tmp.Y(1) = .Y(1) + boxwidth
           End If
           If state = 3 Then
              tmp.X(1) = .X(1) - boxwidth
              tmp.Y(1) = .Y(1) + boxwidth
           End If
           If state = 4 Then
              tmp.X(1) = .X(1) - boxwidth
              tmp.Y(1) = .Y(1) - boxwidth
           End If
        ElseIf .num = 2 Then
           Exit Function
        ElseIf .num = 7 Then
           If state = 1 Then
              tmp.X(1) = .X(1) - (2 * boxwidth)
              tmp.Y(1) = .Y(1) - (2 * boxwidth)
           End If
           If state = 2 Then
              tmp.X(1) = .X(1) + (2 * boxwidth)
              tmp.Y(1) = .Y(1) + (2 * boxwidth)
           End If
        End If
        
        calcBox tmp, bstate
        '*****************************
        For i = 1 To 4
            If tmp.X(i) < fx Then
               tmp.X(1) = tmp.X(1) + boxwidth
               calcBox tmp, bstate
            End If
            If tmp.X(i) + boxwidth = fx + (12 * boxwidth) Then
               tmp.X(1) = tmp.X(1) - (2 * boxwidth)
               calcBox tmp, bstate
            Else
               If tmp.X(i) + boxwidth > fx + (10 * boxwidth) Then
                  tmp.X(1) = tmp.X(1) - boxwidth
                  calcBox tmp, bstate
               End If
            End If
            If tmp.Y(i) + boxwidth >= fy + (20 * boxwidth) Then
               rotate = False
               Exit Function
            End If
        Next i
        
        '**********************
        For i = 1 To pos
            For j = 1 To 4
                If tmp.X(j) = old(i).X And tmp.Y(j) = old(i).Y Then
                   rotate = False
                   Exit Function
                End If
            Next j
        Next i
        '*********************
        For i = 1 To 4
            .X(i) = tmp.X(i)
            .Y(i) = tmp.Y(i)
        Next i
        rotate = True
   End With
End Function


Public Function hitleft(box As tempbox) As Boolean
   With box
        For i = 1 To 4
            If .X(i) = fx Then
                hitleft = True: Exit Function
            End If
        Next i
        hitleft = False
   End With
End Function


Public Function hitright(box As tempbox) As Boolean
   With box
        For i = 1 To 4
            If .X(i) + boxwidth = fx + (10 * boxwidth) Then
                hitright = True: Exit Function
            End If
        Next i
        hitright = False
   End With
End Function


Public Function hitbrixright(box As tempbox) As Boolean
   With box
   If pos <> 0 Then
      For i = 1 To 4
          For j = 1 To pos
              If .X(i) + boxwidth = old(j).X And .Y(i) = old(j).Y Then
                 hitbrixright = True: Exit Function
              End If
          Next j
      Next i
   End If
   hitbrixright = False
   End With
End Function


Public Function hitbrixleft(box As tempbox) As Boolean
   With box
   If pos <> 0 Then
      For i = 1 To 4
          For j = 1 To pos
              If .X(i) = old(j).X + boxwidth And .Y(i) = old(j).Y Then
                 hitbrixleft = True: Exit Function
              End If
          Next j
      Next i
   End If
   hitbrixleft = False
   End With
End Function

'L
Public Sub calcBox1(ByRef box As tempbox, bstate As Integer)
    With box
        Select Case bstate
        Case 1:
            .X(2) = .X(1)
            .Y(2) = .Y(1) + boxwidth
            .X(3) = .X(2)
            .Y(3) = .Y(2) + boxwidth
            .X(4) = .X(3) + boxwidth
            .Y(4) = .Y(3)
        Case 2:
            .X(2) = .X(1) - boxwidth
            .Y(2) = .Y(1)
            .X(3) = .X(2) - boxwidth
            .Y(3) = .Y(2)
            .X(4) = .X(3)
            .Y(4) = .Y(3) + boxwidth
        Case 3:
            .X(2) = .X(1)
            .Y(2) = .Y(1) - boxwidth
            .X(3) = .X(2)
            .Y(3) = .Y(2) - boxwidth
            .X(4) = .X(3) - boxwidth
            .Y(4) = .Y(3)
        Case 4:
            .X(2) = .X(1) + boxwidth
            .Y(2) = .Y(1)
            .X(3) = .X(2) + boxwidth
            .Y(3) = .Y(2)
            .X(4) = .X(3)
            .Y(4) = .Y(3) - boxwidth
        End Select
    End With
End Sub

'Ìï
Public Sub calcBox2(ByRef box As tempbox, bstate As Integer)
    With box
        .X(2) = .X(1) + boxwidth
        .Y(2) = .Y(1)
        .X(3) = .X(1)
        .Y(3) = .Y(1) + boxwidth
        .X(4) = .X(2)
        .Y(4) = .Y(3)
    End With
End Sub


'·´L
Public Sub calcBox3(ByRef box As tempbox, bstate As Integer)
    With box
        Select Case bstate
        Case 1:
            .X(2) = .X(1)
            .Y(2) = .Y(1) + boxwidth
            .X(3) = .X(2)
            .Y(3) = .Y(2) + boxwidth
            .X(4) = .X(3) - boxwidth
            .Y(4) = .Y(3)
        Case 2:
            .X(2) = .X(1) - boxwidth
            .Y(2) = .Y(1)
            .X(3) = .X(2) - boxwidth
            .Y(3) = .Y(2)
            .X(4) = .X(3)
            .Y(4) = .Y(3) - boxwidth
        Case 3:
            .X(2) = .X(1)
            .Y(2) = .Y(1) - boxwidth
            .X(3) = .X(2)
            .Y(3) = .Y(2) - boxwidth
            .X(4) = .X(3) + boxwidth
            .Y(4) = .Y(3)
        Case 4:
            .X(2) = .X(1) + boxwidth
            .Y(2) = .Y(1)
            .X(3) = .X(2) + boxwidth
            .Y(3) = .Y(2)
            .X(4) = .X(3)
            .Y(4) = .Y(3) + boxwidth
        End Select
    End With
End Sub

'·´N
Public Sub calcBox4(ByRef box As tempbox, bstate As Integer)
    With box
        Select Case bstate
        Case 1:
            .X(2) = .X(1)
            .Y(2) = .Y(1) + boxwidth
            .X(3) = .X(2) + boxwidth
            .Y(3) = .Y(2)
            .X(4) = .X(3)
            .Y(4) = .Y(3) + boxwidth
        Case 2:
            .X(2) = .X(1) - boxwidth
            .Y(2) = .Y(1)
            .X(3) = .X(2)
            .Y(3) = .Y(2) + boxwidth
            .X(4) = .X(3) - boxwidth
            .Y(4) = .Y(3)
        Case 3:
            .X(2) = .X(1)
            .Y(2) = .Y(1) - boxwidth
            .X(3) = .X(2) - boxwidth
            .Y(3) = .Y(2)
            .X(4) = .X(3)
            .Y(4) = .Y(3) - boxwidth
        Case 4:
            .X(2) = .X(1) + boxwidth
            .Y(2) = .Y(1)
            .X(3) = .X(2)
            .Y(3) = .Y(2) - boxwidth
            .X(4) = .X(3) + boxwidth
            .Y(4) = .Y(3)
        End Select
    End With
End Sub

'N
Public Sub calcBox5(ByRef box As tempbox, bstate As Integer)
    With box
        Select Case bstate
        Case 1:
            .X(2) = .X(1)
            .Y(2) = .Y(1) + boxwidth
            .X(3) = .X(2) - boxwidth
            .Y(3) = .Y(2)
            .X(4) = .X(3)
            .Y(4) = .Y(3) + boxwidth
        Case 2:
            .X(2) = .X(1) - boxwidth
            .Y(2) = .Y(1)
            .X(3) = .X(2)
            .Y(3) = .Y(2) - boxwidth
            .X(4) = .X(3) - boxwidth
            .Y(4) = .Y(3)
        Case 3:
            .X(2) = .X(1)
            .Y(2) = .Y(1) - boxwidth
            .X(3) = .X(2) + boxwidth
            .Y(3) = .Y(2)
            .X(4) = .X(3)
            .Y(4) = .Y(3) - boxwidth
        Case 4:
            .X(2) = .X(1) + boxwidth
            .Y(2) = .Y(1)
            .X(3) = .X(2)
            .Y(3) = .Y(2) + boxwidth
            .X(4) = .X(3) + boxwidth
            .Y(4) = .Y(3)
        End Select
    End With
End Sub

'T
Public Sub calcBox6(ByRef box As tempbox, bstate As Integer)
    With box
        Select Case bstate
        Case 1:
            .X(2) = .X(1)
            .Y(2) = .Y(1) + boxwidth
            .X(3) = .X(2) + boxwidth
            .Y(3) = .Y(2)
            .X(4) = .X(2) - boxwidth
            .Y(4) = .Y(2)
        Case 2:
            .X(2) = .X(1) - boxwidth
            .Y(2) = .Y(1)
            .X(3) = .X(2)
            .Y(3) = .Y(2) + boxwidth
            .X(4) = .X(2)
            .Y(4) = .Y(2) - boxwidth
        Case 3:
            .X(2) = .X(1)
            .Y(2) = .Y(1) - boxwidth
            .X(3) = .X(2) + boxwidth
            .Y(3) = .Y(2)
            .X(4) = .X(2) - boxwidth
            .Y(4) = .Y(2)
        Case 4:
            .X(2) = .X(1) + boxwidth
            .Y(2) = .Y(1)
            .X(3) = .X(2)
            .Y(3) = .Y(2) - boxwidth
            .X(4) = .X(2)
            .Y(4) = .Y(2) + boxwidth
        End Select
    End With
End Sub


'I ÐÍ
Public Sub calcBox7(ByRef box As tempbox, bstate As Integer)
    With box
        Select Case bstate
        Case 1:
            .X(2) = .X(1)
            .Y(2) = .Y(1) + boxwidth
            .X(3) = .X(2)
            .Y(3) = .Y(2) + boxwidth
            .X(4) = .X(3)
            .Y(4) = .Y(3) + boxwidth
        Case 2:
            .X(2) = .X(1) - boxwidth
            .Y(2) = .Y(1)
            .X(3) = .X(2) - boxwidth
            .Y(3) = .Y(2)
            .X(4) = .X(3) - boxwidth
            .Y(4) = .Y(3)
        Case 3:
            .X(2) = .X(1)
            .Y(2) = .Y(1) + boxwidth
            .X(3) = .X(2)
            .Y(3) = .Y(2) + boxwidth
            .X(4) = .X(3)
            .Y(4) = .Y(3) + boxwidth
        Case 4:
            .X(2) = .X(1) - boxwidth
            .Y(2) = .Y(1)
            .X(3) = .X(2) - boxwidth
            .Y(3) = .Y(2)
            .X(4) = .X(3) - boxwidth
            .Y(4) = .Y(3)
        End Select
    End With
End Sub



Public Sub calcBox(ByRef box As tempbox, ByVal bstate As Integer)
    Select Case box.num
    Case 1:
        calcBox1 box, bstate
    Case 2:
        calcBox2 box, bstate
    Case 3:
        calcBox3 box, bstate
    Case 4:
        calcBox4 box, bstate
    Case 5:
        calcBox5 box, bstate
    Case 6:
        calcBox6 box, bstate
    Case 7:
        calcBox7 box, bstate
    End Select
End Sub



