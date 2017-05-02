VERSION 5.00
Begin VB.Form formMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "俄罗斯方块"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7905
   FillStyle       =   0  'Solid
   LinkTopic       =   "fromMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   7905
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame manualFrame 
      Caption         =   "操作说明"
      Height          =   5000
      Left            =   5800
      TabIndex        =   1
      Top             =   600
      Width           =   2000
      Begin VB.Label Label7 
         Caption         =   "ESC -- 退出"
         Height          =   300
         Left            =   200
         TabIndex        =   8
         Top             =   4000
         Width           =   1600
      End
      Begin VB.Label Label6 
         Caption         =   "N -- 新游戏"
         Height          =   300
         Left            =   200
         TabIndex        =   7
         Top             =   3400
         Width           =   1600
      End
      Begin VB.Label Label5 
         Caption         =   "空格 -- 暂停/继续"
         Height          =   300
         Left            =   200
         TabIndex        =   6
         Top             =   2800
         Width           =   1600
      End
      Begin VB.Label Label4 
         Caption         =   "→ -- 右移"
         Height          =   300
         Left            =   200
         TabIndex        =   5
         Top             =   2200
         Width           =   1400
      End
      Begin VB.Label Label3 
         Caption         =   "← -- 左移"
         Height          =   300
         Left            =   200
         TabIndex        =   4
         Top             =   1600
         Width           =   1400
      End
      Begin VB.Label Label2 
         Caption         =   "↓ -- 加速"
         Height          =   300
         Left            =   200
         TabIndex        =   3
         Top             =   1000
         Width           =   1400
      End
      Begin VB.Label Label1 
         Caption         =   "↑ -- 变形"
         Height          =   300
         Left            =   200
         TabIndex        =   2
         Top             =   400
         Width           =   1400
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   480
      Top             =   5640
   End
   Begin VB.Label overLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "游戏结束"
      BeginProperty Font 
         Name            =   "华文新魏"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   2500
      TabIndex        =   9
      Top             =   2600
      Width           =   3030
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   6060
      Left            =   2500
      Top             =   -10
      Width           =   3030
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   1800
      Left            =   300
      Top             =   0
      Width           =   1800
   End
   Begin VB.Label scoreLabel 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   3240
      Width           =   1815
   End
End
Attribute VB_Name = "formMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const KEY_LEFT = 37         '左箭头，向左移动
Const KEY_UP = 38           '上箭头，改变形状
Const KEY_RIGHT = 39        '右箭头，向右移动
Const KEY_DOWN = 40         '下箭头，加速
Const PAUSE = 32            '空格，暂停、继续
Const NEW_GAME = 78         '字母N，开始新的游戏
Const END_GAME = 27         'ESC，退出游戏

Public resize As Boolean

Dim preNum As Integer
Dim sameCount As Integer


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

   If Timer1.Enabled = True Then
   
   With cBox
        If KeyCode = KEY_RIGHT Then
           If hitright(cBox) = True Then Exit Sub
           If hitbrixright(cBox) = True Then Exit Sub
           clear cBox
           For i = 1 To 4
               .X(i) = .X(i) + boxwidth
           Next i
           display cBox
       End If
       If KeyCode = KEY_LEFT Then
          If hitleft(cBox) = True Then Exit Sub
          If hitbrixleft(cBox) = True Then Exit Sub
          clear cBox
          For i = 1 To 4
              .X(i) = .X(i) - boxwidth
          Next i
          display cBox
       End If
       If KeyCode = KEY_UP Then
          clear cBox
          tempstate = state
          state = state + 1
          If state > .rot Then
             state = 1
          End If
          If rotate(cBox, state) = False Then state = tempstate
          display cBox
       End If
       If KeyCode = KEY_DOWN Then
          For i = 1 To 7
              clear cBox
              If Control = True Then Exit Sub
              For j = 1 To 4
                  .Y(j) = .Y(j) + boxwidth
              Next j
              display cBox
          Next i
          If Control = True Then Exit Sub
       End If
    End With
    End If
       
       If KeyCode = PAUSE Then
          Timer1.Enabled = Not (Timer1.Enabled)
       End If
       If KeyCode = NEW_GAME Then
          Form_Load
       End If
       
       If KeyCode = END_GAME Then
          Form_Unload (1)
       End If
  
End Sub


Private Sub Form_Load()

    formMain.Refresh
    overLabel.Visible = False
    Timer1.Enabled = True
    
    boxwidth = 300: pos = 0
    fx = 2500
    fy = 0
    score = 0
    scoreLabel.Caption = "得分 : 0"
    
    firstBrix
    display cBox
    
    nextBox

End Sub

Private Sub Form_Unload(Cancel As Integer)
   End
End Sub


Private Sub Timer1_Timer()
    If resize = True Then
      For i = 1 To pos
          test (i)
      Next i
      resize = False
   End If
   Dim showNext As Boolean
   showNext = True
   With cBox
        If Control = True Then Exit Sub
        clear cBox
        For i = 1 To 4
            .Y(i) = .Y(i) + boxwidth
            If .Y(i) < boxwidth Then
                showNext = False
            End If
        Next i
        display cBox
   End With
   
    If showNext Then
        clearNextArea
        drawBox nBox
    End If
End Sub


Private Sub display(box As tempbox)
   With box
        For i = 1 To 4
            Line (.X(i), .Y(i))-(.X(i) + boxwidth, .Y(i) + boxwidth), RGB(.r, .g, .b), BF
            fill box, 0, 0, 0, i
        Next i
   End With
End Sub


Private Sub clear(box As tempbox)
   With box
        For i = 1 To 4
            Line (.X(i), .Y(i))-(.X(i) + boxwidth, .Y(i) + boxwidth), RGB(0, 0, 0), BF
            fill box, 0, 0, 0, i
        Next i
   End With
End Sub


Private Sub clearNextArea()
    Line (300, 0)-(2100, 1800), RGB(0, 0, 0), BF
End Sub


Private Sub fill(box As tempbox, rd, gr, bl, j)
   With box
        Line (.X(j), .Y(j))-(.X(j) + boxwidth, .Y(j)), RGB(rd, gr, bl)
        Line (.X(j), .Y(j))-(.X(j), .Y(j) + boxwidth), RGB(rd, gr, bl)
        Line (.X(j), .Y(j) + boxwidth)-(.X(j) + boxwidth, .Y(j) + boxwidth), RGB(rd, gr, bl)
        Line (.X(j) + boxwidth, .Y(j))-(.X(j) + boxwidth, .Y(j) + boxwidth), RGB(rd, gr, bl)
   End With
End Sub


Private Sub border(box As tempbox)
   With box
        For j = 1 To 4
            Line (.X(j), .Y(j))-(.X(j) + boxwidth, .Y(j)), RGB(0, 0, 0)
            Line (.X(j), .Y(j))-(.X(j), .Y(j) + boxwidth), RGB(0, 0, 0)
            Line (.X(j), .Y(j) + boxwidth)-(.X(j) + boxwidth, .Y(j) + boxwidth), RGB(0, 0, 0)
            Line (.X(j) + boxwidth, .Y(j))-(.X(j) + boxwidth, .Y(j) + boxwidth), RGB(0, 0, 0)
        Next j
   End With
End Sub


Private Sub drawBox(box As tempbox)
    With box
        For i = 1 To 4
            Line (.X(i), .Y(i))-(.X(i) + boxwidth, .Y(i) + boxwidth), RGB(.r, .g, .b), BF
        Next i
   End With
   border box
End Sub



Public Function Control() As Boolean
    With cBox
        For i = 1 To 4
            If .Y(i) + boxwidth = fy + (20 * boxwidth) Then
               For k = 1 To 4
                   pos = pos + 1
                   old(pos).X = .X(k)
                   old(pos).Y = .Y(k)
                   old(pos).bl = True
                   old(pos).r = .r
                   old(pos).g = .g
                   old(pos).b = .b
               Next k
               display cBox
               control2
               newbrix
               Control = True
               Exit Function
            End If
        Next i
        '****************************
        For i = 1 To 4
            For l = 1 To pos
                If .Y(i) + boxwidth = old(l).Y And .X(i) = old(l).X Then
                   For k = 1 To 4
                       If .Y(k) = fy Then
                          gameover
                          Control = True
                          Exit Function
                       End If
                       pos = pos + 1
                       old(pos).X = .X(k)
                       old(pos).Y = .Y(k)
                       old(pos).bl = True
                       old(pos).r = .r
                       old(pos).g = .g
                       old(pos).b = .b
                   Next k
                   display cBox
                   control2
                   newbrix
                   Control = True
                   Exit Function
                End If
            Next l
        Next i
        Control = False
    End With
End Function

Public Sub control2()
   Dim fin(20)
   metr = 0
   For i = 1 To 20
       cn = 0
       For j = 1 To pos
           If old(j).Y = fy + (i * boxwidth) Then
              cn = cn + 1
              If cn = 10 Then
                 metr = metr + 1
                 fin(metr) = old(j).Y
              End If
           End If
       Next j
   Next i
   '**********************
   If metr <> 0 Then
      score = score + (metr * 10)
      scoreLabel.Caption = "得分 : " + Str(score)
      For i = 1 To metr
          For j = 1 To pos
              If old(j).Y = fin(i) Then
                 old(j).bl = False
              End If
          Next j
      Next i
   '***********************
      Line (fx, fy)-(fx + (10 * boxwidth), fy + (20 * boxwidth)), RGB(0, 0, 0), BF
      For j = 1 To metr
          For i = 1 To pos
              If old(i).bl = True And old(i).Y < fin(j) Then
                 old(i).Y = old(i).Y + boxwidth
              End If
          Next i
      Next j
   '************************
      num = 0
      For i = 1 To pos
          If old(i).bl = True Then
             num = num + 1
             clean(num).X = old(i).X: clean(num).Y = old(i).Y
             clean(num).r = old(i).r: clean(num).g = old(i).g: clean(num).b = old(i).b
             test (i)
          End If
      Next i
   '***************************
      pos = num
      For i = 1 To num
          With old(i)
               .X = clean(i).X
               .Y = clean(i).Y
               .bl = True
               .r = clean(i).r
               .g = clean(i).g
               .b = clean(i).b
          End With
      Next i
   End If
End Sub

Private Function nextNum() As Integer
    While True
        Randomize
        rndSeek = Int(Rnd * 100)
        Randomize rndSeek
        nnum = Int(Rnd * 7) + 1
        If nnum <> preNum Then
            nextNum = nnum
            sameCount = 0
            preNum = nnum
            Exit Function
        End If
        sameCount = sameCount + 1
        If sameCount < 3 Then
            nextNum = nnum
            Exit Function
        End If
    Wend
End Function


Private Sub nextBox()
    'Randomize
    Dim nnum As Integer
    'nnum = Int(Rnd * 7) + 1
    nnum = nextNum
    
    nBox.X(1) = 900
    nBox.Y(1) = 300
    nBox.r = Int(Rnd * 255) + 100
    nBox.g = Int(Rnd * 255) + 100
    nBox.b = Int(Rnd * 255) + 100
    nBox.num = nnum
    If nnum = 2 Then
        nBox.rot = 1
    ElseIf nnum = 7 Then
        nBox.rot = 2
    Else
        nBox.rot = 4
    End If
    calcBox nBox, 1
    'clearNextArea
    'drawBox nBox
End Sub

Private Sub firstBrix()
    'Randomize
    Dim nnum As Integer
    'nnum = Int(Rnd * 7) + 1
    
    nnum = nextNum
    
    With cBox
        .num = nnum
        .r = Int(Rnd * 255) + 100
        .g = Int(Rnd * 255) + 100
        .b = Int(Rnd * 255) + 100
        .X(1) = fx + (4 * boxwidth)
        If nnum = 2 Then
           .Y(1) = fy - (2 * boxwidth)
           .rot = 1
        ElseIf nnum = 7 Then
            .Y(1) = fy - (4 * boxwidth)
            .rot = 2
        Else
            .Y(1) = fy - (3 * boxwidth)
            .rot = 4
        End If
        state = 1
    End With
    calcBox cBox, 1
End Sub


Public Sub newbrix()
    
    cBox = nBox
    With cBox
        .X(1) = fx + (4 * boxwidth)
        If .num = 2 Then
           .Y(1) = fy - (2 * boxwidth)
        Else
           If .num = 7 Then
              .Y(1) = fy - (4 * boxwidth)
           Else
              .Y(1) = fy - (3 * boxwidth)
           End If
        End If
        state = 1
    End With
    calcBox cBox, 1
    nextBox
End Sub

Public Sub test(i)
   With old(i)
        Line (.X, .Y)-(.X + boxwidth, .Y + boxwidth), RGB(.r, .g, .b), BF
        Line (.X, .Y)-(.X + boxwidth, .Y), RGB(0, 0, 0)
        Line (.X, .Y)-(.X, .Y + boxwidth), RGB(0, 0, 0)
        Line (.X, .Y + boxwidth)-(.X + boxwidth, .Y + boxwidth), RGB(0, 0, 0)
        Line (.X + boxwidth, .Y)-(.X + boxwidth, .Y + boxwidth), RGB(0, 0, 0)
 End With
End Sub

Public Sub gameover()
    Timer1.Enabled = False
    overLabel.Visible = True
End Sub


