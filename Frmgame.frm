VERSION 5.00
Begin VB.Form FrmGame 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "五子棋"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   5205
   DrawStyle       =   1  'Dash
   FillColor       =   &H00004000&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "System"
      Size            =   12
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   Icon            =   "Frmgame.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   5205
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox PicMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   4845
      Left            =   150
      MousePointer    =   99  'Custom
      ScaleHeight     =   4815
      ScaleWidth      =   4815
      TabIndex        =   0
      Top             =   480
      Width           =   4845
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "当前下棋方："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Index           =   1
      Left            =   150
      TabIndex        =   5
      Top             =   180
      Width           =   1530
   End
   Begin VB.Image ImgNow 
      Height          =   225
      Left            =   1590
      Stretch         =   -1  'True
      Top             =   195
      Width           =   210
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "黑方："
      Height          =   240
      Index           =   0
      Left            =   840
      TabIndex        =   4
      Top             =   5370
      Width           =   765
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "白方："
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   2940
      TabIndex        =   3
      Top             =   5370
      Width           =   765
   End
   Begin VB.Label lblPrompt1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "胜0盘，负0盘"
      ForeColor       =   &H00C0E0FF&
      Height          =   240
      Left            =   540
      TabIndex        =   2
      Top             =   5700
      Width           =   1515
   End
   Begin VB.Label lblPrompt2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "胜0盘，负0盘"
      ForeColor       =   &H00C0E0FF&
      Height          =   240
      Left            =   2850
      TabIndex        =   1
      Top             =   5700
      Width           =   1515
   End
   Begin VB.Image Img 
      Height          =   225
      Index           =   0
      Left            =   540
      Picture         =   "Frmgame.frx":030A
      Stretch         =   -1  'True
      Top             =   5370
      Width           =   225
   End
   Begin VB.Image Img 
      Height          =   225
      Index           =   1
      Left            =   2640
      Picture         =   "Frmgame.frx":0FD4
      Stretch         =   -1  'True
      Top             =   5370
      Width           =   225
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu MenuStart 
         Caption         =   "重新开始(&N)"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu menuInfo 
      Caption         =   "信息(&I)"
      Begin VB.Menu menuAbout 
         Caption         =   "关于(&A)"
      End
   End
End
Attribute VB_Name = "FrmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MapBlack(1 To 19, 1 To 19, 0 To 4) As Single
Dim MapWhite(1 To 19, 1 To 19, 0 To 4) As Single
Dim NowOpt As Boolean
Dim CanOpt As Boolean
Dim BlackWin As Integer
Dim BlackLost As Integer
Dim WhiteWin As Integer
Dim WhiteLost As Integer
Const Grid = 240

Private Sub InitializePre()  '初始化优先级
    Dim i As Integer, j As Integer
    For i = 1 To 9
        For j = i To 19 - i
            MapBlack(i, j, 1) = i
            MapBlack(i, j, 2) = i
            MapBlack(i, j, 3) = i
            MapBlack(i, j, 4) = i
            MapBlack(j, i, 1) = i
            MapBlack(j, i, 2) = i
            MapBlack(j, i, 3) = i
            MapBlack(j, i, 4) = i
        Next j
    Next i
    For i = 19 To 11 Step -1
        For j = 20 - i To i
            MapBlack(i, j, 1) = 20 - i
            MapBlack(i, j, 2) = 20 - i
            MapBlack(i, j, 3) = 20 - i
            MapBlack(i, j, 4) = 20 - i
            MapBlack(j, i, 1) = 20 - i
            MapBlack(j, i, 2) = 20 - i
            MapBlack(j, i, 3) = 20 - i
            MapBlack(j, i, 4) = 20 - i
        Next j
    Next i
    MapBlack(10, 10, 1) = 10
    MapBlack(10, 10, 2) = 10
    MapBlack(10, 10, 3) = 10
    MapBlack(10, 10, 4) = 10
    For i = 1 To 19
        For j = 1 To 19
            MapWhite(i, j, 1) = MapBlack(i, j, 1)
            MapWhite(i, j, 2) = MapBlack(i, j, 2)
            MapWhite(i, j, 3) = MapBlack(i, j, 3)
            MapWhite(i, j, 4) = MapBlack(i, j, 4)
        Next j
    Next i
End Sub

Private Function InputeCalcPre() '输入计算优先级别
    Dim i As Integer, j As Integer, Sum As Single
    Dim ii As Integer, jj As Integer
    For i = 1 To 19
        For j = 1 To 19
            If MapWhite(i, j, 0) = 0 And MapBlack(i, j, 0) = 0 Then
               '\
               ii = i - 1
               jj = j - 1
               Sum = 0
               Do While ii > 0 And jj > 0
                  If MapWhite(ii, jj, 0) = 1 Then
                     Sum = Sum + 1
                  Else
                     If MapBlack(ii, jj, 0) = 1 Then
                        If Sum < 4 Then Sum = Sum - 1
                     End If
                     Exit Do
                  End If
                    ii = ii - 1
                    jj = jj - 1
               Loop
               ii = i + 1
               jj = j + 1
               Do While ii < 20 And jj < 20
                  If MapWhite(ii, jj, 0) = 1 Then
                     Sum = Sum + 1
                  Else
                     If MapBlack(ii, jj, 0) = 1 Then
                        If Sum < 4 Then Sum = Sum - 1
                     End If
                     Exit Do
                  End If
                  ii = ii + 1
                  jj = jj + 1
               Loop
               If Sum > 0 Then
                  MapWhite(i, j, 1) = Sum * 1000
               End If
               ii = i - 1
               jj = j - 1
               Sum = 0
              Do While ii > 0 And jj > 0
                  If MapBlack(ii, jj, 0) = 1 Then
                     Sum = Sum + 1
                  Else
                     If MapWhite(ii, jj, 0) = 1 Then
                        If Sum < 4 Then Sum = Sum - 1
                     End If
                     Exit Do
                  End If
                    ii = ii - 1
                    jj = jj - 1
               Loop
               ii = i + 1
               jj = j + 1
               Do While ii < 20 And jj < 20
                  If MapBlack(ii, jj, 0) = 1 Then
                     Sum = Sum + 1
                  Else
                     If MapWhite(ii, jj, 0) = 1 Then
                        If Sum < 4 Then Sum = Sum - 1
                     End If
                     Exit Do
                  End If
                  ii = ii + 1
                  jj = jj + 1
               Loop
               If Sum > 0 Then
                  MapBlack(i, j, 1) = Sum * 1000
               End If
               '/
               ii = i - 1
               jj = j + 1
               Sum = 0
               Do While ii > 0 And jj < 20
                  If MapWhite(ii, jj, 0) = 1 Then
                     Sum = Sum + 1
                  Else
                     If MapBlack(ii, jj, 0) = 1 Then
                        If Sum < 4 Then Sum = Sum - 1
                     End If
                     Exit Do
                  End If
                    ii = ii - 1
                    jj = jj + 1
               Loop
               ii = i + 1
               jj = j - 1
               Do While ii < 20 And jj > 0
                  If MapWhite(ii, jj, 0) = 1 Then
                     Sum = Sum + 1
                  Else
                     If MapBlack(ii, jj, 0) = 1 Then
                        If Sum < 4 Then Sum = Sum - 1
                     End If
                     Exit Do
                  End If
                  ii = ii + 1
                  jj = jj - 1
               Loop
               If Sum > 0 Then
                  MapWhite(i, j, 4) = Sum * 1000
               End If
               ii = i - 1
               jj = j + 1
               Sum = 0
              Do While ii > 0 And jj < 20
                  If MapBlack(ii, jj, 0) = 1 Then
                     Sum = Sum + 1
                  Else
                     If MapWhite(ii, jj, 0) = 1 Then
                        If Sum < 4 Then Sum = Sum - 1
                     End If
                     Exit Do
                  End If
                    ii = ii - 1
                    jj = jj + 1
               Loop
               ii = i + 1
               jj = j - 1
               Do While ii < 20 And jj > 0
                  If MapBlack(ii, jj, 0) = 1 Then
                     Sum = Sum + 1
                  Else
                     If MapWhite(ii, jj, 0) = 1 Then
                        If Sum < 4 Then Sum = Sum - 1
                     End If
                     Exit Do
                  End If
                  ii = ii + 1
                  jj = jj - 1
               Loop
               If Sum > 0 Then
                  MapBlack(i, j, 4) = Sum * 1000
               End If
               
               '-
               ii = i
               jj = j - 1
               Sum = 0
               Do While jj > 0
                  If MapWhite(ii, jj, 0) = 1 Then
                     Sum = Sum + 1
                  Else
                     If MapBlack(ii, jj, 0) = 1 Then
                        If Sum < 4 Then Sum = Sum - 1
                     End If
                     Exit Do
                  End If
                    jj = jj - 1
               Loop
               jj = j + 1
               Do While jj < 20
                  If MapWhite(ii, jj, 0) = 1 Then
                     Sum = Sum + 1
                  Else
                     If MapBlack(ii, jj, 0) = 1 Then
                        If Sum < 4 Then Sum = Sum - 1
                     End If
                     Exit Do
                  End If
                  jj = jj + 1
               Loop
               If Sum > 0 Then
                  MapWhite(i, j, 3) = Sum * 1000
               End If
               jj = j - 1
               Sum = 0
              Do While jj > 0
                  If MapBlack(ii, jj, 0) = 1 Then
                     Sum = Sum + 1
                  Else
                     If MapWhite(ii, jj, 0) = 1 Then
                        If Sum < 4 Then Sum = Sum - 1
                     End If
                     Exit Do
                  End If
                    jj = jj - 1
               Loop
               jj = j + 1
               Do While jj < 20
                  If MapBlack(ii, jj, 0) = 1 Then
                     Sum = Sum + 1
                  Else
                     If MapWhite(ii, jj, 0) = 1 Then
                        If Sum < 4 Then Sum = Sum - 1
                     End If
                     Exit Do
                  End If
                  jj = jj + 1
               Loop
               If Sum > 0 Then
                  MapBlack(i, j, 3) = Sum * 1000
               End If
               
                '|
               ii = i - 1
               jj = j
               Sum = 0
               Do While ii > 0
                  If MapWhite(ii, jj, 0) = 1 Then
                     Sum = Sum + 1
                  Else
                     If MapBlack(ii, jj, 0) = 1 Then
                        If Sum < 4 Then Sum = Sum - 1
                     End If
                     Exit Do
                  End If
                    ii = ii - 1
               Loop
               ii = i + 1
               Do While ii < 20
                  If MapWhite(ii, jj, 0) = 1 Then
                     Sum = Sum + 1
                  Else
                     If MapBlack(ii, jj, 0) = 1 Then
                        If Sum < 4 Then Sum = Sum - 1
                     End If
                     Exit Do
                  End If
                  ii = ii + 1
               Loop
               If Sum > 0 Then
                  MapWhite(i, j, 2) = Sum * 1000
               End If
               ii = i - 1
               Sum = 0
              Do While ii > 0
                  If MapBlack(ii, jj, 0) = 1 Then
                     Sum = Sum + 1
                  Else
                     If MapWhite(ii, jj, 0) = 1 Then
                        If Sum < 4 Then Sum = Sum - 1
                     End If
                     Exit Do
                  End If
                    ii = ii - 1
               Loop
               ii = i + 1
               Do While ii < 20
                  If MapBlack(ii, jj, 0) = 1 Then
                     Sum = Sum + 1
                  Else
                     If MapWhite(ii, jj, 0) = 1 Then
                        If Sum < 4 Then Sum = Sum - 1
                     End If
                     Exit Do
                  End If
                  ii = ii + 1
               Loop
               If Sum > 0 Then
                  MapBlack(i, j, 2) = Sum * 1000
               End If
            Else
               MapWhite(i, j, 1) = 0
               MapWhite(i, j, 2) = 0
               MapWhite(i, j, 3) = 0
               MapWhite(i, j, 4) = 0
               MapBlack(i, j, 1) = 0
               MapBlack(i, j, 2) = 0
               MapBlack(i, j, 3) = 0
               MapBlack(i, j, 4) = 0
            End If
        Next j
    Next i
End Function

'判断胜利条件
Private Function OpinionWin(Opt As Boolean) As String
    Dim i As Integer, j As Integer, k As Integer, Sum As Integer
    Dim ii As Integer, jj As Integer
    If Opt = False Then
          For i = 1 To 19
              For j = 1 To 19
                  If MapBlack(i, j, 0) = 1 Then
                     Sum = Sum + 1
                  Else
                     Sum = 0
                  End If
                  If Sum = 5 Then
                     For k = j To j - 4 Step -1
                         Call ShowMap(Opt, k * Grid, i * Grid, True)
                     Next k
                     OpinionWin = "黑方胜"
                  End If
              Next j
              Sum = 0
          Next i
          For i = 1 To 19
              For j = 1 To 19
                  If MapBlack(j, i, 0) = 1 Then
                     Sum = Sum + 1
                  Else
                     Sum = 0
                  End If
                  If Sum = 5 Then
                     For k = j To j - 4 Step -1
                         Call ShowMap(Opt, i * Grid, k * Grid, True)
                     Next k
                     OpinionWin = "黑方胜"
                  End If
              Next j
              Sum = 0
          Next i
          For i = 1 To 19
              For j = 19 To 1 Step -1
                  ii = i
                  jj = j
                  Do While ii <= 19 And jj <= 19
                     If MapBlack(ii, jj, 0) = 1 Then
                        Sum = Sum + 1
                     Else
                        Sum = 0
                     End If
                    If Sum = 5 Then
                       For k = 0 To 4
                           Call ShowMap(Opt, (jj - k) * Grid, (ii - k) * Grid, True)
                       Next k
                       OpinionWin = "黑方胜"
                    End If
                       jj = jj + 1
                       ii = ii + 1
                  Loop
                  Sum = 0
              Next j
          Next i
          For i = 1 To 19
              For j = 1 To 19
                  ii = i
                  jj = j
                  Do While ii <= 19 And jj >= 1
                     If MapBlack(ii, jj, 0) = 1 Then
                        Sum = Sum + 1
                     Else
                        Sum = 0
                     End If
                    If Sum = 5 Then
                       For k = 0 To 4
                           Call ShowMap(Opt, (jj + k) * Grid, (ii - k) * Grid, True)
                       Next k
                       OpinionWin = "黑方胜"
                    End If
                       jj = jj - 1
                       ii = ii + 1
                  Loop
                  Sum = 0
              Next j
          Next i
    Else
          For i = 1 To 19
              For j = 1 To 19
                  If MapWhite(i, j, 0) = 1 Then
                     Sum = Sum + 1
                  Else
                     Sum = 0
                  End If
                  If Sum = 5 Then
                     For k = j To j - 4 Step -1
                         Call ShowMap(Opt, k * Grid, i * Grid, True)
                     Next k
                     OpinionWin = "白方胜"
                  End If
              Next j
              Sum = 0
          Next i
          For i = 1 To 19
              For j = 1 To 19
                  If MapWhite(j, i, 0) = 1 Then
                     Sum = Sum + 1
                  Else
                     Sum = 0
                  End If
                  If Sum = 5 Then
                     For k = j To j - 4 Step -1
                         Call ShowMap(Opt, i * Grid, k * Grid, True)
                     Next k
                     OpinionWin = "白方胜"
                  End If
              Next j
              Sum = 0
          Next i
          For i = 1 To 19
              For j = 19 To 1 Step -1
                  ii = i
                  jj = j
                  Do While ii <= 19 And jj <= 19
                     If MapWhite(ii, jj, 0) = 1 Then
                        Sum = Sum + 1
                     Else
                        Sum = 0
                     End If
                    If Sum = 5 Then
                       For k = 0 To 4
                           Call ShowMap(Opt, (jj - k) * Grid, (ii - k) * Grid, True)
                       Next k
                       OpinionWin = "白方胜"
                    End If
                       jj = jj + 1
                       ii = ii + 1
                  Loop
                  Sum = 0
              Next j
          Next i
          For i = 1 To 19
              For j = 1 To 19
                  ii = i
                  jj = j
                  Do While ii <= 19 And jj >= 1
                     If MapWhite(ii, jj, 0) = 1 Then
                        Sum = Sum + 1
                     Else
                        Sum = 0
                     End If
                    If Sum = 5 Then
                       For k = 0 To 4
                           Call ShowMap(Opt, (jj + k) * Grid, (ii - k) * Grid, True)
                       Next k
                       OpinionWin = "白方胜"
                    End If
                       jj = jj - 1
                       ii = ii + 1
                  Loop
                  Sum = 0
              Next j
          Next i
    End If
End Function

Private Function ReadMapData()
    Dim i As Integer, j As Integer
    For i = 1 To 19
        For j = 1 To 19
            If MapBlack(i, j, 0) = 1 Then
               Call ShowMap(False, j * Grid, i * Grid)
            End If
            If MapWhite(i, j, 0) = 1 Then
               Call ShowMap(True, j * Grid, i * Grid)
            End If
        Next j
    Next i
End Function

Private Sub ShowMap(Opt As Boolean, X As Integer, Y As Integer, Optional CN As Boolean = False)
    If Opt = False Then
        If CN = True Then
            PicMap.ForeColor = vbBlack
            PicMap.DrawWidth = 15
            PicMap.PSet (X, Y)
            PicMap.ForeColor = vbMagenta
            PicMap.DrawWidth = 13
            PicMap.PSet (X, Y)
            PicMap.ForeColor = vbBlack
            PicMap.DrawWidth = 11
            PicMap.PSet (X, Y)
        Else
            PicMap.ForeColor = vbBlack
            PicMap.DrawWidth = 15
            PicMap.PSet (X, Y)
        End If
    Else
        If CN = True Then
            PicMap.ForeColor = vbBlack
            PicMap.DrawWidth = 15
            PicMap.PSet (X, Y)
            PicMap.ForeColor = vbCyan
            PicMap.DrawWidth = 13
            PicMap.PSet (X, Y)
            PicMap.ForeColor = vbWhite
            PicMap.DrawWidth = 11
            PicMap.PSet (X, Y)
        Else
            PicMap.ForeColor = vbBlack
            PicMap.DrawWidth = 15
            PicMap.PSet (X, Y)
            PicMap.ForeColor = vbWhite
            PicMap.DrawWidth = 13
            PicMap.PSet (X, Y)
        End If
    End If
End Sub

Private Sub ReDrawMap()   '重新绘制地图
    PicMap.Cls
    Dim i As Integer, j As Integer
    PicMap.ForeColor = &H404040
    PicMap.DrawWidth = 1
    For i = 1 To 19
        For j = 1 To 19
            If j <> 19 Then
                PicMap.Line (i * Grid, j * Grid)-(i * Grid, j * Grid + Grid)
            End If
            If i <> 19 Then
                PicMap.Line (i * Grid, j * Grid)-(i * Grid + Grid, j * Grid)
            End If
            If (i = 4 And j = 4) Or (i = 16 And j = 16) Or (i = 16 And j = 4) Or (i = 4 And j = 16) _
                Or (i = 10 And j = 10) Or (i = 4 And j = 10) Or (i = 10 And j = 4) Or (i = 16 And j = 10) _
                Or (i = 10 And j = 16) Then
                PicMap.ForeColor = vbBlack
                PicMap.DrawWidth = 8
                PicMap.PSet (i * Grid, j * Grid)
                PicMap.DrawWidth = 1
                PicMap.ForeColor = &H404040
            End If
        Next j
    Next i
          PicMap.ForeColor = vbBlack
          PicMap.DrawWidth = 2
          PicMap.Line (Grid, Grid)-(Grid, 19 * Grid)
          PicMap.Line (19 * Grid, Grid)-(19 * Grid, 19 * Grid)
          PicMap.Line (Grid, Grid)-(19 * Grid, Grid)
          PicMap.Line (Grid, 19 * Grid)-(19 * Grid, 19 * Grid)
          PicMap.DrawWidth = 1
End Sub

Private Sub ReInputData()
    NowOpt = False
    Dim i As Integer, j As Integer, k As Integer
    For k = 0 To 4
        For j = 1 To 19
            For i = 1 To 19
                MapBlack(i, j, k) = 0
                MapWhite(i, j, k) = 0
            Next i
        Next j
    Next k
End Sub

Private Sub Form_Load()
    Call MenuStart_Click
End Sub

Private Sub menuAbout_Click()
    MsgBox "算法编写比较简单，子力较弱，还望高手指点!" & vbCrLf & _
           "  QQ: 115064582, Email: pariszh@163.com", vbOKOnly + vbInformation, "关于五子棋游戏"
End Sub

Private Sub MenuStart_Click()
    Call ReDrawMap
    Call ReInputData
    Call InitializePre
    ImgNow.Picture = Img(NowOpt).Picture
    CanOpt = True
End Sub


Private Sub PicMap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CanOpt = True Then
    Dim tC As Integer, tR As Integer, tmp As String
    Dim SumB As Integer, SumW As Integer, XXB As Integer, YYB As Integer, XXW As Integer, YYW As Integer
    Dim i As Integer, j As Integer
    tC = X \ Grid
    tR = Y \ Grid
    If tC < X / Grid And X > (tC + 1) * Grid - Grid / 2 Then
       tC = tC + 1
    End If
    If tR < Y / Grid And Y > (tR + 1) * Grid - Grid / 2 Then
       tR = tR + 1
    End If
    If tC > 0 And tR > 0 Then
        If MapBlack(tR, tC, 0) = 0 And MapWhite(tR, tC, 0) = 0 Then  '判断该位置是否能下棋
'            If NowOpt = False Then   '黑方下棋
               MapBlack(tR, tC, 0) = 1
               Call InputeCalcPre
'            End If
            NowOpt = False
            Call ReDrawMap
            Call ReadMapData
            Call ShowMap(NowOpt, tC * Grid, tR * Grid, True)
            tmp = OpinionWin(NowOpt)
            If tmp <> "" Then
              MsgBox tmp, vbInformation + vbOKOnly, "恭喜"
              CanOpt = False
              If NowOpt = False Then
                 BlackWin = BlackWin + 1
                 WhiteLost = WhiteLost + 1
              Else
                 WhiteWin = WhiteWin + 1
                 BlackLost = BlackLost + 1
              End If
                 lblPrompt1.Caption = "胜" & BlackWin & "盘，负" & BlackLost & "盘"
                 lblPrompt2.Caption = "胜" & WhiteWin & "盘，负" & WhiteLost & "盘"
                 Exit Sub
            End If
            XXB = 1: XXW = 1: YYB = 1: YYW = 1
            For i = 1 To 19
                For j = 1 To 19
                    If SumB < MapBlack(i, j, 1) + MapBlack(i, j, 2) Then
                       SumB = MapBlack(i, j, 1) + MapBlack(i, j, 2)
                       XXB = j
                       YYB = i
                    End If
                    If SumB < MapBlack(i, j, 1) + MapBlack(i, j, 3) Then
                       SumB = MapBlack(i, j, 1) + MapBlack(i, j, 3)
                       XXB = j
                       YYB = i
                    End If
                    If SumB < MapBlack(i, j, 1) + MapBlack(i, j, 4) Then
                       SumB = MapBlack(i, j, 1) + MapBlack(i, j, 4)
                       XXB = j
                       YYB = i
                    End If
                    If SumB < MapBlack(i, j, 2) + MapBlack(i, j, 3) Then
                       SumB = MapBlack(i, j, 2) + MapBlack(i, j, 3)
                       XXB = j
                       YYB = i
                    End If
                    If SumB < MapBlack(i, j, 2) + MapBlack(i, j, 4) Then
                       SumB = MapBlack(i, j, 2) + MapBlack(i, j, 4)
                       XXB = j
                       YYB = i
                    End If
                    If SumB < MapBlack(i, j, 3) + MapBlack(i, j, 4) Then
                       SumB = MapBlack(i, j, 3) + MapBlack(i, j, 4)
                       XXB = j
                       YYB = i
                    End If
                    If SumW < MapWhite(i, j, 1) + MapWhite(i, j, 2) Then
                       SumW = MapWhite(i, j, 1) + MapWhite(i, j, 2)
                       XXW = j
                       YYW = i
                    End If
                    If SumW < MapWhite(i, j, 1) + MapWhite(i, j, 3) Then
                       SumW = MapWhite(i, j, 1) + MapWhite(i, j, 3)
                       XXW = j
                       YYW = i
                    End If
                    If SumW < MapWhite(i, j, 1) + MapWhite(i, j, 4) Then
                       SumW = MapWhite(i, j, 1) + MapWhite(i, j, 4)
                       XXW = j
                       YYW = i
                    End If
                    If SumW < MapWhite(i, j, 2) + MapWhite(i, j, 3) Then
                       SumW = MapWhite(i, j, 2) + MapWhite(i, j, 3)
                       XXW = j
                       YYW = i
                    End If
                    If SumW < MapWhite(i, j, 2) + MapWhite(i, j, 4) Then
                       SumW = MapWhite(i, j, 2) + MapWhite(i, j, 4)
                       XXW = j
                       YYW = i
                    End If
                    If SumW < MapWhite(i, j, 3) + MapWhite(i, j, 4) Then
                       SumW = MapWhite(i, j, 3) + MapWhite(i, j, 4)
                       XXW = j
                       YYW = i
                    End If
                Next j
            Next i
                    If SumB > SumW Then
                       MapWhite(YYB, XXB, 0) = 1
                    Else
                       MapWhite(YYW, XXW, 0) = 1
                    End If
                            Call ReDrawMap
                            Call ReadMapData
                    NowOpt = True
                    If SumB > SumW Then
                            Call ShowMap(NowOpt, XXB * Grid, YYB * Grid, True)
                    Else
                            Call ShowMap(NowOpt, XXW * Grid, YYW * Grid, True)
                    End If
            tmp = OpinionWin(NowOpt)
            If tmp <> "" Then
              MsgBox tmp, vbInformation + vbOKOnly, "加油"
              CanOpt = False
              If NowOpt = False Then
                 BlackWin = BlackWin + 1
                 WhiteLost = WhiteLost + 1
              Else
                 WhiteWin = WhiteWin + 1
                 BlackLost = BlackLost + 1
              End If
                lblPrompt1.Caption = "胜" & BlackWin & "盘，负" & BlackLost & "盘"
                lblPrompt2.Caption = "胜" & WhiteWin & "盘，负" & WhiteLost & "盘"
                Exit Sub
            End If

        End If
    End If
End If
End Sub
'源码整理：http://www.codefans.net