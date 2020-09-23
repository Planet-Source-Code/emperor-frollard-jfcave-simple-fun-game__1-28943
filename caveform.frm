VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Frostys Cave Game"
   ClientHeight    =   9000
   ClientLeft      =   -195
   ClientTop       =   1155
   ClientWidth     =   12000
   Icon            =   "caveform.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   0  'User
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrScoreUpdate 
      Interval        =   250
      Left            =   5400
      Top             =   3720
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   10680
      TabIndex        =   0
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label lblGameOver 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   144
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   8175
      Left            =   1365
      TabIndex        =   2
      Top             =   210
      Visible         =   0   'False
      Width           =   8895
   End
   Begin VB.Label lblFinalScore 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   1
      Left            =   6615
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label lblFinalScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Final Score:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   0
      Left            =   5250
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   9240
      TabIndex        =   1
      Top             =   0
      Width           =   2775
   End
   Begin VB.Line linePlayer 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      DrawMode        =   15  'Merge Pen Not
      Index           =   15
      X1              =   224
      X2              =   240
      Y1              =   231
      Y2              =   352
   End
   Begin VB.Line linePlayer 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      DrawMode        =   15  'Merge Pen Not
      Index           =   14
      X1              =   208
      X2              =   224
      Y1              =   231
      Y2              =   352
   End
   Begin VB.Line linePlayer 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      DrawMode        =   15  'Merge Pen Not
      Index           =   13
      X1              =   192
      X2              =   208
      Y1              =   231
      Y2              =   352
   End
   Begin VB.Line linePlayer 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   3
      DrawMode        =   15  'Merge Pen Not
      Index           =   12
      X1              =   176
      X2              =   192
      Y1              =   232
      Y2              =   352
   End
   Begin VB.Line linePlayer 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   4
      DrawMode        =   15  'Merge Pen Not
      Index           =   11
      X1              =   160
      X2              =   176
      Y1              =   232
      Y2              =   352
   End
   Begin VB.Line linePlayer 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   4
      DrawMode        =   15  'Merge Pen Not
      Index           =   10
      X1              =   144
      X2              =   160
      Y1              =   232
      Y2              =   352
   End
   Begin VB.Line linePlayer 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      DrawMode        =   15  'Merge Pen Not
      Index           =   9
      X1              =   128
      X2              =   144
      Y1              =   232
      Y2              =   353
   End
   Begin VB.Line linePlayer 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      DrawMode        =   15  'Merge Pen Not
      Index           =   8
      X1              =   112
      X2              =   128
      Y1              =   232
      Y2              =   352
   End
   Begin VB.Line linePlayer 
      BorderColor     =   &H00808080&
      BorderWidth     =   6
      DrawMode        =   15  'Merge Pen Not
      Index           =   7
      X1              =   96
      X2              =   112
      Y1              =   232
      Y2              =   352
   End
   Begin VB.Line linePlayer 
      BorderColor     =   &H00707070&
      BorderWidth     =   7
      DrawMode        =   15  'Merge Pen Not
      Index           =   6
      X1              =   80
      X2              =   96
      Y1              =   232
      Y2              =   352
   End
   Begin VB.Line linePlayer 
      BorderColor     =   &H00606060&
      BorderWidth     =   8
      DrawMode        =   15  'Merge Pen Not
      Index           =   5
      X1              =   64
      X2              =   80
      Y1              =   232
      Y2              =   352
   End
   Begin VB.Line linePlayer 
      BorderColor     =   &H00505050&
      BorderWidth     =   8
      DrawMode        =   15  'Merge Pen Not
      Index           =   4
      X1              =   48
      X2              =   64
      Y1              =   232
      Y2              =   352
   End
   Begin VB.Line linePlayer 
      BorderColor     =   &H00404040&
      BorderWidth     =   9
      DrawMode        =   15  'Merge Pen Not
      Index           =   3
      X1              =   32
      X2              =   48
      Y1              =   231
      Y2              =   352
   End
   Begin VB.Line linePlayer 
      BorderColor     =   &H00303030&
      BorderWidth     =   10
      DrawMode        =   15  'Merge Pen Not
      Index           =   2
      X1              =   16
      X2              =   32
      Y1              =   232
      Y2              =   352
   End
   Begin VB.Line linePlayer 
      BorderColor     =   &H00202020&
      BorderWidth     =   11
      DrawMode        =   15  'Merge Pen Not
      Index           =   1
      X1              =   0
      X2              =   16
      Y1              =   234
      Y2              =   352
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   50
      X1              =   784
      X2              =   800
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   49
      X1              =   768
      X2              =   784
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   48
      X1              =   752
      X2              =   768
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   47
      X1              =   736
      X2              =   752
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   46
      X1              =   720
      X2              =   736
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   45
      X1              =   704
      X2              =   720
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   44
      X1              =   688
      X2              =   704
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   43
      X1              =   672
      X2              =   688
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   42
      X1              =   656
      X2              =   672
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   41
      X1              =   640
      X2              =   656
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   40
      X1              =   624
      X2              =   640
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   39
      X1              =   608
      X2              =   624
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   38
      X1              =   592
      X2              =   608
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   37
      X1              =   576
      X2              =   592
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   36
      X1              =   560
      X2              =   576
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   35
      X1              =   544
      X2              =   560
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   34
      X1              =   528
      X2              =   544
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   33
      X1              =   512
      X2              =   528
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   32
      X1              =   496
      X2              =   512
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   31
      X1              =   480
      X2              =   496
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   30
      X1              =   464
      X2              =   480
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   29
      X1              =   448
      X2              =   464
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   28
      X1              =   432
      X2              =   448
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   27
      X1              =   416
      X2              =   432
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   26
      X1              =   400
      X2              =   416
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   25
      X1              =   384
      X2              =   400
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   24
      X1              =   368
      X2              =   384
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   23
      X1              =   352
      X2              =   368
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   22
      X1              =   336
      X2              =   352
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   21
      X1              =   320
      X2              =   336
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   20
      X1              =   304
      X2              =   320
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   19
      X1              =   288
      X2              =   304
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   18
      X1              =   272
      X2              =   288
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   17
      X1              =   256
      X2              =   272
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   16
      X1              =   240
      X2              =   256
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   15
      X1              =   224
      X2              =   240
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   14
      X1              =   208
      X2              =   224
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   13
      X1              =   192
      X2              =   208
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   12
      X1              =   176
      X2              =   192
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   11
      X1              =   160
      X2              =   176
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   10
      X1              =   144
      X2              =   160
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   9
      X1              =   128
      X2              =   144
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   8
      X1              =   112
      X2              =   128
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   7
      X1              =   96
      X2              =   112
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   6
      X1              =   80
      X2              =   96
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   5
      X1              =   64
      X2              =   80
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   4
      X1              =   48
      X2              =   64
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   3
      X1              =   32
      X2              =   48
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   2
      X1              =   16
      X2              =   32
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineBot 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   1
      X1              =   0
      X2              =   16
      Y1              =   576
      Y2              =   600
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   50
      X1              =   784
      X2              =   800
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   49
      X1              =   768
      X2              =   784
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   48
      X1              =   752
      X2              =   768
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   47
      X1              =   736
      X2              =   752
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   46
      X1              =   720
      X2              =   736
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   45
      X1              =   704
      X2              =   720
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   44
      X1              =   688
      X2              =   704
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   43
      X1              =   672
      X2              =   688
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   42
      X1              =   656
      X2              =   672
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   41
      X1              =   640
      X2              =   656
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   40
      X1              =   624
      X2              =   640
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   39
      X1              =   608
      X2              =   624
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   38
      X1              =   592
      X2              =   608
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   37
      X1              =   576
      X2              =   592
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   36
      X1              =   560
      X2              =   576
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   35
      X1              =   544
      X2              =   560
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   34
      X1              =   528
      X2              =   544
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   33
      X1              =   512
      X2              =   528
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   32
      X1              =   496
      X2              =   512
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   31
      X1              =   480
      X2              =   496
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   30
      X1              =   464
      X2              =   480
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   29
      X1              =   448
      X2              =   464
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   28
      X1              =   432
      X2              =   448
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   27
      X1              =   416
      X2              =   432
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   26
      X1              =   400
      X2              =   416
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   25
      X1              =   384
      X2              =   400
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   24
      X1              =   368
      X2              =   384
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   23
      X1              =   352
      X2              =   368
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   22
      X1              =   336
      X2              =   352
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   21
      X1              =   320
      X2              =   336
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   20
      X1              =   304
      X2              =   320
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   19
      X1              =   288
      X2              =   304
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   18
      X1              =   272
      X2              =   288
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   17
      X1              =   256
      X2              =   272
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   16
      X1              =   240
      X2              =   256
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   15
      X1              =   224
      X2              =   240
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   14
      X1              =   208
      X2              =   224
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   13
      X1              =   192
      X2              =   208
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   12
      X1              =   176
      X2              =   192
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   11
      X1              =   160
      X2              =   176
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   10
      X1              =   144
      X2              =   160
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   9
      X1              =   128
      X2              =   144
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   8
      X1              =   112
      X2              =   128
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   7
      X1              =   96
      X2              =   112
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   6
      X1              =   80
      X2              =   96
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   5
      X1              =   64
      X2              =   80
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   4
      X1              =   48
      X2              =   64
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   3
      X1              =   32
      X2              =   48
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   2
      X1              =   16
      X2              =   32
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line LineTop 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   15  'Merge Pen Not
      Index           =   1
      X1              =   0
      X2              =   16
      Y1              =   0
      Y2              =   32
   End
   Begin VB.Line lineObstacle 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   3
      DrawMode        =   15  'Merge Pen Not
      Index           =   2
      X1              =   763
      X2              =   763
      Y1              =   189
      Y2              =   266
   End
   Begin VB.Line lineObstacle 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   3
      DrawMode        =   15  'Merge Pen Not
      Index           =   1
      X1              =   742
      X2              =   742
      Y1              =   189
      Y2              =   266
   End
   Begin VB.Line lineObstacle 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   3
      DrawMode        =   15  'Merge Pen Not
      Index           =   0
      X1              =   721
      X2              =   721
      Y1              =   189
      Y2              =   266
   End
   Begin VB.Shape scoreborder 
      BorderColor     =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   390
      Left            =   5145
      Top             =   4200
      Visible         =   0   'False
      Width           =   3210
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'*                  jfCave v Pi (3.14)                  *
'*  You have the rights to use/modify/destroy this code *
'*  or whatever, just please email me first, and        *
'*  possibly even put my name in your credits (if you   *
'*  learned anything from this code.  I take no         *
'*  responsability for blablabla you get the idea       *
'*  If you're curious about how anything was done       *
'*  (SIMPLE questions only please!) email me at         *
'*  drallorf4@hotmail.com                               *
'*                                                      *
'*  (C) 2001 By Jamie Frost a.k.a. Floppy Disk          *
'*                                                      *
'********************************************************
Option Explicit

Private Sub Form_Load()

    frmMain.Show
    cmdExit.Visible = False
    
    initvars   'initalize variables to starting values
    DrawLines  'draw all the lines how they should look when it starts
    
    MainLoop  'duh

End Sub

Private Sub cmdExit_Click()
    ExitApp = True 'flag program for termination...i.e...quit
End Sub

'if the mouse moves, capture where its at
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    intMouseY = Y
End Sub

'keep the "framerate" of the timer reasonable
Private Sub tmrScoreUpdate_Timer()
    frmMain.lblScore.Caption = "Score:" + Str(Score)
End Sub

Sub MainLoop()

    Do Until (LoseFlag)
        DoEvents
        
        'SpeedLimiter
        Do Until GetTickCount >= NextTick
            DoEvents
        Loop: NextTick = GetTickCount + intSpeed
        
        
        ChekMouse            'move according to mouse pos
        UpdateArrays         'update arrays
        ChangeEnd            'change the end value to appropriate cave size
        SetObs               'Start an obstacle on its path
        DrawLines            'draw lines
        CheckForCollisions   'check for collisions
        AddScore             'add appropriate value to score.
    Loop
    If LoseFlag Then

        tmrScoreUpdate.Enabled = False
           
        'GAME OVER!!!
                
        BlowUp         'play animation at y coord of
                       'player for spectacular finish
        ShowStats      'make stats elements visible
        frmMain.lblScore.Caption = "Score:" + Str(Score)
        
        
    End If
    
    Do Until ExitApp
    DoEvents
    Loop
    End
    
End Sub

