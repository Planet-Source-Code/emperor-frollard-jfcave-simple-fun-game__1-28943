VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "jfCave Setup"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   3135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdContact 
      Caption         =   "Contact Me"
      Height          =   435
      Left            =   1260
      TabIndex        =   12
      Top             =   1890
      Width           =   1065
   End
   Begin VB.CommandButton cmdInstructions 
      Caption         =   "Instructions"
      Height          =   435
      Left            =   105
      TabIndex        =   11
      Top             =   1890
      Width           =   960
   End
   Begin VB.Frame frmDifficulty 
      Caption         =   "Difficulty"
      Height          =   1590
      Left            =   1680
      TabIndex        =   7
      Top             =   105
      Width           =   1380
      Begin VB.OptionButton optDiff 
         Caption         =   "HARD"
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   10
         Top             =   1155
         Width           =   1065
      End
      Begin VB.OptionButton optDiff 
         Caption         =   "NORMAL"
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   9
         Top             =   735
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton optDiff 
         Caption         =   "EASY"
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   8
         Top             =   315
         Width           =   1065
      End
   End
   Begin VB.Frame frmSpeed 
      Caption         =   "Speed"
      Height          =   1590
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   1380
      Begin VB.OptionButton optSpd 
         Caption         =   "CRAWLING"
         Height          =   225
         Index           =   4
         Left            =   105
         TabIndex        =   6
         Top             =   1155
         Width           =   1170
      End
      Begin VB.OptionButton optSpd 
         Caption         =   "SLOW"
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   5
         Top             =   945
         Width           =   1170
      End
      Begin VB.OptionButton optSpd 
         Caption         =   "NORMAL"
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   4
         Top             =   735
         Width           =   1170
      End
      Begin VB.OptionButton optSpd 
         Caption         =   "FASTER"
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   3
         Top             =   525
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton optSpd 
         Caption         =   "FASTEST"
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   2
         Top             =   315
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "GO!"
      Height          =   435
      Left            =   2520
      TabIndex        =   0
      Top             =   1890
      Width           =   540
   End
End
Attribute VB_Name = "frmOptions"
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
Private Sub cmdStart_click()
    Dim i As Integer
    Dim intTemp As Integer
    
    
    For i = 0 To 4   'check value of speed
        If optSpd(i) Then intTemp = i
    Next i
    
    Select Case intTemp
        Case 0
            intSpeed = 10
        Case 1
            intSpeed = 12
        Case 2
            intSpeed = 15
        Case 3
            intSpeed = 20
        Case 4
            intSpeed = 30
    End Select

    For i = 0 To 2    'check difficulty
        If optDiff(i) Then intTemp = i
    Next i

    Select Case intTemp 'set difficulty (responsiveness)
        Case 0
            intResponsiveness = 10    'accelerate by 1/10th of the dist to mouse
            intSmallest = 125         'smallest possible cave
        Case 1
            intResponsiveness = 20    'accelerate by 1/20th of the ...blabla
            intSmallest = 100         'smallest possible cave
        Case 2
           intResponsiveness = 30    'same thing again
            intSmallest = 80          'smallest possible cave
    End Select

  Unload Me     'get rid of options dialog
  Load frmMain  'show the game, and get started


End Sub

Private Sub cmdInstructions_Click()
    Dim vbOkOnly As Integer
    vbOkOnly = MsgBox("Ok, this is my FIRST VB program EVER.  I have had some experience before using turbo-pascal, but thats it, so this is a real big thing for me.  How to play:  basically, you fly thru a cave, and you steer with your mouse.  You get more and more points as the cave gets smaller and smaller.  Also, you get WAY less points if you arent actually doing anything...so fly like a teenager who just got his licence, or NO HISCORE FOR YOU! btw, 'Easy' just makes your 'ship' more responsive, and the speed, well if you can't figure that out, then you shouldn't be playing games.  Enjoy", vbOkOnly, "The way it is")
End Sub

Private Sub cmdContact_Click()
    Dim vbOkOnly As Integer
    vbOkOnly = MsgBox("There are many ways you can get a hold of me for well, anything, but I'll only tell you one:  drallorf4@hotmail.com  That is all.", vbOkOnly, "Me")
End Sub

