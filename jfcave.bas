Attribute VB_Name = "MainCode"
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
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

 
 Public Type Obstacle                 ' Obstacle parameters
    X1 As Integer
    X2 As Integer
    Y1 As Integer
    Y2 As Integer
    Active As Boolean
 End Type
 
 Const Pi As Double = 3.14159265358979 'pi...
 Public ExitApp As Boolean            ' true if quit pressed
 Public LoseFlag As Boolean           ' true if wall hit
  
 Public intPlayerPosY(15) As Integer   ' the player's line
 
 Public intTopPos(50) As Integer      ' the stalactites' positions
 Public intBotPos(50) As Integer      ' the stalagmites' positions
 
 Public Obs(2) As Obstacle            ' the obstacle paramaters
 
 Public intCaveHeight As Integer      ' the height of the cave
 
 Public intVector As Integer          ' the players vertical vector
 Public intAccel As Integer           ' acceleration of gravity and 'up'
 Public intResponsiveness As Integer  ' responsiveness of the ship
                                      ' bigger = less
 Public intMouseY As Integer          ' mouse y value
 Public intSpeed As Integer           ' horizontal speed (lower = faster)
 Public intRand As Integer            ' random chance integer
 Public intSmallSpeed As Integer      ' speed @ which cave decreases
                                      ' smaller = faster
 Public intSmallJump As Integer       ' amount to jump (pixels) per smaller
 Public intMaxRand As Integer         ' maxamum random jump by cave itself
 Public intSmallest As Integer        ' Smallest possible cave height
 Public intTimeTilObs As Integer      ' number of frames until an obstacle
                                      ' can be released
 Public NextTick As Long
 Public Score As Long                 ' score
 
 
 
 '**************************** Initialize Variables

Public Sub initvars()
    Dim i As Integer            ' for loop counter
        intTimeTilObs = 300
        intCaveHeight = 560
        'intResponsiveness = 20  ' bigger = worse responce
        'intSpeed = 10            ' smaller = faster
        'intSmallest = 100       ' smallest cave can get
        intSmallJump = 6        ' # of pixels to jump
        intSmallSpeed = 60      ' 1/n chance that itll get smaller
        intMaxRand = 15         ' max amount that cave will jump
    
    For i = 0 To 15             ' set playerline to center
       intPlayerPosY(i) = 300
       
    Next i


    For i = 0 To 50  ' set top and bottom to be at top and bottom
        intTopPos(i) = 20
        intBotPos(i) = 580
    Next i
    
    For i = 0 To 2     'set obstacles outside right side of window
        Obs(i).X1 = 801
        Obs(i).X2 = 801
        Obs(i).Active = False
    Next i
End Sub



'*************************** ScoreKeep
Public Sub AddScore()
    Dim Slope As Single
    Slope = Abs((frmMain.linePlayer(15).Y2 - frmMain.linePlayer(15).Y1) / (frmMain.linePlayer(15).X2 - frmMain.linePlayer(15).X1))
    Slope = (Slope / 2) + 0.5
    Score = Score + Int(Slope * (560 - intCaveHeight))
End Sub

'**************************** Animate explosion
Public Sub BlowUp()
    Dim i As Integer     'for loop again
    Dim intExpFrame As Integer   'frame number of explosion
    
    For intExpFrame = 0 To 16
        frmMain.linePlayer(15).X1 = frmMain.linePlayer(15).X2
    
        For i = 1 To 14 Step 1     'move the players lines
        
            frmMain.linePlayer(i).X2 = frmMain.linePlayer(i + 1).X2
            frmMain.linePlayer(i).Y1 = frmMain.linePlayer(i + 1).Y1
            
            frmMain.linePlayer(i).X1 = frmMain.linePlayer(i + 1).X1
            frmMain.linePlayer(i).Y2 = frmMain.linePlayer(i + 1).Y2
            DoEvents
        Next i
    
        'SpeedLimiter loop
        Do Until GetTickCount >= NextTick
                'DoEvents
        Loop: NextTick = GetTickCount + intSpeed
    Next intExpFrame
End Sub


'**************************** Make end elements visible
Public Sub ShowStats()
    Dim i As Integer    'YET AGAIN...counter
    Dim tempi As Integer
    
    frmMain.lblGameOver.Visible = True      'set "GAME OVER" to visible
    frmMain.cmdExit.Visible = True          'show the quit button
    frmMain.lblScore.Visible = False        'hide the top scorebox
    frmMain.tmrScoreUpdate.Enabled = False  'stop updating the top scorebox
    
    'animate game over sign
    For i = 0 To 8
        tempi = i
        If i > 4 Then tempi = i + 1
        frmMain.lblGameOver.Caption = Left("GAME OVER", tempi)
        'SpeedLimiter
        Do Until GetTickCount >= NextTick
            DoEvents
        Loop: NextTick = GetTickCount + 200
    Next i
    
    frmMain.scoreborder.Visible = True        'show the border around the final score
    frmMain.lblFinalScore(0).Visible = True   'show final score itself
    NextTick = GetTickCount + 500
    
    Do Until GetTickCount >= NextTick
        DoEvents
    Loop
    
    frmMain.lblFinalScore(1).Caption = Str(Score)   'set finalscore value
    frmMain.lblFinalScore(1).Visible = True         'Show the actual Score Number
    
    
    
End Sub




'**************************** Try to start an obstacle
Public Sub SetObs()
    Dim i As Integer              'for loop counter
    Dim intAv As Integer          'available position
    Dim intObsLen As Integer      'length of obstacle
    intObsLen = 0.2 * intCaveHeight
    
    If intTimeTilObs > 0 Then intTimeTilObs = intTimeTilObs - 1
    intAv = 3
    For i = 2 To 0 Step -1     'find out what the first available line is
        If Obs(i).Active = False Then intAv = i
    Next i
    If intAv <> 3 Then 'if theres an available line to start....
        If (Int(Rnd * 10)) = 0 And intTimeTilObs = 0 Then  'then start it
            Obs(intAv).Active = True           'this obstacle is locked and loaded
            intTimeTilObs = intTimeTilObs + 15 'wait minimum of 15 frames to start another
            ' ***set its position
            Obs(intAv).Y1 = Int(Rnd * (intBotPos(50) - intTopPos(50) - intObsLen)) + intTopPos(50)
            Obs(intAv).Y2 = Obs(intAv).Y1 + intObsLen
            Obs(intAv).X1 = 800
            Obs(intAv).X2 = 800
            
        End If
    End If
End Sub

'**************************** Shift all the values left by 1
Public Sub UpdateArrays()
    Dim i As Integer      ' "for" loop counter
    For i = 0 To 49       'shift top and bottom values over by 1
        intTopPos(i) = intTopPos(i + 1)
        intBotPos(i) = intBotPos(i + 1)
    Next i
    For i = 0 To 14       'shift playerline values over by 1
        intPlayerPosY(i) = intPlayerPosY(i + 1)
    Next i
    intPlayerPosY(15) = intPlayerPosY(15) + intAccel
    For i = 0 To 2        'move obstacle left by 16 pixels..or 1 frame
        If Obs(i).Active Then
            Obs(i).X1 = Obs(i).X1 - 16
            Obs(i).X2 = Obs(i).X2 - 16
        End If
    Next i
    For i = 0 To 2        'move obstacle to start if at end and deactivate
        If Obs(i).X1 <= 0 Then
        Obs(i).Active = False
        Obs(i).X1 = 801
        Obs(i).X2 = 801
        End If
    Next i
End Sub





'**************************** Change end value so cave gets smaller
Public Sub ChangeEnd()
    'determine cave height
    intRand = Int(Rnd * intSmallSpeed) + 1       'random chance
    If (intRand = 1) And (intCaveHeight > intSmallest) Then
        intCaveHeight = intCaveHeight - intSmallJump 'shrink cave by smalljump
    End If
    'end determining cave height
    
    'randomly move top of cave
    If Int(Rnd * 2 + 1) = 1 Then      'move down
        If intTopPos(50) < 600 - intMaxRand - intCaveHeight Then
          intTopPos(50) = intTopPos(50) + Int(Rnd * intMaxRand / 2)
        End If
      Else                            'move up
        If intTopPos(50) > 15 + intMaxRand Then
          intTopPos(50) = intTopPos(50) - Int(Rnd * intMaxRand / 2)
        End If
    End If
    
    If Int(Rnd * 2 + 1) = 1 Then      'move down
        intBotPos(50) = intTopPos(50) + intCaveHeight + intMaxRand / 4
      Else
        intBotPos(50) = intTopPos(50) + intCaveHeight - intMaxRand / 4
    End If
End Sub

'**************************** Check for MouseMovement
Public Sub ChekMouse()
    Dim intPY As Integer
    intPY = intPlayerPosY(15) 'good to know where the player actually is
    'accelerate the mouse based on where the mouse is
    intAccel = Int((intMouseY - intPY) / intResponsiveness)
End Sub

'**************************** Check for collisions
Public Sub CheckForCollisions()
    Dim i As Integer   'yet another for loop counter
    
    'see if they hit the top
    If intPlayerPosY(15) <= intTopPos(15) Then
        LoseFlag = True   'run the lose procedure after moving lines
    End If
    'see if they hit the bottom
    If intPlayerPosY(15) >= intBotPos(15) Then
        LoseFlag = True   'run the lose procedure after moving lines
    End If
    
    For i = 0 To 2   'see if they hit an obstacle
        If Obs(i).X1 = 240 Then  'if theres an obstacle at player's x pos
            If (intPlayerPosY(15) >= Obs(i).Y1) And (intPlayerPosY(15) <= Obs(i).Y2) Then
                LoseFlag = True    'run the lose procedure after moving lines
            End If
        End If
    Next i
End Sub

'**************************** Draw the Lines
Public Sub DrawLines()
    Dim i As Integer      ' "for" loop counter...AGAIN
   
    For i = 0 To 49
        frmMain.LineTop(i + 1).Y1 = intTopPos(i)       'draw top lines
        frmMain.LineTop(i + 1).Y2 = intTopPos(i + 1)   'draw top lines
        frmMain.LineBot(i + 1).Y1 = intBotPos(i)       'draw bottom lines
        frmMain.LineBot(i + 1).Y2 = intBotPos(i + 1)   'draw bottom lines
    Next i
    
    For i = 0 To 14
        frmMain.linePlayer(i + 1).Y1 = intPlayerPosY(i)     'draw player lines
        frmMain.linePlayer(i + 1).Y2 = intPlayerPosY(i + 1) 'draw player lines
    Next i
    
    For i = 0 To 2
        frmMain.lineObstacle(i).X1 = Obs(i).X1  'move line i's x1
        frmMain.lineObstacle(i).X2 = Obs(i).X2  'move line i's x2
        frmMain.lineObstacle(i).Y1 = Obs(i).Y1  'move line i's y1
        frmMain.lineObstacle(i).Y2 = Obs(i).Y2  'move line i's y2
    Next i
End Sub
