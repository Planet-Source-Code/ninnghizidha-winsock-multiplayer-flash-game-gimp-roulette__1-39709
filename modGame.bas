Attribute VB_Name = "modGame"
Option Explicit

Public Const cFlashSound = "\sound.swf"      ' The Background-Sound
Public Const cFlashGame = "\game.swf"        ' The Game
Public Const cGameName = "Gimp Roulette"     ' Game-name for the Header

Public gstrDrum As String                    ' The Drum of the Gun
Public gstrLocalNick As String               ' local Nick
Public gstrRemoteNick As String              ' The Remote Nick
Public gstrLocalIP As String                 ' the Local IP
Public gstrRemoteIP As String                ' the remote IP
Public gintLocalPoints As Integer            ' My Points
Public gintRemotePoints As Integer           ' your points
Public glngTime As Long                      ' how log did the game took

'------------
Public Const cFlashSTARTUP = 1               ' StartUp
Public Const cFlashGIMPSTARTUP = 2           ' Weapon in the Middle (StartUp)
Public Const cFlashREMOTEPICKUP = 3          ' RemotePlayer Picks up the Gun
Public Const cFlashREMOTEROLLDRUM = 5        ' RemotePlayer rolles the Drum
Public Const cFlashREMOTENOROLLDRUM = 4      ' RemotePlayer doesnt roll the drum
Public Const cFlashREMOTEPULLTRIGGER = 7     ' RemotePlayer pulls the Trigger and lays down the gun
Public Const cFlashREMOTEPULLTRIGGERDIE = 6  ' RemotePlayer blows away the head.

Public Const cFlashLOCALPICKUP = 8           ' You pick up the Gun
Public Const cFlashLOCALROLLDRUM = 9         ' You roll the drum
Public Const cFlashLOCALNOROLLDRUM = 10      ' You don't roll the drum
Public Const cFlashLOCALPULLTRIGGER = 12     ' You pull the trigger and are alive
Public Const cFlashLOCALPULLTRIGGERDIE = 11  ' You pull the Trigger an die. hehe



Public Function Rotate(pString As String) As String
' rotates the Drum of the gun.

    Dim intSpeed As Integer
    Dim strString As String
    Dim i As Integer
    Const upperbound = 13
    Const lowerbound = 3
    
    strString = pString
    
    Randomize
    intSpeed = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
    
    'roll the Gun, and look at the code: It really rolls :D
    For i = 0 To intSpeed
        strString = Right(strString, Len(strString) - 1) & Left(strString, 1)
    Next
    
    'Debug.Print = "Drum rolled " & intSpeed & " times to " & strString
    Rotate = strString
End Function


Public Function Shoot(pString As String) As String
' pulls the trigger - the bullets move

    Dim strString As String
    
    strString = pString
    strString = Right(strString, Len(strString) - 1) & Left(strString, 1)
    Shoot = strString
    
End Function

Public Function IsHeadshot() As Boolean
' just checks, if bullet was loaded.
    
    IsHeadshot = False
    If Right(gstrDrum, 1) = "1" Then IsHeadshot = True
    ' if bullet is on the right side, it was just shot.
    
End Function


Public Function SetStatus(pstrString)
' The Function, which prints the Status
    
    'frmMain.txtChat.Text = frmMain.txtChat.Text & pstrString & vbCrLf
    frmMain.lstChat.AddItem pstrString
    frmMain.lstChat.ListIndex = frmMain.lstChat.ListCount - 1
End Function


Public Function PlayFlashFrom(pintFrameNo As Integer)
    'frmMain.flshAnimation.StopPlay
    frmMain.flshAnimation.GotoFrame (pintFrameNo)
    frmMain.flshAnimation.StopPlay
    'frmMain.flshAnimation.
    'frmMain.flshAnimation.Play
    'frmMain.flshAnimation.f
End Function

Public Function actionNewGame(pboolIsLocalAction As Boolean)
    If pboolIsLocalAction Then
        SetStatus ("You started a new game.")
        PlayFlashFrom (cFlashSTARTUP)
    Else
        SetStatus (gstrRemoteNick & " joined the your game.")
        SendData cOK
        PlayFlashFrom (cFlashSTARTUP)
    End If
    
End Function

Public Function actionWelcome(pboolIsLocalAction As Boolean)
    If pboolIsLocalAction Then
        SetStatus ("You joined an existing game")
        PlayFlashFrom (cFlashLOCALPULLTRIGGER)
    Else
        SetStatus (gstrRemoteNick & " joined your game..")
        Call RecalibrateWinsock
        PlayFlashFrom (cFlashSTARTUP)
    End If
    
End Function

Public Function actionDisconnect(pboolIsLocalAction As Boolean)
    If pboolIsLocalAction Then
        ' nothing - your are leaving ^^
    Else
        SetStatus (gstrRemoteNick & " left the game ...")
        MsgBox (gstrRemoteNick & " left the game ...")
        gboolConnected = False
    End If
    
End Function


Public Function actionOK(pboolIsLocalAction As Boolean)
    
    gintLocalPoints = 0
    gintRemotePoints = 0
    
    
    gboolConnected = True
    frmMain.tmrDrum.Enabled = True
    
    'Call genPercentBar(frmMain.pbPercentbar, frmMain.lblPercent, CLng("0" & GetOption("Stats", "MaxTriggers")), 0)
    frmMain.lbMaxTriggers = CLng("0" & GetOption("Stats", "MaxTriggers"))
    
    With frmMain
        .cmdNewGame.Visible = False
        .cmdDrum.Enabled = False
        .cmdNoDrum.Enabled = False
        .cmdPullTrigger.Enabled = False
        
        If pboolIsLocalAction Then
            .cmdPickUp.Enabled = False
            gstrDrum = Rotate(gstrDrum)
        Else
            Call ClientJoinedOK
            .cmdPickUp.Enabled = True
        End If
    End With
    
End Function

Public Function actionPickUp(pboolIsLocalAction As Boolean)
    
    If pboolIsLocalAction Then
        SetStatus ("You pick up the gun.")
        PlayFlashFrom cFlashLOCALPICKUP
            frmMain.cmdDrum.Enabled = True
            frmMain.cmdNoDrum.Enabled = True
            frmMain.cmdPickUp.Enabled = False
    Else
        SetStatus (gstrRemoteNick & " picks up the gun.")
        PlayFlashFrom cFlashREMOTEPICKUP
    End If
  
End Function

Public Function actionDrum(pboolIsLocalAction As Boolean)
    If pboolIsLocalAction Then
        SetStatus ("You roll the drum of the gun.")
        PlayFlashFrom cFlashLOCALROLLDRUM
            frmMain.cmdDrum.Enabled = False
            frmMain.cmdNoDrum.Enabled = False
            frmMain.cmdPullTrigger.Enabled = True
    Else
        SetStatus (gstrRemoteNick & " rolls the drum of the gun.")
        PlayFlashFrom cFlashREMOTEROLLDRUM
    End If
    
End Function

Public Function actionNoDrum(pboolIsLocalAction As Boolean)
    If pboolIsLocalAction Then
        SetStatus ("You don't roll the drum of the gun.")
        PlayFlashFrom cFlashLOCALNOROLLDRUM
            frmMain.cmdDrum.Enabled = False
            frmMain.cmdNoDrum.Enabled = False
            frmMain.cmdPullTrigger.Enabled = True
    Else
        SetStatus (gstrRemoteNick & " doesn't roll the drum of the gun.")
        PlayFlashFrom cFlashREMOTENOROLLDRUM
    End If
End Function

Public Function actionPullTrigger(pboolIsLocalAction As Boolean)
Dim SetOptionTemp As Variant
Dim lngStats As Long
Dim lngMaxTriggers As Long
Dim lngMaxTime As Long

    If IsHeadshot() Then
        glngTime = 0 'Set TimeCounter to 0
        
        If pboolIsLocalAction Then
            SetStatus ("You died and " & gstrRemoteNick & " laughs.")
            PlayFlashFrom cFlashLOCALPULLTRIGGERDIE
                frmMain.cmdPullTrigger.Enabled = False
                frmMain.cmdNewGame.Visible = True 'Lose have to push the New-Game-Button
                'Save Stats for Loser
                SetOptionTemp = SetOption("Stats", "Lost", CStr(CLng("0" & GetOption("Stats", "Lost")) + 1))
        Else
            SetStatus (gstrRemoteNick & " died.")
            PlayFlashFrom cFlashREMOTEPULLTRIGGERDIE
                'Save Stats for Winner
                SetOptionTemp = SetOption("Stats", "Won", CStr(CLng("0" & GetOption("Stats", "Won")) + 1))
        End If
        frmMain.tmrDrum.Enabled = False
    Else
        If pboolIsLocalAction Then
            SetStatus ("You pull the trigger, but nothing happens.")
            PlayFlashFrom cFlashLOCALPULLTRIGGER
                frmMain.cmdPullTrigger.Enabled = False
                gintLocalPoints = gintLocalPoints + 1
              
        Else
            SetStatus (gstrRemoteNick & " pulls the trigger, but nothing happens.")
            PlayFlashFrom cFlashREMOTEPULLTRIGGER
                frmMain.cmdPickUp.Enabled = True
                gintRemotePoints = gintRemotePoints + 1
        End If
    End If
End Function




Public Function GotoGame()
Dim Temp As Variant
 
    Temp = SetOption("Settings", "LastIP", frmStart.txtIPAddress.Text)
    Temp = SetOption("Settings", "Nick", frmStart.txtNick.Text)
    
    gstrLocalNick = Trim(Replace(frmStart.txtNick.Text, DataDelemiter, "&"))
    frmMain.flshAnimation.SetVariable "spieler_name", gstrLocalNick
    
    frmStart.Hide
    Unload frmStart
    frmMain.Show
    
    
    gstrLocalIP = frmMain.WinSock.LocalIP
    'If frmStart.chkFirewall.Value = 1 Then
    '    gstrLocalIP = GetInternetIP(True)
    'End If
    
    'frmMain.WinSock.LocalIP
    frmStart.tmrNick.Enabled = False
    
    gstrRemoteNick = "Opponent"
    
End Function

Public Function SetTitle(pstrMessage As String)
Dim strServerMode As String

    If gboolIsServer Then
        strServerMode = "Server-Mode"
    Else
        strServerMode = "Client-Mode"
    End If
    
    frmMain.Caption = cGameName & " " & gstrLocalNick & " - " & strServerMode & ": " & pstrMessage
    
    If pstrMessage = "" Then
        frmMain.Caption = frmMain.Caption & "ok"
    End If
End Function

Public Function DoStatistics()
Dim lngMaxTime As Long
Dim lngMaxTriggers As Long
Dim SetOptionTemp As Variant
    
    lngMaxTime = CLng("0" & GetOption("Stats", "MaxTime"))
    
    
    lngMaxTime = CLng("0" & GetOption("Stats", "MaxTime"))
    lngMaxTriggers = CLng("0" & GetOption("Stats", "MaxTriggers"))
    
     
    
    With frmMain
        .lblMaxTime.Caption = lngMaxTime \ 10 & "s"
    
        If lngMaxTime < glngTime Then
            SetOptionTemp = SetOption("Stats", "MaxTime", CStr(glngTime))
            .lblMaxTime.Caption = lngMaxTime & " (NEW!)"
        End If
        
        If lngMaxTriggers < gintLocalPoints Then
            SetOptionTemp = SetOption("Stats", "MaxTriggers", CStr(gintLocalPoints))
            .lbMaxTriggers.Caption = gintLocalPoints & "ms (NEW!)"
        End If



    'genPercentBar(pPictureBox As PictureBox, pLabelPercent As Label, plngMax As Long, plngPart As Long)

        Call genPercentBar(.pbPercentRounds, .lblPercentRounds, lngMaxTriggers, CLng(gintLocalPoints))
        Call genPercentBar(.pbPercentTime, .lblPercentTime, lngMaxTime, glngTime)
        Call genPercentBar(.pbPercentLost, .lblPercentLost, CLng("0" & GetOption("Stats", "Won")) + CLng("0" & GetOption("Stats", "Lost")), CLng("0" & GetOption("Stats", "Lost")))
    End With
End Function
