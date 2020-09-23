Attribute VB_Name = "modWinSock"
Option Explicit

Public gboolConnected As Boolean
Public gboolIsServer As Boolean
Public Const gboolShowData = False

Public Const DataDelemiter = "+"

Public Const DataTCPPort = "9991"
Public Const DataIdentifier = "data"

Public intCheckAliveCounter As Integer
'Public gboolConnected As Boolean
'-------------
Public Const cWELCOME = "welcome"       ' Joined the Game, important for Server-Remotehost
Public Const cOK = "ok"                 ' Server Running Gimp-Game and listening.
Public Const cTEST = "test"             ' Test-Trigger - no need
Public Const cIDLEBUTALIVE = "idle"     ' Idle but Alive - Check
Public Const cNEWGAME = "newgame"       ' Start a new game
Public Const cDISCONNECT = "disconnect" ' disconnect
'-------------
Public Const cPICKUP = "pickup"         ' pickup the gun
Public Const cROLLDRUM = "rotate"       ' roll the drum with the bullet
Public Const cNOROLLDRUM = "norotate"   ' dont roll the drum
Public Const cPULLTRIGGER = "shoot"     ' pull the trigger



Public Function ProcessDataString(pstrData As String) As String
Dim pstrDataTemp As Variant
Dim strAction As String
Dim boolIsLocalData As Boolean

    pstrDataTemp = Split(pstrData, DataDelemiter)

    gstrDrum = pstrDataTemp(4)
    strAction = pstrDataTemp(3)
        'gstrNick = pstrDataTemp(3)
        
    boolIsLocalData = IsLocalDataString(pstrDataTemp(2), pstrDataTemp(1))
        
    If Not boolIsLocalData Then
        gstrRemoteNick = pstrDataTemp(2)
        gstrRemoteIP = pstrDataTemp(1)
        gintRemotePoints = Int(pstrDataTemp(5))
        frmMain.flshAnimation.SetVariable "gegner_name", gstrRemoteNick
    End If
    
    'Load the actions we got.
    Select Case strAction
        Case cWELCOME: Call actionWelcome(boolIsLocalData)
        Case cOK: Call actionOK(boolIsLocalData)
        Case cTEST:
        Case cIDLEBUTALIVE:
        Case cNEWGAME: Call actionNewGame(boolIsLocalData)
        Case cDISCONNECT: Call actionDisconnect(boolIsLocalData)
        
        Case cPICKUP: Call actionPickUp(boolIsLocalData)
        Case cROLLDRUM: Call actionDrum(boolIsLocalData)
        Case cNOROLLDRUM: Call actionNoDrum(boolIsLocalData)
        Case cPULLTRIGGER: Call actionPullTrigger(boolIsLocalData)
    End Select

End Function

Public Function IsLocalDataString(ByVal pstrNick As String, ByVal pstrIP As String) As Boolean
Dim pstrDataTemp As Variant
    
    'just defines, if an action was/is local. It checks the nick and the IP
    IsLocalDataString = False
    If gstrLocalNick = pstrNick And gstrLocalIP = pstrIP Then IsLocalDataString = True

End Function





Public Function SendData(pstrAction As String) As Boolean
Dim strSenddata As String
    
    strSenddata = DataIdentifier & DataDelemiter
    strSenddata = strSenddata & gstrLocalIP & DataDelemiter ' add LocalIP to DataString
    strSenddata = strSenddata & gstrLocalNick & DataDelemiter ' add Nickname to Datastring
    strSenddata = strSenddata & LCase(pstrAction) & DataDelemiter ' add Action to Datastring
    strSenddata = strSenddata & gstrDrum & DataDelemiter ' add Drum to Datastring
    strSenddata = strSenddata & gintLocalPoints 'Add Counter
    
    DoEvents
    
    If gboolShowData Then ' turn this to true, if you want see the data.
        SetStatus strSenddata
    End If
    
    ' Since its an 2-player-Game, i need the data too - so: process it :)
    ProcessDataString (strSenddata)
    
    'If frmMain.WinSock.State = 1 Then
        On Error Resume Next
            frmMain.WinSock.SendData strSenddata
        
        If Err > 0 Then 'And Not gboolConnected Then
            MsgBox (Err.Description & vbCrLf & "Please wait till connected")
        End If
        On Error GoTo 0
    'End If
  
    
End Function

Public Function IsDataString(pstrData As String) As Boolean
Dim pstrDataTemp As Variant

    If Left(pstrData, Len(DataIdentifier)) = DataIdentifier Then
        pstrDataTemp = Split(pstrData, DataDelemiter)
        If UBound(pstrDataTemp) <> 5 Then
            'corruped DataString
            IsDataString = False
        Else
            IsDataString = True
        End If
    Else
        IsDataString = False
    End If

End Function

Public Function RecalibrateWinsock()
    Call SetTitle("")
    SendData cOK
End Function


Public Function ClientJoinedOK()
    'Join ok, pretty useless :>
    Call SetTitle("")
End Function

