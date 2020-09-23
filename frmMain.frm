VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   Caption         =   "Gimpgame"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10365
   FillColor       =   &H80000012&
   Icon            =   "frmMain.frx":0000
   ScaleHeight     =   6495
   ScaleWidth      =   10365
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame Frame_Chat 
      BackColor       =   &H00008000&
      Caption         =   "Chat-Box"
      Height          =   2775
      Left            =   5880
      TabIndex        =   25
      Top             =   3600
      Width           =   4095
      Begin VB.ListBox lstChat 
         ForeColor       =   &H00004000&
         Height          =   2205
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox txtChatLine 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   120
         MaxLength       =   500
         TabIndex        =   26
         Top             =   2400
         Width           =   3855
      End
   End
   Begin VB.Frame Frame_Stats 
      BackColor       =   &H00008000&
      Height          =   1455
      Left            =   120
      TabIndex        =   11
      Top             =   4800
      Width           =   5655
      Begin VB.PictureBox pbPercentLost 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         FillColor       =   &H000000C0&
         FillStyle       =   7  'Diagonalkreuz
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         ScaleHeight     =   11.25
         ScaleMode       =   2  'Punkt
         ScaleWidth      =   269.25
         TabIndex        =   20
         Top             =   1080
         Width           =   5415
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lost Rounds"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Left            =   120
            TabIndex        =   23
            Top             =   0
            Width           =   1050
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Won"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   4950
            TabIndex        =   22
            Top             =   0
            Width           =   360
         End
         Begin VB.Label lblPercentLost 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3240
            TabIndex        =   21
            Top             =   0
            Width           =   585
         End
      End
      Begin VB.PictureBox pbPercentTime 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         FillColor       =   &H000000C0&
         FillStyle       =   7  'Diagonalkreuz
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         ScaleHeight     =   11.25
         ScaleMode       =   2  'Punkt
         ScaleWidth      =   269.25
         TabIndex        =   16
         Top             =   840
         Width           =   5415
         Begin VB.Label lblPercentTime 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3240
            TabIndex        =   19
            Top             =   0
            Width           =   585
         End
         Begin VB.Label lblMaxTime 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   4860
            TabIndex        =   18
            Top             =   0
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Maximum of Time survived"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Left            =   120
            TabIndex        =   17
            Top             =   0
            Width           =   2250
         End
      End
      Begin VB.PictureBox pbPercentRounds 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         FillColor       =   &H000000C0&
         FillStyle       =   7  'Diagonalkreuz
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         ScaleHeight     =   11.25
         ScaleMode       =   2  'Punkt
         ScaleWidth      =   269.25
         TabIndex        =   12
         Top             =   600
         Width           =   5415
         Begin VB.Label lbLost 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Maximum of Rounds survived"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Left            =   120
            TabIndex        =   15
            Top             =   0
            Width           =   2460
         End
         Begin VB.Label lbMaxTriggers 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   5100
            TabIndex        =   14
            Top             =   0
            Width           =   210
         End
         Begin VB.Label lblPercentRounds 
            Alignment       =   1  'Rechts
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3360
            TabIndex        =   13
            Top             =   0
            Width           =   465
         End
      End
      Begin VB.Label lblStats 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   5415
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame_Settings 
      BackColor       =   &H00008000&
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   2055
      Begin VB.CommandButton cmdNewGame 
         Caption         =   "New Game"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdEnd 
         Caption         =   "Quit the Game"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox chkSound 
         BackColor       =   &H00008000&
         Caption         =   "Background Sound"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         ToolTipText     =   "Background Sound?"
         Top             =   720
         Value           =   1  'Aktiviert
         Width           =   1695
      End
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flshSound 
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   6480
      Visible         =   0   'False
      Width           =   375
      _cx             =   661
      _cy             =   661
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
   Begin MSWinsockLib.Winsock WinSock 
      Left            =   6240
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrDrum 
      Interval        =   300
      Left            =   5880
      Top             =   6480
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flshAnimation 
      Height          =   3570
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10005
      _cx             =   17639
      _cy             =   6299
      FlashVars       =   ""
      Movie           =   "C:\Documents and Settings\jsperlhofer\My Documents\My Pogramms\GimpRoulette\gimproulette2.swf"
      Src             =   "C:\Documents and Settings\jsperlhofer\My Documents\My Pogramms\GimpRoulette\gimproulette2.swf"
      WMode           =   "Window"
      Play            =   0   'False
      Loop            =   0   'False
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   0   'False
      Base            =   ""
      Scale           =   "ShowAll"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
   End
   Begin VB.Frame Frame_Gun 
      BackColor       =   &H00008000&
      Caption         =   "Control the Gun"
      Height          =   1095
      Left            =   2280
      TabIndex        =   0
      Top             =   3600
      Width           =   3495
      Begin VB.CommandButton cmdNoDrum 
         Caption         =   "Don't Roll"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdPickUp 
         Caption         =   "Pick Up the Gun"
         Enabled         =   0   'False
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdDrum 
         Caption         =   "Roll Drum"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdPullTrigger 
         Caption         =   "Pull the Trigger"
         Enabled         =   0   'False
         Height          =   735
         Left            =   2400
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkSound_Click()
    
    If chkSound.Value = 1 Then
        ' turn the sound on.
        frmMain.flshSound.Movie = App.Path & cFlashSound
    Else
        ' to turn the sound off -
        ' no movie, no sound :D
        frmMain.flshSound.Movie = "void"
    End If
End Sub

Private Sub cmdDrum_Click()
    'Rotate the Drum, and send the data
    gstrDrum = Rotate(gstrDrum)
    SendData (cROLLDRUM)
End Sub

Private Sub cmdEnd_Click()
    
    If gboolConnected Then
        SendData cDISCONNECT
    End If
    
    frmMain.WinSock.Close

    frmMain.Hide
    Unload frmMain
    End
End Sub

Private Sub cmdNewGame_Click()
    SendData (cNEWGAME)
End Sub

Private Sub cmdNoDrum_Click()
    SendData (cNOROLLDRUM)
End Sub

Private Sub cmdPickUp_Click()
    SendData (cPICKUP)
End Sub

Private Sub cmdPullTrigger_Click()
    gstrDrum = Shoot(gstrDrum)
    SendData (cPULLTRIGGER)
End Sub

Private Sub Form_Load()
    
    
    gstrDrum = Rotate("100000") 'initiate the drum of the gun:
                                '1 Bullet
                                '5 free slots

    frmMain.flshSound.Movie = App.Path & cFlashSound
    frmMain.flshAnimation.Movie = App.Path & cFlashGame
    PlayFlashFrom (cFlashSTARTUP)
    
    
    ' for the stats :>
    frmMain.lbMaxTriggers = CLng("0" & GetOption("Stats", "Maxtriggers"))
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call cmdEnd_Click
End Sub




Private Sub Form_Resize()

        On Error Resume Next ' if the frame_chat gets to small to work with.
        With frmMain
            'Resizing the Frames.
            .flshAnimation.Width = .Width
            .flshAnimation.Height = 25 * .flshAnimation.Width / 70

            .Frame_Chat.Top = .flshAnimation.Height
            .Frame_Chat.Width = .Width - .Frame_Chat.Left - 2 * frmMain.Frame_Settings.Left
            .Frame_Chat.Height = .Height - .flshAnimation.Height - 5 * frmMain.Frame_Settings.Left
            
            .Frame_Settings.Top = .flshAnimation.Height
            .Frame_Gun.Top = .flshAnimation.Height
            .Frame_Stats.Top = .Frame_Settings.Top + .Frame_Settings.Height + .Frame_Settings.Left
            
            .lstChat.Height = .Frame_Chat.Height - 1.5 * .lstChat.Top - .txtChatLine.Height
            .lstChat.Width = .Frame_Chat.Width - 2 * .lstChat.Left
            .txtChatLine.Top = .lstChat.Height + .lstChat.Top
            .txtChatLine.Width = .Frame_Chat.Width - 2 * .lstChat.Left
            
        End With
        On Error GoTo 0
End Sub


Private Sub Form_Terminate()
    End
End Sub

Private Sub tmrDrum_Timer()
    ' The Stats-Timer .. nothing Special.
    If gboolConnected Then
        lblStats.Caption = "You: " & gintLocalPoints & " - " & gstrRemoteNick & ": " & gintRemotePoints & ""
        glngTime = glngTime + 3 ' 3? Cause it fires all 300ms
    Else
        glngTime = 0
        lblStats.Caption = "You: " & gintLocalPoints & " ..."
    End If
    
    Call DoStatistics
End Sub


Private Sub txtChatLine_KeyPress(KeyAscii As Integer)
Dim strSenddata As String
    If KeyAscii = 13 Then 'enter
        If Trim(txtChatLine.Text) = "" Then
        Else
            strSenddata = "" & gstrLocalNick & ": " & txtChatLine.Text
            If gboolConnected Then
                WinSock.SendData strSenddata
                DoEvents
            Else
                strSenddata = strSenddata & " (offline)"
            End If
            
            SetStatus (strSenddata)
            txtChatLine.Text = ""
        End If
        KeyAscii = 0 'please, dont pleep
    End If

End Sub


Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    Dim strDataRecived As String

    On Error Resume Next
        frmMain.WinSock.GetData strDataRecived
    If Err > 0 Then
        MsgBox Err.Description
    End If
    On Error GoTo 0
    DoEvents

    
    If IsDataString(strDataRecived) Then
         If gboolShowData Then
            SetStatus strDataRecived
         End If
         ProcessDataString (strDataRecived)
    Else
         SetStatus strDataRecived
    End If
    
End Sub


Private Sub Winsock_Connect()
    SendData (cWELCOME)
End Sub


Private Sub winsock_ConnectionRequest(ByVal requestID As Long)
    If WinSock.State <> sckClosed Then WinSock.Close
    WinSock.Accept requestID
End Sub

Private Sub Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "There is no Server running or another Winsock-Error occured." & vbCrLf & vbCrLf & Description
    Load frmStart
    frmStart.Show
    frmMain.Hide
    'if you want to kill the app.
    Unload frmMain
End Sub


