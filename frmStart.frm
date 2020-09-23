VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form frmStart 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Gimp Roulette Startup"
   ClientHeight    =   3570
   ClientLeft      =   3330
   ClientTop       =   1350
   ClientWidth     =   3495
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   3495
   Begin VB.PictureBox pbPercentBar 
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
      ScaleWidth      =   161.25
      TabIndex        =   9
      Top             =   3240
      Width           =   3255
      Begin VB.Label lblPercent 
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
         Left            =   1440
         TabIndex        =   12
         Top             =   0
         Width           =   225
      End
      Begin VB.Label lbWon 
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
         Left            =   2760
         TabIndex        =   11
         Top             =   0
         Width           =   360
      End
      Begin VB.Label lbLost 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lost"
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
         TabIndex        =   10
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Your Connection"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   3255
      Begin VB.CheckBox chkServerMode 
         Caption         =   "Server?"
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   240
         Value           =   1  'Aktiviert
         Width           =   855
      End
      Begin VB.CommandButton cmdStartServer 
         Caption         =   "Start a new Server"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtIPAddress 
         Appearance      =   0  '2D
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         MaxLength       =   20
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Enter IP and Connect"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "or "
         Height          =   255
         Left            =   4440
         TabIndex        =   4
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Your Nickname"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   3255
      Begin VB.TextBox txtNick 
         Appearance      =   0  '2D
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         MaxLength       =   15
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Timer tmrNick 
      Interval        =   1
      Left            =   5760
      Top             =   3120
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash flshStart 
      Height          =   1250
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3500
      _cx             =   6174
      _cy             =   2205
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
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub chkServerMode_Click()
    ' I've used this, cause i don't wanted to use
    ' the common-control-ocx - which some don't have.
    
    If chkServerMode.Value = 0 Then
        txtIPAddress.Visible = True
        cmdConnect.Visible = True
        cmdStartServer.Visible = False
        
    Else
        txtIPAddress.Visible = False
        cmdConnect.Visible = False
        cmdStartServer.Visible = True
    End If
End Sub

Private Sub cmdConnect_Click()
    
    With frmMain.WinSock           ' initiate Winsock control.
            .Close
            .Connect txtIPAddress.Text, DataTCPPort
            
            '.RemotePort = DataUDPLocalPort
            '.Bind DataUDPRemotePort
            '.RemoteHost = txtIPAddress.Text
            '.LocalIP
    End With
    
    DoEvents
    gboolIsServer = False
    Call GotoGame
    
    Call SetTitle("connecting " & txtIPAddress.Text)
    
    ' Send cWelcome, so the Server knows our IP.
    
End Sub

Private Sub cmdStartServer_Click()
        
        With frmMain.WinSock           ' Winsock control.
                
            On Error GoTo Errorhandler
                .Close
                .LocalPort = DataTCPPort
                .Listen
            On Error GoTo 0
        End With
        
        gboolIsServer = True
        Call GotoGame
        Call SetTitle("waiting for Opponent")
    Exit Sub
Errorhandler:
    MsgBox ("Can't run 2 Servers on one PC" & vbCrLf & Err.Description)
End Sub


Private Sub Form_Load()
    txtNick.Text = GetOption("Settings", "Nick")
    txtIPAddress.Text = GetOption("Settings", "LastIP")
    
    'Initiate Game:
    gboolConnected = False
    tmrNick.Enabled = True

    Call genPercentBar(pbPercentBar, lblPercent, CLng("0" & GetOption("Stats", "Won")) + CLng("0" & GetOption("Stats", "Lost")), CLng("0" & GetOption("Stats", "Lost")))

    flshStart.Movie = App.Path & cFlashGame
    flshStart.GotoFrame 1
    
    

End Sub



Private Sub tmrNick_Timer()
' checks, if enough Information was entered, to join the game.

    If Trim(frmStart.txtNick.Text) = "" Then
        cmdConnect.Enabled = False
        cmdStartServer.Enabled = False
    
    
    Else

        cmdStartServer.Enabled = True
    
        If Trim(txtIPAddress.Text) = "" Then
            cmdConnect.Enabled = False
        Else
            cmdConnect.Enabled = True
        End If
    End If
End Sub


