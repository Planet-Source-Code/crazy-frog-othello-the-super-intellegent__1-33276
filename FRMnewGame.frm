VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMnewGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Game"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BTcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton BTstart 
      Caption         =   "Start Game"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   3120
      Width           =   1455
   End
   Begin VB.PictureBox picLogo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2415
      ScaleWidth      =   255
      TabIndex        =   7
      Top             =   0
      Width           =   255
   End
   Begin VB.Frame FrameColor 
      Caption         =   "Your Color"
      Height          =   1575
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   1935
      Begin VB.OptionButton OPyourColor 
         Caption         =   "White"
         Height          =   435
         Index           =   1
         Left            =   720
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton OPyourColor 
         Caption         =   "Blank"
         Height          =   435
         Index           =   0
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   120
         Shape           =   3  'Circle
         Top             =   360
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   120
         Shape           =   3  'Circle
         Top             =   960
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Against"
      Height          =   1815
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton OPcontre 
         Caption         =   "External Program"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton OPcontre 
         Caption         =   "Computer"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton OPcontre 
         Caption         =   "IP Number"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton OPcontre 
         Caption         =   "Man nearby you"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame FrameStyleJeu 
      Caption         =   "Man Vs Man"
      Height          =   3495
      Index           =   0
      Left            =   2520
      TabIndex        =   8
      Top             =   120
      Width           =   2655
   End
   Begin VB.Frame FrameStyleJeu 
      Caption         =   "You Vs IP"
      Height          =   3495
      Index           =   1
      Left            =   2520
      TabIndex        =   10
      Top             =   120
      Width           =   2655
      Begin VB.Frame FrameServer 
         Height          =   1695
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Visible         =   0   'False
         Width           =   2415
         Begin VB.TextBox TXTlocalIP 
            Height          =   315
            Left            =   120
            TabIndex        =   24
            Top             =   960
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Your IP is :"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label LBLwait 
            Caption         =   "WAITING ..."
            Height          =   615
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.OptionButton OPclient 
         Caption         =   "Connect to IP"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton OPserver 
         Caption         =   "Waiting for opponent"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   2055
      End
      Begin VB.Frame FrameClient 
         Height          =   1695
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Visible         =   0   'False
         Width           =   2415
         Begin VB.CommandButton BTconnect 
            Caption         =   "Connect"
            Enabled         =   0   'False
            Height          =   375
            Left            =   1560
            TabIndex        =   17
            Top             =   240
            Width           =   795
         End
         Begin VB.TextBox TXTip 
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label LBLrequest 
            Caption         =   "REQUEST IN PROGRESS ..."
            Height          =   375
            Left            =   960
            TabIndex        =   23
            Top             =   1200
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "IP Number :"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   975
         End
      End
   End
   Begin VB.Frame FrameStyleJeu 
      Caption         =   "You Vs External Program"
      Height          =   3495
      Index           =   3
      Left            =   2520
      TabIndex        =   14
      Top             =   120
      Width           =   2655
   End
   Begin VB.Frame FrameStyleJeu 
      Caption         =   "You Vs Computer"
      Height          =   3495
      Index           =   2
      Left            =   2520
      TabIndex        =   12
      Top             =   120
      Width           =   2655
      Begin MSComctlLib.Slider SLDlevel 
         Height          =   2415
         Left            =   1800
         TabIndex        =   26
         Top             =   360
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   4260
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   1
         Max             =   2
         TickStyle       =   1
      End
      Begin VB.Label Label3 
         Caption         =   "LEVEL :"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   1440
         Width           =   615
      End
   End
End
Attribute VB_Name = "FRMnewGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub BTcancel_Click()
    Unload Me
End Sub

Private Sub BTconnect_Click()
On Error GoTo gestion_erreur
    FRMplateau.TCP.RemoteHost = CStr(TXTip.Text)
    FRMplateau.TCP.RemotePort = 1001
    FRMplateau.TCP.Connect
    LBLrequest.Visible = True
    Exit Sub
gestion_erreur:
    MsgBox Err.Description, vbCritical, "ERREUR n°" & CStr(Err.Number)
End Sub

Private Sub BTreject_Click()

End Sub

Private Sub BTstart_Click()
    Select Case ModeDeJeu
        Case 0
            FRMplateau.DebutePartie
            Unload Me
        Case 1
            If RESEAU Then
                If OPyourColor(0) Then
                    MoiIs Player(1)
                    MonAdversaireIs Player(2)
                    FRMplateau.TCP.SendData "NEWGAME:MOI"
                Else
                    MoiIs Player(2)
                    MonAdversaireIs Player(1)
                    FRMplateau.TCP.SendData "NEWGAME:YOU"
                    FRMplateau.PICplateau.Enabled = False
                End If
                Me.Visible = False
            Else
                MsgBox "NO CONNECTION YET !"
            End If
        Case 2
            Options.Level = SLDlevel.Value
            If OPyourColor(0) Then
                MoiIs Player(1)
                MonAdversaireIs Player(2)
            Else
                MoiIs Player(2)
                MonAdversaireIs Player(1)
                FRMplateau.TimerDessine.Tag = "ORDIPLAY"
            End If
            FRMplateau.DebutePartie
            Unload Me
        Case 3
    End Select

    
End Sub

Private Sub CHKclient_Click()
    
End Sub



Private Sub Form_Click()
    MsgBox Me.Width
End Sub









Private Sub OPclient_Click()
    If OPclient Then
            
        If FRMplateau.TCP.State <> sckClosed Then FRMplateau.TCP.Close
        RESEAU = False
        
        FrameServer.Visible = False
        FrameClient.Visible = True
        FrameColor.Visible = False
        BTstart.Enabled = False
    End If
End Sub

Private Sub OPcontre_Click(Index As Integer)
    FrameStyleJeu(0).Visible = False
    FrameStyleJeu(1).Visible = False
    FrameStyleJeu(2).Visible = False
    FrameStyleJeu(3).Visible = False
    FrameStyleJeu(Index).Visible = True
    ModeDeJeu = Index
End Sub

Private Sub OPserver_Click()
On Error GoTo gestion_erreur
    If OPserver Then
        If FRMplateau.TCP.State <> sckClosed Then FRMplateau.TCP.Close
        FRMplateau.TCP.LocalPort = 1001
        FRMplateau.TCP.Listen
        FrameServer.Visible = True
        FrameColor.Visible = True
        
        FrameClient.Visible = False
        
        RESEAU = False
        BTstart.Enabled = True
        TXTlocalIP = FRMplateau.TCP.LocalIP
    End If
    Exit Sub
gestion_erreur:
    MsgBox Err.Description, vbCritical, "ERREUR n°" & CStr(Err.Number)
End Sub


'==================================================================================
'
'==================================================================================
Private Sub TXTip_Change()
    BTconnect.Enabled = True
End Sub
