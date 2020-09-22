VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FRMoption 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Option"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   Icon            =   "FRMoption.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3360
   ScaleWidth      =   7545
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7560
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Hide in Task Bar"
      Height          =   1215
      Left            =   4320
      TabIndex        =   19
      Top             =   2040
      Width           =   3135
      Begin VB.CommandButton BTchoose 
         Caption         =   "..."
         Height          =   285
         Left            =   2640
         TabIndex        =   24
         Top             =   840
         Width           =   375
      End
      Begin VB.CheckBox CHKhide 
         Caption         =   "Enable"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
      Begin MSComctlLib.Slider SLDico 
         Height          =   675
         Left            =   2730
         TabIndex        =   22
         Top             =   150
         Visible         =   0   'False
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   1191
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   1
         Min             =   1
         Max             =   6
         SelStart        =   1
         Value           =   1
      End
      Begin VB.TextBox TXTshell 
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Text            =   "C:\WINDOWS\calc.exe"
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Left Click action :"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1455
      End
      Begin VB.Image IMico 
         Height          =   480
         Left            =   2160
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox PICplateau 
      AutoSize        =   -1  'True
      Height          =   855
      Left            =   2520
      Picture         =   "FRMoption.frx":030A
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   17
      Tag             =   "1"
      Top             =   360
      Width           =   855
   End
   Begin MSComctlLib.ImageList ILpreview 
      Left            =   7800
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   93
      ImageHeight     =   93
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMoption.frx":246C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMoption.frx":38D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMoption.frx":4D99
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrameCoord 
      Caption         =   "Coordinates"
      Height          =   1815
      Left            =   4320
      TabIndex        =   13
      Top             =   120
      Width           =   3135
      Begin VB.OptionButton OPcoord 
         Caption         =   "None"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Tag             =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OPcoord 
         Caption         =   "Partial"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Tag             =   "1"
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton OPcoord 
         Caption         =   "Complete"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Tag             =   "16"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Image IMpreview 
         Height          =   1440
         Left            =   1560
         Top             =   240
         Width           =   1440
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Anim"
      Height          =   1215
      Left            =   480
      TabIndex        =   8
      Top             =   2040
      Width           =   3615
      Begin VB.CommandButton BTtest 
         Caption         =   "Test"
         Height          =   615
         Left            =   2880
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin MSComctlLib.Slider SLDturnSpeed 
         Height          =   375
         Left            =   1080
         TabIndex        =   9
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin MSComctlLib.Slider SLDdelay 
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
      End
      Begin VB.Label Label3 
         Caption         =   "Delay"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Turn Speed"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame FrameInform 
      Caption         =   "Inform"
      Height          =   1815
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   3615
      Begin VB.CheckBox CHKinformStyle 
         Caption         =   "With numbers"
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Tag             =   "14"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSComctlLib.Slider SLDinformSize 
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   1200
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin VB.OptionButton OPinform 
         Caption         =   "Complete"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Tag             =   "16"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton OPinform 
         Caption         =   "Partial"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Tag             =   "1"
         Top             =   840
         Width           =   855
      End
      Begin VB.OptionButton OPinform 
         Caption         =   "None"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Tag             =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label LBLsize 
         Caption         =   "Size:"
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   375
      End
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
      Height          =   3360
      Left            =   0
      Picture         =   "FRMoption.frx":6330
      ScaleHeight     =   3360
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "FRMoption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim FirstTime As Boolean

Private Plateau As cBitmap





Private Sub BTchoose_Click()
    TXTshell = OuvrirFichierExistant("EXE *")
End Sub
'==================================================================================
'   AFFICHE LA BOITE D OUVERTURE DE FICHIER
'   RENVOIE LE NOM AINSI QUE LE CHEMIN DU FICHIER CHOISI
'   EN CAS D ANNULATION RENVOIE ""
'   PARAMETRES OPTIONELS :
'       LES FILTRES : *.* PAR DEFAUT, A PASSER AINSI : "TXT HTML BAT"
'       LE NOM DU FICHIER SANS LE CHEMIN
'       LE TITRE DE LA BOITE D OUVERTURE
'       NUMERO DE L ERREUR QUI C EST PRODUIT
'==================================================================================
Public Function OuvrirFichierExistant(Optional ByVal TypeFile As String, _
                                    Optional File As String, _
                                    Optional ByVal DialogTitle As String, _
                                    Optional ErrNumber As Long) As String
Dim Tempo As String
Dim Extention As String
Dim BoiteOuverture As CommonDialog

Set BoiteOuverture = FRMoption.CommonDialog1

On Error GoTo ErrHandler
    BoiteOuverture.Flags = cdlOFNHideReadOnly
    BoiteOuverture.Filter = ""
    
    If DialogTitle <> "" Then BoiteOuverture.DialogTitle = DialogTitle
    
    If TypeFile = "" Then
        BoiteOuverture.Filter = "Tous les fichiers (*.*)|*.*"
    Else
        If InStr(1, TypeFile, " ", vbTextCompare) = 0 Then
            BoiteOuverture.Filter = "Fichiers " + UCase(TypeFile) + "|*." + UCase(TypeFile)
        Else
            Tempo = UCase(TypeFile)
            While InStr(1, Tempo, " ", vbTextCompare) <> 0
                Extention = Mid(Tempo, 1, InStr(1, Tempo, " ", vbTextCompare) - 1)
                BoiteOuverture.Filter = BoiteOuverture.Filter & "|Fichiers " + Extention + "|*." + Extention
                Tempo = Replace(Tempo, Extention + " ", "")
            Wend
            BoiteOuverture.Filter = BoiteOuverture.Filter & "|Fichiers " + Tempo + "|*." + Tempo
            BoiteOuverture.Filter = Mid(BoiteOuverture.Filter, 2)
        End If
    End If
    
    BoiteOuverture.FilterIndex = 1
    BoiteOuverture.ShowOpen
    OuvrirFichierExistant = BoiteOuverture.FileName
    File = BoiteOuverture.FileTitle
    ErrNumber = 0
    Exit Function
ErrHandler:
    OuvrirFichierExistant = ""
    ErrNumber = Err.Number
End Function


Private Sub BTtest_Click()
    PICplateau_Click
End Sub

Private Sub CHKhide_Click()
    If CHKhide.Value = vbChecked Then
        
        SLDico.Visible = True
        IMico.Visible = True
        
        IMico.Picture = MDImain.IL_ICO.ListImages.Item(SLDico.Value).Picture
        IMico.DragIcon = IMico.Picture
        
        myNID.hIcon = IMico.DragIcon
   
        ShellNotifyIcon NIM_MODIFY, myNID
    Else
        SLDico.Visible = False
        IMico.Visible = False
        
    End If
End Sub

'==================================================================================
'
'==================================================================================
Private Sub CHKinformStyle_Click()
    If CHKinformStyle.Value = vbChecked Then
        Options.StylePrevien = 30
    Else
        Options.StylePrevien = 16
        
    End If
    If CHKinformStyle.Value = vbChecked Then
                PionBMP(100).Cell = 40
            Else
                PionBMP(100).Cell = 26
            End If
            Actualise
End Sub

Private Sub Form_Activate()
    PionBMP(100).Active = True
    'If MDImain.Toolbar.Buttons(8).Value = tbrUnpressed Then Me.Visible = False
End Sub

'==================================================================================
'
'==================================================================================
Private Sub Form_Load()
    FirstTime = True
         
      IMico.DragIcon = MDImain.IL_ICO.ListImages.Item(1).Picture

    SLDturnSpeed.Value = Options.SpeedAnim \ 10
    SLDdelay.Value = Options.SpeedAction \ 100
    
    Set Plateau = New cBitmap
    Plateau.CreateFromPicture PICplateau.Picture
    
    OPcoord(Options.Coordonnee).Value = True
    
    Select Case Options.StylePrevien
        Case 0
            OPinform(0).Value = True
        Case 16
            OPinform(2).Value = True
            CHKinformStyle.Value = 0
        Case 30
            OPinform(2).Value = True
            CHKinformStyle.Value = vbChecked
        Case Else
            OPinform(1).Value = True
            SLDinformSize.Value = Options.StylePrevien - 16
    End Select
    
End Sub


'==================================================================================
'
'==================================================================================
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If BIGcancel = 0 Then

    Else
        Cancel = 1
        Me.Visible = False
        MDImain.Toolbar.Buttons(8).Value = 0
    End If
End Sub






Private Sub OPcoord_Click(Index As Integer)
    Options.Coordonnee = Index
    IMpreview.Picture = ILpreview.ListImages.Item(Index + 1).Picture
    FRMplateau.DessineCoordonne
End Sub

'==================================================================================
'
'==================================================================================
Private Sub OPinform_Click(Index As Integer)
    Options.StylePrevien = CLng(OPinform(Index).Tag)
    SLDinformSize.Visible = False
    LBLsize.Visible = False
    CHKinformStyle.Visible = False
    Select Case Index
        Case 0
            PionBMP(100).Cell = 16
            Actualise
        Case 1
            SLDinformSize.Visible = True
            LBLsize.Visible = True
            PionBMP(100).Cell = 16 + SLDinformSize.Value
            Actualise
        Case 2
            CHKinformStyle.Visible = True
            If CHKinformStyle.Value = vbChecked Then
                PionBMP(100).Cell = 40
            Else
                PionBMP(100).Cell = 26
            End If
            Actualise
    End Select
End Sub

Private Sub PICplateau_Click()
    Anime
End Sub

'==================================================================================
'
'==================================================================================
Private Sub SLDdelay_Change()
    
End Sub

Private Sub SLDdelay_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Options.SpeedAction = SLDdelay.Value * 100
End Sub

Private Sub SLDico_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 0 Then
        IMico.Picture = MDImain.IL_ICO.ListImages.Item(SLDico.Value).Picture
        IMico.DragIcon = IMico.Picture
        
        myNID.hIcon = IMico.DragIcon
   
        ShellNotifyIcon NIM_MODIFY, myNID
    End If
End Sub

'==================================================================================
'
'==================================================================================
Private Sub SLDinformSize_Change()
    
End Sub

Private Sub SLDinformSize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 0 Then
        Options.StylePrevien = SLDinformSize.Value
        PionBMP(100).Cell = 16 + SLDinformSize.Value
        Actualise
    End If
End Sub


'==================================================================================
'   dessine sur le plateau les pions actif uniquement (pour l'animation)
'==================================================================================
Private Sub Actualise()
Dim lHDC As Long
Dim j As Long
    lHDC = PICplateau.hDC
    Plateau.RenderBitmap lHDC, 0, 0
    For j = 0 To 100
        If PionBMP(j).Active Then
            PionBMP(100).RestoreBackground Plateau.hDC
            PionBMP(100).StoreBackground Plateau.hDC, 10, 10
            PionBMP(100).TransparentDraw Plateau.hDC, 10, 10, PionBMP(100).Cell, True
            PionBMP(100).StageToScreen lHDC, Plateau.hDC
        End If
    Next j
End Sub

'==================================================================================
'
'==================================================================================
Private Sub Anime()
Dim j As Long
Dim i As Long
Dim AttenteTurn As Long
Dim AttentePose As Long

    AttenteTurn = Options.SpeedAnim + 1
    AttentePose = Options.SpeedAction
    
    If PICplateau.Tag = "1" Then
        PICplateau.Tag = "15"
        PionBMP(100).Cell = CLng(PICplateau.Tag)
        Actualise
        Sleep AttentePose
        For j = Player(2).Sprite To Player(1).Sprite Step -1
            For i = 0 To 100
                If PionBMP(i).Active Then
                    PionBMP(i).Cell = j
                End If
            Next i
            Sleep AttenteTurn
            DoEvents
            Actualise
        Next j
    Else
        PICplateau.Tag = "1"
        PionBMP(100).Cell = CLng(PICplateau.Tag)
        Actualise
        Sleep AttentePose
        For j = Player(1).Sprite To Player(2).Sprite Step 1
            For i = 0 To 100
                If PionBMP(i).Active Then
                    PionBMP(i).Cell = j
                End If
            Next i
            Sleep AttenteTurn
            DoEvents
            Actualise
        Next j
    End If
End Sub

Private Sub SLDturnSpeed_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Options.SpeedAnim = (SLDturnSpeed.Max - SLDturnSpeed.Value) * 10
End Sub

