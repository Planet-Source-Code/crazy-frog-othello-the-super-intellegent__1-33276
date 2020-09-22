VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FRMplateau 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   Icon            =   "FRMjeu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   453
   Begin VB.PictureBox PICplateau 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      FillColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5565
      Left            =   255
      ScaleHeight     =   371
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   371
      TabIndex        =   0
      Top             =   0
      Width           =   5565
      Begin VB.CommandButton Command1 
         Caption         =   "About"
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   5160
         Width           =   1095
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
      Height          =   5445
      Left            =   0
      Picture         =   "FRMjeu.frx":030A
      ScaleHeight     =   5445
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Timer TimerDessine 
      Interval        =   1000
      Left            =   6120
      Top             =   240
   End
   Begin MSWinsockLib.Winsock TCP 
      Left            =   6240
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "FRMplateau"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Dim LesCasesAChanger As Collection

Private Type Point
    x As Long
    y As Long
End Type

Dim Direction(7) As Long

Dim CaseJouer As Long
Dim Liberte As Long

Dim LastCoup As Point



Private Type UneTable
    Pion(99) As Long
End Type


Dim LaTable As UneTable
Dim TableRecherche As UneTable
Dim TableTempo As UneTable

Private Plateau As cBitmap


Dim Efface As Long


Private Sub Command1_Click()
MsgBox "Othello Ver-1.00"
MsgBox "A game by Samar Pathania"
MsgBox "Email:-Samar2000in@yahoo.com"
frmAbout.Show vbModal

End Sub

'==================================================================================
'
'==================================================================================
Private Sub Form_Activate()
    DessineTout
    If MDImain.Toolbar.Buttons(4).Value = 0 Then
        FRMgraphic.Visible = False
    Else
        FRMgraphic.Dessine
    End If
    picLogo.Height = Me.ScaleHeight * 15
   
End Sub

'==================================================================================
'
'==================================================================================
Private Sub InitVariables()

    Direction(0) = -11
    Direction(1) = -10
    Direction(2) = -9
    Direction(3) = -1
    Direction(4) = 1
    Direction(5) = 9
    Direction(6) = 10
    Direction(7) = 11
    
    Player(1).Value = 1
    Player(2).Value = -1
    
    Player(1).Number = 1
    Player(2).Number = 2
    
    Player(1).Name = "Gris Foncé"
    Player(2).Name = "Gris Clair"
    
    ModeDeJeu = 0
    
    Options.Level = 1
    
    Efface = 16
    
    Player(1).Sprite = 2
    Player(2).Sprite = 14
    
    
    TailleCase = 37
    TaillePionBMP = 32
    Marge = (TailleCase - TaillePionBMP) \ 2 + 1
    
    Options.SpeedAction = 500
    Options.SpeedAnim = 40
    Options.StylePrevien = 16
    
    
End Sub



'==================================================================================
'
'==================================================================================
Private Sub Form_Load()
Dim i As Long
Dim j As Long

    
    

    InitVariables
    
    FRMplateau.Width = TailleCase * 10
    FRMplateau.Height = TailleCase * 10
    
    PICplateau.Width = TailleCase * 10
    PICplateau.Height = TailleCase * 10

    While FRMplateau.ScaleWidth < PICplateau.Width + PICplateau.Left
        FRMplateau.Width = FRMplateau.Width + 15
    Wend
     
    While FRMplateau.ScaleHeight < PICplateau.Height + PICplateau.Top * 2
        FRMplateau.Height = FRMplateau.Height + 15
    Wend
    
  
    PICplateau.Picture = LoadPicture(App.Path & "\plateau32.gif")

    CreateSpriteResource LesPionBMPs, App.Path & "\sprites32.bmp", 15, 4, RGB(0, 160, 0)
    
    For i = 0 To 100
        CreateSprite LesPionBMPs, PionBMP(i)
    Next
    Set Plateau = New cBitmap
    
    Plateau.CreateFromPicture PICplateau.Picture
    
    DebutePartie
    picLogo.Height = Me.ScaleHeight * 15
    picLogo.Width = 255
    
    
    Me.Top = 30
    Me.Left = 300
End Sub

'==================================================================================
'
'==================================================================================
Public Sub InitDebutPartie()
Dim i As Long
Dim Ligne As ListItem

    For i = 0 To 99
        PionBMP(i).Active = False
        PionBMP(i).Cell = 16
    Next
    
    DessineTout
    For i = 0 To 99
        PionBMP(i).Cell = 0
    Next
    
    
    VideTable LaTable

    PionBMP(44).Cell = Player(2).Sprite
    LaTable.Pion(44) = Player(2).Value
    
    PionBMP(45).Cell = Player(1).Sprite
    LaTable.Pion(45) = Player(1).Value
    
    PionBMP(54).Cell = Player(1).Sprite
    LaTable.Pion(54) = Player(1).Value
    
    PionBMP(55).Cell = Player(2).Sprite
    LaTable.Pion(55) = Player(2).Value
    
    
    FRMhistorique.LVstory.ListItems.Clear
    Set Ligne = FRMhistorique.LVstory.ListItems.Add(, , "0", , 3)
    Ligne.SubItems(4) = "2"
    Ligne.SubItems(5) = "2"
    Ligne.SubItems(6) = "0"
    Ligne.SubItems(7) = 4
    
    If MDImain.Toolbar.Buttons(6).Value = 0 Then FRMhistorique.Visible = False
    
    Select Case ModeDeJeu
        Case 0
            Aqui = 1
            AdversaireJoue Player(1)
        Case 1
            Aqui = 1
            AdversaireJoue Player(1)
        Case 2
            AdversaireJoue Player(1)
            MoiIs Player(1)
            MonAdversaireIs Player(2)
    End Select
    AncienCoup = 100
End Sub

'==================================================================================
'
'==================================================================================
Public Sub DebutePartie()
    InitDebutPartie
    DessineTout
End Sub


Private Sub Form_Paint()
    DessineTout
    picLogo.Height = Me.ScaleHeight * 15
    
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If FRMoption.CHKhide.Value = vbChecked Then
        MDImain.Visible = False
    Else
        MDImain.WindowState = vbMinimized
    End If
    Cancel = BIGcancel
End Sub

'==================================================================================
'
'==================================================================================
Private Sub PICplateau_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)


Dim PeutJouer As Boolean

    LastCoup.x = Round(x \ TailleCase)
    If LastCoup.x = 0 Or LastCoup.x = 9 Then
        Exit Sub
    End If
    
    LastCoup.y = Round(y \ TailleCase)
    If LastCoup.y = 0 Or LastCoup.y = 9 Then
        Exit Sub
    End If
    
    CaseJouer = LastCoup.y * 10 + LastCoup.x
    
    Select Case ModeDeJeu
        Case 0
            If Button = 1 Then
                PionRetourne = NombrePossible(LaTable, CaseJouer, Player(Aqui).Value)
                If PionRetourne > 0 Then
                    PlayIn CaseJouer, Player(Aqui)
                End If
            End If
        Case 1
            If Button = 1 Then
                PionRetourne = NombrePossible(LaTable, CaseJouer, MOI.Value)
                If PionRetourne > 0 Then
                    PlayIn CaseJouer, MOI
                    FRMplateau.TCP.SendData "PLAY:" + CStr(CaseJouer)
                    PICplateau.Enabled = False
                    AdversaireJoue MONADVERSAIRE
                    DessineTout
                    
                End If
            End If
        Case 2
            If Button = 1 Then
                PionRetourne = NombrePossible(LaTable, CaseJouer, MOI.Value)
                If PionRetourne > 0 Then
                    PlayIn CaseJouer, MOI
                    DessineTout
                    If AdversaireJoue(MONADVERSAIRE) Then
                        aRejouer = False
                        DessineTout
                        Sleep Options.SpeedAction
                        OrdinateurJoue
                        AdversaireJoue MOI
                        DessineTout
                    Else
                        aRejouer = True
                        MsgBox "REPLAY"
                    End If
                End If
            End If
    End Select
    
    
End Sub


'==================================================================================
'
'==================================================================================
Private Function AdversaireJoue(toto As APlayer, Optional CombienCase As Long) As Boolean
Dim i As Long
Dim j As Long
Dim Tempo As Long

    AdversaireJoue = False
    CombienCase = 0
    For j = 1 To 8
        For i = j * 10 + 1 To j * 10 + 8
            If PionBMP(i).Cell > 16 Then PionBMP(i).Cell = 16
                Tempo = NombrePossible(LaTable, i, toto.Value)
                If Tempo > 0 Then
                    If Options.StylePrevien > 0 Then
                        Select Case Options.StylePrevien
                            Case 16
                                PionBMP(i).Cell = 16 + Min(Tempo, 14)
                            Case 30
                                PionBMP(i).Cell = 30 + Min(Tempo, 14)
                            Case Else
                                PionBMP(i).Cell = 16 + Options.StylePrevien
                        End Select
                    End If
                        AdversaireJoue = True
                        CombienCase = CombienCase + 1
                End If
        Next i
    Next j
    
End Function
'==================================================================================
'
'==================================================================================
Private Sub Tourne()
    If Aqui = 1 Then Aqui = 2 Else Aqui = 1
End Sub


'==================================================================================
'   active les case a changer pour l'animation qui va suivre
'   change en meme temps la valeur des pions changer
'==================================================================================
Private Sub ActiveCaseAChanger(ByVal LaCase As Long, Joueur As Long)
Dim i As Long
Dim Sens As Long
Dim Adversaire As Long

    Adversaire = -Joueur

    Set LesCasesAChanger = New Collection
    
    For Sens = 0 To 7
        i = 1
        While LaTable.Pion(LaCase + Direction(Sens) * i) = Adversaire
            LesCasesAChanger.Add LaCase + Direction(Sens) * i
            i = i + 1
        Wend
        If LaTable.Pion(LaCase + Direction(Sens) * i) = Joueur Then
            For i = 1 To LesCasesAChanger.Count
                PionBMP(LesCasesAChanger.Item(1)).Active = True
                LaTable.Pion(LesCasesAChanger.Item(1)) = Joueur
                LesCasesAChanger.Remove 1
            Next i
        Else
            For i = 1 To LesCasesAChanger.Count
                LesCasesAChanger.Remove 1
            Next i
        End If
    Next Sens
    
End Sub

Private Function AutreSprite(Joueur As APlayer) As Long
    If Joueur.Sprite = Player(1).Sprite Then
        AutreSprite = Player(2).Sprite
    Else
        AutreSprite = Player(1).Sprite
    End If
    
End Function
'==================================================================================
'   met en rotation les pions actifs
'==================================================================================
Private Sub Animation(ByVal LaCase As Long, Joueur As APlayer)
Dim i As Long
Dim j As Long
Dim AttenteTurn As Long
Dim AttentePose As Long

    AttentePose = Options.SpeedAction
    AttenteTurn = Options.SpeedAnim
    

    PionBMP(LaCase).Active = True
    PionBMP(LaCase).Cell = Joueur.Sprite - Joueur.Value
    LaTable.Pion(LaCase) = Joueur.Value
    If Not aRejouer Then PionBMP(AncienCoup).Cell = AutreSprite(Joueur)
    
    Actualise
    Sleep AttentePose
    PionBMP(LaCase).Active = False
    ActiveCaseAChanger LaCase, Joueur.Value
    
    If Joueur.Value = 1 Then
        For j = Player(2).Sprite To Player(1).Sprite Step -1
            For i = 0 To 99
                If PionBMP(i).Active Then
                    PionBMP(i).Cell = j
                End If
            Next i
            Sleep AttenteTurn
            DoEvents
            Actualise
        Next j
    Else
        For j = Player(1).Sprite To Player(2).Sprite Step 1
            For i = 0 To 99
                If PionBMP(i).Active Then
                    PionBMP(i).Cell = j
                End If
            Next i
            Sleep AttenteTurn
            DoEvents
            Actualise
        Next j
    End If
    
    For i = 0 To 99
        PionBMP(i).Active = False
    Next


End Sub
'==================================================================================
'
'==================================================================================
Private Sub PlayIn(ByVal LaCase As Long, Joueur As APlayer)
    
    Select Case ModeDeJeu
        Case 0
            Animation LaCase, Joueur
            Tourne
            If Not AdversaireJoue(Player(Aqui), Liberte) Then
                Tourne
                MsgBox Player(Aqui).Name & " rejoue"
                AdversaireJoue Player(Aqui)
            End If
        Case 1
            Animation LaCase, Joueur
            If Joueur.Number = MOI.Number Then
                If Not AdversaireJoue(MONADVERSAIRE, Liberte) Then
                    'MsgBox MOI.Name & " rejoue"
                Else
                    PICplateau.Enabled = True
                End If
            Else
                If Not AdversaireJoue(MOI, Liberte) Then
                    'MsgBox MONADVERSAIRE.Name & " rejoue"
                Else
                    PICplateau.Enabled = True
                End If
            End If
        Case 2
            Animation LaCase, Joueur
    End Select
    
    Histoire LaCase, Joueur
    FRMgraphic.Dessine
    DessineTout
    AncienCoup = LaCase
End Sub
'==================================================================================
'
'==================================================================================
Private Sub Histoire(ByVal LaCase As Long, Joueur As APlayer)
Dim Ligne As ListItem
Dim NbNoir As Long
Dim NbBlanc As Long
    Set Ligne = FRMhistorique.LVstory.ListItems.Add(, , FRMhistorique.LVstory.ListItems.Count, , Joueur.Number)
    Ligne.SubItems(Joueur.Number) = LaCase
    Ligne.SubItems(6) = EvaluePositionSimple(LaTable, NbNoir, NbBlanc)
    Ligne.SubItems(5) = NbBlanc
    Ligne.SubItems(4) = NbNoir
    Ligne.SubItems(3) = PionRetourne
    Ligne.SubItems(7) = Liberte
    FRMhistorique.LVstory.SelectedItem = FRMhistorique.LVstory.ListItems(FRMhistorique.LVstory.ListItems.Count)
    
End Sub
'==================================================================================
'   dessine sur le plateau les pions actif uniquement (pour l'animation)
'==================================================================================
Private Sub Actualise()
Dim lHDC As Long
Dim i As Long
    lHDC = PICplateau.hDC
    Plateau.RenderBitmap lHDC, 0, 0
    For i = 0 To 99
        If PionBMP(i).Active Then
            PionBMP(i).RestoreBackground Plateau.hDC
            PionBMP(i).StoreBackground Plateau.hDC, (i Mod 10) * TailleCase + Marge, (i \ 10) * TailleCase + Marge
            PionBMP(i).TransparentDraw Plateau.hDC, (i Mod 10) * TailleCase + Marge, (i \ 10) * TailleCase + Marge, PionBMP(i).Cell, True
            PionBMP(i).StageToScreen lHDC, Plateau.hDC
        End If
    Next i
End Sub
'==================================================================================
'   redissine sur le plateau tout les pions
'==================================================================================
Private Sub DessineTout()
Dim lHDC As Long
Dim i As Long
    lHDC = PICplateau.hDC
    Plateau.RenderBitmap lHDC, 0, 0
    For i = 0 To 99
        If PionBMP(i).Cell <> 0 Then
            PionBMP(i).RestoreBackground Plateau.hDC
            PionBMP(i).StoreBackground Plateau.hDC, (i Mod 10) * TailleCase + Marge, (i \ 10) * TailleCase + Marge
            PionBMP(i).TransparentDraw Plateau.hDC, (i Mod 10) * TailleCase + Marge, (i \ 10) * TailleCase + Marge, PionBMP(i).Cell, True
            PionBMP(i).StageToScreen lHDC, Plateau.hDC
            DoEvents
        End If
    Next i
End Sub

'==================================================================================
'
'==================================================================================

Public Sub DessineCoordonne()
Dim i As Long
    Select Case Options.Coordonnee
        Case 0
            For i = 1 To 8
                PionBMP(i).Cell = 16
                PionBMP(90 + i).Cell = 16
                PionBMP(i * 10).Cell = 16
                PionBMP(i * 10 + 9).Cell = 16
            Next i
        Case 1
            For i = 1 To 8
                PionBMP(i).Cell = 44 + i
                PionBMP(i * 10).Cell = 52 + i
                PionBMP(90 + i).Cell = 16
                PionBMP(i * 10 + 9).Cell = 16
            Next i
        Case 2
            For i = 1 To 8
                PionBMP(i).Cell = 44 + i
                PionBMP(i * 10).Cell = 52 + i
                PionBMP(90 + i).Cell = 44 + i
                PionBMP(i * 10 + 9).Cell = 52 + i
            Next i
        End Select
        DessineTout
        For i = 1 To 8
            PionBMP(i).Cell = 0
            PionBMP(i * 10).Cell = 0
            PionBMP(90 + i).Cell = 0
            PionBMP(i * 10 + 9).Cell = 0
        Next i
        
End Sub
'==================================================================================
'
'==================================================================================
Private Sub CreateSpriteResource( _
        ByRef cR As cSpriteBitmaps, _
        ByVal sFile As String, _
        ByVal cX As Long, _
        ByVal cY As Long, _
        ByVal lTransColor As Long _
    )
    Set cR = New cSpriteBitmaps
    cR.CreateFromFile sFile, cX, cY, , lTransColor
End Sub
'==================================================================================
'
'==================================================================================
Private Sub CreateSprite( _
        ByRef cR As cSpriteBitmaps, _
        ByRef cS As cSprite _
    )
    Set cS = New cSprite
    cS.SpriteData = cR
    cS.Create Me.hDC
End Sub

'==================================================================================
'   renvoie le nombre de pion changer si on joue sur LaCase
'   ET CHANGE CES CASES
'==================================================================================
Private Function ChangeCase(ByVal LaCase As Long, Joueur As APlayer) As Long
Dim i As Long
Dim j As Long
Dim Sens As Long
Dim Tempo As Long
Dim CaseAChanger(8) As Long
Dim Adversaire As Long

    Adversaire = -Joueur.Value

    ChangeCase = 0
    If TableRecherche.Pion(LaCase) = 0 Then
    
    For Sens = 0 To 7
        i = 1
        Tempo = LaCase + Direction(Sens)
        
        If TableRecherche.Pion(Tempo) = Adversaire Then
        
            CaseAChanger(1) = Tempo
            Tempo = LaCase + Direction(Sens) * 2
            
            While TableRecherche.Pion(Tempo) = Adversaire
                i = i + 1
                CaseAChanger(i) = Tempo
                Tempo = LaCase + Direction(Sens) * i
            Wend
            
            If TableRecherche.Pion(Tempo) = Joueur.Value Then
                For j = 1 To i
                    TableRecherche.Pion(CaseAChanger(j)) = Joueur.Value
                Next j
                ChangeCase = ChangeCase + i
            End If
            
        End If
        
    Next Sens
    End If
    
End Function

'==================================================================================
'   renvoie le nombre de pion changer si on joue sur LaCase
'==================================================================================
Private Function NombrePossible(TableR As UneTable, ByVal LaCase As Long, Joueur As Long) As Long
Dim i As Long
Dim j As Long
Dim Sens As Long
Dim Tempo As Long
Dim Adversaire As Long

    Adversaire = -Joueur

    NombrePossible = 0
    If TableR.Pion(LaCase) = 0 Then
    
    For Sens = 0 To 7
        i = 1
        Tempo = LaCase + Direction(Sens)
        
        If TableR.Pion(Tempo) = Adversaire Then
        
            Tempo = LaCase + Direction(Sens) * 2
            
            While TableR.Pion(Tempo) = Adversaire
                i = i + 1
                Tempo = LaCase + Direction(Sens) * (i + 1)
            Wend
            
            If TableR.Pion(Tempo) = Joueur Then
                NombrePossible = NombrePossible + i
            End If
            
        End If
        
    Next Sens
    End If
    
End Function


'==================================================================================
'
'==================================================================================
Private Function EvaluePositionSimple(QuelTable As UneTable, Optional Noir As Long, Optional Blanc As Long) As Long
Dim Tempo As Long
Dim i As Long
Dim j As Long
    Noir = 0
    Blanc = 0
    For j = 1 To 8
        For i = 10 * j + 1 To 10 * j + 8
            If QuelTable.Pion(i) = 1 Then
                Noir = Noir + 1
            End If
            If QuelTable.Pion(i) = -1 Then
                Blanc = Blanc + 1
            End If
        Next i
    Next j
    EvaluePositionSimple = Noir - Blanc
End Function


'==================================================================================
'
'==================================================================================
Private Sub OrdinateurJoue()
Dim Tempo As Long
Dim ValPos As Single
Dim MaxActu As Single
Dim MeilleurCoup As Long
Dim i As Long
Dim j As Long
    MaxActu = -65
    Select Case Options.Level
        Case 0
            For j = 1 To 8
                For i = 10 * j + 1 To 10 * j + 8
                    Tempo = NombrePossible(LaTable, i, MONADVERSAIRE.Value)
                    If Tempo > 0 Then
                        ValPos = Rnd
                        If ValPos > MaxActu Then
                            MeilleurCoup = i
                            MaxActu = ValPos
                        End If
                    End If
                Next
            Next
            
        Case 1
            For j = 1 To 8
                For i = 10 * j + 1 To 10 * j + 8
                    Tempo = NombrePossible(LaTable, i, MONADVERSAIRE.Value)
                    If Tempo > 0 Then
                        ValPos = Tempo + Rnd
                        If ValPos > MaxActu Then
                            MeilleurCoup = i
                            MaxActu = ValPos
                        End If
                    End If
                Next
            Next
            
        Case 2
            For j = 1 To 8
                For i = 10 * j + 1 To 10 * j + 8
                    Tempo = NombrePossible(LaTable, i, MONADVERSAIRE.Value)
                    If Tempo > 0 Then
                        ValPos = Tempo + Rnd + ValeurCase(i)
                        If ValPos > MaxActu Then
                            MeilleurCoup = i
                            MaxActu = ValPos
                        End If
                    End If
                Next
            Next
    End Select
    PlayIn MeilleurCoup, MONADVERSAIRE
End Sub

'==================================================================================
'
'==================================================================================
Private Function ValeurCase(UneCase As Long) As Long
    Select Case UneCase
        Case 11, 18, 81, 88 'les coins
            ValeurCase = 10
        Case 12, 17, 21, 28, 71, 82, 78, 87 'les cases X et C
            ValeurCase = -10
        Case 13, 14, 15, 16, 83, 84, 85, 86, 31, 41, 51, 61, 38, 48, 58, 68 'les bords
            ValeurCase = 5
    End Select
End Function
'==================================================================================
'
'==================================================================================
Private Sub CopyTable(TableSource As UneTable, TableCible As UneTable)
Dim i As Long
Dim j As Long
    For j = 1 To 8
        For i = 10 * j + 1 To 10 * j + 8
            TableCible.Pion(i) = TableSource.Pion(i)
        Next i
    Next j
End Sub

'==================================================================================
'
'==================================================================================
Private Sub VideTable(QuelTable As UneTable)
Dim i As Long
Dim j As Long
    For j = 1 To 8
        For i = 10 * j + 1 To 10 * j + 8
            QuelTable.Pion(i) = 0
        Next i
    Next j
End Sub



'==================================================================================
'
'==================================================================================
Private Sub TCP_ConnectionRequest(ByVal RequestID As Long)
    If FRMplateau.TCP.State <> sckClosed Then FRMplateau.TCP.Close
    FRMplateau.TCP.Accept RequestID
    FRMplateau.TCP.SendData "ACCEPT"
'    FRMnewGame.LBLwait = "REQUEST from " & CStr(RequestID)
    RESEAU = True
End Sub
''==================================================================================
''
''==================================================================================
'Private Sub TCP_DataArrival(ByVal bytesTotal As Long)
'Dim Recu As String
'
'    FRMplateau.TCP.GetData Recu, vbString
'
'    If Recu = "ACCEPT" Then
'        If FRMnewGame.OPclient Then
'            FRMnewGame.LBLrequest.Caption = "REQUEST ACCEPTED"
'            FRMplateau.TCP.SendData "ACCEPT"
'        End If
'        If FRMnewGame.OPserver Then
'            FRMnewGame.LBLwait = "CONNECTED"
'            FRMplateau.TCP.SendData "CONNECTED"
'            RESEAU = True
'        End If
'    End If
'
'    If Recu = "CONNECTED" Then
'        If FRMnewGame.OPclient Then
'            FRMnewGame.LBLrequest.Caption = "CONNECTED"
'            RESEAU = True
'        End If
'    End If
'
'    If Mid(Recu, 1, 8) = "NEWGAME:" Then
'        If Mid(Recu, 9) = "MOI" Then
'            MonAdversaireIs Player(1)
'            MoiIs Player(2)
'            PICplateau.Enabled = False
'        Else
'            MonAdversaireIs Player(2)
'            MoiIs Player(1)
'        End If
'        Unload FRMnewGame
'    End If
'
'    If Mid(Recu, 1, 5) = "PLAY:" Then
'        PlayIn CLng(Mid(Recu, 6)), MONADVERSAIRE
'
'        DessineTout
'    End If
'
'End Sub

'==================================================================================
'
'==================================================================================
Private Sub TimerDessine_Timer()
    DessineTout
    If TimerDessine.Tag = "ORDIPLAY" Then
        TimerDessine.Tag = ""
        OrdinateurJoue
    End If
End Sub
