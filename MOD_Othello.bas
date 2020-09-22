Attribute VB_Name = "MOD_Othello"
Option Explicit

Public BIGcancel As Long



Public RESEAU As Boolean

Public ModeDeJeu As Long

Public LesPionBMPs As cSpriteBitmaps
Public PionBMP(100) As cSprite

Public Type UnHistorique
    Turn As Long
End Type

Public Type APlayer
    Sprite As Long
    Value As Long
    Name As String
    Number As Long
End Type

Public AncienCoup As Long
Public aRejouer As Boolean

Public Type LesOptions
    SpeedAnim As Long
    SpeedAction As Long
    StylePrevien As Long
    Coordonnee As Long
    Level As Long
End Type

Public Options As LesOptions

Public PionRetourne As Long

Public Player(1 To 2) As APlayer
Public MOI As APlayer
Public MONADVERSAIRE As APlayer
Public Aqui As Long


Public TailleCase As Long
Public TaillePionBMP As Long
Public Marge As Long

'==================================================================================
'
'==================================================================================
Public Sub MoiIs(Joueur As APlayer)
     MOI.Name = Joueur.Name
     MOI.Sprite = Joueur.Sprite
    MOI.Value = Joueur.Value
    MOI.Number = Joueur.Number
End Sub

'==================================================================================
'
'==================================================================================
Public Sub MonAdversaireIs(Joueur As APlayer)
    MONADVERSAIRE.Name = Joueur.Name
    MONADVERSAIRE.Sprite = Joueur.Sprite
    MONADVERSAIRE.Value = Joueur.Value
    MONADVERSAIRE.Number = Joueur.Number
End Sub

