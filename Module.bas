Attribute VB_Name = "ModuleFRED"
Option Explicit

'**********************************************************************************
'   FONCTION SUR DES OBJETS VISUELS
'**********************************************************************************

'==================================================================================
'   selectionne tout le contenu d'une zone text
'   à appeler dans l'evenement gotfocus
Public Sub SelectAll(ByRef ZoneText As TextBox)
    ZoneText.SelStart = 0
    ZoneText.SelLength = Len(ZoneText)
End Sub

'==================================================================================
'   classe les colonnes d'une ListView en fonction de la colonne clicker
'   a appeler dans l'evenement ColumnClick
'   ATTENTION utilise le TAG de la ListView mettre 0 dans celui ci avant
Public Sub ClasseLesColonnes(UneListView As ListView, COLONNE As MSComctlLib.ColumnHeader)
    With UneListView
        .Sorted = True
        If .Tag = COLONNE.Index Then   'si on click sur la meme colonne
            If .SortOrder = lvwAscending Then 'inversion de l'ordre de classement
               .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
            .SortKey = COLONNE.Index - 1
        Else
            .SortOrder = lvwAscending 'classe sur cette colonne et par ordre
            .SortKey = COLONNE.Index - 1
        End If
        .Tag = COLONNE.Index   'stock la derniere colonne clicker dans le TAG
    End With
End Sub


'**********************************************************************************
'   FONCTIONS MATHEMATIQUES
'**********************************************************************************

'==================================================================================
'
'==================================================================================
Public Function Max(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then Max = a Else Max = b
End Function

'==================================================================================
'
'==================================================================================
Public Function Min(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then Min = a Else Min = b
End Function

'==================================================================================
' retourne TRUE si la chaine peut etre converti en un long
' il est possible d'obtenir directement le resultat dans le deuxieme paramètre
Public Function IsLong(chaine As String, Optional EntierLong As Long) As Boolean
On Error GoTo non
    EntierLong = CLng(chaine)
    IsLong = True
    Exit Function
non:
    IsLong = False
End Function

'==================================================================================
' retourne TRUE si la chaine peut etre converti en un double
' il est possible d'obtenir directement le resultat dans le deuxieme paramètre
Public Function IsDouble(ByRef chaine As String, Optional NBdouble As Double) As Boolean
On Error GoTo non
    chaine = Replace(chaine, ".", ",")
    NBdouble = CDbl(chaine)
    IsDouble = True
    Exit Function
non:
    IsDouble = False
End Function


