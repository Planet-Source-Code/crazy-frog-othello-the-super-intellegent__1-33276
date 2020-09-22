VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMhistorique 
   Caption         =   "Story"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9810
   Icon            =   "FRMhistorique.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   9810
   Visible         =   0   'False
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
      Height          =   8160
      Left            =   0
      Picture         =   "FRMhistorique.frx":030A
      ScaleHeight     =   8160
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   -840
      Width           =   255
   End
   Begin MSComctlLib.ImageList IL_LV 
      Left            =   8880
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMhistorique.frx":71CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMhistorique.frx":72DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMhistorique.frx":73F0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LVstory 
      Height          =   7335
      Left            =   255
      TabIndex        =   0
      Tag             =   "0"
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   12938
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "IL_LV"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Turn"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "P1"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "P2"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Captured"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Black"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "White"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Difference"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Free"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "FRMhistorique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit







Private Sub Form_Activate()
    FRMgraphic.Dessine
    If MDImain.Toolbar.Buttons(6).Value = tbrUnpressed Then Me.Visible = False
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If BIGcancel = 0 Then
        
    Else
        Cancel = 1
        Me.Visible = False
        MDImain.Toolbar.Buttons(6).Value = 0
    End If
End Sub
'==================================================================================
'
'==================================================================================
Private Sub Form_Resize()
Dim i As Long
    LVstory.Visible = False
    If Me.WindowState <> vbMinimized Then
        picLogo.Top = FRMhistorique.ScaleHeight - picLogo.Height
        LVstory.Width = FRMhistorique.ScaleWidth - LVstory.Left
        LVstory.Height = FRMhistorique.ScaleHeight - LVstory.Top
        For i = 1 To 9
            LVstory.ColumnHeaders(i).Width = LVstory.Width \ 9 - 50
        Next i
        
        LVstory.Visible = True
    End If
    
End Sub

