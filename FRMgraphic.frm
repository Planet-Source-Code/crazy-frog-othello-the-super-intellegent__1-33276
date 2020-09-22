VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FRMgraphic 
   BackColor       =   &H00008000&
   Caption         =   "Graph"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8490
   FillColor       =   &H00FFFFFF&
   Icon            =   "FRMgraphic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5085
   ScaleWidth      =   8490
   Visible         =   0   'False
   Begin MSComctlLib.Slider SLDmax 
      Height          =   1575
      Left            =   8040
      TabIndex        =   1
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   2778
      _Version        =   393216
      Orientation     =   1
      LargeChange     =   1
      Min             =   1
      SelStart        =   5
      TickStyle       =   1
      Value           =   5
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
      Height          =   8565
      Left            =   0
      Picture         =   "FRMgraphic.frx":030A
      ScaleHeight     =   8565
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   -3480
      Width           =   255
   End
   Begin VB.Label LBLnoir 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   180
      Width           =   90
   End
   Begin VB.Image Image2 
      Height          =   570
      Left            =   360
      Picture         =   "FRMgraphic.frx":7748
      Top             =   0
      Width           =   570
   End
   Begin VB.Label LBLblanc 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "2"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1560
      TabIndex        =   3
      Top             =   180
      Width           =   90
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   1320
      Picture         =   "FRMgraphic.frx":88C2
      Top             =   0
      Width           =   570
   End
End
Attribute VB_Name = "FRMgraphic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim Adessine As Boolean







Public Sub Dessine()
    Dim i As Long
    
    DrawWidth = 1
    FRMgraphic.Refresh
    DrawStyle = 0
    ForeColor = RGB(0, 0, 0)
    Line (0, FRMgraphic.ScaleHeight)-(255, FRMgraphic.ScaleHeight)
    For i = 1 To FRMhistorique.LVstory.ListItems.Count Step 2
        Line -(i * FRMgraphic.ScaleWidth / 63, FRMgraphic.ScaleHeight - CLng(FRMhistorique.LVstory.ListItems(i).SubItems(4)) * FRMgraphic.ScaleHeight / (24 + SLDmax.Value * 4))
    Next
    
    ForeColor = RGB(0, 0, 0)
    Line (0, FRMgraphic.ScaleHeight)-(255, FRMgraphic.ScaleHeight)
    DrawStyle = 2
    For i = 1 To FRMhistorique.LVstory.ListItems.Count Step 2
        Line -(i * FRMgraphic.ScaleWidth / 63, FRMgraphic.ScaleHeight - CLng(FRMhistorique.LVstory.ListItems(i).SubItems(7)) * FRMgraphic.ScaleHeight / (24 + SLDmax.Value * 4))
    Next
    DrawStyle = 0
    ForeColor = RGB(255, 255, 255)
    Line (0, FRMgraphic.ScaleHeight)-(255, FRMgraphic.ScaleHeight)
    For i = 1 To FRMhistorique.LVstory.ListItems.Count Step 2
        Line -(i * FRMgraphic.ScaleWidth / 63, FRMgraphic.ScaleHeight - CLng(FRMhistorique.LVstory.ListItems(i).SubItems(5)) * FRMgraphic.ScaleHeight / (24 + SLDmax.Value * 4))
    Next
    
    DrawStyle = 2
    ForeColor = RGB(255, 255, 255)
    Line (0, FRMgraphic.ScaleHeight)-(255, FRMgraphic.ScaleHeight)
    For i = 2 To FRMhistorique.LVstory.ListItems.Count Step 2
        Line -(i * FRMgraphic.ScaleWidth / 63, FRMgraphic.ScaleHeight - CLng(FRMhistorique.LVstory.ListItems(i).SubItems(7)) * FRMgraphic.ScaleHeight / (24 + SLDmax.Value * 4))
    Next
    
    
    LBLnoir = FRMhistorique.LVstory.ListItems(FRMhistorique.LVstory.ListItems.Count).SubItems(4)
    LBLblanc = FRMhistorique.LVstory.ListItems(FRMhistorique.LVstory.ListItems.Count).SubItems(5)
    
    
End Sub

Private Sub Form_Activate()
    FRMgraphic.Dessine
    If MDImain.Toolbar.Buttons(4).Value = tbrUnpressed Then Me.Visible = False
End Sub

Private Sub Form_Click()
    FRMgraphic.Dessine
End Sub

'==================================================================================
'
'==================================================================================
Private Sub Form_Load()
    Me.Width = 480 * 15
    Me.Height = 128 * 15
    Me.BackColor = RGB(0, 132, 0)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Adessine Then
        Adessine = False
        Dessine
    End If
End Sub

'==================================================================================
'
'==================================================================================
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If BIGcancel = 0 Then
       
    Else
        Cancel = 1
        Me.Visible = False
        MDImain.Toolbar.Buttons(4).Value = 0
    End If
End Sub

'==================================================================================
'
'==================================================================================
Private Sub Form_Resize()
Dim i As Long
    If Me.WindowState <> vbMinimized Then
        SLDmax.Left = FRMgraphic.ScaleWidth - SLDmax.Width
        picLogo.Top = FRMgraphic.ScaleHeight - picLogo.Height
        
        Adessine = True
    End If
End Sub




Private Sub SLDmax_Change()
    Dessine
End Sub

Private Sub SLDmax_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 0 Then
        Dessine
    End If
End Sub
