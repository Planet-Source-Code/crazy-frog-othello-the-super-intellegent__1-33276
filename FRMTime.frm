VERSION 5.00
Begin VB.Form FRMTime 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   Icon            =   "FRMTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   2895
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
      Height          =   1455
      Left            =   0
      Picture         =   "FRMTime.frx":030A
      ScaleHeight     =   1455
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Label LBLtimeBlanc 
      BackColor       =   &H00008000&
      Caption         =   "12'34''"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label LBLtimeNoir 
      BackColor       =   &H00008000&
      Caption         =   "12'34''"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   570
      Left            =   480
      Picture         =   "FRMTime.frx":1700
      Top             =   120
      Width           =   570
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   480
      Picture         =   "FRMTime.frx":287A
      Top             =   840
      Width           =   570
   End
End
Attribute VB_Name = "FRMTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Activate()
    If MDImain.Toolbar.Buttons(10).Value = tbrUnpressed Then Me.Visible = False
End Sub

'==================================================================================
'
'==================================================================================
Private Sub Form_Load()
    Me.BackColor = RGB(0, 132, 0)
    LBLtimeBlanc.BackColor = RGB(0, 132, 0)
    LBLtimeNoir.BackColor = RGB(0, 132, 0)
    Me.Top = 30
    Me.Left = 6255
End Sub


'==================================================================================
'
'==================================================================================
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If BIGcancel = 0 Then
    Else
        Cancel = 1
        Me.Visible = False
        MDImain.Toolbar.Buttons(10).Value = 0
    End If
End Sub

'==================================================================================
'
'==================================================================================
Private Sub Form_Resize()
Dim i As Long
    If Me.WindowState <> vbMinimized Then
        
        picLogo.Height = Me.ScaleHeight
        
        
    End If
End Sub


