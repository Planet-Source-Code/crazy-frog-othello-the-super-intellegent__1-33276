VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDImain 
   BackColor       =   &H8000000C&
   Caption         =   "Othello Ver-1.00 by Samar Pathania"
   ClientHeight    =   5550
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   7890
   Icon            =   "MDImain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList IL_ICO 
      Left            =   6480
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":09CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":0CE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":0FFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1318
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1632
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":194C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList IL_TB 
      Left            =   6480
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":1C66
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":21B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":22CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":23DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDImain.frx":24EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   688
      ButtonWidth     =   1561
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      TextAlignment   =   1
      ImageList       =   "IL_TB"
      DisabledImageList=   "IL_TB"
      HotImageList    =   "IL_TB"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Game"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "New"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Exit"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Graph"
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Story"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Option"
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Time"
            ImageIndex      =   5
            Style           =   1
            Value           =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MDImain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DEBUGMODE As Boolean

Private Sub IT_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If FRMoption.TXTshell.Text <> "" Then
            Shell FRMoption.TXTshell.Text, 1
        Else
            Me.WindowState = vbNormal
            Me.Visible = Not Me.Visible
        End If
    Else
        Me.WindowState = vbNormal
        Me.Visible = Not Me.Visible
    End If
End Sub

'==================================================================================
'
'==================================================================================
Private Sub MDIForm_Activate()

    FRMlogo.Top = 0
    FRMlogo.Height = MDImain.ScaleHeight
    FRMplateau.Show
    FRMTime.Show
End Sub

'==================================================================================
'
'==================================================================================
Private Sub MDIForm_Click()
    If FRMoption.CHKhide.Value = vbChecked Then
        MDImain.Visible = False
    End If
End Sub

'==================================================================================
'
'==================================================================================
Private Sub MDIForm_Load()
    MDImain.Width = 640 * Screen.TwipsPerPixelX
    MDImain.Height = 480 * Screen.TwipsPerPixelY
    FRMlogo.Top = 0
    FRMlogo.Height = MDImain.ScaleHeight
    MDImain.BackColor = RGB(83, 83, 83)
    
#If DEBUGMODE Then
    'MsgBox "aie"
        gHW = Me.hwnd
        myNID.cbSize = Len(myNID)
        myNID.hwnd = gHW
        myNID.uID = uID
        myNID.uFlags = NIF_MESSAGE Or NIF_TIP Or NIF_ICON
        myNID.uCallbackMessage = cbNotify
        myNID.hIcon = IL_ICO.ListImages.Item(5).Picture
        myNID.szTip = Me.Caption & Chr(0)
        ShellNotifyIcon NIM_ADD, myNID

        Hook
#End If
    
End Sub

'==================================================================================
'
'==================================================================================
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload FRMlogo
End Sub

'==================================================================================
'
'==================================================================================
Private Sub MDIForm_Resize()
    If WindowState <> vbMinimized Then
        FRMlogo.Top = 0
        FRMlogo.Height = MDImain.ScaleHeight
    Else
        If FRMoption.CHKhide.Value = vbChecked Then
            Me.Visible = False
            MDImain.WindowState = vbNormal
        End If
    End If
End Sub
'==================================================================================
'
'==================================================================================
Private Sub MDIForm_Unload(Cancel As Integer)
#If Not DEBUGMODE Then
    unhook
    ShellNotifyIcon NIM_DELETE, myNID
#End If
End Sub

'==================================================================================
'
'==================================================================================
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Caption
        Case "Game"
            FRMplateau.SetFocus
        Case "Graph"
            If Button.Value = 1 Then
                FRMgraphic.Visible = True
                FRMgraphic.SetFocus
            Else
                FRMgraphic.Visible = False
            End If
      
        Case "Story"
            If Button.Value = 1 Then
                FRMhistorique.Visible = True
                FRMhistorique.SetFocus
            Else
                FRMhistorique.Visible = False
            End If
        Case "Option"
            If Button.Value = 1 Then
                FRMoption.Visible = True
                FRMoption.SetFocus
            Else
                FRMoption.Visible = False
            End If
        Case "Time"
            If Button.Value = 1 Then
                FRMTime.Visible = True
                FRMTime.SetFocus
            Else
                FRMTime.Visible = False
            End If
    End Select
End Sub

'==================================================================================
'
'==================================================================================
Private Sub Toolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Text
        Case "New"
            FRMnewGame.Show vbModal
        Case "Exit"
            Unload Me
    End Select
End Sub
