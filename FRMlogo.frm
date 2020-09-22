VERSION 5.00
Begin VB.Form FRMlogo 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   Icon            =   "FRMlogo.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   1695
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "FRMlogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    FRMlogo.Width = 255
    FRMlogo.Height = MDImain.ScaleHeight
    
    BIGcancel = 1
End Sub

