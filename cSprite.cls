VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================
' Class:    cSprite
' Filename: cSprite.cls
' Author:   SP McMahon
' Date:     14 July 1998
'
' a wrapper for m_cSpriteBitmaps to allow
' independent animation of a number of sprites based
' on the same graphic components.
' ======================================================================
'
' Memory DC to hold a copy of the background we are
' drawing over
Private m_hdcStore As Long
Private m_hBmpStore As Long
Private m_hBmpStoreOld As Long

' Reference to class holding the bitmap data:
Private m_cSpriteBitmap As cSpriteBitmaps

' Current sprite cell for this bitmap
Private m_lCurrentCell As Long
Private m_lX As Long
Private m_lY As Long
Private m_lLastX As Long
Private m_lLastY As Long
Private m_lPrevX As Long
Private m_lPrevY As Long
Private m_lWidth As Long
Private m_lHeight As Long
Private m_bBackstored As Boolean
Private m_bActive As Boolean
Private m_lXDir As Long
Private m_lYDir As Long

'==================================================================================
'
'==================================================================================
Public Property Get XDir() As Long
    XDir = m_lXDir
End Property

Public Property Let XDir(ByVal lXDir As Long)
    m_lXDir = lXDir
End Property
'==================================================================================
'
'==================================================================================
Public Property Get YDir() As Long
    YDir = m_lYDir
End Property

Public Property Let YDir(ByVal lYDir As Long)
    m_lYDir = lYDir
End Property
'==================================================================================
'
'==================================================================================
Public Sub IncrementPosition()
    m_lX = m_lX + m_lXDir
    m_lY = m_lY + m_lYDir
End Sub
'==================================================================================
'
'==================================================================================
Public Property Get Active() As Boolean
    Active = m_bActive
End Property

Public Property Let Active(ByVal bActive As Boolean)
    m_bActive = bActive
End Property
'==================================================================================
'
'==================================================================================
Public Property Get x() As Long
    x = m_lX
End Property

Public Property Let x(ByVal lX As Long)
    m_lX = lX
End Property
'==================================================================================
'
'==================================================================================
Public Property Get y() As Long
    y = m_lY
End Property

Public Property Let y(ByVal lY As Long)
    m_lY = lY
End Property
'==================================================================================
'
'==================================================================================
Public Sub StageToScreen( _
        ByVal lHDC As Long, _
        ByVal lStageHDC As Long _
    )
Static lXOffset As Long, lYOffset As Long
    
    If (BltInTwo) Then
        ' First reblit the position to redraw:
        BitBlt lHDC, PreviousX, PreviousY, m_lWidth, m_lHeight, lStageHDC, PreviousX, PreviousY, SRCCOPY
        ' Now blit the new position:
        BitBlt lHDC, m_lLastX, m_lLastY, m_lWidth, m_lHeight, lStageHDC, m_lLastX, m_lLastY, SRCCOPY
    Else
        ' Reblit the position to redraw and the new sprite
        ' all in one go:
        lXOffset = (m_lLastX - m_lPrevX)
        lYOffset = (m_lLastY - m_lPrevY)
        If (m_lLastX < m_lPrevX) Then
            If (m_lLastY < m_lPrevY) Then
                BitBlt lHDC, m_lLastX, m_lLastY, m_lWidth + Abs(lXOffset), m_lHeight + Abs(lYOffset), lStageHDC, m_lLastX, m_lLastY, SRCCOPY
            Else
                BitBlt lHDC, m_lLastX, m_lPrevY, m_lWidth + Abs(lXOffset), m_lHeight + Abs(lYOffset), lStageHDC, m_lLastX, m_lPrevY, SRCCOPY
            End If
        Else
            If (m_lLastY < m_lPrevY) Then
                BitBlt lHDC, m_lPrevX, m_lLastY, m_lWidth + Abs(lXOffset), m_lHeight + Abs(lYOffset), lStageHDC, m_lPrevX, m_lLastY, SRCCOPY
            Else
                BitBlt lHDC, m_lPrevX, m_lPrevY, m_lWidth + Abs(lXOffset), m_lHeight + Abs(lYOffset), lStageHDC, m_lPrevX, m_lPrevY, SRCCOPY
            End If
        End If
            
    End If

End Sub
'==================================================================================
'
'==================================================================================
Property Get Cell() As Long
    Cell = m_lCurrentCell
End Property

Property Let Cell(iCell As Long)
    m_lCurrentCell = iCell
End Property
'==================================================================================
'
'==================================================================================
Property Get PreviousX() As Long
    PreviousX = m_lPrevX
End Property
'==================================================================================
'
'==================================================================================
Property Get PreviousY() As Long
    PreviousY = m_lPrevY
End Property
'==================================================================================
'
'==================================================================================
Property Get BltInTwo() As Boolean
    If Abs(m_lPrevX - m_lLastX) + Abs(m_lPrevY - m_lLastY) > 24 Then
        BltInTwo = True
    End If
End Property
'==================================================================================
'
'==================================================================================
Property Get Width() As Long
    Width = m_lWidth
End Property
'==================================================================================
'
'==================================================================================
Property Get Height() As Long
    Height = m_lHeight
End Property
'==================================================================================
'
'==================================================================================
Public Sub StoreBackground( _
        ByVal hDC As Long, _
        ByVal lX As Long, _
        ByVal lY As Long _
    )
    ' Store the background at the location:
    BitBlt m_hdcStore, 0, 0, m_lWidth, m_lHeight, hDC, lX, lY, SRCCOPY
    m_bBackstored = True

End Sub
'==================================================================================
'
'==================================================================================
Public Sub TransparentDraw( _
        ByVal hDC As Long, _
        ByVal lX As Long, _
        ByVal lY As Long, _
        ByVal lCell As Long, _
        Optional ByVal bUseBuildDC As Variant _
    )
    m_lCurrentCell = lCell
    m_lPrevX = m_lLastX
    m_lPrevY = m_lLastY
    m_lLastX = lX
    m_lLastY = lY
    m_cSpriteBitmap.TransparentDraw hDC, lX, lY, lCell, bUseBuildDC
End Sub
'==================================================================================
'
'==================================================================================
Public Sub Create( _
        ByVal hDC As Long _
    )
    ' Clear up any DCs/Bitmaps we are already holding:
    ClearUp
    ' Now get the width and height for the sprite:
    With m_cSpriteBitmap
        m_lWidth = .CellWidth
        m_lHeight = .CellHeight
    End With
    ' Create DCs and Bitmaps to hold the backdrop
    GDIMakeDCAndBitmap False, m_hdcStore, m_hBmpStore, m_hBmpStoreOld, m_lWidth, m_lHeight
End Sub
'==================================================================================
'
'==================================================================================
Public Sub RestoreBackground( _
        ByVal hDC As Long _
    )
    If (m_bBackstored) Then
        ' Blit the stored background back onto the
        ' destination HDC:
        BitBlt hDC, m_lLastX, m_lLastY, m_lWidth, m_lHeight, m_hdcStore, 0, 0, SRCCOPY
    End If
End Sub
'==================================================================================
'
'==================================================================================
Property Let SpriteData(cSpriteData As cSpriteBitmaps)
    Set m_cSpriteBitmap = cSpriteData
    m_lWidth = m_cSpriteBitmap.CellWidth
    m_lHeight = m_cSpriteBitmap.CellHeight
End Property
'==================================================================================
'
'==================================================================================
Private Sub ClearUp()
    ' Here we ensure all our Objects are destroyed:
    GDIClearDCBitmap m_hdcStore, m_hBmpStore, m_hBmpStoreOld
    ' We haven't a valid background:
    m_bBackstored = False
End Sub
'==================================================================================
'
'==================================================================================
Private Sub Class_Initialize()
    m_lPrevX = -Screen.Width \ Screen.TwipsPerPixelX
    m_lLastX = m_lPrevX
    m_lPrevY = -Screen.Height \ Screen.TwipsPerPixelY
    m_lLastY = m_lPrevY
End Sub
'==================================================================================
'
'==================================================================================
Private Sub Class_Terminate()
    ClearUp
    Set m_cSpriteBitmap = Nothing
End Sub
