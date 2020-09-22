Attribute VB_Name = "mGlobal"
' Boom2D Library
' --------------
' Written by Bart van de Sande
' Copyright ELSe Software


Option Explicit

Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long

Public DirectX As New DirectX8
Public Direct3D As Direct3D8
Public Direct3DDevice As Direct3DDevice8
Public Direct3DX As New D3DX8
Public Sprites As D3DXSprite

Public Target As RECT

Public TPool As New BOOMTexturePool
Public SPool As New BOOMSpritePool

Public Const PI = 3.1415926

Public ScrollX As Single, ScrollY As Single

Public Function vec2(x!, y!) As D3DVECTOR2
    With vec2
        .x = x: .y = y
    End With
End Function

Public Sub Swap(val1, val2)
    Dim tmpval1, tmpval2
    tmpval1 = val1
    tmpval2 = val2
    
    val1 = tmpval2
    val2 = tmpval1
    
End Sub
