VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOOMGlobals"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function RGBA(r As Integer, g As Integer, b As Integer, a As Integer) As Long
    RGBA = D3DColorRGBA(r, g, b, a)
End Function

Public Function Intersect(Sprite1 As BOOMSprite, Sprite2 As BOOMSprite) As Long
    Dim tmpRect As RECT
    Intersect = IntersectRect(tmpRect, Sprite1.GetDestRect, Sprite2.GetDestRect)
End Function

Public Function vec2(x!, y!) As D3DVECTOR2
    With vec2
        .x = x: .y = y
    End With
End Function

Public Function D2R(degrees As Double) As Double
    D2R = degrees * PI / 180
End Function
Public Function R2D(radians As Double) As Double
    R2D = radians * 180 / PI
End Function
