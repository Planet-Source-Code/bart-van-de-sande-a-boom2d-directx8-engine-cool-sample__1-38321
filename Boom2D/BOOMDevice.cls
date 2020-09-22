VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOOMDevice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public ForeColor As Long
Public CurrentFont As BOOMFont

Public Function CreateFont(font As IFont) As BOOMFont
    Dim tmpFont As New BOOMFont
    tmpFont.Create font
    Set CreateFont = tmpFont
End Function

Public Sub DrawText(text$, x&, y&, x1&, y1&, Optional format As CONST_DTFLAGS = 0)
    Dim rt As RECT
    rt.Left = x: rt.Right = x1
    rt.Top = y: rt.bottom = y1
    Direct3DX.DrawText CurrentFont.GetFont, ForeColor, text, rt, format
End Sub

Public Sub SetScroll(vec As D3DVECTOR2)
    ScrollX = vec.x: ScrollY = vec.y
End Sub
Public Function GetScroll() As D3DVECTOR2
    GetScroll.x = ScrollX: GetScroll.y = ScrollY
End Function
Public Sub Scroll(x!, y!)
    ScrollX = ScrollX + x
    ScrollY = ScrollY + y
End Sub
