VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOOMFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private g_Font As D3DXFont

Private Info As LOGFONT

Public Function GetInfo() As LOGFONT
    GetInfo = Info
End Function

Friend Sub Create(fnt As IFont)
    Set g_Font = Direct3DX.CreateFont(Direct3DDevice, fnt.hFont)
    g_Font.GetLogFont Info ' return info about the font
End Sub
Friend Function GetFont() As D3DXFont
    Set GetFont = g_Font
End Function
