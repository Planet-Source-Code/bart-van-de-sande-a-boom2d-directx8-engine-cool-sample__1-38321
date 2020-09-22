VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOOMTexture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' NOTE: Texture sizes must be power of 2!!!!!

Option Explicit

Private g_Texture As Direct3DTexture8
Private g_Info As D3DXIMAGE_INFO

Public Key As String

Friend Sub CreateFromFile(file$, Optional w& = -1, Optional h& = -1, Optional colorkey As Long = 0, Optional fmt As CONST_D3DFORMAT = 0)
    Dim mipinfo As D3DSURFACE_DESC
    ' create the texture
    
    Set g_Texture = Direct3DX.CreateTextureFromFileEx(Direct3DDevice, file, w, h, 1, 0, fmt, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, colorkey, g_Info, ByVal 0)
    g_Texture.GetLevelDesc 0, mipinfo
    
    ' set the size parameters if needed
    If w = -1 Then
        g_Info.Width = mipinfo.Width
    End If
    If h = -1 Then
        g_Info.Height = mipinfo.Height
    End If
End Sub
Friend Function GetTexture() As Direct3DTexture8
    Set GetTexture = g_Texture
End Function
Friend Function GetWidth() As Long
    GetWidth = g_Info.Width
End Function
Friend Function GetHeight() As Long
    GetHeight = g_Info.Height
End Function
Friend Sub Destroy()
    Set g_Texture = Nothing
End Sub
