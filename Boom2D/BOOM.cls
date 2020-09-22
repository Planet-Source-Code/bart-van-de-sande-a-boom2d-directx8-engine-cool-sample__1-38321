VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public ClearColor As Long

Private lFPS&
Private rTimer As Single

Public Sub Initialize(hWnd&, bWindowed As Boolean, Optional w& = -1, Optional h& = -1, Optional lAdapter As Long = 0, Optional DevType As CONST_D3DDEVTYPE, Optional bVSync As Boolean = True)
    Dim params As D3DPRESENT_PARAMETERS
    Dim dp As D3DDISPLAYMODE
    ' intializes the engine
    
    Set Direct3D = DirectX.Direct3DCreate
    
    ' retrieve the display mode
    Direct3D.GetAdapterDisplayMode lAdapter, dp
    
    If DevType = 0 Then DevType = D3DDEVTYPE_HAL
    
    If bWindowed Then
        GetWindowRect hWnd, Target
        w = Target.Right - Target.Left
        h = Target.bottom - Target.Top
        
        With params
            .BackBufferFormat = dp.format
            .EnableAutoDepthStencil = 0 'FALSE
            '.AutoDepthStencilFormat = D3DFMT_D16
            '.BackBufferWidth = w
            '.BackBufferHeight = h
            .Windowed = 1   'TRUE
            If bVSync Then
                .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
            Else
                .SwapEffect = D3DSWAPEFFECT_COPY
            End If
        End With
    Else
        Target.Left = 0
        Target.Right = w
        Target.Top = 0
        Target.bottom = h
        With params
            .BackBufferFormat = dp.format
            .EnableAutoDepthStencil = 0 'FALSE
            '.AutoDepthStencilFormat = D3DFMT_D16
            .BackBufferWidth = w
            .BackBufferHeight = h
            .Windowed = 0   'FALSE
            .SwapEffect = D3DSWAPEFFECT_FLIP
        End With
    End If
    
    Set Direct3DDevice = Direct3D.CreateDevice(lAdapter, DevType, hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, params)
End Sub

Public Sub Clear()
    Direct3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, ClearColor, 1, 0
    Direct3DDevice.BeginScene
    Sprites.Begin
End Sub
Public Sub Render(Optional hWnd As Long)
    Static fpsTimer!
    Static fpsCounter&
    fpsCounter = fpsCounter + 1
    If Timer - fpsTimer >= 1 Then
        lFPS = fpsCounter
        fpsTimer = Timer
        fpsCounter = 0
    End If
    
    rTimer = Timer - rTimer
    
    Sprites.End
    Direct3DDevice.EndScene
    Direct3DDevice.Present ByVal 0, ByVal 0, hWnd, ByVal 0
End Sub

Public Function GetFPS() As Single
    GetFPS = lFPS
End Function

Public Function CreateDevice(lFlags As Long) As BOOMDevice
    Dim newDevice As New BOOMDevice
    Set CreateDevice = newDevice
    
    ' intialize some other stuff
    Set Sprites = Direct3DX.CreateSprite(Direct3DDevice)
    
End Function
Public Function TimeElapsed() As Single
    TimeElapsed = rTimer
End Function
Public Function GetTexturePool() As BOOMTexturePool
    Set GetTexturePool = TPool
End Function
Public Function GetSpritePool() As BOOMSpritePool
    Set GetSpritePool = SPool
End Function
Public Sub ShowCursor(bShow As Boolean)
    If bShow Then
        Direct3DDevice.ShowCursor 1
    Else
        Direct3DDevice.ShowCursor 0
    End If
End Sub
Private Sub Class_Terminate()
    Direct3DDevice.ShowCursor 1

    TPool.Clear
    
    Set Sprites = Nothing
    Set Direct3DDevice = Nothing
    Set Direct3D = Nothing
    
End Sub
