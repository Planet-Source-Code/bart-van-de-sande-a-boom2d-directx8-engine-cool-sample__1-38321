VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BOOMSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

' SPRITE CHANGEABLE PARAMETERS
Private Color As Long
Private Position As D3DVECTOR2
Private Rotation As Single
Private RotationCenter As D3DVECTOR2
Private Scaling As D3DVECTOR2
Public Active As Boolean

Public XMirror As Boolean
Public YMirror As Boolean

Public IsTile As Boolean

Private vCustomVector1 As D3DVECTOR2
Private vCustomVector2 As D3DVECTOR2

'---------------------------------

' ANIMATION INFO:
Private bAnimated As Boolean
Private aniFlags As Long
Private AniRows As Long, AniCols As Long
Private AniCurrentFrame As Long
Private aniMinFrame As Long, aniMaxFrame As Long
Private AniFrameWidth&, AniFrameHeight&
Private AniInterval As Single

' SPRITE INFO:
Private SpriteWidth As Long
Private SpriteHeight As Long

' SPRITE GRAPHIC POINTER
Private ptTexture

' RENDER VALUES
Private DestRect As RECT
Private SourceRect As RECT

Public Key As String

' SET/GET
Public Sub SetPosition(pos As D3DVECTOR2)
    Position = pos
End Sub
Public Function GetPosition() As D3DVECTOR2
    GetPosition = Position
End Function
Public Sub SetRotation(rot!)
    Rotation = rot
End Sub
Public Function GetRotation() As Single
    GetRotation = Rotation
End Function
Public Sub SetColor(col&)
    Color = col
End Sub
Public Function GetColor() As Long
    GetColor = Color
End Function
Public Sub SetScale(scalevec As D3DVECTOR2)
    Scaling = scalevec
    If bAnimated = False Then
        RotationCenter.x = SpriteWidth / 2 * Scaling.x
        RotationCenter.y = SpriteHeight / 2 * Scaling.y
    Else
        RotationCenter.x = AniFrameWidth / 2 * Scaling.x
        RotationCenter.y = AniFrameHeight / 2 * Scaling.y
    End If
End Sub
Public Function GetScale() As D3DVECTOR2
    GetScale = Scaling
End Function
' ANIMATION:
Public Sub CreateAnimation(animated As Boolean, rows&, cols&, flags&, updateinterval!)
    If animated Then
        bAnimated = True
        AniRows = rows
        AniCols = cols
        aniFlags = flags
        AniFrameWidth = SpriteWidth / cols
        AniFrameHeight = SpriteHeight / rows
        AniInterval = updateinterval
        aniMaxFrame = rows * cols
        aniMinFrame = 1
        AniCurrentFrame = 1
        
        RotationCenter.x = AniFrameWidth / 2 * Scaling.x
        RotationCenter.y = AniFrameHeight / 2 * Scaling.y
    Else
        bAnimated = False
    End If
End Sub
Public Sub CreateAnimation2(animated As Boolean, rows&, cols&, framewidth&, frameheight&, maxframe&, flags&, updateinterval!)
    If animated Then
        bAnimated = True
        AniRows = rows
        AniCols = cols
        aniFlags = flags
        AniFrameWidth = framewidth
        AniFrameHeight = frameheight
        AniInterval = updateinterval
        aniMaxFrame = maxframe
        aniMinFrame = 1
        AniCurrentFrame = 1
        
        RotationCenter.x = AniFrameWidth / 2 * Scaling.x
        RotationCenter.y = AniFrameHeight / 2 * Scaling.y
    Else
        bAnimated = False
    End If
End Sub

Public Function GetAnimationRows() As Long
    GetAnimationRows = AniRows
End Function
Public Function GetAnimationCols() As Long
    GetAnimationCols = AniCols
End Function
Public Function GetCurrentFrame() As Long
    GetCurrentFrame = AniCurrentFrame
End Function
Public Sub SetCurrentFrame(f&)
    AniCurrentFrame = f
End Sub
Public Sub IncreaseFrame()
    AniCurrentFrame = AniCurrentFrame + 1
    If AniCurrentFrame > aniMaxFrame Then
        AniCurrentFrame = aniMinFrame
    End If
End Sub
Public Sub DecreaseFrame()
    AniCurrentFrame = AniCurrentFrame - 1
    If AniCurrentFrame < aniMinFrame Then
        AniCurrentFrame = aniMaxFrame
    End If
End Sub
Public Sub SetAnimationMinFrame(f&)
    aniMinFrame = f
End Sub
Public Sub SetAnimationMaxFrame(f&)
    aniMaxFrame = f
End Sub
Public Function GetAnimationMinFrame() As Long
    GetAnimationMinFrame = aniMinFrame
End Function
Public Function GetAnimationMaxFrame() As Long
    GetAnimationMaxFrame = aniMaxFrame
End Function
Public Function GetDestRect() As RECT
    GetDestRect = DestRect
End Function
Public Sub Move(x!, y!)
    Position.x = Position.x + x
    Position.y = Position.y + y
End Sub
Public Sub SetRotCenter(v As D3DVECTOR2)
    RotationCenter = v
End Sub
Public Function GetRotCenter() As D3DVECTOR2
    GetRotCenter = RotationCenter
End Function
Public Sub Render()
    If Active Then
        
        Dim iCol&, iRow&
        Dim sVec As D3DVECTOR2
        Dim offX!, offY!
        Static fTimer!
    
        sVec = Scaling
        If XMirror Then
            sVec.x = -sVec.x
            If bAnimated Then
                offX = AniFrameWidth
            Else
                offX = SpriteWidth
            End If
        End If
        If YMirror Then
            sVec.y = -sVec.y
            If bAnimated Then
                offY = AniFrameHeight
            Else
                offY = SpriteHeight
            End If
        End If
        
        If bAnimated Then
            ' DESTINIATION RECTANGLE, USED FOR INTERSECTING SPRITES
            DestRect.Left = Position.x
            DestRect.Top = Position.y
            DestRect.Right = Position.x + AniFrameWidth * Scaling.x
            DestRect.bottom = Position.y + AniFrameHeight * Scaling.y
            ' increase, decrease frame if specified
            If Timer - fTimer >= AniInterval And AniInterval <> -1 Then
                fTimer = Timer
                
                If aniFlags = 0 Then
                    AniCurrentFrame = AniCurrentFrame + 1
                Else
                    AniCurrentFrame = AniCurrentFrame - 1
                End If
                
                ' bounds of animation:
                If AniCurrentFrame > aniMaxFrame Then
                    AniCurrentFrame = aniMinFrame
                End If
                If AniCurrentFrame < aniMinFrame Then
                    AniCurrentFrame = aniMaxFrame
                End If
            End If
            
        
            ' calculate current row, col
            If (AniCurrentFrame - 1) / AniCols >= 1 Then
                iRow = Int((AniCurrentFrame - 1) / AniCols)
            Else
                iRow = 0
            End If
            iCol = (AniCurrentFrame - 1) - (iRow * AniCols)
            
            SourceRect.Left = iCol * AniFrameWidth
            SourceRect.Right = SourceRect.Left + AniFrameWidth
            SourceRect.Top = iRow * AniFrameHeight
            SourceRect.bottom = SourceRect.Top + AniFrameHeight
            
            If Active Then
                If ScrollX + Position.x + SpriteWidth > 0 And ScrollX + Position.x < (Target.Right - Target.Left) Then
                    If ScrollY + Position.y + SpriteHeight > 0 And ScrollY + Position.y < (Target.bottom - Target.Top) Then
                        Sprites.Draw TPool(ptTexture).GetTexture, SourceRect, sVec, RotationCenter, Rotation, vec2(Position.x + offX + ScrollX, Position.y + offY + ScrollY), Color
                    End If
                End If
            End If
        Else
            ' DESTINIATION RECTANGLE, USED FOR INTERSECTING SPRITES
            DestRect.Left = Position.x
            DestRect.Top = Position.y
            DestRect.Right = Position.x + SpriteWidth * Scaling.x
            DestRect.bottom = Position.y + SpriteHeight * Scaling.y
            
            If IsTile = False Then
                SourceRect.Left = 0: SourceRect.Right = SpriteWidth
                SourceRect.Top = 0: SourceRect.bottom = SpriteHeight
            End If
            
            If Active Then
                If ScrollX + Position.x + SpriteWidth > 0 And ScrollX + Position.x < (Target.Right - Target.Left) Then
                    If ScrollY + Position.y + SpriteHeight > 0 And ScrollY + Position.y < (Target.bottom - Target.Top) Then
                        Sprites.Draw TPool(ptTexture).GetTexture, ByVal 0, sVec, RotationCenter, Rotation, vec2(Position.x + offX + ScrollX, Position.y + offY + ScrollY), Color
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Sub SetTileRect(x&, y&, x1&, y1&)
    SourceRect.Left = x
    SourceRect.Top = y
    SourceRect.Right = x1
    SourceRect.bottom = y1
End Sub

Public Sub SetCustomVector1(x!, y!)
    vCustomVector1 = vec2(x, y)
End Sub
Public Sub SetCustomVector2(x!, y!)
    vCustomVector2 = vec2(x, y)
End Sub
Public Function GetCustomVector1() As D3DVECTOR2
    GetCustomVector1 = vCustomVector1
End Function
Public Function GetCustomVector2() As D3DVECTOR2
    GetCustomVector2 = vCustomVector2
End Function

Friend Sub CreateFromTexture(tex)
    ptTexture = tex
    
    SpriteWidth = TPool(tex).GetWidth
    SpriteHeight = TPool(tex).GetHeight
    
    ' set default values
    RotationCenter.x = SpriteWidth / 2: RotationCenter.y = SpriteHeight / 2
    Position.x = 0: Position.y = 0
    Rotation = 0
    Color = D3DColorRGBA(255, 255, 255, 255)
    Scaling.x = 1: Scaling.y = 1
    
    Active = True
End Sub


