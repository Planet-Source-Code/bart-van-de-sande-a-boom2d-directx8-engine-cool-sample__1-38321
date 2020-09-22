VERSION 5.00
Begin VB.Form frmSample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Boom2D Sample"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9075
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picScreen 
      BackColor       =   &H00000000&
      Height          =   6015
      Left            =   60
      ScaleHeight     =   397
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   1
      Top             =   60
      Width           =   5235
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sprites"
      Height          =   2715
      Left            =   5340
      TabIndex        =   0
      Top             =   60
      Width           =   3675
      Begin VB.CheckBox chkRandom 
         Caption         =   "Select random color"
         Height          =   255
         Left            =   1140
         TabIndex        =   10
         Top             =   300
         Width           =   1995
      End
      Begin VB.VScrollBar scrAlpha 
         Height          =   1455
         LargeChange     =   10
         Left            =   2880
         Max             =   0
         Min             =   255
         TabIndex        =   9
         Top             =   600
         Value           =   255
         Width           =   255
      End
      Begin VB.PictureBox picColor 
         BackColor       =   &H00000000&
         Height          =   1515
         Left            =   1140
         ScaleHeight     =   1455
         ScaleWidth      =   675
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
      Begin VB.VScrollBar scrColor 
         Height          =   1515
         Index           =   2
         LargeChange     =   10
         Left            =   840
         Max             =   0
         Min             =   255
         TabIndex        =   6
         Top             =   600
         Value           =   255
         Width           =   255
      End
      Begin VB.VScrollBar scrColor 
         Height          =   1515
         Index           =   1
         LargeChange     =   10
         Left            =   540
         Max             =   0
         Min             =   255
         TabIndex        =   5
         Top             =   600
         Value           =   255
         Width           =   255
      End
      Begin VB.VScrollBar scrColor 
         Height          =   1515
         Index           =   0
         LargeChange     =   10
         Left            =   240
         Max             =   0
         Min             =   255
         TabIndex        =   4
         Top             =   600
         Value           =   255
         Width           =   255
      End
      Begin VB.CommandButton cmdAddSprite 
         Caption         =   "Add Sprite"
         Height          =   435
         Left            =   120
         TabIndex        =   2
         Top             =   2160
         Width           =   3435
      End
      Begin VB.Label Label2 
         Caption         =   "Alpha:"
         Height          =   255
         Left            =   2160
         TabIndex        =   8
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Color:"
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Engine written by Bart van de Sande - (C)opyright ELSe Software."
      Height          =   435
      Left            =   5340
      TabIndex        =   12
      Top             =   4380
      Width           =   3675
   End
   Begin VB.Label Label3 
      Caption         =   $"frmSample.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   5340
      TabIndex        =   11
      Top             =   2820
      Width           =   3675
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type EXPLOSION
    active As Boolean
    spriteindex As Long
End Type

Private Engine As New BOOM
Private Device As BOOMDevice
Private Textures As BOOMTexturePool
Private Sprites As BOOMSpritePool
Private bEnd As Boolean
Private Explosions() As EXPLOSION
Private lSpriteNumber As Long

Private Sub AddExplosion(x!, y!)
    Dim i&
    Dim lSlot&
    ' adds an explosion to the world
    ' first, find an open explosion slot
    For i = 1 To UBound(Explosions)
        If Explosions(i).active = False And Explosions(i).spriteindex <> 0 Then
            lSlot = i
            Exit For
        End If
    Next i
    If lSlot = 0 Then
        ' if this explosion doesn't exists yet
        lSlot = UBound(Explosions) + 1
        ReDim Preserve Explosions(1 To lSlot)
        
        ' add a sprite for it
        Sprites.Add "explosion", "exp" & lSlot
        With Sprites("exp" & lSlot)
            .CreateAnimation True, 1, 8, 0, 0.05 ' create the explosion animation
        End With
    End If
    
    ' set some parameters
    With Sprites("exp" & lSlot)
        .SetPosition vec2(x, y)
        .active = True ' activate it
        .SetCurrentFrame 1
    End With
    Explosions(i).active = True
    Explosions(i).spriteindex = lSlot
End Sub

Private Sub HandleExplosions()
    Dim i&
    
    ' draws all explosions
    For i = 1 To UBound(Explosions)
        If Explosions(i).active Then
            Sprites("exp" & Explosions(i).spriteindex).Render
            ' check if the animation is over
            If Sprites("exp" & Explosions(i).spriteindex).GetCurrentFrame >= 7 Then
                Explosions(i).active = False
                Sprites("exp" & Explosions(i).spriteindex).active = False
            End If
        End If
    Next i
End Sub

Private Sub cmdAddSprite_Click()
    Dim i&
    Dim lSlot&
    ' find open slot:
    For i = 1 To Sprites.Count
        If Sprites(i).active = False Then
            If Left(Sprites(i).Key, 6) = "sprite" Then
                lSlot = CLng(Right(Sprites(i).Key, Len(Sprites(i).Key) - 6))
                Debug.Print Sprites(i).Key
                Exit For
            End If
        End If
    Next i
    
    If lSlot = 0 Then
        ' Add a sprite
        lSpriteNumber = lSpriteNumber + 1
        Sprites.Add "animation", "sprite" & lSpriteNumber
        lSlot = lSpriteNumber
    End If

    
    ' Add new one
    With Sprites("sprite" & lSlot)
        .CreateAnimation True, 6, 5, 0, 0.01  ' updateinverval = 10 miliseconds
        .SetPosition vec2((picScreen.ScaleWidth - 64) * Rnd, (picScreen.ScaleHeight - 64) * Rnd)
        .SetCustomVector1 -3 + 6 * Rnd, -3 + 6 * Rnd
        .SetColor rgba(scrColor(0).Value, scrColor(1).Value, scrColor(2).Value, scrAlpha.Value)
        .active = True
    End With
    
    ' If we must select a random color
    If chkRandom.Value = 1 Then
        scrColor(0).Value = 255 * Rnd
        scrColor(1).Value = 255 * Rnd
        scrColor(2).Value = 255 * Rnd
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim test As Long

    Randomize

    Show
    DoEvents
    
    ' Initialize the engine
    Engine.Initialize picScreen.hWnd, True
    
    ' create the device
    Set Device = Engine.CreateDevice(0) ' leave flags zero
    
    ' create the texturepool
    Set Textures = Engine.GetTexturePool
    
    ' create the spritepool
    Set Sprites = Engine.GetSpritePool
    
    ' That's all you need to do to initialize the engine!!
    
    ' Sprites are created from a texture, you can have multiple sprites that use the same texture
    ' so, when you create a sprite you must specify the name of the texture to use, we want to create
    ' sprites, so we must first load a texture
    Textures.AddFromFile App.Path & "\ani.bmp", "animation"
    Textures.AddFromFile App.Path & "\bigexp.bmp", "explosion", , , -16759893
    ' NOTE: Texture sizes must be powers of 2, else they will be resized
    
    ' now we can create a sprite from that texture
    Sprites.Add "animation", "sprite1"
    
    ' set the properties of the sprite
    With Sprites("sprite1")
        lSpriteNumber = 1
        ' we want to create an animation from it
        ' we set the the speed of the animation (updateinterval) the nubmer of rows and columns
        ' for the frames, when you set the updateinterval, the sprite will automaticcaly change
        ' frames when rendering.
        .CreateAnimation True, 6, 5, 0, 0.01  ' updateinverval = 10 miliseconds
        .SetPosition vec2(10, 10)

        .SetCustomVector1 1, 1
    End With
    
    ' initialize this array for the explosions
    ReDim Explosions(1 To 1)
    
    ' Now we'll enter the render loop
    Do Until bEnd
        DoEvents
        
        ' Always clear before rendering
        Engine.Clear
        
        ' Now, we'll draw all the sprites
        ' NOTE: Use this instead of RenderAll, this is usually faster
        ' NOTE: Set Sprites.Active = false to make a sprite invisible
        For i = 1 To Sprites.Count
            If Left(Sprites(i).Key, 6) = "sprite" Then
                ' check if the sprite goes off the screen:
                ' if so, bounce it
                If Sprites(i).GetPosition.x < 0 Then
                    Sprites(i).SetCustomVector1 -Sprites(i).GetCustomVector1.x, Sprites(i).GetCustomVector1.y
                ElseIf Sprites(i).GetPosition.x > picScreen.ScaleWidth - 40 Then
                    Sprites(i).SetCustomVector1 -Sprites(i).GetCustomVector1.x, Sprites(i).GetCustomVector1.y
                End If
                If Sprites(i).GetPosition.y < 0 Then
                    Sprites(i).SetCustomVector1 Sprites(i).GetCustomVector1.x, -Sprites(i).GetCustomVector1.y
                ElseIf Sprites(i).GetPosition.y > picScreen.ScaleHeight - 40 Then
                    Sprites(i).SetCustomVector1 Sprites(i).GetCustomVector1.x, -Sprites(i).GetCustomVector1.y
                End If
                
                ' check if the sprite hits another sprite
                For test = 1 To Sprites.Count
                    If Not test = i Then
                        If Intersect(Sprites(i), Sprites(test)) Then
                            If Sprites(test).active And Sprites(i).active Then
                                If Left(Sprites(test).Key, 6) = "sprite" And Left(Sprites(i).Key, 6) = "sprite" Then
                                    ' yes, we hit another sprite
                                    ' explode:
                                    AddExplosion Sprites(i).GetPosition.x, Sprites(i).GetPosition.y
                                    AddExplosion Sprites(test).GetPosition.x, Sprites(test).GetPosition.y
                                    
                                    Debug.Print Sprites(test).Key & " - " & Sprites(i).Key
                                    
                                    ' make these sprites inactive
                                    Sprites(i).active = False
                                    Sprites(test).active = False
                                End If
                            End If
                        End If
                    End If
                Next test
                If Sprites(i).active Then
                    ' make the sprite move
                    Sprites(i).Move Sprites(i).GetCustomVector1.x, Sprites(i).GetCustomVector1.y
                End If
                ' render it
                Sprites(i).Render
                
            End If
        Next i
        
        ' explosion stuff
        HandleExplosions
        
        ' Render everything to the screen, if you specify another hWnd, everything will be rendered
        ' to that target.
        Engine.Render
    Loop
    End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    bEnd = True
    Cancel = 1
End Sub

