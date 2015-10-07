Attribute VB_Name = "modDX8"
Option Explicit

Public Tex_Misc() As Long
Public Tex_Gui() As Long
Public Tex_Button_N() As Long
Public Tex_Button_C() As Long
Public Tex_Sprite() As Long
Public Tex_Tileset() As Long
Public Tex_PokeIcon() As Long
Public Tex_Background() As Long
Public Tex_FieldFront() As Long
Public Tex_FieldBack() As Long
Public Tex_TitleBar() As Long
Public Tex_PokeFront() As Long
Public Tex_PokeBack() As Long
Public Tex_ItemPic() As Long
Public Tex_PokeSprite() As Long

Public Count_Misc As Long
Public Count_Gui As Long
Public Count_Button_N As Long
Public Count_Button_C As Long
Public Count_Sprite As Long
Public Count_Tileset As Long
Public Count_PokeIcon As Long
Public Count_Background As Long
Public Count_FieldFront As Long
Public Count_FieldBack As Long
Public Count_TitleBar As Long
Public Count_PokeFront As Long
Public Count_PokeBack As Long
Public Count_ItemPic As Long
Public Count_PokeSprite As Long

Public Const Path_Misc As String = "\data\gfx\misc\"
Public Const Path_Gui As String = "\data\gfx\gui\"
Public Const Path_Button As String = "\data\gfx\gui\buttons\"
Public Const Path_Sprite As String = "\data\gfx\sprites\"
Public Const Path_Tileset As String = "\data\gfx\tilesets\"
Public Const Path_PokeIcon As String = "\data\gfx\pokemons\icons\"
Public Const Path_Background As String = "\data\gfx\backgrounds\"
Public Const Path_Field As String = "\data\gfx\fields\"
Public Const Path_TitleBar As String = "\data\gfx\titlebars\"
Public Const Path_PokeFront As String = "\data\gfx\pokemons\front\"
Public Const Path_PokeBack As String = "\data\gfx\pokemons\back\"
Public Const Path_ItemPic As String = "\data\gfx\items\"
Public Const Path_PokeSprite As String = "\data\gfx\pokemons\sprite\"

Public DX As DirectX8
Public D3D8 As Direct3D8
Public Direct3DX8 As D3DX8

Public D3DDevice8 As Direct3DDevice8
Public DispMode As D3DDISPLAYMODE
Public D3DWindow As D3DPRESENT_PARAMETERS

Public Type TLVERTEX
    X As Single
    Y As Single
    z As Single
    RHW As Single
    color As Long
    tu As Single
    tv As Single
End Type

Public gTexture() As TextureRec
Private Const TEXTURE_NULL As Long = 0

Private Type TextureRec
    Texture As Direct3DTexture8
    Width As Long
    Height As Long
    Path As String
    UnloadTimer As Long
    loaded As Boolean
End Type

Private Type D3DXIMAGE_INFO_A
    Width As Long
    Height As Long
    Depth As Long
    MipLevels As Long
    Format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
End Type

Public Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
Public Const FVF_Size As Long = 28

Private mTextureNum As Long
Public CurrentTexture As Long
Public Const ScreenWidth As Long = 800
Public Const ScreenHeight As Long = 608

Public Sub EngineInit()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set DX = New DirectX8
    Set D3D8 = DX.Direct3DCreate()
    Set Direct3DX8 = New D3DX8
    
    If Not EngineInitD3DDevice(D3DCREATE_PUREDEVICE Or D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
        If Not EngineInitD3DDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
            If Not EngineInitD3DDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then
                If Not EngineInitD3DDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then
                    Call MsgBox("Could not init D3DDevice8. Exiting...")
                    Call EngineUnloadDirectX
                    End
                End If
            End If
        End If
    End If

    Call EngineCacheTextures
    Call EngineInitRenderStates
    
    Call EngineInitFontTextures
    Call EngineInitFontSettings
    Call EngineInitFontColors
    
    Exit Sub
errHandler:
    HandleError "EngineInit", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Function EngineInitRenderStates()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    With D3DDevice8
        .SetVertexShader FVF
    
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
    End With
    
    Exit Function
errHandler:
    HandleError "EngineInitRenderStates", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Private Function EngineInitD3DDevice(D3DCREATEFLAGS As CONST_D3DCREATEFLAGS) As Boolean
    On Error GoTo ERRORMSG
    
    D3D8.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
    DispMode.Format = D3DFMT_X8R8G8B8
    D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
    DispMode.Width = ScreenWidth
    DispMode.Height = ScreenHeight
    D3DWindow.BackBufferCount = 1
    D3DWindow.BackBufferFormat = DispMode.Format
    D3DWindow.BackBufferWidth = ScreenWidth
    D3DWindow.BackBufferHeight = ScreenHeight
    D3DWindow.hDeviceWindow = frmMain.hWnd
    D3DWindow.Windowed = True

    If Not D3DDevice8 Is Nothing Then Set D3DDevice8 = Nothing
    Set D3DDevice8 = D3D8.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hWnd, D3DCREATEFLAGS, D3DWindow)
    
    EngineInitD3DDevice = True
    
    Exit Function
ERRORMSG:
    Set D3DDevice8 = Nothing
    EngineInitD3DDevice = False
End Function

Public Sub EngineUnloadDirectX()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Not D3DDevice8 Is Nothing Then Set D3DDevice8 = Nothing
    If Not D3D8 Is Nothing Then Set D3D8 = Nothing

    For i = 1 To mTextureNum
        Set gTexture(i).Texture = Nothing
    Next

    If Not DX Is Nothing Then Set DX = Nothing
    
    Exit Sub
errHandler:
    HandleError "EngineUnloadDirectX", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Function SetTexturePath(ByVal Path As String) As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    mTextureNum = mTextureNum + 1
    ReDim Preserve gTexture(0 To mTextureNum) As TextureRec
    gTexture(mTextureNum).Path = Path
    SetTexturePath = mTextureNum
    gTexture(mTextureNum).loaded = False
    
    Exit Function
errHandler:
    HandleError "SetTexturePath", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub LoadTexture(ByVal TextureNum As Long)
Dim Tex_Info As D3DXIMAGE_INFO_A
Dim Path As String

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Path = gTexture(TextureNum).Path
    
    Select Case gTexture(TextureNum).Width
        Case 0
            Set gTexture(TextureNum).Texture = Direct3DX8.CreateTextureFromFileEx(D3DDevice8, Path, D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, RGB(255, 0, 255), Tex_Info, ByVal 0)
            gTexture(TextureNum).Height = Tex_Info.Height
            gTexture(TextureNum).Width = Tex_Info.Width
        Case Is > 0
            Set gTexture(TextureNum).Texture = Direct3DX8.CreateTextureFromFileEx(D3DDevice8, Path, gTexture(TextureNum).Width, gTexture(TextureNum).Height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, RGB(255, 0, 255), ByVal 0, ByVal 0)
    End Select
    
    gTexture(TextureNum).UnloadTimer = GetTickCount
    gTexture(TextureNum).loaded = True
    
    Exit Sub
errHandler:
    HandleError "LoadTexture", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub UnloadTextures()
Dim Count As Long
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If mTextureNum <= 0 Then Exit Sub
    Count = UBound(gTexture)
    If Count <= 0 Then Exit Sub
    
    For i = 1 To Count
        With gTexture(i)
            If .UnloadTimer > GetTickCount + 150000 Then
                Set .Texture = Nothing
                Call ZeroMemory(ByVal VarPtr(gTexture(i)), LenB(gTexture(i)))
                .UnloadTimer = 0
                .loaded = False
            End If
        End With
    Next
    
    Exit Sub
errHandler:
    HandleError "UnloadTextures", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SetTexture(ByVal Texture As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Texture <> CurrentTexture Then
        If Texture > UBound(gTexture) Then Texture = UBound(gTexture)
        If Texture < 0 Then Texture = 0
        
        If Not Texture = TEXTURE_NULL Then
            If Not gTexture(Texture).loaded Then
                Call LoadTexture(Texture)
            End If
        End If
        
        Call D3DDevice8.SetTexture(0, gTexture(Texture).Texture)
        CurrentTexture = Texture
    End If
    
    Exit Sub
errHandler:
    HandleError "SetTexture", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub EngineCacheTextures()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Count_Misc = 1
    Do While FileExist(App.Path & Path_Misc & Count_Misc & Gfx_Ext)
        Count_Misc = Count_Misc + 1
    Loop
    Count_Misc = Count_Misc - 1
    If Count_Misc > 0 Then
        ReDim Tex_Misc(0 To Count_Misc)
        For i = 1 To Count_Misc
            Tex_Misc(i) = SetTexturePath(App.Path & Path_Misc & i & Gfx_Ext)
        Next
    End If
    
    Count_Gui = 1
    Do While FileExist(App.Path & Path_Gui & Count_Gui & Gfx_Ext)
        Count_Gui = Count_Gui + 1
    Loop
    Count_Gui = Count_Gui - 1
    If Count_Gui > 0 Then
        ReDim Tex_Gui(0 To Count_Gui)
        For i = 1 To Count_Gui
            Tex_Gui(i) = SetTexturePath(App.Path & Path_Gui & i & Gfx_Ext)
        Next
    End If
    
    Count_Button_N = 1
    Do While FileExist(App.Path & Path_Button & Count_Button_N & "_n" & Gfx_Ext)
        Count_Button_N = Count_Button_N + 1
    Loop
    Count_Button_N = Count_Button_N - 1
    If Count_Button_N > 0 Then
        ReDim Tex_Button_N(0 To Count_Button_N)
        For i = 1 To Count_Button_N
            Tex_Button_N(i) = SetTexturePath(App.Path & Path_Button & i & "_n" & Gfx_Ext)
        Next
    End If
    
    Count_Button_C = 1
    Do While FileExist(App.Path & Path_Button & Count_Button_C & "_c" & Gfx_Ext)
        Count_Button_C = Count_Button_C + 1
    Loop
    Count_Button_C = Count_Button_C - 1
    If Count_Button_C > 0 Then
        ReDim Tex_Button_C(0 To Count_Button_C)
        For i = 1 To Count_Button_C
            Tex_Button_C(i) = SetTexturePath(App.Path & Path_Button & i & "_c" & Gfx_Ext)
        Next
    End If
    
    Count_Sprite = 1
    Do While FileExist(App.Path & Path_Sprite & Count_Sprite & Gfx_Ext)
        Count_Sprite = Count_Sprite + 1
    Loop
    Count_Sprite = Count_Sprite - 1
    If Count_Sprite > 0 Then
        ReDim Tex_Sprite(0 To Count_Sprite)
        For i = 1 To Count_Sprite
            Tex_Sprite(i) = SetTexturePath(App.Path & Path_Sprite & i & Gfx_Ext)
        Next
    End If
    
    Count_Tileset = 1
    Do While FileExist(App.Path & Path_Tileset & Count_Tileset & Gfx_Ext)
        Count_Tileset = Count_Tileset + 1
    Loop
    Count_Tileset = Count_Tileset - 1
    If Count_Tileset > 0 Then
        ReDim Tex_Tileset(0 To Count_Tileset)
        For i = 1 To Count_Tileset
            Tex_Tileset(i) = SetTexturePath(App.Path & Path_Tileset & i & Gfx_Ext)
        Next
    End If
    
    Count_PokeIcon = 1
    Do While FileExist(App.Path & Path_PokeIcon & Count_PokeIcon & Gfx_Ext)
        Count_PokeIcon = Count_PokeIcon + 1
    Loop
    Count_PokeIcon = Count_PokeIcon - 1
    If Count_PokeIcon > 0 Then
        ReDim Tex_PokeIcon(0 To Count_PokeIcon)
        For i = 1 To Count_PokeIcon
            Tex_PokeIcon(i) = SetTexturePath(App.Path & Path_PokeIcon & i & Gfx_Ext)
        Next
    End If
    
    Count_Background = 1
    Do While FileExist(App.Path & Path_Background & Count_Background & Gfx_Ext)
        Count_Background = Count_Background + 1
    Loop
    Count_Background = Count_Background - 1
    If Count_Background > 0 Then
        ReDim Tex_Background(0 To Count_Background)
        For i = 1 To Count_Background
            Tex_Background(i) = SetTexturePath(App.Path & Path_Background & i & Gfx_Ext)
        Next
    End If
    
    Count_FieldFront = 1
    Do While FileExist(App.Path & Path_Field & Count_FieldFront & "_f" & Gfx_Ext)
        Count_FieldFront = Count_FieldFront + 1
    Loop
    Count_FieldFront = Count_FieldFront - 1
    If Count_FieldFront > 0 Then
        ReDim Tex_FieldFront(0 To Count_FieldFront)
        For i = 1 To Count_FieldFront
            Tex_FieldFront(i) = SetTexturePath(App.Path & Path_Field & i & "_f" & Gfx_Ext)
        Next
    End If
    
    Count_FieldBack = 1
    Do While FileExist(App.Path & Path_Field & Count_FieldBack & "_b" & Gfx_Ext)
        Count_FieldBack = Count_FieldBack + 1
    Loop
    Count_FieldBack = Count_FieldBack - 1
    If Count_FieldBack > 0 Then
        ReDim Tex_FieldBack(0 To Count_FieldBack)
        For i = 1 To Count_FieldBack
            Tex_FieldBack(i) = SetTexturePath(App.Path & Path_Field & i & "_b" & Gfx_Ext)
        Next
    End If
    
    Count_TitleBar = 0
    Do While FileExist(App.Path & Path_TitleBar & Count_TitleBar & Gfx_Ext)
        Count_TitleBar = Count_TitleBar + 1
    Loop
    Count_TitleBar = Count_TitleBar - 1
    If Count_TitleBar > 0 Then
        ReDim Tex_TitleBar(0 To Count_TitleBar)
        For i = 0 To Count_TitleBar
            Tex_TitleBar(i) = SetTexturePath(App.Path & Path_TitleBar & i & Gfx_Ext)
        Next
    End If
    
    Count_PokeFront = 1
    Do While FileExist(App.Path & Path_PokeFront & Count_PokeFront & Gfx_Ext)
        Count_PokeFront = Count_PokeFront + 1
    Loop
    Count_PokeFront = Count_PokeFront - 1
    If Count_PokeFront > 0 Then
        ReDim Tex_PokeFront(0 To Count_PokeFront)
        For i = 1 To Count_PokeFront
            Tex_PokeFront(i) = SetTexturePath(App.Path & Path_PokeFront & i & Gfx_Ext)
        Next
    End If
    
    Count_PokeBack = 1
    Do While FileExist(App.Path & Path_PokeBack & Count_PokeBack & Gfx_Ext)
        Count_PokeBack = Count_PokeBack + 1
    Loop
    Count_PokeBack = Count_PokeBack - 1
    If Count_PokeBack > 0 Then
        ReDim Tex_PokeBack(0 To Count_PokeBack)
        For i = 1 To Count_PokeBack
            Tex_PokeBack(i) = SetTexturePath(App.Path & Path_PokeBack & i & Gfx_Ext)
        Next
    End If
    
    Count_ItemPic = 1
    Do While FileExist(App.Path & Path_ItemPic & Count_ItemPic & Gfx_Ext)
        Count_ItemPic = Count_ItemPic + 1
    Loop
    Count_ItemPic = Count_ItemPic - 1
    If Count_ItemPic > 0 Then
        ReDim Tex_ItemPic(0 To Count_ItemPic)
        For i = 1 To Count_ItemPic
            Tex_ItemPic(i) = SetTexturePath(App.Path & Path_ItemPic & i & Gfx_Ext)
        Next
    End If
    
    Count_PokeSprite = 1
    Do While FileExist(App.Path & Path_PokeSprite & Count_PokeSprite & Gfx_Ext)
        Count_PokeSprite = Count_PokeSprite + 1
    Loop
    Count_PokeSprite = Count_PokeSprite - 1
    If Count_PokeSprite > 0 Then
        ReDim Tex_PokeSprite(0 To Count_PokeSprite)
        For i = 1 To Count_PokeSprite
            Tex_PokeSprite(i) = SetTexturePath(App.Path & Path_PokeSprite & i & Gfx_Ext)
        Next
    End If
    
    Exit Sub
errHandler:
    HandleError "EngineCacheTextures", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub RenderTexture(ByVal Texture As Long, ByVal DX As Long, ByVal dY As Long, ByVal sX As Long, ByVal sY As Long, ByVal dW As Long, ByVal dH As Long, ByVal sW As Long, ByVal sH As Long, Optional ByVal colour As Long = -1)
Dim Box(0 To 3) As TLVERTEX
Dim i As Long
Dim textureWidth As Long, textureHeight As Long
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Call SetTexture(Texture)
    
    textureWidth = gTexture(Texture).Width
    textureHeight = gTexture(Texture).Height
    
    If Texture <= 0 Or textureWidth <= 0 Or textureHeight <= 0 Then Exit Sub
    
    For i = 0 To 3
        Box(i).RHW = 1
        Box(i).color = colour
    Next

    Box(0).X = DX: Box(0).Y = dY: Box(0).tu = ((sX + 0.5) / textureWidth): Box(0).tv = ((sY + 0.5) / textureHeight)
    Box(1).X = DX + dW: Box(1).tu = (sX + sW + 1) / textureWidth
    Box(2).X = Box(0).X
    Box(3).X = Box(1).X

    Box(2).Y = dY + dH: Box(2).tv = (sY + sH + 1) / textureHeight

    Box(1).Y = Box(0).Y: Box(1).tv = Box(0).tv
    Box(2).tu = Box(0).tu
    Box(3).Y = Box(2).Y: Box(3).tu = Box(1).tu: Box(3).tv = Box(2).tv
    
    Call D3DDevice8.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, Box(0), FVF_Size)
    gTexture(Texture).UnloadTimer = GetTickCount
    
    Exit Sub
errHandler:
    HandleError "RenderTexture", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub RenderTextureByRects(ByVal TextureRec As Long, sRect As RECT, dRect As RECT)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    RenderTexture TextureRec, dRect.Left, dRect.top, sRect.Left, sRect.top, dRect.Right - dRect.Left, dRect.bottom - dRect.top, sRect.Right - sRect.Left, sRect.bottom - sRect.top, D3DColorRGBA(255, 255, 255, 255)

    Exit Sub
errHandler:
    HandleError "RenderTextureByRects", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function GetTextureWidth(ByVal TextureRec As Long) As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    GetTextureWidth = gTexture(TextureRec).Width
    
    Exit Function
errHandler:
    HandleError "GetTextureWidth", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function GetTextureHeight(ByVal TextureRec As Long) As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    GetTextureHeight = gTexture(TextureRec).Height
    
    Exit Function
errHandler:
    HandleError "GetTextureHeight", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub Render_Graphics()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If D3DDevice8.TestCooperativeLevel <> D3D_OK Then
        If D3DDevice8.TestCooperativeLevel = D3DERR_DEVICELOST Then Exit Sub
        Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
        Call D3DDevice8.Reset(D3DWindow)
        Call EngineInitRenderStates
    End If
    
    UnloadTextures
    D3DDevice8.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    D3DDevice8.BeginScene
    
    If InMenu Then Render_Menu
    If InGame Then Render_Game
    
    DrawCursor
    
    D3DDevice8.EndScene
    D3DDevice8.Present ByVal 0, ByVal 0, 0, ByVal 0
    
    If frmMapEditor.Visible Then Render_Tileset
    If frmPokemonEditor.Visible Then Editor_DrawPokeIcon
    If frmItemEditor.Visible Then Editor_DrawItemPic
    
    Exit Sub
errHandler:
    HandleError "Render_Graphics", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub Render_Menu()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Not InMenu Then Exit Sub
    
    RenderTexture Tex_Gui(GuiBackground), 0, 0, 0, 0, ScreenWidth, ScreenHeight, GetTextureWidth(Tex_Gui(GuiBackground)), GetTextureHeight(Tex_Gui(GuiBackground))
    RenderTexture Tex_Gui(GuiLogo), (ScreenWidth / 2) - (GetTextureWidth(Tex_Gui(GuiLogo)) / 2), 100, 0, 0, GetTextureWidth(Tex_Gui(GuiLogo)), GetTextureHeight(Tex_Gui(GuiLogo)), GetTextureWidth(Tex_Gui(GuiLogo)), GetTextureHeight(Tex_Gui(GuiLogo))
    
    DrawLogin
    DrawRegister
    DrawCharSelect
    DrawCharCreate
    
    Exit Sub
errHandler:
    HandleError "Render_Menu", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub UpdateCamera()
Dim offsetX As Long, offsetY As Long
Dim StartX As Long, StartY As Long
Dim EndX As Long, EndY As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    offsetX = Player(MyIndex).xOffset + Pic_Size
    offsetY = Player(MyIndex).yOffset + Pic_Size
    StartX = GetPlayerX(MyIndex) - ((Max_MapX + 1) \ 2) - 1
    StartY = GetPlayerY(MyIndex) - ((Max_MapY + 1) \ 2) - 1

    If StartX < 0 Then
        offsetX = 0

        If StartX = -1 Then
            If Player(MyIndex).xOffset > 0 Then
                offsetX = Player(MyIndex).xOffset
            End If
        End If

        StartX = 0
    End If
    If StartY < 0 Then
        offsetY = 0

        If StartY = -1 Then
            If Player(MyIndex).yOffset > 0 Then
                offsetY = Player(MyIndex).yOffset
            End If
        End If

        StartY = 0
    End If

    EndX = StartX + (Max_MapX + 1) + 1
    EndY = StartY + (Max_MapY + 1) + 1
    If EndX > Map.MaxX Then
        offsetX = 32

        If EndX = Map.MaxX + 1 Then
            If Player(MyIndex).xOffset < 0 Then
                offsetX = Player(MyIndex).xOffset + Pic_Size
            End If
        End If

        EndX = Map.MaxX
        StartX = EndX - Max_MapX - 1
    End If
    If EndY > Map.MaxY Then
        offsetY = 32

        If EndY = Map.MaxY + 1 Then
            If Player(MyIndex).yOffset < 0 Then
                offsetY = Player(MyIndex).yOffset + Pic_Size
            End If
        End If

        EndY = Map.MaxY
        StartY = EndY - Max_MapY - 1
    End If

    'With TileView
    '    .Top = StartY
    '    .bottom = EndY
    '    .Left = StartX
    '    .Right = EndX
    'End With
    
    With TileView
        .top = GetPlayerY(MyIndex) - StartYValue
        .bottom = .top + EndYValue
        .Left = GetPlayerX(MyIndex) - StartXValue
        .Right = .Left + EndXValue
    End With

    'With Camera
    '    .Top = offsetY
    '    .bottom = .Top + ScreenY
    '    .Left = offsetX
    '    .Right = .Left + ScreenX
    'End With
    
    With Camera
        .top = Player(MyIndex).yOffset + Pic_Size
        .bottom = .top + ScreenY
        .Left = Player(MyIndex).xOffset + Pic_Size
        .Right = .Left + ScreenX
    End With
    
    CurX = TileView.Left + ((GlobalX + Camera.Left) \ Pic_Size)
    CurY = TileView.top + ((GlobalY + Camera.top) \ Pic_Size)
    GlobalX_Map = GlobalX + (TileView.Left * Pic_Size) + Camera.Left
    GlobalY_Map = GlobalY + (TileView.top * Pic_Size) + Camera.top

    Exit Sub
errHandler:
    HandleError "UpdateCamera", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub Render_Game()
Dim X As Long, Y As Long
Dim i As Long
    
    UpdateCamera
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Editor = EDITOR_MAP Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.top To TileView.bottom
                If IsValidMapPoint(X, Y) Then
                    RenderTexture Tex_Misc(MiscAlpha), ConvertMapX(X * 32), ConvertMapY(Y * 32), 0, 0, 32, 32, 16, 16
                End If
            Next Y
        Next X
    End If
    If Count_Tileset > 0 Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.top To TileView.bottom
                For i = Layers.Ground To Layers.Mask2
                    DrawMapTile i, X, Y, ConvertMapX(X * 32), ConvertMapY(Y * 32)
                Next i
            Next Y
        Next X
    End If
    
    If Count_Sprite > 0 Then
        For Y = 0 To Map.MaxY
            For i = 1 To HighPlayerIndex
                If IsPlaying(i) Then
                    If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                        If GetPlayerY(i) = Y Then
                            DrawPlayer i
                        End If
                    End If
                End If
            Next
            
            For i = 1 To MAX_MAP_NPC
                If MapNpc(i).Num > 0 Then
                    If MapNpc(i).Y = Y Then
                        Call DrawNpc(i)
                    End If
                End If
            Next
        Next
    End If
    
    If Count_Tileset > 0 Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.top To TileView.bottom
                For i = Layers.Fringe To Layers.Fringe2
                    DrawMapTile i, X, Y, ConvertMapX(X * 32), ConvertMapY(Y * 32)
                Next i
            Next Y
        Next X
    End If
    
    For i = 1 To HighPlayerIndex
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            If Player(i).InBattle > 0 Then
                DrawInBattleIcon i
            End If
        End If
    Next
    
    If MyTarget > 0 Then
        DrawTarget MyTarget
    End If
    For i = 1 To HighPlayerIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) > 0 Then
                If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    DrawPlayerName i
                End If
            End If
        End If
    Next
    For i = 1 To MAX_MAP_NPC
        If MapNpc(i).Num > 0 Then
            DrawNpcName i
        End If
    Next
    
    If Editor = EDITOR_MAP Then
        DrawAttribute
        RenderTexture Tex_Misc(MiscBlank), ConvertMapX(CurX * Pic_Size), ConvertMapY(CurY * Pic_Size), 0, 0, Pic_Size, Pic_Size, 1, 1, D3DColorARGB(120, 0, 0, 0)
        RenderText CurFont, "Map: " & Player(MyIndex).Map, 10, 10, White
        RenderText CurFont, "Player X: " & Player(MyIndex).X, 10, 25, White
        RenderText CurFont, "Player Y: " & Player(MyIndex).Y, 10, 40, White
        RenderText CurFont, "Cursor X: " & CurX, 10, 55, White
        RenderText CurFont, "Cursor Y: " & CurY, 10, 70, White
    Else
        DrawGui
    End If
    
    Exit Sub
errHandler:
    HandleError "Render_Game", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub Render_Tileset()
Dim desRect As D3DRECT
Dim sRect As RECT, dRect As RECT
Dim Height As Long, Width As Long
Dim scrlX As Long, scrlY As Long
Dim X As Long, Y As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    scrlX = frmMapEditor.scrlPictureX.value * Pic_Size: scrlY = frmMapEditor.scrlPictureY.value * Pic_Size
    Width = GetTextureWidth(Tex_Tileset(CurTileset)) - scrlX: Height = GetTextureHeight(Tex_Tileset(CurTileset)) - scrlY
    
    With sRect
        .Left = scrlX
        .top = scrlY
        .Right = .Left + Width
        .bottom = .top + Height
    End With
    With dRect
        .bottom = Height
        .Right = Width
    End With

    With desRect
        .x2 = frmMapEditor.picTileset.ScaleWidth
        .y2 = frmMapEditor.picTileset.ScaleHeight
    End With

    D3DDevice8.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    D3DDevice8.BeginScene
    
    For X = 0 To (frmMapEditor.picTileset.ScaleWidth \ 16)
        For Y = 0 To (frmMapEditor.picTileset.ScaleHeight \ 16)
            RenderTexture Tex_Misc(MiscAlpha), X * 16, Y * 16, 0, 0, 16, 16, 16, 16
        Next
    Next
    
    RenderTextureByRects Tex_Tileset(CurTileset), sRect, dRect
    RenderTexture Tex_Misc(MiscBlank), (EditorTileX * Pic_Size) - scrlX, (EditorTileY * Pic_Size) - scrlY, 0, 0, EditorTileWidth * Pic_Size, EditorTileHeight * Pic_Size, 1, 1, D3DColorARGB(120, 0, 0, 0)
    
    D3DDevice8.EndScene
    D3DDevice8.Present desRect, desRect, frmMapEditor.picTileset.hWnd, ByVal 0
    
    Exit Sub
errHandler:
    HandleError "Render_Tileset", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawButton(ByVal ButtonNum As Byte, Optional ByVal Text As String, Optional ByVal tColor As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Not Buttons(ButtonNum).Visible Then Exit Sub
    
    With Buttons(ButtonNum)
        Select Case .bState
            Case ButtonNormal
                RenderTexture Tex_Button_N(.Pic), .X, .Y, 0, 0, GetTextureWidth(Tex_Button_N(.Pic)), GetTextureHeight(Tex_Button_N(.Pic)), GetTextureWidth(Tex_Button_N(.Pic)), GetTextureHeight(Tex_Button_N(.Pic))
                If Len(Text) > 0 Then RenderText CurFont, Text, .X + ((GetTextureWidth(Tex_Button_N(.Pic)) / 2) - (GetWidth(CurFont, Text) / 2)), .Y + ((GetTextureHeight(Tex_Button_N(.Pic)) / 2) - 8), tColor
            Case ButtonClick
                RenderTexture Tex_Button_C(.Pic), .X, .Y, 0, 0, GetTextureWidth(Tex_Button_C(.Pic)), GetTextureHeight(Tex_Button_C(.Pic)), GetTextureWidth(Tex_Button_C(.Pic)), GetTextureHeight(Tex_Button_C(.Pic))
                If Len(Text) > 0 Then RenderText CurFont, Text, .X + ((GetTextureWidth(Tex_Button_C(.Pic)) / 2) - (GetWidth(CurFont, Text) / 2)), .Y + ((GetTextureHeight(Tex_Button_C(.Pic)) / 2) - 8), tColor
        End Select
    End With
    
    Exit Sub
errHandler:
    HandleError "DrawButton", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawLogin()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Not WindowVisible(WindowType.Menu_Login) Then Exit Sub
    
    RenderTexture Tex_Gui(GuiLogin), GuiLoginX, GuiLoginY, 0, 0, GetTextureWidth(Tex_Gui(GuiLogin)), GetTextureHeight(Tex_Gui(GuiLogin)), GetTextureWidth(Tex_Gui(GuiLogin)), GetTextureHeight(Tex_Gui(GuiLogin)), D3DColorARGB(130, 255, 255, 255)
    
    RenderText CurFont, "Login Window", GuiLoginX + 10, GuiLoginY + 5, White
    RenderText CurFont, "Username:", GuiLoginX + 40, GuiLoginY + 37, White
    RenderText CurFont, "Password:", GuiLoginX + 40, GuiLoginY + 62, White
    RenderText CurFont, "Save Password?", GuiLoginX + 130, GuiLoginY + 88, White
    
    If SaveAccount Then RenderTexture Tex_Misc(MiscCheck), GuiLoginX + 100, GuiLoginY + 90, 0, 0, 11, 11, 11, 11
    
    If CurTextBox = 1 Then
        RenderText CurFont, user & ChatLine, GuiLoginX + 130, GuiLoginY + 37, White
        RenderText CurFont, CensorWord(Pass), GuiLoginX + 130, GuiLoginY + 62, White
    ElseIf CurTextBox = 2 Then
        RenderText CurFont, user, GuiLoginX + 130, GuiLoginY + 37, White
        RenderText CurFont, CensorWord(Pass) & ChatLine, GuiLoginX + 130, GuiLoginY + 62, White
    Else
        RenderText CurFont, user, GuiLoginX + 130, GuiLoginY + 37, White
        RenderText CurFont, CensorWord(Pass), GuiLoginX + 130, GuiLoginY + 62, White
    End If
    
    DrawButton ButtonEnum.LoginAccept, "Accept", White
    DrawButton ButtonEnum.Register, "Register", White
    
    Exit Sub
errHandler:
    HandleError "DrawLogin", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawRegister()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Not WindowVisible(WindowType.Menu_Register) Then Exit Sub
    
    RenderTexture Tex_Gui(GuiRegister), GuiRegisterX, GuiRegisterY, 0, 0, GetTextureWidth(Tex_Gui(GuiRegister)), GetTextureHeight(Tex_Gui(GuiRegister)), GetTextureWidth(Tex_Gui(GuiRegister)), GetTextureHeight(Tex_Gui(GuiRegister)), D3DColorARGB(130, 255, 255, 255)
    
    RenderText CurFont, "Register Window", GuiRegisterX + 10, GuiRegisterY + 5, White
    RenderText CurFont, "Username:", GuiRegisterX + 40, GuiRegisterY + 37, White
    RenderText CurFont, "Password:", GuiRegisterX + 40, GuiRegisterY + 62, White
    RenderText CurFont, "Retype:", GuiRegisterX + 40, GuiRegisterY + 87, White
    
    If CurTextBox = 1 Then
        RenderText CurFont, user & ChatLine, GuiRegisterX + 130, GuiRegisterY + 37, White
        RenderText CurFont, CensorWord(Pass), GuiRegisterX + 130, GuiRegisterY + 62, White
        RenderText CurFont, CensorWord(Pass2), GuiRegisterX + 130, GuiRegisterY + 87, White
    ElseIf CurTextBox = 2 Then
        RenderText CurFont, user, GuiRegisterX + 130, GuiRegisterY + 37, White
        RenderText CurFont, CensorWord(Pass) & ChatLine, GuiRegisterX + 130, GuiRegisterY + 62, White
        RenderText CurFont, CensorWord(Pass2), GuiRegisterX + 130, GuiRegisterY + 87, White
    ElseIf CurTextBox = 3 Then
        RenderText CurFont, user, GuiRegisterX + 130, GuiRegisterY + 37, White
        RenderText CurFont, CensorWord(Pass), GuiRegisterX + 130, GuiRegisterY + 62, White
        RenderText CurFont, CensorWord(Pass2) & ChatLine, GuiRegisterX + 130, GuiRegisterY + 87, White
    Else
        RenderText CurFont, user, GuiRegisterX + 130, GuiRegisterY + 37, White
        RenderText CurFont, CensorWord(Pass), GuiRegisterX + 130, GuiRegisterY + 62, White
        RenderText CurFont, CensorWord(Pass2), GuiRegisterX + 130, GuiRegisterY + 87, White
    End If
    
    DrawButton ButtonEnum.RegisterAccept, "Accept", White
    DrawButton ButtonEnum.Login, "Login", White
    
    Exit Sub
errHandler:
    HandleError "DrawRegister", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawCharSelect()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Not WindowVisible(WindowType.Menu_CharSelect) Then Exit Sub
    
    RenderTexture Tex_Gui(GuiCharSelect), GuiCharSelectX, GuiCharSelectY, 0, 0, GetTextureWidth(Tex_Gui(GuiCharSelect)), GetTextureHeight(Tex_Gui(GuiCharSelect)), GetTextureWidth(Tex_Gui(GuiCharSelect)), GetTextureHeight(Tex_Gui(GuiCharSelect)), D3DColorARGB(130, 255, 255, 255)

    RenderText CurFont, "Character Selection", GuiCharSelectX + 10, GuiCharSelectY + 5, White

    DrawButton ButtonEnum.CharNew1, "New", White
    DrawButton ButtonEnum.CharNew2, "New", White
    DrawButton ButtonEnum.CharNew3, "New", White
    
    DrawButton ButtonEnum.CharUse1
    DrawButton ButtonEnum.CharUse2
    DrawButton ButtonEnum.CharUse3
    If CharSelectSprite(1) > 0 Then RenderTexture Tex_Sprite(CharSelectSprite(1)), GuiCharSelectX + 23, GuiCharSelectY + 51, (GetTextureWidth(Tex_Sprite(CharSelectSprite(1))) / 3), 0, GetTextureWidth(Tex_Sprite(CharSelectSprite(1))) / 3, GetTextureHeight(Tex_Sprite(CharSelectSprite(1))) / 4, GetTextureWidth(Tex_Sprite(CharSelectSprite(1))) / 3, GetTextureHeight(Tex_Sprite(CharSelectSprite(1))) / 4
    If CharSelectSprite(2) > 0 Then RenderTexture Tex_Sprite(CharSelectSprite(2)), GuiCharSelectX + 85, GuiCharSelectY + 51, (GetTextureWidth(Tex_Sprite(CharSelectSprite(2))) / 3), 0, GetTextureWidth(Tex_Sprite(CharSelectSprite(2))) / 3, GetTextureHeight(Tex_Sprite(CharSelectSprite(2))) / 4, GetTextureWidth(Tex_Sprite(CharSelectSprite(2))) / 3, GetTextureHeight(Tex_Sprite(CharSelectSprite(2))) / 4
    If CharSelectSprite(3) > 0 Then RenderTexture Tex_Sprite(CharSelectSprite(3)), GuiCharSelectX + 147, GuiCharSelectY + 51, (GetTextureWidth(Tex_Sprite(CharSelectSprite(3))) / 3), 0, GetTextureWidth(Tex_Sprite(CharSelectSprite(3))) / 3, GetTextureHeight(Tex_Sprite(CharSelectSprite(3))) / 4, GetTextureWidth(Tex_Sprite(CharSelectSprite(3))) / 3, GetTextureHeight(Tex_Sprite(CharSelectSprite(3))) / 4

    DrawButton ButtonEnum.CharDel1, "Del", White
    DrawButton ButtonEnum.CharDel2, "Del", White
    DrawButton ButtonEnum.CharDel3, "Del", White
    
    Exit Sub
errHandler:
    HandleError "DrawCharSelect", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawCharCreate()
Dim mText As String

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Not WindowVisible(WindowType.Menu_CharCreate) Then Exit Sub
    
    RenderTexture Tex_Gui(GuiCharCreate), GuiCharCreateX, GuiCharCreateY, 0, 0, GetTextureWidth(Tex_Gui(GuiCharCreate)), GetTextureHeight(Tex_Gui(GuiCharCreate)), GetTextureWidth(Tex_Gui(GuiCharCreate)), GetTextureHeight(Tex_Gui(GuiCharCreate)), D3DColorARGB(130, 255, 255, 255)
    
    RenderText CurFont, "Character Creation", GuiCharCreateX + 10, GuiCharCreateY + 5, White
    
    RenderText CurFont, "Name:", GuiCharCreateX + 28, GuiCharCreateY + 37, White
    RenderText CurFont, user & ChatLine, GuiCharCreateX + 83, GuiCharCreateY + 37, White
    
    mText = "Starter"
    RenderText CurFont, mText, GuiCharCreateX + 130 + (86 / 2) - (GetWidth(CurFont, mText) / 2), GuiCharCreateY + 65, D3DColorARGB(170, 255, 255, 255)
    mText = "Bulbasaur"
    If SelStarter = 1 Then
        RenderText CurFont, mText, GuiCharCreateX + 130 + (86 / 2) - (GetWidth(CurFont, mText) / 2), GuiCharCreateY + 85, Yellow
    Else
        RenderText CurFont, mText, GuiCharCreateX + 130 + (86 / 2) - (GetWidth(CurFont, mText) / 2), GuiCharCreateY + 85, White
    End If
    mText = "Charmander"
    If SelStarter = 4 Then
        RenderText CurFont, mText, GuiCharCreateX + 130 + (86 / 2) - (GetWidth(CurFont, mText) / 2), GuiCharCreateY + 100, Yellow
    Else
        RenderText CurFont, mText, GuiCharCreateX + 130 + (86 / 2) - (GetWidth(CurFont, mText) / 2), GuiCharCreateY + 100, White
    End If
    mText = "Squirtle"
    If SelStarter = 7 Then
        RenderText CurFont, mText, GuiCharCreateX + 130 + (86 / 2) - (GetWidth(CurFont, mText) / 2), GuiCharCreateY + 115, Yellow
    Else
        RenderText CurFont, mText, GuiCharCreateX + 130 + (86 / 2) - (GetWidth(CurFont, mText) / 2), GuiCharCreateY + 115, White
    End If
    
    Select Case SelGender
        Case GENDER_MALE
            RenderTexture Tex_Button_C(4), GuiCharCreateX + 18, GuiCharCreateY + 72, 0, 0, GetTextureWidth(Tex_Button_C(4)), GetTextureHeight(Tex_Button_C(4)), GetTextureWidth(Tex_Button_C(4)), GetTextureHeight(Tex_Button_C(4))
            RenderTexture Tex_Button_N(4), GuiCharCreateX + 75, GuiCharCreateY + 72, 0, 0, GetTextureWidth(Tex_Button_N(4)), GetTextureHeight(Tex_Button_N(4)), GetTextureWidth(Tex_Button_N(4)), GetTextureHeight(Tex_Button_N(4))
            
            ' Temp Sprite Number
            RenderTexture Tex_Sprite(1), GuiCharCreateX + 24, GuiCharCreateY + 78, (GetTextureWidth(Tex_Sprite(1)) / 3) * GenAnim, 0, GetTextureWidth(Tex_Sprite(1)) / 3, GetTextureHeight(Tex_Sprite(1)) / 4, GetTextureWidth(Tex_Sprite(1)) / 3, GetTextureHeight(Tex_Sprite(1)) / 4
            RenderTexture Tex_Sprite(2), GuiCharCreateX + 80, GuiCharCreateY + 78, (GetTextureWidth(Tex_Sprite(1)) / 3), 0, GetTextureWidth(Tex_Sprite(1)) / 3, GetTextureHeight(Tex_Sprite(1)) / 4, GetTextureWidth(Tex_Sprite(1)) / 3, GetTextureHeight(Tex_Sprite(1)) / 4
        Case GENDER_FEMALE
            RenderTexture Tex_Button_N(4), GuiCharCreateX + 18, GuiCharCreateY + 72, 0, 0, GetTextureWidth(Tex_Button_N(4)), GetTextureHeight(Tex_Button_N(4)), GetTextureWidth(Tex_Button_N(4)), GetTextureHeight(Tex_Button_N(4))
            RenderTexture Tex_Button_C(4), GuiCharCreateX + 75, GuiCharCreateY + 72, 0, 0, GetTextureWidth(Tex_Button_C(4)), GetTextureHeight(Tex_Button_C(4)), GetTextureWidth(Tex_Button_C(4)), GetTextureHeight(Tex_Button_C(4))
            
            ' Temp Sprite Number
            RenderTexture Tex_Sprite(1), GuiCharCreateX + 24, GuiCharCreateY + 78, (GetTextureWidth(Tex_Sprite(1)) / 3), 0, GetTextureWidth(Tex_Sprite(1)) / 3, GetTextureHeight(Tex_Sprite(1)) / 4, GetTextureWidth(Tex_Sprite(1)) / 3, GetTextureHeight(Tex_Sprite(1)) / 4
            RenderTexture Tex_Sprite(2), GuiCharCreateX + 80, GuiCharCreateY + 78, (GetTextureWidth(Tex_Sprite(1)) / 3) * GenAnim, 0, GetTextureWidth(Tex_Sprite(1)) / 3, GetTextureHeight(Tex_Sprite(1)) / 4, GetTextureWidth(Tex_Sprite(1)) / 3, GetTextureHeight(Tex_Sprite(1)) / 4
    End Select
    
    DrawButton ButtonEnum.CharAccept, "Accept", White
    DrawButton ButtonEnum.CharDecline, "Back", White
    
    Exit Sub
errHandler:
    HandleError "DrawCharCreate", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawMapTile(ByVal Layers As Long, ByVal X As Long, ByVal Y As Long, ByVal posX As Long, ByVal posY As Long, Optional ByVal color As Long = -1)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If GettingMap Then Exit Sub
    
    If IsValidMapPoint(X, Y) Then
        With Map.Tile(X, Y)
            RenderTexture Tex_Tileset(.Layer(Layers).Tileset), posX, posY, .Layer(Layers).X * Pic_Size, .Layer(Layers).Y * Pic_Size, Pic_Size, Pic_Size, Pic_Size, Pic_Size, color
        End With
    End If
    
    Exit Sub
errHandler:
    HandleError "DrawMapTile", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawPlayer(ByVal Index As Long)
Dim Sprite As Long, anim As Long
Dim spritetop As Byte
Dim X As Long, Y As Long
Dim sRec As RECT

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Sprite = Player(Index).Sprite
    If Sprite <= 0 Or Sprite > Count_Sprite Then Exit Sub
    
    anim = 1
    
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            If (Player(Index).yOffset > 8) Then anim = Player(Index).Step
        Case DIR_DOWN
            If (Player(Index).yOffset < -8) Then anim = Player(Index).Step
        Case DIR_LEFT
            If (Player(Index).xOffset > 8) Then anim = Player(Index).Step
        Case DIR_RIGHT
            If (Player(Index).xOffset < -8) Then anim = Player(Index).Step
    End Select
    
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 1
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 2
    End Select
    
    With sRec
        .top = spritetop * (GetTextureHeight(Tex_Sprite(Sprite)) / 4)
        .bottom = (GetTextureHeight(Tex_Sprite(Sprite)) / 4)
        .Left = anim * (GetTextureWidth(Tex_Sprite(Sprite)) / 3)
        .Right = (GetTextureWidth(Tex_Sprite(Sprite)) / 3)
    End With

    X = GetPlayerX(Index) * Pic_Size + Player(Index).xOffset - ((GetTextureWidth(Tex_Sprite(Sprite)) / 3 - Pic_Size) / 2)
    If GetTextureHeight(Tex_Sprite(Sprite)) > Pic_Size Then
        Y = GetPlayerY(Index) * Pic_Size + Player(Index).yOffset - ((GetTextureHeight(Tex_Sprite(Sprite)) / 4) - Pic_Size) - 4
    Else
        Y = GetPlayerY(Index) * Pic_Size + Player(Index).yOffset - 4
    End If
    
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            DrawShadow X + ((GetTextureWidth(Tex_Sprite(Sprite)) / 3) / 2) - (GetTextureWidth(Tex_Misc(MiscShadow)) / 2), Y + 23
            RenderTexture Tex_Sprite(Sprite), ConvertMapX(X), ConvertMapY(Y), sRec.Left, sRec.top, sRec.Right, sRec.bottom, sRec.Right, sRec.bottom
            DrawPokeSprite Index
        Case DIR_DOWN
            DrawPokeSprite Index
            DrawShadow X + ((GetTextureWidth(Tex_Sprite(Sprite)) / 3) / 2) - (GetTextureWidth(Tex_Misc(MiscShadow)) / 2), Y + 23
            RenderTexture Tex_Sprite(Sprite), ConvertMapX(X), ConvertMapY(Y), sRec.Left, sRec.top, sRec.Right, sRec.bottom, sRec.Right, sRec.bottom
        Case DIR_LEFT
            DrawPokeSprite Index
            DrawShadow X + ((GetTextureWidth(Tex_Sprite(Sprite)) / 3) / 2) - (GetTextureWidth(Tex_Misc(MiscShadow)) / 2), Y + 23
            RenderTexture Tex_Sprite(Sprite), ConvertMapX(X), ConvertMapY(Y), sRec.Left, sRec.top, sRec.Right, sRec.bottom, sRec.Right, sRec.bottom
        Case DIR_RIGHT
            DrawPokeSprite Index
            DrawShadow X + ((GetTextureWidth(Tex_Sprite(Sprite)) / 3) / 2) - (GetTextureWidth(Tex_Misc(MiscShadow)) / 2), Y + 23
            RenderTexture Tex_Sprite(Sprite), ConvertMapX(X), ConvertMapY(Y), sRec.Left, sRec.top, sRec.Right, sRec.bottom, sRec.Right, sRec.bottom
    End Select
    
    Exit Sub
errHandler:
    HandleError "DrawPlayer", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawNpc(ByVal MapNpcNum As Long)
Dim anim As Byte, X As Long, Y As Long
Dim Sprite As Long, spritetop As Long
Dim sRec As RECT

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If MapNpc(MapNpcNum).Num = 0 Then Exit Sub
    Sprite = NPC(MapNpc(MapNpcNum).Num).Sprite
    If Sprite < 1 Or Sprite > Count_Sprite Then Exit Sub

    anim = 1
    Select Case MapNpc(MapNpcNum).Dir
        Case DIR_UP
            If (MapNpc(MapNpcNum).yOffset > 8) Then anim = MapNpc(MapNpcNum).Step
        Case DIR_DOWN
            If (MapNpc(MapNpcNum).yOffset < -8) Then anim = MapNpc(MapNpcNum).Step
        Case DIR_LEFT
            If (MapNpc(MapNpcNum).xOffset > 8) Then anim = MapNpc(MapNpcNum).Step
        Case DIR_RIGHT
            If (MapNpc(MapNpcNum).xOffset < -8) Then anim = MapNpc(MapNpcNum).Step
    End Select

    Select Case MapNpc(MapNpcNum).Dir
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 1
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 2
    End Select

    With sRec
        .top = spritetop * (GetTextureHeight(Tex_Sprite(Sprite)) / 4)
        .bottom = (GetTextureHeight(Tex_Sprite(Sprite)) / 4)
        .Left = anim * (GetTextureWidth(Tex_Sprite(Sprite)) / 3)
        .Right = (GetTextureWidth(Tex_Sprite(Sprite)) / 3)
    End With

    X = MapNpc(MapNpcNum).X * Pic_Size + MapNpc(MapNpcNum).xOffset - ((GetTextureWidth(Tex_Sprite(Sprite)) / 3 - Pic_Size) / 2)
    If GetTextureHeight(Tex_Sprite(Sprite)) > Pic_Size Then
        Y = MapNpc(MapNpcNum).Y * Pic_Size + MapNpc(MapNpcNum).yOffset - ((GetTextureHeight(Tex_Sprite(Sprite)) / 4) - Pic_Size) - 4
    Else
        Y = MapNpc(MapNpcNum).Y * Pic_Size + MapNpc(MapNpcNum).yOffset - 4
    End If
    
    DrawShadow X + ((GetTextureWidth(Tex_Sprite(Sprite)) / 3) / 2) - (GetTextureWidth(Tex_Misc(MiscShadow)) / 2), Y + 23
    RenderTexture Tex_Sprite(Sprite), ConvertMapX(X), ConvertMapY(Y), sRec.Left, sRec.top, sRec.Right, sRec.bottom, sRec.Right, sRec.bottom
    
    Exit Sub
errHandler:
    HandleError "DrawNpc", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawPokeSprite(ByVal Index As Long)
Dim Sprite As Long, anim As Long
Dim spritetop As Byte
Dim X As Long, Y As Long
Dim x2 As Long, y2 As Long
Dim sRec As RECT

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Player(Index).Pokemon(1).Num <= 0 Then Exit Sub
    Sprite = Pokemon(Player(Index).Pokemon(1).Num).Pic
    If Sprite <= 0 Or Sprite > Count_PokeSprite Then Exit Sub
    
    anim = 1
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            x2 = GetPlayerX(Index)
            y2 = GetPlayerY(Index) + 1
            If (Player(Index).yOffset > 8) Then anim = Player(Index).Step
        Case DIR_DOWN
            x2 = GetPlayerX(Index)
            y2 = GetPlayerY(Index) - 1
            If (Player(Index).yOffset < -8) Then anim = Player(Index).Step
        Case DIR_LEFT
            x2 = GetPlayerX(Index) + 1
            y2 = GetPlayerY(Index)
            If (Player(Index).xOffset > 8) Then anim = Player(Index).Step
        Case DIR_RIGHT
            x2 = GetPlayerX(Index) - 1
            y2 = GetPlayerY(Index)
            If (Player(Index).xOffset < -8) Then anim = Player(Index).Step
    End Select
    
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            If (Player(Index).yOffset > 8) Then anim = Player(Index).Step
        Case DIR_DOWN
            If (Player(Index).yOffset < -8) Then anim = Player(Index).Step
        Case DIR_LEFT
            If (Player(Index).xOffset > 8) Then anim = Player(Index).Step
        Case DIR_RIGHT
            If (Player(Index).xOffset < -8) Then anim = Player(Index).Step
    End Select

    Select Case GetPlayerDir(Index)
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select
    
    With sRec
        .top = spritetop * (GetTextureHeight(Tex_PokeSprite(Sprite)) / 4)
        .bottom = (GetTextureHeight(Tex_PokeSprite(Sprite)) / 4)
        .Left = anim * (GetTextureWidth(Tex_PokeSprite(Sprite)) / 3)
        .Right = (GetTextureWidth(Tex_PokeSprite(Sprite)) / 3)
    End With

    X = x2 * Pic_Size + Player(Index).xOffset - ((GetTextureWidth(Tex_PokeSprite(Sprite)) / 3 - Pic_Size) / 2)
    If GetTextureHeight(Tex_PokeSprite(Sprite)) > Pic_Size Then
        Y = y2 * Pic_Size + Player(Index).yOffset - ((GetTextureHeight(Tex_PokeSprite(Sprite)) / 4) - Pic_Size) - 4
    Else
        Y = y2 * Pic_Size + Player(Index).yOffset - 4
    End If
    
    DrawShadow X + ((GetTextureWidth(Tex_PokeSprite(Sprite)) / 3) / 2) - (GetTextureWidth(Tex_Misc(MiscShadow)) / 2), Y + 23
    RenderTexture Tex_PokeSprite(Sprite), ConvertMapX(X), ConvertMapY(Y), sRec.Left, sRec.top, sRec.Right, sRec.bottom, sRec.Right, sRec.bottom
    
    Exit Sub
errHandler:
    HandleError "DrawPokeSprite", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawShadow(ByVal X As Long, ByVal Y As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    RenderTexture Tex_Misc(MiscShadow), ConvertMapX(X), ConvertMapY(Y), 0, 0, Pic_Size, Pic_Size, Pic_Size, Pic_Size, D3DColorARGB(150, 255, 255, 255)
    
    Exit Sub
errHandler:
    HandleError "DrawShadow", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawGui()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If TitleBarAlpha > 0 Then
        DrawTitleBar
    End If
    
    If Player(MyIndex).InBattle > 0 And FadeType = 0 Then
        DrawBattle
    End If
    DrawFade
    DrawChatbox
    RenderTexture Tex_Gui(GuiPokeView), GuiPokeViewX, GuiPokeViewY, 0, 0, GetTextureWidth(Tex_Gui(GuiPokeView)), GetTextureHeight(Tex_Gui(GuiPokeView)), GetTextureWidth(Tex_Gui(GuiPokeView)), GetTextureHeight(Tex_Gui(GuiPokeView)), D3DColorARGB(130, 255, 255, 255)
    DrawPokeView
    
    DrawTrainerWindow
    DrawInventory
    DrawOption
    
    If InStorage Then DrawStorage
    If InShop > 0 Then DrawShop
    If InTrade Then DrawTrade
    If InTradeConfirm Then DrawTradeConfirm
    
    DrawSelect
    
    If IsLearnMove Then DrawLearnMove
    If IsEvolve Then DrawEvolve
    
    DrawButton ButtonEnum.mPokedex
    DrawButton ButtonEnum.mInventory
    DrawButton ButtonEnum.mCharacter
    DrawButton ButtonEnum.mOptions
    
    If MyTarget > 0 Then DrawTargetMenu MyTarget
    
    DrawInput
    
    Exit Sub
errHandler:
    HandleError "DrawGui", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawChatbox()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Not WindowVisible(WindowType.Main_Chatbox) Then Exit Sub
    
    RenderTexture Tex_Gui(GuiChatbox), GuiChatboxX, GuiChatboxY, 0, 0, GetTextureWidth(Tex_Gui(GuiChatbox)), GetTextureHeight(Tex_Gui(GuiChatbox)), GetTextureWidth(Tex_Gui(GuiChatbox)), GetTextureHeight(Tex_Gui(GuiChatbox)), D3DColorARGB(130, 255, 255, 255)

    DrawButton ButtonEnum.bChatScrollUp
    DrawButton ButtonEnum.bChatScrollDown
    
    If ChatOn Then
        RenderText CurFont, "Chat:", GuiChatboxX + 20, GuiChatboxY + 122, D3DColorARGB(180, 255, 255, 255)
        RenderText CurFont, RenderChatMsg & ChatLine, GuiChatboxX + 60, GuiChatboxY + 122, White
    Else
        RenderText CurFont, "Press 'Enter' to Chat", GuiChatboxX + 160, GuiChatboxY + 122, D3DColorARGB(150, 255, 255, 255)
    End If
    
    RenderChatTextBuffer

    Exit Sub
errHandler:
    HandleError "DrawChatbox", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawAttribute()
Dim X As Long, Y As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If frmMapEditor.SSTab1.Tab <> 1 Then Exit Sub
    
    For X = TileView.Left To TileView.Right
        For Y = TileView.top To TileView.bottom
            If IsValidMapPoint(X, Y) Then
                If Map.Tile(X, Y).Type > 0 Then
                    RenderTexture Tex_Misc(MiscBlank), ConvertMapX(X * Pic_Size), ConvertMapY(Y * Pic_Size), 0, 0, Pic_Size, Pic_Size, 1, 1, GetAttributeColor(Map.Tile(X, Y).Type)
                End If
            End If
        Next
    Next
    
    Exit Sub
errHandler:
    HandleError "DrawAttribute", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawCursor()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    RenderTexture Tex_Misc(MiscCursor), GlobalX, GlobalY, 20 * IsClicked, 0, 20, 20, 20, 20
    
    Exit Sub
errHandler:
    HandleError "DrawCursor", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub Editor_DrawPokeIcon()
Dim desRect As D3DRECT
Dim Scrl As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Scrl = frmPokemonEditor.scrlPokeIcon.value
    
    If Scrl <= 0 Or Scrl > Count_PokeIcon Then
        frmPokemonEditor.picIcon.Cls
        Exit Sub
    End If
    
    With desRect
        .x2 = frmPokemonEditor.picIcon.ScaleWidth
        .y2 = frmPokemonEditor.picIcon.ScaleHeight
    End With

    D3DDevice8.Clear 0, ByVal 0, D3DCLEAR_TARGET, White, 1#, 0
    D3DDevice8.BeginScene
    
    RenderTexture Tex_PokeIcon(Scrl), 0, 0, (GetTextureWidth(Tex_PokeIcon(Scrl)) / 2) * PokeIconAnim, 0, GetTextureWidth(Tex_PokeIcon(Scrl)) / 2, GetTextureHeight(Tex_PokeIcon(Scrl)), GetTextureWidth(Tex_PokeIcon(Scrl)) / 2, GetTextureHeight(Tex_PokeIcon(Scrl))
    
    D3DDevice8.EndScene
    D3DDevice8.Present desRect, desRect, frmPokemonEditor.picIcon.hWnd, ByVal 0
    
    Exit Sub
errHandler:
    HandleError "Editor_DrawPokeIcon", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub Editor_DrawItemPic()
Dim desRect As D3DRECT
Dim Scrl As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Scrl = frmItemEditor.scrlItemPic.value
    
    If Scrl <= 0 Or Scrl > Count_PokeIcon Then
        frmItemEditor.picItem.Cls
        Exit Sub
    End If
    
    With desRect
        .x2 = frmItemEditor.picItem.ScaleWidth
        .y2 = frmItemEditor.picItem.ScaleHeight
    End With

    D3DDevice8.Clear 0, ByVal 0, D3DCLEAR_TARGET, White, 1#, 0
    D3DDevice8.BeginScene
    
    RenderTexture Tex_ItemPic(Scrl), 0, 0, 0, 0, GetTextureWidth(Tex_ItemPic(Scrl)), GetTextureHeight(Tex_ItemPic(Scrl)), GetTextureWidth(Tex_ItemPic(Scrl)), GetTextureHeight(Tex_ItemPic(Scrl))
    
    D3DDevice8.EndScene
    D3DDevice8.Present desRect, desRect, frmItemEditor.picItem.hWnd, ByVal 0
    
    Exit Sub
errHandler:
    HandleError "Editor_DrawItemPic", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawPokeView()
Dim i As Long
Dim X As Long
Dim Width As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    For i = 0 To (MAX_POKEMON - 1)
        X = (GuiPokeViewX + 9) + ((GetTextureWidth(Tex_Gui(GuiPokeSlot)) + 3) * i)
        RenderTexture Tex_Gui(GuiPokeSlot), X, GuiPokeViewY + 7, 0, 0, GetTextureWidth(Tex_Gui(GuiPokeSlot)), GetTextureHeight(Tex_Gui(GuiPokeSlot)), GetTextureWidth(Tex_Gui(GuiPokeSlot)), GetTextureHeight(Tex_Gui(GuiPokeSlot)), D3DColorARGB(190, 255, 255, 255)
        With Player(MyIndex).Pokemon(i + 1)
            If .Num > 0 Then
                If Pokemon(.Num).Pic > 0 And Pokemon(.Num).Pic <= Count_PokeIcon Then
                    RenderTexture Tex_PokeIcon(Pokemon(.Num).Pic), X + 3, GuiPokeViewY + 15, 32 * PokeIconAnim, 0, 32, 32, 32, 32
                End If
                ' Temp: Have Nickname
                RenderText CurFont, Trim$(Pokemon(.Num).Name), X + 37, GuiPokeViewY + 13, D3DColorARGB(150, 255, 255, 255)

                Width = (.CurHP / 67) / (.Stat(Stats.HP) / 67) * 67
                RenderTexture Tex_Misc(MiscBars), X + 51, GuiPokeViewY + 31, 0, 0, Width, GetTextureHeight(Tex_Misc(MiscBars)) / 2, Width, GetTextureHeight(Tex_Misc(MiscBars)) / 2
                Width = (.Exp / 67) / (ExpCalc(.Level) / 67) * 67
                RenderTexture Tex_Misc(MiscBars), X + 51, GuiPokeViewY + 39, 0, GetTextureHeight(Tex_Misc(MiscBars)) / 2, Width, GetTextureHeight(Tex_Misc(MiscBars)) / 2, Width, GetTextureHeight(Tex_Misc(MiscBars)) / 2
                
                RenderTexture Tex_Misc(MiscGender), X + 110, GuiPokeViewY + 15, 8 * .Gender, 0, 8, 11, 8, 11
            End If
        End With
    Next
    
    Exit Sub
errHandler:
    HandleError "DrawPokeView", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawBattle()
Dim X As Long, Y As Long
Dim color As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If MapBackground <= 0 Or MapField <= 0 Then
        MsgBox "Cannot load Battle Background!", vbCritical
        CloseApp
        Exit Sub
    End If
    
    RenderTexture Tex_Gui(GuiBattle), GuiBattleX, GuiBattleY, 0, 0, GetTextureWidth(Tex_Gui(GuiBattle)), GetTextureHeight(Tex_Gui(GuiBattle)), GetTextureWidth(Tex_Gui(GuiBattle)), GetTextureHeight(Tex_Gui(GuiBattle))
    RenderTexture Tex_Background(MapBackground), GuiBattleX, GuiBattleY + 60, 0, 0, 512, 405, GetTextureWidth(Tex_Background(MapBackground)), GetTextureHeight(Tex_Background(MapBackground)), D3DColorARGB(150, 255, 255, 255)
    
    If Not CanUseCmd And Not ExitBattleTmr Then
        RenderText CurFont, "Please wait for a moment...", GuiBattleX + (512 / 2) - (GetWidth(CurFont, "Please wait for a moment...") / 2), GuiBattleY + 60 + (405 / 2) - 8, White
    ElseIf ExitBattleTmr And CanExit Then
        RenderText CurFont, "Click the screen to exit the battle", GuiBattleX + (512 / 2) - (GetWidth(CurFont, "Click here to exit the battle") / 2), GuiBattleY + 60 + (405 / 2) - 8, White
    End If
    
    With Player(MyIndex).Pokemon(CurPoke)
        If CurBarWidth > 70 Then
            color = D3DColorARGB(255, 24, 192, 32)
        ElseIf CurBarWidth <= 70 And CurBarWidth > 30 Then
            color = D3DColorARGB(255, 248, 176, 0)
        ElseIf CurBarWidth <= 30 Then
            color = D3DColorARGB(255, 248, 88, 40)
        End If
        RenderTexture Tex_FieldBack(MapField), GuiBattleX, GuiBattleY + 401, 100, 0, GetTextureWidth(Tex_FieldBack(MapField)) - 100, GetTextureHeight(Tex_FieldBack(MapField)), GetTextureWidth(Tex_FieldBack(MapField)) - 100, GetTextureHeight(Tex_FieldBack(MapField)), D3DColorARGB(150, 255, 255, 255)
        RenderTexture Tex_PokeBack(Pokemon(.Num).Pic), (GuiBattleX + 93) + SwitchPokeX, GuiBattleY + 365, 0, 0, GetTextureWidth(Tex_PokeBack(Pokemon(.Num).Pic)), GetTextureHeight(Tex_PokeBack(Pokemon(.Num).Pic)), GetTextureWidth(Tex_PokeBack(Pokemon(.Num).Pic)), GetTextureHeight(Tex_PokeBack(Pokemon(.Num).Pic)), D3DColorARGB(PokeAlpha, 255, 255, 255)
        
        RenderTexture Tex_Gui(GuiBattleHP), GuiBattleX + 285, GuiBattleY + 354, 0, 0, GetTextureWidth(Tex_Gui(GuiBattleHP)), GetTextureHeight(Tex_Gui(GuiBattleHP)), GetTextureWidth(Tex_Gui(GuiBattleHP)), GetTextureHeight(Tex_Gui(GuiBattleHP))
        RenderText CurFont, Trim$(Pokemon(.Num).Name), 327, 369, Black
        RenderText CurFont, .Level, 470, 369, Black
        RenderTexture Tex_Misc(MiscBattleBars), GuiBattleX + 394, GuiBattleY + 394, 0, 0, CurBarWidth, GetTextureHeight(Tex_Misc(MiscBattleBars)), CurBarWidth, GetTextureHeight(Tex_Misc(MiscBattleBars)), color
        RenderTexture Tex_Misc(MiscGender), GuiBattleX + 428, GuiBattleY + 371, 8 * .Gender, 0, 8, 11, 8, 11
    End With

    With EnemyPokemon
        If CurEnemyBarWidth > 70 Then
            color = D3DColorARGB(255, 24, 192, 32)
        ElseIf CurEnemyBarWidth <= 70 And CurEnemyBarWidth > 30 Then
            color = D3DColorARGB(255, 248, 176, 0)
        ElseIf CurEnemyBarWidth <= 30 Then
            color = D3DColorARGB(255, 248, 88, 40)
        End If
        RenderTexture Tex_FieldFront(MapField), (GuiBattleX + 240) - EnemyPos, GuiBattleY + 150, 0, 0, GetTextureWidth(Tex_FieldFront(MapField)), GetTextureHeight(Tex_FieldFront(MapField)), GetTextureWidth(Tex_FieldFront(MapField)), GetTextureHeight(Tex_FieldFront(MapField)), D3DColorARGB(150, 255, 255, 255)
        RenderTexture Tex_PokeFront(Pokemon(.Num).Pic), (GuiBattleX + 317) - EnemyPos, GuiBattleY + 140, 0, 0, GetTextureWidth(Tex_PokeFront(Pokemon(.Num).Pic)), GetTextureHeight(Tex_PokeFront(Pokemon(.Num).Pic)), GetTextureWidth(Tex_PokeFront(Pokemon(.Num).Pic)), GetTextureHeight(Tex_PokeFront(Pokemon(.Num).Pic)), D3DColorARGB(Capture, 255, 255, 255)
        
        If EnemyPos = 0 Then
            RenderTexture Tex_Gui(GuiBattleFoeHP), GuiBattleX, GuiBattleY + 131, 0, 0, GetTextureWidth(Tex_Gui(GuiBattleFoeHP)), GetTextureHeight(Tex_Gui(GuiBattleFoeHP)), GetTextureWidth(Tex_Gui(GuiBattleFoeHP)), GetTextureHeight(Tex_Gui(GuiBattleFoeHP))
            RenderText CurFont, Trim$(Pokemon(.Num).Name), 25, 146, Black
            RenderText CurFont, .Level, 166, 146, Black
            RenderTexture Tex_Misc(MiscBattleBars), GuiBattleX + 92, GuiBattleY + 171, 0, 0, CurEnemyBarWidth, GetTextureHeight(Tex_Misc(MiscBattleBars)), CurEnemyBarWidth, GetTextureHeight(Tex_Misc(MiscBattleBars)), color
            RenderTexture Tex_Misc(MiscGender), GuiBattleX + 126, GuiBattleY + 148, 8 * .Gender, 0, 8, 11, 8, 11
        End If
    End With
    
    DrawButton ButtonEnum.BattleFight, "FIGHT", White
    DrawButton ButtonEnum.BattleSwitch, "POKEMON", White
    DrawButton ButtonEnum.BattleBag, "BAG", White
    DrawButton ButtonEnum.BattleRun, "RUN", White
    
    DrawButton ButtonEnum.bBattleScrollUp
    DrawButton ButtonEnum.bBattleScrollDown
    
    RenderBattleTextBuffer
    
    DrawMoves
    DrawPokemonSwitch
    
    Exit Sub
errHandler:
    HandleError "DrawBattle", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawTitleBar()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    With Map
        RenderTexture Tex_TitleBar(.Moral), 10, 70, 0, 0, GetTextureWidth(Tex_TitleBar(.Moral)), GetTextureHeight(Tex_TitleBar(.Moral)), GetTextureWidth(Tex_TitleBar(.Moral)), GetTextureHeight(Tex_TitleBar(.Moral)), D3DColorARGB(TitleBarAlpha, 255, 255, 255)
        RenderText CurFont, Trim$(.Name), 20, 85, D3DColorARGB(TitleBarAlpha, 255, 255, 255)
    End With
    
    Exit Sub
errHandler:
    HandleError "DrawTitleBar", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawFade()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    RenderTexture Tex_Misc(MiscBlank), 0, 0, 0, 0, ScreenWidth, ScreenHeight, ScreenWidth, ScreenHeight, D3DColorARGB(FadeAlpha, 0, 0, 0)
    
    Exit Sub
errHandler:
    HandleError "DrawFade", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawInBattleIcon(ByVal Index As Long)
Dim X As Long, Y As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    X = GetPlayerX(Index) * Pic_Size + Player(Index).xOffset + (Pic_Size / 2) - (30 / 2)
    Y = GetPlayerY(Index) * Pic_Size + Player(Index).yOffset - 32
    If GetPlayerSprite(Index) >= 1 And GetPlayerSprite(Index) <= Count_Sprite Then
        Y = GetPlayerY(Index) * Pic_Size + Player(Index).yOffset - (GetTextureHeight(Tex_Sprite(GetPlayerSprite(Index))) / 4) - 23
    End If
    
    RenderTexture Tex_Misc(MiscInBattle), ConvertMapX(X), ConvertMapY(Y), 0, 0, 30, 32, 30, 32

    Exit Sub
errHandler:
    HandleError "DrawInBattleIcon", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawTarget(ByVal Index As Long)
Dim X As Long, Y As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    With Player(Index)
        X = .X * Pic_Size + .xOffset - ((GetTextureWidth(Tex_Sprite(.Sprite)) / 3 - Pic_Size) / 2)
        If GetTextureHeight(Tex_Sprite(.Sprite)) > Pic_Size Then
            Y = .Y * Pic_Size + .yOffset - ((GetTextureHeight(Tex_Sprite(.Sprite)) / 4) - Pic_Size) - 4
        Else
            Y = .Y * Pic_Size + .yOffset - 4
        End If
        
        RenderTexture Tex_Misc(MiscTarget), ConvertMapX((X - 3) - TargetAnim), ConvertMapY((Y - 3) - TargetAnim), 0, 0, 14, 14, 14, 14
        RenderTexture Tex_Misc(MiscTarget), ConvertMapX((X + 25) + TargetAnim), ConvertMapY((Y - 3) - TargetAnim), 14, 0, 14, 14, 14, 14
        RenderTexture Tex_Misc(MiscTarget), ConvertMapX((X - 3) - TargetAnim), ConvertMapY((Y + 38) + TargetAnim), 0, 14, 14, 14, 14, 14
        RenderTexture Tex_Misc(MiscTarget), ConvertMapX((X + 25) + TargetAnim), ConvertMapY((Y + 38) + TargetAnim), 14, 14, 14, 14, 14, 14
    End With

    Exit Sub
errHandler:
    HandleError "DrawTarget", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawMoves()
Dim i As Long, X As Long, Y As Long
Dim mText As String, Width As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Not ShowMoves Then Exit Sub
    
    X = GuiBattleX + 542
    For i = 1 To MAX_POKEMON_MOVES
        Y = GuiBattleY + 342 + ((i - 1) * 20)
        RenderTexture Tex_Gui(13), X, Y, 0, 0, 107, 20, 107, 20
        With Player(MyIndex).Pokemon(CurPoke)
            If .Moves(i).Num > 0 Then
                mText = Trim$(Moves(.Moves(i).Num).Name)
                If .Moves(i).MaxPP > 0 Then
                    Width = (.Moves(i).PP / 107) / (.Moves(i).MaxPP / 107) * 107
                End If
                RenderTexture Tex_Gui(13), X, Y, 0, 0, Width, 20, Width, 20, D3DColorARGB(255, 255, 255, 0)
                If .Moves(i).PP <= 0 Then
                    RenderText CurFont, mText, X + ((107 / 2) - (GetWidth(CurFont, mText) / 2)), Y + 2, Silver
                Else
                    If GlobalX >= X And GlobalX <= X + 107 And GlobalY >= Y + 2 And GlobalY <= Y + 18 Then
                        RenderText CurFont, mText, X + ((107 / 2) - (GetWidth(CurFont, mText) / 2)), Y + 2, Yellow
                    Else
                        RenderText CurFont, mText, X + ((107 / 2) - (GetWidth(CurFont, mText) / 2)), Y + 2, White
                    End If
                End If
            End If
        End With
    Next

    Exit Sub
errHandler:
    HandleError "DrawMoves", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawPokemonSwitch()
Dim X As Long, Y As Long, i As Long
Dim mText As String

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Not ShowPokemonSwitch Then Exit Sub
    
    X = GuiBattleX + 665
    For i = 1 To MAX_POKEMON
        Y = GuiBattleY + 302 + ((i - 1) * 20)
        RenderTexture Tex_Gui(13), X, Y, 0, 0, 107, 20, 107, 20
        With Player(MyIndex)
            If .Pokemon(i).Num > 0 Then
                mText = Trim$(Pokemon(.Pokemon(i).Num).Name)
                If i = CurPoke Then
                    RenderText CurFont, mText, X + ((107 / 2) - (GetWidth(CurFont, mText) / 2)), Y + 2, Green
                ElseIf .Pokemon(i).CurHP <= 0 Then
                    RenderText CurFont, mText, X + ((107 / 2) - (GetWidth(CurFont, mText) / 2)), Y + 2, Silver
                Else
                    If GlobalX >= X And GlobalX <= X + 107 And GlobalY >= Y + 2 And GlobalY <= Y + 18 Then
                        RenderText CurFont, mText, X + ((107 / 2) - (GetWidth(CurFont, mText) / 2)), Y + 2, Yellow
                    Else
                        RenderText CurFont, mText, X + ((107 / 2) - (GetWidth(CurFont, mText) / 2)), Y + 2, White
                    End If
                End If
            End If
        End With
    Next
    
    Exit Sub
errHandler:
    HandleError "DrawPokemonSwitch", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawLearnMove()
Dim X As Long, Y As Long
Dim x2 As Long, y2 As Long
Dim i As Byte, mText As String

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If LearnPokeNum <= 0 Then Exit Sub
    
    With Player(MyIndex).Pokemon(LearnPokeNum)
        RenderTexture Tex_Gui(GuiLearnMove), GuiLearnMoveX, GuiLearnMoveY, 0, 0, GetTextureWidth(Tex_Gui(GuiLearnMove)), GetTextureHeight(Tex_Gui(GuiLearnMove)), GetTextureWidth(Tex_Gui(GuiLearnMove)), GetTextureHeight(Tex_Gui(GuiLearnMove))

        mText = Trim$(Pokemon(.Num).Name) & " is trying to learn " & Trim$(Moves(LearnMoveNum).Name)
        RenderText CurFont, mText, GuiLearnMoveX + 19 + (269 / 2) - (GetWidth(CurFont, mText) / 2), GuiLearnMoveY + 27, White
        mText = "Delete a move to make a room for " & Trim$(Moves(LearnMoveNum).Name)
        RenderText CurFont, mText, GuiLearnMoveX + 19 + (269 / 2) - (GetWidth(CurFont, mText) / 2), GuiLearnMoveY + 42, White
        
        X = GuiLearnMoveX + 100
        For i = 1 To MAX_POKEMON_MOVES
            Y = GuiLearnMoveY + 86 + ((i - 1) * 20)
            If .Moves(i).Num > 0 Then
                mText = Trim$(Moves(.Moves(i).Num).Name)
                If SelectedMove = i Then
                    RenderText CurFont, mText, X + ((107 / 2) - (GetWidth(CurFont, mText) / 2)), Y + 2, Silver
                Else
                    If GlobalX >= X And GlobalX <= X + 107 And GlobalY >= Y + 2 And GlobalY <= Y + 18 Then
                        RenderText CurFont, mText, X + ((107 / 2) - (GetWidth(CurFont, mText) / 2)), Y + 2, Yellow
                    Else
                        RenderText CurFont, mText, X + ((107 / 2) - (GetWidth(CurFont, mText) / 2)), Y + 2, White
                    End If
                End If
            End If
        Next
        x2 = GuiLearnMoveX + 100 + ((107 / 2) - (GetWidth(CurFont, "Replace") / 2)): y2 = GuiLearnMoveY + 175
        If GlobalX >= x2 And GlobalX <= x2 + GetWidth(CurFont, "Replace") And GlobalY >= y2 And GlobalY <= y2 + 16 Then
            RenderText CurFont, "Replace", GuiLearnMoveX + 100 + ((107 / 2) - (GetWidth(CurFont, "Replace") / 2)), GuiLearnMoveY + 175, Yellow
        Else
            RenderText CurFont, "Replace", GuiLearnMoveX + 100 + ((107 / 2) - (GetWidth(CurFont, "Replace") / 2)), GuiLearnMoveY + 175, White
        End If
        x2 = GuiLearnMoveX + 100 + ((107 / 2) - (GetWidth(CurFont, "Cancel") / 2)): y2 = GuiLearnMoveY + 195
        If GlobalX >= x2 And GlobalX <= x2 + GetWidth(CurFont, "Cancel") And GlobalY >= y2 And GlobalY <= y2 + 16 Then
            RenderText CurFont, "Cancel", GuiLearnMoveX + 100 + ((107 / 2) - (GetWidth(CurFont, "Cancel") / 2)), GuiLearnMoveY + 195, Yellow
        Else
            RenderText CurFont, "Cancel", GuiLearnMoveX + 100 + ((107 / 2) - (GetWidth(CurFont, "Cancel") / 2)), GuiLearnMoveY + 195, White
        End If
    End With

    Exit Sub
errHandler:
    HandleError "DrawLearnMove", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawEvolve()
Dim X As Long, Y As Long, mText As String

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Not IsEvolve Then Exit Sub

    With Player(MyIndex).Pokemon(EvolvePoke)
        RenderTexture Tex_Gui(GuiEvolve), GuiEvolveX, GuiEvolveY, 0, 0, GetTextureWidth(Tex_Gui(GuiEvolve)), GetTextureHeight(Tex_Gui(GuiEvolve)), GetTextureWidth(Tex_Gui(GuiEvolve)), GetTextureHeight(Tex_Gui(GuiEvolve)), D3DColorARGB(230, 255, 255, 255)
        mText = Trim$(Pokemon(TmpCurNum).Name) & " is trying to evolve into " & Trim$(Pokemon(TmpEvolveNum).Name)
        X = GuiEvolveX + (GetTextureWidth(Tex_Gui(GuiEvolve)) / 2) - (GetWidth(CurFont, mText) / 2): Y = GuiEvolveY + 30
        RenderText CurFont, mText, X, Y, White
        mText = "Press YES to allow " & Trim$(Pokemon(TmpCurNum).Name) & " to evolve, or press NO to stop it"
        X = GuiEvolveX + (GetTextureWidth(Tex_Gui(GuiEvolve)) / 2) - (GetWidth(CurFont, mText) / 2): Y = GuiEvolveY + 46
        RenderText CurFont, mText, X, Y, White
        
        X = GuiEvolveX + ((GetTextureWidth(Tex_Gui(GuiEvolve)) / 2) - 50)
        Y = GuiEvolveY + 190
        RenderTexture Tex_PokeFront(Pokemon(DrawPokeNum).Pic), X, Y, 0, 0, GetTextureWidth(Tex_PokeFront(Pokemon(DrawPokeNum).Pic)), GetTextureHeight(Tex_PokeFront(Pokemon(DrawPokeNum).Pic)), GetTextureWidth(Tex_PokeFront(Pokemon(DrawPokeNum).Pic)), GetTextureHeight(Tex_PokeFront(Pokemon(DrawPokeNum).Pic)), D3DColorARGB(EvolveAlpha, 255, 255, 255)
        
        DrawButton ButtonEnum.EvolveYes, "YES", White
        DrawButton ButtonEnum.EvolveNo, "NO", White
    End With
    
    Exit Sub
errHandler:
    HandleError "DrawEvolve", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawTrainerWindow()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Not WindowVisible(WindowType.Main_Trainer) Then Exit Sub
    
    RenderTexture Tex_Gui(GuiTrainer), GuiTrainerX, GuiTrainerY, 0, 0, GetTextureWidth(Tex_Gui(GuiTrainer)), GetTextureHeight(Tex_Gui(GuiTrainer)), GetTextureWidth(Tex_Gui(GuiTrainer)), GetTextureHeight(Tex_Gui(GuiTrainer)), D3DColorARGB(240, 255, 255, 255)
    With Player(MyIndex)
        RenderText CurFont, "Name: ", GuiTrainerX + 24, GuiTrainerY + 27, D3DColorARGB(170, 255, 255, 255)
        RenderText CurFont, Trim$(.Name), GuiTrainerX + 24 + GetWidth(CurFont, "Name: "), GuiTrainerY + 27, White
        RenderText CurFont, "Money: ", GuiTrainerX + 24, GuiTrainerY + 43, D3DColorARGB(170, 255, 255, 255)
        RenderText CurFont, Player(MyIndex).Money, GuiTrainerX + 24 + GetWidth(CurFont, "Money: "), GuiTrainerY + 43, White
        RenderText CurFont, "Reputation: ", GuiTrainerX + 24, GuiTrainerY + 58, D3DColorARGB(170, 255, 255, 255)
        RenderText CurFont, "0", GuiTrainerX + 24 + GetWidth(CurFont, "Reputation: "), GuiTrainerY + 58, White
        
        RenderText CurFont, "PvP Stats", GuiTrainerX + 24, GuiTrainerY + 80, White
        RenderText CurFont, "Wins: ", GuiTrainerX + 24, GuiTrainerY + 96, D3DColorARGB(170, 255, 255, 255)
        RenderText CurFont, Player(MyIndex).PvP.win, GuiTrainerX + 24 + GetWidth(CurFont, "Wins: "), GuiTrainerY + 96, White
        RenderText CurFont, "Losses: ", GuiTrainerX + 24, GuiTrainerY + 112, D3DColorARGB(170, 255, 255, 255)
        RenderText CurFont, Player(MyIndex).PvP.Lose, GuiTrainerX + 24 + GetWidth(CurFont, "Losses: "), GuiTrainerY + 112, White
        RenderText CurFont, "Disconnects: ", GuiTrainerX + 24, GuiTrainerY + 128, D3DColorARGB(170, 255, 255, 255)
        RenderText CurFont, Player(MyIndex).PvP.Disconnect, GuiTrainerX + 24 + GetWidth(CurFont, "Disconnects: "), GuiTrainerY + 128, White
        
        RenderText CurFont, "Kanto Badge", GuiTrainerX + 24, GuiTrainerY + 150, White
        
        RenderText CurFont, "Playtime: ", GuiTrainerX + 24, GuiTrainerY + 190, D3DColorARGB(170, 255, 255, 255)
        RenderText CurFont, "0", GuiTrainerX + 24 + GetWidth(CurFont, "Playtime: "), GuiTrainerY + 190, White
    End With

    Exit Sub
errHandler:
    HandleError "DrawTrainerWindow", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawInventory()
Dim X As Long, Pic As Long
Dim x2 As Long, y2 As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Not WindowVisible(WindowType.Main_Inventory) Then Exit Sub
    
    RenderTexture Tex_Gui(GuiInventory), GuiInventoryX, GuiInventoryY, 0, 0, GetTextureWidth(Tex_Gui(GuiInventory)), GetTextureHeight(Tex_Gui(GuiInventory)), GetTextureWidth(Tex_Gui(GuiInventory)), GetTextureHeight(Tex_Gui(GuiInventory)), D3DColorARGB(240, 255, 255, 255)

    For x2 = ButtonEnum.InvItems To ButtonEnum.InvKeyItems
        If CurInvType + 1 = x2 - 31 Then
            Buttons(x2).X = GuiTrainerX - 54
        Else
            Buttons(x2).X = GuiTrainerX - 46
        End If
    Next
    DrawButton ButtonEnum.InvScrollUp
    DrawButton ButtonEnum.InvScrollDown
    DrawButton ButtonEnum.InvItems, "Items", White
    DrawButton ButtonEnum.InvPokeballs, "Pokeballs", White
    DrawButton ButtonEnum.InvTM_HMs, "TM/HMs", White
    DrawButton ButtonEnum.InvBerries, "Berries", White
    DrawButton ButtonEnum.InvKeyItems, "Key Items", White

    x2 = GuiInventoryX + 45
    If GetMaxInv > 0 Then
        If GetMaxInv >= 3 Then
            For X = StartInv To StartInv + 3
                y2 = GuiInventoryY + 39 + ((X - StartInv) * 39)
                With Player(MyIndex).Item(X, CurInvType)
                    If .Num > 0 Then
                        Pic = Item(.Num).Pic
                        RenderTexture Tex_ItemPic(Pic), x2, y2, 0, 0, 32, 32, 32, 32
                        RenderText CurFont, Trim$(Item(.Num).Name), x2 + 40, y2 + 9, White
                        If .value > 0 Then RenderText CurFont, "x" & .value, x2 + 150, y2 + 9, White
                    End If
                End With
            Next
            For X = StartInv To StartInv + 3
                y2 = GuiInventoryY + 39 + ((X - StartInv) * 39)
                If ShowUseItem And UseItemNum = X Then
                    RenderTexture Tex_Gui(GuiSelection), x2 + 32, y2 + 25, 0, 0, GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection)), GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection))
                    If GlobalX >= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Summary") / 2)) And GlobalX <= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Summary") / 2)) + GetWidth(CurFont, "Summary") And GlobalY >= y2 + 28 And GlobalY <= y2 + 28 + 16 Then
                        RenderText CurFont, "Summary", x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Summary") / 2)), y2 + 28, Yellow
                    Else
                        RenderText CurFont, "Summary", x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Summary") / 2)), y2 + 28, White
                    End If
                    RenderTexture Tex_Gui(GuiSelection), x2 + 32, y2 + 45, 0, 0, GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection)), GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection))
                    If GlobalX >= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Use") / 2)) And GlobalX <= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Use") / 2)) + GetWidth(CurFont, "Use") And GlobalY >= y2 + 48 And GlobalY <= y2 + 48 + 16 Then
                        RenderText CurFont, "Use", x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Use") / 2)), y2 + 48, Yellow
                    Else
                        RenderText CurFont, "Use", x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Use") / 2)), y2 + 48, White
                    End If
                    If InShop > 0 Then
                        RenderTexture Tex_Gui(GuiSelection), x2 + 32, y2 + 65, 0, 0, GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection)), GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection))
                        If GlobalX >= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Sell") / 2)) And GlobalX <= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Sell") / 2)) + GetWidth(CurFont, "Sell") And GlobalY >= y2 + 68 And GlobalY <= y2 + 68 + 16 Then
                            RenderText CurFont, "Sell", x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Sell") / 2)), y2 + 68, Yellow
                        Else
                            RenderText CurFont, "Sell", x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Sell") / 2)), y2 + 68, White
                        End If
                    End If
                    If InTrade Then
                        RenderTexture Tex_Gui(GuiSelection), x2 + 32, y2 + 65, 0, 0, GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection)), GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection))
                        If GlobalX >= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Trade") / 2)) And GlobalX <= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Trade") / 2)) + GetWidth(CurFont, "Trade") And GlobalY >= y2 + 68 And GlobalY <= y2 + 68 + 16 Then
                            RenderText CurFont, "Trade", x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Trade") / 2)), y2 + 68, Yellow
                        Else
                            RenderText CurFont, "Trade", x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Trade") / 2)), y2 + 68, White
                        End If
                    End If
                End If
            Next
        Else
            For X = 1 To GetMaxInv
                y2 = GuiInventoryY + 39 + ((X - 1) * 39)
                With Player(MyIndex).Item(X, CurInvType)
                    If .Num > 0 Then
                        Pic = Item(.Num).Pic
                        RenderTexture Tex_ItemPic(Pic), x2, y2, 0, 0, 32, 32, 32, 32
                        RenderText CurFont, Trim$(Item(.Num).Name), x2 + 40, y2 + 9, White
                        If .value > 0 Then RenderText CurFont, "x" & .value, x2 + 150, y2 + 9, White
                    End If
                End With
            Next
            For X = 1 To GetMaxInv
                y2 = GuiInventoryY + 39 + ((X - 1) * 39)
                If ShowUseItem And UseItemNum = X Then
                    RenderTexture Tex_Gui(GuiSelection), x2 + 32, y2 + 25, 0, 0, GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection)), GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection))
                    If GlobalX >= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Summary") / 2)) And GlobalX <= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Summary") / 2)) + GetWidth(CurFont, "Summary") And GlobalY >= y2 + 28 And GlobalY <= y2 + 28 + 16 Then
                        RenderText CurFont, "Summary", x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Summary") / 2)), y2 + 28, Yellow
                    Else
                        RenderText CurFont, "Summary", x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Summary") / 2)), y2 + 28, White
                    End If
                    RenderTexture Tex_Gui(GuiSelection), x2 + 32, y2 + 45, 0, 0, GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection)), GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection))
                    If GlobalX >= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Use") / 2)) And GlobalX <= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Use") / 2)) + GetWidth(CurFont, "Use") And GlobalY >= y2 + 48 And GlobalY <= y2 + 48 + 16 Then
                        RenderText CurFont, "Use", x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Use") / 2)), y2 + 48, Yellow
                    Else
                        RenderText CurFont, "Use", x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Use") / 2)), y2 + 48, White
                    End If
                    If InShop > 0 Then
                        RenderTexture Tex_Gui(GuiSelection), x2 + 32, y2 + 65, 0, 0, GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection)), GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection))
                        If GlobalX >= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Sell") / 2)) And GlobalX <= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Sell") / 2)) + GetWidth(CurFont, "Sell") And GlobalY >= y2 + 68 And GlobalY <= y2 + 68 + 16 Then
                            RenderText CurFont, "Sell", x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Sell") / 2)), y2 + 68, Yellow
                        Else
                            RenderText CurFont, "Sell", x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Sell") / 2)), y2 + 68, White
                        End If
                    End If
                    If InTrade Then
                        RenderTexture Tex_Gui(GuiSelection), x2 + 32, y2 + 65, 0, 0, GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection)), GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection))
                        If GlobalX >= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Trade") / 2)) And GlobalX <= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Trade") / 2)) + GetWidth(CurFont, "Trade") And GlobalY >= y2 + 68 And GlobalY <= y2 + 68 + 16 Then
                            RenderText CurFont, "Trade", x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Trade") / 2)), y2 + 68, Yellow
                        Else
                            RenderText CurFont, "Trade", x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Trade") / 2)), y2 + 68, White
                        End If
                    End If
                End If
            Next
        End If
    End If
    
    If Player(MyIndex).InBattle > 0 Then
        If GlobalX >= GuiInventoryX + 25 And GlobalX <= GuiInventoryX + 25 + GetWidth(CurFont, "Close") And GlobalY >= ((GuiInventoryY + GetTextureHeight(Tex_Gui(GuiInventory))) - 35) And GlobalY <= ((GuiInventoryY + GetTextureHeight(Tex_Gui(GuiInventory))) - 35 + 16) Then
            RenderText CurFont, "Close", GuiInventoryX + 25, (GuiInventoryY + GetTextureHeight(Tex_Gui(GuiInventory))) - 35, Yellow
        Else
            RenderText CurFont, "Close", GuiInventoryX + 25, (GuiInventoryY + GetTextureHeight(Tex_Gui(GuiInventory))) - 35, White
        End If
    End If
    
    If InShop > 0 Then
        RenderText CurFont, "Money: " & Player(MyIndex).Money, GuiInventoryX, GuiInventoryY - 25, White
    End If
    
    Exit Sub
errHandler:
    HandleError "DrawInventory", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawOption()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Not WindowVisible(WindowType.Main_Option) Then Exit Sub
    
    RenderTexture Tex_Gui(GuiOption), GuiOptionX, GuiOptionY, 0, 0, GetTextureWidth(Tex_Gui(GuiOption)), GetTextureHeight(Tex_Gui(GuiOption)), GetTextureWidth(Tex_Gui(GuiOption)), GetTextureHeight(Tex_Gui(GuiOption)), D3DColorARGB(240, 255, 255, 255)
    
    RenderText CurFont, "Music: ", GuiOptionX + 30, GuiOptionY + 30, D3DColorARGB(170, 255, 255, 255)
    RenderText CurFont, "On", GuiOptionX + 130, GuiOptionY + 30, D3DColorARGB(170, 255, 255, 255)
    RenderText CurFont, "Off", GuiOptionX + 197, GuiOptionY + 30, D3DColorARGB(170, 255, 255, 255)
    If Options.Music > 0 Then
        RenderTexture Tex_Misc(MiscCheck), GuiOptionX + 110, GuiOptionY + 33, 0, 0, 11, 11, 11, 11
    Else
        RenderTexture Tex_Misc(MiscCheck), GuiOptionX + 177, GuiOptionY + 33, 0, 0, 11, 11, 11, 11
    End If
    RenderText CurFont, "Sound: ", GuiOptionX + 30, GuiOptionY + 60, D3DColorARGB(170, 255, 255, 255)
    RenderText CurFont, "On", GuiOptionX + 130, GuiOptionY + 60, D3DColorARGB(170, 255, 255, 255)
    RenderText CurFont, "Off", GuiOptionX + 197, GuiOptionY + 60, D3DColorARGB(170, 255, 255, 255)
    If Options.Sound > 0 Then
        RenderTexture Tex_Misc(MiscCheck), GuiOptionX + 110, GuiOptionY + 63, 0, 0, 11, 11, 11, 11
    Else
        RenderTexture Tex_Misc(MiscCheck), GuiOptionX + 177, GuiOptionY + 63, 0, 0, 11, 11, 11, 11
    End If
    
    Exit Sub
errHandler:
    HandleError "DrawOption", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawTargetMenu(ByVal targetindex As Long)
Dim X As Long, Y As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    RenderTexture Tex_Gui(GuiTargetMenu), GuiTargetMenuX, GuiTargetMenuY, 0, 0, GetTextureWidth(Tex_Gui(GuiTargetMenu)), GetTextureHeight(Tex_Gui(GuiTargetMenu)), GetTextureWidth(Tex_Gui(GuiTargetMenu)), GetTextureHeight(Tex_Gui(GuiTargetMenu))
    RenderText CurFont, Trim$(Player(targetindex).Name), GuiTargetMenuX + 83 + (101 / 2) - (GetWidth(CurFont, Trim$(Player(targetindex).Name)) / 2), GuiTargetMenuY + 29, White
    
    X = GuiTargetMenuX + 83 + (101 / 2) - (GetWidth(CurFont, "Trade") / 2)
    Y = GuiTargetMenuY + 49
    If GlobalX >= X And GlobalX <= X + GetWidth(CurFont, "Trade") And GlobalY >= Y And GlobalY <= Y + 16 Then
        RenderText CurFont, "Trade", X, Y, Yellow
    Else
        RenderText CurFont, "Trade", X, Y, White
    End If
    Y = GuiTargetMenuY + 69
    If GlobalX >= X And GlobalX <= X + GetWidth(CurFont, "Battle") And GlobalY >= Y And GlobalY <= Y + 16 Then
        RenderText CurFont, "Battle", X, Y, Yellow
    Else
        RenderText CurFont, "Battle", X, Y, White
    End If
    RenderTexture Tex_Sprite(Player(targetindex).Sprite), GuiTargetMenuX + 31, GuiTargetMenuY + 31, (GetTextureWidth(Tex_Sprite(Player(targetindex).Sprite)) / 3), 0, GetTextureWidth(Tex_Sprite(Player(targetindex).Sprite)) / 3, GetTextureHeight(Tex_Sprite(Player(targetindex).Sprite)) / 4, GetTextureWidth(Tex_Sprite(Player(targetindex).Sprite)) / 3, GetTextureHeight(Tex_Sprite(Player(targetindex).Sprite)) / 4

    Exit Sub
errHandler:
    HandleError "DrawTargetMenu", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawSelect()
Dim i As Long
Dim X As Long, Y As Long, mText As String

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Not ShowSelect Then Exit Sub
    
    For i = 1 To MaxSelection + 1
        X = (ScreenWidth / 2) - (GetTextureWidth(Tex_Gui(GuiSelection)) / 2)
        Y = (ScreenHeight / 2) - ((21 * MaxSelection) / 2) + ((i - 1) * 20)
        RenderTexture Tex_Gui(GuiSelection), X, Y, 0, 0, GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection)), GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection))
        If i > MaxSelection Then
            mText = "Close"
        Else
            If InputData1 = SELECT_POKEMON Then
                mText = Trim$(Pokemon(Player(MyIndex).Pokemon(i).Num).Name)
            ElseIf InputData1 = SELECT_MOVE Then
                mText = Trim$(Moves(Player(MyIndex).Pokemon(InputData2).Moves(i).Num).Name)
            End If
        End If
        If GlobalX >= X + ((107 / 2) - (GetWidth(CurFont, mText) / 2)) And GlobalX <= X + ((107 / 2) - (GetWidth(CurFont, mText) / 2)) + GetWidth(CurFont, mText) And GlobalY >= Y And GlobalY <= Y + 16 Then
            RenderText CurFont, mText, X + ((107 / 2) - (GetWidth(CurFont, mText) / 2)), Y + 3, Yellow
        Else
            RenderText CurFont, mText, X + ((107 / 2) - (GetWidth(CurFont, mText) / 2)), Y + 3, White
        End If
    Next
    
    Exit Sub
errHandler:
    HandleError "DrawSelect", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawStorage()
Dim X As Long, Y As Long, i As Long
Dim mText As String

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Not WindowVisible(WindowType.Main_Storage) Then Exit Sub
    If Not InStorage Then Exit Sub
    
    RenderTexture Tex_Gui(GuiStorage), GuiStorageX, GuiStorageY, 0, 0, GetTextureWidth(Tex_Gui(GuiStorage)), GetTextureHeight(Tex_Gui(GuiStorage)), GetTextureWidth(Tex_Gui(GuiStorage)), GetTextureHeight(Tex_Gui(GuiStorage)), D3DColorARGB(240, 255, 255, 255)
    
    DrawButton ButtonEnum.PCDepositPoke, "Store", White
    DrawButton ButtonEnum.PCClose, "Close", White
    DrawButton ButtonEnum.PCNext, "Next", White
    DrawButton ButtonEnum.PCPrevious, "Previous", White
    
    If GetMaxStoredPokemon > 0 Then
        If GetMaxStoredPokemon >= 12 Then
            For i = StartStorage To StartStorage + 11
                With Player(MyIndex).StoredPokemon(i)
                    If .Num > 0 Then
                        Y = GuiStorageY + 22 + (104 * ((i - StartStorage) \ 4))
                        X = GuiStorageX + 22 + (104 * ((i - StartStorage) Mod 4))
                        
                        RenderTexture Tex_PokeFront(Pokemon(.Num).Pic), X, Y, 0, 0, 100, 100, 100, 100
                        If i = SelStoragePoke Then SelStorageX = X + 70: SelStorageY = Y + 70
                    End If
                End With
            Next
        Else
            For i = 1 To GetMaxStoredPokemon
                With Player(MyIndex).StoredPokemon(i)
                    If .Num > 0 Then
                        Y = GuiStorageY + 22 + (104 * ((i - 1) \ 4))
                        X = GuiStorageX + 22 + (104 * ((i - 1) Mod 4))
                        
                        RenderTexture Tex_PokeFront(Pokemon(.Num).Pic), X, Y, 0, 0, 100, 100, 100, 100
                        If i = SelStoragePoke Then SelStorageX = X + 70: SelStorageY = Y + 70
                    End If
                End With
            Next
        End If
    End If
    If SelStoragePoke > 0 Then
        For i = 0 To 2
            RenderTexture Tex_Gui(GuiSelection), SelStorageX, SelStorageY + (i * 20), 0, 0, GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection)), GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection))
        Next
        X = SelStorageX + ((107 / 2) - (GetWidth(CurFont, "Summary") / 2)): Y = SelStorageY + 2
        If GlobalX >= X And GlobalX <= X + GetWidth(CurFont, "Summary") And GlobalY >= Y And GlobalY <= Y + 16 Then
            RenderText CurFont, "Summary", X, Y, Yellow
        Else
            RenderText CurFont, "Summary", X, Y, White
        End If
        X = SelStorageX + ((107 / 2) - (GetWidth(CurFont, "Withdraw") / 2)): Y = SelStorageY + 22
        If GlobalX >= X And GlobalX <= X + GetWidth(CurFont, "Withdraw") And GlobalY >= Y And GlobalY <= Y + 16 Then
            RenderText CurFont, "Withdraw", X, Y, Yellow
        Else
            RenderText CurFont, "Withdraw", X, Y, White
        End If
        X = SelStorageX + ((107 / 2) - (GetWidth(CurFont, "Release") / 2)): Y = SelStorageY + 42
        If GlobalX >= X And GlobalX <= X + GetWidth(CurFont, "Release") And GlobalY >= Y And GlobalY <= Y + 16 Then
            RenderText CurFont, "Release", X, Y, Yellow
        Else
            RenderText CurFont, "Release", X, Y, White
        End If
    End If

    If ShowStorageSelect Then
        X = GuiStorageX + 28
        For i = 1 To MAX_POKEMON
            Y = GuiStorageY + 208 + ((i - 1) * 20)
            RenderTexture Tex_Gui(GuiSelection), X, Y, 0, 0, GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection)), GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection))
            With Player(MyIndex)
                If .Pokemon(i).Num > 0 Then
                    mText = Trim$(Pokemon(.Pokemon(i).Num).Name)
                    If GlobalX >= X And GlobalX <= X + 107 And GlobalY >= Y + 2 And GlobalY <= Y + 18 Then
                        RenderText CurFont, mText, X + ((GetTextureWidth(Tex_Gui(GuiSelection)) / 2) - (GetWidth(CurFont, mText) / 2)), Y + 2, Yellow
                    Else
                        RenderText CurFont, mText, X + ((GetTextureWidth(Tex_Gui(GuiSelection)) / 2) - (GetWidth(CurFont, mText) / 2)), Y + 2, White
                    End If
                End If
            End With
        Next
    End If

    Exit Sub
errHandler:
    HandleError "DrawStorage", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawShop()
Dim X As Long, Y As Long
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If InShop <= 0 Then Exit Sub
    
    RenderTexture Tex_Gui(GuiShop), GuiShopX, GuiShopY, 0, 0, GetTextureWidth(Tex_Gui(GuiShop)), GetTextureHeight(Tex_Gui(GuiShop)), GetTextureWidth(Tex_Gui(GuiShop)), GetTextureHeight(Tex_Gui(GuiShop)), D3DColorARGB(240, 255, 255, 255)
    RenderText CurFont, Trim$(Shop(InShop).Name), GuiShopX + 10, GuiShopY + 5, White

    DrawButton ButtonEnum.ShopScrollUp
    DrawButton ButtonEnum.ShopScrollDown
    
    For i = ShopStart To ShopStart + 8
        If Shop(InShop).sItem(i).Num > 0 Then
            If ShopSelect = i Then
                RenderText CurFont, i & ": " & Trim$(Item(Shop(InShop).sItem(i).Num).Name), GuiShopX + 20, GuiShopY + 40 + ((i - ShopStart) * 20), Yellow
            Else
                RenderText CurFont, i & ": " & Trim$(Item(Shop(InShop).sItem(i).Num).Name), GuiShopX + 20, GuiShopY + 40 + ((i - ShopStart) * 20), White
            End If
        Else
            RenderText CurFont, i & ": None", GuiShopX + 20, GuiShopY + 40 + ((i - ShopStart) * 20), White
        End If
    Next
    If ShopSelect > 0 Then
        RenderText CurFont, "Price: " & Shop(InShop).sItem(ShopSelect).Price, GuiShopX + (264 / 2) - (GetWidth(CurFont, "Price: " & Shop(InShop).sItem(ShopSelect).Price) / 2), GuiShopY + 235, White
    End If

    X = GuiShopX + 217: Y = GuiShopY + 235
    If GlobalX >= X And GlobalX <= X + GetWidth(CurFont, "Close") And GlobalY >= Y And GlobalY <= Y + 16 Then
        RenderText CurFont, "Close", X, Y, Yellow
    Else
        RenderText CurFont, "Close", X, Y, White
    End If
    X = GuiShopX + 15
    If GlobalX >= X And GlobalX <= X + GetWidth(CurFont, "Buy") And GlobalY >= Y And GlobalY <= Y + 16 Then
        RenderText CurFont, "Buy", X, Y, Yellow
    Else
        RenderText CurFont, "Buy", X, Y, White
    End If
    
    Exit Sub
errHandler:
    HandleError "DrawShop", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawInput()
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If ShowInput Then
        RenderTexture Tex_Gui(GuiCurrency), RenderValX, RenderValY, 0, 0, GetTextureWidth(Tex_Gui(GuiCurrency)), GetTextureHeight(Tex_Gui(GuiCurrency)), GetTextureWidth(Tex_Gui(GuiCurrency)), GetTextureHeight(Tex_Gui(GuiCurrency)), D3DColorARGB(240, 255, 255, 255)
        RenderText CurFont, "Input Value", RenderValX + 10, RenderValY + 5, White
        
        RenderText CurFont, RenderVal & ChatLine, RenderValX + 20, RenderValY + 36, White
        
        If GlobalY >= RenderValY + 60 And GlobalY <= RenderValY + 60 + GetWidth(CurFont, "Close") Then
            If GlobalX >= RenderValX + 20 And GlobalX <= RenderValX + 20 + GetWidth(CurFont, "Confirm") Then
                RenderText CurFont, "Confirm", RenderValX + 20, RenderValY + 60, Yellow
            Else
                RenderText CurFont, "Confirm", RenderValX + 20, RenderValY + 60, White
            End If
            If GlobalX >= RenderValX + 115 And GlobalX <= RenderValX + 115 + GetWidth(CurFont, "Close") Then
                RenderText CurFont, "Close", RenderValX + 115, RenderValY + 60, Yellow
            Else
                RenderText CurFont, "Close", RenderValX + 115, RenderValY + 60, White
            End If
        Else
            RenderText CurFont, "Confirm", RenderValX + 20, RenderValY + 60, White
            RenderText CurFont, "Close", RenderValX + 115, RenderValY + 60, White
        End If
    End If

    Exit Sub
errHandler:
    HandleError "DrawInput", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawTrade()
Dim i As Long, Sprite As Long
Dim y2 As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Not InTrade Then Exit Sub
    
    RenderTexture Tex_Gui(GuiTrade), GuiTradeX, GuiTradeY, 0, 0, GetTextureWidth(Tex_Gui(GuiTrade)), GetTextureHeight(Tex_Gui(GuiTrade)), GetTextureWidth(Tex_Gui(GuiTrade)), GetTextureHeight(Tex_Gui(GuiTrade)), D3DColorARGB(240, 255, 255, 255)
    For i = 1 To MAX_POKEMON
        If Player(MyIndex).Pokemon(i).Num > 0 Then
            Sprite = Pokemon(Player(MyIndex).Pokemon(i).Num).Pic
            If Sprite > 0 Then
                RenderTexture Tex_PokeIcon(Sprite), GuiTradeX + 15 + ((i - 1) * 41), GuiTradeY + 280, (GetTextureWidth(Tex_PokeIcon(Sprite)) / 2) * PokeIconAnim, 0, GetTextureWidth(Tex_PokeIcon(Sprite)) / 2, GetTextureHeight(Tex_PokeIcon(Sprite)), GetTextureWidth(Tex_PokeIcon(Sprite)) / 2, GetTextureHeight(Tex_PokeIcon(Sprite))
            End If
        End If
    Next i
    
    RenderText CurFont, "Trade Window", GuiTradeX + 13, GuiTradeY + 10, White
    
    For i = 1 To MAX_TRADE
        With MyTrade(i)
            Select Case .Type
                Case TRADE_TYPE_ITEM
                    If .ItemNum > 0 And .ItemNum <= Count_Item Then
                        RenderText CurFont, i & ": " & Trim$(Item(.ItemNum).Name) & " x" & .ItemVal, GuiTradeX + 10, GuiTradeY + 40 + ((i - 1) * 20), White
                    Else
                        RenderText CurFont, i & ": ", GuiTradeX + 10, GuiTradeY + 40 + ((i - 1) * 20), White
                    End If
                Case TRADE_TYPE_POKEMON
                    If .Pokemon.Num > 0 Then
                        RenderText CurFont, i & ": " & Trim$(Pokemon(.Pokemon.Num).Name) & " Lv" & .Pokemon.Level, GuiTradeX + 10, GuiTradeY + 40 + ((i - 1) * 20), White
                    Else
                        RenderText CurFont, i & ": ", GuiTradeX + 10, GuiTradeY + 40 + ((i - 1) * 20), White
                    End If
                Case Else
                    RenderText CurFont, i & ": ", GuiTradeX + 10, GuiTradeY + 40 + ((i - 1) * 20), White
            End Select
        End With
    Next
    RenderText CurFont, "Money: ", GuiTradeX + 15, GuiTradeY + 245, White
    
    If ShowTradeSel Then
        For i = 1 To MAX_TRADE
            y2 = GuiTradeY + 40 + ((i - 1) * 20)
            If SelTrade = i Then
                RenderTexture Tex_Gui(GuiSelection), GuiTradeX + 42, y2 + 25, 0, 0, GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection)), GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection))
                If GlobalX >= GuiTradeX + 42 + ((107 / 2) - (GetWidth(CurFont, "Remove") / 2)) And GlobalX <= GuiTradeX + 42 + ((107 / 2) - (GetWidth(CurFont, "Remove") / 2)) + GetWidth(CurFont, "Remove") And GlobalY >= y2 + 28 And GlobalY <= y2 + 28 + 16 Then
                    RenderText CurFont, "Remove", GuiTradeX + 42 + ((107 / 2) - (GetWidth(CurFont, "Remove") / 2)), y2 + 28, Yellow
                Else
                    RenderText CurFont, "Remove", GuiTradeX + 42 + ((107 / 2) - (GetWidth(CurFont, "Remove") / 2)), y2 + 28, White
                End If
            End If
        Next
    End If
    
    DrawButton ButtonEnum.TradeConfirm, "Confirm", White
    
    Exit Sub
errHandler:
    HandleError "DrawTrade", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawTradeConfirm()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Not InTradeConfirm Then Exit Sub
    
    RenderTexture Tex_Gui(GuiTradeConfirm), GuiTradeConfirmX, GuiTradeConfirmY, 0, 0, GetTextureWidth(Tex_Gui(GuiTradeConfirm)), GetTextureHeight(Tex_Gui(GuiTradeConfirm)), GetTextureWidth(Tex_Gui(GuiTradeConfirm)), GetTextureHeight(Tex_Gui(GuiTradeConfirm)), D3DColorARGB(240, 255, 255, 255)
    
    For i = 1 To MAX_TRADE
        With MyTrade(i)
            Select Case .Type
                Case TRADE_TYPE_ITEM
                    If .ItemNum > 0 And .ItemNum <= Count_Item Then
                        RenderText CurFont, i & ": " & Trim$(Item(.ItemNum).Name) & " x" & .ItemVal, GuiTradeConfirmX + 10, GuiTradeConfirmY + 40 + ((i - 1) * 20), White
                    Else
                        RenderText CurFont, i & ": ", GuiTradeConfirmX + 10, GuiTradeConfirmY + 40 + ((i - 1) * 20), White
                    End If
                Case TRADE_TYPE_POKEMON
                    If .Pokemon.Num > 0 Then
                        RenderText CurFont, i & ": " & Trim$(Pokemon(.Pokemon.Num).Name) & " Lv" & .Pokemon.Level, GuiTradeConfirmX + 10, GuiTradeConfirmY + 40 + ((i - 1) * 20), White
                    Else
                        RenderText CurFont, i & ": ", GuiTradeConfirmX + 10, GuiTradeConfirmY + 40 + ((i - 1) * 20), White
                    End If
                Case Else
                    RenderText CurFont, i & ": ", GuiTradeConfirmX + 10, GuiTradeConfirmY + 40 + ((i - 1) * 20), White
            End Select
        End With
        
        If TheirTradeConfirm Then
            If InTradeIndex > 0 Then
                With TheirTrade(i)
                    Select Case .Type
                        Case TRADE_TYPE_ITEM
                            If .ItemNum > 0 And .ItemNum <= Count_Item Then
                                RenderText CurFont, i & ": " & Trim$(Item(.ItemNum).Name) & " x" & .ItemVal, GuiTradeConfirmX + 170, GuiTradeConfirmY + 40 + ((i - 1) * 20), White
                            Else
                                RenderText CurFont, i & ": ", GuiTradeConfirmX + 170, GuiTradeConfirmY + 40 + ((i - 1) * 20), White
                            End If
                        Case TRADE_TYPE_POKEMON
                            If .Pokemon.Num > 0 Then
                                RenderText CurFont, i & ": " & Trim$(Pokemon(.Pokemon.Num).Name) & " Lv" & .Pokemon.Level, GuiTradeConfirmX + 170, GuiTradeConfirmY + 40 + ((i - 1) * 20), White
                            Else
                                RenderText CurFont, i & ": ", GuiTradeConfirmX + 170, GuiTradeConfirmY + 40 + ((i - 1) * 20), White
                            End If
                        Case Else
                            RenderText CurFont, i & ": ", GuiTradeConfirmX + 170, GuiTradeConfirmY + 40 + ((i - 1) * 20), White
                    End Select
                End With
            End If
        End If
    Next
    
    If Not TheirTradeConfirm Then RenderText CurFont, "Waiting for other player..", GuiTradeConfirmX + 170, GuiTradeConfirmY + 80, White

    RenderText CurFont, "Money: ", GuiTradeConfirmX + 15, GuiTradeConfirmY + 245, White
    RenderText CurFont, "Money: ", GuiTradeConfirmX + 175, GuiTradeConfirmY + 245, White
    
    DrawButton ButtonEnum.TradeAccept, "Accept", White
    DrawButton ButtonEnum.TradeDecline, "Decline", White
    
    RenderText CurFont, "Trade Confirmation", GuiTradeConfirmX + 13, GuiTradeConfirmY + 10, White
    
    Exit Sub
errHandler:
    HandleError "DrawTradeConfirm", "modDX8", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub
