Attribute VB_Name = "modGameEditor"
Option Explicit

Public Editor As Byte

Public Const EDITOR_MAP As Byte = 1
Public Const EDITOR_POKEMON As Byte = 2
Public Const EDITOR_MOVE As Byte = 3
Public Const EDITOR_ITEM As Byte = 4
Public Const EDITOR_NPC As Byte = 5
Public Const EDITOR_SHOP As Byte = 6

Public CurTileset As Long
Public CurLayer As Byte
Public CurAttribute As Byte

Public EditorTileX As Long
Public EditorTileY As Long
Public EditorTileWidth As Long
Public EditorTileHeight As Long

Public EditorIndex As Long
Public Data1Index As Long

Public EditorData1 As Long
Public EditorData2 As Long
Public EditorData3 As Long
Public EditorData4 As String

Public InEditorInit As Boolean

Public Pokemon_Changed(1 To Count_Pokemon) As Boolean
Public Move_Changed(1 To Count_Move) As Boolean
Public Item_Changed(1 To Count_Item) As Boolean
Public NPC_Changed(1 To Count_NPC) As Boolean
Public Shop_Changed(1 To Count_Shop) As Boolean

Public Sub MapEditorInit()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Editor = EDITOR_MAP
    
    frmMapEditor.Show
    
    frmMapEditor.cmbTileset.Clear
    For i = 1 To Count_Tileset
        frmMapEditor.cmbTileset.AddItem "Tileset #: " & i
    Next
    frmMapEditor.cmbTileset.ListIndex = 0
    CurTileset = frmMapEditor.cmbTileset.ListIndex + 1
    
    CurLayer = Layers.Ground
    CurAttribute = 0
    frmMapEditor.optLayers(CurLayer).value = True
    frmMapEditor.SSTab1.Tab = 0
    
    If Not gTexture(Tex_Tileset(CurTileset)).loaded Then Call LoadTexture(Tex_Tileset(CurTileset))
    frmMapEditor.scrlPictureY.max = (GetTextureHeight(Tex_Tileset(CurTileset)) \ Pic_Size) - (frmMapEditor.picTileset.Height \ Pic_Size)
    frmMapEditor.scrlPictureX.max = (GetTextureWidth(Tex_Tileset(CurTileset)) \ Pic_Size) - (frmMapEditor.picTileset.Width \ Pic_Size)
    MapEditorTileScroll
    
    EditorTileX = 0
    EditorTileY = 0
    EditorTileWidth = 1
    EditorTileHeight = 1
    
    Exit Sub
errHandler:
    HandleError "MapEditorInit", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorTileScroll()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If GetTextureWidth(Tex_Tileset(CurTileset)) < frmMapEditor.picTileset.Width Then
        frmMapEditor.scrlPictureX.Enabled = False
    Else
        frmMapEditor.scrlPictureX.Enabled = True
    End If
    If GetTextureHeight(Tex_Tileset(CurTileset)) < frmMapEditor.picTileset.Height Then
        frmMapEditor.scrlPictureY.Enabled = False
    Else
        frmMapEditor.scrlPictureY.Enabled = True
    End If
    
    Exit Sub
errHandler:
    HandleError "MapEditorTileScroll", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorChooseTile(Button As Integer, X As Single, Y As Single)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Button = vbLeftButton Then
        EditorTileWidth = 1
        EditorTileHeight = 1
        
        EditorTileX = X \ Pic_Size
        EditorTileY = Y \ Pic_Size
    End If
    
    Exit Sub
errHandler:
    HandleError "MapEditorChooseTile", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorDrag(Button As Integer, X As Single, Y As Single)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Button = vbLeftButton Then
        X = (X \ Pic_Size) + 1
        Y = (Y \ Pic_Size) + 1

        If X < 0 Then X = 0
        If X > GetTextureWidth(Tex_Tileset(CurTileset)) / Pic_Size Then X = GetTextureWidth(Tex_Tileset(CurTileset)) / Pic_Size
        If Y < 0 Then Y = 0
        If Y > GetTextureHeight(Tex_Tileset(CurTileset)) / Pic_Size Then Y = GetTextureHeight(Tex_Tileset(CurTileset)) / Pic_Size
        If X > EditorTileX Then EditorTileWidth = X - EditorTileX
        If Y > EditorTileY Then EditorTileHeight = Y - EditorTileY
    End If
    
    Exit Sub
errHandler:
    HandleError "MapEditorDrag", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorSetTile(ByVal X As Long, ByVal Y As Long, Optional ByVal multitile As Boolean = False)
Dim x2 As Long, y2 As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Not multitile Then
        With Map.Tile(X, Y)
            .Layer(CurLayer).X = EditorTileX
            .Layer(CurLayer).Y = EditorTileY
            .Layer(CurLayer).Tileset = CurTileset
        End With
    Else
        y2 = 0
        For Y = CurY To CurY + EditorTileHeight - 1
            x2 = 0
            For X = CurX To CurX + EditorTileWidth - 1
                If X >= 0 And X <= Map.MaxX Then
                    If Y >= 0 And Y <= Map.MaxY Then
                        With Map.Tile(X, Y)
                            .Layer(CurLayer).X = EditorTileX + x2
                            .Layer(CurLayer).Y = EditorTileY + y2
                            .Layer(CurLayer).Tileset = CurTileset
                        End With
                    End If
                End If
                x2 = x2 + 1
            Next
            y2 = y2 + 1
        Next
    End If
    
    Exit Sub
errHandler:
    HandleError "MapEditorSetTile", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorMouseDown(ByVal Button As Integer, ByVal X As Long, ByVal Y As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Not isInBounds Then Exit Sub
    
    If Button = vbLeftButton Then
        If frmMapEditor.SSTab1.Tab = 0 Then
            If EditorTileWidth = 1 And EditorTileHeight = 1 Then
                MapEditorSetTile CurX, CurY
            Else
                MapEditorSetTile CurX, CurY, True
            End If
        ElseIf frmMapEditor.SSTab1.Tab = 1 Then
            With Map.Tile(CurX, CurY)
                .Type = CurAttribute
                .Data1 = EditorData1
                .Data2 = EditorData2
                .Data3 = EditorData3
                .Data4 = EditorData4
            End With
        End If
    End If

    If Button = vbRightButton Then
        If frmMapEditor.SSTab1.Tab = 0 Then
            With Map.Tile(CurX, CurY)
                .Layer(CurLayer).X = 0
                .Layer(CurLayer).Y = 0
                .Layer(CurLayer).Tileset = 0
            End With
        ElseIf frmMapEditor.SSTab1.Tab = 1 Then
            With Map.Tile(CurX, CurY)
                .Type = 0
                .Data1 = 0
                .Data2 = 0
                .Data3 = 0
                .Data4 = vbNullString
            End With
        End If
    End If
    
    Exit Sub
errHandler:
    HandleError "MapEditorMouseDown", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function isInBounds()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If (CurX >= 0) Then
        If (CurX <= Map.MaxX) Then
            If (CurY >= 0) Then
                If (CurY <= Map.MaxY) Then
                    isInBounds = True
                End If
            End If
        End If
    End If
    
    Exit Function
errHandler:
    HandleError "isInBounds", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub InitProperties()
Dim i As Long
Dim tmpString() As String
Dim pIndex As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    With frmMapProperties
        .txtName.Text = Trim$(Map.Name)
        
        .lstMusic.Clear
        .lstMusic.AddItem "None."
        For i = 1 To UBound(musicCache)
            .lstMusic.AddItem musicCache(i)
        Next i
        
        If .lstMusic.ListCount >= 0 Then
            .lstMusic.ListIndex = 0
            For i = 0 To .lstMusic.ListCount
                If .lstMusic.List(i) = Trim$(Map.Music) Then
                    .lstMusic.ListIndex = i
                End If
            Next
        End If
        
        .cmbMoral.ListIndex = Map.Moral
        .txtMaxX.Text = Map.MaxX
        .txtMaxY.Text = Map.MaxY
        For i = 0 To 3
            .txtLink(i).Text = Map.Link(i)
        Next i
        
        .lstPokemon.Clear
        For i = 1 To MAX_MAP_POKEMON
            If Map.Pokemon(i) > 0 Then
                .lstPokemon.AddItem i & ": " & Trim$(Pokemon(Map.Pokemon(i)).Name)
            Else
                .lstPokemon.AddItem i & ": None"
            End If
        Next i
        .lstPokemon.ListIndex = 0
        .lstNPC.Clear
        For i = 1 To MAX_MAP_NPC
            If Map.NPC(i) > 0 Then
                .lstNPC.AddItem i & ": " & Trim$(NPC(Map.NPC(i)).Name)
            Else
                .lstNPC.AddItem i & ": None"
            End If
        Next i
        .lstNPC.ListIndex = 0
        
        .cmbPokemon.Clear
        .cmbPokemon.AddItem "None"
        For i = 1 To Count_Pokemon
            .cmbPokemon.AddItem i & ": " & Trim$(Pokemon(i).Name)
        Next i
        .cmbNPC.Clear
        .cmbNPC.AddItem "None"
        For i = 1 To Count_NPC
            .cmbNPC.AddItem i & ": " & Trim$(NPC(i).Name)
        Next i
        
        tmpString = Split(.lstPokemon.List(.lstPokemon.ListIndex))
        pIndex = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
        .cmbPokemon.ListIndex = Map.Pokemon(pIndex)
        
        tmpString = Split(.lstNPC.List(.lstNPC.ListIndex))
        pIndex = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
        .cmbNPC.ListIndex = Map.NPC(pIndex)
        
        .txtLvlMin = Map.MinLvl
        .txtLvlMax = Map.MaxLvl
        
        If Map.CurField > Count_FieldFront Then
            Map.CurField = Count_FieldFront
        End If
        .scrlField = Map.CurField
        If Map.CurBack > Count_Background Then
            Map.CurBack = Count_Background
        End If
        .scrlBack.value = Map.CurBack
        
        .Show vbModal
    End With
    
    Exit Sub
errHandler:
    HandleError "InitProperties", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorCancel()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Editor = 0
    Unload frmMapEditor
    SendNeedMap YES
    
    Exit Sub
errHandler:
    HandleError "MapEditorCancel", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorSend()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Map.Rev >= MAX_LONG Then
        Map.Rev = 0
    Else
        Map.Rev = Map.Rev + 1
    End If
    Call SendMap
    Editor = 0
    Unload frmMapEditor
    
    Exit Sub
errHandler:
    HandleError "MapEditorSend", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorFillLayer()
Dim X As Long, Y As Long
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If MsgBox("Are you sure you want to fill this layer?", vbYesNo) = vbYes Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                With Map.Tile(X, Y).Layer(CurLayer)
                    .Tileset = CurTileset
                    .X = EditorTileX
                    .Y = EditorTileY
                End With
            Next
        Next
    End If
    
    Exit Sub
errHandler:
    HandleError "MapEditorFillLayer", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub MapEditorClearLayer()
Dim X As Long, Y As Long
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If MsgBox("Are you sure you want to clear this layer?", vbYesNo) = vbYes Then
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                With Map.Tile(X, Y).Layer(CurLayer)
                    .Tileset = 0
                    .X = 0
                    .Y = 0
                End With
            Next
        Next
    End If
    
    Exit Sub
errHandler:
    HandleError "MapEditorClearLayer", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearAttributeDialogue()
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    frmMapEditor.fraAttribute.Visible = False
    frmMapEditor.fraWarp.Visible = False
    frmMapEditor.fraShop.Visible = False
    
    EditorData1 = 0
    EditorData2 = 0
    EditorData3 = 0
    EditorData4 = vbNullString
    
    Exit Sub
errHandler:
    HandleError "ClearAttributeDialogue", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub PokemonEditorInit()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    EditorIndex = frmPokemonEditor.lstIndex.ListIndex + 1
    
    With frmPokemonEditor
        .txtName.Text = Trim$(Pokemon(EditorIndex).Name)
        .scrlPokeIcon.value = Pokemon(EditorIndex).Pic
        .txtFemaleRate.Text = Pokemon(EditorIndex).FemaleRate
        .txtBaseExp.Text = Pokemon(EditorIndex).BaseExp
        For i = 1 To Stats.Stat_Count - 1
            .txtBaseStat(i).Text = Pokemon(EditorIndex).BaseStat(i)
        Next i
        
        .cmbMoveNum.Clear
        .cmbMoveNum.AddItem "None."
        For i = 1 To Count_Move
            .cmbMoveNum.AddItem i & ": " & Trim$(Moves(i).Name)
        Next
        .lstMove.Clear
        For i = 1 To MAX_MOVES
            If Pokemon(EditorIndex).MoveNum(i) > 0 Then
                .lstMove.AddItem i & ": " & Trim$(Moves(Pokemon(EditorIndex).MoveNum(i)).Name) & " - Lv" & Pokemon(EditorIndex).MoveLevel(i)
            Else
                .lstMove.AddItem i & ": None."
            End If
        Next
        .lstMove.ListIndex = 0
        Data1Index = .lstMove.ListIndex + 1
        
        .cmbMoveNum.ListIndex = Pokemon(EditorIndex).MoveNum(Data1Index)
        .txtMoveLevel.Text = Pokemon(EditorIndex).MoveLevel(Data1Index)
        
        .cmbPType.ListIndex = Pokemon(EditorIndex).pType
        .cmbSType.ListIndex = Pokemon(EditorIndex).sType
        
        .cmbEvolve.Clear
        .cmbEvolve.AddItem "None."
        For i = 1 To Count_Pokemon
            .cmbEvolve.AddItem i & ": " & Trim$(Pokemon(i).Name)
        Next
        .cmbEvolve.ListIndex = Pokemon(EditorIndex).EvolveNum
        .txtEvolveLvl.Text = Pokemon(EditorIndex).EvolveLvl
        
        .txtCatchRate.Text = Pokemon(EditorIndex).CatchRate
    End With
    
    Pokemon_Changed(EditorIndex) = True
    
    Exit Sub
errHandler:
    HandleError "PokemonEditorInit", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub PokemonEditorOk()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_Pokemon
        If Pokemon_Changed(i) Then
            Call SendSavePokemon(i)
        End If
    Next
    
    Unload frmPokemonEditor
    Editor = 0
    ClearPokemonChange
    
    Exit Sub
errHandler:
    HandleError "PokemonEditorOk", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub PokemonEditorCancel()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Unload frmPokemonEditor
    Editor = 0
    ClearPokemonChange
    ClearPokemons
    SendRequestPokemons
    
    Exit Sub
errHandler:
    HandleError "PokemonEditorCancel", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearPokemonChange()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ZeroMemory Pokemon_Changed(1), Count_Pokemon * 2
    
    Exit Sub
errHandler:
    HandleError "ClearPokemonChange", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub MoveEditorInit()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    EditorIndex = frmMoveEditor.lstIndex.ListIndex + 1
    
    With frmMoveEditor
        .txtName.Text = Trim$(Moves(EditorIndex).Name)
        .txtPower.Text = Moves(EditorIndex).Power
        .txtPP.Text = Moves(EditorIndex).PP
        .cmbPType.ListIndex = Moves(EditorIndex).Type
        .cmbAtkType.ListIndex = Moves(EditorIndex).AtkType
    End With
    
    Move_Changed(EditorIndex) = True
    
    Exit Sub
errHandler:
    HandleError "MoveEditorInit", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub MoveEditorOk()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_Move
        If Move_Changed(i) Then
            Call SendSaveMove(i)
        End If
    Next
    
    Unload frmMoveEditor
    Editor = 0
    ClearMoveChange
    
    Exit Sub
errHandler:
    HandleError "MoveEditorOk", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub MoveEditorCancel()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Unload frmMoveEditor
    Editor = 0
    ClearMoveChange
    ClearMoves
    SendRequestMoves
    
    Exit Sub
errHandler:
    HandleError "MoveEditorCancel", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearMoveChange()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ZeroMemory Move_Changed(1), Count_Move * 2
    
    Exit Sub
errHandler:
    HandleError "ClearMoveChange", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ItemEditorInit()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    EditorIndex = frmItemEditor.lstIndex.ListIndex + 1
    
    InEditorInit = True
    With frmItemEditor
        .txtName.Text = Trim$(Item(EditorIndex).Name)
        .scrlItemPic.value = Item(EditorIndex).Pic
        .txtDesc.Text = Trim$(Item(EditorIndex).Desc)
        
        .cmbItemType.ListIndex = Item(EditorIndex).Type
        
        If Item(EditorIndex).Type = ItemType.Items Then
            .fraItem.Visible = True
            .cmbType.ListIndex = Item(EditorIndex).IType
            If Item(EditorIndex).IType = ItemProperties.RestoreHP Or Item(EditorIndex).IType = ItemProperties.RestorePP Then
                .fraValue.Visible = True
                .txtValue.Text = Item(EditorIndex).Data2
            Else
                .fraValue.Visible = False
            End If
        Else
            .fraItem.Visible = False
        End If
        
        If Item(EditorIndex).Type = ItemType.Pokeballs Then
            .fraPokeBall.Visible = True
            .txtCatchRate.Text = Item(EditorIndex).Data3
        Else
            .fraPokeBall.Visible = False
        End If
        
        .txtSellPrice.Text = Item(EditorIndex).Sell
    End With
    InEditorInit = False
    
    Item_Changed(EditorIndex) = True
    
    Exit Sub
errHandler:
    HandleError "ItemEditorInit", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ItemEditorOk()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_Item
        If Item_Changed(i) Then
            Call SendSaveItem(i)
        End If
    Next
    
    Unload frmItemEditor
    Editor = 0
    ClearItemChange
    
    Exit Sub
errHandler:
    HandleError "ItemEditorOk", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ItemEditorCancel()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Unload frmItemEditor
    Editor = 0
    ClearItemChange
    ClearItems
    SendRequestItems
    
    Exit Sub
errHandler:
    HandleError "ItemEditorCancel", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearItemChange()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ZeroMemory Item_Changed(1), Count_Item * 2
    
    Exit Sub
errHandler:
    HandleError "ClearItemChange", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub NPCEditorInit()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    EditorIndex = frmNPCEditor.lstIndex.ListIndex + 1
    
    With frmNPCEditor
        .txtName.Text = Trim$(NPC(EditorIndex).Name)
        .scrlSprite.value = NPC(EditorIndex).Sprite
    End With
    
    NPC_Changed(EditorIndex) = True
    
    Exit Sub
errHandler:
    HandleError "NPCEditorInit", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub NPCEditorOk()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_NPC
        If NPC_Changed(i) Then
            Call SendSaveNPC(i)
        End If
    Next
    
    Unload frmNPCEditor
    Editor = 0
    ClearNPCChange
    
    Exit Sub
errHandler:
    HandleError "NPCEditorOk", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub NPCEditorCancel()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Unload frmNPCEditor
    Editor = 0
    ClearNPCChange
    ClearNPCs
    SendRequestNPCs
    
    Exit Sub
errHandler:
    HandleError "NPCEditorCancel", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearNPCChange()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ZeroMemory NPC_Changed(1), Count_NPC * 2
    
    Exit Sub
errHandler:
    HandleError "ClearNPCChange", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ShopEditorInit()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    EditorIndex = frmShopEditor.lstIndex.ListIndex + 1
    
    With frmShopEditor
        .txtName.Text = Trim$(Shop(EditorIndex).Name)
        
        .cmbItemNum.Clear
        .cmbItemNum.AddItem "None."
        For i = 1 To Count_Item
            .cmbItemNum.AddItem i & ": " & Trim$(Item(i).Name)
        Next
        .lstItem.Clear
        For i = 1 To MAX_SHOP_ITEMS
            If Shop(EditorIndex).sItem(i).Num > 0 Then
                .lstItem.AddItem i & ": " & Trim$(Item(Shop(EditorIndex).sItem(i).Num).Name)
            Else
                .lstItem.AddItem i & ": None"
            End If
        Next
        .lstItem.ListIndex = -1
    End With
    
    Shop_Changed(EditorIndex) = True
    
    Exit Sub
errHandler:
    HandleError "ShopEditorInit", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ShopEditorOk()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_Shop
        If Shop_Changed(i) Then
            Call SendSaveShop(i)
        End If
    Next
    
    Unload frmShopEditor
    Editor = 0
    ClearShopChange
    
    Exit Sub
errHandler:
    HandleError "ShopEditorOk", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ShopEditorCancel()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Unload frmShopEditor
    Editor = 0
    ClearShopChange
    ClearShops
    SendRequestShops
    
    Exit Sub
errHandler:
    HandleError "ShopEditorCancel", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearShopChange()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ZeroMemory Shop_Changed(1), Count_Shop * 2
    
    Exit Sub
errHandler:
    HandleError "ClearShopChange", "modGameEditor", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub
