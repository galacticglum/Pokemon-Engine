Attribute VB_Name = "modHandleData"
Option Explicit

Private Function GetAddress(FunAddr As Long) As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    GetAddress = FunAddr
    
    Exit Function
errHandler:
    HandleError "GetAddress", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub InitMessages()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    HandleDataSub(SAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SCharSelect) = GetAddress(AddressOf HandleCharSelect)
    HandleDataSub(SIndex) = GetAddress(AddressOf HandleIndex)
    HandleDataSub(SHighIndex) = GetAddress(AddressOf HandleHighIndex)
    HandleDataSub(SInGame) = GetAddress(AddressOf HandleInGame)
    HandleDataSub(SPlayerXY) = GetAddress(AddressOf HandlePlayerXY)
    HandleDataSub(SPlayerXYToMap) = GetAddress(AddressOf HandlePlayerXYToMap)
    HandleDataSub(SLeaveMap) = GetAddress(AddressOf HandleLeaveMap)
    HandleDataSub(SCheckForMap) = GetAddress(AddressOf HandleCheckForMap)
    HandleDataSub(SMap) = GetAddress(AddressOf HandleMap)
    HandleDataSub(SMapDone) = GetAddress(AddressOf HandleMapDone)
    HandleDataSub(SPlayerData) = GetAddress(AddressOf HandlePlayerData)
    HandleDataSub(SPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(SPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(SLeft) = GetAddress(AddressOf HandleLeft)
    HandleDataSub(SMsg) = GetAddress(AddressOf HandleMsg)
    HandleDataSub(SEditMap) = GetAddress(AddressOf HandleEditMap)
    HandleDataSub(SEditPokemon) = GetAddress(AddressOf HandleEditPokemon)
    HandleDataSub(SUpdatePokemon) = GetAddress(AddressOf HandleUpdatePokemon)
    HandleDataSub(SPlayerPokemon) = GetAddress(AddressOf HandlePlayerPokemon)
    HandleDataSub(SBattle) = GetAddress(AddressOf HandleBattle)
    HandleDataSub(SEnemyPokemon) = GetAddress(AddressOf HandleEnemyPokemon)
    HandleDataSub(SBattleMsg) = GetAddress(AddressOf HandleBattleMsg)
    HandleDataSub(SPlayMusic) = GetAddress(AddressOf HandlePlayMusic)
    HandleDataSub(SPlaySound) = GetAddress(AddressOf HandlePlaySound)
    HandleDataSub(SEditMove) = GetAddress(AddressOf HandleEditMove)
    HandleDataSub(SUpdateMove) = GetAddress(AddressOf HandleUpdateMove)
    HandleDataSub(SExpCalc) = GetAddress(AddressOf HandleExpCalc)
    HandleDataSub(STarget) = GetAddress(AddressOf HandleTarget)
    HandleDataSub(SUpdatePokemonVital) = GetAddress(AddressOf HandleUpdatePokemonVital)
    HandleDataSub(SUpdateEnemyVital) = GetAddress(AddressOf HandleUpdateEnemyVital)
    HandleDataSub(SBattleResult) = GetAddress(AddressOf HandleBattleResult)
    HandleDataSub(SExitBattle) = GetAddress(AddressOf HandleExitBattle)
    HandleDataSub(SForceSwitch) = GetAddress(AddressOf HandleForceSwitch)
    HandleDataSub(SSwitch) = GetAddress(AddressOf HandleSwitch)
    HandleDataSub(SLearnMove) = GetAddress(AddressOf HandleLearnMove)
    HandleDataSub(SEvolve) = GetAddress(AddressOf HandleEvolve)
    HandleDataSub(SEditItem) = GetAddress(AddressOf HandleEditItem)
    HandleDataSub(SUpdateItem) = GetAddress(AddressOf HandleUpdateItem)
    HandleDataSub(SInventory) = GetAddress(AddressOf HandleInventory)
    HandleDataSub(SSelect) = GetAddress(AddressOf HandleSelect)
    HandleDataSub(SBattleRequest) = GetAddress(AddressOf HandleBattleRequest)
    HandleDataSub(SCaptured) = GetAddress(AddressOf HandleCaptured)
    HandleDataSub(SPlayerStoredPokemon) = GetAddress(AddressOf HandlePlayerStoredPokemon)
    HandleDataSub(SStorage) = GetAddress(AddressOf HandleStorage)
    HandleDataSub(SEditNPC) = GetAddress(AddressOf HandleEditNPC)
    HandleDataSub(SUpdateNPC) = GetAddress(AddressOf HandleUpdateNPC)
    HandleDataSub(SEditShop) = GetAddress(AddressOf HandleEditShop)
    HandleDataSub(SUpdateShop) = GetAddress(AddressOf HandleUpdateShop)
    HandleDataSub(SOpenShop) = GetAddress(AddressOf HandleOpenShop)
    
    HandleDataSub(SSpawnNpc) = GetAddress(AddressOf HandleSpawnNpc)
    HandleDataSub(SNPCClear) = GetAddress(AddressOf HandleNpcClear)
    HandleDataSub(SNpcMove) = GetAddress(AddressOf HandleNpcMove)
    HandleDataSub(SMapNpcData) = GetAddress(AddressOf HandleMapNpcData)
    HandleDataSub(SNpcDir) = GetAddress(AddressOf HandleNpcDir)
    
    HandleDataSub(STradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(STrade) = GetAddress(AddressOf HandleTrade)
    HandleDataSub(SCloseTrade) = GetAddress(AddressOf HandleCloseTrade)
    HandleDataSub(STradeConfirm) = GetAddress(AddressOf HandleTradeConfirm)
    
    Exit Sub
errHandler:
    HandleError "InitMessages", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleData(ByRef Data() As Byte)
Dim Buffer As clsBuffer
Dim MsgType As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MsgType = Buffer.ReadLong
    
    If MsgType < 0 Then
        CloseApp
        Exit Sub
    End If
    If MsgType >= ServerPacket_Count Then
        CloseApp
        Exit Sub
    End If
    
    CallWindowProc HandleDataSub(MsgType), MyIndex, Buffer.ReadBytes(Buffer.length), 0, 0
    
    Exit Sub
errHandler:
    HandleError "HandleData", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub IncomingData(ByVal DataLength As Long)
Dim Buffer() As Byte
Dim pLength As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    frmMain.Socket.GetData Buffer, vbUnicode, DataLength
    
    PlayerBuffer.WriteBytes Buffer()
    
    If PlayerBuffer.length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Do While pLength > 0 And pLength <= PlayerBuffer.length - 4
        If pLength <= PlayerBuffer.length - 4 Then
            PlayerBuffer.ReadLong
            HandleData PlayerBuffer.ReadBytes(pLength)
        End If

        pLength = 0
        If PlayerBuffer.length >= 4 Then pLength = PlayerBuffer.ReadLong(False)
    Loop
    PlayerBuffer.Trim
    DoEvents
    
    Exit Sub
errHandler:
    HandleError "IncomingData", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAlertMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Msg As String
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    Set Buffer = Nothing
    
    Call MsgBox(Msg, vbOKOnly, GameTitle)
    LogOutGame
    
    Exit Sub
errHandler:
    HandleError "HandleAlertMsg", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCharSelect(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim X As Byte
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Options.Username = SaveUser
    If SaveAccount Then
        Options.Password = SavePass
        Options.SavePass = 1
    Else
        Options.Password = vbNullString
        Options.SavePass = 0
    End If
    SaveOption
    
    SaveUser = vbNullString
    SavePass = vbNullString
    
    For X = 1 To MAX_PLAYER_DATA
        CharSelectName(X) = vbNullString
        CharSelectName(X) = Trim$(Buffer.ReadString)
        CharSelectSprite(X) = 0
        CharSelectSprite(X) = Buffer.ReadLong
    Next
    OpenWindow Menu_CharSelect
    Set Buffer = Nothing

    Exit Sub
errHandler:
    HandleError "HandleCharSelect", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleIndex(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MyIndex = Buffer.ReadLong
    HighPlayerIndex = Buffer.ReadLong
    Set Buffer = Nothing

    Exit Sub
errHandler:
    HandleError "HandleIndex", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleHighIndex(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    HighPlayerIndex = Buffer.ReadLong
    Set Buffer = Nothing

    Exit Sub
errHandler:
    HandleError "HandleHighIndex", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleInGame(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    InitGame

    Exit Sub
errHandler:
    HandleError "HandleInGame", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerXY(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Call SetPlayerX(MyIndex, Buffer.ReadLong)
    Call SetPlayerY(MyIndex, Buffer.ReadLong)
    Call SetPlayerDir(MyIndex, Buffer.ReadByte)
    Player(MyIndex).Moving = 0
    Player(MyIndex).xOffset = 0
    Player(MyIndex).yOffset = 0
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandlePlayerXY", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerXYToMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    Call SetPlayerX(i, Buffer.ReadLong)
    Call SetPlayerY(i, Buffer.ReadLong)
    Call SetPlayerDir(i, Buffer.ReadByte)
    Set Buffer = Nothing
    Player(i).Moving = 0
    Player(i).xOffset = 0
    Player(i).yOffset = 0

    Exit Sub
errHandler:
    HandleError "HandleInGame", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleLeaveMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Call ClearPlayer(Buffer.ReadLong)
    Set Buffer = Nothing

    Exit Sub
errHandler:
    HandleError "HandleLeaveMap", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCheckForMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long, MapNum As Long, Rev As Long, NeedMap As Long
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    GettingMap = True
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To HighPlayerIndex
        If i <> MyIndex Then
            SetPlayerMap i, 0
        End If
    Next
    
    ' Clear Other Map Data
    ClearMap
    
    MapNum = Buffer.ReadLong
    Rev = Buffer.ReadLong
    
    NeedMap = YES
    Call ClearMapNpcs
    If FileExist(App.Path & "\bin\_maps\" & MapNum & "_cache.dat") Then
        Call LoadMap(MapNum)
        If Map.Rev <> Rev Then
            NeedMap = YES
        Else
            NeedMap = NO
        End If
    End If
    
    SendNeedMap NeedMap
    
    If Editor = EDITOR_MAP Then
        Editor = 0
        Unload frmMapEditor
        
        ClearAttributeDialogue

        If frmMapProperties.Visible Then
            frmMapProperties.Visible = False
        End If
    End If

    Exit Sub
errHandler:
    HandleError "HandleCheckForMap", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long, X As Long, Y As Long
Dim MapNum As Long
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MapNum = Buffer.ReadLong
    Map.Name = Trim$(Buffer.ReadString)
    Map.Music = Trim$(Buffer.ReadString)
    Map.Rev = Buffer.ReadLong
    Map.Moral = Buffer.ReadByte
    Map.MaxX = Buffer.ReadLong
    Map.MaxY = Buffer.ReadLong
    
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            With Map.Tile(X, Y)
                For i = 0 To Layers.LayerCount - 1
                    .Layer(i).Tileset = Buffer.ReadLong
                    .Layer(i).X = Buffer.ReadLong
                    .Layer(i).Y = Buffer.ReadLong
                Next i
                .Type = Buffer.ReadByte
                .Data1 = Buffer.ReadLong
                .Data2 = Buffer.ReadLong
                .Data3 = Buffer.ReadLong
                .Data4 = Buffer.ReadString
            End With
        Next
    Next
    
    For i = 0 To 3
        Map.Link(i) = Buffer.ReadLong
    Next i
    
    For i = 1 To MAX_MAP_POKEMON
        Map.Pokemon(i) = Buffer.ReadLong
    Next i
    
    Map.MinLvl = Buffer.ReadLong
    Map.MaxLvl = Buffer.ReadLong
    
    For i = 1 To MAX_MAP_NPC
        Map.NPC(i) = Buffer.ReadLong
    Next i
    
    Map.CurField = Buffer.ReadLong
    Map.CurBack = Buffer.ReadLong
    
    Set Buffer = Nothing
    SaveMap MapNum
    
    If Editor = EDITOR_MAP Then
        Editor = 0
        Unload frmMapEditor
        
        ClearAttributeDialogue

        If frmMapProperties.Visible Then
            frmMapProperties.Visible = False
        End If
    End If
    
    Exit Sub
errHandler:
    HandleError "HandleMap", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapDone(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    GettingMap = False
    CanMoveNow = True
    ShowTitleBar = YES
    TitleBarAlpha = 0
    ChangeTitleBar = 0
    
    If Not Trim$(Map.Music) = "None." Then
        PlayMusic Trim$(Map.Music)
    Else
        StopMusic
    End If

    Exit Sub
errHandler:
    HandleError "HandleMapDone", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    Player(i).Name = Trim$(Buffer.ReadString)
    Player(i).Gender = Buffer.ReadByte
    Player(i).Access = Buffer.ReadByte
    SetPlayerSprite i, Buffer.ReadLong
    
    SetPlayerMap i, Buffer.ReadLong
    SetPlayerX i, Buffer.ReadLong
    SetPlayerY i, Buffer.ReadLong
    SetPlayerDir i, Buffer.ReadByte
    
    With Player(i).PvP
        .win = Buffer.ReadLong
        .Lose = Buffer.ReadLong
        .Disconnect = Buffer.ReadLong
    End With
    
    Player(i).Money = Buffer.ReadLong
    
    Player(i).IsVIP = Buffer.ReadByte
    Set Buffer = Nothing
    
    Player(i).Moving = 0
    Player(i).xOffset = 0
    Player(i).yOffset = 0

    Exit Sub
errHandler:
    HandleError "HandlePlayerData", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    Call SetPlayerDir(i, Buffer.ReadByte)
    Set Buffer = Nothing
    
    Player(i).xOffset = 0
    Player(i).yOffset = 0
    Player(i).Moving = 0

    Exit Sub
errHandler:
    HandleError "HandlePlayerData", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    Call SetPlayerX(i, Buffer.ReadLong)
    Call SetPlayerY(i, Buffer.ReadLong)
    Call SetPlayerDir(i, Buffer.ReadByte)
    Player(i).xOffset = 0
    Player(i).yOffset = 0
    Player(i).Moving = Buffer.ReadByte
    Set Buffer = Nothing
    
    Select Case GetPlayerDir(i)
        Case DIR_UP
            Player(i).yOffset = Pic_Size
        Case DIR_DOWN
            Player(i).yOffset = Pic_Size * -1
        Case DIR_LEFT
            Player(i).xOffset = Pic_Size
        Case DIR_RIGHT
            Player(i).xOffset = Pic_Size * -1
    End Select
    
    Exit Sub
errHandler:
    HandleError "HandlePlayerMove", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleLeft(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Call ClearPlayer(Buffer.ReadLong)
    Set Buffer = Nothing

    Exit Sub
errHandler:
    HandleError "HandleLeft", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Call AddText(Buffer.ReadString, CheckColor(Buffer.ReadLong))
    Set Buffer = Nothing

    Exit Sub
errHandler:
    HandleError "HandleMsg", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEditMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If GetPlayerAccess(MyIndex) < ACCESS_MAPPER Then Exit Sub
    
    MapEditorInit
    
    Exit Sub
errHandler:
    HandleError "HandleEditMap", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEditPokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Editor = 0 Then
        With frmPokemonEditor
            Editor = EDITOR_POKEMON
            
            .lstIndex.Clear
            For i = 1 To Count_Pokemon
                .lstIndex.AddItem i & ": " & Trim$(Pokemon(i).Name)
            Next
            .Show
            .lstIndex.ListIndex = 0
            
            PokemonEditorInit
        End With
    End If
    
    Exit Sub
errHandler:
    HandleError "HandleEditPokemon", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdatePokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long
Dim PokemonSize As Long
Dim PokemonData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    PokemonSize = LenB(Pokemon(n))
    ReDim PokemonData(PokemonSize - 1)
    PokemonData = Buffer.ReadBytes(PokemonSize)
    CopyMemory ByVal VarPtr(Pokemon(n)), ByVal VarPtr(PokemonData(0)), PokemonSize
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandleUpdatePokemon", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerPokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Slot As Byte
Dim i As Long
Dim X As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    X = Buffer.ReadLong
    Slot = Buffer.ReadByte
    With Player(X).Pokemon(Slot)
        .Num = Buffer.ReadLong
        .Gender = Buffer.ReadByte
        .CurHP = Buffer.ReadLong
        .Level = Buffer.ReadLong
        For i = 1 To Stats.Stat_Count - 1
            .Stat(i) = Buffer.ReadLong
            .StatIV(i) = Buffer.ReadLong
            .StatEV(i) = Buffer.ReadLong
        Next i
        .Exp = Buffer.ReadLong
        For i = 1 To MAX_POKEMON_MOVES
            .Moves(i).Num = Buffer.ReadLong
            .Moves(i).PP = Buffer.ReadLong
            .Moves(i).MaxPP = Buffer.ReadLong
        Next i
    End With
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandlePlayerPokemon", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBattle(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long
Dim IBattle As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    IBattle = Buffer.ReadByte
    If i = MyIndex Then
        If IBattle > 0 Then
            InitBattle
            CurPoke = Buffer.ReadByte
            Fade = True: FadeType = 1
            ShowMoves = False: ShowPokemonSwitch = False
            ForceSwitch = False
            PokeAlpha = 255
            If Not FileExist(App.Path & MUSIC_PATH & Battle_Wild_Music) Then
                StopMusic
            Else
                PlayMusic Battle_Wild_Music
            End If
            Player(MyIndex).InBattle = IBattle
        End If
    Else
        Player(i).InBattle = IBattle
    End If
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandleBattle", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEnemyPokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ClearEnemyPokemon
    With EnemyPokemon
        .Num = Buffer.ReadLong
        .Gender = Buffer.ReadByte
        .Level = Buffer.ReadLong
        .CurHP = Buffer.ReadLong
        For i = 1 To Stats.Stat_Count - 1
            .Stat(i) = Buffer.ReadLong
            .StatIV(i) = Buffer.ReadLong
            .StatEV(i) = Buffer.ReadLong
        Next i
        .Exp = Buffer.ReadLong
        For i = 1 To MAX_POKEMON_MOVES
            .Moves(i).Num = Buffer.ReadLong
            .Moves(i).PP = Buffer.ReadLong
            .Moves(i).MaxPP = Buffer.ReadLong
        Next i
    End With
    Set Buffer = Nothing
    EnemyPos = 600
    
    Exit Sub
errHandler:
    HandleError "HandleEnemyPokemon", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBattleMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    AddBattleLog Trim$(Buffer.ReadString), CheckColor(Buffer.ReadLong)
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandleBattleMsg", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayMusic(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PlayMusic Trim$(Buffer.ReadString)
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandlePlayMusic", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlaySound(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PlaySound Trim$(Buffer.ReadString)
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandlePlaySound", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEditMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Editor = 0 Then
        With frmMoveEditor
            Editor = EDITOR_MOVE
            
            .lstIndex.Clear
            For i = 1 To Count_Move
                .lstIndex.AddItem i & ": " & Trim$(Moves(i).Name)
            Next
            .Show
            .lstIndex.ListIndex = 0
            
            MoveEditorInit
        End With
    End If
    
    Exit Sub
errHandler:
    HandleError "HandleEditMove", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long
Dim MoveSize As Long
Dim MoveData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    MoveSize = LenB(Moves(n))
    ReDim MoveData(MoveSize - 1)
    MoveData = Buffer.ReadBytes(MoveSize)
    CopyMemory ByVal VarPtr(Moves(n)), ByVal VarPtr(MoveData(0)), MoveSize
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandleUpdateMove", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleExpCalc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer, i As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    For i = 1 To MAX_LEVEL
        ExpCalc(i) = Buffer.ReadLong
    Next
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandleExpCalc", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTarget(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MyTarget = Buffer.ReadLong
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandleTarget", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdatePokemonVital(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Slot As Byte, i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Slot = Buffer.ReadByte
    With Player(MyIndex).Pokemon(Slot)
        .CurHP = Buffer.ReadLong
        For i = 1 To MAX_POKEMON_MOVES
            .Moves(i).PP = Buffer.ReadLong
        Next
    End With
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandleUpdatePokemonVital", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateEnemyVital(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    With EnemyPokemon
        .CurHP = Buffer.ReadLong
        For i = 1 To MAX_POKEMON_MOVES
            .Moves(i).PP = Buffer.ReadLong
        Next
    End With
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandleUpdateEnemyVital", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBattleResult(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    CanUseCmd = True
    
    Exit Sub
errHandler:
    HandleError "HandleBattleResult", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleExitBattle(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    DidWin = Buffer.ReadByte
    Set Buffer = Nothing

    ExitBattleTmr = True
    CanUseCmd = False
    CanExit = False
    
    Exit Sub
errHandler:
    HandleError "HandleExitBattle", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleForceSwitch(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    ForceSwitch = True
    ShowPokemonSwitch = True
    
    Exit Sub
errHandler:
    HandleError "HandleForceSwitch", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSwitch(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Force As Byte
Dim X As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    TmpSwitch = Buffer.ReadByte
    Force = Buffer.ReadByte
    Set Buffer = Nothing
    isSwitch = True
    SwitchPos = YES
    If Force = YES Then
        IsSwitchForce = True
    Else
        IsSwitchForce = False
    End If
        
    Exit Sub
errHandler:
    HandleError "HandleSwitch", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleLearnMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    LearnMoveNum = Buffer.ReadLong
    LearnPokeNum = Buffer.ReadLong
    Set Buffer = Nothing
    IsLearnMove = True
    
    Exit Sub
errHandler:
    HandleError "HandleLearnMove", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEvolve(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim PokeSlot As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PokeSlot = Buffer.ReadLong
    Set Buffer = Nothing
    TmpInEvolve = PokeSlot
    
    Exit Sub
errHandler:
    HandleError "HandleEvolve", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEditItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Editor = 0 Then
        With frmItemEditor
            Editor = EDITOR_ITEM
            
            .lstIndex.Clear
            For i = 1 To Count_Item
                .lstIndex.AddItem i & ": " & Trim$(Item(i).Name)
            Next
            .Show
            .lstIndex.ListIndex = 0
            
            ItemEditorInit
        End With
    End If
    
    Exit Sub
errHandler:
    HandleError "HandleEditItem", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long
Dim ItemSize As Long
Dim ItemData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = Buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandleUpdateItem", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleInventory(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim X As Long, Y As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    For X = 1 To MAX_PLAYER_ITEM
        For Y = 0 To ItemType.Item_Count - 1
            Player(MyIndex).Item(X, Y).Num = Buffer.ReadLong
            Player(MyIndex).Item(X, Y).value = Buffer.ReadLong
        Next
    Next
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandleInventory", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSelect(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    InputData1 = Buffer.ReadLong
    InputData2 = Buffer.ReadLong
    InputData3 = Buffer.ReadLong
    InputData4 = Buffer.ReadLong
    Set Buffer = Nothing
    ShowSelect = True
    
    If InputData1 = SELECT_POKEMON Then
        MaxSelection = CountPlayerPokemon(MyIndex)
    ElseIf InputData1 = SELECT_MOVE Then
        MaxSelection = CountPlayerPokemonMove(MyIndex, InputData2)
    End If
    
    Exit Sub
errHandler:
    HandleError "HandleSelect", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBattleRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim X As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    X = Buffer.ReadLong
    Set Buffer = Nothing
    
    BattleRequestIndex = X
    
    AddText Trim$(Player(X).Name) & " challenge you into a battle!", White
    AddText "/accept - to accept the challenge", White
    AddText "/decline - to decline the challenge", White
    
    Exit Sub
errHandler:
    HandleError "HandleBattleRequest", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCaptured(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    IsCapture = True
    Capture = 255
    
    Exit Sub
errHandler:
    HandleError "HandleCaptured", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerStoredPokemon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Slot As Byte
Dim i As Long
Dim X As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Slot = Buffer.ReadLong
    With Player(MyIndex).StoredPokemon(Slot)
        .Num = Buffer.ReadLong
        .Gender = Buffer.ReadByte
        .CurHP = Buffer.ReadLong
        .Level = Buffer.ReadLong
        For i = 1 To Stats.Stat_Count - 1
            .Stat(i) = Buffer.ReadLong
            .StatIV(i) = Buffer.ReadLong
            .StatEV(i) = Buffer.ReadLong
        Next i
        .Exp = Buffer.ReadLong
        For i = 1 To MAX_POKEMON_MOVES
            .Moves(i).Num = Buffer.ReadLong
            .Moves(i).PP = Buffer.ReadLong
            .Moves(i).MaxPP = Buffer.ReadLong
        Next i
    End With
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandlePlayerStoredPokemon", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleStorage(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    OpenStorage

    Exit Sub
errHandler:
    HandleError "HandlePlayerStoredPokemon", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEditNPC(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Editor = 0 Then
        With frmNPCEditor
            Editor = EDITOR_NPC
            
            .lstIndex.Clear
            For i = 1 To Count_NPC
                .lstIndex.AddItem i & ": " & Trim$(NPC(i).Name)
            Next
            .Show
            .lstIndex.ListIndex = 0
            
            NPCEditorInit
        End With
    End If
    
    Exit Sub
errHandler:
    HandleError "HandleEditNPC", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateNPC(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long
Dim NPCSize As Long
Dim NPCData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    NPCSize = LenB(NPC(n))
    ReDim NPCData(NPCSize - 1)
    NPCData = Buffer.ReadBytes(NPCSize)
    CopyMemory ByVal VarPtr(NPC(n)), ByVal VarPtr(NPCData(0)), NPCSize
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandleUpdateNPC", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEditShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Editor = 0 Then
        With frmShopEditor
            Editor = EDITOR_SHOP
            
            .lstIndex.Clear
            For i = 1 To Count_Shop
                .lstIndex.AddItem i & ": " & Trim$(Shop(i).Name)
            Next
            .Show
            .lstIndex.ListIndex = 0
            
            ShopEditorInit
        End With
    End If
    
    Exit Sub
errHandler:
    HandleError "HandleEditShop", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long
Dim ShopSize As Long
Dim ShopData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    ShopSize = LenB(Shop(n))
    ReDim ShopData(ShopSize - 1)
    ShopData = Buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(n)), ByVal VarPtr(ShopData(0)), ShopSize
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandleUpdateShop", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleOpenShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    InShop = Buffer.ReadLong
    Set Buffer = Nothing
    
    For i = ButtonEnum.ShopScrollUp To ButtonEnum.ShopScrollDown
        Buttons(i).Visible = True
    Next
    OpenInventory
    ShopStart = 1
    
    Exit Sub
errHandler:
    HandleError "HandleOpenShop", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpawnNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    With MapNpc(n)
        .Num = Buffer.ReadLong
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
        .Dir = Buffer.ReadLong
        
        .Moving = 0
        .xOffset = 0
        .yOffset = 0
    End With
    Set Buffer = Nothing
End Sub

Private Sub HandleNpcClear(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim n As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    Call ClearMapNpc(n)
    Set Buffer = Nothing
End Sub

Private Sub HandleNpcMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    With MapNpc(Buffer.ReadLong)
        .X = Buffer.ReadLong
        .Y = Buffer.ReadLong
        .Dir = Buffer.ReadLong
        .xOffset = 0
        .yOffset = 0
        .Moving = Buffer.ReadLong
        Select Case .Dir
            Case DIR_UP
                .yOffset = Pic_Size
            Case DIR_DOWN
                .yOffset = Pic_Size * -1
            Case DIR_LEFT
                .xOffset = Pic_Size
            Case DIR_RIGHT
                .xOffset = Pic_Size * -1
        End Select
    End With
End Sub

Private Sub HandleMapNpcData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    For i = 1 To MAX_MAP_NPC
        With MapNpc(i)
            .Num = Buffer.ReadLong
            .X = Buffer.ReadLong
            .Y = Buffer.ReadLong
            .Dir = Buffer.ReadLong
        End With
    Next
    Set Buffer = Nothing
End Sub

Private Sub HandleNpcDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    With MapNpc(Buffer.ReadLong)
        .Dir = Buffer.ReadLong
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With
    Set Buffer = Nothing
End Sub

Private Sub HandleTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim Count As Long

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    InTradeIndex = Buffer.ReadLong
    Count = Buffer.ReadLong
    Set Buffer = Nothing
    
    AddText Trim$(Player(InTradeIndex).Name) & " would like to trade with you", White
    AddText "type: /accept or /decline", White
    AddText "Trade request expired in " & Count & "sec/s", White
End Sub

Private Sub HandleTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    InTrade = True
    InTradeConfirm = False
    OpenInventory
    ClearTradeSlots

    Buttons(ButtonEnum.TradeConfirm).Visible = True
    Buttons(ButtonEnum.TradeAccept).Visible = False
    Buttons(ButtonEnum.TradeDecline).Visible = False
    
    MyTradeConfirm = False
    TheirTradeConfirm = False
    
    ' Init Trade
End Sub

Private Sub HandleCloseTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    InTrade = False
    InTradeConfirm = False
    InTradeIndex = 0
    CloseInventory
    ClearTradeSlots
    
    Buttons(ButtonEnum.TradeConfirm).Visible = False
    Buttons(ButtonEnum.TradeAccept).Visible = False
    Buttons(ButtonEnum.TradeDecline).Visible = False
        
    MyTradeConfirm = False
    TheirTradeConfirm = False
End Sub

Private Sub HandleTradeConfirm(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim Buffer As clsBuffer
Dim tIndex As Long, i As Long
Dim PokemonSize As Long, PokemonData() As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    tIndex = Buffer.ReadLong
    If tIndex <> MyIndex Then
        For i = 1 To MAX_TRADE
            With TheirTrade(i)
                .Type = Buffer.ReadByte
                
                .ItemNum = Buffer.ReadLong
                .ItemVal = Buffer.ReadLong
                
                PokemonSize = LenB(.Pokemon)
                ReDim PokemonData(PokemonSize - 1)
                PokemonData = Buffer.ReadBytes(PokemonSize)
                CopyMemory ByVal VarPtr(.Pokemon), ByVal VarPtr(PokemonData(0)), PokemonSize
                
                .TempItemSlot = Buffer.ReadLong
                .TempItemType = Buffer.ReadByte
                .TempPokeSlot = Buffer.ReadLong
            End With
        Next
        TheirTradeConfirm = True
    Else
        MyTradeConfirm = True
    End If
    Set Buffer = Nothing
    
    If MyTradeConfirm Then
        InTradeConfirm = True
        InTrade = False
        CloseInventory True
        Buttons(ButtonEnum.TradeConfirm).Visible = False
        
        Buttons(ButtonEnum.TradeAccept).Visible = True
        Buttons(ButtonEnum.TradeDecline).Visible = True
    End If
End Sub
