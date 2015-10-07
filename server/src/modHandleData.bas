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
    
    HandleDataSub(CRegister) = GetAddress(AddressOf HandleRegister)
    HandleDataSub(CLogin) = GetAddress(AddressOf HandleLogin)
    HandleDataSub(CAddChar) = GetAddress(AddressOf HandleAddChar)
    HandleDataSub(CDelChar) = GetAddress(AddressOf HandleDelChar)
    HandleDataSub(CUserChar) = GetAddress(AddressOf HandleUseChar)
    HandleDataSub(CNeedMap) = GetAddress(AddressOf HandleNeedMap)
    HandleDataSub(CPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(CPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(CMsg) = GetAddress(AddressOf HandleMsg)
    HandleDataSub(CRequestNewMap) = GetAddress(AddressOf HandleRequestNewMap)
    HandleDataSub(CRefresh) = GetAddress(AddressOf HandleRefresh)
    HandleDataSub(CRequestEditMap) = GetAddress(AddressOf HandleRequestEditMap)
    HandleDataSub(CMap) = GetAddress(AddressOf HandleMap)
    HandleDataSub(CRequestEditPokemon) = GetAddress(AddressOf HandleRequestEditPokemon)
    HandleDataSub(CRequestPokemons) = GetAddress(AddressOf HandleRequestPokemons)
    HandleDataSub(CSavePokemon) = GetAddress(AddressOf HandleSavePokemon)
    HandleDataSub(CBattleCommand) = GetAddress(AddressOf HandleBattleCommand)
    HandleDataSub(CRequestEditMove) = GetAddress(AddressOf HandleRequestEditMove)
    HandleDataSub(CRequestMoves) = GetAddress(AddressOf HandleRequestMoves)
    HandleDataSub(CSaveMove) = GetAddress(AddressOf HandleSaveMove)
    HandleDataSub(CExpCalc) = GetAddress(AddressOf HandleExpCalc)
    HandleDataSub(CTarget) = GetAddress(AddressOf HandleTarget)
    HandleDataSub(CSwitchComplete) = GetAddress(AddressOf HandleSwitchComplete)
    HandleDataSub(CAdminWarp) = GetAddress(AddressOf HandleAdminWarp)
    HandleDataSub(CReplaceMove) = GetAddress(AddressOf HandleReplaceMove)
    HandleDataSub(CEvolve) = GetAddress(AddressOf HandleEvolve)
    HandleDataSub(CRequestEditItem) = GetAddress(AddressOf HandleRequestEditItem)
    HandleDataSub(CRequestItems) = GetAddress(AddressOf HandleRequestItems)
    HandleDataSub(CSaveItem) = GetAddress(AddressOf HandleSaveItem)
    HandleDataSub(CUseItem) = GetAddress(AddressOf HandleUseItem)
    HandleDataSub(CInitSelect) = GetAddress(AddressOf HandleInitSelect)
    HandleDataSub(CBattleRequest) = GetAddress(AddressOf HandleBattleRequest)
    HandleDataSub(CInitBattle) = GetAddress(AddressOf HandleInitBattle)
    HandleDataSub(CDepositPokemon) = GetAddress(AddressOf HandleDepositPokemon)
    HandleDataSub(CWithdrawPokemon) = GetAddress(AddressOf HandleWithdrawPokemon)
    HandleDataSub(CRequestEditNPC) = GetAddress(AddressOf HandleRequestEditNPC)
    HandleDataSub(CRequestNPCs) = GetAddress(AddressOf HandleRequestNPCs)
    HandleDataSub(CSaveNPC) = GetAddress(AddressOf HandleSaveNPC)
    HandleDataSub(CRequestEditShop) = GetAddress(AddressOf HandleRequestEditShop)
    HandleDataSub(CRequestShops) = GetAddress(AddressOf HandleRequestShops)
    HandleDataSub(CSaveShop) = GetAddress(AddressOf HandleSaveShop)
    HandleDataSub(CBuyItem) = GetAddress(AddressOf HandleBuyItem)
    HandleDataSub(CSellItem) = GetAddress(AddressOf HandleSellItem)
    HandleDataSub(CInitTrade) = GetAddress(AddressOf HandleInitTrade)
    HandleDataSub(CTradeAccept) = GetAddress(AddressOf HandleTradeAccept)
    HandleDataSub(CTradeDecline) = GetAddress(AddressOf HandleTradeDecline)
    HandleDataSub(CCloseTrade) = GetAddress(AddressOf HandleCloseTrade)
    HandleDataSub(CTradeConfirm) = GetAddress(AddressOf HandleTradeConfirm)
    
    Exit Sub
errHandler:
    HandleError "InitMessages", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleData(ByVal index As Long, ByRef Data() As Byte)
Dim buffer As clsBuffer
Dim MsgType As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
        
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MsgType = buffer.ReadLong

    If MsgType < 0 Then Exit Sub
    If MsgType >= ClientPacket_Count Then Exit Sub
    CallWindowProc HandleDataSub(MsgType), index, buffer.ReadBytes(buffer.length), 0, 0
    
    Exit Sub
errHandler:
    HandleError "HandleData", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)
Dim buffer() As Byte
Dim pLength As Long
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If GetPlayerAccess(index) <= 0 Then
        If TempPlayer(index).DataBytes > 1000 Then
            If GetTickCount < TempPlayer(index).DataTimer Then
                Exit Sub
            End If
        End If
        If TempPlayer(index).DataPackets > 25 Then
            If GetTickCount < TempPlayer(index).DataTimer Then
                Exit Sub
            End If
        End If
    End If

    TempPlayer(index).DataBytes = TempPlayer(index).DataBytes + DataLength
    If GetTickCount >= TempPlayer(index).DataTimer Then
        TempPlayer(index).DataTimer = GetTickCount + 1000
        TempPlayer(index).DataBytes = 0
        TempPlayer(index).DataPackets = 0
    End If
    
    frmMain.Socket(index).GetData buffer(), vbUnicode, DataLength
    TempPlayer(index).buffer.WriteBytes buffer()
    
    If TempPlayer(index).buffer.length >= 4 Then
        pLength = TempPlayer(index).buffer.ReadLong(False)
        If pLength < 0 Then Exit Sub
    End If
    Do While pLength > 0 And pLength <= TempPlayer(index).buffer.length - 4
        If pLength <= TempPlayer(index).buffer.length - 4 Then
            TempPlayer(index).DataPackets = TempPlayer(index).DataPackets + 1
            TempPlayer(index).buffer.ReadLong
            HandleData index, TempPlayer(index).buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If TempPlayer(index).buffer.length >= 4 Then
            pLength = TempPlayer(index).buffer.ReadLong(False)
            If pLength < 0 Then Exit Sub
        End If
    Loop

    TempPlayer(index).buffer.Trim
    
    Exit Sub
errHandler:
    HandleError "IncomingData", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleRegister(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim User As String, Pass As String

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If IsConnected(index) Then
        If Not IsLoggedIn(index) Then
            Set buffer = New clsBuffer
            buffer.WriteBytes Data()
            User = Trim$(buffer.ReadString)
            Pass = Trim$(buffer.ReadString)
            
            If App.Major <> buffer.ReadLong Or App.Minor <> buffer.ReadLong Or App.Revision <> buffer.ReadLong Then
                Call SendAlertMsg(index, "Version is outdate! Please update your client before playing.")
                Exit Sub
            End If
            
            Set buffer = Nothing
            
            If Not CheckNameInput(User) Then
                Call SendAlertMsg(index, "Your username/password must be between 3 and 20 characters long and only letters, numbers, spaces, and _ allowed in names")
                Exit Sub
            End If
            If Not CheckNameInput(Pass) Then
                Call SendAlertMsg(index, "Your username/password must be between 3 and 20 characters long and only letters, numbers, spaces, and _ allowed in names")
                Exit Sub
            End If
            
            If Not AccountExist(User) Then
                AddAccount index, User, Pass
                AddLog "'Account: " & User & "/" & GetPlayerIP(index) & "' has been created..."
                Call LoadPlayer(index, User)
                Call LoadPlayerPokemon(index, User)
                SendCharSelect index
            Else
                Call SendAlertMsg(index, "Sorry, that account name is already taken!")
            End If
        End If
    End If
    
    Exit Sub
errHandler:
    HandleError "HandleRegister", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleLogin(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim User As String, Pass As String

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If IsConnected(index) Then
        If Not IsLoggedIn(index) Then
            Set buffer = New clsBuffer
            buffer.WriteBytes Data()
            User = Trim$(buffer.ReadString)
            Pass = Trim$(buffer.ReadString)
            
            If App.Major <> buffer.ReadLong Or App.Minor <> buffer.ReadLong Or App.Revision <> buffer.ReadLong Then
                Call SendAlertMsg(index, "Version is outdate! Please update your client before playing.")
                Exit Sub
            End If
            
            Set buffer = Nothing
            
            If Not AccountExist(User) Then
                Call SendAlertMsg(index, "That account name does not exist.")
                Exit Sub
            End If
            
            If Not CheckNameInput(User) Then
                Call SendAlertMsg(index, "Your username/password must be between 3 and 20 characters long and only letters, numbers, spaces, and _ allowed in names")
                Exit Sub
            End If
            If Not CheckNameInput(Pass) Then
                Call SendAlertMsg(index, "Your username/password must be between 3 and 20 characters long and only letters, numbers, spaces, and _ allowed in names")
                Exit Sub
            End If
            
            If Not isPasswordOK(User, Pass) Then
                Call SendAlertMsg(index, "Incorrect Password!")
                Exit Sub
            End If
            
            Call LoadPlayer(index, User)
            Call LoadPlayerPokemon(index, User)
            SendCharSelect index
            AddLog "'Account: " & User & "/" & GetPlayerIP(index) & "' has logged in..."
        End If
    End If
    
    Exit Sub
errHandler:
    HandleError "HandleLogin", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAddChar(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Gender As Byte, Slot As Byte, Name As String
Dim Starter As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If IsConnected(index) Then
        If IsLoggedIn(index) Then
            Set buffer = New clsBuffer
            buffer.WriteBytes Data()
            Name = Trim$(buffer.ReadString)
            Gender = buffer.ReadByte
            Slot = buffer.ReadByte
            Starter = buffer.ReadLong
            Set buffer = Nothing
                
            If Not CheckNameInput(Name) Then
                Call SendAlertMsg(index, "Your name must be between 3 and 20 characters long and only letters, numbers, spaces, and _ allowed.")
                Exit Sub
            End If
                
            If Len(Trim$(Player(index).PlayerData(Slot).Name)) > 0 Then
                Call SendAlertMsg(index, "You cannot create a character on this slot!")
                Exit Sub
            End If
                
            If FindChar(Name) Then
                Call SendAlertMsg(index, "Sorry, but that name is in use!")
                Exit Sub
            End If
                
            AddChar index, Name, Gender, Slot, Starter
            AddLog "'Character: " & Name & "' has been created!"
            SendCharSelect index
        End If
    End If
    
    Exit Sub
errHandler:
    HandleError "HandleAddChar", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleDelChar(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Slot As Byte
Dim tName As String

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If IsConnected(index) Then
        If IsLoggedIn(index) Then
            Set buffer = New clsBuffer
            buffer.WriteBytes Data()
            Slot = buffer.ReadByte
            Set buffer = Nothing
            
            If Slot <= 0 Or Slot > MAX_PLAYER_DATA Then Exit Sub

            tName = Trim$(Player(index).PlayerData(Slot).Name)

            Call ZeroMemory(ByVal VarPtr(Player(index).PlayerData(Slot)), LenB(Player(index).PlayerData(Slot)))
            Player(index).PlayerData(Slot).Name = vbNullString
            SavePlayer index
            SavePlayerPokemon index
            
            DeleteName tName
    
            AddLog "'Character: " & tName & "' from 'Account: " & Trim$(Player(index).Username) & "' has been deleted!"
            SendCharSelect index
        End If
    End If
    
    Exit Sub
errHandler:
    HandleError "HandleDelChar", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUseChar(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Slot As Byte
Dim tName As String

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If IsConnected(index) Then
        If IsLoggedIn(index) Then
            Set buffer = New clsBuffer
            buffer.WriteBytes Data()
            Slot = buffer.ReadByte
            Set buffer = Nothing
            
            If Slot <= 0 Or Slot > MAX_PLAYER_DATA Then Exit Sub

            tName = Trim$(Player(index).PlayerData(Slot).Name)
            TempPlayer(index).CurSlot = Slot
            
            If Len(tName) > 0 Then
                JoinGame index
            End If
        End If
    End If
    
    Exit Sub
errHandler:
    HandleError "HandleUseChar", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNeedMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim NeedMap As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    NeedMap = buffer.ReadByte
    Set buffer = Nothing
    
    If NeedMap = YES Then
        SendMap index, GetPlayerMap(index)
    End If
    SendMapNpcsTo index, GetPlayerMap(index)
    SendJoinMap index
    
    TempPlayer(index).GettingMap = False
    Set buffer = New clsBuffer
    buffer.WriteLong SMapDone
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandleNeedMap", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim tmpX As Long, tmpY As Long
Dim Dir As Byte
Dim Moving As Byte
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If TempPlayer(index).GettingMap = YES Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    tmpX = buffer.ReadLong
    tmpY = buffer.ReadLong
    Dir = buffer.ReadByte
    Moving = buffer.ReadByte
    Set buffer = Nothing

    If Dir < DIR_UP Or Dir > DIR_DOWN Then Exit Sub
    If Moving < 1 Or Moving > 2 Then Exit Sub
    
    If TempPlayer(index).InBattle > 0 Then
        SendPlayerXY (index)
        Exit Sub
    End If

    If GetPlayerX(index) <> tmpX Then
        SendPlayerXY (index)
        Exit Sub
    End If
    If GetPlayerY(index) <> tmpY Then
        SendPlayerXY (index)
        Exit Sub
    End If
    
    If TempPlayer(index).InTrade > 0 Then
        SendPlayerXY (index)
        Exit Sub
    End If
    
    Call PlayerMove(index, Dir, Moving)
    
    Exit Sub
errHandler:
    HandleError "HandlePlayerMove", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerDir(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Dir As Byte
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If TempPlayer(index).GettingMap = YES Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Dir = buffer.ReadByte
    Set buffer = Nothing

    If Dir < DIR_UP Or Dir > DIR_DOWN Then Exit Sub

    Call SetPlayerDir(index, Dir)
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerDir
    buffer.WriteLong index
    buffer.WriteLong GetPlayerDir(index)
    SendDataToMapBut index, GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandlePlayerDir", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Msg As String, tText As String
Dim sMsgType As Byte
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    sMsgType = buffer.ReadByte
    Set buffer = Nothing
    
    Select Case sMsgType
        Case MsgType.MapMsg
            tText = "[Map] " & GetPlayerName(index) & ": " & Msg
            Call SendMsgToMap(GetPlayerMap(index), tText, Silver)
        Case MsgType.GlobalMsg
            tText = "[Global] " & GetPlayerName(index) & ": " & Msg
            Call SendMsgToAll(tText, White)
    End Select
    
    Exit Sub
errHandler:
    HandleError "HandleMapMsg", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleRequestNewMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Dir As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Dir = buffer.ReadByte
    Set buffer = Nothing

    If Dir < DIR_UP Or Dir > DIR_DOWN Then Exit Sub
    Call PlayerMove(index, Dir, MOVING_WALKING)
    
    Exit Sub
errHandler:
    HandleError "HandleRequestNewMap", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleRefresh(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
    SendMsg index, "Map refresh", Green
    
    Exit Sub
errHandler:
    HandleError "HandleRefresh", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleRequestEditMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If GetPlayerAccess(index) < ACCESS_MAPPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong SEditMap
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandleRequestEditMap", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long, x As Long, y As Long
Dim MapNum As Long
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MapNum = GetPlayerMap(index)
    ClearMap MapNum
    Map(MapNum).Name = Trim$(buffer.ReadString)
    Map(MapNum).Music = Trim$(buffer.ReadString)
    Map(MapNum).Rev = buffer.ReadLong
    Map(MapNum).Moral = buffer.ReadByte
    Map(MapNum).MaxX = buffer.ReadLong
    Map(MapNum).MaxY = buffer.ReadLong
    ReDim Map(MapNum).Tile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY
            With Map(MapNum).Tile(x, y)
                For i = 0 To Layers.LayerCount - 1
                    .Layer(i).Tileset = buffer.ReadLong
                    .Layer(i).x = buffer.ReadLong
                    .Layer(i).y = buffer.ReadLong
                Next i
                .Type = buffer.ReadByte
                .Data1 = buffer.ReadLong
                .Data2 = buffer.ReadLong
                .Data3 = buffer.ReadLong
                .Data4 = buffer.ReadString
            End With
        Next
    Next
    For i = 0 To 3
        Map(MapNum).Link(i) = buffer.ReadLong
    Next i
    For i = 1 To MAX_MAP_POKEMON
        Map(MapNum).Pokemon(i) = buffer.ReadLong
    Next i
    Map(MapNum).MinLvl = buffer.ReadLong
    Map(MapNum).MaxLvl = buffer.ReadLong
    For i = 1 To MAX_MAP_NPC
        Map(MapNum).Npc(i) = buffer.ReadLong
        ClearMapNpc MapNum, i
    Next i
    
    Map(MapNum).CurField = buffer.ReadLong
    Map(MapNum).CurBack = buffer.ReadLong
    Set buffer = Nothing

    ' spawn other things
    Call SendMapNpcsToMap(MapNum)
    Call SpawnMapNpcs(MapNum)
    
    SaveMap MapNum
    MapCache_Create MapNum
    
    For i = 1 To HighPlayerIndex
        If IsPlaying(i) Then
            PlayerWarp i, GetPlayerMap(i), GetPlayerX(i), GetPlayerY(i)
        End If
    Next i
    
    Exit Sub
errHandler:
    HandleError "HandleMap", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleRequestEditPokemon(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If GetPlayerAccess(index) < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong SEditPokemon
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandleRequestEditPokemon", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleRequestPokemons(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    SendPokemons index
    
    Exit Sub
errHandler:
    HandleError "HandleRequestPokemons", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSavePokemon(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim PokemonNum As Long
Dim PokemonSize As Long
Dim PokemonData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If GetPlayerAccess(index) < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PokemonNum = buffer.ReadLong
    If PokemonNum < 0 Or PokemonNum > Count_Pokemon Then Exit Sub
    PokemonSize = LenB(Pokemon(PokemonNum))
    ReDim PokemonData(PokemonSize - 1)
    PokemonData = buffer.ReadBytes(PokemonSize)
    CopyMemory ByVal VarPtr(Pokemon(PokemonNum)), ByVal VarPtr(PokemonData(0)), PokemonSize
    Set buffer = Nothing
    
    Call SendUpdatePokemonToAll(PokemonNum)
    Call SavePokemon(PokemonNum)
    
    Exit Sub
errHandler:
    HandleError "HandleSavePokemon", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBattleCommand(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Data1 As Long, Data2 As Long, Data3 As Long
Dim Cmd As Byte, FoeIndex As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Cmd = buffer.ReadByte
    Data1 = buffer.ReadLong
    Data2 = buffer.ReadLong
    Data3 = buffer.ReadLong
    Set buffer = Nothing
    
    Select Case Cmd
        Case BATTLE_COMMAND_FIGHT
            With Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(TempPlayer(index).InBattlePoke)
                If .Moves(Data1).PP > 0 Then
                    If TempPlayer(index).InBattle = BATTLE_WILD Then
                        InitBattleVsNPC index, Data1
                    ElseIf TempPlayer(index).InBattle = BATTLE_TRAINER Then
                        TempPlayer(index).MoveSet = Data1
                    End If
                End If
            End With
        Case BATTLE_COMMAND_SWITCH
            With Player(index).PlayerData(TempPlayer(index).CurSlot)
                If .Pokemon(Data1).Num > 0 Then
                    If TempPlayer(index).InBattle = BATTLE_WILD Then
                        TempPlayer(index).InBattlePoke = Data1
                        SendUpdatePokemonVital index, index, TempPlayer(index).InBattlePoke
                        SendSwitch index, Data1, Data2
                    ElseIf TempPlayer(index).InBattle = BATTLE_TRAINER Then
                        TempPlayer(index).InBattlePoke = Data1
                        FoeIndex = TempPlayer(index).BattleRequest
                        TempPlayer(FoeIndex).EnemyPokemon = .Pokemon(TempPlayer(index).InBattlePoke)
                        SendEnemyPokemon FoeIndex
                        SendSwitch index, Data1, Data2
                    End If
                End If
            End With
        Case BATTLE_COMMAND_RUN
            If TempPlayer(index).InBattle = BATTLE_WILD Then
                If Data1 = YES Then
                    SendBattleMsg index, "You have successfully escaped!", White
                    ExitBattle index
                Else
                    GetEscapeChance index
                End If
            ElseIf TempPlayer(index).InBattle = BATTLE_TRAINER Then
                ' forfiet
            End If
    End Select
    
    Exit Sub
errHandler:
    HandleError "HandleBattleCommand", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleRequestEditMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If GetPlayerAccess(index) < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong SEditMove
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandleRequestEditMove", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleRequestMoves(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    SendMoves index
    
    Exit Sub
errHandler:
    HandleError "HandleRequestMoves", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSaveMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim MoveNum As Long
Dim MoveSize As Long
Dim MoveData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If GetPlayerAccess(index) < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MoveNum = buffer.ReadLong
    If MoveNum < 0 Or MoveNum > Count_Move Then Exit Sub
    MoveSize = LenB(Moves(MoveNum))
    ReDim MoveData(MoveSize - 1)
    MoveData = buffer.ReadBytes(MoveSize)
    CopyMemory ByVal VarPtr(Moves(MoveNum)), ByVal VarPtr(MoveData(0)), MoveSize
    Set buffer = Nothing
    
    Call SendUpdateMoveToAll(MoveNum)
    Call SaveMove(MoveNum)
    
    Exit Sub
errHandler:
    HandleError "HandleSaveMove", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleExpCalc(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, i As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    For i = 1 To MAX_LEVEL
        ExpCalc(i) = buffer.ReadLong
    Next
    Set buffer = Nothing
    SaveExpCalc
    For i = 1 To HighPlayerIndex
        If IsPlaying(i) Then
            SendExpCalc i
        End If
    Next
    
    Exit Sub
errHandler:
    HandleError "HandleExpCalc", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTarget(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, target As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    TempPlayer(index).target = buffer.ReadLong
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandleTarget", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSwitchComplete(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If TempPlayer(index).InBattle = BATTLE_TRAINER Then
        TempPlayer(index).MoveSet = 5
    ElseIf TempPlayer(index).InBattle = BATTLE_WILD Then
        NpcVsPlayer index
        If Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(TempPlayer(index).InBattlePoke).CurHP <= 0 Then
            If CheckPokemon(index) > 0 Then
                SendForceSwitch index
            Else
                ExitBattle index, YES
            End If
        End If
        SendBattleMsg index, EndLine, Cyan
        SendBattleResult index
    End If
    
    Exit Sub
errHandler:
    HandleError "HandleSwitchComplete", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAdminWarp(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, x As Long, y As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    x = buffer.ReadLong
    y = buffer.ReadLong
    If GetPlayerAccess(index) > 0 Then
        SetPlayerX index, x
        SetPlayerY index, y
        SendPlayerXYToMap index
    End If
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandleAdminWarp", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleReplaceMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim PokeSlot As Long, MoveSlot As Long, MoveNum As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PokeSlot = buffer.ReadLong
    MoveSlot = buffer.ReadLong
    MoveNum = buffer.ReadLong
    Set buffer = Nothing
    
    If PokeSlot <= 0 Then Exit Sub
    With Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(PokeSlot)
        SendMsg index, Trim$(Pokemon(.Num).Name) & " forgot " & Trim$(Moves(.Moves(MoveSlot).Num).Name), Green
    End With
    LearnMove index, PokeSlot, MoveSlot, MoveNum
    
    Exit Sub
errHandler:
    HandleError "HandleReplaceMove", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEvolve(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim PokeSlot As Long
Dim i As Long, NewMove As Long, FreeMoveSlot As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PokeSlot = buffer.ReadLong
    Set buffer = Nothing

    If PokeSlot <= 0 Then Exit Sub
    With Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(PokeSlot)
        If .Num > 0 Then
            If Pokemon(.Num).EvolveNum > 0 Then
                .Num = Pokemon(.Num).EvolveNum
                For i = 1 To Stats.Stat_Count - 1
                    .Stat(i) = GetPlayerPokeSlotStat(index, PokeSlot, i)
                Next
                NewMove = CheckLearnMove(index, PokeSlot, .Level)
                If NewMove > 0 Then
                    FreeMoveSlot = CheckFreeMoveSlot(index, , PokeSlot)
                    If FreeMoveSlot > 0 Then
                        LearnMove index, PokeSlot, FreeMoveSlot, NewMove
                    Else
                        SendLearnMove index, NewMove, PokeSlot
                    End If
                End If
            End If
        End If
    End With
    SendPlayerPokemon index, PokeSlot
    
    Exit Sub
errHandler:
    HandleError "HandleEvolve", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleRequestEditItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If GetPlayerAccess(index) < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong SEditItem
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandleRequestEditItem", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleRequestItems(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    SendItems index
    
    Exit Sub
errHandler:
    HandleError "HandleRequestItems", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSaveItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim ItemNum As Long
Dim ItemSize As Long
Dim ItemData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If GetPlayerAccess(index) < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ItemNum = buffer.ReadLong
    If ItemNum < 0 Or ItemNum > Count_Item Then Exit Sub
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    ItemData = buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(ItemNum)), ByVal VarPtr(ItemData(0)), ItemSize
    Set buffer = Nothing
    
    Call SendUpdateItemToAll(ItemNum)
    Call SaveItem(ItemNum)
    
    Exit Sub
errHandler:
    HandleError "HandleSaveItem", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUseItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim ItemSlot As Long
Dim CurInvType As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ItemSlot = buffer.ReadLong
    CurInvType = buffer.ReadByte
    Set buffer = Nothing
    
    UseItem index, ItemSlot, CurInvType
    
    Exit Sub
errHandler:
    HandleError "HandleUseItem", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleInitSelect(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim SelectData1 As Long, SelectData2 As Long, SelectData3 As Long, SelectData4 As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    SelectData1 = buffer.ReadLong
    SelectData2 = buffer.ReadLong
    SelectData3 = buffer.ReadLong
    SelectData4 = buffer.ReadLong
    Set buffer = Nothing
    InitSelectUseItem index, SelectData1, SelectData2, SelectData3, SelectData4
    
    Exit Sub
errHandler:
    HandleError "HandleInitSelect", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBattleRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim x As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    x = buffer.ReadLong
    Set buffer = Nothing
    
    If Not IsPlaying(x) Then
        SendMsg index, "Player is not playing!", Red
        Exit Sub
    End If
    If GetPlayerMap(x) <> GetPlayerMap(index) Then
        SendMsg index, "You must be with the same map as the request target!", Red
        Exit Sub
    End If
    If TempPlayer(x).InBattle > 0 Then
        SendMsg index, "Player is currently In-battle!", Red
        Exit Sub
    End If
    If TempPlayer(x).BattleRequest > 0 And TempPlayer(x).BattleRequest <> index Then
        SendMsg index, "Player have already receive a battle request!", Red
        Exit Sub
    End If
    If TempPlayer(x).InTradeRequest > 0 Then
        SendMsg index, "Player have a trade request, please invite later!", Red
        Exit Sub
    End If
    If CheckPokemon(x) <= 0 Then
        SendMsg index, "Player has been wiped out!", Red
        Exit Sub
    End If
    If CheckPokemon(index) <= 0 Then
        SendMsg index, "You can't request this battle!", Red
        Exit Sub
    End If
    
    TempPlayer(index).target = 0
    SendTarget index
    SendMsg index, "Battle request sent!", Green
    TempPlayer(index).BattleRequest = x
    SendBattleRequest x, index
    
    Exit Sub
errHandler:
    HandleError "HandleBattleRequest", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleInitBattle(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim x As Long, Cmd As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    x = buffer.ReadLong
    Cmd = buffer.ReadByte
    Set buffer = Nothing
    
    If Cmd = NO Then
        SendMsg x, "Reject your challenge request!", Red
        Exit Sub
    End If
    If Not IsPlaying(x) Then
        SendMsg index, "Player is not playing!", Red
        Exit Sub
    End If
    If GetPlayerMap(x) <> GetPlayerMap(index) Then
        SendMsg index, "You must be with the same map as the request target!", Red
        Exit Sub
    End If
    If TempPlayer(x).InBattle > 0 Then
        SendMsg index, "Player is currently In-battle!", Red
        Exit Sub
    End If
    If TempPlayer(x).BattleRequest > 0 And TempPlayer(x).BattleRequest <> index Then
        SendMsg index, "Player have already receive a battle request!", Red
        Exit Sub
    End If
    If TempPlayer(x).InTradeRequest > 0 Then
        SendMsg index, "Player have a trade request, please invite later!", Red
        Exit Sub
    End If
    If CheckPokemon(x) <= 0 Then
        SendMsg index, "Player has been wiped out!", Red
        Exit Sub
    End If
    If CheckPokemon(index) <= 0 Then
        SendMsg index, "You can't accept this battle!", Red
        Exit Sub
    End If
    
    TempPlayer(index).target = 0: TempPlayer(x).target = 0
    SendTarget index: SendTarget x
    SendMsg x, "Challenge Accepted, Initiate Battle!", Green
    SendMsg index, "Challenge Accepted, Initiate Battle!", Green
    TempPlayer(index).BattleRequest = x
    TempPlayer(x).BattleRequest = index
    InitPlayerVsPlayer index, x
    
    Exit Sub
errHandler:
    HandleError "HandleInitBattle", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleDepositPokemon(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Slot As Byte
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Slot = buffer.ReadByte
    Set buffer = Nothing
    
    If Slot <= 0 Or Slot > MAX_POKEMON Then Exit Sub
    If CountPokemon(index) = 1 Then
        SendMsg index, "This is the last pokemon on your slot!", Red
        Exit Sub
    End If
    With Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(Slot)
        If .Num > 0 Then
            .CurHP = .Stat(Stats.HP)
            For i = 1 To MAX_POKEMON_MOVES
                .Moves(i).PP = .Moves(i).MaxPP
            Next
            DepositPokemon index, Slot
        End If
    End With
    
    Exit Sub
errHandler:
    HandleError "HandleDepositPokemon", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleWithdrawPokemon(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Slot As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Slot = buffer.ReadLong
    Set buffer = Nothing
    
    If Slot <= 0 Or Slot > MAX_STORAGE_POKEMON Then Exit Sub
    If CountPokemon(index) = MAX_POKEMON Then
        SendMsg index, "You already have " & MAX_POKEMON & " on your slot!", Red
        Exit Sub
    End If
    With Player(index).PlayerData(TempPlayer(index).CurSlot).StoredPokemon(Slot)
        If .Num > 0 Then
            WithdrawPokemon index, Slot
        End If
    End With
    
    Exit Sub
errHandler:
    HandleError "HandleWithdrawPokemon", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleRequestEditNPC(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If GetPlayerAccess(index) < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong SEditNPC
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandleRequestEditNPC", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleRequestNPCs(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    SendNPCs index
    
    Exit Sub
errHandler:
    HandleError "HandleRequestNPCs", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSaveNPC(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim NpcNum As Long
Dim NPCSize As Long
Dim NPCData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If GetPlayerAccess(index) < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    NpcNum = buffer.ReadLong
    If NpcNum < 0 Or NpcNum > Count_NPC Then Exit Sub
    NPCSize = LenB(Npc(NpcNum))
    ReDim NPCData(NPCSize - 1)
    NPCData = buffer.ReadBytes(NPCSize)
    CopyMemory ByVal VarPtr(Npc(NpcNum)), ByVal VarPtr(NPCData(0)), NPCSize
    Set buffer = Nothing
    
    Call SendUpdateNPCToAll(NpcNum)
    Call SaveNPC(NpcNum)
    
    Exit Sub
errHandler:
    HandleError "HandleSaveNPC", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleRequestEditShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If GetPlayerAccess(index) < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong SEditShop
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "HandleRequestEditShop", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleRequestShops(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    SendShops index
    
    Exit Sub
errHandler:
    HandleError "HandleRequestShops", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSaveShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim ShopNum As Long
Dim ShopSize As Long
Dim ShopData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If GetPlayerAccess(index) < ACCESS_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ShopNum = buffer.ReadLong
    If ShopNum < 0 Or ShopNum > Count_Shop Then Exit Sub
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(ShopNum)), ByVal VarPtr(ShopData(0)), ShopSize
    Set buffer = Nothing
    
    Call SendUpdateShopToAll(ShopNum)
    Call SaveShop(ShopNum)
    
    Exit Sub
errHandler:
    HandleError "HandleSaveShop", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBuyItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Slot As Long, xVal As Long
Dim ShopNum As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ShopNum = buffer.ReadLong
    Slot = buffer.ReadLong
    xVal = buffer.ReadLong
    Set buffer = Nothing
    
    If ShopNum > 0 Then
        With Shop(ShopNum).sItem(Slot)
            If .Num > 0 Then
                If CanAfford(index, .Price * xVal) Then
                    GiveItem index, .Num, xVal
                    TakeMoney index, .Price * xVal
                Else
                    SendMsg index, "You don't have enough money to afford this item", Red
                End If
            End If
        End With
    End If
    
    Exit Sub
errHandler:
    HandleError "HandleBuyItem", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSellItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim Slot As Long, xVal As Long, IType As Byte
Dim Money As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Slot = buffer.ReadLong
    IType = buffer.ReadByte
    xVal = buffer.ReadLong
    Set buffer = Nothing
    
    With Player(index).PlayerData(TempPlayer(index).CurSlot).Item(Slot, IType)
        If .Num > 0 Then
            If xVal > .Value Then xVal = .Value
            Money = Item(.Num).Sell * xVal
            TakeItem index, .Num, xVal
            GiveMoney index, Money
        End If
    End With
    
    Exit Sub
errHandler:
    HandleError "HandleSellItem", "modHandleData", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleInitTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    Set buffer = Nothing
    
    If TempPlayer(index).InTradeRequest > 0 Then
        SendMsg index, "You cannot request another trade request for " & TempPlayer(index).InTradeReqCount & "sec/s", Red
        Exit Sub
    End If
    
    If TempPlayer(index).BattleRequest > 0 Then
        SendMsg index, "You cannot request trade if you receive a battle invitation!", Red
        Exit Sub
    End If
    
    If i > 0 And i <= MAX_PLAYER Then
        If IsPlaying(i) Then
            If TempPlayer(i).InTradeRequest > 0 Then
                SendMsg index, "Selected player already have another Trade Request!", Red
                Exit Sub
            End If
            
            If GetPlayerMap(i) <> GetPlayerMap(index) Then
                SendMsg index, "You must be on the same map as the trade target!", Red
                Exit Sub
            End If
            
            If TempPlayer(i).BattleRequest > 0 Then
                SendMsg index, "Player have received a battle request, please request later", Red
                Exit Sub
            End If
            
            TempPlayer(i).InTradeRequest = index
            TempPlayer(i).InTradeReqCount = 30
            TempPlayer(index).InTradeRequest = i
            TempPlayer(index).InTradeReqCount = 30
            
            SendMsg index, "Trade request sent!", White
            SendMsg index, "Trade request expire in " & TempPlayer(index).InTradeReqCount & "sec/s", White
            SendTradeRequest i
        End If
    End If
End Sub

Private Sub HandleTradeAccept(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long

    i = TempPlayer(index).InTradeRequest
    
    If TempPlayer(index).InTradeRequest > 0 And TempPlayer(index).InTradeRequest <> i Then
        SendMsg index, "You cannot request trade if you receive another trade request", Red
        Exit Sub
    End If
    
    If TempPlayer(index).BattleRequest > 0 Then
        SendMsg index, "You cannot request trade if you receive a battle invitation!", Red
        Exit Sub
    End If
    
    If i > 0 And i <= MAX_PLAYER Then
        If IsPlaying(i) Then
            If TempPlayer(i).InTradeRequest > 0 Then
                If TempPlayer(i).InTradeRequest <> index Then
                    SendMsg index, "Selected player already have another Trade Request!", Red
                    Exit Sub
                End If
            End If
            
            If GetPlayerMap(i) <> GetPlayerMap(index) Then
                SendMsg index, "You must be on the same map as the trade target!", Red
                Exit Sub
            End If
            
            If TempPlayer(i).BattleRequest > 0 Then
                SendMsg index, "Player have received a battle request, please request later", Red
                Exit Sub
            End If
            
            TempPlayer(i).InTrade = index
            TempPlayer(index).InTrade = i
            
            TempPlayer(i).InTradeRequest = 0
            TempPlayer(i).InTradeReqCount = 0
            TempPlayer(index).InTradeRequest = 0
            TempPlayer(index).InTradeReqCount = 0
            TempPlayer(i).MyTradeConfirm = False
            TempPlayer(index).MyTradeConfirm = False
            SendTrade i
            SendTrade index
            Exit Sub
        End If
    End If
    
    SendMsg index, "Trade request has been cancelled!", Red
    TempPlayer(index).InTradeRequest = 0
    TempPlayer(index).InTradeReqCount = 0
End Sub

Private Sub HandleTradeDecline(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long

    i = TempPlayer(index).InTradeRequest
    
    If i > 0 And i <= MAX_PLAYER Then
        If IsPlaying(i) Then
            SendMsg i, "Trade request decline!", Red
            TempPlayer(i).InTradeRequest = 0
            TempPlayer(i).InTradeRequest = 0
        End If
    End If
    SendMsg index, "Trade request decline!", Red
    TempPlayer(index).InTradeRequest = 0
    TempPlayer(index).InTradeReqCount = 0
End Sub

Private Sub HandleCloseTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long

    i = TempPlayer(index).InTrade
    If i > 0 And i <= MAX_PLAYER Then
        If IsPlaying(i) Then
            SendMsg i, "Trade has been cancelled!", Red
            TempPlayer(i).InTrade = 0
            SendCloseTrade i
        End If
    End If
    TempPlayer(index).InTrade = 0
End Sub

Private Sub HandleTradeConfirm(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer
Dim i As Long, tIndex As Long
Dim PokemonSize As Long, PokemonData() As Byte

    tIndex = TempPlayer(index).InTrade
    If tIndex <= 0 Or tIndex > MAX_PLAYER Then Exit Sub
    If Not IsPlaying(tIndex) Then Exit Sub
    If GetPlayerMap(tIndex) <> GetPlayerMap(index) Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    With TempPlayer(index)
        .MyTradeConfirm = True
        For i = 1 To MAX_TRADE
            .MyTrade(i).Type = buffer.ReadByte
            
            .MyTrade(i).ItemNum = buffer.ReadLong
            .MyTrade(i).ItemVal = buffer.ReadLong
            
            PokemonSize = LenB(.MyTrade(i).Pokemon)
            ReDim PokemonData(PokemonSize - 1)
            PokemonData = buffer.ReadBytes(PokemonSize)
            CopyMemory ByVal VarPtr(.MyTrade(i).Pokemon), ByVal VarPtr(PokemonData(0)), PokemonSize
            
            .MyTrade(i).TempItemSlot = buffer.ReadLong
            .MyTrade(i).TempItemType = buffer.ReadByte
            .MyTrade(i).TempPokeSlot = buffer.ReadLong
        Next i
    End With
    Set buffer = Nothing
    
    For i = 1 To MAX_TRADE
        TempPlayer(tIndex).TheirTrade(i) = TempPlayer(index).MyTrade(i)
    Next i
    SendTradeConfirm index, index
    SendTradeConfirm index, tIndex
End Sub
