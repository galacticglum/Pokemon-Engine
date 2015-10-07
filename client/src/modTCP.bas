Attribute VB_Name = "modTCP"
Option Explicit

Public PlayerBuffer As clsBuffer

Public Sub TcpInit()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set PlayerBuffer = New clsBuffer
    frmMain.Socket.RemoteHost = Trim$(Options.SaveIp)
    frmMain.Socket.RemotePort = Options.SavePort
    
    Exit Sub
errHandler:
    HandleError "TcpInit", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub DestroyTCP()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    frmMain.Socket.close
    
    Exit Sub
errHandler:
    HandleError "DestroyTCP", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function IsPlaying(ByVal Index As Long) As Boolean
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If LenB(GetPlayerName(Index)) > 0 Then
        IsPlaying = True
    End If

    Exit Function
errHandler:
    HandleError "IsPlaying", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function ConnectToServer() As Boolean
Dim Wait As Long
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If
    
    Wait = GetTickCount
    frmMain.Socket.close
    frmMain.Socket.Connect
    
    Do While (Not IsConnected) And (GetTickCount <= Wait + 3000)
        DoEvents
    Loop
    
    ConnectToServer = IsConnected
    
    Exit Function
errHandler:
    HandleError "ConnectToServer", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function IsConnected() As Boolean
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If frmMain.Socket.State = sckConnected Then
        IsConnected = True
    End If
    
    Exit Function
errHandler:
    HandleError "IsConnected", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub SendData(ByRef Data() As Byte)
Dim Buffer As clsBuffer
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If IsConnected Then
        Set Buffer = New clsBuffer
                
        Buffer.WriteLong (UBound(Data) - LBound(Data)) + 1
        Buffer.WriteBytes Data()
        frmMain.Socket.SendData Buffer.ToArray()
    End If
    
    Exit Sub
errHandler:
    HandleError "SendData", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRegisterData(ByVal rUser As String, ByVal rPass As String)
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRegister
    Buffer.WriteString rUser
    Buffer.WriteString rPass
    Buffer.WriteLong App.Major
    Buffer.WriteLong App.Minor
    Buffer.WriteLong App.Revision
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendRegisterData", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendLogin(ByVal lUser As String, ByVal lPass As String)
Dim Buffer As clsBuffer
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CLogin
    Buffer.WriteString lUser
    Buffer.WriteString lPass
    Buffer.WriteLong App.Major
    Buffer.WriteLong App.Minor
    Buffer.WriteLong App.Revision
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    Exit Sub
errHandler:
    HandleError "SendLogin", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendAddChar(ByVal aName As String, ByVal Gender As Byte, ByVal Slot As Byte, ByVal StarterNum As Long)
Dim Buffer As clsBuffer
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CAddChar
    Buffer.WriteString aName
    Buffer.WriteByte Gender
    Buffer.WriteByte Slot
    Buffer.WriteLong StarterNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    Exit Sub
errHandler:
    HandleError "SendAddChar", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDelChar(ByVal Slot As Byte)
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDelChar
    Buffer.WriteByte Slot
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    Exit Sub
errHandler:
    HandleError "SendDelChar", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUseChar(ByVal Slot As Byte)
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUserChar
    Buffer.WriteByte Slot
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    Exit Sub
errHandler:
    HandleError "SendUseChar", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendNeedMap(ByVal NeedMap As Byte)
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CNeedMap
    Buffer.WriteByte NeedMap
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    Exit Sub
errHandler:
    HandleError "SendNeedMap", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerMove()
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerMove
    Buffer.WriteLong GetPlayerX(MyIndex)
    Buffer.WriteLong GetPlayerY(MyIndex)
    Buffer.WriteByte GetPlayerDir(MyIndex)
    Buffer.WriteByte Player(MyIndex).Moving
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendPlayerMove", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerDir()
Dim Buffer As clsBuffer
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CPlayerDir
    Buffer.WriteByte GetPlayerDir(MyIndex)
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendPlayerDir", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMsg(ByVal Msg As String, ByVal sMsgType As MsgType)
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CMsg
    Buffer.WriteString Msg
    Buffer.WriteByte sMsgType
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendMapMsg", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerRequestNewMap()
Dim Buffer As clsBuffer
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestNewMap
    Buffer.WriteByte GetPlayerDir(MyIndex)
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendPlayerRequestNewMap", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRefresh()
Dim Buffer As clsBuffer
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRefresh
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendRefresh", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditMap()
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditMap
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    Exit Sub
errHandler:
    HandleError "SendRequestEditMap", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMap()
Dim Buffer As clsBuffer
Dim X As Long, Y As Long, i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CMap
    Buffer.WriteString Trim$(Map.Name)
    Buffer.WriteString Trim$(Map.Music)
    Buffer.WriteLong Map.Rev
    Buffer.WriteByte Map.Moral
    Buffer.WriteLong Map.MaxX
    Buffer.WriteLong Map.MaxY
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            With Map.Tile(X, Y)
                For i = 0 To Layers.LayerCount - 1
                    Buffer.WriteLong .Layer(i).Tileset
                    Buffer.WriteLong .Layer(i).X
                    Buffer.WriteLong .Layer(i).Y
                Next i
                Buffer.WriteByte .Type
                Buffer.WriteLong .Data1
                Buffer.WriteLong .Data2
                Buffer.WriteLong .Data3
                Buffer.WriteString .Data4
            End With
        Next
    Next
    For i = 0 To 3
        Buffer.WriteLong Map.Link(i)
    Next i
    For i = 1 To MAX_MAP_POKEMON
        Buffer.WriteLong Map.Pokemon(i)
    Next i
    Buffer.WriteLong Map.MinLvl
    Buffer.WriteLong Map.MaxLvl
    For i = 1 To MAX_MAP_NPC
        Buffer.WriteLong Map.NPC(i)
    Next i
    
    Buffer.WriteLong Map.CurField
    Buffer.WriteLong Map.CurBack
    SendData Buffer.ToArray
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendMap", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditPokemon()
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditPokemon
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendRequestEditPokemon", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestPokemons()
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestPokemons
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendRequestPokemons", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSavePokemon(ByVal PokemonNum As Long)
Dim Buffer As clsBuffer
Dim PokemonSize As Long
Dim PokemonData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    PokemonSize = LenB(Pokemon(PokemonNum))
    ReDim PokemonData(PokemonSize - 1)
    CopyMemory PokemonData(0), ByVal VarPtr(Pokemon(PokemonNum)), PokemonSize
    Buffer.WriteLong CSavePokemon
    Buffer.WriteLong PokemonNum
    Buffer.WriteBytes PokemonData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendSavePokemon", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBattleCommand(ByVal Cmd As Byte, Optional ByVal Data1 As Long = 0, Optional ByVal Data2 As Long = 0, Optional ByVal Data3 As Long)
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CBattleCommand
    Buffer.WriteByte Cmd
    Buffer.WriteLong Data1
    Buffer.WriteLong Data2
    Buffer.WriteLong Data3
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendBattleCommand", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditMove()
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditMove
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendRequestEditMove", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestMoves()
Dim Buffer As clsBuffer
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestMoves
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendRequestMoves", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveMove(ByVal MoveNum As Long)
Dim Buffer As clsBuffer
Dim MoveSize As Long
Dim MoveData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    MoveSize = LenB(Moves(MoveNum))
    ReDim MoveData(MoveSize - 1)
    CopyMemory MoveData(0), ByVal VarPtr(Moves(MoveNum)), MoveSize
    Buffer.WriteLong CSaveMove
    Buffer.WriteLong MoveNum
    Buffer.WriteBytes MoveData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendSaveMove", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendExpCalc()
Dim Buffer As clsBuffer, i As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CExpCalc
    For i = 1 To MAX_LEVEL
        Buffer.WriteLong ExpCalc(i)
    Next
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendExpCalc", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub PlayerTarget(ByVal Target As Long)
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If MyTarget = Target Then
        MyTarget = 0
    Else
        MyTarget = Target
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong CTarget
    Buffer.WriteLong Target
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    Exit Sub
errHandler:
    HandleError "PlayerTarget", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSwitchComplete()
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSwitchComplete
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    Exit Sub
errHandler:
    HandleError "SendSwitchComplete", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub AdminWarp(ByVal X As Long, ByVal Y As Long)
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CAdminWarp
    Buffer.WriteLong X
    Buffer.WriteLong Y
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    Exit Sub
errHandler:
    HandleError "AdminWarp", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendReplaceMove(ByVal PokeSlot As Long, ByVal MoveSlot As Long, ByVal MoveNum As Long)
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CReplaceMove
    Buffer.WriteLong PokeSlot
    Buffer.WriteLong MoveSlot
    Buffer.WriteLong MoveNum
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    Exit Sub
errHandler:
    HandleError "SendReplaceMove", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendEvolve(ByVal PokeSlot As Long)
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CEvolve
    Buffer.WriteLong PokeSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    Exit Sub
errHandler:
    HandleError "SendEvolve", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditItem()
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditItem
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendRequestEditItem", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestItems()
Dim Buffer As clsBuffer
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestItems
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendRequestItems", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveItem(ByVal ItemNum As Long)
Dim Buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    Buffer.WriteLong CSaveItem
    Buffer.WriteLong ItemNum
    Buffer.WriteBytes ItemData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendSaveItem", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUseItem(ByVal ItemSlot As Long)
Dim Buffer As clsBuffer
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CUseItem
    Buffer.WriteLong ItemSlot
    Buffer.WriteByte CurInvType
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendUseItem", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendInitSelect()
Dim Buffer As clsBuffer
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CInitSelect
    Buffer.WriteLong OutputData1
    Buffer.WriteLong OutputData2
    Buffer.WriteLong OutputData3
    Buffer.WriteLong OutputData4
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendInitSelect", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBattleRequest(ByVal Index As Long)
Dim Buffer As clsBuffer
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CBattleRequest
    Buffer.WriteLong Index
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendBattleRequest", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendInitBattle(ByVal Cmd As Byte)
Dim Buffer As clsBuffer
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CInitBattle
    Buffer.WriteLong BattleRequestIndex
    Buffer.WriteByte Cmd
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendInitBattle", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDepositPokemon(ByVal PokeSlot As Byte)
Dim Buffer As clsBuffer
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CDepositPokemon
    Buffer.WriteByte PokeSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendDepositPokemon", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendWithdrawPokemon(ByVal StorageSlot As Long)
Dim Buffer As clsBuffer
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong CWithdrawPokemon
    Buffer.WriteLong StorageSlot
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendWithdrawPokemon", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditNPC()
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditNPC
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendRequestEditNPC", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestNPCs()
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestNPCs
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendRequestNPCs", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveNPC(ByVal NPCNum As Long)
Dim Buffer As clsBuffer
Dim NPCSize As Long
Dim NPCData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    NPCSize = LenB(NPC(NPCNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(NPC(NPCNum)), NPCSize
    Buffer.WriteLong CSaveNPC
    Buffer.WriteLong NPCNum
    Buffer.WriteBytes NPCData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendSaveNPC", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestEditShop()
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestEditShop
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendRequestEditShop", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendRequestShops()
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CRequestShops
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendRequestShops", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSaveShop(ByVal ShopNum As Long)
Dim Buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(ShopNum)), ShopSize
    Buffer.WriteLong CSaveShop
    Buffer.WriteLong ShopNum
    Buffer.WriteBytes ShopData
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendSaveShop", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBuyItem(ByVal ShopItem As Long, ByVal Val As Long)
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CBuyItem
    Buffer.WriteLong InShop
    Buffer.WriteLong ShopItem
    Buffer.WriteLong Val
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendBuyItem", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSellItem(ByVal ItemSlot As Long, ByVal Val As Long)
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong CSellItem
    Buffer.WriteLong ItemSlot
    Buffer.WriteByte CurInvType
    Buffer.WriteLong Val
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendSellItem", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendInitTrade(ByVal Index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CInitTrade
    Buffer.WriteLong Index
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendTradeAccept()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CTradeAccept
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendTradeDecline()
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong CTradeDecline
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Public Sub SendTradeConfirm()
Dim Buffer As clsBuffer
Dim i As Long
Dim PokemonSize As Long, PokemonData() As Byte

    Set Buffer = New clsBuffer
    Buffer.WriteLong CTradeConfirm
    For i = 1 To MAX_TRADE
        With MyTrade(i)
            Buffer.WriteByte .Type
            
            Buffer.WriteLong .ItemNum
            Buffer.WriteLong .ItemVal
            
            PokemonSize = LenB(.Pokemon)
            ReDim PokemonData(PokemonSize - 1)
            CopyMemory PokemonData(0), ByVal VarPtr(.Pokemon), PokemonSize
            Buffer.WriteBytes PokemonData
            
            Buffer.WriteLong .TempItemSlot
            Buffer.WriteByte .TempItemType
            Buffer.WriteLong .TempPokeSlot
        End With
    Next
    SendData Buffer.ToArray()
    Set Buffer = Nothing
End Sub
