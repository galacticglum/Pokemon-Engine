Attribute VB_Name = "modTCP"
Option Explicit

Public Function IsConnected(ByVal index As Long) As Boolean
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If index <= 0 Or index > MAX_PLAYER Then Exit Function
    IsConnected = False
    If frmMain.Socket(index).State = sckConnected Then
        IsConnected = True
    End If
    
    Exit Function
errHandler:
    HandleError "IsConnected", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function IsLoggedIn(ByVal index As Long) As Boolean
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If index <= 0 Or index > MAX_PLAYER Then Exit Function
    IsLoggedIn = False
    If Len(Trim$(Player(index).Username)) > 0 Then
        IsLoggedIn = True
    End If
    
    Exit Function
errHandler:
    HandleError "IsLoggedIn", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function IsPlaying(ByVal index As Long) As Boolean
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If index <= 0 Or index > MAX_PLAYER Then Exit Function
    IsPlaying = False
    If IsConnected(index) Then
        If TempPlayer(index).InGame Then
            IsPlaying = True
        End If
    End If

    Exit Function
errHandler:
    HandleError "IsPlaying", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub SendDataTo(ByVal index As Long, ByRef Data() As Byte)
Dim buffer As clsBuffer
Dim TempData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If IsConnected(index) Then
        Set buffer = New clsBuffer
        TempData = Data
        
        buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        buffer.WriteBytes TempData()
              
        frmMain.Socket(index).SendData buffer.ToArray()
    End If
    
    Exit Sub
errHandler:
    HandleError "SendDataTo", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDataToAll(ByRef Data() As Byte)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To HighPlayerIndex
        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If
    Next

    Exit Sub
errHandler:
    HandleError "SendDataToAll", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDataToAllBut(ByVal index As Long, ByRef Data() As Byte)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To HighPlayerIndex
        If IsPlaying(i) Then
            If i <> index Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next

    Exit Sub
errHandler:
    HandleError "SendDataToAllBut", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDataToMap(ByVal MapNum As Long, ByRef Data() As Byte)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To HighPlayerIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next

    Exit Sub
errHandler:
    HandleError "SendDataToMap", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendDataToMapBut(ByVal index As Long, ByVal MapNum As Long, ByRef Data() As Byte)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To HighPlayerIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                If i <> index Then
                    Call SendDataTo(i, Data)
                End If
            End If
        End If
    Next

    Exit Sub
errHandler:
    HandleError "SendDataToMap", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function FindOpenPlayerSlot() As Long
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    FindOpenPlayerSlot = 0
    For i = 1 To MAX_PLAYER
        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If
    Next

    Exit Function
errHandler:
    HandleError "FindOpenPlayerSlot", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub AcceptConnection(ByVal index As Long, ByVal SocketId As Long)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If (index = 0) Then
        i = FindOpenPlayerSlot

        If i <> 0 Then
            frmMain.Socket(i).Close
            frmMain.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If

    Exit Sub
errHandler:
    HandleError "AcceptConnection", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub CloseSocket(ByVal index As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If index > 0 Then
        Call LeftGame(index)
        AddLog "Connection from " & GetPlayerIP(index) & " has been terminated!"
        frmMain.Socket(index).Close
        Call ClearPlayer(index)
    End If

    Exit Sub
errHandler:
    HandleError "CloseSocket", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SocketConnected(ByVal index As Long)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If index <> 0 Then
        AddLog "Receiving connection from " & GetPlayerIP(index) & "..."
        HighPlayerIndex = 0
        For i = MAX_PLAYER To 1 Step -1
            If IsConnected(i) Then
                HighPlayerIndex = i
                Exit For
            End If
        Next
        SendHighIndex
    End If
    
    Exit Sub
errHandler:
    HandleError "SocketConnected", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendAlertMsg(ByVal index As Long, ByVal Msg As String)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong SAlertMsg
    buffer.WriteString Msg
    SendDataTo index, buffer.ToArray()
    DoEvents
    CloseSocket index
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendAlertMsg", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendCharSelect(ByVal index As Long)
Dim buffer As clsBuffer
Dim x As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong SCharSelect
    For x = 1 To MAX_PLAYER_DATA
        buffer.WriteString Player(index).PlayerData(x).Name
        buffer.WriteLong Player(index).PlayerData(x).Sprite
        ' Player Sprite, Level
    Next
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendAlertMsg", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendIndex(ByVal index As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SIndex
    buffer.WriteLong index
    buffer.WriteLong HighPlayerIndex
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing

    Exit Sub
errHandler:
    HandleError "SendAlertMsg", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendHighIndex()
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SHighIndex
    buffer.WriteLong HighPlayerIndex
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendHighIndex", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendInGame(ByVal index As Long)
Dim buffer As clsBuffer
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong SInGame
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendInGame", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function PlayerData(ByVal index As Long) As Byte()
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If index <= 0 Or index > MAX_PLAYER Then Exit Function
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerData
    buffer.WriteLong index
    buffer.WriteString GetPlayerName(index)
    buffer.WriteByte GetPlayerGender(index)
    
    buffer.WriteByte Player(index).PlayerData(TempPlayer(index).CurSlot).Access
    
    buffer.WriteLong GetPlayerSprite(index)
    
    buffer.WriteLong GetPlayerMap(index)
    buffer.WriteLong GetPlayerX(index)
    buffer.WriteLong GetPlayerY(index)
    buffer.WriteByte GetPlayerDir(index)
    
    With Player(index).PlayerData(TempPlayer(index).CurSlot).PvP
        buffer.WriteLong .Win
        buffer.WriteLong .Lose
        buffer.WriteLong .Disconnect
    End With
    
    buffer.WriteLong Player(index).PlayerData(TempPlayer(index).CurSlot).Money
    
    buffer.WriteByte Player(index).PlayerData(TempPlayer(index).CurSlot).IsVIP

    PlayerData = buffer.ToArray()
    Set buffer = Nothing

    Exit Function
errHandler:
    HandleError "PlayerData", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub SendPlayerData(ByVal index As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    SendDataToMap GetPlayerMap(index), PlayerData(index)
    
    Exit Sub
errHandler:
    HandleError "SendPlayerData", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerXY(ByVal index As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerXY
    buffer.WriteLong GetPlayerX(index)
    buffer.WriteLong GetPlayerY(index)
    buffer.WriteByte GetPlayerDir(index)
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendPlayerXY", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerXYToMap(ByVal index As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerXYToMap
    buffer.WriteLong index
    buffer.WriteLong GetPlayerX(index)
    buffer.WriteLong GetPlayerY(index)
    buffer.WriteByte GetPlayerDir(index)
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendPlayerXYToMap", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendJoinMap(ByVal index As Long)
Dim buffer As clsBuffer
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    For i = 1 To HighPlayerIndex
        If IsPlaying(i) Then
            If i <> index Then
                If GetPlayerMap(i) = GetPlayerMap(index) Then
                    SendDataTo index, PlayerData(i)
                    SendPlayerPokemons i
                End If
            End If
        End If
    Next
    SendDataToMap GetPlayerMap(index), PlayerData(index)
    SendPlayerPokemons index
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendJoinMap", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendLeaveMap(ByVal index As Long, ByVal OldMap As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SLeaveMap
    buffer.WriteLong index
    SendDataToMapBut index, OldMap, buffer.ToArray()
    Set buffer = Nothing

    Exit Sub
errHandler:
    HandleError "SendLeaveMap", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendCheckForMap(ByVal index As Long, ByVal MapNum As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong SCheckForMap
    buffer.WriteLong MapNum
    buffer.WriteLong Map(MapNum).Rev
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendCheckForMap", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub MapCache_Create(ByVal MapNum As Long)
Dim buffer As clsBuffer
Dim x As Long, y As Long, i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong MapNum
    buffer.WriteString Trim$(Map(MapNum).Name)
    buffer.WriteString Trim$(Map(MapNum).Music)
    buffer.WriteLong Map(MapNum).Rev
    buffer.WriteByte Map(MapNum).Moral
    buffer.WriteLong Map(MapNum).MaxX
    buffer.WriteLong Map(MapNum).MaxY
    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY
            With Map(MapNum).Tile(x, y)
                For i = 0 To Layers.LayerCount - 1
                    buffer.WriteLong .Layer(i).Tileset
                    buffer.WriteLong .Layer(i).x
                    buffer.WriteLong .Layer(i).y
                Next i
                buffer.WriteByte .Type
                buffer.WriteLong .Data1
                buffer.WriteLong .Data2
                buffer.WriteLong .Data3
                buffer.WriteString .Data4
            End With
        Next
    Next
    For i = 0 To 3
        buffer.WriteLong Map(MapNum).Link(i)
    Next i
    For i = 1 To MAX_MAP_POKEMON
        buffer.WriteLong Map(MapNum).Pokemon(i)
    Next i
    buffer.WriteLong Map(MapNum).MinLvl
    buffer.WriteLong Map(MapNum).MaxLvl
    For i = 1 To MAX_MAP_NPC
        buffer.WriteLong Map(MapNum).Npc(i)
    Next i
    
    buffer.WriteLong Map(MapNum).CurField
    buffer.WriteLong Map(MapNum).CurBack
    MapCache(MapNum).Data = buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "MapCache_Create", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub CreateFullMapCache()
Dim i As Long
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_Map
        Call MapCache_Create(i)
    Next

    Exit Sub
errHandler:
    HandleError "CreateFullMapCache", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMap(ByVal index As Long, ByVal MapNum As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong SMap
    buffer.WriteBytes MapCache(MapNum).Data()
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing

    Exit Sub
errHandler:
    HandleError "SendMap", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerMove(ByVal index As Long, ByVal Moving As Byte, Optional ByVal sendToSelf As Boolean = False)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerMove
    buffer.WriteLong index
    buffer.WriteLong GetPlayerX(index)
    buffer.WriteLong GetPlayerY(index)
    buffer.WriteByte GetPlayerDir(index)
    buffer.WriteByte Moving
    If Not sendToSelf Then
        SendDataToMapBut index, GetPlayerMap(index), buffer.ToArray()
    Else
        SendDataToMap GetPlayerMap(index), buffer.ToArray()
    End If
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendPlayerMove", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendLeftGame(ByVal index As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SLeft
    buffer.WriteLong index
    SendDataToAllBut index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendLeftGame", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SMsg
    buffer.WriteString Msg
    buffer.WriteLong Color
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendMsg", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMsgToAll(ByVal Msg As String, ByVal Color As Long)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    For i = 1 To HighPlayerIndex
        If IsPlaying(i) Then
            SendMsg i, Msg, Color
        End If
    Next

    Exit Sub
errHandler:
    HandleError "SendMsgToAll", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMsgToMap(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Long)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    For i = 1 To HighPlayerIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                SendMsg i, Msg, Color
            End If
        End If
    Next

    Exit Sub
errHandler:
    HandleError "SendMsgToMap", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPokemons(ByVal index As Long)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_Pokemon
        If LenB(Trim$(Pokemon(i).Name)) > 0 Then
            Call SendUpdatePokemonTo(index, i)
        End If
    Next
    
    Exit Sub
errHandler:
    HandleError "SendPokemons", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUpdatePokemonToAll(ByVal PokemonNum As Long)
Dim buffer As clsBuffer
Dim PokemonSize As Long
Dim PokemonData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    PokemonSize = LenB(Pokemon(PokemonNum))
    ReDim PokemonData(PokemonSize - 1)
    CopyMemory PokemonData(0), ByVal VarPtr(Pokemon(PokemonNum)), PokemonSize
    buffer.WriteLong SUpdatePokemon
    buffer.WriteLong PokemonNum
    buffer.WriteBytes PokemonData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendUpdatePokemonToAll", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUpdatePokemonTo(ByVal index As Long, ByVal PokemonNum As Long)
Dim buffer As clsBuffer
Dim PokemonSize As Long
Dim PokemonData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    PokemonSize = LenB(Pokemon(PokemonNum))
    ReDim PokemonData(PokemonSize - 1)
    CopyMemory PokemonData(0), ByVal VarPtr(Pokemon(PokemonNum)), PokemonSize
    buffer.WriteLong SUpdatePokemon
    buffer.WriteLong PokemonNum
    buffer.WriteBytes PokemonData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendUpdatePokemonTo", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerPokemon(ByVal index As Long, ByVal PokemonSlot As Byte)
Dim buffer As clsBuffer
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerPokemon
    buffer.WriteLong index
    buffer.WriteByte PokemonSlot
    With Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(PokemonSlot)
        buffer.WriteLong .Num
        buffer.WriteByte .Gender
        buffer.WriteLong .CurHP
        buffer.WriteLong .Level
        For i = 1 To Stats.Stat_Count - 1
            buffer.WriteLong .Stat(i)
            buffer.WriteLong .StatIV(i)
            buffer.WriteLong .StatEV(i)
        Next i
        buffer.WriteLong .Exp
        For i = 1 To MAX_POKEMON_MOVES
            buffer.WriteLong .Moves(i).Num
            buffer.WriteLong .Moves(i).PP
            buffer.WriteLong .Moves(i).MaxPP
        Next i
    End With
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendPlayerPokemon", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerPokemons(ByVal index As Long)
Dim i As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To MAX_POKEMON
        SendPlayerPokemon index, i
    Next

    Exit Sub
errHandler:
    HandleError "SendPlayerPokemons", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBattle(ByVal index As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong SBattle
    buffer.WriteLong index
    buffer.WriteByte TempPlayer(index).InBattle
    buffer.WriteByte TempPlayer(index).InBattlePoke
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendBattle", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendEnemyPokemon(ByVal index As Long)
Dim buffer As clsBuffer
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEnemyPokemon
    With TempPlayer(index).EnemyPokemon
        buffer.WriteLong .Num
        buffer.WriteByte .Gender
        buffer.WriteLong .Level
        buffer.WriteLong .CurHP
        For i = 1 To Stats.Stat_Count - 1
            buffer.WriteLong .Stat(i)
            buffer.WriteLong .StatIV(i)
            buffer.WriteLong .StatEV(i)
        Next i
        buffer.WriteLong .Exp
        For i = 1 To MAX_POKEMON_MOVES
            buffer.WriteLong .Moves(i).Num
            buffer.WriteLong .Moves(i).PP
            buffer.WriteLong .Moves(i).MaxPP
        Next i
    End With
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendEnemyPokemon", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBattleMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong SBattleMsg
    buffer.WriteString Msg
    buffer.WriteLong Color
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendBattleMsg", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayMusic(ByVal index As Long, ByVal Music As String)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SPlayMusic
    buffer.WriteString Music
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing

    Exit Sub
errHandler:
    HandleError "SendPlayMusic", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayMusicToMap(ByVal MapNum As Long, ByVal Music As String)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To HighPlayerIndex
        If GetPlayerMap(i) = MapNum Then
            SendPlayMusic i, Music
        End If
    Next

    Exit Sub
errHandler:
    HandleError "SendPlayMusicToMap", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayMusicToAll(ByVal Music As String)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To HighPlayerIndex
        If IsPlaying(i) Then
            SendPlayMusic i, Music
        End If
    Next

    Exit Sub
errHandler:
    HandleError "SendPlayMusicToAll", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlaySound(ByVal index As Long, ByVal Sound As String)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SPlaySound
    buffer.WriteString Sound
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing

    Exit Sub
errHandler:
    HandleError "SendPlaySound", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlaySoundToMap(ByVal MapNum As Long, ByVal Sound As String)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To HighPlayerIndex
        If GetPlayerMap(i) = MapNum Then
            SendPlaySound i, Sound
        End If
    Next

    Exit Sub
errHandler:
    HandleError "SendPlaySoundToMap", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlaySoundToAll(ByVal Sound As String)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To HighPlayerIndex
        If IsPlaying(i) Then
            SendPlaySound i, Sound
        End If
    Next

    Exit Sub
errHandler:
    HandleError "SendPlaySoundToAll", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMoves(ByVal index As Long)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_Move
        If LenB(Trim$(Moves(i).Name)) > 0 Then
            Call SendUpdateMoveTo(index, i)
        End If
    Next
    
    Exit Sub
errHandler:
    HandleError "SendMoves", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUpdateMoveToAll(ByVal MoveNum As Long)
Dim buffer As clsBuffer
Dim MoveSize As Long
Dim MoveData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    MoveSize = LenB(Moves(MoveNum))
    ReDim MoveData(MoveSize - 1)
    CopyMemory MoveData(0), ByVal VarPtr(Moves(MoveNum)), MoveSize
    buffer.WriteLong SUpdateMove
    buffer.WriteLong MoveNum
    buffer.WriteBytes MoveData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendUpdateMoveToAll", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUpdateMoveTo(ByVal index As Long, ByVal MoveNum As Long)
Dim buffer As clsBuffer
Dim MoveSize As Long
Dim MoveData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    MoveSize = LenB(Moves(MoveNum))
    ReDim MoveData(MoveSize - 1)
    CopyMemory MoveData(0), ByVal VarPtr(Moves(MoveNum)), MoveSize
    buffer.WriteLong SUpdateMove
    buffer.WriteLong MoveNum
    buffer.WriteBytes MoveData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendUpdateMoveTo", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendExpCalc(ByVal index As Long)
Dim buffer As clsBuffer, i As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SExpCalc
    For i = 1 To MAX_LEVEL
        buffer.WriteLong ExpCalc(i)
    Next
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendExpCalc", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendTarget(ByVal index As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteLong STarget
    buffer.WriteLong TempPlayer(index).target
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendTarget", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUpdatePokemonVital(ByVal index As Long, ByVal TargetIndex As Long, ByVal PokeSlot As Byte)
Dim buffer As clsBuffer
Dim x As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SUpdatePokemonVital
    buffer.WriteByte PokeSlot
    With Player(TargetIndex).PlayerData(TempPlayer(TargetIndex).CurSlot).Pokemon(PokeSlot)
        buffer.WriteLong .CurHP
        For x = 1 To MAX_POKEMON_MOVES
            buffer.WriteLong .Moves(x).PP
        Next
    End With
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendUpdatePokemonVital", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUpdateEnemyVital(ByVal index As Long)
Dim buffer As clsBuffer
Dim x As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SUpdateEnemyVital
    With TempPlayer(index).EnemyPokemon
        buffer.WriteLong .CurHP
        For x = 1 To MAX_POKEMON_MOVES
            buffer.WriteLong .Moves(x).PP
        Next
    End With
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendUpdateEnemyVital", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBattleResult(ByVal index As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong SBattleResult
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendBattleResult", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendExitBattle(ByVal index As Long, Optional ByVal Didwin As Byte = 0)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong SExitBattle
    buffer.WriteByte Didwin
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendExitBattle", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendForceSwitch(ByVal index As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong SForceSwitch
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendForceSwitch", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSwitch(ByVal index As Long, ByVal PokeSlot As Byte, Optional ByVal Force As Byte = NO)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSwitch
    buffer.WriteByte PokeSlot
    buffer.WriteByte Force
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendSwitch", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendLearnMove(ByVal index As Long, ByVal MoveNum As Long, ByVal PokeSlot As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong SLearnMove
    buffer.WriteLong MoveNum
    buffer.WriteLong PokeSlot
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendLearnMove", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendEvolve(ByVal index As Long, ByVal PokeSlot As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEvolve
    buffer.WriteLong PokeSlot
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendEvolve", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendItems(ByVal index As Long)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_Item
        If LenB(Trim$(Item(i).Name)) > 0 Then
            Call SendUpdateItemTo(index, i)
        End If
    Next
    
    Exit Sub
errHandler:
    HandleError "SendItems", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUpdateItemToAll(ByVal ItemNum As Long)
Dim buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    buffer.WriteLong SUpdateItem
    buffer.WriteLong ItemNum
    buffer.WriteBytes ItemData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendUpdateItemToAll", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUpdateItemTo(ByVal index As Long, ByVal ItemNum As Long)
Dim buffer As clsBuffer
Dim ItemSize As Long
Dim ItemData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    buffer.WriteLong SUpdateItem
    buffer.WriteLong ItemNum
    buffer.WriteBytes ItemData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendUpdateItemTo", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendInventory(ByVal index As Long)
Dim buffer As clsBuffer
Dim x As Long, y As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SInventory
    With Player(index).PlayerData(TempPlayer(index).CurSlot)
        For x = 1 To MAX_PLAYER_ITEM
            For y = 0 To ItemType.Item_Count - 1
                buffer.WriteLong .Item(x, y).Num
                buffer.WriteLong .Item(x, y).Value
            Next
        Next
    End With
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendInventory", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSelect(ByVal index As Long, Optional ByVal Data1 As Long = 0, Optional ByVal Data2 As Long = 0, Optional ByVal Data3 As Long = 0, Optional ByVal Data4 As Long = 0)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SSelect
    buffer.WriteLong Data1
    buffer.WriteLong Data2
    buffer.WriteLong Data3
    buffer.WriteLong Data4
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendSelect", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendBattleRequest(ByVal index As Long, ByVal RequestIndex As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SBattleRequest
    buffer.WriteLong RequestIndex
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendBattleRequest", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendCaptured(ByVal index As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SCaptured
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendCaptured", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerStoredPokemon(ByVal index As Long, ByVal Slot As Long)
Dim buffer As clsBuffer
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerStoredPokemon
    buffer.WriteLong Slot
    With Player(index).PlayerData(TempPlayer(index).CurSlot).StoredPokemon(Slot)
        buffer.WriteLong .Num
        buffer.WriteByte .Gender
        buffer.WriteLong .CurHP
        buffer.WriteLong .Level
        For i = 1 To Stats.Stat_Count - 1
            buffer.WriteLong .Stat(i)
            buffer.WriteLong .StatIV(i)
            buffer.WriteLong .StatEV(i)
        Next i
        buffer.WriteLong .Exp
        For i = 1 To MAX_POKEMON_MOVES
            buffer.WriteLong .Moves(i).Num
            buffer.WriteLong .Moves(i).PP
            buffer.WriteLong .Moves(i).MaxPP
        Next i
    End With
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendPlayerStoredPokemon", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendPlayerStoredPokemons(ByVal index As Long)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To MAX_STORAGE_POKEMON
        SendPlayerStoredPokemon index, i
    Next

    Exit Sub
errHandler:
    HandleError "SendPlayerStoredPokemons", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendStorage(ByVal index As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong SStorage
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendStorage", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendNPCs(ByVal index As Long)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_NPC
        If LenB(Trim$(Npc(i).Name)) > 0 Then
            Call SendUpdateNPCTo(index, i)
        End If
    Next
    
    Exit Sub
errHandler:
    HandleError "SendNPCs", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUpdateNPCToAll(ByVal NpcNum As Long)
Dim buffer As clsBuffer
Dim NPCSize As Long
Dim NPCData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    NPCSize = LenB(Npc(NpcNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(Npc(NpcNum)), NPCSize
    buffer.WriteLong SUpdateNPC
    buffer.WriteLong NpcNum
    buffer.WriteBytes NPCData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendUpdateNPCToAll", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUpdateNPCTo(ByVal index As Long, ByVal NpcNum As Long)
Dim buffer As clsBuffer
Dim NPCSize As Long
Dim NPCData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    NPCSize = LenB(Npc(NpcNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(Npc(NpcNum)), NPCSize
    buffer.WriteLong SUpdateNPC
    buffer.WriteLong NpcNum
    buffer.WriteBytes NPCData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendUpdateNPCTo", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendShops(ByVal index As Long)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_Shop
        If LenB(Trim$(Shop(i).Name)) > 0 Then
            Call SendUpdateShopTo(index, i)
        End If
    Next
    
    Exit Sub
errHandler:
    HandleError "SendShops", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUpdateShopToAll(ByVal ShopNum As Long)
Dim buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(ShopNum)), ShopSize
    buffer.WriteLong SUpdateShop
    buffer.WriteLong ShopNum
    buffer.WriteBytes ShopData
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendUpdateShopToAll", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendUpdateShopTo(ByVal index As Long, ByVal ShopNum As Long)
Dim buffer As clsBuffer
Dim ShopSize As Long
Dim ShopData() As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(ShopNum)), ShopSize
    buffer.WriteLong SUpdateShop
    buffer.WriteLong ShopNum
    buffer.WriteBytes ShopData
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendUpdateShopTo", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendOpenShop(ByVal index As Long, ByVal ShopNum As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SOpenShop
    buffer.WriteLong ShopNum
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendOpenShop", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendSpawnNpc(ByVal MapNum As Long, ByVal MapNpcNum As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SSpawnNpc
    buffer.WriteLong MapNpcNum
    buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Num
    buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).x
    buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).y
    buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Dir
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendSpawnNpc", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendClearNPC(ByVal MapNum As Long, ByVal MapNpcNum As Long)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SNPCClear
    buffer.WriteLong MapNpcNum
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendClearNPC", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Movement As Byte)
Dim buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SNpcMove
    buffer.WriteLong MapNpcNum
    buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).x
    buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).y
    buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Dir
    buffer.WriteLong Movement
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendNpcMove", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMapNpcsToMap(ByVal MapNum As Long)
Dim buffer As clsBuffer
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SMapNpcData
    For i = 1 To MAX_MAP_NPC
        buffer.WriteLong MapNpc(MapNum).Npc(i).Num
        buffer.WriteLong MapNpc(MapNum).Npc(i).x
        buffer.WriteLong MapNpc(MapNum).Npc(i).y
        buffer.WriteLong MapNpc(MapNum).Npc(i).Dir
    Next
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendMapNpcsToMap", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendMapNpcsTo(ByVal index As Long, ByVal MapNum As Long)
Dim buffer As clsBuffer
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Set buffer = New clsBuffer
    buffer.WriteLong SMapNpcData
    For i = 1 To MAX_MAP_NPC
        buffer.WriteLong MapNpc(MapNum).Npc(i).Num
        buffer.WriteLong MapNpc(MapNum).Npc(i).x
        buffer.WriteLong MapNpc(MapNum).Npc(i).y
        buffer.WriteLong MapNpc(MapNum).Npc(i).Dir
    Next
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "SendMapNpcsTo", "modTCP", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendTradeRequest(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong STradeRequest
    buffer.WriteLong TempPlayer(index).InTradeRequest
    buffer.WriteLong TempPlayer(index).InTradeReqCount
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendTrade(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong STrade
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendCloseTrade(ByVal index As Long)
Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SCloseTrade
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Public Sub SendTradeConfirm(ByVal index As Long, ByVal SendTo As Long)
Dim buffer As clsBuffer
Dim i As Long
Dim PokemonSize As Long, PokemonData() As Byte
    
    Set buffer = New clsBuffer
    buffer.WriteLong STradeConfirm
    buffer.WriteLong index
    For i = 1 To MAX_TRADE
        With TempPlayer(index).MyTrade(i)
            buffer.WriteByte .Type

            buffer.WriteLong .ItemNum
            buffer.WriteLong .ItemVal

            PokemonSize = LenB(.Pokemon)
            ReDim PokemonData(PokemonSize - 1)
            CopyMemory PokemonData(0), ByVal VarPtr(.Pokemon), PokemonSize
            buffer.WriteBytes PokemonData
            
            buffer.WriteLong .TempItemSlot
            buffer.WriteByte .TempItemType
            buffer.WriteLong .TempPokeSlot
        End With
    Next
    SendDataTo SendTo, buffer.ToArray()
    Set buffer = Nothing
End Sub
