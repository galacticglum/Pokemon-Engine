Attribute VB_Name = "modPlayer"
Option Explicit

Public Sub AddAccount(ByVal index As Long, ByVal User As String, ByVal Pass As String)
Dim FileName As String

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If index <= 0 Or index > MAX_PLAYER Then Exit Sub
    
    ClearPlayer index
    Player(index).Username = User
    Player(index).Password = Pass
    Call SavePlayer(index)
    Call SavePlayerPokemon(index)
    
    Exit Sub
errHandler:
    HandleError "AddAccount", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerIP(ByVal index As Long) As String
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If index <= 0 Or index > MAX_PLAYER Then Exit Function
    GetPlayerIP = frmMain.Socket(index).RemoteHostIP
    
    Exit Function
errHandler:
    HandleError "AddAccount", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub AddChar(ByVal index As Long, ByVal Name As String, ByVal Gender As Byte, Slot As Byte, ByVal Starter As Long)
Dim FileName As String, F As Long
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    With Player(index).PlayerData(Slot)
        If Len(Trim$(.Name)) = 0 Then
            .Name = Name
            .Gender = Gender
            
            .Access = 0
            
            Select Case Gender
                Case GENDER_MALE
                    .Sprite = 1 ' Temp
                Case GENDER_FEMALE
                    .Sprite = 2 ' Temp
            End Select
            
            .Map = START_MAP
            .x = START_X
            .y = START_Y
            .Dir = DIR_DOWN
            
            .Checkpoint.Map = .Map
            .Checkpoint.x = .x
            .Checkpoint.y = .y
            
            .Money = 3000
            
            .TempData = Starter
      
            FileName = App.Path & "\bin\players\namelist.txt"
            F = FreeFile
            Open FileName For Append As #F
                Print #F, Name
            Close #F
            SavePlayer index
            SavePlayerPokemon index
        End If
    End With
    
    Exit Sub
errHandler:
    HandleError "AddChar", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerSprite(ByVal index As Long) As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If index <= 0 Or index > MAX_PLAYER Then Exit Function
    GetPlayerSprite = Player(index).PlayerData(TempPlayer(index).CurSlot).Sprite

    Exit Function
errHandler:
    HandleError "GetPlayerSprite", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerSprite(ByVal index As Long, ByVal SpriteNum As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If index <= 0 Or index > MAX_PLAYER Then Exit Sub
    Player(index).PlayerData(TempPlayer(index).CurSlot).Sprite = SpriteNum

    Exit Sub
errHandler:
    HandleError "SetPlayerSprite", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerName(ByVal index As Long) As String
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If index <= 0 Or index > MAX_PLAYER Then Exit Function
    GetPlayerName = Trim$(Player(index).PlayerData(TempPlayer(index).CurSlot).Name)

    Exit Function
errHandler:
    HandleError "GetPlayerName", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function GetPlayerGender(ByVal index As Long) As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If index <= 0 Or index > MAX_PLAYER Then Exit Function
    GetPlayerGender = Player(index).PlayerData(TempPlayer(index).CurSlot).Gender

    Exit Function
errHandler:
    HandleError "GetPlayerGender", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function GetPlayerMap(ByVal index As Long) As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If index <= 0 Or index > MAX_PLAYER Then Exit Function
    GetPlayerMap = Player(index).PlayerData(TempPlayer(index).CurSlot).Map

    Exit Function
errHandler:
    HandleError "GetPlayerMap", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerMap(ByVal index As Long, ByVal MapNum As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If index <= 0 Or index > MAX_PLAYER Then Exit Sub
    Player(index).PlayerData(TempPlayer(index).CurSlot).Map = MapNum

    Exit Sub
errHandler:
    HandleError "SetPlayerMap", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerX(ByVal index As Long) As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If index <= 0 Or index > MAX_PLAYER Then Exit Function
    GetPlayerX = Player(index).PlayerData(TempPlayer(index).CurSlot).x

    Exit Function
errHandler:
    HandleError "GetPlayerX", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerX(ByVal index As Long, ByVal xVal As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If index <= 0 Or index > MAX_PLAYER Then Exit Sub
    Player(index).PlayerData(TempPlayer(index).CurSlot).x = xVal

    Exit Sub
errHandler:
    HandleError "SetPlayerX", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerY(ByVal index As Long) As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If index <= 0 Or index > MAX_PLAYER Then Exit Function
    GetPlayerY = Player(index).PlayerData(TempPlayer(index).CurSlot).y

    Exit Function
errHandler:
    HandleError "GetPlayerY", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerY(ByVal index As Long, ByVal yVal As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If index <= 0 Or index > MAX_PLAYER Then Exit Sub
    Player(index).PlayerData(TempPlayer(index).CurSlot).y = yVal

    Exit Sub
errHandler:
    HandleError "SetPlayerY", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerDir(ByVal index As Long) As Byte
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If index <= 0 Or index > MAX_PLAYER Then Exit Function
    GetPlayerDir = Player(index).PlayerData(TempPlayer(index).CurSlot).Dir

    Exit Function
errHandler:
    HandleError "GetPlayerDir", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerDir(ByVal index As Long, ByVal xDir As Byte)
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If index <= 0 Or index > MAX_PLAYER Then Exit Sub
    Player(index).PlayerData(TempPlayer(index).CurSlot).Dir = xDir

    Exit Sub
errHandler:
    HandleError "SetPlayerDir", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerAccess(ByVal index As Long) As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If index <= 0 Or index > MAX_PLAYER Then Exit Function
    If TempPlayer(index).CurSlot = 0 Then Exit Function
    GetPlayerAccess = Player(index).PlayerData(TempPlayer(index).CurSlot).Access
    
    Exit Function
errHandler:
    HandleError "GetPlayerDir", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub PlayerWarp(ByVal index As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
Dim OldMap As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Not IsPlaying(index) Then Exit Sub
    If MapNum <= 0 Or MapNum > Count_Map Then Exit Sub
    
    If x > Map(MapNum).MaxX Then x = Map(MapNum).MaxX
    If y > Map(MapNum).MaxY Then y = Map(MapNum).MaxY
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    
    If MapNum = GetPlayerMap(index) Then
        SendPlayerXYToMap index
    End If
    
    OldMap = GetPlayerMap(index)
    If OldMap <> MapNum Then
        SendLeaveMap index, OldMap
        ClearAllTarget index
        TempPlayer(index).target = 0
        SendTarget index
    End If
    
    SetPlayerMap index, MapNum
    SetPlayerX index, x
    SetPlayerY index, y
    
    PlayerOnMap(MapNum) = True
    TempPlayer(index).GettingMap = True
    SendCheckForMap index, MapNum
    
    Exit Sub
errHandler:
    HandleError "PlayerWarp", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal Movement As Long, Optional ByVal sendToSelf As Boolean = False)
Dim Moved As Byte
Dim NewMapY As Long, NewMapX As Long
Dim Lvl As Long, Num As Long
Dim GetAppear As Long
Dim MapNum As Long, x As Long, y As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Not IsPlaying(index) Or Dir < DIR_UP Or Dir > DIR_DOWN Or Movement < 1 Or Movement > 2 Then Exit Sub

    Call SetPlayerDir(index, Dir)
    Moved = NO
    
    Select Case Dir
        Case DIR_UP
            If GetPlayerY(index) > 0 Then
                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> Attributes.Blocked Then
                    Call SetPlayerY(index, GetPlayerY(index) - 1)
                    SendPlayerMove index, Movement, sendToSelf
                    Moved = YES
                End If
            Else
                If Map(GetPlayerMap(index)).Link(DIR_UP) > 0 Then
                    NewMapY = Map(Map(GetPlayerMap(index)).Link(DIR_UP)).MaxY
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Link(DIR_UP), GetPlayerX(index), NewMapY)
                    Moved = YES
                End If
            End If

        Case DIR_DOWN
            If GetPlayerY(index) < Map(GetPlayerMap(index)).MaxY Then
                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> Attributes.Blocked Then
                    Call SetPlayerY(index, GetPlayerY(index) + 1)
                    SendPlayerMove index, Movement, sendToSelf
                    Moved = YES
                End If
            Else
                If Map(GetPlayerMap(index)).Link(DIR_DOWN) > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Link(DIR_DOWN), GetPlayerX(index), 0)
                    Moved = YES
                End If
            End If

        Case DIR_LEFT
            If GetPlayerX(index) > 0 Then
                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> Attributes.Blocked Then
                    Call SetPlayerX(index, GetPlayerX(index) - 1)
                    SendPlayerMove index, Movement, sendToSelf
                    Moved = YES
                End If
            Else
                If Map(GetPlayerMap(index)).Link(DIR_LEFT) > 0 Then
                    NewMapX = Map(Map(GetPlayerMap(index)).Link(DIR_LEFT)).MaxX
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Link(DIR_LEFT), NewMapX, GetPlayerY(index))
                    Moved = YES
                End If
            End If

        Case DIR_RIGHT
            If GetPlayerX(index) < Map(GetPlayerMap(index)).MaxX Then
                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> Attributes.Blocked Then
                    Call SetPlayerX(index, GetPlayerX(index) + 1)
                    SendPlayerMove index, Movement, sendToSelf
                    Moved = YES
                End If
            Else
                If Map(GetPlayerMap(index)).Link(DIR_RIGHT) > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Link(DIR_RIGHT), 0, GetPlayerY(index))
                    Moved = YES
                End If
            End If
    End Select

    With Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index))
        If .Type = Attributes.mShop Then
            x = .Data1
            SendOpenShop index, x
            Moved = YES
        End If
        
        If .Type = Attributes.Warp Then
            MapNum = .Data1
            x = .Data2
            y = .Data3
            If MapNum > 0 And MapNum <= Count_Map Then
                Call PlayerWarp(index, MapNum, x, y)
            End If
            Moved = YES
        End If
        
        If .Type = Attributes.TallGrass Then
            Lvl = Random(Map(GetPlayerMap(index)).MinLvl, Map(GetPlayerMap(index)).MaxLvl)
            GetAppear = Random(1, 3)
            If GetAppear = 1 Then
                Num = Random(1, MAX_MAP_POKEMON)
                If Map(GetPlayerMap(index)).Pokemon(Num) > 0 Then
                    If CheckPokemon(index) > 0 Then
                        InitPlayerVsNpc index, Map(GetPlayerMap(index)).Pokemon(Num), Lvl
                    End If
                End If
            End If
            Moved = YES
        End If
        
        If .Type = Attributes.Heal Then
            RestoreAllPokemon index
            Moved = YES
        End If
        
        If .Type = Attributes.Checkpoint Then
            Player(index).PlayerData(TempPlayer(index).CurSlot).Checkpoint.Map = GetPlayerMap(index)
            Player(index).PlayerData(TempPlayer(index).CurSlot).Checkpoint.x = GetPlayerX(index)
            Player(index).PlayerData(TempPlayer(index).CurSlot).Checkpoint.y = GetPlayerY(index)
            Moved = YES
        End If
        
        If .Type = Attributes.Storage Then
            InitStorage index
            Moved = YES
        End If
    End With
    
    If Moved = NO Then PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
    
    Exit Sub
errHandler:
    HandleError "PlayerMove", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function FindPlayer(ByVal Name As String) As Long
Dim i As Long
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To HighPlayerIndex
        If IsPlaying(i) Then
            If (GetPlayerName(i)) = (Trim$(Name)) Then
                FindPlayer = i
                Exit Function
            End If
        End If
    Next
    FindPlayer = 0
    
    Exit Function
errHandler:
    HandleError "FindPlayer", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub GivePlayerPokemon(ByVal index As Long, ByVal PokemonNum As Long, ByVal Level As Long)
Dim Slot As Byte
Dim i As Long, n As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Slot = FindPokemonSlot(index)
    If Slot <= 0 Or Slot > MAX_POKEMON Then Exit Sub
    
    With Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(Slot)
        .Num = PokemonNum
        If Rnd <= Pokemon(PokemonNum).FemaleRate Then
            .Gender = GENDER_FEMALE
        Else
            .Gender = GENDER_MALE
        End If
        .Level = Level
        For i = 1 To Stats.Stat_Count - 1
            .Stat(i) = GetPlayerPokeSlotStat(index, Slot, i)
            .StatIV(i) = Random(1, 31)
            .StatEV(i) = 0
        Next
        .CurHP = .Stat(Stats.HP)
        .Exp = 0
        
        GetPlayerPokemonMove index, Slot
    End With
    SendPlayerPokemon index, Slot

    SavePlayerPokemon index

    Exit Sub
errHandler:
    HandleError "GivePlayerPokemon", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub TakePlayerPokemon(ByVal index As Long, ByVal Slot As Byte)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Slot <= 0 Or Slot > MAX_POKEMON Then Exit Sub

    Call ZeroMemory(ByVal VarPtr(Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(Slot)), LenB(Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(Slot)))
    SendPlayerPokemon index, Slot
    UpdatePokemon index
    
    SavePlayerPokemon index
    
    Exit Sub
errHandler:
    HandleError "TakePlayerPokemon", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function FindPokemonSlot(ByVal index As Long) As Byte
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    FindPokemonSlot = 0
    For i = MAX_POKEMON To 1 Step -1
        If Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(i).Num = 0 Then
            FindPokemonSlot = i
        End If
    Next

    Exit Function
errHandler:
    HandleError "FindPokemonSlot", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub ClearAllTarget(ByVal index As Long)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To HighPlayerIndex
        If IsPlaying(i) Then
            If TempPlayer(i).target = index Then
                TempPlayer(i).target = 0
                SendTarget i
            End If
        End If
    Next

    Exit Sub
errHandler:
    HandleError "ClearAllTarget", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function CheckPokemon(ByVal index As Long) As Long
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    CheckPokemon = 0
    For i = MAX_POKEMON To 1 Step -1
        If Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(i).Num > 0 Then
            If Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(i).CurHP > 0 Then
                CheckPokemon = i
            End If
        End If
    Next

    Exit Function
errHandler:
    HandleError "CheckPokemon", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function CountPokemon(ByVal index As Long) As Long
Dim i As Long, count As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    count = 0
    For i = 1 To MAX_POKEMON
        If Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(i).Num > 0 Then
            count = count + 1
        End If
    Next
    CountPokemon = count

    Exit Function
errHandler:
    HandleError "CountPokemon", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub GetPlayerPokemonMove(ByVal index As Long, ByVal PokeSlot As Long)
Dim x As Long
Dim i As Long, MoveLvl As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For x = 1 To MAX_MOVES
        With Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(PokeSlot)
            If .Num > 0 Then
                If Pokemon(.Num).MoveNum(x) > 0 Then
                    i = CheckFreeMoveSlot(index, , PokeSlot)
                    MoveLvl = Pokemon(.Num).MoveLevel(x)
                    If i > 0 And MoveLvl <= .Level And Not CheckSameMove(index, x, , PokeSlot) Then
                        .Moves(i).Num = Pokemon(.Num).MoveNum(x)
                        .Moves(i).PP = Moves(Pokemon(.Num).MoveNum(x)).PP
                        .Moves(i).MaxPP = Moves(Pokemon(.Num).MoveNum(x)).PP
                    End If
                End If
            End If
        End With
    Next
    
    Exit Sub
errHandler:
    HandleError "GetPlayerPokemonMove", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub GetEnemyPokemonMove(ByVal index As Long)
Dim x As Long
Dim i As Long, MoveLvl As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For x = 1 To MAX_MOVES
        With TempPlayer(index).EnemyPokemon
            If .Num > 0 Then
                If Pokemon(.Num).MoveNum(x) > 0 Then
                    i = CheckFreeMoveSlot(index, YES)
                    MoveLvl = Pokemon(.Num).MoveLevel(x)
                    If ((i > 0) And (MoveLvl <= .Level) And Not CheckSameMove(index, x, YES)) Then
                        .Moves(i).Num = Pokemon(.Num).MoveNum(x)
                        .Moves(i).PP = Moves(Pokemon(.Num).MoveNum(x)).PP
                        .Moves(i).MaxPP = Moves(Pokemon(.Num).MoveNum(x)).PP
                        i = CheckFreeMoveSlot(index, YES)
                    End If
                End If
            End If
        End With
    Next
    
    Exit Sub
errHandler:
    HandleError "GetEnemyPokemonMove", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function CheckFreeMoveSlot(ByVal index As Long, Optional ByVal IsEnemy As Byte = 0, Optional ByVal PokeSlot As Long) As Long
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    CheckFreeMoveSlot = 0
    For i = 1 To MAX_POKEMON_MOVES
        If IsEnemy > 0 Then
            If TempPlayer(index).EnemyPokemon.Moves(i).Num = 0 Then
                CheckFreeMoveSlot = i
                Exit Function
            End If
        Else
            If Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(PokeSlot).Moves(i).Num = 0 Then
                CheckFreeMoveSlot = i
                Exit Function
            End If
        End If
    Next
    
    Exit Function
errHandler:
    HandleError "CheckFreeMoveSlot", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function CheckSameMove(ByVal index As Long, ByVal MoveNum As Long, Optional ByVal IsEnemy As Byte = 0, Optional ByVal PokeSlot As Long) As Boolean
Dim x As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    CheckSameMove = False
    For x = 1 To MAX_POKEMON_MOVES
        If IsEnemy > 0 Then
            If TempPlayer(index).EnemyPokemon.Moves(x).Num = MoveNum Then
                CheckSameMove = True
                Exit Function
            End If
        Else
            If Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(PokeSlot).Moves(x).Num = MoveNum Then
                CheckSameMove = True
                Exit Function
            End If
        End If
    Next

    Exit Function
errHandler:
    HandleError "CheckSameMove", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub UpdatePokemon(ByVal index As Long)
Dim i As Long
Dim TmpPokemon As PlayerPokemonRec

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    With Player(index).PlayerData(TempPlayer(index).CurSlot)
        i = 2
        Do While i <= MAX_POKEMON
            If .Pokemon(i).Num > 0 Then
                If .Pokemon(i - 1).Num = 0 Then
                    TmpPokemon = .Pokemon(i)
                    .Pokemon(i - 1) = TmpPokemon
                    Call ZeroMemory(ByVal VarPtr(.Pokemon(i)), LenB(.Pokemon(i)))
                    
                    SendPlayerPokemon index, i - 1
                    SendPlayerPokemon index, i
                End If
            End If
            i = i + 1
        Loop
    End With

    Exit Sub
errHandler:
    HandleError "UpdatePokemon", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub RestoreAllPokemon(ByVal index As Long)
Dim i As Long, x As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    With Player(index).PlayerData(TempPlayer(index).CurSlot)
        For i = 1 To MAX_POKEMON
            .Pokemon(i).CurHP = .Pokemon(i).Stat(Stats.HP)
            For x = 1 To MAX_POKEMON_MOVES
                .Pokemon(i).Moves(x).PP = .Pokemon(i).Moves(x).MaxPP
            Next
        Next
    End With
    SendPlayerPokemons index
    
    Exit Sub
errHandler:
    HandleError "RestoreAllPokemon", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function HandleLevel(ByVal index As Long, ByVal PokeSlot As Long, ByVal Exp As Long) As Boolean
Dim ExpRollover As Long
Dim NewMove As Long
Dim Lvl As Long
Dim i As Long, FreeMoveSlot As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    HandleLevel = False
    With Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(PokeSlot)
        If .Level >= MAX_LEVEL Then Exit Function
        If Player(index).PlayerData(TempPlayer(index).CurSlot).IsVIP = YES Then
            Exp = (Exp * 2)
        End If
        .Exp = .Exp + Exp
        Lvl = .Level
        For i = 1 To Stats.Stat_Count - 1
            If .StatEV(i) >= MAX_EV Then
                .StatEV(i) = MAX_EV
            Else
                .StatEV(i) = .StatEV(i) + (Sqr(Pokemon(TempPlayer(index).EnemyPokemon.Num).BaseStat(i)) / 4)
                If .StatEV(i) >= MAX_EV Then .StatEV(i) = MAX_EV
            End If
        Next
        Do While .Exp >= ExpCalc(Lvl) And .Level < MAX_LEVEL
            ExpRollover = .Exp - ExpCalc(Lvl)
            .Level = .Level + 1
            For i = 1 To Stats.Stat_Count - 1
                .Stat(i) = GetPlayerPokeSlotStat(index, PokeSlot, i)
            Next
            .Exp = ExpRollover
            NewMove = CheckLearnMove(index, PokeSlot, .Level)
            If NewMove > 0 Then
                FreeMoveSlot = CheckFreeMoveSlot(index, , PokeSlot)
                If FreeMoveSlot > 0 Then
                    LearnMove index, PokeSlot, FreeMoveSlot, NewMove
                Else
                    SendLearnMove index, NewMove, PokeSlot
                End If
            End If
            Lvl = Lvl + 1
            SendMsg index, Trim$(Pokemon(.Num).Name) & " grew to Lv." & .Level & "!", Green
            HandleLevel = True
        Loop
    End With
    SendPlayerPokemon index, PokeSlot
    
    Exit Function
errHandler:
    HandleError "HandleLevel", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function CheckLearnMove(ByVal index As Long, ByVal PokeSlot As Long, ByVal Lvl As Long) As Long
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    For i = 1 To MAX_MOVES
        With Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(PokeSlot)
            If Pokemon(.Num).MoveNum(i) > 0 Then
                If Pokemon(.Num).MoveLevel(i) = Lvl Then
                    CheckLearnMove = Pokemon(.Num).MoveNum(i)
                End If
            End If
        End With
    Next

    Exit Function
errHandler:
    HandleError "CheckLearnMove", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub LearnMove(ByVal index As Long, ByVal PokeSlot As Long, ByVal MoveSlot As Long, ByVal MoveNum As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    With Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(PokeSlot)
        .Moves(MoveSlot).Num = MoveNum
        .Moves(MoveSlot).MaxPP = Moves(MoveNum).PP
        .Moves(MoveSlot).PP = Moves(MoveNum).PP
        SendMsg index, Trim$(Pokemon(.Num).Name) & " learned " & Trim$(Moves(MoveNum).Name) & "!", Green
    End With
    SendPlayerPokemon index, PokeSlot

    Exit Sub
errHandler:
    HandleError "LearnMove", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function CheckEvolve(ByVal index As Long, ByVal PokeSlot As Long) As Boolean
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    CheckEvolve = False
    With Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(PokeSlot)
        If Pokemon(.Num).EvolveNum > 0 Then
            If .Level >= Pokemon(.Num).EvolveLvl Then
                CheckEvolve = True
            End If
        End If
    End With
    
    Exit Function
errHandler:
    HandleError "CheckEvolve", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function TryCatchPokemon(ByVal index As Long, ByVal BonusBall As Single) As Boolean
Dim a As Single, Bonusstatus As Single
Dim b() As Long, i As Long, bRate As Single

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ReDim b(1 To 4)
    
    Bonusstatus = 1
    With TempPlayer(index).EnemyPokemon
        a = (((3 * .Stat(Stats.HP) - 2 * .CurHP) * Pokemon(.Num).CatchRate * BonusBall) / (3 * .Stat(Stats.HP))) * Bonusstatus
        If a >= 255 Then
            TryCatchPokemon = True
            Exit Function
        Else
            bRate = ((2 ^ 16) - 1) * (4 * (Sqr((a / ((2 ^ 8) - 1)))))
            'bRate = 1048560 / (Sqr(Sqr((16711680 / a))))
            'bRate = ((2 ^ 16) - 1) * (4 * (Sqr(a))) / (4 * (Sqr(((2 ^ 8) - 1))))
            For i = 1 To 4
                b(i) = Random(0, 65535)
                If bRate >= b(i) Then
                    TryCatchPokemon = True
                    Exit Function
                End If
                SendBattleMsg index, "The pokeball shake..", White
            Next
        End If
    End With
    TryCatchPokemon = False
    
    Exit Function
errHandler:
    HandleError "TryCatchPokemon", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub GiveItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim InvSlot As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    With Player(index).PlayerData(TempPlayer(index).CurSlot)
        If ItemNum > 0 Then
            InvSlot = FindInvFreeSlot(index, ItemNum, ItemVal)
            If InvSlot > 0 Then
                .Item(InvSlot, Item(ItemNum).Type).Num = ItemNum
                If Not Item(ItemNum).Type = ItemType.KeyItems Then
                    .Item(InvSlot, Item(ItemNum).Type).Value = .Item(InvSlot, Item(ItemNum).Type).Value + ItemVal
                End If
            End If
        End If
    End With
    UpdateInventory index
    SendInventory index
    
    SavePlayer index

    Exit Sub
errHandler:
    HandleError "GiveItem", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub TakeItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
Dim InvSlot As Long, IType As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    With Player(index).PlayerData(TempPlayer(index).CurSlot)
        If ItemNum > 0 Then
            IType = Item(ItemNum).Type
            InvSlot = FindInvItemNum(index, ItemNum)
            If InvSlot > 0 Then
                Select Case IType
                    Case ItemType.KeyItems
                        .Item(InvSlot, IType).Num = 0
                        .Item(InvSlot, IType).Value = 0
                    Case Else
                        If .Item(InvSlot, IType).Value > 0 Then
                            .Item(InvSlot, IType).Value = .Item(InvSlot, IType).Value - ItemVal
                            If .Item(InvSlot, IType).Value <= 0 Then
                                .Item(InvSlot, IType).Num = 0
                                .Item(InvSlot, IType).Value = 0
                            End If
                        End If
                End Select
            End If
        End If
    End With
    UpdateInventory index
    SendInventory index
    
    SavePlayer index

    Exit Sub
errHandler:
    HandleError "TakeItem", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function FindInvItemNum(ByVal index As Long, ByVal ItemNum As Long)
Dim i As Long, IType As Byte
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    FindInvItemNum = 0
    With Player(index).PlayerData(TempPlayer(index).CurSlot)
        IType = Item(ItemNum).Type
        For i = 1 To MAX_PLAYER_ITEM
            If .Item(i, IType).Num = ItemNum Then
                FindInvItemNum = i
                Exit Function
            End If
        Next
    End With
    
    Exit Function
errHandler:
    HandleError "FindInvItemNum", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function FindInvFreeSlot(ByVal index As Long, ByVal ItemNum As Long, Optional ByVal vVal As Long = 0) As Long
Dim i As Long, IType As Byte
Dim x As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    FindInvFreeSlot = 0
    With Player(index).PlayerData(TempPlayer(index).CurSlot)
        IType = Item(ItemNum).Type
        Select Case IType
            Case ItemType.KeyItems
                For i = 1 To MAX_PLAYER_ITEM
                    If .Item(i, IType).Num = 0 Then
                        FindInvFreeSlot = i
                        Exit Function
                    End If
                Next
            Case Else
                x = GetSameInvItem(index, ItemNum, vVal)
                If x > 0 Then
                    FindInvFreeSlot = x
                Else
                    For i = 1 To MAX_PLAYER_ITEM
                        If .Item(i, IType).Num = 0 Then
                            FindInvFreeSlot = i
                            Exit Function
                        End If
                    Next
                End If
        End Select
    End With
    
    Exit Function
errHandler:
    HandleError "FindInvFreeSlot", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function GetSameInvItem(ByVal index As Long, ByVal ItemNum As Long, ByVal vVal As Long) As Long
Dim i As Long, IType As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    IType = Item(ItemNum).Type
    GetSameInvItem = 0
    With Player(index).PlayerData(TempPlayer(index).CurSlot)
        For i = 1 To MAX_PLAYER_ITEM
            If .Item(i, IType).Num = ItemNum Then
                If .Item(i, IType).Value + vVal <= MAX_PLAYER_INV_VALUE Then
                    GetSameInvItem = i
                    Exit Function
                End If
            End If
        Next
    End With
    
    Exit Function
errHandler:
    HandleError "GetSameInvItem", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub UpdateInventory(ByVal index As Long)
Dim x As Long, y As Long
Dim TmpInv As PlayerItemRec

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    With Player(index).PlayerData(TempPlayer(index).CurSlot)
        For x = 0 To ItemType.Item_Count - 1
            y = 2
            Do While y <= MAX_PLAYER_ITEM And .Item(y - 1, x).Num = 0
                If .Item(y, x).Num > 0 Then
                    TmpInv = .Item(y, x)
                    .Item(y - 1, x) = TmpInv
                    Call ZeroMemory(ByVal VarPtr(.Item(y, x)), LenB(.Item(y, x)))
                End If
                y = y + 1
            Loop
        Next
    End With

    Exit Sub
errHandler:
    HandleError "UpdateInventory", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub UseItem(ByVal index As Long, ByVal ItemSlot As Long, ByVal CurInvType As Byte)
Dim DidCatch As Boolean
Dim x As Long, i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If TempPlayer(index).MoveSet = 5 Then Exit Sub
    If ItemSlot <= 0 Then Exit Sub
    
    With Player(index).PlayerData(TempPlayer(index).CurSlot)
        If .Item(ItemSlot, CurInvType).Num = 0 Then Exit Sub
        
        Select Case Item(.Item(ItemSlot, CurInvType).Num).Type
            Case ItemType.Items
                Select Case Item(.Item(ItemSlot, CurInvType).Num).IType
                    Case ItemProperties.RestoreHP, ItemProperties.RestorePP
                        If CountPlayerPokemon(index) > 0 Then
                            SendSelect index, SELECT_POKEMON, .Item(ItemSlot, CurInvType).Num
                        End If
                End Select
            Case ItemType.Pokeballs
                If TempPlayer(index).InBattle = BATTLE_WILD Then
                    DidCatch = TryCatchPokemon(index, Item(.Item(ItemSlot, CurInvType).Num).Data3)
                    If DidCatch Then
                        x = FindPokemonSlot(index)
                        If x > 0 Then
                            .Pokemon(x) = TempPlayer(index).EnemyPokemon
                            .Pokemon(x).Exp = 0
                            SendPlayerPokemon index, x
                            
                            SendBattleMsg index, "You successfully captured " & Trim$(Pokemon(TempPlayer(index).EnemyPokemon.Num).Name), White
                            SendCaptured index
                            ExitBattle index, NO
                        Else
                            x = FindPokemonStorage(index)
                            If x > 0 Then
                                .StoredPokemon(x) = TempPlayer(index).EnemyPokemon
                                .StoredPokemon(x).CurHP = .StoredPokemon(x).Stat(Stats.HP)
                                For i = 1 To MAX_POKEMON_MOVES
                                    .StoredPokemon(x).Moves(i).PP = .StoredPokemon(x).Moves(i).MaxPP
                                Next
                                SendPlayerStoredPokemon index, x
                                
                                SendBattleMsg index, "You successfully captured " & Trim$(Pokemon(TempPlayer(index).EnemyPokemon.Num).Name), White
                                SendCaptured index
                                ExitBattle index, NO
                            End If
                        End If
                    Else
                        SendBattleMsg index, "Wild " & Trim$(Pokemon(TempPlayer(index).EnemyPokemon.Num).Name) & " broke the ball!", White
                        NpcVsPlayer index
                        If .Pokemon(TempPlayer(index).InBattlePoke).CurHP <= 0 Then
                            If CheckPokemon(index) > 0 Then
                                SendForceSwitch index
                            Else
                                ExitBattle index, YES
                            End If
                        End If
                        SendBattleMsg index, EndLine, Cyan
                    End If
                    TakeItem index, .Item(ItemSlot, CurInvType).Num, 1
                ElseIf TempPlayer(index).InBattle = BATTLE_TRAINER Then
                    SendBattleMsg index, "You cannot use Pokeballs on Trainer's Pokemon", Red
                End If
        End Select
    End With

    Exit Sub
errHandler:
    HandleError "UseItem", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub InitSelectUseItem(ByVal index As Long, ByVal Data1 As Long, ByVal Data2 As Long, ByVal Data3 As Long, ByVal Data4 As Long)
Dim FoeIndex As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If TempPlayer(index).MoveSet = 5 Then Exit Sub
    
    With Player(index).PlayerData(TempPlayer(index).CurSlot)
        If Data2 > 0 Then
            Select Case Item(Data2).Type
                Case ItemType.Items
                    Select Case Item(Data2).IType
                        Case ItemProperties.RestoreHP
                            If Data1 = SELECT_POKEMON Then
                                .Pokemon(Data3).CurHP = .Pokemon(Data3).CurHP + Item(Data2).Data2
                                If .Pokemon(Data3).CurHP > .Pokemon(Data3).Stat(Stats.HP) Then
                                    .Pokemon(Data3).CurHP = .Pokemon(Data3).Stat(Stats.HP)
                                End If
                                SendUpdatePokemonVital index, index, Data3
                                If TempPlayer(index).InBattle = BATTLE_TRAINER Then
                                    'FoeIndex = TempPlayer(index).BattleRequest
                                    'TempPlayer(FoeIndex).EnemyPokemon = .Pokemon(Data3)
                                    'SendUpdateEnemyVital FoeIndex
                                End If
                                SendMsg index, Trim$(Pokemon(.Pokemon(Data3).Num).Name) & " HP have been restored!", Green
                            End If
                            TakeItem index, Data2, 1
                            If TempPlayer(index).InBattle = BATTLE_WILD Then
                                NpcVsPlayer index
                                If .Pokemon(Data3).CurHP <= 0 Then
                                    If CheckPokemon(index) > 0 Then
                                        SendForceSwitch index
                                    Else
                                        ExitBattle index, YES
                                    End If
                                End If
                                SendBattleMsg index, EndLine, Cyan
                            ElseIf TempPlayer(index).InBattle = BATTLE_TRAINER Then
                                TempPlayer(index).MoveSet = 5
                            End If
                        Case ItemProperties.RestorePP
                            If Data1 = SELECT_POKEMON Then
                                If CountPlayerPokemonMove(index, Data3) Then
                                    SendSelect index, SELECT_MOVE, Data2, Data3
                                End If
                            ElseIf Data1 = SELECT_MOVE Then
                                .Pokemon(Data4).Moves(Data3).PP = .Pokemon(Data4).Moves(Data3).PP + Item(Data2).Data2
                                If .Pokemon(Data4).Moves(Data3).PP > .Pokemon(Data4).Moves(Data3).MaxPP Then
                                    .Pokemon(Data4).Moves(Data3).PP = .Pokemon(Data4).Moves(Data3).MaxPP
                                End If
                                SendUpdatePokemonVital index, index, Data4
                                FoeIndex = TempPlayer(index).BattleRequest
                                TempPlayer(FoeIndex).EnemyPokemon = .Pokemon(Data4)
                                SendUpdateEnemyVital FoeIndex
                                SendMsg index, Trim$(Pokemon(.Pokemon(Data4).Num).Name) & "'s " & Trim$(Moves(.Pokemon(Data4).Moves(Data3).Num).Name) & " PP have been restored!", Green
                            End If
                            TakeItem index, Data2, 1
                            If TempPlayer(index).InBattle = BATTLE_WILD Then
                                NpcVsPlayer index
                                If .Pokemon(Data4).CurHP <= 0 Then
                                    If CheckPokemon(index) > 0 Then
                                        SendForceSwitch index
                                    Else
                                        ExitBattle index, YES
                                    End If
                                End If
                                SendBattleMsg index, EndLine, Cyan
                            ElseIf TempPlayer(index).InBattle = BATTLE_TRAINER Then
                                TempPlayer(index).MoveSet = 5
                            End If
                    End Select
            End Select
        End If
    End With

    Exit Sub
errHandler:
    HandleError "InitSelectUseItem", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function CountPlayerPokemon(ByVal index As Long) As Long
Dim i As Long, x As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    x = 0
    For i = 1 To MAX_POKEMON
        If Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(i).Num > 0 Then
            x = x + 1
        End If
    Next
    CountPlayerPokemon = x

    Exit Function
errHandler:
    HandleError "CountPlayerPokemon", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function CountPlayerPokemonMove(ByVal index As Long, ByVal PokeSlot As Long) As Long
Dim i As Long, x As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    x = 0
    For i = 1 To MAX_POKEMON_MOVES
        If Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(PokeSlot).Moves(i).Num > 0 Then
            x = x + 1
        End If
    Next
    CountPlayerPokemonMove = x

    Exit Function
errHandler:
    HandleError "CountPlayerPokemonMove", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub ClosePlayerBattle(ByVal index As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    TempPlayer(index).InBattle = 0
    TempPlayer(index).InBattlePoke = 0
    TempPlayer(index).EscCount = 0
    TempPlayer(index).BattleRequest = 0
    ClearEnemyPokemon index
    SendBattle index
    SendExitBattle index

    Exit Sub
errHandler:
    HandleError "ClosePlayerBattle", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub GivePvPPoints(ByVal index As Long, ByVal PointType As Byte)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    With Player(index).PlayerData(TempPlayer(index).CurSlot).PvP
        Select Case PointType
            Case 1
                .Win = .Win + 1
            Case 2
                .Lose = .Lose + 1
            Case 3
                .Disconnect = .Disconnect + 1
        End Select
    End With
    SendPlayerData index

    Exit Sub
errHandler:
    HandleError "GivePvPPoints", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub DepositPokemon(ByVal index As Long, ByVal PokeSlot As Long)
Dim i As Long
Dim TmpPoke As PlayerPokemonRec

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    With Player(index).PlayerData(TempPlayer(index).CurSlot)
        i = FindPokemonStorage(index)
        If .Pokemon(PokeSlot).Num > 0 Then
            If i > 0 Then
                TmpPoke = .Pokemon(PokeSlot)
                Call ZeroMemory(ByVal VarPtr(.Pokemon(PokeSlot)), LenB(.Pokemon(PokeSlot)))
                .StoredPokemon(i) = TmpPoke
                UpdatePokemon index
                SendPlayerPokemon index, PokeSlot
                SendPlayerStoredPokemon index, i
            Else
                SendMsg index, "PC is already full!", Red
            End If
        End If
    End With

    Exit Sub
errHandler:
    HandleError "DepositPokemon", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub WithdrawPokemon(ByVal index As Long, ByVal StorageSlot As Long)
Dim i As Long
Dim TmpPoke As PlayerPokemonRec

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    With Player(index).PlayerData(TempPlayer(index).CurSlot)
        i = FindPokemonSlot(index)
        If .StoredPokemon(StorageSlot).Num > 0 Then
            If i > 0 Then
                TmpPoke = .StoredPokemon(StorageSlot)
                Call ZeroMemory(ByVal VarPtr(.StoredPokemon(StorageSlot)), LenB(.StoredPokemon(StorageSlot)))
                .Pokemon(i) = TmpPoke
                UpdateStoredPokemon index
                SendPlayerPokemon index, i
                SendPlayerStoredPokemon index, StorageSlot
            Else
                SendMsg index, "Pokemon slot is already full!", Red
            End If
        End If
    End With

    Exit Sub
errHandler:
    HandleError "WithdrawPokemon", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function FindPokemonStorage(ByVal index As Long) As Long
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To MAX_STORAGE_POKEMON
        If Player(index).PlayerData(TempPlayer(index).CurSlot).StoredPokemon(i).Num = 0 Then
            FindPokemonStorage = i
            Exit Function
        End If
    Next

    Exit Function
errHandler:
    HandleError "FindPokemonStorage", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub UpdateStoredPokemon(ByVal index As Long)
Dim i As Long
Dim TmpPokemon As PlayerPokemonRec

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    With Player(index).PlayerData(TempPlayer(index).CurSlot)
        i = 2
        Do While i <= MAX_STORAGE_POKEMON
            If .StoredPokemon(i).Num > 0 Then
                If .StoredPokemon(i - 1).Num = 0 Then
                    TmpPokemon = .StoredPokemon(i)
                    .StoredPokemon(i - 1) = TmpPokemon
                    Call ZeroMemory(ByVal VarPtr(.StoredPokemon(i)), LenB(.StoredPokemon(i)))
    
                    SendPlayerStoredPokemon index, i - 1
                    SendPlayerStoredPokemon index, i
                End If
            End If
            i = i + 1
        Loop
    End With

    Exit Sub
errHandler:
    HandleError "UpdateStoredPokemon", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub InitStorage(ByVal index As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    SendStorage index
    
    Exit Sub
errHandler:
    HandleError "InitStorage", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub GiveMoney(ByVal index As Long, ByVal Value As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    SendMsg index, "You got $" & Value, White
    Player(index).PlayerData(TempPlayer(index).CurSlot).Money = Player(index).PlayerData(TempPlayer(index).CurSlot).Money + Value
    SendPlayerData index
    
    Exit Sub
errHandler:
    HandleError "GiveMoney", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function CanAfford(ByVal index As Long, ByVal Value As Long) As Boolean
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    CanAfford = False
    If Player(index).PlayerData(TempPlayer(index).CurSlot).Money >= Value Then
        CanAfford = True
    End If

    Exit Function
errHandler:
    HandleError "CanAfford", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub TakeMoney(ByVal index As Long, ByVal Value As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Player(index).PlayerData(TempPlayer(index).CurSlot).Money = Player(index).PlayerData(TempPlayer(index).CurSlot).Money - Value
    SendPlayerData index
    
    Exit Sub
errHandler:
    HandleError "TakeMoney", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function CheckPlayerOnMap(ByVal MapNum As Long) As Long
Dim i As Long
Dim count As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    count = 0
    For i = 1 To MAX_PLAYER
        If TempPlayer(i).CurSlot > 0 Then
            If IsPlaying(i) And Player(i).PlayerData(TempPlayer(i).CurSlot).Map = MapNum Then
                count = count + 1
            End If
        End If
    Next
    CheckPlayerOnMap = count
    

    Exit Function
errHandler:
    HandleError "CanAfford", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function
