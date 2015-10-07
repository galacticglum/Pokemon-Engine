Attribute VB_Name = "modPlayer"
Option Explicit

Public Function GetPlayerSprite(ByVal Index As Long) As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Index <= 0 Or Index > MAX_PLAYER Then Exit Function
    GetPlayerSprite = Player(Index).Sprite

    Exit Function
errHandler:
    HandleError "GetPlayerSprite", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerSprite(ByVal Index As Long, ByVal SpriteNum As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Index <= 0 Or Index > MAX_PLAYER Then Exit Sub
    Player(Index).Sprite = SpriteNum

    Exit Sub
errHandler:
    HandleError "SetPlayerSprite", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerName(ByVal Index As Long) As String
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Index <= 0 Or Index > MAX_PLAYER Then Exit Function
    GetPlayerName = Trim$(Player(Index).Name)

    Exit Function
errHandler:
    HandleError "GetPlayerName", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function GetPlayerGender(ByVal Index As Long) As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Index <= 0 Or Index > MAX_PLAYER Then Exit Function
    GetPlayerGender = Player(Index).Gender

    Exit Function
errHandler:
    HandleError "GetPlayerGender", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function GetPlayerMap(ByVal Index As Long) As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Index <= 0 Or Index > MAX_PLAYER Then Exit Function
    GetPlayerMap = Player(Index).Map

    Exit Function
errHandler:
    HandleError "GetPlayerMap", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Index <= 0 Or Index > MAX_PLAYER Then Exit Sub
    Player(Index).Map = MapNum

    Exit Sub
errHandler:
    HandleError "SetPlayerMap", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerX(ByVal Index As Long) As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Index <= 0 Or Index > MAX_PLAYER Then Exit Function
    GetPlayerX = Player(Index).X

    Exit Function
errHandler:
    HandleError "GetPlayerX", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerX(ByVal Index As Long, ByVal xVal As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Index <= 0 Or Index > MAX_PLAYER Then Exit Sub
    Player(Index).X = xVal

    Exit Sub
errHandler:
    HandleError "SetPlayerX", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerY(ByVal Index As Long) As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Index <= 0 Or Index > MAX_PLAYER Then Exit Function
    GetPlayerY = Player(Index).Y

    Exit Function
errHandler:
    HandleError "GetPlayerY", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerY(ByVal Index As Long, ByVal yVal As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Index <= 0 Or Index > MAX_PLAYER Then Exit Sub
    Player(Index).Y = yVal

    Exit Sub
errHandler:
    HandleError "SetPlayerY", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerDir(ByVal Index As Long) As Byte
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Index <= 0 Or Index > MAX_PLAYER Then Exit Function
    GetPlayerDir = Player(Index).Dir

    Exit Function
errHandler:
    HandleError "GetPlayerDir", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerDir(ByVal Index As Long, ByVal xDir As Byte)
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Index <= 0 Or Index > MAX_PLAYER Then Exit Sub
    Player(Index).Dir = xDir

    Exit Sub
errHandler:
    HandleError "SetPlayerDir", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerAccess(ByVal Index As Long) As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Index <= 0 Or Index > MAX_PLAYER Then Exit Function
    GetPlayerAccess = Player(Index).Access
    
    Exit Function
errHandler:
    HandleError "GetPlayerDir", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function IsTryingToMove() As Boolean
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If HoldDown Or HoldUp Or HoldLeft Or HoldRight Then
        IsTryingToMove = True
    End If
    
    Exit Function
errHandler:
    HandleError "IsTryingToMove", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub ProcessMovement(ByVal Index As Long)
Dim MovementSpeed As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Select Case Player(Index).Moving
        Case MOVING_WALKING: MovementSpeed = WALK_SPEED
        Case Else: Exit Sub
    End Select
    
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            Player(Index).yOffset = Player(Index).yOffset - MovementSpeed
            If Player(Index).yOffset < 0 Then Player(Index).yOffset = 0
        Case DIR_DOWN
            Player(Index).yOffset = Player(Index).yOffset + MovementSpeed
            If Player(Index).yOffset > 0 Then Player(Index).yOffset = 0
        Case DIR_LEFT
            Player(Index).xOffset = Player(Index).xOffset - MovementSpeed
            If Player(Index).xOffset < 0 Then Player(Index).xOffset = 0
        Case DIR_RIGHT
            Player(Index).xOffset = Player(Index).xOffset + MovementSpeed
            If Player(Index).xOffset > 0 Then Player(Index).xOffset = 0
    End Select

    If Player(Index).Moving > 0 Then
        If GetPlayerDir(Index) = DIR_RIGHT Or GetPlayerDir(Index) = DIR_DOWN Then
            If (Player(Index).xOffset >= 0) And (Player(Index).yOffset >= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 0 Then
                    Player(Index).Step = 2
                Else
                    Player(Index).Step = 0
                End If
            End If
        Else
            If (Player(Index).xOffset <= 0) And (Player(Index).yOffset <= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 0 Then
                    Player(Index).Step = 2
                Else
                    Player(Index).Step = 0
                End If
            End If
        End If
    End If
    
    Exit Sub
errHandler:
    HandleError "ProcessMovement", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function CanMove() As Boolean
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    CanMove = True

    If Player(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If
    
    If IsLearnMove Or IsEvolve Then
        CanMove = False
        Exit Function
    End If
    
    If Player(MyIndex).InBattle > 0 Then
        CanMove = False
        Exit Function
    End If
    
    If InStorage Then
        CanMove = False
        Exit Function
    End If
    
    If InShop > 0 Then
        CanMove = False
        Exit Function
    End If

    If InTradeConfirm Or InTrade Then
        CanMove = False
        Exit Function
    End If

    If HoldUp Then
        Call SetPlayerDir(MyIndex, DIR_UP)

        If GetPlayerY(MyIndex) > 0 Then
            If CheckDirection(DIR_UP) Then
                CanMove = False

                If GetPlayerDir(MyIndex) <> DIR_UP Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If
        Else
            If Editor = 0 Then
                If Map.Link(DIR_UP) > 0 Then
                    SendPlayerRequestNewMap
                    GettingMap = True
                    CanMoveNow = False
                End If
                CanMove = False
                Exit Function
            End If
        End If
    End If

    If HoldDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)

        If GetPlayerY(MyIndex) < Map.MaxY Then
            If CheckDirection(DIR_DOWN) Then
                CanMove = False

                If GetPlayerDir(MyIndex) <> DIR_DOWN Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If
        Else
            If Editor = 0 Then
                If Map.Link(DIR_DOWN) > 0 Then
                    SendPlayerRequestNewMap
                    GettingMap = True
                    CanMoveNow = False
                End If
                CanMove = False
                Exit Function
            End If
        End If
    End If

    If HoldLeft Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)

        If GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_LEFT) Then
                CanMove = False

                If GetPlayerDir(MyIndex) <> DIR_LEFT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If
        Else
            If Editor = 0 Then
                If Map.Link(DIR_LEFT) > 0 Then
                    SendPlayerRequestNewMap
                    GettingMap = True
                    CanMoveNow = False
                End If
                CanMove = False
                Exit Function
            End If
        End If
    End If

    If HoldRight Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)

        If GetPlayerX(MyIndex) < Map.MaxX Then
            If CheckDirection(DIR_RIGHT) Then
                CanMove = False

                If GetPlayerDir(MyIndex) <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If

                Exit Function
            End If
        Else
            If Editor = 0 Then
                If Map.Link(DIR_RIGHT) > 0 Then
                    SendPlayerRequestNewMap
                    GettingMap = True
                    CanMoveNow = False
                End If
                CanMove = False
                Exit Function
            End If
        End If
    End If

    Exit Function
errHandler:
    HandleError "CanMove", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function CheckDirection(ByVal Direction As Byte) As Boolean
Dim X As Long, Y As Long
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    CheckDirection = False
    Select Case Direction
        Case DIR_UP
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) - 1
        Case DIR_DOWN
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) + 1
        Case DIR_LEFT
            X = GetPlayerX(MyIndex) - 1
            Y = GetPlayerY(MyIndex)
        Case DIR_RIGHT
            X = GetPlayerX(MyIndex) + 1
            Y = GetPlayerY(MyIndex)
    End Select

    If Map.Tile(X, Y).Type = Attributes.Blocked Then
        CheckDirection = True
        Exit Function
    End If
    
    For i = 1 To MAX_MAP_NPC
        If MapNpc(i).Num > 0 Then
            If MapNpc(i).X = X Then
                If MapNpc(i).Y = Y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next

    Exit Function
errHandler:
    HandleError "CheckDirection", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub CheckMovement()
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If IsTryingToMove Then
        If CanMove Then
            If Not CtrlDown Then
                Player(MyIndex).Moving = MOVING_WALKING
    
                Select Case GetPlayerDir(MyIndex)
                    Case DIR_UP
                        Call SendPlayerMove
                        Player(MyIndex).yOffset = Pic_Size
                        Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                    Case DIR_DOWN
                        Call SendPlayerMove
                        Player(MyIndex).yOffset = Pic_Size * -1
                        Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                    Case DIR_LEFT
                        Call SendPlayerMove
                        Player(MyIndex).xOffset = Pic_Size
                        Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                    Case DIR_RIGHT
                        Call SendPlayerMove
                        Player(MyIndex).xOffset = Pic_Size * -1
                        Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
                End Select
                
                If Player(MyIndex).X >= 0 And Player(MyIndex).X <= Map.MaxX Then
                    If Player(MyIndex).Y >= 0 And Player(MyIndex).Y <= Map.MaxY Then
                        If Map.Tile(Player(MyIndex).X, Player(MyIndex).Y).Type = Attributes.Warp Then
                            GettingMap = True
                        End If
                    End If
                End If
            End If
        End If
    End If
 
    Exit Sub
errHandler:
    HandleError "CheckMovement", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub InitBattle()
Dim X As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    MapBackground = Map.CurBack
    MapField = Map.CurField

    For X = ButtonEnum.BattleFight To ButtonEnum.BattleRun
        Buttons(X).Visible = True
    Next X
    For X = ButtonEnum.bBattleScrollUp To ButtonEnum.bBattleScrollDown
        Buttons(X).Visible = True
    Next X
    BattleScroll = MaxBattleLine
    
    WindowVisible(WindowType.Main_Trainer) = False
    WindowVisible(WindowType.Main_Inventory) = False
    WindowVisible(WindowType.Main_Option) = False
    For X = ButtonEnum.InvScrollUp To ButtonEnum.InvScrollDown
        Buttons(X).Visible = False
    Next X
    Capture = 255
    
    InStorage = False
    InShop = 0
    InTradeConfirm = False
    InTrade = False
    
    Exit Sub
errHandler:
    HandleError "InitBattle", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub CloseBattle()
Dim X As Long
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For X = ButtonEnum.BattleFight To ButtonEnum.BattleRun
        Buttons(X).Visible = False
    Next X
    For X = ButtonEnum.bBattleScrollUp To ButtonEnum.bBattleScrollDown
        Buttons(X).Visible = False
    Next X
    For i = 1 To ChatTextBufferSize
        BattleTextBuffer(i).Text = vbNullString
        BattleTextBuffer(i).color = 0
    Next
    BattleScroll = MaxBattleLine
    totalBattleLines = 0
    
    WindowVisible(WindowType.Main_Inventory) = False
    For X = ButtonEnum.InvScrollUp To ButtonEnum.InvScrollDown
        Buttons(X).Visible = False
    Next X

    Exit Sub
errHandler:
    HandleError "CloseBattle", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub FindTarget()
Dim i As Long, X As Long, Y As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If InStorage Or InShop > 0 Or InTrade Or InTradeConfirm Then Exit Sub

    For i = 1 To HighPlayerIndex
        If IsPlaying(i) And GetPlayerMap(MyIndex) = GetPlayerMap(i) Then
            If GetPlayerX(i) = CurX And GetPlayerY(i) = CurY Then
                If i <> MyIndex Then
                    PlayerTarget i
                    Exit Sub
                End If
            End If
        End If
    Next
    
    Exit Sub
errHandler:
    HandleError "FindTarget", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ExitBattle()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Player(MyIndex).InBattle = 0
    CurPoke = 0
    ClearEnemyPokemon
    ForceSwitch = False
    
    If Not Trim$(Map.Music) = "None." Then
        PlayMusic Trim$(Map.Music)
    Else
        StopMusic
    End If
    
    CloseBattle
    
    Exit Sub
errHandler:
    HandleError "ExitBattle", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub InitEvolve(ByVal PokeSlot As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If PokeSlot <= 0 Or PokeSlot > MAX_POKEMON Then Exit Sub
    
    With Player(MyIndex).Pokemon(PokeSlot)
        If .Num <= 0 Then Exit Sub
        If Pokemon(.Num).EvolveNum > 0 Then
            EvolvePoke = PokeSlot
            TmpCurNum = .Num
            TmpEvolveNum = Pokemon(.Num).EvolveNum
            IsEvolve = True
            Buttons(ButtonEnum.EvolveYes).Visible = True
            Buttons(ButtonEnum.EvolveNo).Visible = True
            DrawPokeNum = TmpCurNum
            EvolveAlpha = 255
            EvolvePos = 1
            If Not FileExist(App.Path & MUSIC_PATH & Evolve_Music) Then
                StopMusic
            Else
                PlayMusic Evolve_Music
            End If
            
            TmpInEvolve = 0
        End If
    End With
    
    Exit Sub
errHandler:
    HandleError "InitEvolve", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub CloseEvolve()
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    EvolvePoke = 0
    IsEvolve = False
    Buttons(ButtonEnum.EvolveYes).Visible = False
    Buttons(ButtonEnum.EvolveNo).Visible = False
    
    If Not Trim$(Map.Music) = "None." Then
        PlayMusic Trim$(Map.Music)
    Else
        StopMusic
    End If
    
    TmpEvolveNum = 0
    TmpCurNum = 0
    DrawPokeNum = 0
    EvolveAlpha = 255
    EvolvePos = 1
    
    Exit Sub
errHandler:
    HandleError "CloseEvolve", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function CountPlayerPokemon(ByVal Index As Long) As Long
Dim i As Long, X As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    X = 0
    For i = 1 To MAX_POKEMON
        If Player(Index).Pokemon(i).Num > 0 Then
            X = X + 1
        End If
    Next
    CountPlayerPokemon = X

    Exit Function
errHandler:
    HandleError "CountPlayerPokemon", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function CountPlayerPokemonMove(ByVal Index As Long, ByVal PokeSlot As Long) As Long
Dim i As Long, X As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    X = 0
    For i = 1 To MAX_POKEMON_MOVES
        If Player(Index).Pokemon(PokeSlot).Moves(i).Num > 0 Then
            X = X + 1
        End If
    Next
    CountPlayerPokemonMove = X

    Exit Function
errHandler:
    HandleError "CountPlayerPokemonMove", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function IsStoragePoke(ByVal X As Single, ByVal Y As Single) As Long
Dim tempRec As RECT
Dim i As Long
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    IsStoragePoke = 0
    If GetMaxStoredPokemon > 0 Then
        If GetMaxStoredPokemon >= 12 Then
            For i = StartStorage To StartStorage + 11
                With Player(MyIndex).StoredPokemon(i)
                    If .Num > 0 And .Num <= MAX_STORAGE_POKEMON Then
                        With tempRec
                            .top = GuiStorageY + 22 + (104 * ((i - StartStorage) \ 4)) + 30
                            .bottom = .top + 70
                            .Left = GuiStorageX + 22 + (104 * ((i - StartStorage) Mod 4)) + 30
                            .Right = .Left + 70
                        End With
                        
                        If X >= tempRec.Left And X <= tempRec.Right Then
                            If Y >= tempRec.top And Y <= tempRec.bottom Then
                                IsStoragePoke = i
                                Exit Function
                            End If
                        End If
                    End If
                End With
            Next
        Else
            For i = 1 To GetMaxStoredPokemon
                With Player(MyIndex).StoredPokemon(i)
                    If .Num > 0 And .Num <= MAX_STORAGE_POKEMON Then
                        With tempRec
                            .top = GuiStorageY + 22 + (104 * ((i - 1) \ 4)) + 30
                            .bottom = .top + 70
                            .Left = GuiStorageX + 22 + (104 * ((i - 1) Mod 4)) + 30
                            .Right = .Left + 70
                        End With
                        
                        If X >= tempRec.Left And X <= tempRec.Right Then
                            If Y >= tempRec.top And Y <= tempRec.bottom Then
                                IsStoragePoke = i
                                Exit Function
                            End If
                        End If
                    End If
                End With
            Next
        End If
    End If

    Exit Function
errHandler:
    HandleError "IsStoragePoke", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function GetInvItemNum(ByVal InvSlot As Long, ByVal InvType As Byte) As Long
    If InvSlot > 0 And InvSlot <= MAX_PLAYER_ITEM Then
        If InvType > 0 And InvType <= ItemType.Item_Count - 1 Then
            GetInvItemNum = Player(MyIndex).Item(InvSlot, InvType).Num
        End If
    End If
End Function
