Attribute VB_Name = "modLoop"
Option Explicit

Public Declare Function GetTickCount Lib "Kernel32" () As Long

Public Sub AppLoop()
Dim Tick As Long
Dim Tmr25 As Long
Dim WalkTmr As Long
Dim ChtTmr As Long
Dim Tmr10 As Long
Dim Tmr500 As Long
Dim Tmr150 As Long
Dim Tmr1000 As Long
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Do While AppOpen
        Tick = GetTickCount

        If InGame Then
            If Tmr25 < Tick Then
                Call CheckKeys
                If GetForegroundWindow() = frmMain.hWnd Then
                    Call CheckInputKeys
                End If
                
                If CanMoveNow Then
                    Call CheckMovement
                End If
                
                If TargetSwitch = 1 Then
                    TargetAnim = TargetAnim + 1
                    If TargetAnim >= 5 Then
                        TargetAnim = 5
                        TargetSwitch = 0
                    End If
                ElseIf TargetSwitch = 0 Then
                    TargetAnim = TargetAnim - 1
                    If TargetAnim <= 0 Then
                        TargetAnim = 0
                        TargetSwitch = 1
                    End If
                End If
                
                Tmr25 = Tick + 25
            End If
            
            If Tmr1000 < Tick Then
                If ChangeTitleBar > 0 Then
                    ChangeTitleBar = ChangeTitleBar - 1
                    If ChangeTitleBar <= 0 Then ChangeTitleBar = 0
                End If
                
                Tmr1000 = Tick + 1000
            End If
            
            If Tmr10 < Tick Then
                If ShowTitleBar = YES Then
                    TitleBarAlpha = TitleBarAlpha + 5
                    If TitleBarAlpha >= 255 Then
                        ShowTitleBar = NO
                        ChangeTitleBar = 2
                    End If
                ElseIf ShowTitleBar = NO And ChangeTitleBar = 0 Then
                    TitleBarAlpha = TitleBarAlpha - 5
                    If TitleBarAlpha <= 0 Then TitleBarAlpha = 0
                End If
                
                If Fade Then
                    If FadeType = 1 Then
                        FadeAlpha = FadeAlpha + 5
                        If FadeAlpha >= 255 Then
                            FadeAlpha = 255
                            FadeType = 0
                        End If
                    ElseIf FadeType = 0 Then
                        FadeAlpha = FadeAlpha - 5
                        If FadeAlpha <= 0 Then
                            FadeAlpha = 0
                            FadeType = 0
                            Fade = False
                            CanUseCmd = True
                        End If
                    End If
                End If
                
                If EnemyPos > 0 Then
                    EnemyPos = EnemyPos - 7
                    If EnemyPos <= 0 Then
                        EnemyPos = 0
                    End If
                End If
                
                If isSwitch Then
                    If SwitchPos = YES Then
                        SwitchPokeX = SwitchPokeX - 3
                        PokeAlpha = PokeAlpha - 5
                        If PokeAlpha <= 10 Then
                            PokeAlpha = 0
                            SwitchPos = NO
                            CurPoke = TmpSwitch
                            TmpSwitch = 0
                        End If
                    ElseIf SwitchPos = NO Then
                        SwitchPokeX = SwitchPokeX + 3
                        If SwitchPokeX >= 0 Then SwitchPokeX = 0
                        PokeAlpha = PokeAlpha + 5
                        If PokeAlpha >= 255 Then
                            PokeAlpha = 255
                            isSwitch = False
                            If Not IsSwitchForce Then
                                SendSwitchComplete
                            Else
                                CanUseCmd = True
                                ForceSwitch = False
                            End If
                        End If
                    End If
                End If
                
                If IsCapture Then
                    Capture = Capture - 5
                    If Capture <= 0 Then
                        Capture = 0
                        IsCapture = False
                    End If
                End If
                
                If IsEvolve Then
                    If EvolvePos = 1 Then
                        EvolveAlpha = EvolveAlpha - 9
                        If EvolveAlpha <= 0 Then
                            EvolveAlpha = 0
                            EvolvePos = 0
                            If DrawPokeNum = TmpCurNum Then
                                DrawPokeNum = TmpEvolveNum
                            Else
                                DrawPokeNum = TmpCurNum
                            End If
                        End If
                    ElseIf EvolvePos = 0 Then
                        EvolveAlpha = EvolveAlpha + 9
                        If EvolveAlpha >= 255 Then
                            EvolveAlpha = 255
                            EvolvePos = 1
                        End If
                    End If
                End If
                     
                Tmr10 = Tick + 10
            End If
            
            If WalkTmr < Tick Then
                For i = 1 To HighPlayerIndex
                    If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                        Call ProcessMovement(i)
                    End If
                Next i
                For i = 1 To MAX_MAP_NPC
                    If MapNpc(i).Num > 0 Then
                        ProcessNpcMovement i
                    End If
                Next i
                
                UpdateBarWidth
                
                WalkTmr = Tick + 35
            End If
            
            If ChtTmr < Tick Then
                If ChatScrollUp Then ScrollChatBox 0
                If ChatScrollDown Then ScrollChatBox 1
                If BattleScrollUp Then ScrollBattleBox 0
                If BattleScrollDown Then ScrollBattleBox 1
                If IsInvScrollUp Then ScrollInv 0
                If IsInvScrollDown Then ScrollInv 1
                If InShopScrollUp Then ScrollShop 0
                If InShopScrollDown Then ScrollShop 1
                
                ChtTmr = Tick + 50
            End If
            
            If Tmr150 < Tick Then
                If PokeIconAnim = 0 Then
                    PokeIconAnim = 1
                Else
                    PokeIconAnim = 0
                End If
                
                Tmr150 = Tick + 150
            End If
            If Tmr500 < Tick Then
                If ChatLine = "|" Then
                    ChatLine = vbNullString
                Else
                    ChatLine = "|"
                End If
                Tmr500 = Tick + 500
            End If
            
            InGame = IsConnected
        End If
        
        If InMenu Then
            If Tmr150 < Tick Then
                If GenAnim = 0 Then
                    GenAnim = 2
                Else
                    GenAnim = 0
                End If
                
                Tmr150 = Tick + 150
            End If
            If Tmr500 < Tick Then
                If ChatLine = "|" Then
                    ChatLine = vbNullString
                Else
                    ChatLine = "|"
                End If
                Tmr500 = Tick + 500
            End If
        End If
        
        If Not InGame And Not InMenu Then CloseApp
        
        If frmMain.WindowState <> vbMinimized Then
            Render_Graphics
        End If
        DoEvents
    Loop
    
    Exit Sub
errHandler:
    HandleError "AppLoop", "modLoop", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateBarWidth()
Dim BarWidth As Long
Dim EnemyBarWidth As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Player(MyIndex).InBattle > 0 Then
        BarWidth = (Player(MyIndex).Pokemon(CurPoke).CurHP / 100) / (Player(MyIndex).Pokemon(CurPoke).Stat(Stats.HP) / 100) * 100
        If CurBarWidth < BarWidth Then
            CurBarWidth = CurBarWidth + 3
            UpdatingVital = True
            CanUseCmd = False
            If CurBarWidth >= BarWidth Then
                CurBarWidth = BarWidth
                UpdatingVital = False
                If Not ExitBattleTmr Then
                    CanUseCmd = True
                End If
            End If
        ElseIf CurBarWidth > BarWidth Then
            CurBarWidth = CurBarWidth - 3
            UpdatingVital = True
            CanUseCmd = False
            If CurBarWidth <= BarWidth Then
                CurBarWidth = BarWidth
                UpdatingVital = False
                If Not ExitBattleTmr Then
                    CanUseCmd = True
                End If
            End If
        End If
        
        EnemyBarWidth = (EnemyPokemon.CurHP / 100) / (EnemyPokemon.Stat(Stats.HP) / 100) * 100
        If CurEnemyBarWidth < EnemyBarWidth Then
            CurEnemyBarWidth = CurEnemyBarWidth + 3
            UpdatingVital = True
            CanUseCmd = False
            If CurEnemyBarWidth >= EnemyBarWidth Then
                CurEnemyBarWidth = EnemyBarWidth
                UpdatingVital = False
                If Not ExitBattleTmr Then
                    CanUseCmd = True
                End If
            End If
        ElseIf CurEnemyBarWidth > EnemyBarWidth Then
            CurEnemyBarWidth = CurEnemyBarWidth - 3
            UpdatingVital = True
            CanUseCmd = False
            If CurEnemyBarWidth <= EnemyBarWidth Then
                CurEnemyBarWidth = EnemyBarWidth
                UpdatingVital = False
                If Not ExitBattleTmr Then
                    CanUseCmd = True
                End If
            End If
        End If
    End If
    
    If ExitBattleTmr Then
        If Not UpdatingVital Then
            CanExit = True
            If DidWin > 0 Then
                If Not CurMusic = Victory_Wild_Music Then
                    If Not FileExist(App.Path & MUSIC_PATH & Victory_Wild_Music) Then
                        StopMusic
                    Else
                        PlayMusic Victory_Wild_Music
                    End If
                End If
            End If
        End If
    End If
    
    Exit Sub
errHandler:
    HandleError "UpdateBarWidth", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub ScrollInv(ByVal Direction As Byte)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If CurInvType > 0 Then
        If Direction = 0 Then
            StartInv = StartInv - 1
            If StartInv <= 1 Then
                StartInv = 1
            End If
        Else
            StartInv = StartInv + 1
            If StartInv >= GetMaxInv - 3 Then
                StartInv = GetMaxInv - 3
            End If
        End If
    End If
    
    Exit Sub
errHandler:
    HandleError "ScrollInv", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function GetMaxInv() As Long
Dim X As Long, Y As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Y = 0
    If CurInvType >= 0 Then
        For X = 1 To MAX_PLAYER_ITEM
            If Player(MyIndex).Item(X, CurInvType).Num > 0 Then
                Y = Y + 1
            End If
        Next
    End If
    GetMaxInv = Y
    
    Exit Function
errHandler:
    HandleError "GetMaxInv", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function GetMaxStoredPokemon() As Long
Dim X As Long, Y As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Y = 0
    For X = 1 To MAX_STORAGE_POKEMON
        If Player(MyIndex).StoredPokemon(X).Num > 0 Then
            Y = Y + 1
        End If
    Next
    GetMaxStoredPokemon = Y
    
    Exit Function
errHandler:
    HandleError "GetMaxStoredPokemon", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub ChangeStorageTab(ByVal isNext As Boolean)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If isNext Then
        If StartStorage + 12 <= GetMaxStoredPokemon Then
            StartStorage = StartStorage + 12
        End If
    Else
        If StartStorage > 12 Then
            StartStorage = StartStorage - 12
        End If
    End If
    
    Exit Sub
errHandler:
    HandleError "ChangeStorageTab", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub ScrollShop(ByVal Direction As Byte)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Direction = 0 Then
        ShopStart = ShopStart - 1
        If ShopStart <= 1 Then
            ShopStart = 1
        End If
    Else
        ShopStart = ShopStart + 1
        If ShopStart + 8 >= MAX_SHOP_ITEMS Then
            ShopStart = MAX_SHOP_ITEMS - 8
        End If
    End If
    
    Exit Sub
errHandler:
    HandleError "ScrollShop", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub
