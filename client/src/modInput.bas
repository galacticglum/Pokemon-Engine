Attribute VB_Name = "modInput"
Option Explicit

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public HoldUp As Boolean
Public HoldDown As Boolean
Public HoldLeft As Boolean
Public HoldRight As Boolean
Public ShiftDown As Boolean
Public CtrlDown As Boolean

Public Sub CheckKeys()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If GetKeyState(vbKeyUp) >= 0 Then HoldUp = False
    If GetKeyState(vbKeyDown) >= 0 Then HoldDown = False
    If GetKeyState(vbKeyLeft) >= 0 Then HoldLeft = False
    If GetKeyState(vbKeyRight) >= 0 Then HoldRight = False
    If GetKeyState(vbKeyShift) >= 0 Then ShiftDown = False
    If GetKeyState(vbKeyControl) >= 0 Then CtrlDown = False
    
    Exit Sub
errHandler:
    HandleError "CheckKeys", "modInput", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckInputKeys()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If GetKeyState(vbKeyControl) < 0 Then
        CtrlDown = True
    Else
        CtrlDown = False
    End If
    
    If GetKeyState(vbKeyShift) < 0 Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If
    
    If GetKeyState(vbKeyUp) < 0 Then
        HoldUp = True
        HoldDown = False
        HoldLeft = False
        HoldRight = False
        Exit Sub
    Else
        HoldUp = False
    End If
    
    If GetKeyState(vbKeyDown) < 0 Then
        HoldUp = False
        HoldDown = True
        HoldLeft = False
        HoldRight = False
        Exit Sub
    Else
        HoldDown = False
    End If
    
    If GetKeyState(vbKeyLeft) < 0 Then
        HoldUp = False
        HoldDown = False
        HoldLeft = True
        HoldRight = False
        Exit Sub
    Else
        HoldLeft = False
    End If
    
    If GetKeyState(vbKeyRight) < 0 Then
        HoldUp = False
        HoldDown = False
        HoldLeft = False
        HoldRight = True
        Exit Sub
    Else
        HoldRight = False
    End If
    
    Exit Sub
errHandler:
    HandleError "CheckInputKeys", "modInput", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If InMenu Then
        If Y > GuiCharCreateY + 72 And Y <= GuiCharCreateY + 72 + GetTextureHeight(Tex_Button_N(4)) Then
            If X > GuiCharCreateX + 23 And X <= GuiCharCreateX + 23 + GetTextureWidth(Tex_Button_N(4)) Then
                SelGender = GENDER_MALE
            ElseIf X > GuiCharCreateX + 77 And X <= GuiCharCreateX + 77 + GetTextureWidth(Tex_Button_N(4)) Then
                SelGender = GENDER_FEMALE
            End If
        End If
    End If
    
    If InGame Then
        If Editor = EDITOR_MAP Then
            MapEditorMouseDown Button, X, Y
        Else
            If Button = vbLeftButton Then
                FindTarget
            ElseIf Button = vbRightButton Then
                If ShiftDown Then
                    If GetPlayerAccess(MyIndex) > 0 Then
                        AdminWarp CurX, CurY
                    End If
                End If
            End If
        End If
    End If
    
    For i = 1 To ButtonEnum.MaxButton
        With Buttons(i)
            If .Visible Then
                If X >= .X And X <= .X + GetTextureWidth(Tex_Button_N(.Pic)) Then
                    If Y >= .Y And Y <= .Y + GetTextureHeight(Tex_Button_N(.Pic)) Then
                        If .bState = 0 Then
                            Select Case i
                                Case ButtonEnum.bChatScrollUp
                                    If InGame Then
                                        ChatScrollUp = True
                                    End If
                                Case ButtonEnum.bChatScrollDown
                                    If InGame Then
                                        ChatScrollDown = True
                                    End If
                                Case ButtonEnum.bBattleScrollUp
                                    If InGame And Player(MyIndex).InBattle Then
                                        BattleScrollUp = True
                                    End If
                                Case ButtonEnum.bBattleScrollDown
                                    If InGame And Player(MyIndex).InBattle Then
                                        BattleScrollDown = True
                                    End If
                                Case ButtonEnum.InvScrollUp
                                    If InGame Then
                                        If WindowVisible(WindowType.Main_Inventory) Then
                                            IsInvScrollUp = True
                                            UseItemNum = 0
                                            ShowUseItem = False
                                        End If
                                    End If
                                Case ButtonEnum.InvScrollDown
                                    If InGame Then
                                        If WindowVisible(WindowType.Main_Inventory) Then
                                            IsInvScrollDown = True
                                            UseItemNum = 0
                                            ShowUseItem = False
                                        End If
                                    End If
                                Case ButtonEnum.ShopScrollUp
                                    If InGame Then
                                        If InShop > 0 Then
                                            InShopScrollUp = True
                                        End If
                                    End If
                                Case ButtonEnum.ShopScrollDown
                                    If InGame Then
                                        If InShop > 0 Then
                                            InShopScrollDown = True
                                        End If
                                    End If
                            End Select
                            If Not LastButtonClick = i Then
                                .bState = 1
                                LastButtonClick = i
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Next
    
    Exit Sub
errHandler:
    HandleError "HandleMouseDown", "modInput", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleClick()
Dim i As Long, X As Long, Y As Long
Dim x2 As Long, y2 As Long, mText As String
Dim DidConfirm As Boolean

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If InGame Then
        If Not ShowInput Then
            If Player(MyIndex).InBattle > 0 Then
                If ShowMoves And Not IsLearnMove Then
                    X = GuiBattleX + 542
                    For i = 1 To MAX_POKEMON_MOVES
                        Y = GuiBattleY + 342 + ((i - 1) * 22)
                        With Player(MyIndex).Pokemon(CurPoke)
                            If .Moves(i).Num > 0 Then
                                If GlobalX >= X And GlobalX <= X + 107 And GlobalY >= Y And GlobalY <= Y + 16 Then
                                    If .Moves(i).PP <= 0 Then
                                        AddBattleLog "There's no PP left for this move!", White
                                        ShowMoves = False
                                        Exit For
                                    Else
                                        SendBattleCommand BATTLE_COMMAND_FIGHT, i
                                        CanUseCmd = False
                                        ShowMoves = False
                                        Exit For
                                    End If
                                End If
                            End If
                        End With
                    Next i
                End If
                If ShowPokemonSwitch And Not IsLearnMove Then
                    X = GuiBattleX + 665
                    For i = 1 To MAX_POKEMON
                        Y = GuiBattleY + 302 + ((i - 1) * 22)
                        With Player(MyIndex)
                            If GlobalX >= X And GlobalX <= X + 107 And GlobalY >= Y And GlobalY <= Y + 16 Then
                                If .Pokemon(i).Num > 0 Then
                                    If i = CurPoke Then
                                        AddBattleLog Trim$(Pokemon(.Pokemon(CurPoke).Num).Name) & " is already in battle!", White
                                        If Not ForceSwitch Then ShowPokemonSwitch = False
                                        Exit For
                                    ElseIf .Pokemon(i).CurHP <= 0 Then
                                        AddBattleLog Trim$(Pokemon(.Pokemon(i).Num).Name) & " has no energy left to battle!", White
                                        If Not ForceSwitch Then ShowPokemonSwitch = False
                                        Exit For
                                    Else
                                        If ForceSwitch Then
                                            SendBattleCommand BATTLE_COMMAND_SWITCH, i, YES
                                        Else
                                            SendBattleCommand BATTLE_COMMAND_SWITCH, i, NO
                                        End If
                                        CanUseCmd = False
                                        ShowPokemonSwitch = False
                                        Exit For
                                    End If
                                End If
                            End If
                        End With
                    Next i
                End If
            End If
            If IsLearnMove Then
                X = GuiLearnMoveX + 100
                For i = 1 To MAX_POKEMON_MOVES
                    Y = GuiLearnMoveY + 86 + ((i - 1) * 20)
                    With Player(MyIndex).Pokemon(LearnPokeNum)
                        If .Moves(i).Num > 0 Then
                            If GlobalX >= X And GlobalX <= X + 107 And GlobalY >= Y + 2 And GlobalY <= Y + 18 Then
                                If Not SelectedMove = i Then SelectedMove = i
                            End If
                        End If
                    End With
                Next
                x2 = GuiLearnMoveX + 100 + ((107 / 2) - (GetWidth(CurFont, "Replace") / 2)): y2 = GuiLearnMoveY + 175
                If GlobalX >= x2 And GlobalX <= x2 + GetWidth(CurFont, "Replace") And GlobalY >= y2 And GlobalY <= y2 + 16 Then
                    If SelectedMove > 0 Then
                        SendReplaceMove LearnPokeNum, SelectedMove, LearnMoveNum
                        SelectedMove = 0
                        LearnPokeNum = 0
                        LearnMoveNum = 0
                        IsLearnMove = False
                    End If
                End If
                x2 = GuiLearnMoveX + 100 + ((107 / 2) - (GetWidth(CurFont, "Cancel") / 2)): y2 = GuiLearnMoveY + 195
                If GlobalX >= x2 And GlobalX <= x2 + GetWidth(CurFont, "Cancel") And GlobalY >= y2 And GlobalY <= y2 + 16 Then
                    SelectedMove = 0
                    LearnPokeNum = 0
                    LearnMoveNum = 0
                    IsLearnMove = False
                End If
            End If
            
            If ExitBattleTmr And CanExit Then
                CanExit = False
                ExitBattleTmr = False
                ExitBattle
                If TmpInEvolve > 0 Then InitEvolve TmpInEvolve
            End If
            
            If WindowVisible(WindowType.Main_Option) Then
                If GlobalY >= GuiOptionY + 33 And GlobalY <= GuiOptionY + 44 Then
                    If GlobalX >= GuiOptionX + 110 And GlobalX <= GuiOptionX + 121 Then
                        Options.Music = YES
                        If Not Trim$(Map.Music) = "None." Then
                            PlayMusic Trim$(Map.Music)
                        Else
                            StopMusic
                        End If
                        SaveOption
                    ElseIf GlobalX >= GuiOptionX + 177 And GlobalX <= GuiOptionX + 188 Then
                        Options.Music = NO
                        StopMusic
                        SaveOption
                    End If
                ElseIf GlobalY >= GuiOptionY + 63 And GlobalY <= GuiOptionY + 74 Then
                    If GlobalX >= GuiOptionX + 110 And GlobalX <= GuiOptionX + 121 Then
                        Options.Sound = YES
                        SaveOption
                    ElseIf GlobalX >= GuiOptionX + 177 And GlobalX <= GuiOptionX + 188 Then
                        Options.Sound = NO
                        StopSound
                        SaveOption
                    End If
                End If
            End If
            
            If MyTarget > 0 Then
                X = GuiTargetMenuX + 83 + (101 / 2) - (GetWidth(CurFont, "Trade") / 2)
                Y = GuiTargetMenuY + 49
                If GlobalX >= X And GlobalX <= X + GetWidth(CurFont, "Trade") And GlobalY >= Y And GlobalY <= Y + 16 Then
                    If Not InStorage Then
                        SendInitTrade MyTarget
                        InTradeIndex = MyTarget
                        MyTarget = 0
                    End If
                End If
                Y = GuiTargetMenuY + 69
                If GlobalX >= X And GlobalX <= X + GetWidth(CurFont, "Battle") And GlobalY >= Y And GlobalY <= Y + 16 Then
                    If Not InStorage Then
                        SendBattleRequest MyTarget
                        MyTarget = 0
                    End If
                End If
            End If
            
            If WindowVisible(WindowType.Main_Inventory) Then
                If Player(MyIndex).InBattle > 0 Then
                    If GlobalX >= GuiInventoryX + 25 And GlobalX <= GuiInventoryX + 25 + GetWidth(CurFont, "Close") And GlobalY >= ((GuiInventoryY + GetTextureHeight(Tex_Gui(GuiInventory))) - 35) And GlobalY <= ((GuiInventoryY + GetTextureHeight(Tex_Gui(GuiInventory))) - 35 + 16) Then
                        CloseInventory
                    End If
                End If
                x2 = GuiInventoryX + 45
                If GetMaxInv > 0 Then
                    If GetMaxInv >= 4 Then
                        For X = StartInv To StartInv + 3
                            y2 = GuiInventoryY + 39 + ((X - StartInv) * 39)
                            With Player(MyIndex).Item(X, CurInvType)
                                If .Num > 0 Then
                                    If GlobalX >= x2 And GlobalX <= x2 + 32 And GlobalY >= y2 And GlobalY <= y2 + 32 Then
                                        If UseItemNum = X Then
                                            UseItemNum = 0
                                            ShowUseItem = False
                                        Else
                                            UseItemNum = X
                                            ShowUseItem = True
                                        End If
                                        DidConfirm = True
                                        Exit For
                                    End If
                                End If
                            End With
                        Next
                        For X = StartInv To StartInv + 3
                            y2 = GuiInventoryY + 39 + ((X - StartInv) * 39)
                            If ShowUseItem And UseItemNum = X Then
                                If GlobalX >= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Summary") / 2)) And GlobalX <= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Summary") / 2)) + GetWidth(CurFont, "Summary") And GlobalY >= y2 + 28 And GlobalY <= y2 + 28 + 16 Then
                                    ' View Summary
                                    UseItemNum = 0
                                    ShowUseItem = False
                                End If
                                If GlobalX >= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Use") / 2)) And GlobalX <= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Use") / 2)) + GetWidth(CurFont, "Use") And GlobalY >= y2 + 48 And GlobalY <= y2 + 48 + 16 Then
                                    If Not InShop > 0 And Not InTrade Then
                                        SendUseItem UseItemNum
                                        UseItemNum = 0
                                        ShowUseItem = False
                                        CloseInventory
                                    End If
                                End If
                                If InShop > 0 Then
                                    If GlobalX >= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Sell") / 2)) And GlobalX <= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Sell") / 2)) + GetWidth(CurFont, "Sell") And GlobalY >= y2 + 68 And GlobalY <= y2 + 68 + 16 Then
                                        InitInput 1, GlobalX, GlobalY, UseItemNum
                                        ShowUseItem = False
                                    End If
                                End If
                                If InTrade Then
                                    If CurInvType <> ItemType.TM_HMs And CurInvType <> ItemType.KeyItems Then
                                        If GlobalX >= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Trade") / 2)) And GlobalX <= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Trade") / 2)) + GetWidth(CurFont, "Trade") And GlobalY >= y2 + 68 And GlobalY <= y2 + 68 + 16 Then
                                            InitInput 3, GlobalX, GlobalY, UseItemNum
                                            ShowUseItem = False
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        If Not DidConfirm Then
                            UseItemNum = 0
                            ShowUseItem = False
                        End If
                    Else
                        For X = 1 To GetMaxInv
                            y2 = GuiInventoryY + 39 + ((X - 1) * 39)
                            With Player(MyIndex).Item(X, CurInvType)
                                If .Num > 0 Then
                                    If GlobalX >= x2 And GlobalX <= x2 + 32 And GlobalY >= y2 And GlobalY <= y2 + 32 Then
                                        If Not ShowUseItem Then
                                            If UseItemNum = X Then
                                                UseItemNum = 0
                                                ShowUseItem = False
                                            Else
                                                UseItemNum = X
                                                ShowUseItem = True
                                            End If
                                        End If
                                        DidConfirm = True
                                        Exit For
                                    End If
                                End If
                            End With
                        Next
                        For X = 1 To GetMaxInv
                            y2 = GuiInventoryY + 39 + ((X - 1) * 39)
                            If ShowUseItem And UseItemNum = X Then
                                If GlobalX >= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Summary") / 2)) And GlobalX <= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Summary") / 2)) + GetWidth(CurFont, "Summary") And GlobalY >= y2 + 28 And GlobalY <= y2 + 28 + 16 Then
                                    ' View Summary
                                    UseItemNum = 0
                                    ShowUseItem = False
                                End If
                                If GlobalX >= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Use") / 2)) And GlobalX <= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Use") / 2)) + GetWidth(CurFont, "Use") And GlobalY >= y2 + 48 And GlobalY <= y2 + 48 + 16 Then
                                    If Not InShop > 0 And Not InTrade Then
                                        SendUseItem UseItemNum
                                        UseItemNum = 0
                                        ShowUseItem = False
                                        CloseInventory
                                    End If
                                End If
                                If InShop > 0 Then
                                    If GlobalX >= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Sell") / 2)) And GlobalX <= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Sell") / 2)) + GetWidth(CurFont, "Sell") And GlobalY >= y2 + 68 And GlobalY <= y2 + 68 + 16 Then
                                        InitInput 1, GlobalX, GlobalY, UseItemNum
                                        ShowUseItem = False
                                    End If
                                End If
                                If InTrade Then
                                    If CurInvType <> ItemType.TM_HMs And CurInvType <> ItemType.KeyItems Then
                                        If GlobalX >= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Trade") / 2)) And GlobalX <= x2 + 32 + ((107 / 2) - (GetWidth(CurFont, "Trade") / 2)) + GetWidth(CurFont, "Trade") And GlobalY >= y2 + 68 And GlobalY <= y2 + 68 + 16 Then
                                            InitInput 3, GlobalX, GlobalY, UseItemNum
                                            ShowUseItem = False
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        If Not DidConfirm Then
                            UseItemNum = 0
                            ShowUseItem = False
                        End If
                    End If
                End If
            End If
            
            If ShowSelect Then
                For i = 1 To MaxSelection + 1
                    X = (ScreenWidth / 2) - (GetTextureWidth(Tex_Gui(GuiSelection)) / 2)
                    Y = (ScreenHeight / 2) - ((21 * MaxSelection) / 2) + ((i - 1) * 20)
                    If i > MaxSelection Then
                        mText = "Close"
                        If GlobalX >= X + ((107 / 2) - (GetWidth(CurFont, mText) / 2)) And GlobalX <= X + ((107 / 2) - (GetWidth(CurFont, mText) / 2)) + GetWidth(CurFont, mText) And GlobalY >= Y And GlobalY <= Y + 16 Then
                            OutputData1 = 0: OutputData2 = 0: OutputData3 = 0: OutputData4 = 0
                            InputData1 = 0: InputData2 = 0: InputData3 = 0: InputData4 = 0
                            ShowSelect = False
                        End If
                    Else
                        If InputData1 = SELECT_POKEMON Then
                            mText = Trim$(Pokemon(Player(MyIndex).Pokemon(i).Num).Name)
                        ElseIf InputData1 = SELECT_MOVE Then
                            mText = Trim$(Moves(Player(MyIndex).Pokemon(InputData2).Moves(i).Num).Name)
                        End If
                        If GlobalX >= X + ((107 / 2) - (GetWidth(CurFont, mText) / 2)) And GlobalX <= X + ((107 / 2) - (GetWidth(CurFont, mText) / 2)) + GetWidth(CurFont, mText) And GlobalY >= Y And GlobalY <= Y + 16 Then
                            OutputData1 = InputData1
                            OutputData2 = InputData2
                            OutputData3 = i
                            OutputData4 = InputData3
                            SendInitSelect
                            OutputData1 = 0: OutputData2 = 0: OutputData3 = 0: OutputData4 = 0
                            InputData1 = 0: InputData2 = 0: InputData3 = 0: InputData4 = 0
                            ShowSelect = False
                        End If
                    End If
                Next
            End If
            
            If InStorage Then
                If ShowStorageSelect Then
                    X = GuiStorageX + 28
                    For i = 1 To MAX_POKEMON
                        Y = GuiStorageY + 208 + ((i - 1) * 20)
                        RenderTexture Tex_Gui(GuiSelection), X, Y, 0, 0, GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection)), GetTextureWidth(Tex_Gui(GuiSelection)), GetTextureHeight(Tex_Gui(GuiSelection))
                        With Player(MyIndex)
                            If .Pokemon(i).Num > 0 Then
                                mText = Trim$(Pokemon(.Pokemon(i).Num).Name)
                                If GlobalX >= X And GlobalX <= X + 107 And GlobalY >= Y + 2 And GlobalY <= Y + 18 Then
                                    SendDepositPokemon i
                                    ShowStorageSelect = False
                                    Exit For
                                End If
                            End If
                        End With
                    Next
                Else
                    If SelStoragePoke > 0 Then
                        If GlobalX >= SelStorageX And GlobalX <= SelStorageX + 107 And GlobalY >= SelStorageY And GlobalY <= SelStorageY + 60 Then
                            X = SelStorageX + ((107 / 2) - (GetWidth(CurFont, "Summary") / 2)): Y = SelStorageY + 2
                            If GlobalX >= X And GlobalX <= X + GetWidth(CurFont, "Summary") And GlobalY >= Y And GlobalY <= Y + 16 Then
                                ' Summary
                            End If
                            X = SelStorageX + ((107 / 2) - (GetWidth(CurFont, "Withdraw") / 2)): Y = SelStorageY + 22
                            If GlobalX >= X And GlobalX <= X + GetWidth(CurFont, "Withdraw") And GlobalY >= Y And GlobalY <= Y + 16 Then
                                SendWithdrawPokemon SelStoragePoke
                                SelStoragePoke = 0
                            End If
                            X = SelStorageX + ((107 / 2) - (GetWidth(CurFont, "Release") / 2)): Y = SelStorageY + 42
                            If GlobalX >= X And GlobalX <= X + GetWidth(CurFont, "Release") And GlobalY >= Y And GlobalY <= Y + 16 Then
                                ' Release
                            End If
                        Else
                            SelStoragePoke = IsStoragePoke(GlobalX, GlobalY)
                        End If
                    Else
                        SelStoragePoke = IsStoragePoke(GlobalX, GlobalY)
                    End If
                End If
            End If
            
            If InShop > 0 Then
                For X = ShopStart To ShopStart + 8
                    If Shop(InShop).sItem(X).Num > 0 Then
                        If GlobalX >= GuiShopX + 20 And GlobalX <= GuiShopX + 20 + 200 Then
                            If GlobalY >= GuiShopY + 40 + ((X - ShopStart) * 20) And GlobalY <= GuiShopY + 40 + ((X - ShopStart) * 20) + 16 Then
                                If ShopSelect <> X Then
                                    ShopSelect = X
                                Else
                                    ShopSelect = 0
                                End If
                            End If
                        End If
                    End If
                Next
        
                X = GuiShopX + 220: Y = GuiShopY + 235
                If GlobalX >= X And GlobalX <= X + GetWidth(CurFont, "Close") And GlobalY >= Y And GlobalY <= Y + 16 Then
                    InShop = 0
                    For Y = ButtonEnum.ShopScrollUp To ButtonEnum.ShopScrollDown
                        Buttons(Y).Visible = True
                    Next
                    CloseInventory
                End If
                X = GuiShopX + 15
                If GlobalX >= X And GlobalX <= X + GetWidth(CurFont, "Buy") And GlobalY >= Y And GlobalY <= Y + 16 Then
                    If ShopSelect > 0 Then
                        InitInput 2, GlobalX, GlobalY, ShopSelect
                    End If
                End If
            End If
            
            If InTrade Then
                X = GuiTradeX + 244
                Y = GuiTradeY + 13
                If GlobalX >= X And GlobalX <= X + 7 And GlobalY >= Y And GlobalY <= Y + 8 Then
                    CloseTrade
                End If
                For i = 1 To MAX_TRADE
                    y2 = GuiTradeY + 40 + ((i - 1) * 20)
                    With MyTrade(i)
                        If .Type > 0 Then
                            If GlobalX >= GuiTradeX + 10 And GlobalX <= GuiTradeX + 242 And GlobalY >= y2 And GlobalY <= y2 + 16 Then
                                If Not ShowTradeSel Then
                                    If SelTrade = i Then
                                        ShowTradeSel = False
                                        SelTrade = 0
                                    Else
                                        ShowTradeSel = True
                                        SelTrade = i
                                    End If
                                End If
                                DidConfirm = True
                                Exit For
                            End If
                        End If
                    End With
                Next
                If ShowTradeSel Then
                    For i = 1 To MAX_TRADE
                        y2 = GuiTradeY + 40 + ((i - 1) * 20)
                        If SelTrade = i Then
                            If GlobalX >= GuiTradeX + 42 + ((107 / 2) - (GetWidth(CurFont, "Remove") / 2)) And GlobalX <= GuiTradeX + 42 + ((107 / 2) - (GetWidth(CurFont, "Remove") / 2)) + GetWidth(CurFont, "Remove") And GlobalY >= y2 + 28 And GlobalY <= y2 + 28 + 16 Then
                                ClearTradeSlot SelTrade
                                ShowTradeSel = False
                                SelTrade = 0
                            End If
                        End If
                    Next
                End If
                If Not DidConfirm Then
                    ShowTradeSel = False
                    SelTrade = 0
                End If
                
                For i = 1 To MAX_POKEMON
                    If Player(MyIndex).Pokemon(i).Num > 0 Then
                        X = GuiTradeX + 15 + ((i - 1) * 41)
                        If GlobalX >= X And GlobalX <= X + 32 And GlobalY >= GuiTradeY + 280 And GlobalY <= GuiTradeY + 280 + 32 Then
                            If i = 1 Then
                                AddText "You cannot trade the pokemon on your first slot!", Red
                            Else
                                AddTrade TRADE_TYPE_POKEMON, i
                            End If
                            Exit For
                        End If
                    End If
                Next i
            End If
            
            If InTradeConfirm Then
                X = GuiTradeConfirmX + 314
                Y = GuiTradeConfirmY + 13
                If GlobalX >= X And GlobalX <= X + 7 And GlobalY >= Y And GlobalY <= Y + 8 Then
                    CloseTrade
                End If
            End If
        End If
        
        If ShowInput Then
            If GlobalY >= RenderValY + 60 And GlobalY <= RenderValY + 60 + GetWidth(CurFont, "Close") Then
                If GlobalX >= RenderValX + 20 And GlobalX <= RenderValX + 20 + GetWidth(CurFont, "Confirm") Then
                    ConfirmInput
                End If
                If GlobalX >= RenderValX + 115 And GlobalX <= RenderValX + 115 + GetWidth(CurFont, "Close") Then
                    CloseInput
                End If
            End If
        End If
    End If
    
    If InMenu Then
        If GlobalX >= GuiLoginX + 99 And GlobalX <= GuiLoginX + 110 Then
            If GlobalY >= GuiLoginY + 89 And GlobalY <= GuiLoginY + 100 Then
                SaveAccount = Not SaveAccount
            End If
        End If

        If WindowVisible(WindowType.Menu_CharCreate) Then
            mText = "Bulbasaur"
            X = GuiCharCreateX + 130 + (86 / 2) - (GetWidth(CurFont, mText) / 2)
            Y = GuiCharCreateY + 85
            If GlobalX >= X And GlobalX <= X + GetWidth(CurFont, mText) And GlobalY >= Y And GlobalY <= Y + 16 Then
                SelStarter = 1
            End If
            mText = "Charmander"
            X = GuiCharCreateX + 130 + (86 / 2) - (GetWidth(CurFont, mText) / 2)
            Y = GuiCharCreateY + 100
            If GlobalX >= X And GlobalX <= X + GetWidth(CurFont, mText) And GlobalY >= Y And GlobalY <= Y + 16 Then
                SelStarter = 4
            End If
            mText = "Squirtle"
            X = GuiCharCreateX + 130 + (86 / 2) - (GetWidth(CurFont, mText) / 2)
            Y = GuiCharCreateY + 115
            If GlobalX >= X And GlobalX <= X + GetWidth(CurFont, mText) And GlobalY >= Y And GlobalY <= Y + 16 Then
                SelStarter = 7
            End If
        End If
    End If
    
    Exit Sub
errHandler:
    HandleError "HandleClick", "modInput", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    GlobalX = X: GlobalY = Y
    If InGame Then
        GlobalX_Map = GlobalX + (TileView.Left * Pic_Size) + Camera.Left
        GlobalY_Map = GlobalY + (TileView.top * Pic_Size) + Camera.top
        
        CurX = TileView.Left + ((X + Camera.Left) \ Pic_Size)
        CurY = TileView.top + ((Y + Camera.top) \ Pic_Size)
        
        If Editor = EDITOR_MAP Then
            MapEditorMouseDown Button, X, Y
        End If
    End If
    
    Exit Sub
errHandler:
    HandleError "HandleMouseMove", "modInput", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If InMenu Then
        
    End If
    If InGame Then
 
    End If
    
    For i = 1 To ButtonEnum.MaxButton
        With Buttons(i)
            If .Visible Then
                If X >= .X And X <= .X + GetTextureWidth(Tex_Button_N(.Pic)) Then
                    If Y >= .Y And Y <= .Y + GetTextureHeight(Tex_Button_N(.Pic)) Then
                        If .bState = 1 Then
                            Select Case i
                                ' Menu ~
                                Case ButtonEnum.LoginAccept
                                    If InMenu Then MenuState MenuStateLogin
                                Case ButtonEnum.Register
                                    If InMenu Then OpenWindow Menu_Register
                                Case ButtonEnum.RegisterAccept
                                    If InMenu Then MenuState MenuStateRegister
                                Case ButtonEnum.Login
                                    If InMenu Then OpenWindow Menu_Login
                                Case ButtonEnum.CharNew1, ButtonEnum.CharNew2, ButtonEnum.CharNew3
                                    If InMenu Then
                                        OpenWindow Menu_CharCreate
                                        CharSelected = i - 4
                                    End If
                                Case ButtonEnum.CharUse1, ButtonEnum.CharUse2, ButtonEnum.CharUse3
                                    If InMenu Then
                                        CharSelected = i - 7
                                        MenuState MenuStateUseChar
                                    End If
                                Case ButtonEnum.CharDel1, ButtonEnum.CharDel2, ButtonEnum.CharDel3
                                    If InMenu Then
                                        ' Temporary Deleting Interface
                                        CharSelected = i - 10
                                        If MsgBox("Are you sure you want to delete this character?", vbYesNo) = vbYes Then
                                            MenuState MenuStateDelChar
                                        End If
                                    End If
                                Case ButtonEnum.CharAccept
                                    If InMenu Then MenuState MenuStateAddChar
                                Case ButtonEnum.CharDecline
                                    If InMenu Then OpenWindow Menu_CharSelect
                                ' Menu End ~
                                ' Main ~
                                Case ButtonEnum.bChatScrollUp
                                    ChatScrollUp = False
                                Case ButtonEnum.bChatScrollDown
                                    ChatScrollDown = False
                                Case ButtonEnum.mPokedex
                                
                                Case ButtonEnum.mCharacter
                                    If InGame Then
                                        If Not InShop > 0 And Not InTrade And Not InTradeConfirm Then
                                            If Not Player(MyIndex).InBattle > 0 Then
                                                WindowVisible(WindowType.Main_Trainer) = Not WindowVisible(WindowType.Main_Trainer)
                                                CloseInventory
                                                WindowVisible(WindowType.Main_Option) = False
                                            End If
                                        End If
                                    End If
                                Case ButtonEnum.mInventory
                                    If InGame Then
                                        If Not InShop > 0 And Not InTrade And Not InTradeConfirm Then
                                            If Player(MyIndex).InBattle > 0 Then
                                                CloseInventory
                                            Else
                                                If WindowVisible(WindowType.Main_Inventory) Then
                                                    CloseInventory
                                                Else
                                                    OpenInventory
                                                End If
                                            End If
                                            WindowVisible(WindowType.Main_Trainer) = False
                                            WindowVisible(WindowType.Main_Option) = False
                                        End If
                                    End If
                                Case ButtonEnum.mOptions
                                    If InGame Then
                                        If Not InShop > 0 And Not InTrade And Not InTradeConfirm Then
                                            If Not Player(MyIndex).InBattle > 0 Then
                                                WindowVisible(WindowType.Main_Option) = Not WindowVisible(WindowType.Main_Option)
                                                CloseInventory
                                                WindowVisible(WindowType.Main_Trainer) = False
                                            End If
                                        End If
                                    End If
                                Case ButtonEnum.BattleFight
                                    If InGame Then
                                        If Not ForceSwitch And Not IsLearnMove Then
                                            If Player(MyIndex).InBattle > 0 And EnemyPos = 0 And CanUseCmd And Not ShowSelect Then
                                                If Not WindowVisible(WindowType.Main_Inventory) Then
                                                    ShowMoves = Not ShowMoves
                                                    ShowPokemonSwitch = False
                                                End If
                                            End If
                                        End If
                                    End If
                                Case ButtonEnum.BattleSwitch
                                    If InGame Then
                                        If Not ForceSwitch And Not IsLearnMove Then
                                            If Player(MyIndex).InBattle > 0 And EnemyPos = 0 And CanUseCmd And Not ShowSelect Then
                                                If Not WindowVisible(WindowType.Main_Inventory) Then
                                                    ShowPokemonSwitch = Not ShowPokemonSwitch
                                                    ShowMoves = False
                                                End If
                                            End If
                                        End If
                                    End If
                                Case ButtonEnum.BattleBag
                                    If InGame Then
                                        If Not ForceSwitch And Not IsLearnMove Then
                                            If Player(MyIndex).InBattle > 0 And EnemyPos = 0 And CanUseCmd And Not ShowSelect Then
                                                If Not WindowVisible(WindowType.Main_Inventory) Then
                                                    ShowMoves = False
                                                    ShowPokemonSwitch = False
                                                    OpenInventory
                                                End If
                                            End If
                                        End If
                                    End If
                                Case ButtonEnum.BattleRun
                                    If InGame Then
                                        If Not IsLearnMove Then
                                            If Player(MyIndex).InBattle > 0 And EnemyPos = 0 And CanUseCmd And Not ShowSelect Then
                                                If Not WindowVisible(WindowType.Main_Inventory) Then
                                                    If Player(MyIndex).InBattle = BATTLE_TRAINER Then
                                                        AddBattleLog "You can't run away from Trainer's Battle, Type: /forfiet - to forfiet the battle!", Red
                                                    Else
                                                        If ForceSwitch Then
                                                            SendBattleCommand BATTLE_COMMAND_RUN, YES
                                                        Else
                                                            SendBattleCommand BATTLE_COMMAND_RUN
                                                        End If
                                                        CanUseCmd = False
                                                        ShowMoves = False
                                                        ShowPokemonSwitch = False
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Case ButtonEnum.bBattleScrollUp
                                    If InGame Then
                                        If Player(MyIndex).InBattle > 0 Then
                                            BattleScrollUp = False
                                        End If
                                    End If
                                Case ButtonEnum.bBattleScrollDown
                                    If InGame Then
                                        If Player(MyIndex).InBattle > 0 Then
                                            BattleScrollDown = False
                                        End If
                                    End If
                                Case ButtonEnum.EvolveYes
                                    If InGame Then
                                        If IsEvolve And Not IsLearnMove Then
                                            SendEvolve EvolvePoke
                                            CloseEvolve
                                        End If
                                    End If
                                Case ButtonEnum.EvolveNo
                                    If InGame Then
                                        If IsEvolve And Not IsLearnMove Then
                                            CloseEvolve
                                        End If
                                    End If
                                Case ButtonEnum.InvScrollUp
                                    If InGame Then
                                        If WindowVisible(WindowType.Main_Inventory) Then
                                            IsInvScrollUp = False
                                        End If
                                    End If
                                Case ButtonEnum.InvScrollDown
                                    If InGame Then
                                        If WindowVisible(WindowType.Main_Inventory) Then
                                            IsInvScrollDown = False
                                        End If
                                    End If
                                Case ButtonEnum.InvItems
                                    If InGame Then
                                        If WindowVisible(WindowType.Main_Inventory) Then
                                            CurInvType = ItemType.Items: StartInv = 1
                                            UseItemNum = 0
                                            ShowUseItem = False
                                        End If
                                    End If
                                Case ButtonEnum.InvPokeballs
                                    If InGame Then
                                        If WindowVisible(WindowType.Main_Inventory) Then
                                            CurInvType = ItemType.Pokeballs: StartInv = 1
                                            UseItemNum = 0
                                            ShowUseItem = False
                                        End If
                                    End If
                                Case ButtonEnum.InvTM_HMs
                                    If InGame Then
                                        If WindowVisible(WindowType.Main_Inventory) Then
                                            CurInvType = ItemType.TM_HMs: StartInv = 1
                                            UseItemNum = 0
                                            ShowUseItem = False
                                        End If
                                    End If
                                Case ButtonEnum.InvBerries
                                    If InGame Then
                                        If WindowVisible(WindowType.Main_Inventory) Then
                                            CurInvType = ItemType.Berries: StartInv = 1
                                            UseItemNum = 0
                                            ShowUseItem = False
                                        End If
                                    End If
                                Case ButtonEnum.InvKeyItems
                                    If InGame Then
                                        If WindowVisible(WindowType.Main_Inventory) Then
                                            CurInvType = ItemType.KeyItems: StartInv = 1
                                            UseItemNum = 0
                                            ShowUseItem = False
                                        End If
                                    End If
                                Case ButtonEnum.PCDepositPoke
                                    If InGame Then
                                        If WindowVisible(WindowType.Main_Storage) Then
                                            ShowStorageSelect = Not ShowStorageSelect
                                            SelStoragePoke = 0
                                        End If
                                    End If
                                Case ButtonEnum.PCClose
                                    If InGame Then
                                        If WindowVisible(WindowType.Main_Storage) Then
                                            CloseStorage
                                            SelStoragePoke = 0
                                        End If
                                    End If
                                Case ButtonEnum.PCNext
                                    If InGame Then
                                        If WindowVisible(WindowType.Main_Storage) Then
                                            ChangeStorageTab True
                                            SelStoragePoke = 0
                                        End If
                                    End If
                                Case ButtonEnum.PCPrevious
                                    If InGame Then
                                        If WindowVisible(WindowType.Main_Storage) Then
                                            ChangeStorageTab False
                                            SelStoragePoke = 0
                                        End If
                                    End If
                                Case ButtonEnum.ShopScrollUp
                                    InShopScrollUp = False
                                Case ButtonEnum.ShopScrollDown
                                    InShopScrollDown = False
                                Case ButtonEnum.TradeConfirm
                                    If InGame Then
                                        If InTrade Then
                                            SendTradeConfirm
                                        End If
                                    End If
                                
                                Case ButtonEnum.TradeAccept
                                    If InGame Then
                                        If InTradeConfirm Then
                                            ' Accept Trade
                                        End If
                                    End If
                                Case ButtonEnum.TradeDecline
                                    If InGame Then
                                        If InTradeConfirm Then
                                            CloseTrade
                                        End If
                                    End If
                                    
                                ' Main End ~
                                Case Else
                            End Select
                        End If
                    End If
                End If
            End If
        End With
    Next
    ResetButtonState
    
    Exit Sub
errHandler:
    HandleError "HandleMouseUp", "modInput", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ResetButtonState()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    For i = 1 To ButtonEnum.MaxButton
        Buttons(i).bState = 0
    Next
    LastButtonClick = 0
    ChatScrollDown = False
    ChatScrollUp = False
    BattleScrollDown = False
    BattleScrollUp = False
    IsInvScrollUp = False
    IsInvScrollDown = False
    InShopScrollUp = False
    InShopScrollDown = False
    
    Exit Sub
errHandler:
    HandleError "ResetButtonState", "modInput", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleMenuKeyPress(ByVal KeyAscii As Integer)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If KeyAscii = vbKeyTab Or KeyAscii = vbKeyReturn Then
        If WindowVisible(WindowType.Menu_Login) Then
            If CurTextBox = 1 Then
                CurTextBox = 2
            ElseIf CurTextBox = 2 Then
                If KeyAscii = vbKeyTab Then
                    CurTextBox = 1
                ElseIf KeyAscii = vbKeyReturn Then
                    If user = vbNullString Then
                        CurTextBox = 1
                    Else
                        MenuState MenuStateLogin
                    End If
                End If
            End If
        ElseIf WindowVisible(WindowType.Menu_Register) Then
            If CurTextBox = 1 Then
                CurTextBox = 2
            ElseIf CurTextBox = 2 Then
                CurTextBox = 3
            ElseIf CurTextBox = 3 Then
                If KeyAscii = vbKeyTab Then
                    CurTextBox = 1
                ElseIf KeyAscii = vbKeyReturn Then
                    If user = vbNullString Then
                        CurTextBox = 1
                    Else
                        MenuState MenuStateRegister
                    End If
                End If
            End If
        ElseIf WindowVisible(WindowType.Menu_CharCreate) Then
            If KeyAscii = vbKeyReturn Then
                MenuState MenuStateAddChar
            End If
        End If
    End If
    
    If WindowVisible(WindowType.Menu_Login) Then
        If CurTextBox = 1 Then
            If Len(user) < (MAX_STRING - 1) Or KeyAscii = vbKeyBack Then user = EnterText(user, KeyAscii)
        ElseIf CurTextBox = 2 Then
            If Len(Pass) < (MAX_STRING - 1) Or KeyAscii = vbKeyBack Then Pass = EnterText(Pass, KeyAscii)
        End If
    ElseIf WindowVisible(WindowType.Menu_Register) Then
        If CurTextBox = 1 Then
            If Len(user) < (MAX_STRING - 1) Or KeyAscii = vbKeyBack Then user = EnterText(user, KeyAscii)
        ElseIf CurTextBox = 2 Then
            If Len(Pass) < (MAX_STRING - 1) Or KeyAscii = vbKeyBack Then Pass = EnterText(Pass, KeyAscii)
        ElseIf CurTextBox = 3 Then
            If Len(Pass2) < (MAX_STRING - 1) Or KeyAscii = vbKeyBack Then Pass2 = EnterText(Pass2, KeyAscii)
        End If
    ElseIf WindowVisible(WindowType.Menu_CharCreate) Then
        If Len(user) < (MAX_STRING - 1) Or KeyAscii = vbKeyBack Then user = EnterText(user, KeyAscii)
    End If
    
    Exit Sub
errHandler:
    HandleError "HandleMenuKeyPress", "modInput", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleMainKeyPress(ByVal KeyAscii As Integer)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Editor = EDITOR_MAP Then Exit Sub
    
    If ShowInput Then
        ChatMsg = vbNullString
        RenderChatMsg = UpdateChatText(ChatMsg, MaxMsg)
        ChatOn = False
        
        If KeyAscii = vbKeyBack Or IsNumeric(ChrW$(KeyAscii)) Then
            InputVal = EnterText(InputVal, KeyAscii)
            If Val(InputVal) > 99 Then
                InputVal = 99
            End If
            RenderVal = UpdateChatText(InputVal, 110)
        End If
    Else
        If KeyAscii = vbKeyReturn Then
            If IsLearnMove Or IsEvolve Then
                ChatMsg = vbNullString
                RenderChatMsg = UpdateChatText(ChatMsg, MaxMsg)
                ChatOn = False
            Else
                If ChatOn = False Then
                    ChatOn = True
                Else
                    If Len(ChatMsg) > 0 Then
                        SendChat
                    Else
                        ChatMsg = vbNullString
                        RenderChatMsg = UpdateChatText(ChatMsg, MaxMsg)
                        ChatOn = False
                    End If
                End If
            End If
            Exit Sub
        End If
        
        If ChatOn Then
            If KeyAscii <> vbKeyReturn Or KeyAscii <> vbKeyTab Then
                ChatMsg = EnterText(ChatMsg, KeyAscii)
            End If
            RenderChatMsg = UpdateChatText(ChatMsg, MaxMsg)
        End If
    End If
    
    Exit Sub
errHandler:
    HandleError "HandleMainKeyPress", "modInput", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SendChat()
Dim Command() As String
Dim MyText As String
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Left$(ChatMsg, 1) = "'" Then
        MyText = Mid$(ChatMsg, 2, Len(ChatMsg) - 1)

        If Len(MyText) > 0 Then
            Call SendMsg(MyText, GlobalMsg)
        End If

        ChatMsg = vbNullString
        MyText = vbNullString
        RenderChatMsg = UpdateChatText(ChatMsg, MaxMsg)
        ChatOn = False
        Exit Sub
    End If
                
    If Left$(ChatMsg, 1) = "/" Then
        Command = Split(ChatMsg, Space(1))

        Select Case Command(0)
            Case "/editmap"
                If GetPlayerAccess(MyIndex) >= ACCESS_MAPPER Then
                    SendRequestEditMap
                Else
                    Call AddText("Invalid command!", Red)
                End If
            Case "/editpokemon"
                If GetPlayerAccess(MyIndex) >= ACCESS_DEVELOPER Then
                    Call SendRequestEditPokemon
                Else
                    Call AddText("Invalid command!", Red)
                End If
            Case "/editmove"
                If GetPlayerAccess(MyIndex) >= ACCESS_DEVELOPER Then
                    Call SendRequestEditMove
                Else
                    Call AddText("Invalid command!", Red)
                End If
            Case "/editexpcalc"
                If GetPlayerAccess(MyIndex) >= ACCESS_DEVELOPER Then
                    frmExpCalc.Show
                Else
                    Call AddText("Invalid command!", Red)
                End If
            Case "/edititem"
                If GetPlayerAccess(MyIndex) >= ACCESS_DEVELOPER Then
                    Call SendRequestEditItem
                Else
                    Call AddText("Invalid command!", Red)
                End If
            Case "/editnpc"
                If GetPlayerAccess(MyIndex) >= ACCESS_DEVELOPER Then
                    Call SendRequestEditNPC
                Else
                    Call AddText("Invalid command!", Red)
                End If
            Case "/editshop"
                If GetPlayerAccess(MyIndex) >= ACCESS_DEVELOPER Then
                    Call SendRequestEditShop
                Else
                    Call AddText("Invalid command!", Red)
                End If
            Case "/accept"
                If BattleRequestIndex > 0 Then
                    SendInitBattle YES
                    BattleRequestIndex = 0
                    GoTo continue
                End If
                If InTradeIndex > 0 Then
                    SendTradeAccept
                    GoTo continue
                End If
            Case "/decline"
                If BattleRequestIndex > 0 Then
                    SendInitBattle NO
                    BattleRequestIndex = 0
                    GoTo continue
                End If
                If InTradeIndex > 0 Then
                    SendTradeDecline
                    GoTo continue
                End If
            Case Else
                Call AddText("Invalid command!", Red)
        End Select

continue:
        ChatMsg = vbNullString
        MyText = vbNullString
        RenderChatMsg = UpdateChatText(ChatMsg, MaxMsg)
        ChatOn = False
        Exit Sub
    End If
    
    If Len(ChatMsg) > 0 Then
        Call SendMsg(ChatMsg, MapMsg)
    End If

    ChatMsg = vbNullString
    MyText = vbNullString
    RenderChatMsg = UpdateChatText(ChatMsg, MaxMsg)
    ChatOn = False
    
    Exit Sub
errHandler:
    HandleError "SendChat", "modInput", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub
