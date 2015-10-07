Attribute VB_Name = "modMain"
Option Explicit

Sub Main()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ChkDir App.Path & "\", "bin"
    ChkDir App.Path & "\bin\", "report"
    ChkDir App.Path & "\bin\", "_maps"
    ChkDir App.Path & "\bin\", "font"
    ChkDir App.Path & "\", "data"
    ChkDir App.Path & "\data\", "gfx"
    ChkDir App.Path & "\data\", "sfx"
    ChkDir App.Path & "\data\sfx\", "music"
    ChkDir App.Path & "\data\sfx\", "sound"
    ChkDir App.Path & "\data\gfx\", "misc"
    ChkDir App.Path & "\data\gfx\", "gui"
    ChkDir App.Path & "\data\gfx\", "sprites"
    ChkDir App.Path & "\data\gfx\", "tilesets"
    ChkDir App.Path & "\data\gfx\", "pokemons"
    ChkDir App.Path & "\data\gfx\", "backgrounds"
    ChkDir App.Path & "\data\gfx\", "fields"
    ChkDir App.Path & "\data\gfx\", "titlebars"
    ChkDir App.Path & "\data\gfx\pokemons\", "icons"
    ChkDir App.Path & "\data\gfx\pokemons\", "front"
    ChkDir App.Path & "\data\gfx\pokemons\", "back"
    ChkDir App.Path & "\data\gfx\pokemons\", "sprite"
    ChkDir App.Path & "\data\gfx\gui\", "buttons"
    ChkDir App.Path & "\data\gfx\", "items"
    
    LoadOption
    
    InitSound
    EngineInit

    Call TcpInit
    Call InitMessages
    
    CurFont = Font_Georgia
    InitButton
    
    InitMenu
    
    If Not FileExist(App.Path & MUSIC_PATH & Menu_Music) Then
        StopMusic
    Else
        PlayMusic Menu_Music
    End If

    frmMain.Show
    AppOpen = True
    AppLoop
    
    Exit Sub
errHandler:
    Err.Clear
    CloseApp
    Exit Sub
End Sub

Public Sub LogOutGame()
Dim X As Long

    InitMenu
    ' Close All In-Game data
    ' Clear All In-Game Data
    ClearMap
    MyIndex = 0
    Editor = 0
    
    DestroyTCP
    For X = 1 To MAX_PLAYER_DATA
        CharSelectName(X) = vbNullString
        CharSelectSprite(X) = 0
    Next
    CharSelected = 0
End Sub

Public Sub CloseApp()
    EngineUnloadDirectX
    CloseSound
    CloseAllForms
    AppOpen = False
    End
End Sub
Private Sub CloseAllForms()
Dim Frm As Form

    For Each Frm In VB.Forms
        Unload Frm
    Next
End Sub

Private Sub InitButton()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    For i = 1 To ButtonEnum.MaxButton
        With Buttons(i)
            .bState = 0
            .Visible = False
        End With
    Next
    
    ' Menu ~
    With Buttons(ButtonEnum.LoginAccept)
        .Pic = 1
        .X = GuiLoginX + 70
        .Y = GuiLoginY + 112
    End With
    With Buttons(ButtonEnum.Register)
        .Pic = 2
        .X = GuiLoginX + 175
        .Y = GuiLoginY + 112
    End With
    
    With Buttons(ButtonEnum.RegisterAccept)
        .Pic = 1
        .X = GuiRegisterX + 70
        .Y = GuiRegisterY + 115
    End With
    With Buttons(ButtonEnum.Login)
        .Pic = 2
        .X = GuiRegisterX + 175
        .Y = GuiRegisterY + 115
    End With
    
    With Buttons(ButtonEnum.CharAccept)
        .Pic = 1
        .X = GuiCharCreateX + 30
        .Y = GuiCharCreateY + 147
    End With
    With Buttons(ButtonEnum.CharDecline)
        .Pic = 2
        .X = GuiCharCreateX + 130
        .Y = GuiCharCreateY + 147
    End With
    
    With Buttons(ButtonEnum.CharNew1)
        .Pic = 3
        .X = GuiCharSelectX + 17
        .Y = GuiCharSelectY + 45
    End With
    With Buttons(ButtonEnum.CharNew2)
        .Pic = 3
        .X = GuiCharSelectX + 79
        .Y = GuiCharSelectY + 45
    End With
    With Buttons(ButtonEnum.CharNew3)
        .Pic = 3
        .X = GuiCharSelectX + 141
        .Y = GuiCharSelectY + 45
    End With
    With Buttons(ButtonEnum.CharUse1)
        .Pic = 4
        .X = GuiCharSelectX + 17
        .Y = GuiCharSelectY + 45
    End With
    With Buttons(ButtonEnum.CharUse2)
        .Pic = 4
        .X = GuiCharSelectX + 79
        .Y = GuiCharSelectY + 45
    End With
    With Buttons(ButtonEnum.CharUse3)
        .Pic = 4
        .X = GuiCharSelectX + 141
        .Y = GuiCharSelectY + 45
    End With
    With Buttons(ButtonEnum.CharDel1)
        .Pic = 5
        .X = GuiCharSelectX + 17
        .Y = GuiCharSelectY + 115
    End With
    With Buttons(ButtonEnum.CharDel2)
        .Pic = 5
        .X = GuiCharSelectX + 79
        .Y = GuiCharSelectY + 115
    End With
    With Buttons(ButtonEnum.CharDel3)
        .Pic = 5
        .X = GuiCharSelectX + 141
        .Y = GuiCharSelectY + 115
    End With
    ' Menu End ~
    ' Main
    With Buttons(ButtonEnum.bChatScrollUp)
        .Pic = 6
        .X = GuiChatboxX + 393
        .Y = GuiChatboxY + 11
    End With
    With Buttons(ButtonEnum.bChatScrollDown)
        .Pic = 7
        .X = GuiChatboxX + 393
        .Y = GuiChatboxY + 102
    End With
    With Buttons(ButtonEnum.mPokedex)
        .Pic = 10
        .X = GuiChatboxX + 620
        .Y = GuiChatboxY + 100
    End With
    With Buttons(ButtonEnum.mCharacter)
        .Pic = 9
        .X = GuiChatboxX + 665
        .Y = GuiChatboxY + 100
    End With
    With Buttons(ButtonEnum.mInventory)
        .Pic = 8
        .X = GuiChatboxX + 710
        .Y = GuiChatboxY + 100
    End With
    With Buttons(ButtonEnum.mOptions)
        .Pic = 11
        .X = GuiChatboxX + 755
        .Y = GuiChatboxY + 100
    End With
    
    With Buttons(ButtonEnum.BattleFight)
        .Pic = 14
        .X = GuiBattleX + 542
        .Y = GuiBattleY + 422
    End With
    With Buttons(ButtonEnum.BattleSwitch)
        .Pic = 12
        .X = GuiBattleX + 665
        .Y = GuiBattleY + 422
    End With
    With Buttons(ButtonEnum.BattleBag)
        .Pic = 13
        .X = GuiBattleX + 542
        .Y = GuiBattleY + 484
    End With
    With Buttons(ButtonEnum.BattleRun)
        .Pic = 15
        .X = GuiBattleX + 665
        .Y = GuiBattleY + 484
    End With
    
    With Buttons(ButtonEnum.bBattleScrollUp)
        .Pic = 6
        .X = GuiBattleX + 757
        .Y = GuiBattleY + 88
    End With
    With Buttons(ButtonEnum.bBattleScrollDown)
        .Pic = 7
        .X = GuiBattleX + 757
        .Y = GuiBattleY + 371
    End With
    
    With Buttons(ButtonEnum.EvolveYes)
        .Pic = 1
        .X = GuiEvolveX + 164
        .Y = GuiEvolveY + 348
    End With
    With Buttons(ButtonEnum.EvolveNo)
        .Pic = 2
        .X = GuiEvolveX + 289
        .Y = GuiEvolveY + 348
    End With
    
    With Buttons(ButtonEnum.InvScrollUp)
        .Pic = 6
        .X = GuiTrainerX + 226
        .Y = GuiTrainerY + 22
    End With
    With Buttons(ButtonEnum.InvScrollDown)
        .Pic = 7
        .X = GuiTrainerX + 226
        .Y = GuiTrainerY + 194
    End With
    
    With Buttons(ButtonEnum.InvItems)
        .Pic = 16
        .X = GuiTrainerX - 46
        .Y = GuiTrainerY + 28
    End With
    With Buttons(ButtonEnum.InvPokeballs)
        .Pic = 16
        .X = GuiTrainerX - 46
        .Y = GuiTrainerY + 53
    End With
    With Buttons(ButtonEnum.InvTM_HMs)
        .Pic = 16
        .X = GuiTrainerX - 46
        .Y = GuiTrainerY + 78
    End With
    With Buttons(ButtonEnum.InvBerries)
        .Pic = 16
        .X = GuiTrainerX - 46
        .Y = GuiTrainerY + 103
    End With
    With Buttons(ButtonEnum.InvKeyItems)
        .Pic = 16
        .X = GuiTrainerX - 46
        .Y = GuiTrainerY + 128
    End With
    
    With Buttons(ButtonEnum.PCDepositPoke)
        .Pic = 16
        .X = GuiStorageX + 28
        .Y = GuiStorageY + 330
    End With
    With Buttons(ButtonEnum.PCClose)
        .Pic = 16
        .X = GuiStorageX + 341
        .Y = GuiStorageY + 330
    End With
    
    With Buttons(ButtonEnum.PCPrevious)
        .Pic = 16
        .X = GuiStorageX + 119
        .Y = GuiStorageY + 325
    End With
    With Buttons(ButtonEnum.PCNext)
        .Pic = 16
        .X = GuiStorageX + 250
        .Y = GuiStorageY + 325
    End With
    
    With Buttons(ButtonEnum.ShopScrollUp)
        .Pic = 6
        .X = GuiShopX + 228
        .Y = GuiShopY + 39
    End With
    With Buttons(ButtonEnum.ShopScrollDown)
        .Pic = 7
        .X = GuiShopX + 228
        .Y = GuiShopY + 213
    End With
    
    With Buttons(ButtonEnum.TradeConfirm)
        .Pic = 1
        .X = GuiTradeX + 89
        .Y = GuiTradeY + 322
    End With
    
    With Buttons(ButtonEnum.TradeAccept)
        .Pic = 1
        .X = GuiTradeConfirmX + 78
        .Y = GuiTradeConfirmY + 273
    End With
    With Buttons(ButtonEnum.TradeDecline)
        .Pic = 2
        .X = GuiTradeConfirmX + 168
        .Y = GuiTradeConfirmY + 273
    End With
    ' Main End ~

    Exit Sub
errHandler:
    HandleError "InitButton", "modMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub InitMenu()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    InMenu = True
    InGame = False
    OpenWindow Menu_Login
    
    Exit Sub
errHandler:
    HandleError "InitMenu", "modMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub InitGame()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    GettingMap = True
    
    InMenu = False
    InGame = True
    
    ChatScroll = MaxChatLine
    OpenWindow Main_Chatbox
    BattleScroll = MaxBattleLine
    
    For i = WindowType.Menu_Login To WindowType.Menu_CharCreate
        WindowVisible(i) = False
    Next i
    For i = ButtonEnum.LoginAccept To ButtonEnum.CharDecline
        Buttons(i).Visible = False
    Next i
    
    Exit Sub
errHandler:
    HandleError "InitGame", "modMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub OpenWindow(ByVal WindowNum As WindowType)
Dim X As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    CurTextBox = 0
    user = vbNullString: Pass = vbNullString
    
    For X = ButtonEnum.CharNew1 To ButtonEnum.CharDel3
        Buttons(X).Visible = False
    Next
    
    Select Case WindowNum
        ' Menu ~
        Case Menu_Login
            WindowVisible(WindowType.Menu_Register) = False
            WindowVisible(WindowType.Menu_Login) = True
            WindowVisible(WindowType.Menu_CharSelect) = False
            WindowVisible(WindowType.Menu_CharCreate) = False
            
            Buttons(ButtonEnum.LoginAccept).Visible = True
            Buttons(ButtonEnum.Register).Visible = True
            Buttons(ButtonEnum.RegisterAccept).Visible = False
            Buttons(ButtonEnum.Login).Visible = False
            Buttons(ButtonEnum.CharAccept).Visible = False
            Buttons(ButtonEnum.CharDecline).Visible = False
            CurTextBox = 1
            If Len(Options.Username) > 0 Then
                user = Trim$(Options.Username)
            End If
            If SaveAccount Then
                If Len(Options.Password) > 0 Then
                    Pass = Trim$(Options.Password)
                End If
            End If
        Case Menu_Register
            WindowVisible(WindowType.Menu_Register) = True
            WindowVisible(WindowType.Menu_Login) = False
            WindowVisible(WindowType.Menu_CharSelect) = False
            WindowVisible(WindowType.Menu_CharCreate) = False
            
            Buttons(ButtonEnum.RegisterAccept).Visible = True
            Buttons(ButtonEnum.Login).Visible = True
            Buttons(ButtonEnum.LoginAccept).Visible = False
            Buttons(ButtonEnum.Register).Visible = False
            Buttons(ButtonEnum.CharAccept).Visible = False
            Buttons(ButtonEnum.CharDecline).Visible = False
            CurTextBox = 1
        Case Menu_CharSelect
            WindowVisible(WindowType.Menu_Register) = False
            WindowVisible(WindowType.Menu_Login) = False
            WindowVisible(WindowType.Menu_CharSelect) = True
            WindowVisible(WindowType.Menu_CharCreate) = False
            
            Buttons(ButtonEnum.LoginAccept).Visible = False
            Buttons(ButtonEnum.Register).Visible = False
            Buttons(ButtonEnum.RegisterAccept).Visible = False
            Buttons(ButtonEnum.Login).Visible = False
            Buttons(ButtonEnum.CharAccept).Visible = False
            Buttons(ButtonEnum.CharDecline).Visible = False
            
            For X = 1 To MAX_PLAYER_DATA
                ' Sprite, Level
                If Len(Trim$(CharSelectName(X))) > 0 Then
                    Buttons(X + 7).Visible = True
                    Buttons(X + 10).Visible = True
                Else
                    Buttons(X + 4).Visible = True
                End If
            Next
        Case Menu_CharCreate
            WindowVisible(WindowType.Menu_Register) = False
            WindowVisible(WindowType.Menu_Login) = False
            WindowVisible(WindowType.Menu_CharSelect) = False
            WindowVisible(WindowType.Menu_CharCreate) = True
            
            Buttons(ButtonEnum.LoginAccept).Visible = False
            Buttons(ButtonEnum.Register).Visible = False
            Buttons(ButtonEnum.RegisterAccept).Visible = False
            Buttons(ButtonEnum.Login).Visible = False
            Buttons(ButtonEnum.CharAccept).Visible = True
            Buttons(ButtonEnum.CharDecline).Visible = True
            SelGender = GENDER_MALE
            SelStarter = 1
        ' Menu End ~
        ' Main ~
        Case Main_Chatbox
            WindowVisible(WindowType.Main_Chatbox) = True
            Buttons(ButtonEnum.bChatScrollUp).Visible = True
            Buttons(ButtonEnum.bChatScrollDown).Visible = True
            
            For X = ButtonEnum.mPokedex To ButtonEnum.mOptions
                Buttons(X).Visible = True
            Next X
            CloseInventory
        ' Main End ~
    End Select
    ResetButtonState
    
    Exit Sub
errHandler:
    HandleError "InitMenu", "modMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub MenuState(ByVal MnuState As Byte)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If ConnectToServer Then
        Select Case MnuState
            Case MenuStateLogin
                If CheckNameInput(user) Then
                    If CheckNameInput(Pass) Then
                        SendLogin user, Pass
                        SaveUser = user
                        SavePass = Pass
                    End If
                End If
            Case MenuStateRegister
                If CheckNameInput(user) Then
                    If CheckNameInput(Pass) Then
                        If Pass = Pass2 Then
                            SendRegisterData user, Pass
                        End If
                    End If
                End If
            Case MenuStateAddChar
                If CheckNameInput(user) Then
                    If CharSelected > 0 And CharSelected <= MAX_PLAYER_DATA Then
                        SendAddChar user, SelGender, CharSelected, SelStarter
                    End If
                End If
            Case MenuStateDelChar
                If CharSelected > 0 And CharSelected <= MAX_PLAYER_DATA Then
                    SendDelChar CharSelected
                End If
            Case MenuStateUseChar
                If CharSelected > 0 And CharSelected <= MAX_PLAYER_DATA Then
                    SendUseChar CharSelected
                End If
        End Select
    Else
        Call MsgBox("Server is offline! Please try again later.", vbCritical)
    End If
    
    Exit Sub
errHandler:
    HandleError "MenuState", "modMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub OpenInventory()
Dim X As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If InTradeConfirm Then Exit Sub
    
    WindowVisible(WindowType.Main_Inventory) = True
    Buttons(ButtonEnum.InvScrollUp).Visible = True
    Buttons(ButtonEnum.InvScrollDown).Visible = True
    For X = ButtonEnum.InvItems To ButtonEnum.InvKeyItems
        Buttons(X).Visible = True
    Next
    CurInvType = ItemType.Items
    StartInv = 1
    
    UseItemNum = 0
    ShowUseItem = False
    
    Exit Sub
errHandler:
    HandleError "OpenInventory", "modMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub CloseInventory(Optional ByVal Forced As Boolean = False)
Dim X As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If InShop > 0 Then Exit Sub
    If Not Forced Then
        If InTrade Or InTradeConfirm Then Exit Sub
    End If
    
    WindowVisible(WindowType.Main_Inventory) = False
    Buttons(ButtonEnum.InvScrollUp).Visible = False
    Buttons(ButtonEnum.InvScrollDown).Visible = False
    For X = ButtonEnum.InvItems To ButtonEnum.InvKeyItems
        Buttons(X).Visible = False
    Next
    CurInvType = ItemType.Items
    StartInv = 1
    
    UseItemNum = 0
    ShowUseItem = False
    
    Exit Sub
errHandler:
    HandleError "CloseInventory", "modMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub OpenStorage()
Dim X As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    WindowVisible(WindowType.Main_Storage) = True
    For X = ButtonEnum.PCDepositPoke To ButtonEnum.PCPrevious
        Buttons(X).Visible = True
    Next
    InStorage = True
    StartStorage = 1
    ShowStorageSelect = False
    
    Exit Sub
errHandler:
    HandleError "OpenStorage", "modMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub CloseStorage()
Dim X As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    WindowVisible(WindowType.Main_Storage) = False
    For X = ButtonEnum.PCDepositPoke To ButtonEnum.PCPrevious
        Buttons(X).Visible = False
    Next
    InStorage = False
    StartStorage = 1
    ShowStorageSelect = False
    
    Exit Sub
errHandler:
    HandleError "CloseStorage", "modMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub InitInput(ByVal IType As Byte, ByVal X As Long, ByVal Y As Long, Optional ByVal xData As Long = 0)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ShowInput = True
    InputType = IType
    InputData = xData
    InputVal = "1"
    RenderVal = UpdateChatText(InputVal, 110)
    RenderValX = X
    RenderValY = Y
    
    Exit Sub
errHandler:
    HandleError "InitInput", "modMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ConfirmInput()
Dim xVal As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If ShowInput Then
        If IsNumeric(InputVal) Then
            xVal = Val(InputVal)
            If xVal <= 0 Then xVal = 1
            Select Case InputType
                Case 1
                    SendSellItem InputData, xVal
                    UseItemNum = 0
                    ShowUseItem = False
                Case 2
                    SendBuyItem InputData, xVal
                    ShopSelect = 0
                Case 3
                    If xVal > Player(MyIndex).Item(InputData, CurInvType).value Then
                        xVal = Player(MyIndex).Item(InputData, CurInvType).value
                    End If
                    AddTrade TRADE_TYPE_ITEM, , InputData, xVal
                    ShowTradeSel = False
                Case Else
            End Select
            
            CloseInput
        End If
    End If
    
    Exit Sub
errHandler:
    HandleError "ConfirmInput", "modMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub CloseInput()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ShowInput = False
    InputType = 0
    InputVal = vbNullString
    RenderVal = UpdateChatText(InputVal, 110)
    RenderValX = 0
    RenderValY = 0
    InputData = 0
    
    Exit Sub
errHandler:
    HandleError "CloseInput", "modMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub CloseTrade()
Dim Buffer As clsBuffer

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    InTrade = False
    InTradeConfirm = False
    InTradeIndex = 0
    Set Buffer = New clsBuffer
    Buffer.WriteLong CCloseTrade
    SendData Buffer.ToArray()
    Set Buffer = Nothing
    
    Exit Sub
errHandler:
    HandleError "CloseTrade", "modMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub
