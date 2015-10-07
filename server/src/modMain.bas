Attribute VB_Name = "modMain"
Option Explicit

Sub Main()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    AddLog "Checking Directory..."
    ChkDir App.Path & "\", "bin"
    ChkDir App.Path & "\bin\", "report"
    ChkDir App.Path & "\bin\", "players"
    ChkDir App.Path & "\bin\", "maps"
    ChkDir App.Path & "\bin\", "pokemons"
    ChkDir App.Path & "\bin\", "moves"
    ChkDir App.Path & "\bin\", "items"
    ChkDir App.Path & "\bin\", "npcs"
    ChkDir App.Path & "\bin\", "shops"
    
    frmMain.Show
    
    AddLog "Setting up Socket..."
    frmMain.Socket(0).RemoteHost = frmMain.Socket(0).LocalIP
    frmMain.Socket(0).LocalPort = 7001
    
    For i = 1 To MAX_PLAYER
        Call ClearPlayer(i)
        Load frmMain.Socket(i)
    Next
    
    Call InitMessages
    
    AddLog "Clearing Maps..."
    ClearMaps
    AddLog "Clearing Pokemons..."
    ClearPokemons
    AddLog "Clearing Moves..."
    ClearMoves
    AddLog "Clearing Items..."
    ClearItems
    AddLog "Clearing NPCs..."
    ClearNPCs
    AddLog "Clearing Shops..."
    ClearShops
    AddLog "Clearing Map NPCs..."
    ClearMapNpcs
    AddLog "Loading Maps..."
    LoadMaps
    AddLog "Loading Pokemons..."
    LoadPokemons
    AddLog "Loading Moves..."
    LoadMoves
    AddLog "Loading Items..."
    LoadItems
    AddLog "Loading NPCs..."
    LoadNPCs
    AddLog "Loading Shops..."
    LoadShops
    AddLog "Loading ExpCalc..."
    LoadExpCalc
    AddLog "Spawning All Map Npcs..."
    SpawnAllMapNpcs
    CreateFullMapCache
    
    frmMain.Socket(0).Listen
    AddLog "Initialization complete..."
    
    AppOpen = True
    AppLoop

    Exit Sub
errHandler:
    Err.Clear
    CloseApp
    Exit Sub
End Sub

Public Sub CloseApp()
Dim i As Long

    For i = 1 To MAX_PLAYER
        Unload frmMain.Socket(i)
    Next
    Unload frmMain
    AppOpen = False
    End
End Sub

Public Sub AddLog(ByVal Text As String)
Dim s As String

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    s = frmMain.txtLog.Text
    If Len(s) = 0 Then
        frmMain.txtLog.Text = Text
    Else
        frmMain.txtLog.Text = s & vbNewLine & Text
    End If
    frmMain.txtLog.SelStart = Len(frmMain.txtLog.Text)
    
    Exit Sub
errHandler:
    HandleError "AddLog", "modMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function isNameLegal(ByVal KeyAscii As Integer) As Boolean
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 95) Or (KeyAscii = 32) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        isNameLegal = True
    End If
    
    Exit Function
errHandler:
    HandleError "isNameLegal", "modMain", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function CheckNameInput(ByVal Name As String) As Boolean
Dim i As Long, n As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Len(Name) <= 2 Or Len(Name) > (MAX_STRING - 1) Then
        CheckNameInput = False
        Exit Function
    End If
    
    For i = 1 To Len(Name)
        n = AscW(Mid$(Name, i, 1))

        If Not isNameLegal(n) Then
            CheckNameInput = False
            Exit Function
        End If
    Next
    CheckNameInput = True
    
    Exit Function
errHandler:
    HandleError "CheckNameInput", "modMain", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub JoinGame(ByVal index As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Not IsPlaying(index) Then
        TempPlayer(index).InGame = True
        
        frmMain.lvwInfo.ListItems(index).SubItems(1) = GetPlayerIP(index)
        frmMain.lvwInfo.ListItems(index).SubItems(2) = Trim$(Player(index).Username)
        frmMain.lvwInfo.ListItems(index).SubItems(3) = GetPlayerName(index)
        
        If Player(index).PlayerData(TempPlayer(index).CurSlot).TempData > 0 Then
            GivePlayerPokemon index, Player(index).PlayerData(TempPlayer(index).CurSlot).TempData, 5
            Player(index).PlayerData(TempPlayer(index).CurSlot).TempData = 0
        End If
        
        SendIndex index
        
        SendPokemons index
        SendMoves index
        SendExpCalc index
        SendItems index
        SendInventory index
        SendNPCs index
        SendShops index
        
        SendPlayerStoredPokemons index
        
        PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
        SendInGame index
        
        AddLog "'Player: " & GetPlayerName(index) & "' has joined the game..."
        
        SendMsg index, "Welcome to 'Pokemon Engine'", White
    End If
    
    Exit Sub
errHandler:
    HandleError "JoinGame", "modMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub LeftGame(ByVal index As Long)
Dim x As Long, i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If TempPlayer(index).InGame Then
        TempPlayer(index).InGame = False
        
        If TempPlayer(index).InBattle = BATTLE_TRAINER Then
            x = TempPlayer(index).BattleRequest
            SendMsg x, Trim$(Player(index).PlayerData(TempPlayer(index).CurSlot).Name) & " has been disconnected!", Red
            ClosePlayerBattle x
            GivePvPPoints index, 3
            GivePvPPoints x, 3
        End If

        i = TempPlayer(index).InTradeRequest
        If i > 0 And i <= MAX_PLAYER Then
            If IsPlaying(i) Then
                SendMsg i, "Trade request has been cancelled!", Red
                TempPlayer(i).InTradeRequest = 0
                TempPlayer(i).InTradeRequest = 0
            End If
        End If
        i = TempPlayer(index).InTrade
        If i > 0 And i <= MAX_PLAYER Then
            If IsPlaying(i) Then
                SendMsg i, "Trade has been cancelled!", Red
                TempPlayer(i).InTrade = 0
                SendCloseTrade i
            End If
        End If
        
        Call SavePlayer(index)
        Call SavePlayerPokemon(index)
        Call SendLeftGame(index)
        ClearAllTarget index
        AddLog "'Player: " & GetPlayerName(index) & "' has left the game..."
    End If
    Call ClearPlayer(index)
    
    Exit Sub
errHandler:
    HandleError "JoinGame", "modMain", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

