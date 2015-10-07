Attribute VB_Name = "modDatabase"
Option Explicit

Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc)
Dim FileName As String

    MsgBox "Run-time error '" & erNumber & "': " & erDesc & ".", vbCritical, AppTitle

    FileName = App.Path & "\bin\report\errors.txt"
    Open FileName For Append As #1
        Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
        Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
        Print #1, ""
    Close #1
    
    CloseApp
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
    
    Exit Sub
errHandler:
    HandleError "ChkDir", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function GetVar(file As String, Header As String, Var As String) As String
Dim sSpaces As String
Dim szReturn As String

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), file)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
    
    Exit Function
errHandler:
    HandleError "GetVar", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub PutVar(file As String, Header As String, Var As String, Value As String)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call WritePrivateProfileString$(Header, Var, Value, file)
    
    Exit Sub
errHandler:
    HandleError "PutVar", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function Random(ByVal Low As Long, ByVal High As Long) As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Randomize
    Random = Int((High - Low + 1) * Rnd) + Low
    
    Exit Function
errHandler:
    HandleError "Random", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function RandomDigit(ByVal Low As Single, ByVal High As Single) As Single
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Randomize
    RandomDigit = ((High - Low + 1) * Rnd) + Low
    
    Exit Function
errHandler:
    HandleError "RandomDigit", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function FileExist(ByVal FileName As String) As Boolean
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If LenB(Dir(FileName)) > 0 Then
        FileExist = True
    End If
    
    Exit Function
errHandler:
    HandleError "FileExist", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub ClearPlayer(ByVal index As Long)
Dim x As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call ZeroMemory(ByVal VarPtr(TempPlayer(index)), LenB(TempPlayer(index)))
    Set TempPlayer(index).buffer = New clsBuffer
    
    Call ZeroMemory(ByVal VarPtr(Player(index)), LenB(Player(index)))
    Player(index).Username = vbNullString
    Player(index).Password = vbNullString
    
    For x = 1 To MAX_PLAYER_DATA
        Player(index).PlayerData(x).Name = vbNullString
    Next
    
    frmMain.lvwInfo.ListItems(index).SubItems(1) = vbNullString
    frmMain.lvwInfo.ListItems(index).SubItems(2) = vbNullString
    frmMain.lvwInfo.ListItems(index).SubItems(3) = vbNullString
    
    Exit Sub
errHandler:
    HandleError "ClearTempPlayer", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SavePlayersOnline()
Dim i As Long

    For i = 1 To HighPlayerIndex
        If IsLoggedIn(i) > 0 Then
            Call SavePlayer(i)
            Call SavePlayerPokemon(i)
        End If
    Next
End Sub

Public Sub SavePlayer(ByVal index As Long)
Dim FileName As String
Dim F As Long
Dim x As Byte
Dim y As Long
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    FileName = App.Path & "\bin\players\" & Trim$(Player(index).Username) & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Put #F, , Player(index).Username
        Put #F, , Player(index).Password
        
        For i = 1 To MAX_PLAYER_DATA
            Put #F, , Player(index).PlayerData(i).Name
            Put #F, , Player(index).PlayerData(i).Gender
            
            Put #F, , Player(index).PlayerData(i).Access
            
            Put #F, , Player(index).PlayerData(i).Sprite
            
            Put #F, , Player(index).PlayerData(i).Map
            Put #F, , Player(index).PlayerData(i).x
            Put #F, , Player(index).PlayerData(i).y
            Put #F, , Player(index).PlayerData(i).Dir
            
            Put #F, , Player(index).PlayerData(i).Checkpoint.Map
            Put #F, , Player(index).PlayerData(i).Checkpoint.x
            Put #F, , Player(index).PlayerData(i).Checkpoint.y
            
            For x = 1 To MAX_PLAYER_ITEM
                For y = 0 To ItemType.Item_Count - 1
                    With Player(index).PlayerData(i)
                        Put #F, , .Item(x, y).Num
                        Put #F, , .Item(x, y).Value
                    End With
                Next
            Next
            
            Put #F, , Player(index).PlayerData(i).PvP.Win
            Put #F, , Player(index).PlayerData(i).PvP.Lose
            Put #F, , Player(index).PlayerData(i).PvP.Disconnect
            
            Put #F, , Player(index).PlayerData(i).Money
            
            Put #F, , Player(index).PlayerData(i).IsVIP
        Next
    Close #F
    
    Exit Sub
errHandler:
    HandleError "SavePlayer", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadPlayer(ByVal index As Long, ByVal User As String)
Dim FileName As String
Dim F As Long
Dim x As Byte
Dim y As Long
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    FileName = App.Path & "\bin\players\" & Trim$(User) & ".dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Get #F, , Player(index).Username
        Get #F, , Player(index).Password
        
        For i = 1 To MAX_PLAYER_DATA
            Get #F, , Player(index).PlayerData(i).Name
            Get #F, , Player(index).PlayerData(i).Gender
            
            Get #F, , Player(index).PlayerData(i).Access
            
            Get #F, , Player(index).PlayerData(i).Sprite
            
            Get #F, , Player(index).PlayerData(i).Map
            Get #F, , Player(index).PlayerData(i).x
            Get #F, , Player(index).PlayerData(i).y
            Get #F, , Player(index).PlayerData(i).Dir
            
            Get #F, , Player(index).PlayerData(i).Checkpoint.Map
            Get #F, , Player(index).PlayerData(i).Checkpoint.x
            Get #F, , Player(index).PlayerData(i).Checkpoint.y
            
            For x = 1 To MAX_PLAYER_ITEM
                For y = 0 To ItemType.Item_Count - 1
                    With Player(index).PlayerData(i)
                        Get #F, , .Item(x, y).Num
                        Get #F, , .Item(x, y).Value
                    End With
                Next
            Next
            
            Get #F, , Player(index).PlayerData(i).PvP.Win
            Get #F, , Player(index).PlayerData(i).PvP.Lose
            Get #F, , Player(index).PlayerData(i).PvP.Disconnect
            
            Get #F, , Player(index).PlayerData(i).Money
            
            Get #F, , Player(index).PlayerData(i).IsVIP
        Next
    Close #F

    Exit Sub
errHandler:
    HandleError "LoadPlayer", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SavePlayerPokemon(ByVal index As Long)
Dim FileName As String
Dim F As Long
Dim x As Byte
Dim i As Long, y As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    FileName = App.Path & "\bin\players\" & Trim$(Player(index).Username) & "_pokemon.dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        For x = 1 To MAX_PLAYER_DATA
            With Player(index).PlayerData(x)
                For i = 1 To MAX_POKEMON
                    Put #F, , .Pokemon(i).Num
                    Put #F, , .Pokemon(i).Gender
                    Put #F, , .Pokemon(i).Level
                    Put #F, , .Pokemon(i).CurHP
                    
                    For y = 1 To Stats.Stat_Count - 1
                        Put #F, , .Pokemon(i).Stat(y)
                        Put #F, , .Pokemon(i).StatEV(y)
                        Put #F, , .Pokemon(i).StatIV(y)
                    Next y
                    
                    Put #F, , .Pokemon(i).Exp
                    
                    For y = 1 To MAX_POKEMON_MOVES
                        Put #F, , .Pokemon(i).Moves(y).Num
                        Put #F, , .Pokemon(i).Moves(y).PP
                        Put #F, , .Pokemon(i).Moves(y).MaxPP
                    Next y
                Next
                For i = 1 To MAX_STORAGE_POKEMON
                    Put #F, , .StoredPokemon(i).Num
                    Put #F, , .StoredPokemon(i).Gender
                    Put #F, , .StoredPokemon(i).Level
                    Put #F, , .StoredPokemon(i).CurHP
                    
                    For y = 1 To Stats.Stat_Count - 1
                        Put #F, , .StoredPokemon(i).Stat(y)
                        Put #F, , .StoredPokemon(i).StatEV(y)
                        Put #F, , .StoredPokemon(i).StatIV(y)
                    Next y
                    
                    Put #F, , .StoredPokemon(i).Exp
                    
                    For y = 1 To MAX_POKEMON_MOVES
                        Put #F, , .StoredPokemon(i).Moves(y).Num
                        Put #F, , .StoredPokemon(i).Moves(y).PP
                        Put #F, , .StoredPokemon(i).Moves(y).MaxPP
                    Next y
                Next
            End With
        Next
    Close #F
    
    Exit Sub
errHandler:
    HandleError "SavePlayerPokemon", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadPlayerPokemon(ByVal index As Long, ByVal User As String)
Dim FileName As String
Dim F As Long
Dim x As Byte
Dim i As Long, y As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    FileName = App.Path & "\bin\players\" & Trim$(User) & "_pokemon.dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        For x = 1 To MAX_PLAYER_DATA
            With Player(index).PlayerData(x)
                For i = 1 To MAX_POKEMON
                    Get #F, , .Pokemon(i).Num
                    Get #F, , .Pokemon(i).Gender
                    Get #F, , .Pokemon(i).Level
                    Get #F, , .Pokemon(i).CurHP
                    
                    For y = 1 To Stats.Stat_Count - 1
                        Get #F, , .Pokemon(i).Stat(y)
                        Get #F, , .Pokemon(i).StatEV(y)
                        Get #F, , .Pokemon(i).StatIV(y)
                    Next y
                    
                    Put #F, , .Pokemon(i).Exp
                    
                    For y = 1 To MAX_POKEMON_MOVES
                        Get #F, , .Pokemon(i).Moves(y).Num
                        Get #F, , .Pokemon(i).Moves(y).PP
                        Get #F, , .Pokemon(i).Moves(y).MaxPP
                    Next y
                Next
                For i = 1 To MAX_STORAGE_POKEMON
                    Get #F, , .StoredPokemon(i).Num
                    Get #F, , .StoredPokemon(i).Gender
                    Get #F, , .StoredPokemon(i).Level
                    Get #F, , .StoredPokemon(i).CurHP
                    
                    For y = 1 To Stats.Stat_Count - 1
                        Get #F, , .StoredPokemon(i).Stat(y)
                        Get #F, , .StoredPokemon(i).StatEV(y)
                        Get #F, , .StoredPokemon(i).StatIV(y)
                    Next y
                    
                    Put #F, , .StoredPokemon(i).Exp
                    
                    For y = 1 To MAX_POKEMON_MOVES
                        Get #F, , .StoredPokemon(i).Moves(y).Num
                        Get #F, , .StoredPokemon(i).Moves(y).PP
                        Get #F, , .StoredPokemon(i).Moves(y).MaxPP
                    Next y
                Next
            End With
        Next
    Close #F

    Exit Sub
errHandler:
    HandleError "LoadPlayerPokemon", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function AccountExist(ByVal Name As String) As Boolean
Dim FileName As String

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    FileName = App.Path & "\bin\players\" & Trim(Name) & ".dat"
    If FileExist(FileName) Then AccountExist = True

    Exit Function
errHandler:
    HandleError "AccountExist", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function isPasswordOK(ByVal User As String, ByVal Pass As String) As Boolean
Dim FileName As String, F As Long
Dim Pass2 As String * MAX_STRING

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If AccountExist(User) Then
        FileName = App.Path & "\bin\players\" & Trim$(User) & ".dat"
        F = FreeFile
        
        Open FileName For Binary As #F
            Get #F, MAX_STRING, Pass2
        Close #F

        If UCase$(Trim$(Pass)) = UCase$(Trim$(Pass2)) Then isPasswordOK = True
    End If

    Exit Function
errHandler:
    HandleError "isPasswordOK", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function FindChar(ByVal Name As String) As Boolean
Dim FileName As String, F As Long
Dim s As String

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    FileName = App.Path & "\bin\players\namelist.txt"
    F = FreeFile
    
    If Not FileExist(FileName) Then
        Open FileName For Output As #F
        Close #F
    End If
    
    Open FileName For Input As #F
        Do While Not EOF(F)
            Input #F, s
                If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
                FindChar = True
            Close #F
                Exit Function
            End If
        Loop
    Close #F
    
    Exit Function
errHandler:
    HandleError "FindChar", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub DeleteName(ByVal Name As String)
Dim F1 As Long, F2 As Long
Dim s As String
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call FileCopy(App.Path & "\bin\players\namelist.txt", App.Path & "\bin\players\namelisttemp.txt")

    F1 = FreeFile
    Open App.Path & "\bin\players\namelisttemp.txt" For Input As #F1
        F2 = FreeFile
        Open App.Path & "\bin\players\namelist.txt" For Output As #F2
            Do While Not EOF(F1)
                Input #F1, s
        
                If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
                    Print #F2, s
                End If
            Loop
        Close #F1
    Close #F2
    Call Kill(App.Path & "\bin\players\namelisttemp.txt")
    
    Exit Sub
errHandler:
    HandleError "DeleteName", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadMap(ByVal MapNum As Long)
Dim FileName As String
Dim F As Long
Dim x As Long, y As Long, i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    FileName = App.Path & "\bin\maps\" & MapNum & "_map.dat"
    F = FreeFile
    
    If Not FileExist(FileName) Then
        ClearMap MapNum
        SaveMap MapNum
        Exit Sub
    End If
    
    Open FileName For Binary As #F
        Get #F, , Map(MapNum).Name
        Get #F, , Map(MapNum).Music
        
        Get #F, , Map(MapNum).Rev
        
        Get #F, , Map(MapNum).Moral
        
        Get #F, , Map(MapNum).MaxX
        Get #F, , Map(MapNum).MaxY
        
        ReDim Map(MapNum).Tile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
        For x = 0 To Map(MapNum).MaxX
            For y = 0 To Map(MapNum).MaxY
                With Map(MapNum).Tile(x, y)
                    For i = 0 To Layers.LayerCount - 1
                        Get #F, , .Layer(i).Tileset
                        Get #F, , .Layer(i).x
                        Get #F, , .Layer(i).y
                    Next i
                    Get #F, , .Type
                    Get #F, , .Data1
                    Get #F, , .Data2
                    Get #F, , .Data3
                    Get #F, , .Data4
                End With
            Next
        Next
        
        For i = 0 To 3
            Get #F, , Map(MapNum).Link(i)
        Next i
        
        For i = 1 To MAX_MAP_POKEMON
            Get #F, , Map(MapNum).Pokemon(i)
        Next i
        
        Get #F, , Map(MapNum).MinLvl
        Get #F, , Map(MapNum).MaxLvl
        
        For i = 1 To MAX_MAP_NPC
            Get #F, , Map(MapNum).Npc(i)
        Next i
        
        Get #F, , Map(MapNum).CurField
        Get #F, , Map(MapNum).CurBack
    Close #F
    DoEvents
    
    Exit Sub
errHandler:
    HandleError "LoadMap", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadMaps()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Count_Map <= 0 Then Exit Sub
    
    For i = 1 To Count_Map
        Call LoadMap(i)
    Next
    
    Exit Sub
errHandler:
    HandleError "LoadMaps", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearMap(ByVal MapNum As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call ZeroMemory(ByVal VarPtr(Map(MapNum)), LenB(Map(MapNum)))
    Map(MapNum).Name = vbNullString
    Map(MapNum).MaxX = Max_MapX
    Map(MapNum).MaxY = Max_MapY
    ReDim Map(MapNum).Tile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
    
    Exit Sub
errHandler:
    HandleError "ClearMap", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearMaps()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_Map
        Call ClearMap(i)
    Next
    
    Exit Sub
errHandler:
    HandleError "ClearMaps", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveMap(ByVal MapNum As Long)
Dim FileName As String
Dim F As Long
Dim x As Long, y As Long, i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    FileName = App.Path & "\bin\maps\" & MapNum & "_map.dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Put #F, , Map(MapNum).Name
        Put #F, , Map(MapNum).Music
        
        Put #F, , Map(MapNum).Rev
        
        Put #F, , Map(MapNum).Moral
        
        Put #F, , Map(MapNum).MaxX
        Put #F, , Map(MapNum).MaxY
        
        For x = 0 To Map(MapNum).MaxX
            For y = 0 To Map(MapNum).MaxY
                With Map(MapNum).Tile(x, y)
                    For i = 0 To Layers.LayerCount - 1
                        Put #F, , .Layer(i).Tileset
                        Put #F, , .Layer(i).x
                        Put #F, , .Layer(i).y
                    Next i
                    Put #F, , .Type
                    Put #F, , .Data1
                    Put #F, , .Data2
                    Put #F, , .Data3
                    Put #F, , .Data4
                End With
            Next
        Next
        
        For i = 0 To 3
            Put #F, , Map(MapNum).Link(i)
        Next i
        
        For i = 1 To MAX_MAP_POKEMON
            Put #F, , Map(MapNum).Pokemon(i)
        Next i
        
        Put #F, , Map(MapNum).MinLvl
        Put #F, , Map(MapNum).MaxLvl
        
        For i = 1 To MAX_MAP_NPC
            Put #F, , Map(MapNum).Npc(i)
        Next i
        
        Put #F, , Map(MapNum).CurField
        Put #F, , Map(MapNum).CurBack
    Close #F
    DoEvents
    
    Exit Sub
errHandler:
    HandleError "SaveMap", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SavePokemon(ByVal PokemonNum As Long)
Dim FileName As String
Dim F As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    FileName = App.Path & "\bin\Pokemons\" & PokemonNum & "_pokemon.dat"
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Pokemon(PokemonNum)
    Close #F
    
    Exit Sub
errHandler:
    HandleError "SavePokemon", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SavePokemons()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_Pokemon
        Call SavePokemon(i)
    Next
    
    Exit Sub
errHandler:
    HandleError "SavePokemons", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadPokemon(ByVal PokemonNum As Long)
Dim FileName As String
Dim i As Long
Dim F As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    FileName = App.Path & "\bin\Pokemons\" & PokemonNum & "_pokemon.dat"

    If Not FileExist(FileName) Then
        Call ClearPokemon(PokemonNum)
        Call SavePokemon(PokemonNum)
        Exit Sub
    End If

    F = FreeFile
    Open FileName For Binary As #F
        Get #F, , Pokemon(PokemonNum)
    Close #F
    
    Exit Sub
errHandler:
    HandleError "LoadPokemon", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadPokemons()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_Pokemon
        Call LoadPokemon(i)
    Next
    
    Exit Sub
errHandler:
    HandleError "LoadPokemons", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearPokemon(ByVal index As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call ZeroMemory(ByVal VarPtr(Pokemon(index)), LenB(Pokemon(index)))
    Pokemon(index).Name = vbNullString
    
    Exit Sub
errHandler:
    HandleError "ClearPokemon", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearPokemons()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_Pokemon
        Call ClearPokemon(i)
    Next
    
    Exit Sub
errHandler:
    HandleError "ClearPokemons", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearEnemyPokemon(ByVal index As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Call ZeroMemory(ByVal VarPtr(TempPlayer(index).EnemyPokemon), LenB(TempPlayer(index).EnemyPokemon))
    
    Exit Sub
errHandler:
    HandleError "ClearEnemyPokemon", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveMove(ByVal MoveNum As Long)
Dim FileName As String
Dim F As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    FileName = App.Path & "\bin\Moves\Move" & MoveNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Moves(MoveNum)
    Close #F
    
    Exit Sub
errHandler:
    HandleError "SaveMove", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveMoves()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_Move
        Call SaveMove(i)
    Next
    
    Exit Sub
errHandler:
    HandleError "SaveMoves", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadMove(ByVal MoveNum As Long)
Dim FileName As String
Dim i As Long
Dim F As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    FileName = App.Path & "\bin\Moves\Move" & MoveNum & ".dat"

    If Not FileExist(FileName) Then
        Call ClearMove(MoveNum)
        Call SaveMove(MoveNum)
        Exit Sub
    End If

    F = FreeFile
    Open FileName For Binary As #F
        Get #F, , Moves(MoveNum)
    Close #F
    
    Exit Sub
errHandler:
    HandleError "LoadMove", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadMoves()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_Move
        Call LoadMove(i)
    Next
    
    Exit Sub
errHandler:
    HandleError "LoadMoves", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearMove(ByVal index As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call ZeroMemory(ByVal VarPtr(Moves(index)), LenB(Moves(index)))
    Moves(index).Name = vbNullString
    
    Exit Sub
errHandler:
    HandleError "ClearMove", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearMoves()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_Move
        Call ClearMove(i)
    Next
    
    Exit Sub
errHandler:
    HandleError "ClearMoves", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadExpCalc()
Dim FileName As String, i As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    FileName = App.Path & "\bin\expcalc.ini"
    
    If Not FileExist(FileName) Then
        ClearExpCalc
        SaveExpCalc
        Exit Sub
    End If
    
    For i = 1 To MAX_LEVEL
        ExpCalc(i) = Val(GetVar(FileName, "Experience", "Level" & i))
    Next
    
    Exit Sub
errHandler:
    HandleError "LoadExpCalc", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveExpCalc()
Dim FileName As String, i As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    FileName = App.Path & "\bin\expcalc.ini"
    
    For i = 1 To MAX_LEVEL
        PutVar FileName, "Experience", "Level" & i, Str(ExpCalc(i))
    Next
    
    Exit Sub
errHandler:
    HandleError "SaveExpCalc", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearExpCalc()
Dim i As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To MAX_LEVEL
        ExpCalc(i) = 5 * i
    Next
    
    Exit Sub
errHandler:
    HandleError "ClearExpCalc", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveItem(ByVal ItemNum As Long)
Dim FileName As String
Dim F As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    FileName = App.Path & "\bin\Items\Item" & ItemNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Item(ItemNum)
    Close #F
    
    Exit Sub
errHandler:
    HandleError "SaveItem", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveItems()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_Item
        Call SaveItem(i)
    Next
    
    Exit Sub
errHandler:
    HandleError "SaveItems", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadItem(ByVal ItemNum As Long)
Dim FileName As String
Dim i As Long
Dim F As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    FileName = App.Path & "\bin\Items\Item" & ItemNum & ".dat"

    If Not FileExist(FileName) Then
        Call ClearItem(ItemNum)
        Call SaveItem(ItemNum)
        Exit Sub
    End If

    F = FreeFile
    Open FileName For Binary As #F
        Get #F, , Item(ItemNum)
    Close #F
    
    Exit Sub
errHandler:
    HandleError "LoadItem", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadItems()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_Item
        Call LoadItem(i)
    Next
    
    Exit Sub
errHandler:
    HandleError "LoadItems", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearItem(ByVal index As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call ZeroMemory(ByVal VarPtr(Item(index)), LenB(Item(index)))
    Item(index).Name = vbNullString
    Item(index).Type = ItemType.Items
    Item(index).Desc = vbNullString
    
    Exit Sub
errHandler:
    HandleError "ClearItem", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearItems()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_Item
        Call ClearItem(i)
    Next
    
    Exit Sub
errHandler:
    HandleError "ClearItems", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveNPC(ByVal NpcNum As Long)
Dim FileName As String
Dim F As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    FileName = App.Path & "\bin\NPCs\NPC" & NpcNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Npc(NpcNum)
    Close #F
    
    Exit Sub
errHandler:
    HandleError "SaveNPC", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveNPCs()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_NPC
        Call SaveNPC(i)
    Next
    
    Exit Sub
errHandler:
    HandleError "SaveNPCs", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadNPC(ByVal NpcNum As Long)
Dim FileName As String
Dim i As Long
Dim F As Long
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    FileName = App.Path & "\bin\NPCs\NPC" & NpcNum & ".dat"

    If Not FileExist(FileName) Then
        Call ClearNPC(NpcNum)
        Call SaveNPC(NpcNum)
        Exit Sub
    End If

    F = FreeFile
    Open FileName For Binary As #F
        Get #F, , Npc(NpcNum)
    Close #F
    
    Exit Sub
errHandler:
    HandleError "LoadNPC", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadNPCs()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_NPC
        Call LoadNPC(i)
    Next
    
    Exit Sub
errHandler:
    HandleError "LoadNPCs", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearNPC(ByVal index As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call ZeroMemory(ByVal VarPtr(Npc(index)), LenB(Npc(index)))
    Npc(index).Name = vbNullString
    
    Exit Sub
errHandler:
    HandleError "ClearNPC", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearNPCs()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_NPC
        Call ClearNPC(i)
    Next
    
    Exit Sub
errHandler:
    HandleError "ClearNPCs", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearMapNpc(ByVal MapNum As Long, ByVal NpcNum As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call ZeroMemory(ByVal VarPtr(MapNpc(MapNum).Npc(NpcNum)), LenB(MapNpc(MapNum).Npc(NpcNum)))
    
    Exit Sub
errHandler:
    HandleError "ClearMapNpc", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearMapNpcs()
Dim x As Long, y As Long
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For x = 1 To Count_Map
        For y = 1 To MAX_MAP_NPC
            Call ClearMapNpc(x, y)
        Next
    Next
    
    Exit Sub
errHandler:
    HandleError "ClearMapNpcs", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveShop(ByVal ShopNum As Long)
Dim FileName As String
Dim F As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    FileName = App.Path & "\bin\Shops\Shop" & ShopNum & ".dat"
    F = FreeFile
    Open FileName For Binary As #F
        Put #F, , Shop(ShopNum)
    Close #F
    
    Exit Sub
errHandler:
    HandleError "SaveShop", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveShops()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_Shop
        Call SaveShop(i)
    Next
    
    Exit Sub
errHandler:
    HandleError "SaveShops", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadShop(ByVal ShopNum As Long)
Dim FileName As String
Dim i As Long
Dim F As Long
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    FileName = App.Path & "\bin\Shops\Shop" & ShopNum & ".dat"

    If Not FileExist(FileName) Then
        Call ClearShop(ShopNum)
        Call SaveShop(ShopNum)
        Exit Sub
    End If

    F = FreeFile
    Open FileName For Binary As #F
        Get #F, , Shop(ShopNum)
    Close #F
    
    Exit Sub
errHandler:
    HandleError "LoadShop", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadShops()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_Shop
        Call LoadShop(i)
    Next
    
    Exit Sub
errHandler:
    HandleError "LoadShops", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearShop(ByVal index As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call ZeroMemory(ByVal VarPtr(Shop(index)), LenB(Shop(index)))
    Shop(index).Name = vbNullString
    
    Exit Sub
errHandler:
    HandleError "ClearShop", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearShops()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To Count_Shop
        Call ClearShop(i)
    Next
    
    Exit Sub
errHandler:
    HandleError "ClearShops", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

