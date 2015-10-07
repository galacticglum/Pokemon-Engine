Attribute VB_Name = "modDatabase"
Option Explicit

Private Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc)
Dim FileName As String

    MsgBox "Run-time error '" & erNumber & "': " & erDesc & ".", vbCritical, GameTitle

    FileName = App.Path & "\bin\report\errors.txt"
    Open FileName For Append As #1
        Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
        Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
        Print #1, ""
    Close #1
    
    CloseApp
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

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
    
    Exit Sub
errHandler:
    HandleError "ChkDir", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

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

Public Sub PutVar(file As String, Header As String, Var As String, value As String)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call WritePrivateProfileString$(Header, Var, value, file)
    
    Exit Sub
errHandler:
    HandleError "PutVar", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearMap()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call ZeroMemory(ByVal VarPtr(Map), LenB(Map))
    Map.Name = vbNullString
    Map.MaxX = Max_MapX
    Map.MaxY = Max_MapY
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
    
    Exit Sub
errHandler:
    HandleError "ClearMap", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadMap(ByVal MapNum As Long)
Dim FileName As String
Dim F As Long
Dim X As Long, Y As Long, i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    FileName = App.Path & "\bin\_maps\" & MapNum & "_cache.dat"
    F = FreeFile
    
    If Not FileExist(FileName) Then Exit Sub
    
    Open FileName For Binary As #F
        Get #F, , Map.Name
        Get #F, , Map.Music
         
        Get #F, , Map.Rev
        
        Get #F, , Map.Moral
        
        Get #F, , Map.MaxX
        Get #F, , Map.MaxY
        
        ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)
        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                With Map.Tile(X, Y)
                    For i = 0 To Layers.LayerCount - 1
                        Get #F, , .Layer(i).Tileset
                        Get #F, , .Layer(i).X
                        Get #F, , .Layer(i).Y
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
            Get #F, , Map.Link(i)
        Next i
        
        For i = 1 To MAX_MAP_POKEMON
            Get #F, , Map.Pokemon(i)
        Next i
        
        Get #F, , Map.MinLvl
        Get #F, , Map.MaxLvl
        
        For i = 1 To MAX_MAP_NPC
            Get #F, , Map.NPC(i)
        Next i
        
        Get #F, , Map.CurField
        Get #F, , Map.CurBack
    Close #F
    DoEvents
    
    Exit Sub
errHandler:
    HandleError "LoadMap", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveMap(ByVal MapNum As Long)
Dim FileName As String
Dim F As Long
Dim X As Long, Y As Long, i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    FileName = App.Path & "\bin\_maps\" & MapNum & "_cache.dat"
    F = FreeFile
    
    Open FileName For Binary As #F
        Put #F, , Map.Name
        Put #F, , Map.Music
         
        Put #F, , Map.Rev
        
        Put #F, , Map.Moral
        
        Put #F, , Map.MaxX
        Put #F, , Map.MaxY

        For X = 0 To Map.MaxX
            For Y = 0 To Map.MaxY
                With Map.Tile(X, Y)
                    For i = 0 To Layers.LayerCount - 1
                        Put #F, , .Layer(i).Tileset
                        Put #F, , .Layer(i).X
                        Put #F, , .Layer(i).Y
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
            Put #F, , Map.Link(i)
        Next i
        
        For i = 1 To MAX_MAP_POKEMON
            Put #F, , Map.Pokemon(i)
        Next i
        
        Put #F, , Map.MinLvl
        Put #F, , Map.MaxLvl
        
        For i = 1 To MAX_MAP_NPC
            Put #F, , Map.NPC(i)
        Next i
        
        Put #F, , Map.CurField
        Put #F, , Map.CurBack
    Close #F
    DoEvents
    
    Exit Sub
errHandler:
    HandleError "SaveMap", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearPlayer(ByVal Index As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Player(Index).Name = vbNullString
    
    Exit Sub
errHandler:
    HandleError "ClearPlayer", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function IsValidMapPoint(ByVal X As Long, ByVal Y As Long) As Boolean
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    IsValidMapPoint = False
    If X < 0 Then Exit Function
    If Y < 0 Then Exit Function
    If X > Map.MaxX Then Exit Function
    If Y > Map.MaxY Then Exit Function
    IsValidMapPoint = True
    
    Exit Function
errHandler:
    HandleError "IsValidMapPoint", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function ConvertMapX(ByVal X As Long) As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ConvertMapX = X - (TileView.Left * Pic_Size) - Camera.Left
    
    Exit Function
errHandler:
    HandleError "ConvertMapX", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function ConvertMapY(ByVal Y As Long) As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ConvertMapY = Y - (TileView.top * Pic_Size) - Camera.top
    
    Exit Function
errHandler:
    HandleError "ConvertMapY", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub SaveOption()
Dim FileName As String

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    FileName = App.Path & "\bin\config.ini"
    
    PutVar FileName, "Account", "Username", Trim$(Options.Username)
    PutVar FileName, "Account", "Password", Trim$(Options.Password)
    PutVar FileName, "Account", "SavePass", Val(Options.SavePass)
    
    PutVar FileName, "Settings", "Game IP", Trim$(Options.SaveIp)
    PutVar FileName, "Settings", "Game Port", Val(Options.SavePort)
    
    PutVar FileName, "Audio", "Music", Val(Options.Music)
    PutVar FileName, "Audio", "Sound", Val(Options.Sound)
    
    Exit Sub
errHandler:
    HandleError "SaveOption", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadOption()
Dim FileName As String

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    FileName = App.Path & "\bin\config.ini"
    
    If Not FileExist(FileName) Then
        ClearOption
        SaveOption
        SaveAccount = False
        Exit Sub
    End If
    
    Options.Username = Trim$(GetVar(FileName, "Account", "Username"))
    Options.Password = Trim$(GetVar(FileName, "Account", "Password"))
    Options.SavePass = Val(GetVar(FileName, "Account", "SavePass"))
    
    Options.SaveIp = Trim$(GetVar(FileName, "Settings", "Game IP"))
    Options.SavePort = Val(GetVar(FileName, "Settings", "Game Port"))
    
    Options.Music = Val(GetVar(FileName, "Audio", "Music"))
    Options.Sound = Val(GetVar(FileName, "Audio", "Sound"))
    
    If Options.SavePass > 0 Then
        SaveAccount = True
    Else
        SaveAccount = False
    End If
    
    Exit Sub
errHandler:
    HandleError "LoadOption", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearOption()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Options.Username = vbNullString
    Options.Password = vbNullString
    Options.SavePass = 0
    Options.SaveIp = "localhost"
    Options.SavePort = 7001
    Options.Music = 0
    Options.Sound = 0

    Exit Sub
errHandler:
    HandleError "LoadOption", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearPokemon(ByVal Index As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call ZeroMemory(ByVal VarPtr(Pokemon(Index)), LenB(Pokemon(Index)))
    Pokemon(Index).Name = vbNullString
    
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

Public Sub ClearEnemyPokemon()
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    Call ZeroMemory(ByVal VarPtr(EnemyPokemon), LenB(EnemyPokemon))
    
    Exit Sub
errHandler:
    HandleError "ClearEnemyPokemon", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearMove(ByVal Index As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call ZeroMemory(ByVal VarPtr(Moves(Index)), LenB(Moves(Index)))
    Moves(Index).Name = vbNullString
    
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

Public Sub ClearItem(ByVal Index As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Item(Index).Name = vbNullString
    Item(Index).Type = ItemType.Items
    Item(Index).Desc = vbNullString
    
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

Public Sub ClearNPC(ByVal Index As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call ZeroMemory(ByVal VarPtr(NPC(Index)), LenB(NPC(Index)))
    NPC(Index).Name = vbNullString
    
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

Public Sub ClearMapNpc(ByVal Index As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call ZeroMemory(ByVal VarPtr(MapNpc(Index)), LenB(MapNpc(Index)))
    
    Exit Sub
errHandler:
    HandleError "ClearMapNpc", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearMapNpcs()
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For i = 1 To MAX_MAP_NPC
        Call ClearMapNpc(i)
    Next
    
    Exit Sub
errHandler:
    HandleError "ClearMapNpcs", "modDatabase", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ClearShop(ByVal Index As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    Shop(Index).Name = vbNullString
    
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
