Attribute VB_Name = "modLogic"
Option Explicit

Public Sub UpdateMapLogic()
Dim x As Long, MapNum As Long, NpcNum As Long
Dim i As Long
Dim TickCount As Long

    For MapNum = 1 To Count_Map
        If CheckPlayerOnMap(MapNum) > 0 Then
            TickCount = GetTickCount
            
            For x = 1 To MAX_MAP_NPC
                NpcNum = MapNpc(MapNum).Npc(x).Num
                
                If Map(MapNum).Npc(x) > 0 And MapNpc(MapNum).Npc(x).Num > 0 Then
                    i = Int(Rnd * 2)
    
                    If i = 1 Then
                        i = Int(Rnd * 4)

                        If CanNpcMove(MapNum, x, i) Then
                            Call NpcMove(MapNum, x, i, MOVING_WALKING)
                        End If
                    End If
                End If
                
                If MapNpc(MapNum).Npc(x).Num = 0 And Map(MapNum).Npc(x) > 0 Then
                    Call SpawnNpc(x, MapNum)
                End If
            Next
        End If
        DoEvents
    Next
End Sub

Public Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long)
Dim NpcNum As Long
Dim i As Long
Dim x As Long, y As Long
Dim Spawned As Boolean

    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPC Or MapNum <= 0 Or MapNum > Count_Map Then Exit Sub
    NpcNum = Map(MapNum).Npc(MapNpcNum)

    If NpcNum > 0 Then
        MapNpc(MapNum).Npc(MapNpcNum).Num = NpcNum
        MapNpc(MapNum).Npc(MapNpcNum).Dir = Int(Rnd * 4)
        
        'For x = 0 To Max_MapX
        '    For y = 0 To Max_MapY
        '        If Map(MapNum).Tile(x, y).Attribute = Attributes.NpcSpawn Then
        '            If Map(MapNum).Tile(x, y).Data1 = MapNpcNum Then
        '                MapNpc(MapNum).Npc(MapNpcNum).x = x
        '                MapNpc(MapNum).Npc(MapNpcNum).y = y
        '                MapNpc(MapNum).Npc(MapNpcNum).Dir = Map(MapNum).Tile(x, y).Data2
        '                Spawned = True
        '                Exit For
        '            End If
        '        End If
        '    Next y
        'Next x
        
        If Not Spawned Then
            For i = 1 To 100
                x = Random(0, Max_MapX)
                y = Random(0, Max_MapY)
    
                If x > Max_MapX Then x = Max_MapX
                If y > Max_MapY Then y = Max_MapY
    
                If NpcTileIsOpen(MapNum, x, y) Then
                    MapNpc(MapNum).Npc(MapNpcNum).x = x
                    MapNpc(MapNum).Npc(MapNpcNum).y = y
                    Spawned = True
                    Exit For
                End If
            Next
        End If

        If Not Spawned Then
            For x = 0 To Max_MapX
                For y = 0 To Max_MapY
                    If NpcTileIsOpen(MapNum, x, y) Then
                        MapNpc(MapNum).Npc(MapNpcNum).x = x
                        MapNpc(MapNum).Npc(MapNpcNum).y = y
                        Spawned = True
                    End If
                Next
            Next
        End If

        If Spawned Then
            SendSpawnNpc MapNum, MapNpcNum
        End If
    Else
        MapNpc(MapNum).Npc(MapNpcNum).Num = 0
        SendClearNPC MapNum, MapNpcNum
    End If
End Sub

Public Sub SpawnMapNpcs(ByVal MapNum As Long)
Dim i As Long

    For i = 1 To MAX_MAP_NPC
        Call SpawnNpc(i, MapNum)
    Next
End Sub

Public Sub SpawnAllMapNpcs()
Dim i As Long

    For i = 1 To Count_Map
        Call SpawnMapNpcs(i)
    Next
End Sub

Public Function NpcTileIsOpen(ByVal MapNum As Long, ByVal x As Long, ByVal y As Long) As Boolean
Dim i As Long

    NpcTileIsOpen = True

    For i = 1 To MAX_PLAYER
        If TempPlayer(i).CurSlot > 0 Then
            If IsPlaying(i) And Player(i).PlayerData(TempPlayer(i).CurSlot).Map = MapNum Then
                If Player(i).PlayerData(TempPlayer(i).CurSlot).x = x Then
                    If Player(i).PlayerData(TempPlayer(i).CurSlot).y = y Then
                        NpcTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If
        End If
    Next

    For i = 1 To MAX_MAP_NPC
        If MapNpc(MapNum).Npc(i).Num > 0 Then
            If MapNpc(MapNum).Npc(i).x = x Then
                If MapNpc(MapNum).Npc(i).y = y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If
    Next

    If Map(MapNum).Tile(x, y).Type <> Attributes.Walkable Then
        NpcTileIsOpen = False
        Exit Function
    End If
End Function

Public Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Byte) As Boolean
Dim i As Long
Dim n As Long
Dim x As Long
Dim y As Long

    If MapNum <= 0 Or MapNum > Count_Map Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPC Or Dir < DIR_UP Or Dir > DIR_RIGHT Then Exit Function
    x = MapNpc(MapNum).Npc(MapNpcNum).x
    y = MapNpc(MapNum).Npc(MapNpcNum).y
    
    CanNpcMove = True

    Select Case Dir
        Case DIR_UP
            If y > 0 Then
                n = Map(MapNum).Tile(x, y - 1).Type
                If n <> Attributes.Walkable Then
                    CanNpcMove = False
                    Exit Function
                End If

                For i = 1 To MAX_PLAYER
                    If TempPlayer(i).CurSlot > 0 Then
                        If IsPlaying(i) And Player(i).PlayerData(TempPlayer(i).CurSlot).Map = MapNum Then
                            If (Player(i).PlayerData(TempPlayer(i).CurSlot).x = MapNpc(MapNum).Npc(MapNpcNum).x) And (Player(i).PlayerData(TempPlayer(i).CurSlot).y = MapNpc(MapNum).Npc(MapNpcNum).y - 1) Then
                                CanNpcMove = False
                                Exit Function
                            End If
                        End If
                    End If
                Next
                For i = 1 To MAX_MAP_NPC
                    If (i <> MapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) Then
                        If (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(MapNpcNum).x) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(MapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next
            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN
            If y < Max_MapY Then
                n = Map(MapNum).Tile(x, y + 1).Type
                If n <> Attributes.Walkable Then
                    CanNpcMove = False
                    Exit Function
                End If

                For i = 1 To MAX_PLAYER
                    If TempPlayer(i).CurSlot > 0 Then
                        If IsPlaying(i) And Player(i).PlayerData(TempPlayer(i).CurSlot).Map = MapNum Then
                            If (Player(i).PlayerData(TempPlayer(i).CurSlot).x = MapNpc(MapNum).Npc(MapNpcNum).x) And (Player(i).PlayerData(TempPlayer(i).CurSlot).y = MapNpc(MapNum).Npc(MapNpcNum).y + 1) Then
                                CanNpcMove = False
                                Exit Function
                            End If
                        End If
                    End If
                Next
                For i = 1 To MAX_MAP_NPC
                    If (i <> MapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) Then
                        If (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(MapNpcNum).x) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(MapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next
            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT
            If x > 0 Then
                n = Map(MapNum).Tile(x - 1, y).Type
                If n <> Attributes.Walkable Then
                    CanNpcMove = False
                    Exit Function
                End If

                For i = 1 To MAX_PLAYER
                    If TempPlayer(i).CurSlot > 0 Then
                        If IsPlaying(i) And Player(i).PlayerData(TempPlayer(i).CurSlot).Map = MapNum Then
                            If (Player(i).PlayerData(TempPlayer(i).CurSlot).x = MapNpc(MapNum).Npc(MapNpcNum).x - 1) And (Player(i).PlayerData(TempPlayer(i).CurSlot).y = MapNpc(MapNum).Npc(MapNpcNum).y) Then
                                CanNpcMove = False
                                Exit Function
                            End If
                        End If
                    End If
                Next
                For i = 1 To MAX_MAP_NPC
                    If (i <> MapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) Then
                        If (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(MapNpcNum).x - 1) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(MapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next
            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT
            If x < Max_MapX Then
                n = Map(MapNum).Tile(x + 1, y).Type
                If n <> Attributes.Walkable Then
                    CanNpcMove = False
                    Exit Function
                End If

                For i = 1 To MAX_PLAYER
                    If TempPlayer(i).CurSlot > 0 Then
                        If IsPlaying(i) And Player(i).PlayerData(TempPlayer(i).CurSlot).Map = MapNum Then
                            If (Player(i).PlayerData(TempPlayer(i).CurSlot).x = MapNpc(MapNum).Npc(MapNpcNum).x + 1) And (Player(i).PlayerData(TempPlayer(i).CurSlot).y = MapNpc(MapNum).Npc(MapNpcNum).y) Then
                                CanNpcMove = False
                                Exit Function
                            End If
                        End If
                    End If
                Next
                For i = 1 To MAX_MAP_NPC
                    If (i <> MapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) Then
                        If (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(MapNpcNum).x + 1) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(MapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next
            Else
                CanNpcMove = False
            End If
    End Select
End Function

Public Sub NpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long, ByVal Movement As Long)
    If MapNum <= 0 Or MapNum > Count_Map Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPC Or Dir < DIR_UP Or Dir > DIR_DOWN Then Exit Sub
    If Movement <> MOVING_WALKING Then Exit Sub
    
    MapNpc(MapNum).Npc(MapNpcNum).Dir = Dir
    Select Case Dir
        Case DIR_UP
            MapNpc(MapNum).Npc(MapNpcNum).y = MapNpc(MapNum).Npc(MapNpcNum).y - 1
        Case DIR_DOWN
            MapNpc(MapNum).Npc(MapNpcNum).y = MapNpc(MapNum).Npc(MapNpcNum).y + 1
        Case DIR_LEFT
            MapNpc(MapNum).Npc(MapNpcNum).x = MapNpc(MapNum).Npc(MapNpcNum).x - 1
        Case DIR_RIGHT
            MapNpc(MapNum).Npc(MapNpcNum).x = MapNpc(MapNum).Npc(MapNpcNum).x + 1
    End Select
    SendNpcMove MapNum, MapNpcNum, Movement
End Sub

Public Sub NpcDir(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long)
Dim buffer As clsBuffer

    If MapNum <= 0 Or MapNum > Count_Map Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPC Or Dir < DIR_UP Or Dir > DIR_DOWN Then Exit Sub
    MapNpc(MapNum).Npc(MapNpcNum).Dir = Dir
    
    Set buffer = New clsBuffer
    buffer.WriteLong SNpcDir
    buffer.WriteLong MapNpcNum
    buffer.WriteLong Dir
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

