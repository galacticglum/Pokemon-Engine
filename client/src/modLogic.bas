Attribute VB_Name = "modLogic"
Option Explicit

Public Sub ProcessNpcMovement(ByVal MapNpcNum As Long)
Dim MovementSpeed As Long

    Select Case MapNpc(MapNpcNum).Moving
        Case MOVING_WALKING: MovementSpeed = WALK_SPEED
        Case Else: Exit Sub
    End Select
    
    Select Case MapNpc(MapNpcNum).Dir
        Case DIR_UP
            MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset - MovementSpeed
            If MapNpc(MapNpcNum).yOffset < 0 Then MapNpc(MapNpcNum).yOffset = 0
        Case DIR_DOWN
            MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset + MovementSpeed
            If MapNpc(MapNpcNum).yOffset > 0 Then MapNpc(MapNpcNum).yOffset = 0
        Case DIR_LEFT
            MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset - MovementSpeed
            If MapNpc(MapNpcNum).xOffset < 0 Then MapNpc(MapNpcNum).xOffset = 0
        Case DIR_RIGHT
            MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset + MovementSpeed
            If MapNpc(MapNpcNum).xOffset > 0 Then MapNpc(MapNpcNum).xOffset = 0
    End Select
    
    If MapNpc(MapNpcNum).Moving > 0 Then
        If MapNpc(MapNpcNum).Dir = DIR_RIGHT Or MapNpc(MapNpcNum).Dir = DIR_DOWN Then
            If (MapNpc(MapNpcNum).xOffset >= 0) And (MapNpc(MapNpcNum).yOffset >= 0) Then
                MapNpc(MapNpcNum).Moving = 0
                If MapNpc(MapNpcNum).Step = 0 Then
                    MapNpc(MapNpcNum).Step = 2
                Else
                    MapNpc(MapNpcNum).Step = 0
                End If
            End If
        Else
            If (MapNpc(MapNpcNum).xOffset <= 0) And (MapNpc(MapNpcNum).yOffset <= 0) Then
                MapNpc(MapNpcNum).Moving = 0
                If MapNpc(MapNpcNum).Step = 0 Then
                    MapNpc(MapNpcNum).Step = 2
                Else
                    MapNpc(MapNpcNum).Step = 0
                End If
            End If
        End If
    End If
End Sub

Public Sub AddTrade(ByVal TradeType As Byte, Optional ByVal PokeSlot As Long = 0, Optional ByVal ItemSlot As Long = 0, Optional ByVal ItemVal As Long = 0)
Dim i As Long
Dim ItemNum As Long

    i = CheckFreeTradeSlot
    
    If Not TradeExist(TradeType, PokeSlot, ItemSlot, CurInvType) Then
        If i > 0 Then
            With MyTrade(i)
                .Type = TradeType
                If .Type = TRADE_TYPE_ITEM Then
                    ItemNum = GetInvItemNum(ItemSlot, CurInvType)
                    If ItemNum > 0 And ItemNum < Count_Item Then
                        .ItemNum = ItemNum
                        .ItemVal = ItemVal
                        
                        .TempItemSlot = ItemSlot
                        .TempItemType = CurInvType
                    End If
                ElseIf .Type = TRADE_TYPE_POKEMON Then
                    If PokeSlot > 0 And PokeSlot < MAX_POKEMON Then
                        .Pokemon = Player(MyIndex).Pokemon(PokeSlot)
                        
                        .TempPokeSlot = PokeSlot
                    End If
                End If
            End With
            UpdateTrade
        End If
    End If
End Sub

Public Function CheckFreeTradeSlot() As Long
Dim i As Long
    
    CheckFreeTradeSlot = 0
    For i = 1 To MAX_TRADE
        If MyTrade(i).Type = 0 Then CheckFreeTradeSlot = i
    Next
End Function

Public Sub ClearTradeSlot(ByVal TradeSlot As Long)
    Call ZeroMemory(ByVal VarPtr(MyTrade(TradeSlot)), LenB(MyTrade(TradeSlot)))
    UpdateTrade
End Sub

Public Sub ClearTradeSlots()
Dim i As Long

    For i = 1 To MAX_TRADE
        Call ClearTradeSlot(i)
    Next
End Sub

Public Sub UpdateTrade()
Dim i As Long
Dim TmpTrade As TradeRec

    For i = 1 To MAX_TRADE - 1
        If MyTrade(i).Type = 0 Then
            If MyTrade(i + 1).Type > 0 Then
                TmpTrade = MyTrade(i + 1)
                MyTrade(i) = TmpTrade
                ClearTradeSlot (i + 1)
            End If
        End If
    Next
End Sub

Public Function TradeExist(ByVal TradeType As Byte, Optional ByVal PokeSlot As Long = 0, Optional ByVal ItemSlot As Long = 0, Optional ByVal ItemType As Byte = 0) As Boolean
Dim i As Long

    TradeExist = False
    For i = 1 To MAX_TRADE
        With MyTrade(i)
            If .Type = TradeType Then
                If .Type = TRADE_TYPE_ITEM Then
                    If .TempItemSlot = ItemSlot Then
                        If .TempItemType = ItemType Then TradeExist = True
                    End If
                ElseIf .Type = TRADE_TYPE_POKEMON Then
                    If .TempPokeSlot = PokeSlot Then TradeExist = True
                End If
            End If
        End With
    Next
End Function
