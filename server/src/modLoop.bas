Attribute VB_Name = "modLoop"
Option Explicit

Public Declare Function GetTickCount Lib "Kernel32" () As Long

Public Sub AppLoop()
Dim Tick As Long
Dim LastUpdatePlayers As Long
Dim Tmr35 As Long, Tmr500 As Long, Tmr1000 As Long
Dim i As Long
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Do While AppOpen
        Tick = GetTickCount
        
        If LastUpdatePlayers < Tick Then
            SavePlayersOnline
            LastUpdatePlayers = Tick + 300000
        End If
        
        If Tmr500 < Tick Then
            UpdateMapLogic
            
            Tmr500 = Tick + 500
        End If
        
        If Tmr1000 < Tick Then
            For i = 1 To HighPlayerIndex
                If IsPlaying(i) Then
                    If TempPlayer(i).InTradeRequest > 0 Then
                        TempPlayer(i).InTradeReqCount = TempPlayer(i).InTradeReqCount - 1
                        If TempPlayer(i).InTradeReqCount <= 0 Then
                            TempPlayer(i).InTradeReqCount = 0
                            TempPlayer(i).InTradeRequest = 0
                        End If
                    End If
                End If
            Next
            
            Tmr1000 = Tick + 1000
        End If
        
        If Tmr35 < Tick Then
            For i = 1 To HighPlayerIndex
                If IsPlaying(i) Then
                    BattleLoop i
                End If
            Next
            
            Tmr35 = Tick + 35
        End If
        
        DoEvents
    Loop

    Exit Sub
errHandler:
    HandleError "AppLoop", "modLoop", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub BattleLoop(ByVal index As Long)
Dim FoeIndex As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If TempPlayer(index).Checked = YES Then Exit Sub
    
    If TempPlayer(index).InBattle = BATTLE_TRAINER Then
        FoeIndex = TempPlayer(index).BattleRequest
        If TempPlayer(index).MoveSet > 0 Then
            If TempPlayer(FoeIndex).MoveSet > 0 Then
                InitBattleVsPlayer index, FoeIndex, TempPlayer(index).MoveSet, TempPlayer(FoeIndex).MoveSet
                TempPlayer(FoeIndex).Checked = YES
            End If
        End If
    End If

    Exit Sub
errHandler:
    HandleError "BattleLoop", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub
