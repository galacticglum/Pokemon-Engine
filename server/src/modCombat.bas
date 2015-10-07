Attribute VB_Name = "modCombat"
Option Explicit

Private CheckNpcMove(1 To MAX_POKEMON_MOVES) As NpcMove

Private Type NpcMove
    InputNum As Long
End Type

Public Function GetBattleExp(ByVal index As Long) As Long
Dim BattleType As Single, IsExpShare As Single, IsTrade As Single
Dim PokeNum As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    BattleType = 1 ' 1.5 if won on Trainer
    IsExpShare = 1 ' 2 if pokemon have exp.share
    IsTrade = 1 ' 1.5 if is trade
    
    PokeNum = TempPlayer(index).EnemyPokemon.Num
    With Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(TempPlayer(index).InBattlePoke)
        GetBattleExp = ((((BattleType * Pokemon(PokeNum).BaseExp * TempPlayer(index).EnemyPokemon.Level) / (5 * IsExpShare)) * (((2 * TempPlayer(index).EnemyPokemon.Level + 10) ^ 2.5) / ((TempPlayer(index).EnemyPokemon.Level + .Level + 10) ^ 2.5))) + 1) * IsTrade ' * ItemBooster
    End With
    
    Exit Function
errHandler:
    HandleError "GetBattleExp", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function CheckEffective(ByVal PType As PokeType, ByVal MoveType As PokeType) As Single
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Select Case PType
        Case PokeType.Normal
            If MoveType = PokeType.Fight Then
                CheckEffective = 2
            ElseIf MoveType = PokeType.Ghost Then
                CheckEffective = 0
            Else
                CheckEffective = 1
            End If
        Case PokeType.Fight
            If MoveType = PokeType.Flying Or MoveType = PokeType.Flying Then
                CheckEffective = 2
            ElseIf MoveType = PokeType.Rock Or MoveType = PokeType.Bug Then
                CheckEffective = 0.5
            Else
                CheckEffective = 1
            End If
        Case PokeType.Flying
            If MoveType = PokeType.Rock Or MoveType = PokeType.Electric Or MoveType = PokeType.Ice Then
                CheckEffective = 2
            ElseIf MoveType = PokeType.Fight Or MoveType = PokeType.Bug Then
                CheckEffective = 0.5
            ElseIf MoveType = PokeType.Ground Then
                CheckEffective = 0
            Else
                CheckEffective = 1
            End If
        Case PokeType.Poison
            If MoveType = PokeType.Ground Or MoveType = PokeType.Psychic Then
                CheckEffective = 2
            ElseIf MoveType = PokeType.Fight Or MoveType = PokeType.Bug Or MoveType = PokeType.Poison Then
                CheckEffective = 0.5
            Else
                CheckEffective = 1
            End If
        Case PokeType.Ground
            If MoveType = PokeType.Water Or MoveType = PokeType.Grass Or MoveType = PokeType.Ice Then
                CheckEffective = 2
            ElseIf MoveType = PokeType.Rock Or MoveType = PokeType.Poison Then
                CheckEffective = 0.5
            Else
                CheckEffective = 1
            End If
        Case PokeType.Rock
            If MoveType = PokeType.Fight Or MoveType = PokeType.Ground Or MoveType = PokeType.Steel Or MoveType = PokeType.Water Or MoveType = PokeType.Grass Then
                CheckEffective = 2
            ElseIf MoveType = PokeType.Normal Or MoveType = PokeType.Flying Or MoveType = PokeType.Poison Or MoveType = PokeType.Fire Then
                CheckEffective = 0.5
            Else
                CheckEffective = 1
            End If
        Case PokeType.Bug
            If MoveType = PokeType.Flying Or MoveType = PokeType.Rock Or MoveType = PokeType.Fire Then
                CheckEffective = 2
            ElseIf MoveType = PokeType.Fight Or MoveType = PokeType.Ground Or MoveType = PokeType.Grass Then
                CheckEffective = 0.5
            Else
                CheckEffective = 1
            End If
        Case PokeType.Ghost
            If MoveType = PokeType.Ghost Or MoveType = PokeType.Dark Then
                CheckEffective = 2
            ElseIf MoveType = PokeType.Poison Or MoveType = PokeType.Bug Then
                CheckEffective = 0.5
            ElseIf MoveType = PokeType.Fight Or MoveType = PokeType.Normal Then
                CheckEffective = 0
            Else
                CheckEffective = 1
            End If
        Case PokeType.Steel
            If MoveType = PokeType.Fight Or MoveType = PokeType.Ground Or MoveType = PokeType.Fire Then
                CheckEffective = 2
            ElseIf MoveType = PokeType.Flying Or MoveType = PokeType.Normal Or MoveType = PokeType.Rock Or MoveType = PokeType.Bug Or MoveType = PokeType.Steel Or MoveType = PokeType.Grass Or MoveType = PokeType.Psychic Or MoveType = PokeType.Ice Or MoveType = PokeType.Dragon Then
                CheckEffective = 0.5
            ElseIf MoveType = PokeType.Poison Then
                CheckEffective = 0
            Else
                CheckEffective = 1
            End If
        Case PokeType.Fire
            If MoveType = PokeType.Ground Or MoveType = PokeType.Rock Or MoveType = PokeType.Water Then
                CheckEffective = 2
            ElseIf MoveType = PokeType.Bug Or MoveType = PokeType.Steel Or MoveType = PokeType.Fire Or MoveType = PokeType.Grass Or MoveType = PokeType.Ice Then
                CheckEffective = 0.5
            Else
                CheckEffective = 1
            End If
        Case PokeType.Water
            If MoveType = PokeType.Grass Or MoveType = PokeType.Electric Then
                CheckEffective = 2
            ElseIf MoveType = PokeType.Steel Or MoveType = PokeType.Fire Or MoveType = PokeType.Water Or MoveType = PokeType.Ice Then
                CheckEffective = 0.5
            Else
                CheckEffective = 1
            End If
        Case PokeType.Grass
            If MoveType = PokeType.Flying Or MoveType = PokeType.Poison Or MoveType = PokeType.Bug Or MoveType = PokeType.Fire Or MoveType = PokeType.Ice Then
                CheckEffective = 2
            ElseIf MoveType = PokeType.Ground Or MoveType = PokeType.Water Or MoveType = PokeType.Grass Or MoveType = PokeType.Electric Then
                CheckEffective = 0.5
            Else
                CheckEffective = 1
            End If
        Case PokeType.Electric
            If MoveType = PokeType.Ground Then
                CheckEffective = 2
            ElseIf MoveType = PokeType.Flying Or MoveType = PokeType.Steel Or MoveType = PokeType.Electric Then
                CheckEffective = 0.5
            Else
                CheckEffective = 1
            End If
        Case PokeType.Psychic
            If MoveType = PokeType.Bug Or MoveType = PokeType.Ghost Or MoveType = PokeType.Dark Then
                CheckEffective = 2
            ElseIf MoveType = PokeType.Fight Or MoveType = PokeType.Psychic Then
                CheckEffective = 0.5
            Else
                CheckEffective = 1
            End If
        Case PokeType.Ice
            If MoveType = PokeType.Fight Or MoveType = PokeType.Rock Or MoveType = PokeType.Steel Or MoveType = PokeType.Fire Then
                CheckEffective = 2
            ElseIf MoveType = PokeType.Ice Then
                CheckEffective = 0.5
            Else
                CheckEffective = 1
            End If
        Case PokeType.Dragon
            If MoveType = PokeType.Ice Or MoveType = PokeType.Dragon Then
                CheckEffective = 2
            ElseIf MoveType = PokeType.Fire Or MoveType = PokeType.Water Or MoveType = PokeType.Grass Or MoveType = PokeType.Electric Then
                CheckEffective = 0.5
            Else
                CheckEffective = 1
            End If
        Case PokeType.Dark
            If MoveType = PokeType.Fight Or MoveType = PokeType.Bug Then
                CheckEffective = 2
            ElseIf MoveType = PokeType.Ghost Or MoveType = PokeType.Dark Then
                CheckEffective = 0.5
            ElseIf MoveType = PokeType.Psychic Then
                CheckEffective = 0
            Else
                CheckEffective = 1
            End If
        Case Else
            CheckEffective = 1
    End Select
    
    Exit Function
errHandler:
    HandleError "CheckEffective", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function GetPokemonDamage(ByVal index As Long, ByVal MoveNum As Long) As Long
Dim Modify As Single, Stab As Single

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    With Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(TempPlayer(index).InBattlePoke)
        If Pokemon(.Num).PType = Moves(MoveNum).Type Then
            Stab = 1.5
        Else
            Stab = 1
        End If
        Modify = Stab * CheckEffective(Pokemon(TempPlayer(index).EnemyPokemon.Num).PType, Moves(MoveNum).Type) * RandomDigit(0.85, 1)
        If Moves(MoveNum).AtkType = PhysicalAttack Then
            GetPokemonDamage = (((2 * .Level + 10) / 250) * (.Stat(Stats.Atk) / TempPlayer(index).EnemyPokemon.Stat(Stats.Def)) * Moves(MoveNum).Power + 2) * Modify
        ElseIf Moves(MoveNum).AtkType = SpecialAttack Then
            GetPokemonDamage = (((2 * .Level + 10) / 250) * (.Stat(Stats.SpAtk) / TempPlayer(index).EnemyPokemon.Stat(Stats.SpDef)) * Moves(MoveNum).Power + 2) * Modify
        End If
    End With
    
    Exit Function
errHandler:
    HandleError "GetPokemonDamage", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function GetEnemyDamage(ByVal index As Long, ByVal MoveNum As Long) As Long
Dim Modify As Single, Stab As Single

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    With Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(TempPlayer(index).InBattlePoke)
        If Pokemon(TempPlayer(index).EnemyPokemon.Num).PType = Moves(MoveNum).Type Then
            Stab = 1.5
        Else
            Stab = 1
        End If
        Modify = Stab * CheckEffective(Pokemon(.Num).PType, Moves(MoveNum).Type) * RandomDigit(0.85, 1)
        If Moves(MoveNum).AtkType = PhysicalAttack Then
            GetEnemyDamage = (((2 * .Level + 10) / 250) * (TempPlayer(index).EnemyPokemon.Stat(Stats.Atk) / .Stat(Stats.Def)) * Moves(MoveNum).Power + 2) * Modify
        ElseIf Moves(MoveNum).AtkType = SpecialAttack Then
            GetEnemyDamage = (((2 * .Level + 10) / 250) * (TempPlayer(index).EnemyPokemon.Stat(Stats.SpAtk) / .Stat(Stats.SpDef)) * Moves(MoveNum).Power + 2) * Modify
        End If
    End With
    
    Exit Function
errHandler:
    HandleError "GetEnemyDamage", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function GetPlayerPokeSlotStat(ByVal index As Long, ByVal Slot As Byte, ByVal Stat As Stats) As Long
Dim IV As Long, EV As Long, BStat() As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Slot <= 0 Or Slot > MAX_POKEMON Then Exit Function
    With Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(Slot)
        If .Num <= 0 Or .Num > Count_Pokemon Then Exit Function
        ReDim BStat(1 To Stats.Stat_Count - 1)
        BStat(Stat) = Pokemon(.Num).BaseStat(Stat)
        IV = .StatIV(Stat): EV = .StatEV(Stat)
        ' Temp: Pos_Nature
        If Stat = Stats.HP Then
            GetPlayerPokeSlotStat = (((2 * BStat(Stat) + IV + (EV / 4) + 100) * .Level) / 100 + 10)
        Else
            GetPlayerPokeSlotStat = (((2 * BStat(Stat) + IV + (EV / 4)) * .Level) / 100 + 5) * POS_NATURE
        End If
    End With
    
    Exit Function
errHandler:
    HandleError "GetPlayerPokeSlotStat", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function GetEnemyStat(ByVal index As Long, ByVal Stat As Stats) As Long
Dim IV As Long, EV As Long, BStat() As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    With TempPlayer(index).EnemyPokemon
        If .Num <= 0 Or .Num > Count_Pokemon Then Exit Function
        ReDim BStat(1 To Stats.Stat_Count - 1)
        BStat(Stat) = Pokemon(.Num).BaseStat(Stat)
        IV = .StatIV(Stat): EV = .StatEV(Stat)
        ' Temp: Pos_Nature
        If Stat = Stats.HP Then
            GetEnemyStat = (((2 * BStat(Stat) + IV + (EV / 4) + 100) * .Level) / 100 + 10)
        Else
            GetEnemyStat = (((2 * BStat(Stat) + IV + (EV / 4)) * .Level) / 100 + 5) * POS_NATURE
        End If
    End With
    
    Exit Function
errHandler:
    HandleError "GetEnemyStat", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub InitPlayerVsPlayer(ByVal index As Long, ByVal FoeIndex As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ClearEnemyPokemon index
    ClearEnemyPokemon FoeIndex
    TempPlayer(index).InBattle = BATTLE_TRAINER
    TempPlayer(FoeIndex).InBattle = BATTLE_TRAINER
    TempPlayer(index).InBattlePoke = CheckPokemon(index)
    TempPlayer(FoeIndex).InBattlePoke = CheckPokemon(FoeIndex)
    With Player(FoeIndex).PlayerData(TempPlayer(FoeIndex).CurSlot)
        TempPlayer(index).EnemyPokemon = .Pokemon(TempPlayer(FoeIndex).InBattlePoke)
    End With
    With Player(index).PlayerData(TempPlayer(index).CurSlot)
        TempPlayer(FoeIndex).EnemyPokemon = .Pokemon(TempPlayer(index).InBattlePoke)
    End With
    SendBattle index
    SendBattle FoeIndex
    SendEnemyPokemon index
    SendEnemyPokemon FoeIndex
    
    TempPlayer(index).MoveSet = 0
    TempPlayer(FoeIndex).MoveSet = 0
    
    Exit Sub
errHandler:
    HandleError "InitPlayerVsPlayer", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub InitPlayerVsNpc(ByVal index As Long, ByVal PokeNum As Long, ByVal Level As Long)
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ClearEnemyPokemon index
    With TempPlayer(index).EnemyPokemon
        .Num = PokeNum
        If Rnd <= Pokemon(PokeNum).FemaleRate Then
            .Gender = GENDER_FEMALE
        Else
            .Gender = GENDER_MALE
        End If
        .Level = Level
        For i = 1 To Stats.Stat_Count - 1
            .Stat(i) = GetEnemyStat(index, i)
            .StatIV(i) = Random(1, 31)
            .StatEV(i) = 0
        Next
        .CurHP = .Stat(Stats.HP)
        .Exp = 0
    
        GetEnemyPokemonMove index
    End With
    TempPlayer(index).InBattle = BATTLE_WILD
    TempPlayer(index).InBattlePoke = CheckPokemon(index)
    SendBattle index
    SendEnemyPokemon index
    
    SendBattleMsg index, "Wild " & Trim$(Pokemon(PokeNum).Name) & " appeared!", White
    
    Exit Sub
errHandler:
    HandleError "InitPlayerVsNpc", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub InitBattleVsNPC(ByVal index As Long, ByVal MoveSlot As Long)
Dim CurSlot As Long
Dim Exp As Long
Dim DidLevelUp As Boolean

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    CurSlot = TempPlayer(index).CurSlot
    With Player(index).PlayerData(CurSlot).Pokemon(TempPlayer(index).InBattlePoke)
        If .Stat(Stats.Spd) >= TempPlayer(index).EnemyPokemon.Stat(Stats.Spd) Then
            PlayerVsNpc index, MoveSlot
            If TempPlayer(index).EnemyPokemon.CurHP > 0 Then
                NpcVsPlayer index
                If .CurHP <= 0 Then
                    If CheckPokemon(index) > 0 Then
                        SendForceSwitch index
                    Else
                        ExitBattle index, YES
                    End If
                End If
            Else
                Exp = GetBattleExp(index)
                If Player(index).PlayerData(TempPlayer(index).CurSlot).IsVIP = YES Then
                    GiveMoney index, (Random(1, 100) * 2)
                Else
                    GiveMoney index, Random(1, 100)
                End If
                DidLevelUp = HandleLevel(index, TempPlayer(index).InBattlePoke, Exp)
                SendBattleMsg index, Trim$(Pokemon(.Num).Name) & " gained " & Exp & " EXP.Points!", Green
                ExitBattle index, , DidLevelUp, YES
                SendBattleMsg index, EndLine, Cyan
                Exit Sub
            End If
        Else
            NpcVsPlayer index
            If .CurHP > 0 Then
                PlayerVsNpc index, MoveSlot
                If TempPlayer(index).EnemyPokemon.CurHP <= 0 Then
                    Exp = GetBattleExp(index)
                    If Player(index).PlayerData(TempPlayer(index).CurSlot).IsVIP = YES Then
                        GiveMoney index, (Random(1, 100) * 2)
                    Else
                        GiveMoney index, Random(1, 100)
                    End If
                    DidLevelUp = HandleLevel(index, TempPlayer(index).InBattlePoke, Exp)
                    SendBattleMsg index, Trim$(Pokemon(.Num).Name) & " gained " & Exp & " EXP.Points!", Green
                    ExitBattle index, , DidLevelUp, YES
                    SendBattleMsg index, EndLine, Cyan
                    Exit Sub
                End If
            Else
                If CheckPokemon(index) > 0 Then
                    SendForceSwitch index
                Else
                    ExitBattle index, YES
                End If
            End If
        End If
    End With
    SendBattleMsg index, EndLine, Cyan
    SendBattleResult index
    
    Exit Sub
errHandler:
    HandleError "InitBattleVsNPC", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub InitBattleVsPlayer(ByVal index As Long, ByVal FoeIndex As Long, ByVal MoveSlot As Long, ByVal FoeMoveSlot As Long)
Dim pSlot As Long, pPoke As Long
Dim fSlot As Long, fPoke As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    pSlot = TempPlayer(index).CurSlot: pPoke = TempPlayer(index).InBattlePoke
    fSlot = TempPlayer(FoeIndex).CurSlot: fPoke = TempPlayer(FoeIndex).InBattlePoke
    With Player(index).PlayerData(pSlot).Pokemon(pPoke)
        If .Stat(Stats.Spd) > TempPlayer(index).EnemyPokemon.Stat(Stats.Spd) Then
            PlayerVsPlayer index, FoeIndex, MoveSlot
            If TempPlayer(index).EnemyPokemon.CurHP > 0 Then
                PlayerVsPlayer FoeIndex, index, FoeMoveSlot
                If .CurHP <= 0 Then
                    If CheckPokemon(index) > 0 Then
                        SendForceSwitch index
                    Else
                        GivePvPPoints index, 2
                        GivePvPPoints FoeIndex, 1
                        ExitBattle index, YES
                        ExitBattle FoeIndex, NO
                        TempPlayer(index).MoveSet = 0
                        TempPlayer(FoeIndex).MoveSet = 0
                        Exit Sub
                    End If
                End If
            Else
                If CheckPokemon(FoeIndex) > 0 Then
                    SendForceSwitch FoeIndex
                Else
                    GivePvPPoints index, 1
                    GivePvPPoints FoeIndex, 2
                    ExitBattle index, NO
                    ExitBattle FoeIndex, YES
                    TempPlayer(index).MoveSet = 0
                    TempPlayer(FoeIndex).MoveSet = 0
                    Exit Sub
                End If
            End If
        Else
            PlayerVsPlayer FoeIndex, index, FoeMoveSlot
            If .CurHP > 0 Then
                PlayerVsPlayer index, FoeIndex, MoveSlot
                If TempPlayer(index).EnemyPokemon.CurHP <= 0 Then
                    If CheckPokemon(FoeIndex) > 0 Then
                        SendForceSwitch FoeIndex
                    Else
                        GivePvPPoints index, 1
                        GivePvPPoints FoeIndex, 2
                        ExitBattle index, NO
                        ExitBattle FoeIndex, YES
                        TempPlayer(index).MoveSet = 0
                        TempPlayer(FoeIndex).MoveSet = 0
                        Exit Sub
                    End If
                End If
            Else
                If CheckPokemon(index) > 0 Then
                    SendForceSwitch index
                Else
                    GivePvPPoints index, 2
                    GivePvPPoints FoeIndex, 1
                    ExitBattle index, YES
                    ExitBattle FoeIndex, NO
                    TempPlayer(index).MoveSet = 0
                    TempPlayer(FoeIndex).MoveSet = 0
                    Exit Sub
                End If
            End If
        End If
    End With
    TempPlayer(index).MoveSet = 0
    TempPlayer(FoeIndex).MoveSet = 0
    SendBattleMsg index, EndLine, Cyan
    SendBattleMsg FoeIndex, EndLine, Cyan
    SendBattleResult index
    SendBattleResult FoeIndex
    
    Exit Sub
errHandler:
    HandleError "InitBattleVsPlayer", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub PlayerVsNpc(ByVal index As Long, ByVal MoveSlot As Long)
Dim Damage As Long
Dim Crit As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Rnd <= 0.1 Then
        Crit = 2
    Else
        Crit = 1
    End If

    With Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(TempPlayer(index).InBattlePoke)
        If MoveSlot > 0 Then
            SendBattleMsg index, Trim$(Pokemon(.Num).Name) & " used " & Trim$(Moves(.Moves(MoveSlot).Num).Name) & "!", White
            Damage = GetPokemonDamage(index, .Moves(MoveSlot).Num) * Crit
            SendBattleMsg index, "Wild " & Trim$(Pokemon(TempPlayer(index).EnemyPokemon.Num).Name) & " deals " & Damage & " damage!", White
            If Crit = 2 Then SendBattleMsg index, "A critical hit!", White
            Select Case CheckEffective(Pokemon(TempPlayer(index).EnemyPokemon.Num).PType, Moves(.Moves(MoveSlot).Num).Type)
                Case 2
                    SendBattleMsg index, "It's super effective!", White
                Case 0.5
                    SendBattleMsg index, "It's not very effective!", White
                Case 0
                    SendBattleMsg index, "It's not effective!", White
            End Select
            If Damage >= TempPlayer(index).EnemyPokemon.CurHP Then
                TempPlayer(index).EnemyPokemon.CurHP = 0
                SendBattleMsg index, "Wild " & Trim$(Pokemon(TempPlayer(index).EnemyPokemon.Num).Name) & " fainted!", White
            Else
                TempPlayer(index).EnemyPokemon.CurHP = TempPlayer(index).EnemyPokemon.CurHP - Damage
            End If
        Else
            ' Player Struggle
            SendBattleMsg index, Trim$(Pokemon(.Num).Name) & " used Struggle!", White
        End If
    
        If MoveSlot > 0 Then .Moves(MoveSlot).PP = .Moves(MoveSlot).PP - 1
        SendUpdatePokemonVital index, index, TempPlayer(index).InBattlePoke
        SendUpdateEnemyVital index
    End With
    
    Exit Sub
errHandler:
    HandleError "PlayerVsNpc", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub NpcVsPlayer(ByVal index As Long)
Dim x As Long
Dim Damage As Long
Dim Crit As Byte

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Rnd <= 0.1 Then
        Crit = 2
    Else
        Crit = 1
    End If
    
    With Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(TempPlayer(index).InBattlePoke)
        x = GetNpcMove(index)
        If x > 0 Then
            SendBattleMsg index, "Wild " & Trim$(Pokemon(TempPlayer(index).EnemyPokemon.Num).Name) & " used " & Trim$(Moves(TempPlayer(index).EnemyPokemon.Moves(x).Num).Name) & "!", White
            Damage = GetEnemyDamage(index, x) * Crit
            SendBattleMsg index, Trim$(Pokemon(.Num).Name) & " deals " & Damage & " damage!", White
            If Crit = 2 Then SendBattleMsg index, "A critical hit!", White
            Select Case CheckEffective(Pokemon(.Num).PType, Moves(TempPlayer(index).EnemyPokemon.Moves(x).Num).Type)
                Case 2
                    SendBattleMsg index, "It's super effective!", White
                Case 0.5
                    SendBattleMsg index, "It's not very effective!", White
                Case 0
                    SendBattleMsg index, "It's not effective!", White
            End Select
            If Damage >= .CurHP Then
                .CurHP = 0
            Else
                .CurHP = .CurHP - Damage
            End If
        Else
            ' Enemy Struggle
            SendBattleMsg index, "Wild " & Trim$(Pokemon(TempPlayer(index).EnemyPokemon.Num).Name) & " used Struggle!", White
        End If
    
        If x > 0 Then TempPlayer(index).EnemyPokemon.Moves(x).PP = TempPlayer(index).EnemyPokemon.Moves(x).PP - 1
        SendUpdatePokemonVital index, index, TempPlayer(index).InBattlePoke
        SendUpdateEnemyVital index
    End With
    
    Exit Sub
errHandler:
    HandleError "NpcVsPlayer", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub PlayerVsPlayer(ByVal index As Long, ByVal FoeIndex As Long, ByVal MoveSlot As Long)
Dim Damage As Long
Dim Crit As Byte
Dim fPoke As Long, fSlot As Long
Dim pPoke As Long, pSlot As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If MoveSlot > 4 Then Exit Sub

    If Rnd <= 0.1 Then
        Crit = 2
    Else
        Crit = 1
    End If
    
    pSlot = TempPlayer(index).CurSlot: fSlot = TempPlayer(FoeIndex).CurSlot
    pPoke = TempPlayer(index).InBattlePoke: fPoke = TempPlayer(FoeIndex).InBattlePoke
    
    With Player(index).PlayerData(pSlot).Pokemon(pPoke)
        If MoveSlot > 0 Then
            SendBattleMsg index, Trim$(Pokemon(.Num).Name) & " used " & Trim$(Moves(.Moves(MoveSlot).Num).Name) & "!", White
            SendBattleMsg FoeIndex, "Foe " & Trim$(Pokemon(.Num).Name) & " used " & Trim$(Moves(.Moves(MoveSlot).Num).Name) & "!", White
            Damage = GetPokemonDamage(index, MoveSlot) * Crit
            SendBattleMsg index, "Foe " & Trim$(Pokemon(TempPlayer(index).EnemyPokemon.Num).Name) & " deals " & Damage & " damage!", White
            SendBattleMsg FoeIndex, Trim$(Pokemon(TempPlayer(index).EnemyPokemon.Num).Name) & " deals " & Damage & " damage!", White
            If Crit = 2 Then
                SendBattleMsg index, "A critical hit!", White
                SendBattleMsg FoeIndex, "A critical hit!", White
            End If
            Select Case CheckEffective(Pokemon(TempPlayer(index).EnemyPokemon.Num).PType, Moves(.Moves(MoveSlot).Num).Type)
                Case 2
                    SendBattleMsg index, "It's super effective!", White
                    SendBattleMsg FoeIndex, "It's super effective!", White
                Case 0.5
                    SendBattleMsg index, "It's not very effective!", White
                    SendBattleMsg FoeIndex, "It's not very effective!", White
                Case 0
                    SendBattleMsg index, "It's not effective!", White
                    SendBattleMsg FoeIndex, "It's not very effective!", White
            End Select
            If Damage >= TempPlayer(index).EnemyPokemon.CurHP Then
                TempPlayer(index).EnemyPokemon.CurHP = 0
            Else
                TempPlayer(index).EnemyPokemon.CurHP = TempPlayer(index).EnemyPokemon.CurHP - Damage
            End If
        Else
            ' Enemy Struggle
            SendBattleMsg index, Trim$(Pokemon(.Num).Name) & " used Struggle!", White
            SendBattleMsg FoeIndex, "Foe " & Trim$(Pokemon(.Num).Name) & " used Struggle!", White
        End If
        
        If MoveSlot > 0 Then .Moves(MoveSlot).PP = .Moves(MoveSlot).PP - 1
        
        Player(FoeIndex).PlayerData(fSlot).Pokemon(fPoke) = TempPlayer(index).EnemyPokemon
        SendUpdatePokemonVital FoeIndex, FoeIndex, fPoke
        SendUpdateEnemyVital index
    End With
    
    Exit Sub
errHandler:
    HandleError "PlayerVsPlayer", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function CountNpcMove(ByVal index As Long) As Long
Dim i As Long, x As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    For i = 1 To MAX_POKEMON_MOVES
        With TempPlayer(index).EnemyPokemon
            If .Moves(i).Num > 0 Then
                If .Moves(i).PP > 0 Then
                    x = x + 1
                    CheckNpcMove(i).InputNum = x
                Else
                    CheckNpcMove(i).InputNum = 0
                End If
            Else
                CheckNpcMove(i).InputNum = 0
            End If
        End With
    Next
    CountNpcMove = x
    
    Exit Function
errHandler:
    HandleError "CountNpcMove", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function GetNpcMove(ByVal index As Long) As Long
Dim r As Long, x As Long, y As Long
Dim i As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    GetNpcMove = 0
    x = CountNpcMove(index)
    If x > 1 Then
        r = Random(1, x)
        For y = 1 To MAX_POKEMON_MOVES
            If CheckNpcMove(y).InputNum = r Then
                GetNpcMove = y
                Exit Function
            End If
        Next
    ElseIf x = 1 Then
        GetNpcMove = 1
        Exit Function
    End If
    i = i

    Exit Function
errHandler:
    HandleError "GetNpcMove", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub GetEscapeChance(ByVal index As Long)
Dim F As Long, Chance As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If TempPlayer(index).InBattlePoke = 0 Then Exit Sub
    
    With Player(index).PlayerData(TempPlayer(index).CurSlot).Pokemon(TempPlayer(index).InBattlePoke)
        If TempPlayer(index).EnemyPokemon.Stat(Stats.Spd) <= 0 Then
            F = (((.Stat(Stats.Spd) * 128) / 1) + 30 * TempPlayer(index).EscCount) Mod 256
        Else
            F = (((.Stat(Stats.Spd) * 128) / TempPlayer(index).EnemyPokemon.Stat(Stats.Spd)) + 30 * TempPlayer(index).EscCount) Mod 256
        End If
        If F > 255 Then
            SendBattleMsg index, "You have successfully escaped!", White
            ExitBattle index
        Else
            Chance = Random(0, 255)
            If Chance < F Then
                SendBattleMsg index, "You have successfully escaped!", White
                ExitBattle index
            Else
                SendBattleMsg index, "You have failed to escape!", White
                SendBattleResult index
                TempPlayer(index).EscCount = TempPlayer(index).EscCount + 1
                NpcVsPlayer index
                If .CurHP <= 0 Then
                    If CheckPokemon(index) > 0 Then
                        SendForceSwitch index
                    Else
                        ExitBattle index, YES
                    End If
                End If
                SendBattleMsg index, EndLine, Cyan
            End If
        End If
    End With
    
    Exit Sub
errHandler:
    HandleError "ExitBattle", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ExitBattle(ByVal index As Long, Optional ByVal Forced As Byte = NO, Optional ByVal DidLevelUp As Boolean = False, Optional ByVal Didwin As Byte = 0)
Dim CanEvolve As Boolean

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If DidLevelUp Then
        CanEvolve = CheckEvolve(index, TempPlayer(index).InBattlePoke)
        If CanEvolve Then
            SendEvolve index, TempPlayer(index).InBattlePoke
        End If
    End If
    TempPlayer(index).InBattle = 0
    TempPlayer(index).InBattlePoke = 0
    TempPlayer(index).EscCount = 0
    TempPlayer(index).BattleRequest = 0
    TempPlayer(index).MoveSet = 0
    ClearEnemyPokemon index
    SendBattle index
    If Forced = YES Then
        SendExitBattle index, Didwin
        With Player(index).PlayerData(TempPlayer(index).CurSlot).Checkpoint
            PlayerWarp index, .Map, .x, .y
        End With
        RestoreAllPokemon index
        SendMsg index, "You have been wiped out!", Red
    Else
        SendExitBattle index, Didwin
    End If

    Exit Sub
errHandler:
    HandleError "ExitBattle", "modPlayer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub
