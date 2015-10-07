Attribute VB_Name = "modEnumeration"
Option Explicit

Public Enum ServerPackets
    SAlertMsg = 1
    SCharSelect
    SIndex
    SHighIndex
    SInGame
    SPlayerXY
    SPlayerXYToMap
    SLeaveMap
    SCheckForMap
    SMap
    SMapDone
    SPlayerData
    SPlayerDir
    SPlayerMove
    SLeft
    SMsg
    SEditMap
    SEditPokemon
    SUpdatePokemon
    SPlayerPokemon
    SBattle
    SEnemyPokemon
    SBattleMsg
    SPlayMusic
    SPlaySound
    SEditMove
    SUpdateMove
    SExpCalc
    STarget
    SUpdatePokemonVital
    SUpdateEnemyVital
    SBattleResult
    SExitBattle
    SForceSwitch
    SSwitch
    SLearnMove
    SEvolve
    SEditItem
    SUpdateItem
    SInventory
    SSelect
    SBattleRequest
    SCaptured
    SPlayerStoredPokemon
    SStorage
    SEditNPC
    SUpdateNPC
    SEditShop
    SUpdateShop
    SOpenShop
    
    SSpawnNpc
    SNPCClear
    SNpcMove
    SMapNpcData
    SNpcDir
    
    STradeRequest
    STrade
    SCloseTrade
    STradeConfirm
    ServerPacket_Count
End Enum

Public Enum ClientPackets
    CRegister = 1
    CLogin
    CAddChar
    CDelChar
    CUserChar
    CNeedMap
    CPlayerMove
    CPlayerDir
    CMsg
    CRequestNewMap
    CRefresh
    CRequestEditMap
    CMap
    CRequestEditPokemon
    CRequestPokemons
    CSavePokemon
    CBattleCommand
    CRequestEditMove
    CRequestMoves
    CSaveMove
    CExpCalc
    CTarget
    CSwitchComplete
    CAdminWarp
    CReplaceMove
    CEvolve
    CRequestEditItem
    CRequestItems
    CSaveItem
    CUseItem
    CInitSelect
    CBattleRequest
    CInitBattle
    CDepositPokemon
    CWithdrawPokemon
    CRequestEditNPC
    CRequestNPCs
    CSaveNPC
    CRequestEditShop
    CRequestShops
    CSaveShop
    CBuyItem
    CSellItem
    CInitTrade
    CTradeAccept
    CTradeDecline
    CCloseTrade
    CTradeConfirm
    ClientPacket_Count
End Enum

Public HandleDataSub(ServerPacket_Count) As Long

Public Enum WindowType
    ' Menu~
    Menu_Login = 1
    Menu_Register
    Menu_CharSelect
    Menu_CharCreate
    ' Menu End~
    ' Main~
    Main_Chatbox
    Main_Trainer
    Main_Inventory
    Main_Option
    Main_Storage
    ' Main End~
    Window_Count
End Enum

Public Enum ButtonEnum
    ' Menu~
    LoginAccept = 1
    Register
    RegisterAccept
    Login
    CharNew1
    CharNew2
    CharNew3
    CharUse1
    CharUse2
    CharUse3
    CharDel1
    CharDel2
    CharDel3
    CharAccept
    CharDecline
    ' Menu End~
    ' Main~
    bChatScrollUp
    bChatScrollDown
    mPokedex
    mCharacter
    mInventory
    mOptions
    BattleFight
    BattleSwitch
    BattleBag
    BattleRun
    bBattleScrollUp
    bBattleScrollDown
    EvolveYes
    EvolveNo
    InvScrollUp
    InvScrollDown
    InvItems
    InvPokeballs
    InvTM_HMs
    InvBerries
    InvKeyItems
    PCDepositPoke
    PCClose
    PCNext
    PCPrevious
    ShopScrollUp
    ShopScrollDown
    TradeConfirm
    TradeAccept
    TradeDecline
    ' Main End~
    MaxButton
End Enum

Public Enum Layers
    Ground = 0
    mask
    Mask2
    Fringe
    Fringe2
    LayerCount
End Enum

Public Enum Attributes
    Walkable = 0
    Blocked
    TallGrass
    Heal
    Checkpoint
    Storage
    Warp
    mShop
    AttributeCount
End Enum

Public Enum MsgType
    MapMsg = 1
    GlobalMsg
    Msg_Count
End Enum

Public Enum Stats
    HP = 1
    Atk
    Def
    SpAtk
    SpDef
    Spd
    Stat_Count
End Enum

Public Enum PokeType
    Normal = 0
    Fight
    Flying
    Poison
    Ground
    Rock
    Bug
    Ghost
    Steel
    Fire
    Water
    Grass
    Electric
    Psychic
    Ice
    Dragon
    Dark
    Type_Count
End Enum

Public Enum ItemType
    Items = 0
    Pokeballs
    TM_HMs
    Berries
    KeyItems
    Item_Count
End Enum

Public Enum ItemProperties
    None = 0
    RestoreHP
    RestorePP
End Enum
