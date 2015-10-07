Attribute VB_Name = "modType"
Option Explicit

Public Buttons(1 To ButtonEnum.MaxButton) As ButtonRec
Public Player(1 To MAX_PLAYER) As PlayerRec
Public Map As MapRec
Public Options As OptionRec
Public Pokemon(1 To Count_Pokemon) As PokemonRec
Public EnemyPokemon As PlayerPokemonRec
Public Moves(1 To Count_Move) As MoveRec
Public ExpCalc(1 To MAX_LEVEL) As Long
Public Item(1 To Count_Item) As ItemRec
Public NPC(1 To Count_NPC) As NPCRec
Public Shop(1 To Count_Shop) As ShopRec

Public MapNpc(1 To MAX_MAP_NPC) As MapNpcRec

Public MyTrade(1 To MAX_TRADE) As TradeRec
Public MyTradeConfirm As Boolean
Public TheirTrade(1 To MAX_TRADE) As TradeRec
Public TheirTradeConfirm As Boolean

Private Type OptionRec
    Username As String * MAX_STRING
    Password As String * MAX_STRING
    SavePass As Byte
    
    SaveIp As String
    SavePort As Long
    
    Music As Byte
    Sound As Byte
End Type

Public Type LayerRec
    Tileset As Long
    X As Long
    Y As Long
End Type

Public Type TileRec
    Layer(0 To Layers.LayerCount - 1) As LayerRec
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As String
End Type

Public Type MapRec
    Name As String * MAX_STRING
    Music As String * MAX_STRING
    
    Rev As Long
    
    Moral As Byte
    
    MaxX As Long
    MaxY As Long

    Tile() As TileRec
    Link(0 To 3) As Long
    
    Pokemon(1 To MAX_MAP_POKEMON) As Long
    MinLvl As Long
    MaxLvl As Long
    NPC(1 To MAX_MAP_NPC) As Long
    
    CurField As Long
    CurBack As Long
End Type

Private Type ButtonRec
    bState As Byte
    X As Long
    Y As Long
    Pic As Long
    Visible As Boolean
End Type

Private Type PokemonMoveRec
    Num As Long
    PP As Long
    MaxPP As Long
End Type

Private Type PlayerPokemonRec
    Num As Long
    
    Gender As Byte
    
    Level As Long
    
    CurHP As Long
    Stat(1 To Stats.Stat_Count - 1) As Long
    StatIV(1 To Stats.Stat_Count - 1) As Long
    StatEV(1 To Stats.Stat_Count - 1) As Long
    
    Exp As Long
    
    Moves(1 To MAX_POKEMON_MOVES) As PokemonMoveRec
End Type

Private Type PlayerItemRec
    Num As Long
    value As Long
End Type

Private Type PvPRec
    win As Long
    Lose As Long
    Disconnect As Long
End Type

Public Type PlayerRec
    Name As String * MAX_STRING
    Gender As Byte
    
    Access As Byte
    
    Sprite As Long
    
    Map As Long
    X As Long
    Y As Long
    Dir As Byte
    
    Pokemon(1 To MAX_POKEMON) As PlayerPokemonRec
    StoredPokemon(1 To MAX_STORAGE_POKEMON) As PlayerPokemonRec
    Item(1 To MAX_PLAYER_ITEM, 0 To ItemType.Item_Count - 1) As PlayerItemRec
    
    PvP As PvPRec
    
    Money As Long
    
    IsVIP As Byte
    
    Moving As Byte
    xOffset As Long
    yOffset As Long
    Step As Byte
    InBattle As Byte
End Type

Public Type PokemonRec
    Name As String * MAX_STRING
    Pic As Long
    
    BaseStat(1 To Stats.Stat_Count - 1) As Long
    FemaleRate As Double
    BaseExp As Long
    
    MoveNum(1 To MAX_MOVES) As Long
    MoveLevel(1 To MAX_MOVES) As Long
    
    pType As Byte
    sType As Byte
    
    EvolveNum As Long
    EvolveLvl As Long
    
    CatchRate As Byte
End Type

Public Type MoveRec
    Name As String * MAX_STRING
    
    Power As Long
    PP As Long
    
    AtkType As Byte
    Type As Byte
End Type

Public Type ItemRec
    Name As String * MAX_STRING
    
    Type As Byte
    Pic As Long
    
    Desc As String * MAX_BYTE
    
    IType As Byte
    Data1 As Byte
    Data2 As Long
    Data3 As Single
    
    Sell As Long
End Type

Public Type NPCRec
    Name As String * MAX_STRING
    
    Sprite As Long
End Type

Private Type ShopItem
    Num As Long
    Price As Long
End Type

Public Type ShopRec
    Name As String * MAX_STRING
    
    sItem(1 To MAX_SHOP_ITEMS) As ShopItem
End Type

Public Type MapNpcRec
    Num As Long
    
    X As Long
    Y As Long
    Dir As Long

    Moving As Byte
    Step As Long
    xOffset As Long
    yOffset As Long
End Type

Public Type TradeRec
    Type As Byte
    
    ItemNum As Long
    ItemVal As Long
    Pokemon As PlayerPokemonRec
    
    TempItemSlot As Long
    TempItemType As Byte
    TempPokeSlot As Long
End Type
