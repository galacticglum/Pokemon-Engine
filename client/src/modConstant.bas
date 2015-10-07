Attribute VB_Name = "modConstant"
Option Explicit

Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lparam As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long

Public Const GameTitle As String = "Pokemon Engine"

Public Const Gfx_Ext As String = ".png"

Public Const FormMainHeight As Long = 9500
Public Const FormMainWidth As Long = 12095

Public Const Pic_Size As Byte = 32

Public Const MAX_PLAYER As Byte = 100
Public Const MAX_STRING As Byte = 21
Public Const MAX_PLAYER_DATA As Byte = 3
Public Const MAX_POKEMON As Byte = 6
Public Const MAX_STORAGE_POKEMON As Byte = 99
Public Const MAX_MAP_POKEMON As Byte = 30
Public Const MAX_LEVEL As Byte = 100
Public Const MAX_MOVES As Byte = 30
Public Const MAX_POKEMON_MOVES As Byte = 4
Public Const MAX_PLAYER_ITEM As Long = 99
Public Const MAX_PLAYER_INV_VALUE As Long = 99
Public Const MAX_SHOP_ITEMS As Long = 35
Public Const MAX_MAP_NPC As Long = 35
Public Const MAX_TRADE As Long = 10
Public Const Count_Map As Long = 100
Public Const Count_Pokemon As Long = 100
Public Const Count_Move As Long = 100
Public Const Count_Item As Long = 100
Public Const Count_NPC As Long = 100
Public Const Count_Shop As Long = 100

Public Const Max_MapX As Long = 24
Public Const Max_MapY As Long = 18

Public Const MAX_BYTE As Byte = 255
Public Const MAX_INTEGER As Integer = 32767
Public Const MAX_LONG As Long = 2147483647

Public Const GENDER_MALE As Byte = 0
Public Const GENDER_FEMALE As Byte = 1
Public Const GENDER_NONE As Byte = 2

Public Const YES As Byte = 1
Public Const NO As Byte = 0

Public Const DIR_UP As Byte = 0
Public Const DIR_LEFT As Byte = 1
Public Const DIR_RIGHT As Byte = 2
Public Const DIR_DOWN As Byte = 3

Public Const MOVING_WALKING As Byte = 1

Public Const WALK_SPEED As Byte = 6

Public Const ACCESS_MODERATOR As Byte = 1
Public Const ACCESS_MAPPER As Byte = 2
Public Const ACCESS_DEVELOPER As Byte = 3
Public Const ACCESS_ADMIN As Byte = 4

Public Const PhysicalAttack As Byte = 0
Public Const SpecialAttack As Byte = 1
Public Const NeutralAttack As Byte = 2

Public Const MAP_MORAL_NONE As Byte = 0
Public Const MAP_MORAL_FIELD As Byte = 1
Public Const MAP_MORAL_CITY As Byte = 2
Public Const MAP_MORAL_VILLAGE As Byte = 3
Public Const MAP_MORAL_CAVE As Byte = 4
Public Const MAP_MORAL_FOREST As Byte = 5
Public Const MAP_MORAL_RIVERSIDE As Byte = 6
Public Const MAP_MORAL_SEA As Byte = 7

Public Const BATTLE_COMMAND_FIGHT As Byte = 1
Public Const BATTLE_COMMAND_SWITCH As Byte = 2
Public Const BATTLE_COMMAND_BAG As Byte = 3
Public Const BATTLE_COMMAND_RUN As Byte = 4

Public Const BATTLE_WILD As Byte = 1
Public Const BATTLE_TRAINER As Byte = 2

Public Const GuiBackground As Byte = 1
Public Const GuiLogo As Byte = 2
Public Const GuiLogin As Byte = 3
Public Const GuiRegister As Byte = 4
Public Const GuiCharSelect As Byte = 5
Public Const GuiCharCreate As Byte = 6
Public Const GuiChatbox As Byte = 7
Public Const GuiPokeView As Byte = 8
Public Const GuiPokeSlot As Byte = 9
Public Const GuiBattle As Byte = 10
Public Const GuiBattleFoeHP As Byte = 11
Public Const GuiBattleHP As Byte = 12
Public Const GuiLearnMove As Byte = 14
Public Const GuiEvolve As Byte = 15
Public Const GuiTrainer As Byte = 16
Public Const GuiInventory As Byte = 17
Public Const GuiOption As Byte = 18
Public Const GuiTargetMenu As Byte = 19
Public Const GuiSelection As Byte = 20
Public Const GuiStorage As Byte = 21
Public Const GuiShop As Byte = 22
Public Const GuiCurrency As Byte = 23
Public Const GuiTrade As Byte = 24
Public Const GuiTradeConfirm As Byte = 25

Public Const GuiLoginX As Long = 234
Public Const GuiLoginY As Long = 390
Public Const GuiRegisterX As Long = 234
Public Const GuiRegisterY As Long = 390
Public Const GuiCharSelectX As Long = 297
Public Const GuiCharSelectY As Long = 380
Public Const GuiCharCreateX As Long = 281
Public Const GuiCharCreateY As Long = 380
Public Const GuiChatboxX As Long = 0
Public Const GuiChatboxY As Long = 465
Public Const GuiPokeViewX As Long = 0
Public Const GuiPokeViewY As Long = 0
Public Const GuiBattleX As Long = 0
Public Const GuiBattleY As Long = 0
Public Const GuiLearnMoveX As Long = 247
Public Const GuiLearnMoveY As Long = 193
Public Const GuiEvolveX As Long = 132
Public Const GuiEvolveY As Long = 101
Public Const GuiTrainerX As Long = 520
Public Const GuiTrainerY As Long = 330
Public Const GuiInventoryX As Long = 520
Public Const GuiInventoryY As Long = 330
Public Const GuiOptionX As Long = 520
Public Const GuiOptionY As Long = 330
Public Const GuiTargetMenuX As Long = 10
Public Const GuiTargetMenuY As Long = 70
Public Const GuiStorageX As Long = 172
Public Const GuiStorageY As Long = 115
Public Const GuiShopX As Long = 180
Public Const GuiShopY As Long = 174
Public Const GuiTradeX As Long = 205
Public Const GuiTradeY As Long = 125
Public Const GuiTradeConfirmX As Long = 234
Public Const GuiTradeConfirmY As Long = 133

Public Const ButtonNormal As Byte = 0
Public Const ButtonClick As Byte = 1

Public Const MenuStateLogin As Byte = 1
Public Const MenuStateRegister As Byte = 2
Public Const MenuStateAddChar As Byte = 3
Public Const MenuStateDelChar As Byte = 4
Public Const MenuStateUseChar As Byte = 5

Public Const MiscBlank As Byte = 1
Public Const MiscCheck As Byte = 2
Public Const MiscShadow As Byte = 3
Public Const MiscAlpha As Byte = 4
Public Const MiscCursor As Byte = 5
Public Const MiscBars As Byte = 6
Public Const MiscBattleBars As Byte = 7
Public Const MiscInBattle As Byte = 8
Public Const MiscGender As Byte = 9
Public Const MiscTarget As Byte = 10

Public Const HalfX As Integer = ((Max_MapX + 1) / 2) * Pic_Size
Public Const HalfY As Integer = ((Max_MapY + 1) / 2) * Pic_Size
Public Const ScreenX As Integer = (Max_MapX + 1) * Pic_Size
Public Const ScreenY As Integer = (Max_MapY + 1) * Pic_Size
Public Const StartXValue As Integer = ((Max_MapX + 2) / 2)
Public Const StartYValue As Integer = ((Max_MapY + 2) / 2)
Public Const EndXValue As Integer = (Max_MapX + 1) + 1
Public Const EndYValue As Integer = (Max_MapY + 1) + 1
Public Const Half_PIC_X As Integer = Pic_Size / 2
Public Const Half_PIC_Y As Integer = Pic_Size / 2

Public Const SELECT_POKEMON As Byte = 1
Public Const SELECT_MOVE As Byte = 2

Public Const TRADE_TYPE_POKEMON As Byte = 1
Public Const TRADE_TYPE_ITEM As Byte = 2
