Attribute VB_Name = "modConstant"
Option Explicit

Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal length As Long)
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByRef Msg() As Byte, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const AppTitle As String = "Pokemon Server"

Public Const MAX_PLAYER As Long = 100
Public Const MAX_STRING As Byte = 21
Public Const MAX_PLAYER_DATA As Byte = 3
Public Const MAX_POKEMON As Byte = 6
Public Const MAX_STORAGE_POKEMON As Byte = 99
Public Const MAX_MAP_POKEMON As Byte = 30
Public Const MAX_LEVEL As Byte = 100
Public Const MAX_MOVES As Byte = 30
Public Const MAX_POKEMON_MOVES As Byte = 4
Public Const MAX_EV As Long = 510
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

Public Const ACCESS_MODERATOR As Byte = 1
Public Const ACCESS_MAPPER As Byte = 2
Public Const ACCESS_DEVELOPER As Byte = 3
Public Const ACCESS_ADMIN As Byte = 4

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

Public Const POS_NATURE As Single = 1.1
Public Const NEU_NATURE As Single = 1
Public Const NEG_NATURE As Single = 0.9

Public Const PhysicalAttack As Byte = 0
Public Const SpecialAttack As Byte = 1
Public Const NeutralAttack As Byte = 2

Public Const Black As Long = 1
Public Const White As Long = 2
Public Const Silver As Long = 3
Public Const DarkGrey As Long = 4
Public Const Red As Long = 5
Public Const Yellow As Long = 6
Public Const Cyan As Long = 7
Public Const Green As Long = 8
Public Const Blue As Long = 9
Public Const Pink As Long = 10

Public Const EndLine As String = "------------------------"

Public Const START_MAP As Long = 1
Public Const START_X As Long = 8
Public Const START_Y As Long = 9

Public Const SELECT_POKEMON As Byte = 1
Public Const SELECT_MOVE As Byte = 2
