Attribute VB_Name = "modGlobal"
Option Explicit

Public AppOpen As Boolean
Public InMenu As Boolean
Public InGame As Boolean

Public LastButtonClick As Byte
Public WindowVisible(1 To WindowType.Window_Count - 1) As Boolean

Public ChatLine As String * 1
Public user As String
Public Pass As String
Public Pass2 As String
Public SelGender As Byte
Public GenAnim As Byte

Public SaveAccount As Boolean
Public SaveUser As String
Public SavePass As String

Public CurTextBox As Byte

Public CharSelectName(1 To MAX_PLAYER_DATA) As String
Public CharSelectSprite(1 To MAX_PLAYER_DATA) As Long
' CharSelect Level

Public CharSelected As Byte

Public MyIndex As Long
Public HighPlayerIndex As Long

Public GettingMap As Boolean
Public CanMoveNow As Boolean

Public Camera As RECT
Public TileView As RECT

Public IsClicked As Byte
Public CurX As Long
Public CurY As Long
Public GlobalX As Long
Public GlobalY As Long
Public GlobalX_Map As Long
Public GlobalY_Map As Long

Public ChatOn As Boolean
Public ChatMsg As String
Public RenderChatMsg As String

Public PokeIconAnim As Byte

Public ShowTitleBar As Byte
Public TitleBarAlpha As Long
Public ChangeTitleBar As Byte

Public Fade As Boolean
Public FadeType As Byte
Public FadeAlpha As Long

Public TargetAnim As Long
Public TargetSwitch As Long
Public MyTarget As Long

Public EnemyPos As Long
Public ShowMoves As Boolean
Public ForceSwitch As Boolean
Public ShowPokemonSwitch As Boolean
Public CanUseCmd As Boolean
Public CurBarWidth As Long
Public CurEnemyBarWidth As Long
Public CurPoke As Byte

Public ExitBattleTmr As Boolean
Public CanExit As Boolean
Public DidWin As Byte

Public SwitchPokeX As Long
Public SwitchPos As Byte
Public PokeAlpha As Long
Public TmpSwitch As Byte
Public isSwitch As Boolean
Public IsSwitchForce As Boolean

Public IsLearnMove As Boolean
Public LearnPokeNum As Long
Public SelectedMove As Byte
Public LearnMoveNum As Long

Public TmpInEvolve As Byte
Public IsEvolve As Boolean
Public EvolvePoke As Byte
Public TmpEvolveNum As Long
Public TmpCurNum As Long
Public DrawPokeNum As Long
Public EvolveAlpha As Long
Public EvolvePos As Byte

Public IsInvScrollUp As Boolean
Public IsInvScrollDown As Boolean
Public StartInv As Long
Public CurInvType As Byte

Public UpdatingVital As Boolean

Public UseItemNum As Long
Public ShowUseItem As Boolean

Public ShowSelect As Boolean
Public InputData1 As Long
Public InputData2 As Long
Public InputData3 As Long
Public InputData4 As Long
Public OutputData1 As Long
Public OutputData2 As Long
Public OutputData3 As Long
Public OutputData4 As Long
Public MaxSelection As Long

Public BattleRequestIndex As Long

Public IsCapture As Boolean
Public Capture As Long

Public InStorage As Boolean
Public ShowStorageSelect As Boolean
Public StartStorage As Long
Public SelStoragePoke As Long
Public SelStorageX As Long
Public SelStorageY As Long

Public InShop As Long
Public InShopScrollDown As Boolean
Public InShopScrollUp As Boolean
Public ShopStart As Long
Public ShopSelect As Long

Public InTrade As Boolean
Public InTradeConfirm As Boolean
Public InTradeIndex As Long

Public ShowInput As Boolean
Public InputType As Byte
Public InputVal As String
Public RenderVal As String
Public RenderValX As Long
Public RenderValY As Long
Public InputData As Long

Public SelStarter As Long

Public MapBackground As Long
Public MapField As Long

Public ShowTradeSel As Boolean
Public SelTrade As Byte
