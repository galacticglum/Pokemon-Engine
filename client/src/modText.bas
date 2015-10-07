Attribute VB_Name = "modText"
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type CharVA
    Vertex(0 To 3) As TLVERTEX
End Type

Private Type VFH
    BitmapWidth As Long
    BitmapHeight As Long
    CellWidth As Long
    CellHeight As Long
    BaseCharOffset As Byte
    CharWidth(0 To 255) As Byte
    CharVA(0 To 255) As CharVA
End Type

Private Type CustomFont
    HeaderInfo As VFH
    Texture As Direct3DTexture8
    RowPitch As Integer
    RowFactor As Single
    ColFactor As Single
    CharHeight As Byte
    TextureSize As POINTAPI
End Type

Public CurFont As CustomFont

Public Font_Georgia As CustomFont
Private Const Font_Path As String = "\bin\font\"

Public Black As Long
Public White As Long
Public Silver As Long
Public DarkGrey As Long
Public Red As Long
Public Yellow As Long
Public Cyan As Long
Public Green As Long
Public Blue As Long
Public Pink As Long

Public ChatVA() As TLVERTEX
Public ChatVAS() As TLVERTEX

Public Const ChatTextBufferSize As Integer = 200
Public ChatBufferChunk As Single

Public Type ChatTextBuffer
    Text As String
    color As Long
End Type

Public ChatArrayUbound As Long
Public ChatVB As Direct3DVertexBuffer8
Public ChatVBS As Direct3DVertexBuffer8
Public ChatTextBuffer(1 To ChatTextBufferSize) As ChatTextBuffer

Public Const ChatWidth As Long = 360
Public Const MaxChatLine As Long = 7
Public Const MaxMsg As Long = 320

Public ChatScroll As Long
Public ChatScrollUp As Boolean
Public ChatScrollDown As Boolean
Public totalChatLines As Long

Public BattleVA() As TLVERTEX
Public BattleVAS() As TLVERTEX

Public Const BattleTextBufferSize As Integer = 200
Public BattleBufferChunk As Single
Public BattleArrayUbound As Long
Public BattleVB As Direct3DVertexBuffer8
Public BattleVBS As Direct3DVertexBuffer8
Public BattleTextBuffer(1 To ChatTextBufferSize) As ChatTextBuffer

Public Const BattleWidth As Long = 190
Public Const MaxBattleLine As Long = 21

Public BattleScroll As Long
Public BattleScrollUp As Boolean
Public BattleScrollDown As Boolean
Public totalBattleLines As Long

Public Sub EngineInitFontTextures()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If D3DDevice8.TestCooperativeLevel <> D3D_OK Then Exit Sub

    Set Font_Georgia.Texture = Direct3DX8.CreateTextureFromFileEx(D3DDevice8, App.Path & Font_Path & "georgia" & Gfx_Ext, 256, 256, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFF000000, ByVal 0, ByVal 0)
    Font_Georgia.TextureSize.X = 256
    Font_Georgia.TextureSize.Y = 256
    
    Exit Sub
errHandler:
    HandleError "EngineInitFontTextures", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub EngineInitFontSettings()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    LoadFontHeader Font_Georgia, "georgia.dat"
    
    Exit Sub
errHandler:
    HandleError "EngineInitFontSettings", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub EngineInitFontColors()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Black = D3DColorARGB(255, 0, 0, 0)
    White = D3DColorARGB(255, 255, 255, 255)
    Silver = D3DColorARGB(255, 192, 192, 192)
    DarkGrey = D3DColorARGB(255, 102, 102, 102)
    Red = D3DColorARGB(255, 201, 0, 0)
    Yellow = D3DColorARGB(255, 255, 255, 0)
    Cyan = D3DColorARGB(255, 16, 224, 237)
    Green = D3DColorARGB(255, 119, 188, 84)
    Blue = D3DColorARGB(255, 16, 104, 237)
    Pink = D3DColorARGB(255, 255, 118, 221)
    
    Exit Sub
errHandler:
    HandleError "EngineInitFontColors", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub LoadFontHeader(ByRef theFont As CustomFont, ByVal FileName As String)
Dim FileNum As Byte
Dim LoopChar As Long
Dim Row As Single
Dim u As Single
Dim v As Single

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    FileNum = FreeFile
    Open App.Path & Font_Path & FileName For Binary As #FileNum
        Get #FileNum, , theFont.HeaderInfo
    Close #FileNum
    
    theFont.CharHeight = theFont.HeaderInfo.CellHeight - 4
    theFont.RowPitch = theFont.HeaderInfo.BitmapWidth \ theFont.HeaderInfo.CellWidth
    theFont.ColFactor = theFont.HeaderInfo.CellWidth / theFont.HeaderInfo.BitmapWidth
    theFont.RowFactor = theFont.HeaderInfo.CellHeight / theFont.HeaderInfo.BitmapHeight
    
    For LoopChar = 0 To 255
        Row = (LoopChar - theFont.HeaderInfo.BaseCharOffset) \ theFont.RowPitch
        u = ((LoopChar - theFont.HeaderInfo.BaseCharOffset) - (Row * theFont.RowPitch)) * theFont.ColFactor
        v = Row * theFont.RowFactor
        
        With theFont.HeaderInfo.CharVA(LoopChar)
            .Vertex(0).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(0).RHW = 1
            .Vertex(0).tu = u
            .Vertex(0).tv = v
            .Vertex(0).X = 0
            .Vertex(0).Y = 0
            .Vertex(0).z = 0
            .Vertex(1).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).RHW = 1
            .Vertex(1).tu = u + theFont.ColFactor
            .Vertex(1).tv = v
            .Vertex(1).X = theFont.HeaderInfo.CellWidth
            .Vertex(1).Y = 0
            .Vertex(1).z = 0
            .Vertex(2).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).RHW = 1
            .Vertex(2).tu = u
            .Vertex(2).tv = v + theFont.RowFactor
            .Vertex(2).X = 0
            .Vertex(2).Y = theFont.HeaderInfo.CellHeight
            .Vertex(2).z = 0
            .Vertex(3).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).RHW = 1
            .Vertex(3).tu = u + theFont.ColFactor
            .Vertex(3).tv = v + theFont.RowFactor
            .Vertex(3).X = theFont.HeaderInfo.CellWidth
            .Vertex(3).Y = theFont.HeaderInfo.CellHeight
            .Vertex(3).z = 0
        End With
    Next LoopChar
    
    Exit Sub
errHandler:
    HandleError "LoadFontHeader", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub RenderText(ByRef theFont As CustomFont, ByVal Text As String, ByVal X As Long, ByVal Y As Long, ByVal color As Long)
Dim TempVA(0 To 3)  As TLVERTEX
Dim TempVAS(0 To 3) As TLVERTEX
Dim TempStr() As String
Dim Count As Integer
Dim Ascii() As Byte
Dim Row As Integer
Dim u As Single
Dim v As Single
Dim i As Long
Dim j As Long
Dim KeyPhrase As Byte
Dim TempColor As Long
Dim ResetColor As Byte
Dim srcRECT As RECT
Dim v2 As D3DVECTOR2
Dim v3 As D3DVECTOR2
Dim yOffset As Single

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If LenB(Text) = 0 Then Exit Sub
    
    TempStr = Split(Text, vbCrLf)

    TempColor = color
    
    D3DDevice8.SetTexture 0, theFont.Texture
    CurrentTexture = -1
    
    For i = 0 To UBound(TempStr)
        If Len(TempStr(i)) > 0 Then
            yOffset = i * theFont.CharHeight
            Count = 0
            Ascii() = StrConv(TempStr(i), vbFromUnicode)
            
            For j = 1 To Len(TempStr(i))
                Call CopyMemory(TempVA(0), theFont.HeaderInfo.CharVA(Ascii(j - 1)).Vertex(0), FVF_Size * 4)
                
                TempVA(0).X = X + Count
                TempVA(0).Y = Y + yOffset
                TempVA(1).X = TempVA(1).X + X + Count
                TempVA(1).Y = TempVA(0).Y
                TempVA(2).X = TempVA(0).X
                TempVA(2).Y = TempVA(2).Y + TempVA(0).Y
                TempVA(3).X = TempVA(1).X
                TempVA(3).Y = TempVA(2).Y
                
                TempVA(0).color = TempColor
                TempVA(1).color = TempColor
                TempVA(2).color = TempColor
                TempVA(3).color = TempColor

                Call D3DDevice8.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, TempVA(0), FVF_Size)
                
                Count = Count + theFont.HeaderInfo.CharWidth(Ascii(j - 1))
                
                If ResetColor Then
                    ResetColor = 0
                    TempColor = color
                End If
            Next j
        End If
    Next i
    
    Exit Sub
errHandler:
    HandleError "RenderText", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function GetWidth(ByRef theFont As CustomFont, ByVal Text As String) As Integer
Dim LoopI As Integer

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If LenB(Text) = 0 Then Exit Function
    
    For LoopI = 1 To Len(Text)
        GetWidth = GetWidth + theFont.HeaderInfo.CharWidth(Asc(Mid$(Text, LoopI, 1)))
    Next LoopI
    
    Exit Function
errHandler:
    HandleError "GetWidth", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function CensorWord(ByVal sString As String) As String
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    CensorWord = String(Len(sString), "*")
    
    Exit Function
errHandler:
    HandleError "GetWidth", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function EnterText(ByVal tempString As String, ByVal KeyAscii As Integer) As String
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If (KeyAscii = vbKeyBack) Then
        If LenB(tempString) > 0 Then tempString = Mid$(tempString, 1, Len(tempString) - 1)
    End If
    If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) Then
        tempString = tempString & ChrW$(KeyAscii)
    End If
    EnterText = tempString
    
    Exit Function
errHandler:
    HandleError "EnterText", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function isNameLegal(ByVal KeyAscii As Integer) As Boolean
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 95) Or (KeyAscii = 32) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        isNameLegal = True
    End If
    
    Exit Function
errHandler:
    HandleError "isNameLegal", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function CheckNameInput(ByVal Name As String) As Boolean
Dim i As Long, n As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Len(Name) <= 2 Or Len(Name) > (MAX_STRING - 1) Then
        CheckNameInput = False
        Exit Function
    End If
    
    For i = 1 To Len(Name)
        n = AscW(Mid$(Name, i, 1))

        If Not isNameLegal(n) Then
            CheckNameInput = False
            Exit Function
        End If
    Next
    CheckNameInput = True
    
    Exit Function
errHandler:
    HandleError "CheckNameInput", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub DrawPlayerName(ByVal Index As Long)
Dim X As Long, Y As Long
Dim Text As String, TextSize As Long, color As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If Index <= 0 Or Index > HighPlayerIndex Then Exit Sub

    Text = GetPlayerName(Index)
    color = White
    
    If Player(Index).IsVIP = YES Then
        Text = "[VIP] " & Text
        color = Yellow
    End If
    
    TextSize = GetWidth(CurFont, Text)
    
    Select Case GetPlayerAccess(Index)
        Case ACCESS_MODERATOR
            color = Silver
        Case ACCESS_MAPPER
            color = Cyan
        Case ACCESS_DEVELOPER
            color = Blue
        Case ACCESS_ADMIN
            color = Green
    End Select
    
    X = GetPlayerX(Index) * Pic_Size + Player(Index).xOffset + (Pic_Size / 2) - (TextSize / 2)
    If GetPlayerSprite(Index) >= 1 And GetPlayerSprite(Index) <= Count_Sprite Then
        Y = GetPlayerY(Index) * Pic_Size + Player(Index).yOffset - (GetTextureHeight(Tex_Sprite(GetPlayerSprite(Index))) / 4) + 12
    Else
        Y = GetPlayerY(Index) * Pic_Size + Player(Index).yOffset - 32
    End If
    
    RenderText CurFont, Text, ConvertMapX(X), ConvertMapY(Y), color

    Exit Sub
errHandler:
    HandleError "DrawPlayerName", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawNpcName(ByVal MapNpcNum As Long)
Dim X As Long, Y As Long
Dim Text As String, TextSize As Long, color As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPC Then Exit Sub

    Text = Trim$(NPC(MapNpc(MapNpcNum).Num).Name)
    TextSize = GetWidth(CurFont, Text)
    
    color = Silver
    
    X = MapNpc(MapNpcNum).X * Pic_Size + MapNpc(MapNpcNum).xOffset + (Pic_Size / 2) - (TextSize / 2)
    If NPC(MapNpc(MapNpcNum).Num).Sprite >= 1 And NPC(MapNpc(MapNpcNum).Num).Sprite <= Count_Sprite Then
        Y = MapNpc(MapNpcNum).Y * Pic_Size + MapNpc(MapNpcNum).yOffset - (GetTextureHeight(Tex_Sprite(NPC(MapNpc(MapNpcNum).Num).Sprite)) / 4) + 12
    Else
        Y = MapNpc(MapNpcNum).Y * Pic_Size + MapNpc(MapNpcNum).yOffset - 32
    End If
    
    RenderText CurFont, Text, ConvertMapX(X), ConvertMapY(Y), color

    Exit Sub
errHandler:
    HandleError "DrawNPCName", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function UpdateChatText(ByVal Text As String, ByVal MaxWidth As Long) As String
Dim i As Long, X As Long
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    If GetWidth(CurFont, Text) > MaxWidth Then
        For i = Len(Text) To 1 Step -1
            X = X + CurFont.HeaderInfo.CharWidth(Asc(Mid$(Text, i, 1)))
            If X > MaxWidth Then
                UpdateChatText = Right$(Text, Len(Text) - i + 1)
                Exit For
            End If
        Next
    Else
        UpdateChatText = Text
    End If
    
    Exit Function
errHandler:
    HandleError "UpdateChatText", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub RenderChatTextBuffer()
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    D3DDevice8.SetTexture 0, CurFont.Texture
    CurrentTexture = -1

    If ChatArrayUbound > 0 Then
        D3DDevice8.SetStreamSource 0, ChatVBS, FVF_Size
        D3DDevice8.DrawPrimitive D3DPT_TRIANGLELIST, 0, (ChatArrayUbound + 1) \ 3
        D3DDevice8.SetStreamSource 0, ChatVB, FVF_Size
        D3DDevice8.DrawPrimitive D3DPT_TRIANGLELIST, 0, (ChatArrayUbound + 1) \ 3
    End If
    
    Exit Sub
errHandler:
    HandleError "RenderChatTextBuffer", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateChatArray()
Dim Chunk As Integer
Dim Count As Integer
Dim LoopC As Byte
Dim Ascii As Byte
Dim Row As Long
Dim pos As Long
Dim u As Single
Dim v As Single
Dim X As Single
Dim Y As Single
Dim y2 As Single
Dim i As Long
Dim j As Long
Dim Size As Integer
Dim KeyPhrase As Byte
Dim ResetColor As Byte
Dim TempColor As Long
Dim yOffset As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    yOffset = 14

    If ChatBufferChunk <= 1 Then ChatBufferChunk = 1
    Chunk = ChatScroll
    Size = 0

    For LoopC = (Chunk * ChatBufferChunk) - (MaxChatLine - 1) To Chunk * ChatBufferChunk
        If LoopC > ChatTextBufferSize Then Exit For
        Size = Size + Len(ChatTextBuffer(LoopC).Text)
    Next
    
    Size = Size - j
    ChatArrayUbound = Size * 6 - 1
    If ChatArrayUbound < 0 Then Exit Sub
    ReDim ChatVA(0 To ChatArrayUbound)
    ReDim ChatVAS(0 To ChatArrayUbound)
    
    X = GuiChatboxX + 20
    Y = GuiChatboxY + 45

    For LoopC = (Chunk * ChatBufferChunk) - (MaxChatLine - 1) To Chunk * ChatBufferChunk
        If LoopC > ChatTextBufferSize Then Exit For
        If ChatBufferChunk * Chunk > ChatTextBufferSize Then ChatBufferChunk = ChatBufferChunk - 1
        
        TempColor = ChatTextBuffer(LoopC).color
        y2 = Y - (LoopC * yOffset) + (Chunk * ChatBufferChunk * yOffset) - 32
        Count = 0
        If LenB(ChatTextBuffer(LoopC).Text) <> 0 Then
            For j = 1 To Len(ChatTextBuffer(LoopC).Text)
                Ascii = Asc(Mid$(ChatTextBuffer(LoopC).Text, j, 1))
                Row = (Ascii - CurFont.HeaderInfo.BaseCharOffset) \ CurFont.RowPitch
                u = ((Ascii - CurFont.HeaderInfo.BaseCharOffset) - (Row * CurFont.RowPitch)) * CurFont.ColFactor
                v = Row * CurFont.RowFactor

                With ChatVA(0 + (6 * pos))
                    .color = TempColor
                    .X = (X) + Count
                    .Y = (y2)
                    .tu = u
                    .tv = v
                    .RHW = 1
                End With
                With ChatVA(1 + (6 * pos))
                    .color = TempColor
                    .X = (X) + Count
                    .Y = (y2) + CurFont.HeaderInfo.CellHeight
                    .tu = u
                    .tv = v + CurFont.RowFactor
                    .RHW = 1
                End With
                With ChatVA(2 + (6 * pos))
                    .color = TempColor
                    .X = (X) + Count + CurFont.HeaderInfo.CellWidth
                    .Y = (y2) + CurFont.HeaderInfo.CellHeight
                    .tu = u + CurFont.ColFactor
                    .tv = v + CurFont.RowFactor
                    .RHW = 1
                End With
                ChatVA(3 + (6 * pos)) = ChatVA(0 + (6 * pos))
                With ChatVA(4 + (6 * pos))
                    .color = TempColor
                    .X = (X) + Count + CurFont.HeaderInfo.CellWidth
                    .Y = (y2)
                    .tu = u + CurFont.ColFactor
                    .tv = v
                    .RHW = 1
                End With
                ChatVA(5 + (6 * pos)) = ChatVA(2 + (6 * pos))
                pos = pos + 1
                Count = Count + CurFont.HeaderInfo.CharWidth(Ascii)
                If ResetColor Then
                    ResetColor = 0
                    TempColor = ChatTextBuffer(LoopC).color
                End If
            Next
        End If
    Next LoopC
        
    If Not D3DDevice8 Is Nothing Then
        Set ChatVBS = D3DDevice8.CreateVertexBuffer(FVF_Size * pos * 6, 0, FVF, D3DPOOL_MANAGED)
        D3DVertexBuffer8SetData ChatVBS, 0, FVF_Size * pos * 6, 0, ChatVAS(0)
        Set ChatVB = D3DDevice8.CreateVertexBuffer(FVF_Size * pos * 6, 0, FVF, D3DPOOL_MANAGED)
        D3DVertexBuffer8SetData ChatVB, 0, FVF_Size * pos * 6, 0, ChatVA(0)
    End If
    Erase ChatVAS()
    Erase ChatVA()
    
    Exit Sub
errHandler:
    HandleError "UpdateChatArray", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub AddText(ByVal Text As String, ByVal color As Long)
Dim TempSplit() As String
Dim TSLoop As Long
Dim lastSpace As Long
Dim Size As Long
Dim i As Long
Dim b As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    TempSplit = Split(Text, vbCrLf)
    
    For TSLoop = 0 To UBound(TempSplit)
        Size = 0
        b = 1
        lastSpace = 1
        
        For i = 1 To Len(TempSplit(TSLoop))
            Select Case Mid$(TempSplit(TSLoop), i, 1)
                Case " ": lastSpace = i
                Case "_": lastSpace = i
                Case "-": lastSpace = i
            End Select
            
            Size = Size + CurFont.HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), i, 1)))
            
            If Size > ChatWidth Then
                If i - lastSpace > 10 Then
                    AddToChatTextBuffer_Overflow Trim$(Mid$(TempSplit(TSLoop), b, (i - 1) - b)), color
                    b = i - 1
                    Size = 0
                Else
                    AddToChatTextBuffer_Overflow Trim$(Mid$(TempSplit(TSLoop), b, lastSpace - b)), color
                    b = lastSpace + 1
                    Size = GetWidth(CurFont, Mid$(TempSplit(TSLoop), lastSpace, i - lastSpace))
                End If
            End If
            
            If i = Len(TempSplit(TSLoop)) Then
                If b <> i Then AddToChatTextBuffer_Overflow Mid$(TempSplit(TSLoop), b, i), color
            End If
        Next i
    Next TSLoop
    
    If CurFont.RowPitch = 0 Then Exit Sub
    If ChatScroll > MaxChatLine Then ChatScroll = ChatScroll + 1
    UpdateChatArray
    
    Exit Sub
errHandler:
    HandleError "AddText", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub AddToChatTextBuffer_Overflow(ByVal Text As String, ByVal color As Long)
Dim LoopC As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For LoopC = (ChatTextBufferSize - 1) To 1 Step -1
        ChatTextBuffer(LoopC + 1) = ChatTextBuffer(LoopC)
    Next LoopC
    
    ChatTextBuffer(1).Text = Text
    ChatTextBuffer(1).color = color
    
    totalChatLines = totalChatLines + 1
    If totalChatLines > ChatTextBufferSize - 1 Then totalChatLines = ChatTextBufferSize - 1
    
    Exit Sub
errHandler:
    HandleError "AddToChatTextBuffer_Overflow", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ScrollChatBox(ByVal Direction As Byte)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If totalChatLines < MaxChatLine Then
        ChatScroll = MaxChatLine
        UpdateChatArray
        Exit Sub
    End If
    If Direction = 0 Then
        ChatScroll = ChatScroll + 1
    Else
        ChatScroll = ChatScroll - 1
    End If
    If ChatScroll < MaxChatLine Then ChatScroll = MaxChatLine
    If ChatScroll > totalChatLines Then ChatScroll = totalChatLines
    UpdateChatArray
    
    Exit Sub
errHandler:
    HandleError "ScrollChatBox", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function CheckColor(ByVal Num As Long) As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Select Case Num
        Case 1: CheckColor = Black
        Case 2: CheckColor = White
        Case 3: CheckColor = Silver
        Case 4: CheckColor = DarkGrey
        Case 5: CheckColor = Red
        Case 6: CheckColor = Yellow
        Case 7: CheckColor = Cyan
        Case 8: CheckColor = Green
        Case 9: CheckColor = Blue
        Case 10: CheckColor = Pink
        Case Else: CheckColor = Black
    End Select
    
    Exit Function
errHandler:
    HandleError "CheckColor", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function GetAttributeColor(ByVal AttributeNum As Attributes) As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Select Case AttributeNum
        Case Attributes.Blocked
            GetAttributeColor = D3DColorARGB(150, 50, 50, 50)
        Case Attributes.TallGrass
            GetAttributeColor = D3DColorARGB(100, 0, 255, 0)
        Case Attributes.Heal
            GetAttributeColor = D3DColorARGB(100, 255, 0, 255)
        Case Attributes.Checkpoint
            GetAttributeColor = D3DColorARGB(100, 0, 0, 255)
        Case Attributes.Storage
            GetAttributeColor = D3DColorARGB(100, 150, 75, 255)
        Case Attributes.Warp
            GetAttributeColor = D3DColorARGB(100, 0, 255, 255)
        Case Attributes.mShop
            GetAttributeColor = D3DColorARGB(100, 255, 255, 120)
        Case Else
            GetAttributeColor = D3DColorARGB(0, 255, 255, 255)
    End Select
    
    Exit Function
errHandler:
    HandleError "CheckColor", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Sub RenderBattleTextBuffer()
    If Not App.LogMode = 0 Then On Error GoTo errHandler

    D3DDevice8.SetTexture 0, CurFont.Texture
    CurrentTexture = -1

    If BattleArrayUbound > 0 Then
        D3DDevice8.SetStreamSource 0, BattleVBS, FVF_Size
        D3DDevice8.DrawPrimitive D3DPT_TRIANGLELIST, 0, (BattleArrayUbound + 1) \ 3
        D3DDevice8.SetStreamSource 0, BattleVB, FVF_Size
        D3DDevice8.DrawPrimitive D3DPT_TRIANGLELIST, 0, (BattleArrayUbound + 1) \ 3
    End If
    
    Exit Sub
errHandler:
    HandleError "RenderBattleTextBuffer", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateBattleArray()
Dim Chunk As Integer
Dim Count As Integer
Dim LoopC As Byte
Dim Ascii As Byte
Dim Row As Long
Dim pos As Long
Dim u As Single
Dim v As Single
Dim X As Single
Dim Y As Single
Dim y2 As Single
Dim i As Long
Dim j As Long
Dim Size As Integer
Dim KeyPhrase As Byte
Dim ResetColor As Byte
Dim TempColor As Long
Dim yOffset As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    yOffset = 14

    If BattleBufferChunk <= 1 Then BattleBufferChunk = 1
    Chunk = BattleScroll
    Size = 0

    For LoopC = (Chunk * BattleBufferChunk) - (MaxBattleLine - 1) To Chunk * BattleBufferChunk
        If LoopC > BattleTextBufferSize Then Exit For
        Size = Size + Len(BattleTextBuffer(LoopC).Text)
    Next
    
    Size = Size - j
    BattleArrayUbound = Size * 6 - 1
    If BattleArrayUbound < 0 Then Exit Sub
    ReDim BattleVA(0 To BattleArrayUbound)
    ReDim BattleVAS(0 To BattleArrayUbound)
    
    X = GuiBattleX + 541
    Y = GuiBattleY + 118

    For LoopC = (Chunk * BattleBufferChunk) - (MaxBattleLine - 1) To Chunk * BattleBufferChunk
        If LoopC > BattleTextBufferSize Then Exit For
        If BattleBufferChunk * Chunk > BattleTextBufferSize Then BattleBufferChunk = BattleBufferChunk - 1
        
        TempColor = BattleTextBuffer(LoopC).color
        y2 = Y - (LoopC * yOffset) + (Chunk * BattleBufferChunk * yOffset) - 32
        Count = 0
        If LenB(BattleTextBuffer(LoopC).Text) <> 0 Then
            For j = 1 To Len(BattleTextBuffer(LoopC).Text)
                Ascii = Asc(Mid$(BattleTextBuffer(LoopC).Text, j, 1))
                Row = (Ascii - CurFont.HeaderInfo.BaseCharOffset) \ CurFont.RowPitch
                u = ((Ascii - CurFont.HeaderInfo.BaseCharOffset) - (Row * CurFont.RowPitch)) * CurFont.ColFactor
                v = Row * CurFont.RowFactor

                With BattleVA(0 + (6 * pos))
                    .color = TempColor
                    .X = (X) + Count
                    .Y = (y2)
                    .tu = u
                    .tv = v
                    .RHW = 1
                End With
                With BattleVA(1 + (6 * pos))
                    .color = TempColor
                    .X = (X) + Count
                    .Y = (y2) + CurFont.HeaderInfo.CellHeight
                    .tu = u
                    .tv = v + CurFont.RowFactor
                    .RHW = 1
                End With
                With BattleVA(2 + (6 * pos))
                    .color = TempColor
                    .X = (X) + Count + CurFont.HeaderInfo.CellWidth
                    .Y = (y2) + CurFont.HeaderInfo.CellHeight
                    .tu = u + CurFont.ColFactor
                    .tv = v + CurFont.RowFactor
                    .RHW = 1
                End With
                BattleVA(3 + (6 * pos)) = BattleVA(0 + (6 * pos))
                With BattleVA(4 + (6 * pos))
                    .color = TempColor
                    .X = (X) + Count + CurFont.HeaderInfo.CellWidth
                    .Y = (y2)
                    .tu = u + CurFont.ColFactor
                    .tv = v
                    .RHW = 1
                End With
                BattleVA(5 + (6 * pos)) = BattleVA(2 + (6 * pos))
                pos = pos + 1
                Count = Count + CurFont.HeaderInfo.CharWidth(Ascii)
                If ResetColor Then
                    ResetColor = 0
                    TempColor = BattleTextBuffer(LoopC).color
                End If
            Next
        End If
    Next LoopC
        
    If Not D3DDevice8 Is Nothing Then
        Set BattleVBS = D3DDevice8.CreateVertexBuffer(FVF_Size * pos * 6, 0, FVF, D3DPOOL_MANAGED)
        D3DVertexBuffer8SetData BattleVBS, 0, FVF_Size * pos * 6, 0, BattleVAS(0)
        Set BattleVB = D3DDevice8.CreateVertexBuffer(FVF_Size * pos * 6, 0, FVF, D3DPOOL_MANAGED)
        D3DVertexBuffer8SetData BattleVB, 0, FVF_Size * pos * 6, 0, BattleVA(0)
    End If
    Erase BattleVAS()
    Erase BattleVA()
    
    Exit Sub
errHandler:
    HandleError "UpdateBattleArray", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub AddBattleLog(ByVal Text As String, ByVal color As Long)
Dim TempSplit() As String
Dim TSLoop As Long
Dim lastSpace As Long
Dim Size As Long
Dim i As Long
Dim b As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    TempSplit = Split(Text, vbCrLf)
    
    For TSLoop = 0 To UBound(TempSplit)
        Size = 0
        b = 1
        lastSpace = 1
        
        For i = 1 To Len(TempSplit(TSLoop))
            Select Case Mid$(TempSplit(TSLoop), i, 1)
                Case " ": lastSpace = i
                Case "_": lastSpace = i
                Case "-": lastSpace = i
            End Select
            
            Size = Size + CurFont.HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), i, 1)))
            
            If Size > BattleWidth Then
                If i - lastSpace > 10 Then
                    AddToBattleTextBuffer_Overflow Trim$(Mid$(TempSplit(TSLoop), b, (i - 1) - b)), color
                    b = i - 1
                    Size = 0
                Else
                    AddToBattleTextBuffer_Overflow Trim$(Mid$(TempSplit(TSLoop), b, lastSpace - b)), color
                    b = lastSpace + 1
                    Size = GetWidth(CurFont, Mid$(TempSplit(TSLoop), lastSpace, i - lastSpace))
                End If
            End If
            
            If i = Len(TempSplit(TSLoop)) Then
                If b <> i Then AddToBattleTextBuffer_Overflow Mid$(TempSplit(TSLoop), b, i), color
            End If
        Next i
    Next TSLoop
    
    If CurFont.RowPitch = 0 Then Exit Sub
    If BattleScroll > MaxBattleLine Then BattleScroll = BattleScroll + 1
    UpdateBattleArray
    
    Exit Sub
errHandler:
    HandleError "AddBattleLog", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Private Sub AddToBattleTextBuffer_Overflow(ByVal Text As String, ByVal color As Long)
Dim LoopC As Long

    If Not App.LogMode = 0 Then On Error GoTo errHandler

    For LoopC = (BattleTextBufferSize - 1) To 1 Step -1
        BattleTextBuffer(LoopC + 1) = BattleTextBuffer(LoopC)
    Next LoopC
    
    BattleTextBuffer(1).Text = Text
    BattleTextBuffer(1).color = color
    
    totalBattleLines = totalBattleLines + 1
    If totalBattleLines > BattleTextBufferSize - 1 Then totalBattleLines = BattleTextBufferSize - 1
    
    Exit Sub
errHandler:
    HandleError "AddToBattleTextBuffer_Overflow", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub ScrollBattleBox(ByVal Direction As Byte)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If totalBattleLines < MaxBattleLine Then
        BattleScroll = MaxBattleLine
        UpdateBattleArray
        Exit Sub
    End If
    If Direction = 0 Then
        BattleScroll = BattleScroll + 1
    Else
        BattleScroll = BattleScroll - 1
    End If
    If BattleScroll < MaxBattleLine Then BattleScroll = MaxBattleLine
    If BattleScroll > totalBattleLines Then BattleScroll = totalBattleLines
    UpdateBattleArray
    
    Exit Sub
errHandler:
    HandleError "ScrollBattleBox", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function KeepTwoDigit(ByVal Val As Long) As String
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If Val > 9 Then
        KeepTwoDigit = Val
    Else
        KeepTwoDigit = "0" & Val
    End If
    
    Exit Function
errHandler:
    HandleError "KeepTwoDigit", "modText", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function
