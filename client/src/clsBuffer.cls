VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private buffer() As Byte
Private buffersize As Long
Private WriteHead As Long
Private ReadHead As Long

Private Sub Class_Initialize()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    Flush
    
    Exit Sub
errHandler:
    HandleError "Class_Initialize", "clsBuffer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub PreAllocate(ByVal nLength As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    WriteHead = 0
    ReadHead = 0
    buffersize = nLength - 1
    ReDim buffer(0 To buffersize)
    
    Exit Sub
errHandler:
    HandleError "PreAllocate", "clsBuffer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub Allocate(ByVal nLength As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If buffersize = 0 And nLength > 1 Then nLength = nLength - 1
    buffersize = buffersize + nLength
    ReDim Preserve buffer(0 To buffersize)
    
    Exit Sub
errHandler:
    HandleError "Allocate", "clsBuffer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub Flush()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    WriteHead = 0
    ReadHead = 0
    buffersize = 0
    ReDim buffer(0)

    Exit Sub
errHandler:
    HandleError "Flush", "clsBuffer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub Trim()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If ReadHead >= count Then
        Flush
    End If
    
    Exit Sub
errHandler:
    HandleError "Trim", "clsBuffer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub WriteByte(ByVal nByte As Byte)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If WriteHead > buffersize Then Allocate 1
 
    buffer(WriteHead) = nByte
    WriteHead = WriteHead + 1
    
    Exit Sub
errHandler:
    HandleError "WriteByte", "clsBuffer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub WriteBytes(ByRef nByte() As Byte)
Dim nLength As Long
        
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    nLength = (UBound(nByte) - LBound(nByte)) + 1
    
    If WriteHead + nLength - 1 > buffersize Then Allocate nLength
    
    CopyMemory buffer(WriteHead), nByte(0), nLength
    WriteHead = WriteHead + nLength
    
    Exit Sub
errHandler:
    HandleError "WriteBytes", "clsBuffer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub WriteInteger(ByVal nInteger As Integer)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If WriteHead + 1 > buffersize Then Allocate 2
        
    CopyMemory buffer(WriteHead), nInteger, 2
    WriteHead = WriteHead + 2
    
    Exit Sub
errHandler:
    HandleError "WriteInteger", "clsBuffer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub WriteLong(ByVal nLong As Long)
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If WriteHead + 3 > buffersize Then Allocate 4
    
    CopyMemory buffer(WriteHead), nLong, 4
    WriteHead = WriteHead + 4
    
    Exit Sub
errHandler:
    HandleError "WriteLong", "clsBuffer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Sub WriteString(ByRef nString As String)
Dim sBytes() As Byte
Dim sLength As Long
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    sLength = Len(nString)
    sBytes = StrConv(nString, vbFromUnicode)
    
    WriteLong sLength
    
    If sLength <= 0 Then Exit Sub
    
    If WriteHead + sLength - 1 > buffersize Then Allocate sLength
      
    CopyMemory buffer(WriteHead), sBytes(0), sLength
    WriteHead = WriteHead + sLength
    
    Exit Sub
errHandler:
    HandleError "WriteString", "clsBuffer", Err.Number, Err.Description
    Err.Clear
    Exit Sub
End Sub

Public Function ReadByte(Optional MoveReadHead As Boolean = True) As Byte
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If ReadHead > buffersize Then Exit Function
    
    ReadByte = buffer(ReadHead)
    If MoveReadHead Then ReadHead = ReadHead + 1
    
    Exit Function
errHandler:
    HandleError "ReadByte", "clsBuffer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function ReadBytes(ByVal nLength As Long, Optional MoveReadHead As Boolean = True) As Byte()
Dim data() As Byte
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If nLength = 0 Then Exit Function
    If ReadHead + nLength - 1 > buffersize Then Exit Function
        
    ReDim data(nLength - 1)
        
    CopyMemory data(0), buffer(ReadHead), nLength
    If MoveReadHead Then ReadHead = ReadHead + nLength
    
    ReadBytes = data
    
    Exit Function
errHandler:
    HandleError "ReadBytes", "clsBuffer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function ReadInteger(Optional MoveReadHead As Boolean = True) As Integer
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If ReadHead + 1 > buffersize Then Exit Function
    
    CopyMemory ReadInteger, buffer(ReadHead), 2
    If MoveReadHead Then ReadHead = ReadHead + 2
    
    Exit Function
errHandler:
    HandleError "ReadInteger", "clsBuffer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function ReadLong(Optional MoveReadHead As Boolean = True) As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    If ReadHead + 3 > buffersize Then Exit Function
    
    CopyMemory ReadLong, buffer(ReadHead), 4
    If MoveReadHead Then ReadHead = ReadHead + 4

    Exit Function
errHandler:
    HandleError "ReadLong", "clsBuffer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function ReadString(Optional MoveReadHead As Boolean = True) As String
Dim sLength As Long
Dim sBytes() As Byte
    
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    sLength = ReadLong(False)
    If sLength <= 0 Then
        If MoveReadHead Then ReadHead = ReadHead + 4
        Exit Function
    End If
    
    ReDim sBytes(sLength - 1)
    
    CopyMemory sBytes(0), buffer(ReadHead + 4), sLength
    
    ReadString = StrConv(sBytes, vbUnicode)
    If MoveReadHead Then ReadHead = ReadHead + sLength + 4

    Exit Function
errHandler:
    HandleError "ReadString", "clsBuffer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function count() As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    count = (UBound(buffer) - LBound(buffer)) + 1

    Exit Function
errHandler:
    HandleError "Count", "clsBuffer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function length() As Long
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    length = count - ReadHead
    
    Exit Function
errHandler:
    HandleError "Length", "clsBuffer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function ToArray() As Byte()
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ToArray = buffer()
    
    Exit Function
errHandler:
    HandleError "ToArray", "clsBuffer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

Public Function ToString() As String
    If Not App.LogMode = 0 Then On Error GoTo errHandler
    
    ToString = StrConv(buffer, vbUnicode)
    
    Exit Function
errHandler:
    HandleError "ToString", "clsBuffer", Err.Number, Err.Description
    Err.Clear
    Exit Function
End Function

