<div align="center">

## ApiClipboard


</div>

### Description

Allows you to copy the content of the clipboard to one side, and then restore it at a later time. Useful if you want to swap things in and out of the system clipboard programatically.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Duncan Jones](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/duncan-jones.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/duncan-jones-apiclipboard__1-22024/archive/master.zip)

### API Declarations

```
'\\ --[ApiClipboard]-----------------------------------------------------------
'\\ Extends the Visual basic clipboard object by use of the Api
'\\ ---------------------------------------------------------------------------
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetClipboardOwner Lib "user32" () As Long
Private Declare Function GetClipboardViewer Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function SetClipboardViewer Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
Public Enum enClipboardFormats
  CF_BITMAP = 2
  CF_DIB = 8
  CF_DIF = 5
  CF_ENHMETAFILE = 14
  CF_METAFILEPICT = 3
  CF_OEMTEXT = 7
  CF_PALETTE = 9
  CF_PENDATA = 10
  CF_RIFF = 11
  CF_SYLK = 4
  CF_TEXT = 1
  CF_TIFF = 6
  CF_UNICODETEXT = 13
  CF_WAVE = 12
End Enum
'\\ API Global memory class
'\\ Global memory management functions
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Enum enGlobalmemoryAllocationConstants
  GMEM_FIXED = &H0
  GMEM_DISCARDABLE = &H100
  GMEM_MOVEABLE = &H2
  GMEM_NOCOMPACT = &H10
  GMEM_NODISCARD = &H20
  GMEM_ZEROINIT = &H40
End Enum
```


### Source Code

```
'\\ APIClipboard class ---------------------------
Option Explicit
Public ParenthWnd As Long
Private myMemory As ApiGlobalmemory
Private mLastFormat As Long
Public Property Get BackedUp() As Boolean
  BackedUp = Not (myMemory Is Nothing)
End Property
'\\ --[Backup]------------------------------------------------------
'\\ Makes an in-memory copy of the clipboard's contents so that they
'\\ can be restored easily
'\\ ----------------------------------------------------------------
Public Sub Backup()
Dim lRet As Long
Dim AllFormats As Collection
Dim lFormat As Long
'\\ Need to get all the formats first...
Set AllFormats = Me.ClipboardFormats
lRet = OpenClipboard(ParenthWnd)
If Err.LastDllError > 0 Then
  Call ReportError(Err.LastDllError, "ApiClipboard:Backup", APIDispenser.LastSystemError)
End If
If lRet Then
  If AllFormats.Count > 0 Then
    '\\ Get the first format that holds any data
    For lFormat = 0 To AllFormats.Count - 1
      lRet = GetClipboardData(lFormat)
      If lRet > 0 Then
        Set myMemory = New ApiGlobalmemory
        Call myMemory.CopyFromHandle(lRet)
        '\\ Keep a note of this format
        mLastFormat = lFormat
        Exit For
      End If
      'clipboard
    Next lFormat
  End If
  lRet = CloseClipboard()
End If
End Sub
Public Property Get ClipboardFormats() As Collection
Dim lRet As Long
Dim colFormats As Collection
lRet = OpenClipboard(ParenthWnd)
If Err.LastDllError > 0 Then
  Call ReportError(Err.LastDllError, "ApiClipboard:Backup", APIDispenser.LastSystemError)
End If
If lRet > 0 Then
  Set colFormats = New Collection
  '\\ Get the first available format
  lRet = EnumClipboardFormats(0)
  If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiClipboard:Backup", APIDispenser.LastSystemError)
  End If
  While lRet > 0
    colFormats.Add lRet
    '\\ Get the next available format
    lRet = EnumClipboardFormats(lRet)
    If Err.LastDllError > 0 Then
      Call ReportError(Err.LastDllError, "ApiClipboard:Backup", APIDispenser.LastSystemError)
    End If
  Wend
  '\\ Close the clipboard object to make it available to other apps.
  lRet = CloseClipboard()
End If
Set ClipboardFormats = colFormats
End Property
'\\ --[Restore]-----------------------------------------------------
'\\ Takes the in-memory copy of the clipboard object and restores it
'\\ to the clipboard.
'\\ ----------------------------------------------------------------
Public Sub Restore()
Dim lRet As Long
If Me.BackedUp Then
  lRet = OpenClipboard(ParenthWnd)
  If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "ApiClipboard:Restore", APIDispenser.LastSystemError)
  End If
  If lRet Then
    myMemory.AllocationType = GMEM_FIXED
    lRet = SetClipboardData(mLastFormat, myMemory.Handle)
    myMemory.Free
    If Err.LastDllError > 0 Then
      Call ReportError(Err.LastDllError, "ApiClipboard:Backup", APIDispenser.LastSystemError)
    End If
    lRet = CloseClipboard()
    If Err.LastDllError > 0 Then
      Call ReportError(Err.LastDllError, "ApiClipboard:Backup", APIDispenser.LastSystemError)
    End If
  End If
End If
End Sub
Public Property Get Text() As String
Dim sRet As String
If Clipboard.GetFormat(vbCFText) Then
  sRet = Clipboard.GetText()
End If
End Property
Private Sub Class_Initialize()
End Sub
Private Sub Class_Terminate()
Set myMemory = Nothing
End Sub
'\\ APIGlobalmemory class ------------------------
Option Explicit
Private mMyData() As Byte
Private mMyDataSize As Long
Private mHmem As Long
Private mAllocationType As enGlobalmemoryAllocationConstants
Public Property Let AllocationType(ByVal newType As enGlobalmemoryAllocationConstants)
mAllocationType = newType
End Property
Public Property Get AllocationType() As enGlobalmemoryAllocationConstants
  AllocationType = mAllocationType
End Property
Private Sub CopyDataToGlobal()
Dim lRet As Long
If mHmem > 0 Then
  lRet = GlobalLock(mHmem)
  If lRet > 0 Then
    Call CopyMemory(ByVal mHmem, mMyData(0), mMyDataSize)
    Call GlobalUnlock(mHmem)
  End If
End If
End Sub
Public Sub CopyFromHandle(ByVal hMemHandle As Long)
Dim lRet As Long
Dim lPtr As Long
lRet = GlobalSize(hMemHandle)
If lRet > 0 Then
  mMyDataSize = lRet
  lPtr = GlobalLock(hMemHandle)
  If lPtr > 0 Then
    ReDim mMyData(0 To mMyDataSize - 1) As Byte
    CopyMemory mMyData(0), ByVal lPtr, mMyDataSize
    Call GlobalUnlock(hMemHandle)
  End If
End If
End Sub
Public Sub CopyToHandle(ByVal hMemHandle As Long)
Dim lSize As Long
Dim lPtr As Long
'\\ Don't copy if its empty
If Not Me.IsEmpty Then
  lSize = GlobalSize(hMemHandle)
  '\\ Don't attempt to copy if zero size...
  If lSize > 0 Then
    If lPtr > 0 Then
      CopyMemory ByVal lPtr, mMyData(0), lSize
      Call GlobalUnlock(hMemHandle)
    End If
  End If
End If
End Sub
'\\ --[Handle]------------------------------------------------------
'\\ Returns a Global Memroy handle that is valid and filled with the
'\\ info held in this object's private byte array
'\\ ----------------------------------------------------------------
Public Property Get Handle() As Long
If mHmem = 0 Then
  If mMyDataSize > 0 Then
    mHmem = GlobalAlloc(AllocationType, mMyDataSize)
  End If
End If
Call CopyDataToGlobal
Handle = mHmem
End Property
Public Property Get IsEmpty() As Boolean
  IsEmpty = (mMyDataSize = 0)
End Property
Public Sub Free()
If mHmem > 0 Then
  Call GlobalFree(mHmem)
  mHmem = 0
  mMyDataSize = 0
  ReDim mMyData(0) As Byte
End If
End Sub
Private Sub Class_Terminate()
If mHmem > 0 Then
  Call GlobalFree(mHmem)
End If
End Sub
```

