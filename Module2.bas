Attribute VB_Name = "modHook"
Option Explicit

Public Declare Function CallNextHookEx Lib "user32.dll " (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lparam As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32 " Alias "RtlMoveMemory " (lpDest As Any, lpSource As Any, ByVal cBytes As Long)
Public Declare Sub keybd_event Lib "user32 " (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)


Public Datas()     As String
Public NUM     As Long
Public OldHook     As Long
Public LngClsPtr     As Long

Public Function BackHook(ByVal nCode As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
  If nCode < 0 Then
      BackHook = CallNextHookEx(OldHook, nCode, wParam, lparam)
      Exit Function
  End If
  
  ResolvePointer(LngClsPtr).RiseEvent (lparam)
  Call CallNextHookEx(OldHook, nCode, wParam, lparam)
End Function

Private Function ResolvePointer(ByVal lpObj As Long) As ClsHook

    Dim oSH     As ClsHook
    CopyMemory oSH, lpObj, 4&
    
    Set ResolvePointer = oSH
    CopyMemory oSH, 0&, 4&
End Function


