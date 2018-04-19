VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event KeyDown(KeyCode As Integer, Shift As Integer)

Private Type EVENTMSG
          wMsg   As Long
          lParamLow   As Long
          lParamHigh   As Long
          msgTime   As Long
          hWndMsg   As Long
End Type

Private Const WH_JOURNALRECORD = 0

Private Const WM_KEYDOWN = &H100

Private Declare Function SetWindowsHookEx Lib "user32.dll " Alias "SetWindowsHookExA " (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32.dll " (ByVal hHook As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32.dll " (ByVal vKey As Long) As Integer

Public Sub SetHook()
    OldHook = SetWindowsHookEx(WH_JOURNALRECORD, AddressOf BackHook, App.hInstance, 0)
End Sub

Public Sub UnHook()
    Call UnhookWindowsHookEx(OldHook)
End Sub

Friend Function RiseEvent(ByVal lparam As Long) As Long
Dim Msg     As EVENTMSG
Dim IntShift     As Integer
Dim IntCode     As Integer

  CopyMemory Msg, ByVal lparam, Len(Msg)
  
  IntShift = 0
      Select Case Msg.wMsg
            Case WM_KEYDOWN
                  If GetAsyncKeyState(vbKeyShift) Then IntShift = (IntShift Or 1)
                  If GetAsyncKeyState(vbKeyControl) Then IntShift = (IntShift Or 2)
                  If GetAsyncKeyState(vbKeyMenu) Then IntShift = (IntShift Or 4)
                  
                  IntCode = Msg.lParamLow And &HFF
                  Debug.Print Msg.lParamLow
                  Debug.Print &HFF
                  RaiseEvent KeyDown(IntCode, IntShift)
      End Select
End Function

Private Sub Class_Initialize()
  LngClsPtr = ObjPtr(Me)
  MsgBox LngClsPtr
End Sub


