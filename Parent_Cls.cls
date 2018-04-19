VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Parent_Cls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'////////////////////////////////////////////////////////////////
'
'插件名称：
'
'插件作者：人闲花落 QQ：449806776
'
'版权声明：您可以修改或共享发布此插件，但必须注明原创作者信息
'
'VB爱好者：交流QQ群——19871152
'
'////////////////////////////////////////////////////////////////
Option Explicit

Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Public Sub SetMdi(Fhw As Long, Zhw As Long, SetMove As String) '将自动代码列表窗口 设置为代码窗口的子窗体
Dim Rt() As String, rctemp As RECT
    
    OldParent = SetParent(Zhw, Fhw)              '指定本程序为打开程序的父窗口
    Rt = Split(SetMove, ",") '示例："10,10,200,300,"
    GetWindowRect Zhw, rctemp                 '指定子程序窗体位置
    MoveWindow Val(Zhw), Val(Rt(0)), Val(Rt(1)), Val(Rt(2)), Val(Rt(3)), 1

End Sub

Public Sub OutMdi(Zhw As Long) '释放子窗体  自动代码列表窗口
    If OldParent Then
        SetParent Zhw, OldParent
        OldParent = 0
    End If
End Sub

