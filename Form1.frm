VERSION 5.00
Begin VB.Form JS_Frm 
   Caption         =   "Form1"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5865
      Top             =   1530
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3315
      Top             =   4110
   End
   Begin VB.ListBox List2 
      Height          =   2040
      Left            =   90
      TabIndex        =   8
      Top             =   2610
      Width           =   3780
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   4110
      TabIndex        =   7
      Top             =   525
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   1950
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2070
      Width           =   2355
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5910
      Top             =   30
   End
   Begin VB.CommandButton Command1 
      Caption         =   "清空消息监视列表"
      Height          =   450
      Left            =   4245
      TabIndex        =   1
      Top             =   4125
      Width           =   1905
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3780
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "当前代码窗口："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4065
      TabIndex        =   6
      Top             =   225
      Width           =   1260
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "代码窗口字符跟踪："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4065
      TabIndex        =   4
      Top             =   1755
      Width           =   1620
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x=：y="
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   4080
      TabIndex        =   3
      Top             =   1410
      Width           =   2400
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "代码窗口光标位置："
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4065
      TabIndex        =   2
      Top             =   1080
      Width           =   1620
   End
End
Attribute VB_Name = "JS_Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private Declare Function GetCaretPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
x As Long '点在X坐标(横坐标)上的坐标值
y As Long '点在Y坐标(纵坐标)上的坐标值
End Type
'光标位置。
Dim lpPoint As POINTAPI
Dim iKey As String

Private Sub Command1_Click()
    List1.Clear
    List2.Clear
End Sub

Private Sub Form_Load()
    Me.Hide
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim ls As String, h As Long

  If InStr(1, Tkey, "|" & Chr(KeyAscii) & "|") <> 0 Then
        'If AutoCodeFrm.Visible = True And Chr(KeyAscii) = "(" Then 'And AutoCodeFrm.ATlv.SelectedItem.SmallIcon = 1
        '    iKey = "{LEFT}{BACKSPACE " & Len(JS_Frm.Text1) + 1 & "}" & AutoCodeFrm.ATlv.SelectedItem.Text & "{" & Chr(KeyAscii) & "}"
        '    Timer3.Enabled = True
        'End If
        Call NoTextInput '清空输入文本，停止AutoCode
  Else
        'Me.Caption = KeyAscii
        If Len(Text1.Text) = 0 And KeyAscii = 32 Then KeyAscii = 0: Exit Sub
        
        If Trim(Text1) = "" And KeyAscii = 32 Then Exit Sub
        If Trim(Text1) = "" And KeyAscii = 8 Then Exit Sub
        If KeyAscii = 8 And Len(Text1) = 1 Then Call NoTextInput: Exit Sub
        ls = Text1 & Chr(KeyAscii)
        If KeyAscii = 8 And Len(Text1) > 1 Then ls = Left(Text1, Len(Text1) - 1)
    '-------------------------------------------------------------------------------
        h = FindWindowEx(0, 0&, "NameListWndClass", vbNullString)
        If h <> 0 Then
            If GetShow(h) = True Then
                Call NoTextInput '清空输入文本，停止AutoCode
                Exit Sub
            End If
        End If
    '-------------------------------------------------------------------------------

    '-------------------------------------------------------------------------------
        Call IntAutoList(ls) '装载自动代码列表
  End If
End Sub

Private Sub Timer1_Timer()
On Error GoTo myErr
Dim p As POINTAPI, h As Long
        h = FindWindowEx(0, 0&, "NameListWndClass", vbNullString)
    If h <> 0 Then
        If GetShow(h) = True Then
            Call NoTextInput '清空输入文本，停止AutoCode
        End If
    End If
    GetCaretPos p
    Label2.Caption = "x=" & p.x & " ：" & "y=" & p.y
    If VBInstance.ActiveCodePane Is Nothing Then
        If PrevProcPtr <> 0 Then NoTextInput: UnHookCodeWindow '卸载HOOK
        Text2 = "无"
        TS = "ok"
        Unload AutoCodeFrm
    Else
        If PrevProcPtr = 0 Then
            Load AutoCodeFrm
            HookCodeWindow VBInstance.MainWindow.hwnd, VBInstance.ActiveCodePane.Window.Caption
        End If
        If VBInstance.ActiveCodePane.Window.Caption <> Text2.Text Then
            NoTextInput
            UnHookCodeWindow '卸载HOOK
            Unload AutoCodeFrm
            Load AutoCodeFrm
            HookCodeWindow VBInstance.MainWindow.hwnd, VBInstance.ActiveCodePane.Window.Caption
        End If
        Text2 = VBInstance.ActiveCodePane.Window.Caption
        
    End If
myErr:
End Sub

Private Sub Timer2_Timer()
    Timer2.Enabled = False
    Text1.Text = ""
End Sub

Private Sub Timer3_Timer()
    Timer3.Enabled = False
    FKinput = True
        SendKeys iKey, True
    FKinput = False
End Sub
