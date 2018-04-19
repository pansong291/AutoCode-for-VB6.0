VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9405
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   14100
   _ExtentX        =   24871
   _ExtentY        =   16589
   _Version        =   393216
   Description     =   "Add-In Project Template"
   DisplayName     =   "My Add-In"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'VBProjects集合: 通过该集合可以访问VB实例中所包含的工程
'Windows集合：通过该集合可以访问所有的窗口，包括控件栏、属性栏以及工程中的窗体等。
'CodePanes集合: 通过该集合可以访问所有的代码窗口，可以获得代码窗口中的代码以及改变其中的代码
'CommandBars集合：通过该集合可以访问VB实例中的所有命令栏，包括支持快速菜单的命令栏。
'Events 集合: 通过该集合插件可以访问VB中的所有事件对象

Option Explicit

Public FormDisplayed          As Boolean
Dim iVBInstance             As VBIDE.VBE
Dim mcbMenuCommandBar         As Office.CommandBarControl
Dim mfrmAddIn                 As New frmAddIn
Public WithEvents MenuHandler As CommandBarEvents          '命令栏事件句柄
Attribute MenuHandler.VB_VarHelpID = -1
Public WithEvents mButton    As CommandBarEvents
Attribute mButton.VB_VarHelpID = -1

Sub Hide()
    
    On Error Resume Next
    
    FormDisplayed = False
    mfrmAddIn.Hide
   
End Sub

Sub Show()
  
    On Error Resume Next
    
    If mfrmAddIn Is Nothing Then
        Set mfrmAddIn = New frmAddIn
    End If
    
    Set VBInstance = iVBInstance
    Set Connect = Me
    FormDisplayed = True
    
    mfrmAddIn.Show 1, Me
End Sub

'------------------------------------------------------
'这个方法添加外接程序到 VB
'------------------------------------------------------
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error GoTo error_handler
    
    '保存 vb 实例
    Set iVBInstance = Application
    
    '这里是一个设置断点以及测试不同外接程序
    '对象,属性及方法的适当位置
    'Debug.Print VBInstance.FullName

    If ConnectMode = ext_cm_External Then
        '用于让向导工具栏来启动此向导
        Me.Show
    Else
        'Set mcbMenuCommandBar = AddToAddInCommandBar("启动AutoCode")
        '吸取事件
        'Set Me.MenuHandler = iVBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
        
            'iVBInstance.CommandBars(1).Controls.Add Type:=10, Before:=2
            'iVBInstance.CommandBars(1).Controls(2).Caption = "我的菜单"
            'iVBInstance.CommandBars(1).Controls(2).OnAction = "OKcmd"
            'iVBInstance.CommandBars(1).Controls(2).Visible = True
            Set iButton = iVBInstance.CommandBars(1).Controls.Add(Type:=msoControlButton, temporary:=True)
                With iButton
                    .Caption = " AutoCode "
                    .ToolTipText = "AutoCode"
                    .Style = msoButtonCaption 'msoButtonIconAndCaption 'msoButtonIcon 'msoButtonCaption
                    '.BeginGroup = True
                    .State = msoButtonUp
                End With
                Set mButton = iVBInstance.Events.CommandBarEvents(iButton)
    End If
  
    If ConnectMode = ext_cm_AfterStartup Then
        If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
            '设置这个到连接显示的窗体
            Me.Show
        End If
    End If
  
    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

'------------------------------------------------------
'这个方法从 VB 中删除外接程序
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    
    '删除命令栏条目
    iButton.Delete
    
    '关闭外接程序
    If FormDisplayed Then
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "1"
        FormDisplayed = False
    Else
        SaveSetting App.Title, "Settings", "DisplayOnConnect", "0"
    End If
    
    Unload mfrmAddIn
    Set mfrmAddIn = Nothing

End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    If GetSetting(App.Title, "Settings", "DisplayOnConnect", "0") = "1" Then
        '设置这个到连接显示的窗体
        Me.Show
    End If
End Sub

'当 IDE 中的菜单被单击时,这个事件被激活
'Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    'Me.Show
'End Sub

'Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl
'    Dim cbMenuCommandBar As Office.CommandBarControl  '命令栏对象
'    Dim cbMenu As Object
'
'    On Error GoTo AddToAddInCommandBarErr
'
'    '察看能否找到外接程序菜单
'    Set cbMenu = iVBInstance.CommandBars("Add-Ins")
'    If cbMenu Is Nothing Then
'        '没有有效的外接程序,过程失败
'        Exit Function
'    End If
'
'    '添加它到命令栏
'    Set cbMenuCommandBar = cbMenu.Controls.Add(1)
'    '设置标题
'
'    cbMenuCommandBar.Caption = sCaption
'
'    Set AddToAddInCommandBar = cbMenuCommandBar
'
'    Exit Function
'
'AddToAddInCommandBarErr:
'
'End Function

Private Sub mButton_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Me.Show
End Sub
