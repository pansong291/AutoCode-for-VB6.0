VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "疯鱼AutoCode for VB6.0 学习版"
   ClientHeight    =   3450
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "停止AutoCode服务"
      Enabled         =   0   'False
      Height          =   540
      Left            =   3045
      TabIndex        =   2
      Top             =   2760
      Width           =   2145
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   75
      TabIndex        =   1
      Top             =   -15
      Width           =   5895
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   45
         TabIndex        =   3
         Top             =   1995
         Width           =   5790
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "输入，按回车键选择；ESC键取消自动列表。"
            ForeColor       =   &H0000FF00&
            Height          =   180
            Left            =   645
            TabIndex        =   5
            Top             =   345
            Width           =   3510
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "说明：本插件可辅助在空白行输入VB关键字或者控件名字时，自动匹配"
            ForeColor       =   &H0000FF00&
            Height          =   180
            Left            =   105
            TabIndex        =   4
            Top             =   90
            Width           =   5580
         End
      End
      Begin VB.Image Image1 
         Height          =   1860
         Left            =   45
         Picture         =   "frmAddIn.frx":0000
         Top             =   120
         Width           =   5790
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "启动AutoCode服务"
      Height          =   540
      Left            =   735
      TabIndex        =   0
      Top             =   2760
      Width           =   2145
   End
End
Attribute VB_Name = "frmAddIn"
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

'mid(VBInstance.VBProjects.StartProject.VBComponents(List1.List(List1.ListIndex)).CodeModule.Lines(a,1),b-3,3)
'Debug.Print Me.TextWidth("中"), Me.TextHeight("中")
'VBInstance.ActiveCodePane.CodeModule.Parent.Type 5=窗体的代码  6=MDI窗体代码 1=模块代码  2=类模块  8=用户控件模块
'CodeModule.Members 代码过程集合

Option Explicit






Private Sub Command1_Click()
    Me.Hide
    JS_Frm.Timer1.Enabled = False
    If PrevProcPtr <> 0 Then NoTextInput: UnHookCodeWindow
    Unload AutoCodeFrm
    Command3.Enabled = True
    Command1.Enabled = False
End Sub

Private Sub Command3_Click() '开启AutoCode服务
    Me.Hide
    Load JS_Frm
    Call Initialization '初始化
    JS_Frm.Timer1.Enabled = True
    Command3.Enabled = False
    Command1.Enabled = True
End Sub
    'HookCodeWindow VBInstance.MainWindow.hwnd, VBInstance.ActiveCodePane.Window.Caption
    'Call Initialization



'=================================================================================================================
'Sub GetFrm()
'Dim mCop As Object
''获得当前启动工程中的所有对象
'    For Each mCop In VBInstance.VBProjects.StartProject.VBComponents
'        If mCop.Type = vbext_ct_VBForm Then
'            List1.AddItem mCop.Name '如果对象是窗体类型就将其添加到ListBox中
'        End If
'    Next
'
'    If List1.ListCount < 1 Then
'        MsgBox "工程中没有添加控件的窗体"
'        Connect.Hide
'    Else
'        List1.ListIndex = 0
'    End If
'    Command1.Caption = "Add Code"
'End Sub

'Private Sub Command1_Click()
'Dim xComp As VBComponent
'Dim xModule As VBComponent
'Dim xForm As VBForm
'Dim xControl As VBControl
'Dim xCode As CodeModule
'
'
'
'    '获得用户选择的窗体对象
'    Set xComp = VBInstance.VBProjects.StartProject.VBComponents(List1.List(List1.ListIndex))
'    '获得窗体设计器对象
'    Set xForm = xComp.Designer
'
'    '添加一个CommandButton到窗体上
'    Set xControl = xForm.VBControls.Add("VB.CommandButton")
'    '设定控件的名称
'    xControl.Properties("Name") = "cmdButton"
'    '添加控件的Click事件代码
'    xComp.CodeModule.CreateEventProc "Click", "cmdButton"
'
'    '添加一个新模块到工程中
'    Set xModule = VBInstance.VBProjects.StartProject.VBComponents.Add(vbext_ct_StdModule)
'    '设定模块名称
'    xModule.Properties("Name") = "ModulTemp"
'    '获得对象的代码对象
'    Set xCode = xModule.CodeModule
'
'Dim astr As String
'
'    '添加mClick子程序到新模块中
'    astr = "Public Sub mClick()" + Chr(13) + Chr(10) + Chr(vbKeyTab) + "MsgBox ""You click a button!""" + Chr(13) + Chr(10) + "End Sub"
'    xCode.AddFromString astr
'
'Dim lCount As Long
'
'    '在cmdButton的Click事件中添加执行mClick子程序
'        lCount = xComp.CodeModule.ProcBodyLine("cmdButton_Click", vbext_pk_Proc)
'    If lCount <> 0 Then
'        xComp.CodeModule.InsertLines lCount + 1, "mClick"
'    End If
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Me.Hide
End Sub
