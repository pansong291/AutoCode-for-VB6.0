Attribute VB_Name = "Pub_Mod"
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

Public VBInstance As VBIDE.VBE 'IDE接口
Public Connect As Connect '连接接口
Public SetFonz As VBIDE.Properties
Public SetFonzs As VBIDE.Property

Public iButton     As Office.CommandBarButton

'---------------------------------------------------------------------------------
Public hWndCodeWindow As Long '代码窗口句柄
Public PrevProcPtr As Long
Public OldParent As Long '自动代码列表 旧的父句柄
Public AtCodeParent As New Parent_Cls
'---------------------------------------------------------------------------------

Public PubHs() As String 'VB所有的函数
Public PubYj() As String 'VB所有关键字和语句
Public Tkey As String '当输入这些字符时，取消自动代码列表
Public Okey As String '当输入这些字符时，记录输入并且判断特殊键入：删除 空格 下划线
Public SPkey As String '可以用空格输入的关键字
'---------------------------------------------------------------------------------
Public ListType As Integer '自动代码顺序类型：1=先函数后语句  2=先语句后函数

Public SetFont As String '代码窗口的字体
Public SetFontSize As Integer '代码窗口的字号
Public FontHeight As Integer '代码窗口字体高度

Public FKinput As Boolean '判断是否正在自动输入

Public TS As String

