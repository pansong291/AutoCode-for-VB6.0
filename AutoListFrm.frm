VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form AutoCodeFrm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2790
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   1470
      ScaleHeight     =   345
      ScaleWidth      =   1275
      TabIndex        =   1
      Top             =   1830
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2085
      Top             =   1710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoListFrm.frx":0000
            Key             =   ""
            Object.Tag             =   "函数"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoListFrm.frx":0112
            Key             =   ""
            Object.Tag             =   "窗体"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoListFrm.frx":0224
            Key             =   ""
            Object.Tag             =   "控件"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AutoListFrm.frx":0336
            Key             =   ""
            Object.Tag             =   "语句"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ATlv 
      Height          =   2490
      Left            =   -30
      TabIndex        =   0
      Top             =   -30
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   4392
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "name"
         Object.Width           =   4375
      EndProperty
   End
End
Attribute VB_Name = "AutoCodeFrm"
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

Option Explicit

Private Sub ATlv_DblClick()
Dim ls As String
        ls = "{BACKSPACE " & Len(JS_Frm.Text1) & "}" & ATlv.SelectedItem.Text
        Call NoTextInput '清空输入文本，停止AutoCode
        FKinput = True
        SendKeys ls, True
        FKinput = False
End Sub

Private Sub ATlv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    SetListItemColor ATlv, Picture1
    VBInstance.ActiveCodePane.Window.SetFocus
End Sub

Private Sub Form_Load()
    Me.Visible = False
End Sub

Sub SetLIC()
    SetListItemColor ATlv, Picture1
End Sub

Private Sub SetListItemColor(lv As ListView, picBg As PictureBox)

Dim i As Integer
Dim mItem As ListItems
picBg.Cls
lv.Parent.ScaleMode = vbTwips

picBg.Width = lv.ColumnHeaders(1).Width

picBg.Height = lv.ListItems(1).Height * (lv.ListItems.Count)

picBg.ScaleHeight = lv.ListItems.Count

picBg.ScaleWidth = 1

picBg.DrawWidth = 1

'If TS <> "" Then MsgBox "4"
'-----------------------------

'custom.such as

'------------------------------
Set mItem = lv.ListItems

For i = 1 To lv.ListItems.Count

If mItem(i).Selected = False Then
    picBg.Line (0, i - 1)-(1, i), RGB(255, 255, 255), BF
    mItem(i).ForeColor = &H0&
Else
    picBg.Line (0, i - 1)-(1, i), RGB(15, 66, 145), BF
     mItem(i).ForeColor = &HFFFFFF
End If

Next

lv.Picture = picBg.Image

End Sub

