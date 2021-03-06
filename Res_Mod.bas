Attribute VB_Name = "Sub_Res_Mod"
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
Private Type POINTAPI
x As Long '点在X坐标(横坐标)上的坐标值
y As Long '点在Y坐标(纵坐标)上的坐标值
End Type
Private Declare Function GetCaretPos Lib "user32.dll" (lpPoint As POINTAPI) As Long


Function Initialization()  '初始化

Dim sz As String
    'SetFont = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\VBA\Microsoft Visual Basic", "FontFace")
    SetFont = "Droid Sans Mono" '上面是自动读取vb6字体设置信息的
    'sz = GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\VBA\Microsoft Visual Basic", "FontHeight")
    sz = "9" '上面是自动读取vb6字号设置信息的
    SetFontSize = sz
    SetFont = Replace(SetFont, Chr(0), "")
    If SetFontSize = 0 Then MsgBox "获取vb6选项字号出错，系统将以默认设置启动。", 64, "提示": SetFontSize = 10
    If SetFont = "" Then MsgBox "获取vb6选项字体出错，系统将以默认设置启动。", 64, "提示": SetFont = "宋体"
    ListType = 1
    Call GetFontHeight
    Call IntHsKey
    Load AutoCodeFrm
    
End Function



'=========================================================================================================
Function GetFontHeight() '获取代码窗口字体高度
    JS_Frm.FontName = SetFont
    JS_Frm.FontSize = SetFontSize
    FontHeight = JS_Frm.TextHeight("高") / 15
End Function

'=========================================================================================================
Function IntHsKey() '装载函数 关键字 语句
    Dim ls1 As String, ls2 As String
    
    ls1 = "Abs,Array,Asc,Atn,CallByName,Choose,Chr,CStr,CInt,CDate,Command,Cos,CreateObject,CurDir,CVErr,Date,DateAdd,DateDiff,DatePart,DateSerial,DateValue,Day," & _
           "DDB,Dir,DoEvents,Environ,EOF,Error,Exp,FileAttr,FileDateTime,FileLen,Filter,Format,FormatCurrency,FormatDateTime,FormatNumber," & _
           "FormatPercent,FreeFile,FV,GetAllSettings,GetAttr,GetObject,GetSetting,Hex,Hour,IIf,IMEStatus,Input,InputBox,InStr,InStrRev,Int、Fix," & _
           "IPmt,IRR,IsArray,IsDate,IsEmpty,IsError,IsMissing,IsNull,IsNumeric,IsObject,Join,LBound,LCase,Left,Len,Loc,LOF,Log,LTrim,LoadPicture,LoadResData,LoadResPicture,LoadResString,RTrim,Trim," & _
           "Mid,Minute,MIRR,Month,MonthName,MsgBox,Now,NPer,NPV,Oct,Partition,Pmt,PPmt,PV,QBColor,Rate,Replace,RGB,Right,Rnd,Round,Second,Seek," & _
           "Sgn,Shell,Sin,SLN,Space,Spc,Split,Sqr,Str,StrComp,StrConv,StrReverse,String,Switch,SYD,Tab,Tan,Time,Timer,TimeSerial,TimeValue," & _
           "TypeName,UBound,UCase,Val,VarType,Weekday,WeekdayName,Year,"
   
   ls2 = "App,AppActivate,Beep,Call,ChDir,ChDrive,Close,Const,Date,Declare,Deftype,DeleteSetting," & _
          "Dim,Do,Loop,End,Enum,Erase,Error,Event,Exit,FileCopy,For,Each,Next,Function,Get," & _
          "GoSub,Return,GoTo,If,IIf,Then,Else,Implements,Input,Kill,Let,Line,Lock,Unlock," & _
          "LSet,Mid,MkDir,Name,On Error GoTo,On Error Resume Next,Open,Option Base," & _
          "Option Compare,Option Explicit,Option Private,Print,Private,Property Get," & _
          "Property Let,Property Set,Public,Put,RaiseEvent,Randomize,ReDim,Rem,Reset," & _
          "Resume,RmDir,RSet,SaveSetting,Seek,Select,Case,SendKeys,Set,SetAttr,Static,Screen," & _
          "Stop,Sub,Time,Type,While,Wend,With,Write,Nothing,Null,Me,"
          
    'SPkey = "|function|DateAdd|DateDiff|DatePart|DateSerial|DateValue|||||||||||||||||||||||||||||||"
    
    PubHs = Split(ls1, ",")
    PubYj = Split(ls2, ",")

    Tkey = "|~|!|@|#|$|%|^|&|*|(|)|=|+|*|-|/|`|[|]|{|}|\|'|;|:|""|<|>|?|,|.||||"
    
    
End Function
'=========================================================================================================
Function NoTextInput() '清空输入文本，停止AutoCode
On Error Resume Next

    JS_Frm.Timer2.Enabled = True '清空输入文本
    AtCodeParent.OutMdi AutoCodeFrm.hwnd
    If AutoCodeFrm.Visible = False Then Exit Function
    AutoCodeFrm.Visible = False
    AutoCodeFrm.ATlv.ListItems.Clear
    VBInstance.ActiveCodePane.Window.SetFocus
End Function
'=========================================================================================================
Function IntAutoList(key As String)  '装载自动代码列表
Dim ItemX As ListItem
Dim i As Integer
Dim p As POINTAPI
'If TS <> "" Then MsgBox "0"
With AutoCodeFrm
    GetCaretPos p
    If p.x < 0 Then p.x = 3
    If p.y < 0 Then p.y = 3
    'If TS <> "" Then MsgBox "1"
    AtCodeParent.SetMdi hWndCodeWindow, .hwnd, p.x & "," & p.y + FontHeight & ",192,168,"
    'If TS <> "" Then MsgBox "2"
    .ATlv.ListItems.Clear

    Call IntAT_Frm(LCase(key))
    Call IntAT_Frm_Ct(LCase(key))
    Call IntAT_Hs(LCase(key))
    Call IntAT_Yj(LCase(key))
    
    'If TS <> "" Then MsgBox "3"
    
    If .ATlv.ListItems.Count = 0 Then '判断是否有匹配的关键字
        .Visible = False
        AtCodeParent.OutMdi .hwnd
        VBInstance.ActiveCodePane.Window.SetFocus
        Exit Function
    End If
    'If TS <> "" Then MsgBox "4"
    .ATlv.ListItems(1).Selected = True
    .ATlv.ListItems(1).ForeColor = &HFFFFFF
    .ATlv.ListItems(1).EnsureVisible
    .Visible = True
    'If TS <> "" Then MsgBox "1"
    If .ATlv.ListItems.Count > 9 Then
        .ATlv.ColumnHeaders(1).Width = 2480
    Else
        .ATlv.ColumnHeaders(1).Width = 2750
    End If
    Call AutoCodeFrm.SetLIC 'LV背景行绘制颜色
    
    'If TS <> "" Then MsgBox "7"
    VBInstance.ActiveCodePane.Window.SetFocus
End With
End Function

'=========================================================================================================
    Function IntAT_Frm(key As String)  '装载工程内的窗体名字列表
        On Error GoTo myErr
        Dim ItemX As ListItem
        Dim mCop As VBIDE.VBComponent
        With AutoCodeFrm.ATlv
            For Each mCop In VBInstance.VBProjects.StartProject.VBComponents
                If mCop.Type = vbext_ct_ActiveXDesigner Or mCop.Type = vbext_ct_UserControl Or mCop.Type = vbext_ct_VBForm Or mCop.Type = vbext_ct_VBMDIForm Then
                    
                    If Left(LCase(mCop.Name), Len(key)) = key Then
                    If mCop.Type = vbext_ct_UserControl Then
                        Set ItemX = .ListItems.Add(, "k" & .ListItems.Count + 1, mCop.Name, , 3)
                    Else
                        Set ItemX = .ListItems.Add(, "k" & .ListItems.Count + 1, mCop.Name, , 2)
                    End If
                    End If
                End If
            Next
        End With
myErr:
        Err.Clear
    End Function
'=========================================================================================================
    Function IntAT_Frm_Ct(key As String) '装载当前窗体的控件名字列表
        On Error GoTo myErr
        Dim ItemX As ListItem
        Dim xForm As Object
        Dim xControl As Object

        If VBInstance.ActiveCodePane.CodeModule.Parent.Type = vbext_ct_ActiveXDesigner Or VBInstance.ActiveCodePane.CodeModule.Parent.Type = vbext_ct_UserControl Or VBInstance.ActiveCodePane.CodeModule.Parent.Type = vbext_ct_VBForm Or VBInstance.ActiveCodePane.CodeModule.Parent.Type = vbext_ct_VBMDIForm Then
            With AutoCodeFrm.ATlv
            Set xForm = VBInstance.VBProjects.StartProject.VBComponents(VBInstance.ActiveCodePane.CodeModule.Parent.Name).Designer
            For Each xControl In xForm.ContainedVBControls
                If Not (xControl.ControlObject Is Nothing) Then
                    If Left(LCase(xControl.ControlObject.Name), Len(key)) = key Then
                        Set ItemX = .ListItems.Add(, "k" & .ListItems.Count + 1, xControl.ControlObject.Name, , 3)
                    End If
                End If
            Next
            End With
        End If
myErr:
        Err.Clear
    End Function
'=========================================================================================================
    Function IntAT_Hs(key As String) '装载函数名字列表
        On Error GoTo myErr
        Dim ItemX As ListItem, i As Integer

        With AutoCodeFrm.ATlv
            For i = 0 To UBound(PubHs) - 1
                If PubHs(i) <> "" Then
                    If Left(LCase(PubHs(i)), Len(key)) = key Then
                        Set ItemX = .ListItems.Add(, "k" & .ListItems.Count + 1, PubHs(i), , 1)
                    End If
                End If
            Next
        End With
myErr:
        Err.Clear
    End Function
'=========================================================================================================
    Function IntAT_Yj(key As String) '装载语句名字列表
        On Error GoTo myErr
        Dim ItemX As ListItem, i As Integer

        With AutoCodeFrm.ATlv
            For i = 0 To UBound(PubYj) - 1
                If PubYj(i) <> "" Then
                    If Left(LCase(PubYj(i)), Len(key)) = key Then
                        Set ItemX = .ListItems.Add(, "k" & .ListItems.Count + 1, PubYj(i), , 4)
                    End If
                End If
            Next
        End With
myErr:
        Err.Clear
    End Function
'=========================================================================================================
Function GetShow(hwnd As Long) As Boolean '判断指定控件是否可见
    Dim ihwnd As Long
        ihwnd = GetWindowLong(hwnd, -16)
        If ihwnd& And &H10000000 Then
            GetShow = True
        Else
            GetShow = False
        End If
End Function





