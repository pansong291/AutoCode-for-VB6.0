Attribute VB_Name = "Sub_Res_Mod"
'////////////////////////////////////////////////////////////////
'
'²å¼þÃû³Æ£º
'
'²å¼þ×÷Õß£ºÈËÏÐ»¨Âä QQ£º449806776
'
'°æÈ¨ÉùÃ÷£ºÄú¿ÉÒÔÐÞ¸Ä»ò¹²Ïí·¢²¼´Ë²å¼þ£¬µ«±ØÐë×¢Ã÷Ô­´´×÷ÕßÐÅÏ¢
'
'VB°®ºÃÕß£º½»Á÷QQÈº¡ª¡ª19871152
'
'////////////////////////////////////////////////////////////////
Option Explicit
Private Type POINTAPI
x As Long 'µãÔÚX×ø±ê(ºá×ø±ê)ÉÏµÄ×ø±êÖµ
y As Long 'µãÔÚY×ø±ê(×Ý×ø±ê)ÉÏµÄ×ø±êÖµ
End Type
Private Declare Function GetCaretPos Lib "user32.dll" (lpPoint As POINTAPI) As Long


Function Initialization()  '³õÊ¼»¯

Dim sz As String
    'SetFont = GetStringValue("HKEY_CURRENT_USER\Software\Microsoft\VBA\Microsoft Visual Basic", "FontFace")
    SetFont = "Droid Sans Mono" 'ÉÏÃæÊÇ×Ô¶¯¶ÁÈ¡vb6×ÖÌåÉèÖÃÐÅÏ¢µÄ
    'sz = GetDWORDValue("HKEY_CURRENT_USER\Software\Microsoft\VBA\Microsoft Visual Basic", "FontHeight")
    sz = "9" 'ÉÏÃæÊÇ×Ô¶¯¶ÁÈ¡vb6×ÖºÅÉèÖÃÐÅÏ¢µÄ
    SetFontSize = sz
    SetFont = Replace(SetFont, Chr(0), "")
    If SetFontSize = 0 Then MsgBox "»ñÈ¡vb6Ñ¡Ïî×ÖºÅ³ö´í£¬ÏµÍ³½«ÒÔÄ¬ÈÏÉèÖÃÆô¶¯¡£", 64, "ÌáÊ¾": SetFontSize = 10
    If SetFont = "" Then MsgBox "»ñÈ¡vb6Ñ¡Ïî×ÖÌå³ö´í£¬ÏµÍ³½«ÒÔÄ¬ÈÏÉèÖÃÆô¶¯¡£", 64, "ÌáÊ¾": SetFont = "ËÎÌå"
    ListType = 1
    Call GetFontHeight
    Call IntHsKey
    Load AutoCodeFrm
    
End Function



'=========================================================================================================
Function GetFontHeight() '»ñÈ¡´úÂë´°¿Ú×ÖÌå¸ß¶È
    JS_Frm.FontName = SetFont
    JS_Frm.FontSize = SetFontSize
    FontHeight = JS_Frm.TextHeight("¸ß") / 15
End Function

'=========================================================================================================
Function IntHsKey() '×°ÔØº¯Êý ¹Ø¼ü×Ö Óï¾ä
    Dim ls1 As String, ls2 As String
    
    ls1 = "Abs,Array,Asc,Atn,CallByName,Choose,Chr,CStr,CInt,CDate,Command,Cos,CreateObject,CurDir,CVErr,Date,DateAdd,DateDiff,DatePart,DateSerial,DateValue,Day," & _
           "DDB,Dir,DoEvents,Environ,EOF,Error,Exp,FileAttr,FileDateTime,FileLen,Filter,Format,FormatCurrency,FormatDateTime,FormatNumber," & _
           "FormatPercent,FreeFile,FV,GetAllSettings,GetAttr,GetObject,GetSetting,Hex,Hour,IIf,IMEStatus,Input,InputBox,InStr,InStrRev,Int¡¢Fix," & _
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
Function NoTextInput() 'Çå¿ÕÊäÈëÎÄ±¾£¬Í£Ö¹AutoCode
On Error Resume Next

    JS_Frm.Timer2.Enabled = True 'Çå¿ÕÊäÈëÎÄ±¾
    AtCodeParent.OutMdi AutoCodeFrm.hwnd
    If AutoCodeFrm.Visible = False Then Exit Function
    AutoCodeFrm.Visible = False
    AutoCodeFrm.ATlv.ListItems.Clear
    VBInstance.ActiveCodePane.Window.SetFocus
End Function
'=========================================================================================================
Function IntAutoList(key As String)  '×°ÔØ×Ô¶¯´úÂëÁÐ±í
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
    
    If .ATlv.ListItems.Count = 0 Then 'ÅÐ¶ÏÊÇ·ñÓÐÆ¥ÅäµÄ¹Ø¼ü×Ö
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
    Call AutoCodeFrm.SetLIC 'LV±³¾°ÐÐ»æÖÆÑÕÉ«
    
    'If TS <> "" Then MsgBox "7"
    VBInstance.ActiveCodePane.Window.SetFocus
End With
End Function

'=========================================================================================================
    Function IntAT_Frm(key As String)  '×°ÔØ¹¤³ÌÄÚµÄ´°ÌåÃû×ÖÁÐ±í
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
    Function IntAT_Frm_Ct(key As String) '×°ÔØµ±Ç°´°ÌåµÄ¿Ø¼þÃû×ÖÁÐ±í
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
    Function IntAT_Hs(key As String) '×°ÔØº¯ÊýÃû×ÖÁÐ±í
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
    Function IntAT_Yj(key As String) '×°ÔØÓï¾äÃû×ÖÁÐ±í
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
Function GetShow(hwnd As Long) As Boolean 'ÅÐ¶ÏÖ¸¶¨¿Ø¼þÊÇ·ñ¿É¼û
    Dim ihwnd As Long
        ihwnd = GetWindowLong(hwnd, -16)
        If ihwnd& And &H10000000 Then
            GetShow = True
        Else
            GetShow = False
        End If
End Function





