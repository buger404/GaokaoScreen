VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'======================================
'   该类模块是由Emerald创建的 页面控制器 模板
'======================================
'   页面绘制器
    Private Type DDate
        Date As Date
        isDisplay As Boolean
        Name As String
    End Type
    Dim Page As GPage
    Dim ScrBlurDC As Long, mode As Boolean, switchTime As Long, setTime As Long
    Dim useTime As Long, pos As POINT
    Dim setMode As Boolean
    Dim setTab As Integer, editDate As Integer
    Dim yPro As Single, mPro As Single, dpro As Single, dName As String, dDisplay As Boolean
    Dim lockBgn As Integer, lockTime As Long, hideTime As Long, hideMinute As Long
    Dim Dates() As DDate
'======================================
Private Sub ReadDates()
    ReDim Dates(0)
    Dim d() As String, t() As String
    d = Split(ESave.GetData("DateList"), vbCrLf)
    For I = 0 To UBound(d) - 1
        t = Split(d(I), "|")
        ReDim Preserve Dates(UBound(Dates) + 1)
        With Dates(UBound(Dates))
            .Name = t(0)
            .Date = t(1)
            .isDisplay = (t(2) = "√")
        End With
    Next
End Sub
Private Sub WriteDates()
    Dim Str As String
    For I = 1 To UBound(Dates)
        Str = Str & Dates(I).Name & "|" & Dates(I).Date & "|" & IIf(Dates(I).isDisplay, "√", "×") & vbCrLf
    Next
    ESave.PutData "DateList", Str
End Sub
Public Sub Update()
    '应用界面每一帧刷新会调用的方法（页面绘制代码请放在此处）
    
    Page.Clear argb(0, 0, 0, 0) '清空画布

    If (GetTickCount - hideTime) / 1000 / 60 <= hideMinute Then
        useTime = GetTickCount
        Exit Sub
    End If
    
    If (GetTickCount - lockTime) / 1000 >= 30 * 60 Then
        lockTime = GetTickCount
        lockBgn = Int(Rnd * 23) + 1
    End If

    If Hour(Now) = 22 And Minute(Now) >= 29 And App.LogMode <> 0 Then
        Page.DrawImage "ohmygod.png", 20 + Sin((GetTickCount Mod 100) / 100 * 6.28) * 2, 10, alpha:=0.7 + Sin((GetTickCount Mod 1000) / 1000 * 3.14) * 0.3
        Exit Sub
    End If
    
    Dim pro1 As Single, pro2 As Single, pros1 As Single, pros2 As Single
    Dim Gaokao As Date, leftH As Single, leftM As Single
    For I = 1 To UBound(Dates)
        If Dates(I).isDisplay Then
            Gaokao = Dates(I).Date
            Exit For
        End If
    Next
    leftH = 24 - Hour(Now) - IIf(Minute(Now) = 0, 0, 1)
    leftM = 60 - Minute(Now) - IIf(Minute(Now) = 0, 60, 1)
    
    Dim cursor As POINT
    GetCursorPos cursor
    If cursor.X <> pos.X Or cursor.y <> pos.y Then useTime = GetTickCount
    pos = cursor
    
    'Progress 1
    If Not mode Then
        pro1 = (GetTickCount - switchTime) / 500
        If pro1 > 1 Then pro1 = 1
        pro1 = Cubic(pro1, 0, 1, 1, 1)
        pros1 = (GetTickCount - switchTime - 600) / 500
        If pros1 < 0 Then pros1 = 0
        If pros1 > 1 Then pros1 = 1
        pros1 = Cubic(pros1, 0, 1, 1, 1)
    Else
        pro1 = 1 - (GetTickCount - switchTime) / 500
        If pro1 < 0 Then pro1 = 0
        pro1 = Cubic(pro1, 0, 0, 0, 1)
        pros1 = pro1
    End If
    
    If pro1 = 0 Then GoTo bypass1
    Dim pro3 As Single
    If setMode Then
        pro3 = (GetTickCount - setTime) / 500
        If pro3 > 1 Then pro3 = 1
        pro3 = Cubic(pro3, 0, 1, 1, 1)
    Else
        pro3 = 1
    End If
    
    If setMode Then PaintDC ScrBlurDC, Page.CDC, 0, 0, 0, 0, GW, GH, alpha:=pro3
    Page.DrawImage "bar2.png", 160 - 500 * (1 - pro1), 10, alpha:=pro1 * 0.6
    Page.Writes format(Hour(Now), "00") & ":" & format(Minute(Now), "00"), 242, 19, 32, Color:=argb(255 * pros1, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
    If setMode Then Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(180 * pro3, 0, 0, 16)
    Page.DrawImage "bar1.png", 20 - 200 * (1 - pro1), 10, alpha:=pro1 * 0.8 - IIf(setMode, pro3 * 0.8, 0)
    Page.DrawImage "bar3.png", 20 - 200 * (1 - pro1), 10, alpha:=IIf(setMode, pro3, 0)
    If CheckMouse2 = mMouseUp Then
        setMode = Not setMode
        If setMode Then BlurTo ScrBlurDC, GetDC(0), AppWindow, 60
        If setMode Then setTime = GetTickCount
        If Not setMode Then ECore.NewTransform transFadeIn
        setTab = 0
        Mouse.State = 0
        Exit Sub
    End If
    If GetTickCount - useTime >= 15 * 60000 Then
        mode = Not mode
        useTime = GetTickCount
        If App.LogMode <> 0 Then SetWindowPos AppWindow.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
        BlurTo ScrBlurDC, GetDC(0), AppWindow, 0
        switchTime = GetTickCount
        Mouse.State = 0
    End If
    If DateDiff("d", Now, Gaokao) < 0 Then
        Page.Writes "已到期", 92, 19, 32, Color:=argb(255 * pros1, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
    Else
        Page.Writes DateDiff("d", Now, Gaokao) & "天", 92, 19, 32, Color:=argb(255 * pros1, 255, 255, 255), align:=StringAlignmentCenter, style:=FontStyleBold
    End If
    
    If setMode Then
        Page.Writes "2022届高三八班倒计时程序" & vbCrLf & "版本号：0.5.17，制作：Buger404", 30, GH - 120, 24, argb(180 * pro3, 255, 255, 255)
        Page.DrawImage "menubtn1.png", 30 - 200 * (1 - pro3), 100, alpha:=pro3
        If CheckMouse2 = mMouseUp Then
            mode = Not mode
            useTime = GetTickCount
            If App.LogMode <> 0 Then SetWindowPos AppWindow.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
            BlurTo ScrBlurDC, GetDC(0), AppWindow, 0
            lockTime = 0
            setMode = False
            switchTime = GetTickCount
            Mouse.State = 0
        End If
        Page.Writes "立即显示锁屏", 50 - 200 * (1 - pro3), 100 + 7, 24, argb(255 * pro3, 32, 32, 32), style:=FontStyleBold
        Page.DrawImage "menubtn1.png", 30 - 200 * (1 - pro3), 170, alpha:=pro3
        If CheckMouse2 = mMouseUp Then
            hideTime = GetTickCount: hideMinute = 30
            setMode = False: ECore.NewTransform
        End If
        Page.Writes "隐藏30分钟", 50 - 200 * (1 - pro3), 170 + 7, 24, argb(255 * pro3, 32, 32, 32), style:=FontStyleBold
        Page.DrawImage "menubtn1.png", 30 - 200 * (1 - pro3), 240, alpha:=pro3
        If CheckMouse2 = mMouseUp Then
            hideTime = GetTickCount: hideMinute = 60
            setMode = False: ECore.NewTransform
        End If
        Page.Writes "隐藏1小时", 50 - 200 * (1 - pro3), 240 + 7, 24, argb(255 * pro3, 32, 32, 32), style:=FontStyleBold
        Page.DrawImage "menubtn2.png", 30 - 200 * (1 - pro3), 310, alpha:=pro3
        If CheckMouse2 = mMouseUp Then
            Unload AppWindow
            End
        End If
        Page.Writes "彻底关闭倒计时", 50 - 200 * (1 - pro3), 310 + 7, 24, argb(255 * pro3, 255, 255, 255), style:=FontStyleBold
        Page.DrawImage "menubtn1.png", 30 - 200 * (1 - pro3), 380, alpha:=pro3
        If CheckMouse2 = mMouseUp Then
            setMode = False: ECore.NewTransform
        End If
        Page.Writes "关闭此菜单", 50 - 200 * (1 - pro3), 380 + 7, 24, argb(255 * pro3, 32, 32, 32), style:=FontStyleBold
        Page.DrawImage "setdialog.png", GW - 580 * pro3, 60, alpha:=pro3
        Select Case setTab
            Case 0
                Page.Writes "倒计时列表设置", GW - (580 - 30) * pro3, 80, 24, argb(255 * pro3, 0, 0, 0), style:=FontStyleBold
                If UBound(Dates) <= 7 Then
                    Page.DrawImage "add.png", GW - 160 * pro3, 0, alpha:=pro3
                    If CheckMouse2 = mMouseUp Then editDate = 0: setTab = 1: ECore.NewTransform
                End If
                For I = 1 To UBound(Dates)
                    Page.Writes Dates(I).Name & " - " & year(Dates(I).Date) & "." & Month(Dates(I).Date) & "." & Day(Dates(I).Date), GW - (580 - 30) * pro3, 110 + 50 * I, 24, argb(255 * pro3, 64, 64, 64), 500, 40
                    If CheckMouse2 = mMouseUp Then
                        editDate = I: setTab = 1: ECore.NewTransform
                        yPro = (year(Dates(I).Date) - year(Now)) / 10
                        mPro = (Month(Dates(I).Date) - 1) / 11
                        dpro = (Day(Dates(I).Date) - 1) / 30
                        dName = Dates(I).Name
                        dDisplay = Dates(I).isDisplay
                    End If
                Next
            Case 1
                Page.Writes IIf(editDate = 0, "添加新的倒计时", "修改倒计时"), GW - (580 - 30) * pro3, 80, 24, argb(255 * pro3, 0, 0, 0), style:=FontStyleBold
                Dim sy As Long, sm As Long, sd As Long
                sy = year(Now) + yPro * 10
                Page.Writes "年份：" & sy, GW - (580 - 30) * pro3, 160, 24, argb(255 * pro3, 64, 64, 64)
                Page.ShowSliderBar yPro, "pro.png", GW - (580 - 230) * pro3, 165
                yPro = Round(yPro * 10) / 10
                sm = 1 + mPro * 11
                Page.Writes "月份：" & sm, GW - (580 - 30) * pro3, 200, 24, argb(255 * pro3, 64, 64, 64)
                Page.ShowSliderBar mPro, "pro.png", GW - (580 - 230) * pro3, 205
                mPro = Round(mPro * 11) / 11
                sd = 1 + dpro * 30
                Page.Writes "日   ：" & sd, GW - (580 - 30) * pro3, 240, 24, argb(255 * pro3, 64, 64, 64)
                Page.ShowSliderBar dpro, "pro.png", GW - (580 - 230) * pro3, 245
                dpro = Round(dpro * 30) / 30
                Dim dWar As Boolean
                If sm = 2 And sd > 28 Then
                    If sy Mod 4 = 0 Or sy Mod 400 = 0 Then
                        If sd > 29 Then
                            Page.Writes "就算是闰年2月也没有这么多天吧？", GW - (580 - 30) * pro3, 320, 24, argb(255 * pro3, 255, 0, 0)
                            dWar = True
                        Else
                            Page.Writes "是罕见的日期呢~", GW - (580 - 30) * pro3, 320, 24, argb(255 * pro3, 255, 0, 0)
                        End If
                    Else
                        Page.Writes "？2月有这个日期吗？", GW - (580 - 30) * pro3, 320, 24, argb(255 * pro3, 255, 0, 0)
                        dWar = True
                    End If
                End If
                If sm = 4 Or sm = 6 Or sm = 9 Or sm = 11 Then
                    If sd > 30 Then
                        Page.Writes "这个日期永远也不会到来。", GW - (580 - 30) * pro3, 320, 24, argb(255 * pro3, 255, 0, 0)
                        dWar = True
                    End If
                End If
                Page.Writes "名称：", GW - (580 - 30) * pro3, 400, 24, argb(255 * pro3, 64, 64, 64)
                Page.ShowEdit dName, 0, GW - (580 - 130) * pro3, 400, 350, 40, argb(255, 64, 64, 64), argb(255, 242, 242, 242), argb(255, 242, 242, 242), argb(255, 232, 232, 232), size:=24
                Page.ShowCheckBox dDisplay, "checkbox.png", GW - (580 - 30) * pro3, 470, "这是高考日期", argb(255, 64, 64, 64), 24
                If dName = "" Then
                    Page.Writes "名称不能为空哦。", GW - (580 - 30) * pro3, 320, 24, argb(255 * pro3, 255, 0, 0)
                    dWar = True
                End If
                If editDate <> 0 Then
                    Page.Writes "删除这个倒计时...", GW - (580 - 30) * pro3, 520, 24, argb(255 * pro3, 255, 0, 0), 500, 40
                    If CheckMouse2 = mMouseUp Then
                        For I = editDate To UBound(Dates) - 1
                            Dates(I) = Dates(I + 1)
                        Next
                        ReDim Preserve Dates(UBound(Dates) - 1)
                        WriteDates
                        setTab = 0: ECore.NewTransform
                    End If
                End If
                If Not dWar Then
                    Page.DrawImage "okbtn.png", GW - (580 - 34) * pro3, 590, alpha:=1
                    If CheckMouse2 = mMouseUp Then
                        If dDisplay Then
                            For I = 1 To UBound(Dates)
                                If Dates(I).isDisplay Then Dates(I).isDisplay = False
                            Next
                        End If
                        If editDate = 0 Then
                            ReDim Preserve Dates(UBound(Dates) + 1)
                            With Dates(UBound(Dates))
                                .Date = CDate(sy & "/" & sm & "/" & sd)
                                .isDisplay = dDisplay
                                .Name = dName
                            End With
                        Else
                            With Dates(editDate)
                                .Date = CDate(sy & "/" & sm & "/" & sd)
                                .isDisplay = dDisplay
                                .Name = dName
                            End With
                        End If
                        WriteDates
                        setTab = 0: ECore.NewTransform
                    End If
                End If
        End Select
    End If

    'Progress 2
bypass1:
    If mode Then
        pro2 = (GetTickCount - switchTime) / 500
        If pro2 > 1 Then pro2 = 1
        pro2 = Cubic(pro2, 0, 1, 1, 1)
        pros2 = (GetTickCount - switchTime - 600) / 500
        If pros2 < 0 Then pros2 = 0
        If pros2 > 1 Then pros2 = 1
        pros2 = Cubic(pros2, 0, 1, 1, 1)
    Else
        pro2 = 1 - (GetTickCount - switchTime) / 500
        If pro2 < 0 Then pro2 = 0
        pro2 = Cubic(pro2, 0, 0, 0, 1)
        pros2 = pro2
    End If
    If pro2 = 0 Then GoTo bypass2
    Page.DrawImage "wallspaper (" & lockBgn & ").jpg", 0, 0, alpha:=pro2
    'PaintDC ScrBlurDC, Page.CDC, 120 * pro2, 120 * pro2, 120 * pro2, 120 * pro2, GW - 240 * pro2, GH - 240 * pro2, alpha:=pro2
    'Page.Paint 0, 120 * pro2, 120 * pro2, GW - 240 * pro2 + 1, GH - 240 * pro2 + 1, argb(pro2 * 200, 1, 1, 1)
    Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(pro2 * 230, 1, 1, 1)

    Dim p(), leftD As Single, DY As Long
    ReDim p(1)
    leftD = DateDiff("d", Now, Gaokao)
    DY = 120 + 0.5 * (GH - 240)
    p(0) = 0: p(1) = DY + Sin(GetTickCount / 6000) * (60 * (leftD / 365))
    For I = 0 To 30
        ReDim Preserve p(UBound(p) + 2)
        p(UBound(p) - 1) = 0 + (I * 1#) / 30 * GW
        p(UBound(p)) = DY + Sin(GetTickCount / 6000 + I * pros2 / 30 * 5) * (60 * (leftD / 365))
    Next
    ReDim Preserve p(UBound(p) + 2)
    p(UBound(p) - 1) = GW
    p(UBound(p)) = GH
    ReDim Preserve p(UBound(p) + 2)
    p(UBound(p) - 1) = 0
    p(UBound(p)) = GH
    ReDim Preserve p(UBound(p) + 2)
    p(UBound(p) - 1) = 0
    p(UBound(p)) = DY + Sin(GetTickCount / 6000) * (60 * (leftD / 365))
    Page.PaintPolygon2 argb(10 * pros2, 255, 255, 255), 0, p, False
    ReDim p(1)
    leftD = DateDiff("d", Now, Gaokao)
    DY = 100 + 0.5 * (GH - 240)
    p(0) = 0: p(1) = DY + Sin(GetTickCount / 3000 - 3) * (60 * (leftD / 365))
    For I = 0 To 30
        ReDim Preserve p(UBound(p) + 2)
        p(UBound(p) - 1) = 0 + (I * 1#) / 30 * (GW)
        p(UBound(p)) = DY + Sin(GetTickCount / 3000 + I * pros2 / 30 * 5 - 3) * (60 * (leftD / 365))
    Next
    ReDim Preserve p(UBound(p) + 2)
    p(UBound(p) - 1) = GW
    p(UBound(p)) = 120 + GH
    ReDim Preserve p(UBound(p) + 2)
    p(UBound(p) - 1) = 0
    p(UBound(p)) = 120 + GH
    ReDim Preserve p(UBound(p) + 2)
    p(UBound(p) - 1) = 0
    p(UBound(p)) = DY + Sin(GetTickCount / 3000 - 3) * (60 * (leftD / 365))
    Page.PaintPolygon2 argb(10 * pros2, 255, 255, 255), 0, p, False
    
    Page.Writes "仅剩                天", GW / 2 - 20, GH / 2 - 50, 48, argb(150 * pros2, 255, 255, 255), style:=FontStyleBold, align:=StringAlignmentCenter
    Page.Writes DateDiff("d", Now, Gaokao), GW / 2, GH / 2 - 100, 102, argb(255 * pros2, 255, 255, 255), style:=FontStyleBold, align:=StringAlignmentCenter
    Page.Writes Month(Now) & "月" & Day(Now) & "日  " & format(Hour(Now), "00") & ":" & format(Minute(Now), "00"), GW / 2, GH / 2 + 100 - 50, 36, argb(150 * pros2, 255, 255, 255), style:=FontStyleBold, align:=StringAlignmentCenter
    Dim diffDay As Long
    DY = GH / 2 + 180 - 50
    For I = 1 To UBound(Dates)
        If Not Dates(I).isDisplay Then
            diffDay = DateDiff("d", Now, Dates(I).Date)
            If diffDay >= 0 Then
                Page.Writes "距离" & Dates(I).Name & "仅剩" & diffDay & "天", GW / 2, DY, 36, argb(220 * pros2, 255, 255, 255), style:=FontStyleBold, align:=StringAlignmentCenter
                DY = DY + 50
            End If
        End If
    Next
    
    Page.Writes "高三八班加油!!!  (≧^≦)/", GW / 2, 20, 24, argb(150 * pros2, 255, 255, 255), style:=FontStyleBold, align:=StringAlignmentCenter

    'Page.Writes "高三时间仅剩 " & format(Int((leftD + (leftH * 60 + leftM) / (24 * 60)) / 365 * 10000) / 100, "00.00") & "%", GW / 2, GH - 94, 32, argb(100 * pros2, 255, 255, 255), align:=StringAlignmentCenter

bypass2:
    If Mouse.State >= 1 And mode And Mouse.button = 1 Then
        mode = Not mode
        useTime = GetTickCount
        switchTime = GetTickCount
        Mouse.State = 0
    End If
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '鼠标滚轮事件
    'Direction:方向，Depth：深度
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '动画消息接收
End Sub
Public Sub Enter()
    '页面进入事件
    ReadDates
    dName = "xx考试"
    lockBgn = Int(Rnd * 23) + 1
End Sub
Public Sub Leave()
    '页面离开事件
End Sub
Public Sub Gotfocus()
    '窗口取得焦点事件
End Sub
Public Sub Lostfocus()
    '窗口失去焦点事件
End Sub
Public Sub AnimationDone(id As Integer)
    '动画播放结束事件
    'id：结束的动画ID
End Sub
Private Sub Class_Initialize()

    Set Page = New GPage  '创建页面绘制器
    Page.Create Me  '将绘制器与控制器绑定
    Page.Res.NewImages App.path & "\assets", 0.3 '导入资源
    Page.Res.NewImages App.path & "\assets\locks", GW, GH
    Page.Res.newImage App.path & "\assets\lock.jpg", GW, GH, "lockbg"
    Page.Res.ApplyBlurEffect "lockbg", 60, 0
    ScrBlurDC = CreateCDC(GW, GH)
    ECore.Add Page, "AppPage"  '创建页面
    
    '===============================================
    '   如果需要添加动画，请放置在此处
    
    '===============================================
End Sub
