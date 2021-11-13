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
'   ����ģ������Emerald������ ҳ������� ģ��
'======================================
'   ҳ�������
    Dim Page As GPage
    Dim ScrBlurDC As Long, mode As Boolean, switchTime As Long
    Dim useTime As Long, pos As POINT
'======================================

Public Sub Update()
    'Ӧ�ý���ÿһ֡ˢ�»���õķ�����ҳ����ƴ�������ڴ˴���
    
    Page.Clear argb(0, 0, 0, 0) '��ջ���

    Dim pro1 As Single, pro2 As Single, pros1 As Single, pros2 As Single
    Dim Gaokao As Date, leftH As Single, leftM As Single
    Gaokao = #6/7/2022#
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
    Page.DrawImage "bar.png", GW / 2, 45 - 300 * (1 - pro1), alpha:=pro1, pos:=posOnCenter
    If CheckMouse2 = mMouseUp Or GetTickCount - useTime >= 15 * 60000 Then
        mode = Not mode
        useTime = GetTickCount
        SetWindowPos AppWindow.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
        BlurTo ScrBlurDC, GetDC(0), AppWindow, 0
        switchTime = GetTickCount
        Mouse.State = 0
    End If
    Page.Writes DateDiff("d", Now, Gaokao) & "��  " & format(leftH, "00") & ":" & format(leftM, "00"), GW / 2 + 210, 8 - 300 * (1 - pro1), 32, Color:=argb(255 * pros1, 255, 255, 255), align:=StringAlignmentFar, style:=FontStyleBold
    
    Page.PaintArc GW / 2 - 190, 29, 36, 36, 360 * pros1, -90, Color:=argb(255, 96, 96, 96), size:=5, pos:=posOnCenter, style:=1
    Page.PaintArc GW / 2 - 190, 29, 36, 36, ((leftH * 60 + leftM) / (24 * 60)) * 360 * pros1, -90, Color:=argb(255, 255, 255, 255), size:=5, pos:=posOnCenter, style:=1
    
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
    Page.DrawImage "lockbg", 0, 0, alpha:=pro2
    PaintDC ScrBlurDC, Page.CDC, 120 * pro2, 120 * pro2, 120 * pro2, 120 * pro2, GW - 240 * pro2, GH - 240 * pro2, alpha:=pro2
    Page.Paint 0, 120 * pro2, 120 * pro2, GW - 240 * pro2 + 1, GH - 240 * pro2 + 1, argb(pro2 * 200, 1, 1, 1)
    Page.Paint 0, 0, 0, GW + 1, GH + 1, argb(pro2 * 150, 1, 1, 1)

    Dim p(), leftD As Single, dY As Long
    ReDim p(1)
    leftD = DateDiff("d", Now, Gaokao)
    dY = 120 + leftD / 365 * (GH - 240)
    p(0) = 120: p(1) = dY + Sin(GetTickCount / 6000) * (60 * (leftD / 365))
    For I = 0 To 30
        ReDim Preserve p(UBound(p) + 2)
        p(UBound(p) - 1) = 120 + (I * 1#) / 30 * (GW - 240)
        p(UBound(p)) = dY + Sin(GetTickCount / 6000 + I * pros2 / 30 * 5) * (60 * (leftD / 365))
    Next
    ReDim Preserve p(UBound(p) + 2)
    p(UBound(p) - 1) = 120 + (GW - 240)
    p(UBound(p)) = 120 + GH - 240
    ReDim Preserve p(UBound(p) + 2)
    p(UBound(p) - 1) = 120 - 1
    p(UBound(p)) = 120 + GH - 240 + 1
    ReDim Preserve p(UBound(p) + 2)
    p(UBound(p) - 1) = 120 - 1
    p(UBound(p)) = dY + Sin(GetTickCount / 6000) * (60 * (leftD / 365))
    Page.PaintPolygon2 argb(30 * pros2, 255, 255, 255), 0, p, False
    ReDim p(1)
    leftD = DateDiff("d", Now, Gaokao)
    dY = 100 + leftD / 365 * (GH - 240)
    p(0) = 120: p(1) = dY + Sin(GetTickCount / 3000 - 3) * (60 * (leftD / 365))
    For I = 0 To 30
        ReDim Preserve p(UBound(p) + 2)
        p(UBound(p) - 1) = 120 + (I * 1#) / 30 * (GW - 240)
        p(UBound(p)) = dY + Sin(GetTickCount / 3000 + I * pros2 / 30 * 5 - 3) * (60 * (leftD / 365))
    Next
    ReDim Preserve p(UBound(p) + 2)
    p(UBound(p) - 1) = 120 + (GW - 240)
    p(UBound(p)) = 120 + GH - 240
    ReDim Preserve p(UBound(p) + 2)
    p(UBound(p) - 1) = 120 - 1
    p(UBound(p)) = 120 + GH - 240 + 1
    ReDim Preserve p(UBound(p) + 2)
    p(UBound(p) - 1) = 120 - 1
    p(UBound(p)) = dY + Sin(GetTickCount / 3000 - 3) * (60 * (leftD / 365))
    Page.PaintPolygon2 argb(30 * pros2, 255, 255, 255), 0, p, False
    
    Page.Writes DateDiff("d", Now, Gaokao) & "��", 200, GH / 2 - 150, 96, argb(255 * pros2, 255, 255, 255), style:=FontStyleBold
    Page.Writes format(leftH, "00") & "Сʱ " & format(leftM, "00") & "����", 200, GH / 2 + 0, 80, argb(120 * pros2, 255, 255, 255), style:=FontStyleBold
    
    Page.Writes "����ʱ���ʣ " & format(Int((leftD + (leftH * 60 + leftM) / (24 * 60)) / 365 * 10000) / 100, "00.00") & "%", 120, GH - 94, 36, argb(150 * pros2, 255, 255, 255)
    Page.Writes format(Hour(Now), "00") & ":" & format(Minute(Now), "00"), GW - 120, GH - 94, 36, argb(150 * pros2, 255, 255, 255), align:=StringAlignmentFar
    
bypass2:
    If Mouse.button = 2 And Mouse.State >= 1 Then
        Mouse.button = 0: Mouse.State = 0
        If MsgBox("ȷ��Ҫ�˳��߿�����ʱ��", 48 Or vbYesNo, "��Ҫ���ҹص���QWQ") = vbYes Then
            Unload AppWindow
            End
        End If
    End If
    If Mouse.State >= 1 And mode And Mouse.button = 1 Then
        mode = Not mode
        useTime = GetTickCount
        switchTime = GetTickCount
        Mouse.State = 0
    End If
End Sub
Public Sub Wheel(Direction As Integer, Depth As Single)
    '�������¼�
    'Direction:����Depth�����
End Sub
Public Sub AnimationMsg(id As String, msg As String)
    '������Ϣ����
End Sub
Public Sub Enter()
    'ҳ������¼�
End Sub
Public Sub Leave()
    'ҳ���뿪�¼�
End Sub
Public Sub Gotfocus()
    '����ȡ�ý����¼�
End Sub
Public Sub Lostfocus()
    '����ʧȥ�����¼�
End Sub
Public Sub AnimationDone(id As Integer)
    '�������Ž����¼�
    'id�������Ķ���ID
End Sub
Private Sub Class_Initialize()
    Set Page = New GPage  '����ҳ�������
    Page.Create Me  '�����������������
    Page.Res.NewImages App.path & "\assets", 0.4 '������Դ
    Page.Res.newImage App.path & "\assets\lock.jpg", GW, GH, "lockbg"
    Page.Res.ApplyBlurEffect "lockbg", 60, 0
    ScrBlurDC = CreateCDC(GW, GH)
    ECore.Add Page, "AppPage"  '����ҳ��
    
    '===============================================
    '   �����Ҫ���Ӷ�����������ڴ˴�
    
    '===============================================
End Sub