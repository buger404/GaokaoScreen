VERSION 5.00
Begin VB.Form AppWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "高考屏幕"
   ClientHeight    =   6670
   ClientLeft      =   10
   ClientTop       =   10
   ClientWidth     =   9660
   LinkTopic       =   "AppWindow"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   667
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   966
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   9000
      Top             =   240
   End
End
Attribute VB_Name = "AppWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================
'   该类模块是由Emerald创建的 界面容器窗口（应用窗口） 模板
'==================================================
'   页面管理器
    Dim EC As GMan
'==================================================
'   在此处放置你的页面控制器类模块声明语句
    Dim AppPage As AppPage
'==================================================

Private Sub Form_Load()
    Me.Move 0, 0, Screen.Width / Screen.TwipsPerPixelX + 1, Screen.Height / Screen.TwipsPerPixelY + 1
    If App.LogMode <> 0 Then SetWindowPos AppWindow.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    StartEmerald Me.Hwnd, Screen.Width / Screen.TwipsPerPixelX + 1, Screen.Height / Screen.TwipsPerPixelY + 1 '初始化Emerald（在此处可以修改窗口大小）
    Set EF = New GFont
    EF.AddFont App.path & "\ui.ttf"
    EF.MakeFont "Aa马上行楷"
   
    Set EC = New GMan   '创建页面管理器
    EC.Layered False
    '创建存档（可选），存档key的问题请查看Emerald的wiki
    Set ESave = New GSaving
    ESave.Create "GaokaoScreen.2022.Buger404", "我只会心疼gie gie~"
    
    '创建音乐列表（可选）
    'Set MusicList = New GMusicList
    'MusicList.Create App.path & "\music"

    '开始显示界面
    Me.Show
    DrawTimer.Enabled = True
    
    '在此处实例化你的页面控制器
    '=============================================
    '示例：TestPage.cls
    '     Set TestPage = New TestPage
    '公共部分：Dim TestPage As TestPage
        Set AppPage = New AppPage
    '=============================================

    '设置活动页面（在此处设置则为你的启动页面）
    EC.ActivePage = "AppPage"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '终止绘制
    DrawTimer.Enabled = False
    '释放Emerald资源
    EndEmerald
End Sub

Private Sub DrawTimer_Timer()
    '绘制界面并刷新窗口画面
    EC.Display
    DoEvents
End Sub

'============================================================
' 事件映射
Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, y As Single)
    '发送鼠标信息
    UpdateMouse X, y, 1, button
End Sub
Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, y As Single)
    '发送鼠标信息
    If Mouse.State = 0 Then
        UpdateMouse X, y, 0, button
    Else
        Mouse.X = X: Mouse.y = y
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, y As Single)
    '发送鼠标信息
    UpdateMouse X, y, 2, button
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    '发送字符输入
    If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
End Sub
'============================================================
