VERSION 5.00
Begin VB.Form AppWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "�߿���Ļ"
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
   StartUpPosition =   2  '��Ļ����
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
'   ����ģ������Emerald������ �����������ڣ�Ӧ�ô��ڣ� ģ��
'==================================================
'   ҳ�������
    Dim EC As GMan
'==================================================
'   �ڴ˴��������ҳ���������ģ���������
    Dim AppPage As AppPage
'==================================================

Private Sub Form_Load()
    Me.Move 0, 0, Screen.Width / Screen.TwipsPerPixelX + 1, Screen.Height / Screen.TwipsPerPixelY + 1
    If App.LogMode <> 0 Then SetWindowPos AppWindow.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    StartEmerald Me.Hwnd, Screen.Width / Screen.TwipsPerPixelX + 1, Screen.Height / Screen.TwipsPerPixelY + 1 '��ʼ��Emerald���ڴ˴������޸Ĵ��ڴ�С��
    Set EF = New GFont
    EF.AddFont App.path & "\ui.ttf"
    EF.MakeFont "Aa�����п�"
   
    Set EC = New GMan   '����ҳ�������
    EC.Layered False
    '�����浵����ѡ�����浵key��������鿴Emerald��wiki
    Set ESave = New GSaving
    ESave.Create "GaokaoScreen.2022.Buger404", "��ֻ������gie gie~"
    
    '���������б���ѡ��
    'Set MusicList = New GMusicList
    'MusicList.Create App.path & "\music"

    '��ʼ��ʾ����
    Me.Show
    DrawTimer.Enabled = True
    
    '�ڴ˴�ʵ�������ҳ�������
    '=============================================
    'ʾ����TestPage.cls
    '     Set TestPage = New TestPage
    '�������֣�Dim TestPage As TestPage
        Set AppPage = New AppPage
    '=============================================

    '���ûҳ�棨�ڴ˴�������Ϊ�������ҳ�棩
    EC.ActivePage = "AppPage"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '��ֹ����
    DrawTimer.Enabled = False
    '�ͷ�Emerald��Դ
    EndEmerald
End Sub

Private Sub DrawTimer_Timer()
    '���ƽ��沢ˢ�´��ڻ���
    EC.Display
    DoEvents
End Sub

'============================================================
' �¼�ӳ��
Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, y As Single)
    '���������Ϣ
    UpdateMouse X, y, 1, button
End Sub
Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, y As Single)
    '���������Ϣ
    If Mouse.State = 0 Then
        UpdateMouse X, y, 0, button
    Else
        Mouse.X = X: Mouse.y = y
    End If
End Sub
Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, y As Single)
    '���������Ϣ
    UpdateMouse X, y, 2, button
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    '�����ַ�����
    If TextHandle <> 0 Then WaitChr = WaitChr & Chr(KeyAscii)
End Sub
'============================================================
