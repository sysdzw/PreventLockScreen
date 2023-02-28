Attribute VB_Name = "modTuopan"
Option Explicit

Public Const DefaultIconIndex = 1    'ͼ��ȱʡ����
Public Const WM_LBUTTONDOWN = &H201    '��������
Public Const WM_RBUTTONDOWN = &H204    '������Ҽ�
Public Const WM_RBUTTONUP = &H205
Public Const NIM_ADD = 0    '���ͼ��
Public Const NIM_MODIFY = 1    '�޸�ͼ��
Public Const NIM_DELETE = 2    'ɾ��ͼ��
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = 1    'message ��Ч
Public Const NIF_ICON = 2    'ͼ���������ӡ��޸ġ�ɾ������Ч
Public Const NIF_TIP = 4    'ToolTip(��ʾ����Ч

'ͼ�����
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
'�жϴ����Ƿ���С��
Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
'���ô���λ�ú�״̬��position���Ĺ���
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'��������
'֪ͨ��ͼ��״̬
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'���ͼ����֪ͨ��
Public Function Icon_Add(iHwnd As Long, sTips As String, hIcon As Long, IconID As Long) As Long
    '����˵����iHwnd�����ھ����sTips��������Ƶ�֪ͨ��ͼ����ʱ��ʾ����ʾ����
    'hIcon��ͼ������IconID��ͼ��Id��
Dim IconVa As NOTIFYICONDATA
    With IconVa
        .hwnd = iHwnd
        .szTip = sTips + Chr$(0)
        .hIcon = hIcon
        .uID = IconID
        .uCallbackMessage = WM_MOUSEMOVE
        .cbSize = Len(IconVa)
        .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
        Icon_Add = Shell_NotifyIcon(NIM_ADD, IconVa)
    End With
End Function
'ɾ��֪ͨ��ͼ��(����˵��ͬIcon_Add)
Function Icon_Del(iHwnd As Long, lIndex As Long) As Long
Dim IconVa As NOTIFYICONDATA
Dim L As Long
    With IconVa
        .hwnd = iHwnd
        .uID = lIndex
        .cbSize = Len(IconVa)
    End With
    Icon_Del = Shell_NotifyIcon(NIM_DELETE, IconVa)
End Function

