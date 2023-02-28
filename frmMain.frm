VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "�������ƶ�������"
   ClientHeight    =   2655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   6720
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command2 
      Caption         =   "ͣ ֹ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Text            =   "5"
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ�������ƶ�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      TabIndex        =   3
      Top             =   390
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����������:"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   390
      Width           =   1500
   End
   Begin VB.Menu menu 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu mnuQuit 
         Caption         =   "�˳�(&E)"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "����(&A)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim isDraw As Boolean
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Sub Form_Load()
    Call Icon_Add(Me.hwnd, Me.Caption, Me.Icon, 0) '������ͼ�����֪ͨ��
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lMsg As Single
    lMsg = X / Screen.TwipsPerPixelX
    Select Case lMsg
    Case WM_RBUTTONUP
        SetForegroundWindow (hwnd)
        PopupMenu menu
    Case WM_LBUTTONDOWN
        Me.WindowState = 0
        Me.Show
'        Call Icon_Del(Form1.hwnd, 0) '��ʾ������ʱɾ������
    End Select
End Sub

Private Sub Form_Resize() '�жϴ����Ƿ���С��״̬�������ǰ���С����Ŧ���һ�η���Resize�¼�
    If IsIconic(Me.hwnd) <> 0 Then
        Me.Visible = False
        Call Icon_Add(Me.hwnd, Me.Caption, Me.Icon, 0)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Icon_Del(Me.hwnd, 0)
    End
End Sub

Private Sub mnuAbout_Click()
    Dim strInfo$
    strInfo = "�������ƶ������� V" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf & _
        "  ����:sysdzw" & vbCrLf & _
        "  ��ҳ:https://blog.csdn.net/sysdzw" & vbCrLf & _
        "  Q  Q:171977759" & vbCrLf & _
        "  ����:sysdzw@163.com" & vbCrLf & vbCrLf & _
        "2023-02-28"
        
'    Call Icon_Del(Form1.hwnd, 0)
    MsgBox strInfo, vbInformation
'    Call Icon_Add(Me.hwnd, Me.Caption, Me.Icon, 0)
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Command1.Enabled = False
    Command2.Enabled = True
    isDraw = True
    Dim w As New clsWindow
    Dim i%
    
    Randomize
    Do While isDraw
        If Int(Rnd * 100) Mod 2 = 0 Then
            drawACircle
        Else
            DrawALine
        End If
        w.Wait Val(Text1.Text) * 1000 '�ȴ�N����
    Loop
    Command1.Enabled = True
    Command2.Enabled = False
End Sub
'�������Բ
Private Sub drawACircle()
    Dim w As New clsWindow
    Dim X As Double, Y As Double
    Dim sW&, sH&
    Dim k As Single
    Dim R As Double
    
    sW = Screen.Width \ 15
    sH = Screen.Height \ 15
    Const pi As Single = 3.14159
    
    Randomize
    
    X = (sW - 300) * Rnd + 300
    Y = (sH - 500) * Rnd + 500
    R = sH * Rnd + sH / 4
    Me.Caption = R

    Do While k < 2 * pi
        w.SetCursor Cos(k) * R / 4 + X, Sin(k) * R / 4 + Y, , , 5
        k = k + pi / 180
        DoEvents
    Loop
End Sub
'���������
Private Sub DrawALine()
    Dim w As New clsWindow
    Dim x1&, y1&, xPad&, yPad&, lngWidth&, i&, intRndType%, intRndType2%
    Dim sW&, sH&
    sW = Screen.Width \ 15
    sH = Screen.Height \ 15
    
    Randomize
    x1 = sW * Rnd
    y1 = sH * Rnd

    xPad = IIf(x1 > sW / 2, -1, 1)
    yPad = IIf(y1 > sH / 2, -1, 1)
    
    lngWidth = sH * Rnd / 2 + sH / 4
    intRndType = Int(Rnd * 2)
    intRndType2 = Int(Rnd * 2)
    
    For i = 1 To lngWidth
        If intRndType = 0 Then
            If intRndType2 = 0 Then
                x1 = x1 + xPad
            Else
                y1 = y1 + yPad
            End If
        Else
            x1 = x1 + xPad
            y1 = y1 + yPad
        End If
        w.SetCursor x1, y1
        w.Wait 5
    Next
End Sub

Private Sub Command2_Click()
    isDraw = False
End Sub

'Private Sub SetRndPoint()
'    Dim w As New clsWindow
'    Randomize
'    Do
'        x1 = Screen.Width / 15 * Rnd
'        y1 = Screen.Height / 15 * Rnd
'        w.SetCursor x1, y1 '����ƶ�����Ļ������һ����
'        w.Wait 5000 '�ȴ�5����
'    Loop
'End Sub
