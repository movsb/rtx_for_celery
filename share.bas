Attribute VB_Name = "Share"
Option Explicit

Public g_txtLog As String
Public g_msgUsers As String


' ���໯�����Խ����Զ�����Ϣ
Private Const WM_USER As Long = &H400
Private Const WM_COPYDATA As Long = &H4A
Public Const GWL_WNDPROC = (-4)
Public g_WndProc As Long

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function CallWindowProc Lib "USER32.DLL" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "USER32.DLL" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Public Const WM_RBUTTONUP As Long = &H205
Public Const WM_LBUTTONDBLCLK As Long = &H203

Public Function SubWndProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If hWnd = frmMain.hWnd And iMsg = WM_COPYDATA Then
        Dim cds As COPYDATASTRUCT
        CopyMemory cds, ByVal lParam, Len(cds)
        
        ' ����Ϊ������Ϣ
        If cds.dwData = 1 Then
            If cds.cbData <= 4096 Then
                Dim str As String
                Dim buf(1 To 4096) As Byte
                CopyMemory buf(1), ByVal cds.lpData, cds.cbData ' �������������
                str = StrConv(buf, vbUnicode)
                ' ֱ�������ﴦ��ᱨCOM�Զ�����������Ū����ʱ��ȥ����
                g_msgUsers = str
                frmMain.Timer1.Interval = 100
                frmMain.Timer1.Enabled = True
            End If
        End If
    ElseIf hWnd = frmMain.hWnd And iMsg = WM_USER + 128 Then
         If lParam = WM_RBUTTONUP Then
            frmMain.Timer2.Enabled = False
            frmMain.show_trayicon True
            frmMain.reset_trayicon
            frmMain.Hide
        ElseIf lParam = WM_LBUTTONDBLCLK Then
            ' ����ǰ״̬����С�������л�Ϊ��ԭ
            If frmMain.WindowState = 1 Then
                frmMain.WindowState = 0
            End If
            frmMain.Show
        End If
    End If

    SubWndProc = CallWindowProc(g_WndProc, hWnd, iMsg, wParam, lParam)
End Function

