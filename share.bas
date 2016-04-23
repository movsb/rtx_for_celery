Attribute VB_Name = "Share"
Option Explicit

Public g_txtLog As String
Public g_msgUsers As String


' 子类化窗口以接收自定义消息
Private Const WM_APP As Long = 32768
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

Public Function SubWndProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If hWnd = frmMain.hWnd And iMsg = WM_COPYDATA Then
        Dim cds As COPYDATASTRUCT
        CopyMemory cds, ByVal lParam, Len(cds)
        
        ' 类型为发送消息
        If cds.dwData = 1 Then
            If cds.cbData <= 4096 Then
                Dim str As String
                Dim buf(1 To 4096) As Byte
                CopyMemory buf(1), ByVal cds.lpData, cds.cbData
                str = StrConv(buf, vbUnicode)
                ' 直接在这里处理会报COM自动化错误，所以弄个定时器去操作
                g_msgUsers = str
                frmMain.Timer1.Interval = 100
                frmMain.Timer1.Enabled = True
            End If
        End If
    End If

    SubWndProc = CallWindowProc(g_WndProc, hWnd, iMsg, wParam, lParam)
End Function

