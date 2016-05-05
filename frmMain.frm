VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RTX for Celery"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Left            =   3720
      Top             =   2280
   End
   Begin VB.Timer Timer1 
      Left            =   1440
      Top             =   3720
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "Log"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4895
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' RTX用户状态改变接口
Dim WithEvents Presence As RTXCAPILib.RTXCPresence
Attribute Presence.VB_VarHelpID = -1

' 发送消息接口
Public g_imObj As RTXCMODULEINTERFACELib.IRTXIM

' 用户名到列表item的映射
Dim user_map As New Scripting.Dictionary

' 托盘图标
Private Type NotifyIconData
    Size              As Long
    Handle            As Long
    ID                As Long
    Flags             As Long
    CallBackMessage   As Long
    Icon              As Long
    Tip               As String * 64
End Type

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal Message As Long, Data As NotifyIconData) As Boolean

Private Const AddIcon = &H0
Private Const ModifyIcon = &H1
Private Const DeleteIcon = &H2
Private Const WM_USER = &H400
Private Const MessageFlag = &H1
Private Const IconFlag = &H2
Private Const TipFlag = &H4

Private g_icon_data As NotifyIconData

' 用来更新listview状态文本的子过程，记录日志
Private Sub update_status_text(ByVal user As String, ByVal RTXPresence As RTXCAPILib.RTX_PRESENCE, ByVal bLog As Boolean)
    If user_map.Exists(user) Then
        Dim s As String
        Dim t As String
        Dim item As ListItem

        If RTXPresence = RTX_PRESENCE_ONLINE Then
            s = "Online"
        ElseIf RTXPresence = RTX_PRESENCE_AWAY Then
            s = "Away"
        ElseIf RTXPresence = RTX_PRESENCE_OFFLINE Then
            s = "Offline"
        Else
            s = "(Unknown)"
        End If

        ' 到列表
        Set item = user_map.item(user)
        t = Format(Now, "MM-dd hh:nn")
        item.SubItems(2) = s
        item.SubItems(3) = t
        
        If bLog Then
            ' 更新日志
            Dim strLog As String
            strLog = item.SubItems(1) & vbTab & "->  " & s & vbTab & "@  " & t
    
            ' 到窗口
            g_txtLog = g_txtLog & strLog & vbCrLf
            
            ' 到托盘图标文本
            g_icon_data.Tip = item.SubItems(1) & " -> " & s & " @ " & t & vbNullChar
            Shell_NotifyIcon ModifyIcon, g_icon_data
            Timer2.Interval = 1000
            Timer2.Enabled = True
    
            ' 到文件
            Dim fileName As String
            Dim iFile As Integer
            iFile = FreeFile
            fileName = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "status.log"
            Open fileName For Append As #iFile
            Print #iFile, strLog
            Close #iFile
        End If
    End If
End Sub

' 点击“退出”按钮时
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHide_Click()
    Me.Hide
End Sub

' 显示“日志”
Private Sub cmdLog_Click()
    frmLog.Show 1, Me
End Sub

' 点击“刷新”手动刷新列表
Private Sub cmdRefresh_Click()
    Dim i As Integer
    For i = 1 To ListView1.ListItems.Count
        Dim user, status As String
        user = ListView1.ListItems(i).Text
        status = Presence.RTXPresence(user)
        update_status_text user, status, False
    Next
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

' 主窗体加载时
Private Sub Form_Load()
    ' 不应该多次运行
    If App.PrevInstance Then
        MsgBox "RTX for Celery 已经运行。"
        End
    End If
    
    ' 托盘图标
    With g_icon_data
        .Size = Len(g_icon_data)
        .Handle = Me.hWnd
        .ID = vbNull
        .Flags = IconFlag Or TipFlag Or MessageFlag
        .CallBackMessage = WM_USER + 128
        .Icon = Me.Icon
        .Tip = "RTX for Celery" & vbNullChar
    End With
    Call Shell_NotifyIcon(AddIcon, g_icon_data)
    
    ' 不太清楚，但可以使得按ESC关闭窗口
    Me.KeyPreview = True
    
    ' 打开消息对话框接口，不知道为什么在WndProc里面拿不到
    Set g_imObj = CreateObject("RTXClient.RTXAPI").GetObject("AppRoot").GetAppObject("RTXPlugin.IM")
                
    ' RTX用户在线状态对象
    Set Presence = CreateObject("RTXClient.RTXAPI").GetObject("kernalRoot").Presence
    
    ' 列表初始化
    With ListView1
        .FullRowSelect = True
        .GridLines = True
        .LabelEdit = lvwManual
        .HideSelection = False
    End With
    
    With ListView1.ColumnHeaders
        .Add , , "Account", 0
        .Add , , "User", 1000
        .Add , , "Status", 960
        .Add , , "Time", 1200
    End With
    
    ' 读取被监控的用户列表
    ' 格式一定是：小写RTX用户名,自定义显示名称
    Dim fileName As String
    Dim iFile As Integer
    iFile = FreeFile
    fileName = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & "users.txt"
    Open fileName For Input As #iFile
        Do While Not EOF(iFile)
            Dim textline As String
            Dim tokens() As String
            Dim i As Integer
            Dim item As ListItem
            
            Line Input #iFile, textline
            tokens = Split(textline, ",")
            
            Set item = ListView1.ListItems.Add(, , tokens(0))
            item.SubItems(1) = tokens(1)
            user_map.Add tokens(0), item
        Loop
    Close #iFile
    
    ' 子类化以接收自定义消息
    g_WndProc = SetWindowLong(frmMain.hWnd, GWL_WNDPROC, AddressOf SubWndProc)
    
    ' 刚打开时自动刷新列表状态
    cmdRefresh_Click
End Sub

' RTX用户在线状态改变回调
Private Sub Presence_OnPresenceChange(ByVal Account As String, ByVal RTXPresence As RTXCAPILib.RTX_PRESENCE)
    update_status_text Account, RTXPresence, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetWindowLong frmMain.hWnd, GWL_WNDPROC, g_WndProc
    Shell_NotifyIcon DeleteIcon, g_icon_data
End Sub

' 在这里直接调用发消息接口
Private Sub Timer1_Timer()
    frmMain.Timer1.Enabled = False
    g_imObj.SendIMEx g_msgUsers, "", ""
End Sub

Public Sub show_trayicon(ByVal bshow As Boolean)
    If bshow Then
        g_icon_data.Icon = Me.Icon
    Else
        g_icon_data.Icon = vbNull
    End If

    Shell_NotifyIcon ModifyIcon, g_icon_data
End Sub

Public Sub reset_trayicon()
    g_icon_data.Tip = "RTX for Celery" & vbNullChar
    Shell_NotifyIcon ModifyIcon, g_icon_data
End Sub
Private Sub Timer2_Timer()
    show_trayicon g_icon_data.Icon = vbNull
End Sub
