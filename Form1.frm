VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RTX for Celery"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4125
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4125
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' RTX用户状态改变接口
Dim WithEvents Presence As RTXCAPILib.RTXCPresence
Attribute Presence.VB_VarHelpID = -1

' 用户名到列表item的映射
Dim user_map As New Scripting.Dictionary

' 用来更新listview状态文本的子过程
Private Sub update_status_text(ByVal user As String, ByVal RTXPresence As RTXCAPILib.RTX_PRESENCE)
    Dim s As String
    If RTXPresence = RTX_PRESENCE_ONLINE Then
        s = "Online"
    ElseIf RTXPresence = RTX_PRESENCE_AWAY Then
        s = "Away"
    ElseIf RTXPresence = RTX_PRESENCE_OFFLINE Then
        s = "Offline"
    Else
        s = "(Unknown)"
    End If
        If user_map.Exists(user) Then
        Dim item As ListItem
        Set item = user_map.item(user)
        item.SubItems(2) = s
        item.SubItems(3) = FormatDateTime(Now, vbShortTime)
    End If
End Sub

' 点击“退出”按钮时
Private Sub cmdExit_Click()
    Unload Me
End Sub

' 点击“刷新”手动刷新列表
Private Sub cmdRefresh_Click()
    Dim i As Integer
    For i = 1 To ListView1.ListItems.Count
        Dim user, status As String
        user = ListView1.ListItems(i).Text
        status = Presence.RTXPresence(user)
        update_status_text user, status
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
    
    Me.KeyPreview = True
    
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
        .Add , , "Time", 800
    End With
    
    ' 读取被监控的用户列表
    ' 格式一定是：小写RTX用户名,自定义显示名称
    Dim iFile As Integer
    iFile = FreeFile
    Open "C:\Users\Tao\Desktop\rtx_for_celery\users.txt" For Input As #iFile
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
End Sub

' RTX用户在线状态改变回调
Private Sub Presence_OnPresenceChange(ByVal Account As String, ByVal RTXPresence As RTXCAPILib.RTX_PRESENCE)
    update_status_text Account, RTXPresence
End Sub
