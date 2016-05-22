VERSION 5.00
Begin VB.Form Form999 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "嘿~核心在这里"
   ClientHeight    =   2490
   ClientLeft      =   -60
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "Form999.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2396.391
   ScaleMode       =   0  'User
   ScaleWidth      =   4606.531
   Begin VB.CommandButton cmdFree 
      Caption         =   "释放"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdDo 
      Caption         =   "执行"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtProgram 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Text            =   "StudentMain.exe"
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "进程名："
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   735
      Width           =   735
   End
End
Attribute VB_Name = "Form999"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'作者：张佑杰
'
'名称：Form999.frm
'
'描述：劫持程序到自己窗口里的代码
'
'网站：https://www.johnzhang.xyz/
'
'邮箱：zsgsdesign@gmail.com
'
'遵循MIT协议，二次开发请注明原作者！
'****************************************************************************
Option Explicit
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Const GW_HWNDNEXT = 2
Private old_parent As Long
Private child_hwnd As Long
Private Const TH32CS_SNAPPROCESS = &H2&

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

Public Function GetPsPid(sProcess As String) As Long
    Dim lSnapShot    As Long
    Dim lNextProcess As Long
    Dim tPE          As PROCESSENTRY32
    lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If lSnapShot <> -1 Then
        tPE.dwSize = Len(tPE)
        lNextProcess = Process32First(lSnapShot, tPE)
        Do While lNextProcess
            If LCase$(sProcess) = LCase$(Left(tPE.szExeFile, InStr(1, tPE.szExeFile, Chr(0)) - 1)) Then
                Dim lProcess  As Long
                Dim lExitCode As Long
                GetPsPid = tPE.th32ProcessID
                CloseHandle lProcess
            End If
            lNextProcess = Process32Next(lSnapShot, tPE)
        Loop
        CloseHandle (lSnapShot)
    End If
End Function

Private Function InstanceToWnd(ByVal target_pid As Long) As Long
Dim test_hwnd As Long
Dim test_pid As Long
Dim test_thread_id As Long

    ' 获取窗体的句柄。
    test_hwnd = FindWindow(ByVal 0&, ByVal 0&)

    ' 我们用 Do While 循环来找到目标窗口。
    ' 当已经发现目标或说 test_hwnd 不等于0，我们就跳出。
    Do While test_hwnd <> 0
        ' 看看这个窗口是否有父窗体，如果没有，他是一个最高层的窗口。
        If GetParent(test_hwnd) = 0 Then
            ' 这是一个最高层窗口，看看如果它具有目标实例句柄。
            test_thread_id = GetWindowThreadProcessId(test_hwnd, test_pid)
            If test_pid = target_pid Then
                ' 这是一个目标。
                InstanceToWnd = test_hwnd
                Exit Do
            End If
        End If

        ' 检查下一个窗口。
        test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
    Loop
End Function

Private Sub cmdDo_Click()
Dim pid As Long
Dim buf As String
Dim buf_len As Long
Dim styles As Long

    ' 获取PID（原来是运行的我把它修改了）。
    pid = GetPsPid(txtProgram.Text)
    If pid = 0 Then
        MsgBox "当前没有找到这个进程！"
        Exit Sub
    End If

    ' 获取这个窗口的句柄。
    child_hwnd = InstanceToWnd(pid)

    '让程序呆在这个MDI窗体里！
    old_parent = SetParent(child_hwnd, MDIForm1.hwnd)
    
    Me.cmdFree.Enabled = True
    
End Sub

Private Sub cmdFree_Click()
    If GetPsPid(txtProgram.Text) = 0 Then MsgBox "当前没有找到这个进程！", vbInformation, "提示!"
    SetParent child_hwnd, old_parent
    cmdDo.Enabled = True
    cmdFree.Enabled = False
End Sub

Private Sub exe_Click()
cmdDo_Click
End Sub

Private Sub put_Click()
cmdFree_Click
End Sub

Private Sub Form_Load()

End Sub
