VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "可怜的极域~"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   -105
   ClientWidth     =   4410
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   2  '屏幕中心
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'作者：张佑杰
'
'名称：MDIForm1.frm
'
'描述：显示母窗体的代码
'
'网站：https://www.johnzhang.xyz/
'
'邮箱：zsgsdesign@gmail.com
'
'遵循MIT协议，二次开发请注明原作者！
'****************************************************************************
Option Explicit
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MINIMIZEBOX = &H20000
Private Const GWL_STYLE = (-16)
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long


Private Sub MDIForm_Load()
  Dim lWnd As Long
Me.Caption = "电子教室学生端攻击系统-v" & App.Major & "." & App.Minor
Dim lStyle     As Long

        lStyle = GetWindowLong(Me.hwnd, GWL_STYLE)
        lStyle = lStyle And Not WS_MAXIMIZEBOX             '最大化
        lStyle = lStyle And Not WS_MINIMIZEBOX             '最小化
        lStyle = lStyle And Not WS_THICKFRAME             '可改变大小的边框
        SetWindowLong Me.hwnd, GWL_STYLE, lStyle
End Sub
