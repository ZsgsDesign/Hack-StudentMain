VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "电子教室学生端攻击系统-v4.0"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   4560
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0442
   ScaleHeight     =   3840
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重开客户端"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2040
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "请勿将该系统用于不符合法律、道德的行为，本软件仅供技术交流，擅自用于不正当行为所造成的后果自行承担！"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   3240
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " 张佑杰 设计制作"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.Menu menu 
      Caption         =   "caidan"
      Visible         =   0   'False
      Begin VB.Menu way1 
         Caption         =   "攻击并卸载(不推荐)"
      End
      Begin VB.Menu way2 
         Caption         =   "只关闭程序"
      End
      Begin VB.Menu way3 
         Caption         =   "关闭程序-2"
      End
      Begin VB.Menu way4 
         Caption         =   "劫持程序（可脱离屏幕控制，未测试）"
      End
   End
   Begin VB.Menu menufile 
      Caption         =   "文件"
      Begin VB.Menu exit 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu menuhelp 
      Caption         =   "帮助"
      Begin VB.Menu help 
         Caption         =   "帮助"
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "关于"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'作者：张佑杰
'
'名称：Form1.frm
'
'描述：主程序所在子窗体的代码
'
'网站：https://www.johnzhang.xyz/
'
'邮箱：zsgsdesign@gmail.com
'
'遵循MIT协议，二次开发请注明原作者！
'****************************************************************************
Private Sub about_Click()
frmAbout.Show vbModal
End Sub

Private Sub Command2_Click()
On Error Resume Next

If GetSetting("HackingSystem", "User", "Way") = "way2" Then
Command2.Visible = False
Command1.Enabled = True
Command1.Width = 3855
SaveSetting "HackingSystem", "User", "Way", ""
X = Shell("C:\Program Files\Mythware\e-Learning Class\StudentMain.exe")
SaveSetting "HackingSystem", "User", "Way", ""
MsgBox "完成！", vbOKOnly, "提示"
Else
End If

If GetSetting("HackingSystem", "User", "Way") = "way3" Then
Name "C:\Program Files\Mythware\e-Learning Class\StudentMain1.exe" As "C:\Program Files\Mythware\e-Learning Class\StudentMain.exe"
Command2.Visible = False
Command1.Enabled = True
Command1.Width = 3855
SaveSetting "HackingSystem", "User", "Way", ""
X = Shell("C:\Program Files\Mythware\e-Learning Class\StudentMain.exe")
MsgBox "完成！", vbOKOnly, "提示"
Else
End If

If GetSetting("HackingSystem", "User", "Way") = "way4" Then
Name "C:\Program Files\Mythware\e-Learning Class\MediaPlayer1.exe" As "C:\Program Files\Mythware\e-Learning Class\MediaPlayer.exe"
Name "C:\Program Files\Mythware\e-Learning Class\TDChalk1.exe" As "C:\Program Files\Mythware\e-Learning Class\TDChalk.exe"
Name "C:\Program Files\Mythware\e-Learning Class\MasterHelper1.exe" As "C:\Program Files\Mythware\e-Learning Class\MasterHelper.exe"
Name "C:\Program Files\Mythware\e-Learning Class\TDOvrSet1.exe" As "C:\Program Files\Mythware\e-Learning Class\TDOvrSet.exe"
Name "C:\Program Files\Mythware\e-Learning Class\VRecordClt_Personal1.exe" As "C:\Program Files\Mythware\e-Learning Class\VRecordClt_Personal.exe"
Command2.Visible = False
Command1.Enabled = True
Command1.Width = 3855
SaveSetting "HackingSystem", "User", "Way", ""
X = Shell("C:\Program Files\Mythware\e-Learning Class\StudentMain.exe")
MsgBox "完成！", vbOKOnly, "提示"
Else
End If

End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
On Error Resume Next
If GetSetting("HackingSystem", "User", "Way") = "way2" Then
Command1.Width = 1575
Command2.Visible = True
Command1.Enabled = False
Else
End If

If GetSetting("HackingSystem", "User", "Way") = "way3" Then
Command1.Width = 1575
Command2.Visible = True
Command1.Enabled = False
Else
End If
If GetSetting("HackingSystem", "User", "Way") = "way4" Then
Command1.Width = 1575
Command2.Visible = True
Command1.Enabled = False
Else
End If

Me.Caption = "电子教室学生端攻击系统-v" & App.Major & "." & App.Minor
End Sub

Private Sub help_Click()
Frmtip.Show vbModal
End Sub

Private Sub way1_Click()
On Error Resume Next
If MsgBox("请注意完全关闭后无法再接受文件，确定继续？", vbYesNo, "提示") = vbYes Then

X = Shell("ntsd -c q -pn StudentMain.exe", 1)
MsgBox "请在弹出窗口中(如果有)按回车再继续", vbExclamation, "提示"
Kill "C:\Program Files\Mythware\e-Learning Class\*.dll"
Kill "C:\Program Files\Mythware\e-Learning Class\TDChalk.exe"
Kill "C:\Program Files\Mythware\e-Learning Class\TDOvrSet.exe"
Kill "C:\Program Files\Mythware\e-Learning Class\*.dat"
Kill "C:\Program Files\Mythware\e-Learning Class\*.manifest"
Kill "C:\Program Files\Mythware\e-Learning Class\thirdpartylicense.rtf"
Kill "C:\Program Files\Mythware\e-Learning Class\VRecordClt_Personal.exe"
Kill "C:\Program Files\Mythware\e-Learning Class\*.htm"
Kill "C:\Program Files\Mythware\e-Learning Class\*.jpg"
Kill "C:\Program Files\Mythware\e-Learning Class\StudentMain.exe"
Kill "C:\Program Files\Mythware\e-Learning Class\Shutdown.exe"
Kill "C:\Program Files\Mythware\e-Learning Class\MediaPlayer.exe"
Kill "C:\Program Files\Mythware\e-Learning Class\MasterHelper.exe"
Kill "C:\Program Files\Mythware\e-Learning Class\*.wav"
Kill "C:\Program Files\Mythware\e-Learning Class\InstHelpApp.exe"
Kill "C:\Program Files\Mythware\e-Learning Class\Error.exe"
Kill "C:\Program Files\Mythware\e-Learning Class\cad.exe"
Kill "C:\Program Files\Mythware\e-Learning Class\Language\2052\*.*"
Kill "C:\Program Files\Mythware\e-Learning Class\Real\*.*"
Kill "C:\Program Files\Mythware\e-Learning Class\uninst\*.*"
SaveSetting "HackingSystem", "User", "Way", "way1"
MsgBox "完成！", vbOKOnly, "提示"
Else
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next

PopupMenu menu
End Sub

Private Sub way2_Click()
On Error Resume Next
X = Shell("ntsd -c q -pn StudentMain.exe", 1)
MsgBox "请在弹出窗口中(如果有)按回车再继续", vbExclamation, "提示"
Command1.Width = 1575
Command2.Visible = True
Command1.Enabled = False
SaveSetting "HackingSystem", "User", "Way", "way2"
MsgBox "完成！", vbOKOnly, "提示"


End Sub
Private Sub way3_Click()
On Error Resume Next


X = Shell("ntsd -c q -pn StudentMain.exe", 1)
MsgBox "请在弹出窗口中(如果有)按回车再继续", vbExclamation, "提示"
Name "C:\Program Files\Mythware\e-Learning Class\StudentMain.exe" As "C:\Program Files\Mythware\e-Learning Class\StudentMain1.exe"
Command1.Width = 1575
Command2.Visible = True
Command1.Enabled = False
SaveSetting "HackingSystem", "User", "Way", "way3"
MsgBox "完成！", vbOKOnly, "提示"

End Sub
Private Sub way4_Click()
On Error Resume Next


'X = Shell("ntsd -c q -pn StudentMain.exe", 1)
MsgBox "请继续", vbExclamation, "提示"
'Name "C:\Program Files\Mythware\e-Learning Class\MediaPlayer.exe" As "C:\Program Files\Mythware\e-Learning Class\MediaPlayer1.exe"
'Name "C:\Program Files\Mythware\e-Learning Class\TDChalk.exe" As "C:\Program Files\Mythware\e-Learning Class\TDChalk1.exe"
'Name "C:\Program Files\Mythware\e-Learning Class\MasterHelper.exe" As "C:\Program Files\Mythware\e-Learning Class\MasterHelper1.exe"
'Name "C:\Program Files\Mythware\e-Learning Class\TDOvrSet.exe" As "C:\Program Files\Mythware\e-Learning Class\TDOvrSet1.exe"
'Name "C:\Program Files\Mythware\e-Learning Class\VRecordClt_Personal.exe" As "C:\Program Files\Mythware\e-Learning Class\VRecordClt_Personal1.exe"
'Command1.Width = 1575
'Command2.Visible = True
'Command1.Enabled = False
'SaveSetting "HackingSystem", "User", "Way", "way4"
'MsgBox "完成！", vbOKOnly, "提示"
Form999.Visible = True

End Sub
