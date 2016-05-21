VERSION 5.00
Begin VB.Form Frmtip 
   BackColor       =   &H00FFFFFF&
   Caption         =   "帮助"
   ClientHeight    =   3000
   ClientLeft      =   9645
   ClientTop       =   6330
   ClientWidth     =   5685
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5685
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "下一条帮助(&N)"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   120
      Picture         =   "Form2.frx":0442
      ScaleHeight     =   2685
      ScaleWidth      =   3705
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "四种方式..."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   540
         TabIndex        =   4
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   180
         TabIndex        =   3
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Frmtip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************
'作者：张佑杰
'
'名称：Form2.frm
'
'描述：显示帮助的代码
'
'网站：https://www.johnzhang.xyz/
'
'邮箱：zsgsdesign@gmail.com
'
'遵循MIT协议，二次开发请注明原作者！
'****************************************************************************
Option Explicit

' 内存中的提示数据库。
Dim Tips As New Collection

' 提示文件名称
Const TIP_FILE = "TIPOFDAY.TXT"

' 当前正在显示的提示集合的索引。
Dim CurrentTip As Long


Private Sub DoNextTip()

    ' 随机选择一条提示。
    CurrentTip = Int((Tips.Count * Rnd) + 1)
    
    ' 或者，您可以按顺序遍历提示

'    CurrentTip = CurrentTip + 1
'    If Tips.Count < CurrentTip Then
'        CurrentTip = 1
'    End If
    
    ' 显示它。
    Frmtip.DisplayCurrentTip
    
End Sub

Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' 从文件中读出的每条提示。
    Dim InFile As Integer   ' 文件的描述符。
    
    ' 包含下一个自由文件描述符。
    InFile = FreeFile
    
    ' 确定为指定文件。
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' 在打开前确保文件存在。
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' 从文本文件中读取集合。
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' 随机显示一条提示。
    DoNextTip
    
    LoadTips = True
    
End Function



Private Sub cmdNextTip_Click()
    DoNextTip
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim ShowAtStartup As Long
    
    ' 察看在启动时是否将被显示
    ShowAtStartup = GetSetting(App.EXEName, "Options", "在启动时显示提示", 1)
    If ShowAtStartup = 0 Then
        Unload Me
        Exit Sub
    End If
        
    
    ' 随机寻找
    Randomize
    
    ' 读取提示文件并且随机显示一条提示。
    If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        lblTipText.Caption = "文件 " & TIP_FILE & " 没有被找到吗? " & vbCrLf & vbCrLf & _
           "创建文本文件名为 " & TIP_FILE & " 使用记事本每行写一条提示。 " & _
           "然后将它存放在应用程序所在的目录 "
    End If

    
End Sub

Public Sub DisplayCurrentTip()
    If Tips.Count > 0 Then
        lblTipText.Caption = Tips.Item(CurrentTip)
    End If
End Sub
