Attribute VB_Name = "Module1"
'****************************************************************************
'作者：Skin Sharp
'
'名称：ModSkin.bas
'
'描述：磨砂玻璃效果通用库模块的代码
'
'网站：http://www.skinsharp.com/htdocs/index.htm
'
'商业授权，为保护作者版权请购买后将模块代码添加到下方使用！
'****************************************************************************

Public Declare Function SkinH_Attach Lib "SkinH.dll" () As Long

Public Declare Function SkinH_AttachEx Lib "SkinH.dll" (ByVal lpSkinFile As String, ByVal lpPasswd As String) As Long

Public Declare Function SkinH_AttachExt Lib "SkinH.dll" (ByVal lpSkinFile As String, ByVal lpPasswd As String, ByVal nHue As Integer, ByVal nSat As Integer, ByVal nBri As Integer) As Long

Public Declare Function SkinH_AttachRes Lib "SkinH.dll" (lpRes As Any, ByVal nSize As Long, ByVal lpPasswd As String, ByVal nHue As Integer, ByVal nSat As Integer, ByVal nBri As Integer) As Long

Public Declare Function SkinH_AdjustHSV Lib "SkinH.dll" (ByVal nHue As Integer, ByVal nSat As Integer, ByVal nBri As Integer) As Long

Public Declare Function SkinH_Detach Lib "SkinH.dll" () As Long

Public Declare Function SkinH_DetachEx Lib "SkinH.dll" (ByVal hwnd As Long) As Long

Public Declare Function SkinH_SetAero Lib "SkinH.dll" (ByVal hwnd As Long) As Long

Public Declare Function SkinH_SetWindowAlpha Lib "SkinH.dll" (ByVal hwnd As Long, ByVal nAlpha As Integer) As Long

Public Declare Function SkinH_SetMenuAlpha Lib "SkinH.dll" (ByVal nAlpha As Integer) As Long

Public Declare Function SkinH_GetColor Lib "SkinH.dll" (ByVal hwnd As Long, ByVal nPosX As Integer, ByVal nPosY As Integer) As Long

Public Declare Function SkinH_Map Lib "SkinH.dll" (ByVal hwnd As Long, ByVal nType As Integer) As Long

Public Declare Function SkinH_LockUpdate Lib "SkinH.dll" (ByVal hwnd As Long, ByVal nLocked As Integer) As Long

Public Declare Function SkinH_SetBackColor Lib "SkinH.dll" (ByVal hwnd As Long, ByVal nRed As Integer, ByVal nGreen As Integer, ByVal nBlue As Integer) As Long

Public Declare Function SkinH_SetForeColor Lib "SkinH.dll" (ByVal hwnd As Long, ByVal nRed As Integer, ByVal nGreen As Integer, ByVal nBlue As Integer) As Long

Public Declare Function SkinH_SetWindowMovable Lib "SkinH.dll" (ByVal hwnd As Long, ByVal bMove As Integer) As Long

Public Declare Function SkinH_AdjustAero Lib "SkinH.dll" (ByVal nAlpha As Integer, ByVal nShwDark As Integer, ByVal nShwSharp As Integer, ByVal nShwSize As Integer, ByVal nX As Integer, ByVal nY As Integer, ByVal nRed As Integer, ByVal nGreen As Integer, ByVal nBlue As Integer) As Long

Public Declare Function SkinH_NineBlt Lib "SkinH.dll" (ByVal hDtDC As Long, ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer, ByVal nMRect As Integer) As Long

Public Declare Function SkinH_SetTitleMenuBar Lib "SkinH.dll" (ByVal hwnd As Long, ByVal bEnable As Integer, ByVal nMenuY As Integer, ByVal nTopOffs As Integer, ByVal nRightOffs As Integer) As Long

Public Declare Function SkinH_SetFont Lib "SkinH.dll" (ByVal hwnd As Long, ByVal hFont As Long) As Long

Public Declare Function SkinH_SetFontEx Lib "SkinH.dll" (ByVal hwnd As Long, ByVal szFace As String, ByVal nHeight As Integer, ByVal nWidth As Integer, ByVal nWeight As Integer, ByVal nItalic As Integer, ByVal nUnderline As Integer, ByVal nStrikeOut As Integer) As Long

Public Declare Function SkinH_VerifySign Lib "SkinH.dll" () As Long


