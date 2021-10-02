VERSION 5.00
Object = "{C7D9E622-6CAA-42B3-9460-0E9DA4FD385A}#1.0#0"; "FBSLib.ocx"
Object = "{9A226D6F-2658-4445-8D35-5C19D42676FE}#1.0#0"; "BSE.ocx"
Begin VB.Form FrmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "日志文件设置..."
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
   Icon            =   "FrmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   568
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   425
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtDescription 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   5895
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1920
      Width           =   6135
   End
   Begin VB.CommandButton cmdDonate 
      Caption         =   "赞助作者1分钱  >>"
      Height          =   435
      Left            =   4080
      TabIndex        =   7
      Top             =   7920
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   120
      TabIndex        =   5
      Top             =   90
      Width           =   6165
      Begin VB.TextBox TxtPath 
         Height          =   285
         Left            =   180
         MaxLength       =   500
         TabIndex        =   0
         Top             =   585
         Width           =   4785
      End
      Begin VB.CommandButton CmdBrowser 
         Caption         =   "浏览..."
         Height          =   320
         Left            =   5130
         TabIndex        =   1
         Top             =   570
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "SQL日志文件的保存目录(文件名为SQL.txt):"
         Height          =   240
         Left            =   180
         TabIndex        =   6
         Top             =   270
         Width           =   4200
      End
   End
   Begin VB.CommandButton CmdOpen 
      Caption         =   "打开该目录(&O)"
      Height          =   320
      Left            =   120
      TabIndex        =   2
      Top             =   1260
      Width           =   1500
   End
   Begin BSE_Engine.BSE BSE1 
      Left            =   -2370
      Top             =   570
      _ExtentX        =   6588
      _ExtentY        =   1085
   End
   Begin FBSLib.FolderBrowser FolderBrowser1 
      Left            =   1470
      Top             =   1050
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "取消"
      Height          =   320
      Left            =   5145
      TabIndex        =   4
      Top             =   1260
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定"
      Height          =   320
      Left            =   3720
      TabIndex        =   3
      Top             =   1260
      Width           =   1100
   End
   Begin VB.Label lblHomePage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "意见反馈：totop@163.com"
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   8160
      Width           =   2070
   End
   Begin VB.Image ImageZFB 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   7500
      Left            =   12360
      Picture         =   "FrmOptions.frx":0CCA
      Stretch         =   -1  'True
      Top             =   240
      Width           =   4950
   End
   Begin VB.Image ImageWX 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   7500
      Left            =   6600
      Picture         =   "FrmOptions.frx":2E1B3
      Stretch         =   -1  'True
      Top             =   240
      Width           =   5475
   End
End
Attribute VB_Name = "FrmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'sql日志文件的保存目录
Private mSqlPath As String

Public Property Get SqlPath() As String
    SqlPath = mSqlPath
End Property
Public Property Let SqlPath(ByVal pData As String)
    mSqlPath = pData
End Property

Private Sub CmdBrowser_Click()
    On Error GoTo Errhandle
    
    FolderBrowser1.ShowFolderBrowser
    If FolderBrowser1.Folder <> "" Then
        TxtPath.Text = FolderBrowser1.Folder
    End If

    Exit Sub
    
Errhandle:
       
    FrmFolderBrowser.Show vbModal
    
End Sub

Private Sub cmdDonate_Click()
    If InStr(1, cmdDonate.Caption, ">") > 0 Then
        Me.Width = 17580
        cmdDonate.Caption = "<<  不想赞助"
    Else
        Me.Width = 6465
        cmdDonate.Caption = "赞助作者1分钱  >>"
    End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Dim pLocalFso As New Scripting.FileSystemObject
    Dim i As Long
    Dim pTempStr As String
    Dim pErrorstring As String
    
    If Len(Trim(TxtPath.Text)) = 0 Then
        MsgBox "目录非法为空." & Space(10)
        Exit Sub
    ElseIf pLocalFso.FolderExists(Trim(TxtPath.Text)) = True Then
        FrmShowData.SqlPath = Trim(TxtPath.Text)
    Else
        i = MsgBox("目录不存在: " & Trim(TxtPath.Text) & "，是否创建该目录？", vbYesNo + vbExclamation, "提示")
        If i = vbYes Then
            If CreateLocalDirectory(Trim(TxtPath.Text), pTempStr, pErrorstring) = True Then
                
                TxtPath.Text = pTempStr
            Else
                MsgBox pErrorstring, vbOKOnly + vbExclamation, "提示"
                Exit Sub
            End If
        Else
            Exit Sub
        End If
        
    End If
    
    '写入配置文件中
    Call WritePrivateProfileString("SQL", "SaveFolder", TxtPath.Text, App.Path & "\Configure.ini")
    
    Set pLocalFso = Nothing
    
    Unload Me
End Sub

'功能说明：创建目录（FSO 只能创建一级目录，该函数采用递归算法可创建任意多级目录！）
Private Function CreateLocalDirectory(ByVal TheLocalDirectoryToBeCreated As String, ByRef TheStandardDirectoryAfterCreated As String, Optional ByRef Errorstring As String) As Boolean
    Dim Tempstr1 As String
    Dim Tempstr2 As String
    Dim Flag As Boolean
    Dim i As Long
    Dim LocalFso As Scripting.FileSystemObject
    
    On Error GoTo CreateErr
    Flag = True
    Set LocalFso = New Scripting.FileSystemObject
    TheStandardDirectoryAfterCreated = Trim(Replace(TheLocalDirectoryToBeCreated, "/", "\"))
    If LocalFso.FolderExists(TheStandardDirectoryAfterCreated) = False Then
        i = InStrRev(TheStandardDirectoryAfterCreated, "\", Len(TheStandardDirectoryAfterCreated))
        If i <> 0 And i > 3 Then
            Tempstr1 = Left(TheStandardDirectoryAfterCreated, i - 1)
            If CreateLocalDirectory(Tempstr1, Tempstr2, Errorstring) = False Then
                Flag = False
            End If
        End If
    End If
    
    If Flag = True Then
        If LocalFso.FolderExists(TheStandardDirectoryAfterCreated) = False Then
            TheStandardDirectoryAfterCreated = LocalFso.CreateFolder(TheStandardDirectoryAfterCreated).Path
        End If
    End If

CreateOK:
    CreateLocalDirectory = Flag
    Set LocalFso = Nothing
    If Flag = False Then
        Errorstring = "Function.CreateLocalDirectory(" & TheStandardDirectoryAfterCreated & ")--->" & Errorstring
    End If
    Exit Function
    
CreateErr:
    With Err
        Errorstring = "Err.Description:" & .Description & ",   Err.Source:" & .Source
        Flag = False
    End With
    Resume CreateOK
End Function


Private Sub CmdOpen_Click()
    Dim LocalFso As New Scripting.FileSystemObject
    
    If LocalFso.FolderExists(Trim(TxtPath.Text)) Then
        Call Shell("explorer.exe " & Trim(TxtPath.Text), vbNormalFocus)
    Else
        MsgBox "目录不存在: " & Trim(TxtPath.Text), vbOKOnly + vbExclamation, "提示"
    End If
    
    Set LocalFso = Nothing
End Sub

Private Sub Form_Load()
    
    '按钮皮肤初始化
    If gConApplayCmdSkin = 1 Then
        BSE1.SchemeStyle = gConCmdSkinType
        BSE1.EndSubClassing
        BSE1.InitSubClassing
    End If
    
    TxtPath = mSqlPath

    txtDescription.Text = "轻量级数据库查询客户端功能简介：" & vbCrLf & vbCrLf
    txtDescription.Text = txtDescription.Text & "***********************************************" & vbCrLf & vbCrLf
    txtDescription.Text = txtDescription.Text & "1、支持SQLServer、Oracle、Sybase、DB2数据库平台" & vbCrLf
    txtDescription.Text = txtDescription.Text & "（1）支持DB2，要求安装 DB2 Run-Time Client Lite" & vbCrLf
    txtDescription.Text = txtDescription.Text & "（2）支持Oracle，要求安装ODAC组件（OLEDB）" & vbCrLf
    txtDescription.Text = txtDescription.Text & "（3）支持Sybase，本软件已安装Sybase OLEDB组件" & vbCrLf
    txtDescription.Text = txtDescription.Text & "（4）支持SQL Server，本软件已安装ADO组件" & vbCrLf
    txtDescription.Text = txtDescription.Text & "2、单句执行：可以提交单一的SQL语句给服务器执行。" & vbCrLf
    txtDescription.Text = txtDescription.Text & "3、多句一次执行：可以将多个SQL语句用空格或换行的方式进行隔离，程序会一次性提交给数据库服务器执行。" & vbCrLf
    txtDescription.Text = txtDescription.Text & "4、多句依次执行：可以将多个SQL语句用分号‘;’隔开，程序会逐句提交。" & vbCrLf
    txtDescription.Text = txtDescription.Text & "5、执行指定的SQL语句：可以在SQL编辑框中选中部分语句执行，在SQL命令框中通过鼠标连续单击3次可以选择当前行。" & vbCrLf
    txtDescription.Text = txtDescription.Text & "6、在事务中执行SQL语句：将执行按钮上方的Trans 框选中，那么程序对于每次提交的SQL语句都是放在事务中执行的，执行不成功程序会进行回滚操作。程序默认不进行事务处理，因为有些语句不能在事务中执行。" & vbCrLf & vbCrLf
    txtDescription.Text = txtDescription.Text & "**********************************************" & vbCrLf & vbCrLf & "主界面快捷键说明：" & vbCrLf & "F1：显示帮助" & vbCrLf & "F2：撤销对SQL命令框中内容的最近一次编辑" & vbCrLf & "F3：取消对SQL命令框中内容执行的撤销操作" & vbCrLf & "F5：执行SQL命令" & vbCrLf & "F9：显示所选数据库的所有表对象" & vbCrLf & "F10：调整SQL命令框的大小至最大，或还原至正常大小" & vbCrLf

End Sub
