VERSION 5.00
Object = "{C7D9E622-6CAA-42B3-9460-0E9DA4FD385A}#1.0#0"; "FBSLib.ocx"
Object = "{9A226D6F-2658-4445-8D35-5C19D42676FE}#1.0#0"; "BSE.ocx"
Begin VB.Form FrmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��־�ļ�����..."
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
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtDescription 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��������1��Ǯ  >>"
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
         Caption         =   "���..."
         Height          =   320
         Left            =   5130
         TabIndex        =   1
         Top             =   570
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "SQL��־�ļ��ı���Ŀ¼(�ļ���ΪSQL.txt):"
         Height          =   240
         Left            =   180
         TabIndex        =   6
         Top             =   270
         Width           =   4200
      End
   End
   Begin VB.CommandButton CmdOpen 
      Caption         =   "�򿪸�Ŀ¼(&O)"
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
      Caption         =   "ȡ��"
      Height          =   320
      Left            =   5145
      TabIndex        =   4
      Top             =   1260
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "ȷ��"
      Height          =   320
      Left            =   3720
      TabIndex        =   3
      Top             =   1260
      Width           =   1100
   End
   Begin VB.Label lblHomePage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���������totop@163.com"
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

'sql��־�ļ��ı���Ŀ¼
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
        cmdDonate.Caption = "<<  ��������"
    Else
        Me.Width = 6465
        cmdDonate.Caption = "��������1��Ǯ  >>"
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
        MsgBox "Ŀ¼�Ƿ�Ϊ��." & Space(10)
        Exit Sub
    ElseIf pLocalFso.FolderExists(Trim(TxtPath.Text)) = True Then
        FrmShowData.SqlPath = Trim(TxtPath.Text)
    Else
        i = MsgBox("Ŀ¼������: " & Trim(TxtPath.Text) & "���Ƿ񴴽���Ŀ¼��", vbYesNo + vbExclamation, "��ʾ")
        If i = vbYes Then
            If CreateLocalDirectory(Trim(TxtPath.Text), pTempStr, pErrorstring) = True Then
                
                TxtPath.Text = pTempStr
            Else
                MsgBox pErrorstring, vbOKOnly + vbExclamation, "��ʾ"
                Exit Sub
            End If
        Else
            Exit Sub
        End If
        
    End If
    
    'д�������ļ���
    Call WritePrivateProfileString("SQL", "SaveFolder", TxtPath.Text, App.Path & "\Configure.ini")
    
    Set pLocalFso = Nothing
    
    Unload Me
End Sub

'����˵��������Ŀ¼��FSO ֻ�ܴ���һ��Ŀ¼���ú������õݹ��㷨�ɴ�������༶Ŀ¼����
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
        MsgBox "Ŀ¼������: " & Trim(TxtPath.Text), vbOKOnly + vbExclamation, "��ʾ"
    End If
    
    Set LocalFso = Nothing
End Sub

Private Sub Form_Load()
    
    '��ťƤ����ʼ��
    If gConApplayCmdSkin = 1 Then
        BSE1.SchemeStyle = gConCmdSkinType
        BSE1.EndSubClassing
        BSE1.InitSubClassing
    End If
    
    TxtPath = mSqlPath

    txtDescription.Text = "���������ݿ��ѯ�ͻ��˹��ܼ�飺" & vbCrLf & vbCrLf
    txtDescription.Text = txtDescription.Text & "***********************************************" & vbCrLf & vbCrLf
    txtDescription.Text = txtDescription.Text & "1��֧��SQLServer��Oracle��Sybase��DB2���ݿ�ƽ̨" & vbCrLf
    txtDescription.Text = txtDescription.Text & "��1��֧��DB2��Ҫ��װ DB2 Run-Time Client Lite" & vbCrLf
    txtDescription.Text = txtDescription.Text & "��2��֧��Oracle��Ҫ��װODAC�����OLEDB��" & vbCrLf
    txtDescription.Text = txtDescription.Text & "��3��֧��Sybase��������Ѱ�װSybase OLEDB���" & vbCrLf
    txtDescription.Text = txtDescription.Text & "��4��֧��SQL Server��������Ѱ�װADO���" & vbCrLf
    txtDescription.Text = txtDescription.Text & "2������ִ�У������ύ��һ��SQL����������ִ�С�" & vbCrLf
    txtDescription.Text = txtDescription.Text & "3�����һ��ִ�У����Խ����SQL����ÿո���еķ�ʽ���и��룬�����һ�����ύ�����ݿ������ִ�С�" & vbCrLf
    txtDescription.Text = txtDescription.Text & "4���������ִ�У����Խ����SQL����÷ֺš�;�����������������ύ��" & vbCrLf
    txtDescription.Text = txtDescription.Text & "5��ִ��ָ����SQL��䣺������SQL�༭����ѡ�в������ִ�У���SQL�������ͨ�������������3�ο���ѡ��ǰ�С�" & vbCrLf
    txtDescription.Text = txtDescription.Text & "6����������ִ��SQL��䣺��ִ�а�ť�Ϸ���Trans ��ѡ�У���ô�������ÿ���ύ��SQL��䶼�Ƿ���������ִ�еģ�ִ�в��ɹ��������лع�����������Ĭ�ϲ�������������Ϊ��Щ��䲻����������ִ�С�" & vbCrLf & vbCrLf
    txtDescription.Text = txtDescription.Text & "**********************************************" & vbCrLf & vbCrLf & "�������ݼ�˵����" & vbCrLf & "F1����ʾ����" & vbCrLf & "F2��������SQL����������ݵ����һ�α༭" & vbCrLf & "F3��ȡ����SQL�����������ִ�еĳ�������" & vbCrLf & "F5��ִ��SQL����" & vbCrLf & "F9����ʾ��ѡ���ݿ�����б����" & vbCrLf & "F10������SQL�����Ĵ�С����󣬻�ԭ��������С" & vbCrLf

End Sub
