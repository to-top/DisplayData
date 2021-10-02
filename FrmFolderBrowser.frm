VERSION 5.00
Begin VB.Form FrmFolderBrowser 
   Caption         =   "浏览目录"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4125
   Icon            =   "FrmFolderBrowser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   4125
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdExit 
      Caption         =   "取消"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   4260
      Width           =   1215
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   4260
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2250
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   3885
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   105
      TabIndex        =   1
      Top             =   660
      Width           =   3915
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "FrmFolderBrowser.frx":08CA
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "FrmFolderBrowser.frx":1194
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "子目录:"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "驱动器:"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   885
   End
End
Attribute VB_Name = "FrmFolderBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mFolder As String


Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdOK_Click()
    mFolder = Dir1.Path

    FrmOptions.TxtPath.Text = mFolder

    Unload Me
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Resize()
    If Me.Width - 330 > 0 Then
        Drive1.Width = Me.Width - 330
        Dir1.Width = Me.Width - 330
    End If
    
    If Me.Height - 3030 > 0 Then
        Dir1.Height = Me.Height - 3030
    End If
    
    If Me.Width - 2805 > 0 Then
        CmdOK.Left = Me.Width - 2805
    End If
    
    If Me.Width - 1485 > 0 Then
        CmdExit.Left = Me.Width - 1485
    End If
    
    If Me.Height - 1020 > 0 Then
        CmdOK.Top = Me.Height - 1020
        CmdExit.Top = Me.Height - 1020
    End If
End Sub


