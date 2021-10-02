VERSION 5.00
Begin VB.Form FrmSql 
   Caption         =   "DisplayData"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10830
   Icon            =   "FrmSql.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   10830
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.TextBox TxtSQL 
      Height          =   5505
      HideSelection   =   0   'False
      Left            =   300
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   960
      Width           =   9735
   End
End
Attribute VB_Name = "FrmSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF10 Then
        FrmShowData.TxtSQL.Text = TxtSQL.Text
        If FrmShowData.WindowState = vbNormal Then
            FrmShowData.Move Me.Left, Me.Top, Me.Width, Me.Height
        End If
        Me.Hide
        FrmShowData.Show
        KeyCode = 0
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = "DisplayData " & App.Major & "." & App.Minor & "." & App.Revision
    TxtSQL.ToolTipText = "SQL Commands"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    FrmShowData.TxtSQL.Text = TxtSQL.Text
    If FrmShowData.WindowState = vbNormal Then
        FrmShowData.Move Me.Left, Me.Top, Me.Width, Me.Height
    End If
    FrmShowData.Show
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        TxtSQL.Move 0, 0, Me.Width, Me.Height
    End If
End Sub

