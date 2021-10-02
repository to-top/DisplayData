VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{9A226D6F-2658-4445-8D35-5C19D42676FE}#1.0#0"; "BSE.ocx"
Begin VB.Form FrmShowData 
   Caption         =   "DisplayData"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11640
   Icon            =   "FrmShowData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   11640
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox TxtSQL 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00404040&
      Height          =   2775
      HideSelection   =   0   'False
      Left            =   4440
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      ToolTipText     =   "SQL Commands"
      Top             =   480
      Width           =   5775
   End
   Begin BSE_Engine.BSE BSE1 
      Left            =   150
      Top             =   6600
      _ExtentX        =   6588
      _ExtentY        =   1085
   End
   Begin VB.Frame FrameLogin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2970
      Left            =   690
      TabIndex        =   13
      Top             =   360
      Width           =   3555
      Begin VB.TextBox TxtSID 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   2400
         TabIndex        =   2
         Top             =   570
         Width           =   1065
      End
      Begin VB.CommandButton CmdExcel 
         Height          =   500
         Left            =   1710
         Picture         =   "FrmShowData.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "将查询结果集导出为Excel..."
         Top             =   2400
         Width           =   580
      End
      Begin VB.TextBox txtServer 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   1290
         TabIndex        =   1
         Top             =   570
         Width           =   1065
      End
      Begin VB.TextBox txtUID 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   1290
         TabIndex        =   4
         Top             =   1290
         Width           =   2200
      End
      Begin VB.TextBox txtPWD 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1290
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1650
         Width           =   2200
      End
      Begin VB.ComboBox CBOtype 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         ItemData        =   "FrmShowData.frx":1194
         Left            =   1290
         List            =   "FrmShowData.frx":11A4
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   180
         Width           =   2200
      End
      Begin VB.TextBox txtPort 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   1290
         TabIndex        =   3
         Top             =   930
         Width           =   2200
      End
      Begin VB.ComboBox CBOdatabases 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         ItemData        =   "FrmShowData.frx":11CD
         Left            =   1290
         List            =   "FrmShowData.frx":11CF
         TabIndex        =   6
         Text            =   "CBOdatabases"
         Top             =   2010
         Width           =   2205
      End
      Begin VB.CommandButton CmdClear 
         Height          =   500
         Left            =   1125
         Picture         =   "FrmShowData.frx":11D1
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "清除数据项"
         Top             =   2400
         Width           =   580
      End
      Begin VB.CommandButton CmdGo 
         Height          =   500
         Left            =   2880
         Picture         =   "FrmShowData.frx":1A9B
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "执行SQL语句"
         Top             =   2400
         Width           =   580
      End
      Begin VB.CommandButton CmdTranGo 
         Height          =   500
         Left            =   2295
         Picture         =   "FrmShowData.frx":2365
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "事务中执行SQL"
         Top             =   2400
         Width           =   580
      End
      Begin VB.Label LblServer 
         Caption         =   "Server:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   19
         Top             =   660
         Width           =   1200
      End
      Begin VB.Label Label2 
         Caption         =   "User:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   18
         Top             =   1350
         Width           =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   17
         Top             =   1740
         Width           =   1200
      End
      Begin VB.Label LblDatabase 
         Caption         =   "Database:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   16
         Top             =   2100
         Width           =   1200
      End
      Begin VB.Label Label5 
         Caption         =   "ServerType:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   15
         Top             =   270
         Width           =   1200
      End
      Begin VB.Label Label6 
         Caption         =   "ServerPort:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   14
         Top             =   990
         Width           =   1650
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   90
         MouseIcon       =   "FrmShowData.frx":2C2F
         MousePointer    =   99  'Custom
         Picture         =   "FrmShowData.frx":34F9
         Stretch         =   -1  'True
         ToolTipText     =   "日志文件设置..."
         Top             =   2400
         Width           =   450
      End
   End
   Begin TabDlg.SSTab SSTabResults 
      Height          =   4005
      Left            =   690
      TabIndex        =   12
      Top             =   3360
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   7064
      _Version        =   393216
      Style           =   1
      Tabs            =   11
      TabsPerRow      =   12
      TabHeight       =   520
      TabCaption(0)   =   " 1  "
      TabPicture(0)   =   "FrmShowData.frx":3DC3
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "GridData(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   " 2  "
      TabPicture(1)   =   "FrmShowData.frx":3DDF
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GridData(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   " 3  "
      TabPicture(2)   =   "FrmShowData.frx":3DFB
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "GridData(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   " 4  "
      TabPicture(3)   =   "FrmShowData.frx":3E17
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "GridData(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   " 5  "
      TabPicture(4)   =   "FrmShowData.frx":3E33
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "GridData(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   " 6  "
      TabPicture(5)   =   "FrmShowData.frx":3E4F
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "GridData(5)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   " 7  "
      TabPicture(6)   =   "FrmShowData.frx":3E6B
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "GridData(6)"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   " 8  "
      TabPicture(7)   =   "FrmShowData.frx":3E87
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "GridData(7)"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   " 9  "
      TabPicture(8)   =   "FrmShowData.frx":3EA3
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "GridData(8)"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   " 10 "
      TabPicture(9)   =   "FrmShowData.frx":3EBF
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "GridData(9)"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "Output"
      TabPicture(10)  =   "FrmShowData.frx":3EDB
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "PRG"
      Tab(10).Control(1)=   "TxtInformation"
      Tab(10).ControlCount=   2
      Begin MSDataGridLib.DataGrid GridData 
         Height          =   3660
         Index           =   0
         Left            =   30
         TabIndex        =   21
         Top             =   320
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   6456
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   4210752
         HeadLines       =   1
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin RichTextLib.RichTextBox TxtInformation 
         Height          =   3660
         Left            =   -74970
         TabIndex        =   20
         ToolTipText     =   "SQL语句的执行情况"
         Top             =   330
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   6456
         _Version        =   393217
         BackColor       =   16777215
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"FrmShowData.frx":3EF7
      End
      Begin MSDataGridLib.DataGrid GridData 
         Height          =   3660
         Index           =   1
         Left            =   -74970
         TabIndex        =   22
         ToolTipText     =   "查询结果"
         Top             =   320
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   6456
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   4210752
         HeadLines       =   1
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid GridData 
         Height          =   3660
         Index           =   2
         Left            =   -74970
         TabIndex        =   23
         ToolTipText     =   "查询结果"
         Top             =   320
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   6456
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   4210752
         HeadLines       =   1
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid GridData 
         Height          =   3660
         Index           =   3
         Left            =   -74970
         TabIndex        =   24
         ToolTipText     =   "查询结果"
         Top             =   320
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   6456
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   4210752
         HeadLines       =   1
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid GridData 
         Height          =   3660
         Index           =   4
         Left            =   -74970
         TabIndex        =   25
         ToolTipText     =   "查询结果"
         Top             =   320
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   6456
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   4210752
         HeadLines       =   1
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid GridData 
         Height          =   3660
         Index           =   5
         Left            =   -74970
         TabIndex        =   26
         ToolTipText     =   "查询结果"
         Top             =   320
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   6456
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   4210752
         HeadLines       =   1
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid GridData 
         Height          =   3660
         Index           =   6
         Left            =   -74970
         TabIndex        =   27
         ToolTipText     =   "查询结果"
         Top             =   320
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   6456
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   4210752
         HeadLines       =   1
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid GridData 
         Height          =   3660
         Index           =   7
         Left            =   -74970
         TabIndex        =   28
         ToolTipText     =   "查询结果"
         Top             =   320
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   6456
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   4210752
         HeadLines       =   1
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid GridData 
         Height          =   3660
         Index           =   8
         Left            =   -74970
         TabIndex        =   29
         ToolTipText     =   "查询结果"
         Top             =   320
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   6456
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   4210752
         HeadLines       =   1
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid GridData 
         Height          =   3660
         Index           =   9
         Left            =   -74970
         TabIndex        =   30
         ToolTipText     =   "查询结果"
         Top             =   320
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   6456
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   4210752
         HeadLines       =   1
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   3
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ProgressBar PRG 
         Height          =   225
         Left            =   -68500
         TabIndex        =   31
         Top             =   45
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "FrmShowData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLocalRecordset As ADODB.Recordset
Private mFrmHeight As Long
Private mFrmWidth As Long
Private mSQLarray(100) As String
Private mSQLQXCXarray(100) As String

Private mCurrentCXIndex As Long
Private mCurrentCXMaxIndex As Long

Private mCurrentQXCXMaxIndex As Long
Private mCurrentQXCXIndex As Long

'sql命令框的原始尺寸
Private mTop As Long
Private mLeft As Long
Private mHeight As Long
Private mWidth As Long
Private mSQL_CurrentIsFullMode As Boolean  'sql命令框当前是完全模式

Private mDatabasetype As Integer

Private mServer As String
Private mSID As String
Private mPort As String
Private mDatabase As String
Private mUser As String
Private mPassword As String

'SQL文件的保存路径
Private mSqlPath As String

Private mBorderHeight As Long  '系统caption 的高度
Private mCaptionHeight As Long '系统Border 的高度

'是否启用了消息钩子
Private mHookEnabled As Boolean

'当前显示的DBGRID
Private mCurrentDBGrid As DataGrid

'当前SQL命令框是否最大化
Private mSQLMAX As Boolean

'是否第一次load
Private mFirstLoad As Boolean

'SQL文件的保存路径
Public Property Let SqlPath(ByVal pData As String)
    mSqlPath = pData
End Property

Public Property Get CurrentDBGrid() As DataGrid
    Set CurrentDBGrid = mCurrentDBGrid
End Property

Private Sub CBOdatabases_Change()
    mDatabase = Trim(CBOdatabases.Text)
    CBOdatabases.ToolTipText = mDatabase
    Call SetCaption
End Sub

Private Sub CBOdatabases_Click()
    mDatabase = Trim(CBOdatabases.Text)
    CBOdatabases.ToolTipText = mDatabase
    Call SetCaption
End Sub

Private Sub CBOdatabases_dropdown()
    Dim pExists As Boolean '先前选择的数据库是否存在于列表中
    Dim Errorstring As String
    Dim Flag As Boolean
'    Dim LocalSql As String
    Dim LocalConnection As ADODB.Connection
    Dim LocalRS As ADODB.Recordset
    Dim mDatabasetype As Long
    
    
    On Error GoTo ErrorHandleOK
    Me.MousePointer = vbHourglass
    CmdGo.Enabled = False
    CmdTranGo.Enabled = False
    CmdClear.Enabled = False
    CmdExcel.Enabled = False
    DoEvents
    Flag = True
    
    If InStr(1, LCase(CBOtype.Text), "sql") > 0 Then
        mDatabasetype = 0
    End If
    If InStr(1, LCase(CBOtype.Text), "sybase") > 0 Then
        mDatabasetype = 1
    End If
    If InStr(1, LCase(CBOtype.Text), "oracle") > 0 Then
        mDatabasetype = 2
    End If
    If InStr(1, LCase(CBOtype.Text), "db2") > 0 Then
        mDatabasetype = 3
    End If
    
    If mDatabasetype <> 2 Then
        If Trim(txtServer) = "" Or Trim(txtUID) = "" Then
            Flag = False
            Errorstring = "信息不全：数据库服务和用户名不能为空！" & Space(5)
        End If
    Else
        If Trim(txtServer) = "" Or Trim(TxtSID) = "" Or Trim(txtUID) = "" Then
            Flag = False
            Errorstring = "信息不全：数据库服务的主机HOST(IP)、SID和用户名不能为空！" & Space(5)
        End If
    End If
    
    If Flag = True And Trim(txtPort) <> "" Then
        If IsNumeric(Trim(txtPort)) = False Then
            Flag = False
            Errorstring = "数据库服务器侦听的端口号(Server Port)必须是正整数！"
            txtPort.SetFocus
            txtPort.SelStart = 0
            txtPort.SelLength = Len(txtPort)
        Else
            If InStr(1, Trim(txtPort), ".") > 0 Or Val(Trim(txtPort)) <= 0 Then
                    Flag = False
                    Errorstring = "数据库服务器侦听的端口号(Server Port)必须是正整数！"
                    txtPort.SetFocus
                    txtPort.SelStart = 0
                    txtPort.SelLength = Len(txtPort)
            End If
        End If
    End If

    If Flag = True Then
        'db2暂不支持获取所有数据库名
        If mDatabasetype = 3 Then
            GoTo DatabaseOK
        Else
        
            If GetADOConnection(mDatabasetype, Trim(txtServer), Trim(TxtSID), Trim(txtUID), Trim(txtPWD), LocalConnection, "master", Trim(txtPort), Errorstring) = False Then
                Flag = False
            End If
        End If
    End If
    
    If Flag = True Then
        CBOdatabases.Clear
        CBOdatabases.Refresh
        If mDatabasetype = 0 Or mDatabasetype = 1 Then
            Set LocalRS = LocalConnection.Execute("sp_databases")
        Else
            'Set LocalRS = LocalConnection.Execute("select distinct tablespace_name from dba_free_space")
            Set LocalRS = LocalConnection.Execute("select distinct owner as schema from dba_catalog")
        End If
        If Not (LocalRS.EOF And LocalRS.BOF) Then
            While LocalRS.EOF = False
                If Trim(LocalRS.Fields(0) & "") <> "" Then
                    CBOdatabases.AddItem Trim(LocalRS.Fields(0) & "")
                    If mDatabase = Trim(LocalRS.Fields(0) & "") Then
                        pExists = True
                    End If
                    LocalRS.MoveNext
                End If
            Wend
            
            If pExists = True Then
                CBOdatabases.Text = mDatabase
            End If
        End If
    End If
    
DatabaseOK:
    Set LocalRS = Nothing
    Set LocalConnection = Nothing
    CmdGo.Enabled = True
    CmdTranGo.Enabled = True
    CmdClear.Enabled = True
    CmdExcel.Enabled = True
    Me.MousePointer = vbDefault
    If Flag = False Then
        MsgBox Errorstring, vbOKOnly + vbCritical, "错误"
    End If
    Exit Sub
    
ErrorHandleOK:
    Flag = False
    Errorstring = "Err.Number:" & Err.Number & Space(3) & "Err.Source:" & Err.Source & Space(3) & "Err.Description:" & Err.Description
    Resume DatabaseOK
End Sub



Private Sub CBOtype_Click()
    If InStr(1, LCase(CBOtype.Text), "sql") > 0 Then
        mDatabasetype = 0
    End If
    If InStr(1, LCase(CBOtype.Text), "sybase") > 0 Then
        mDatabasetype = 1
    End If
    If InStr(1, LCase(CBOtype.Text), "oracle") > 0 Then
        mDatabasetype = 2
    End If
    If InStr(1, LCase(CBOtype.Text), "db2") > 0 Then
        mDatabasetype = 3
    End If
        
    If InStr(1, LCase(CBOtype.Text), "oracle") > 0 Then
'        CBOdatabases.Text = ""
'        CBOdatabases.ForeColor = vbWhite
'        CBOdatabases.BackColor = &H80000003
'        CBOdatabases.Enabled = False
        LblServer.Caption = "HOST && SID:"
        LblDatabase.Caption = "Schemas:"
        txtServer.Width = 1000
        TxtSID.Visible = True
    Else
'        CBOdatabases.ForeColor = vbBlack
'        CBOdatabases.BackColor = vbWhite
'        CBOdatabases.Enabled = True
        LblServer.Caption = "Server:"
        LblDatabase.Caption = "Databases:"
        txtServer.Width = 2200
        TxtSID.Visible = False
    End If
End Sub

Private Sub CmdClear_Click()

    Image1.Visible = True
'    Status.Visible = False

    Call AddToCXarray
        
    TxtSQL.Text = ""
    
    Call ClearData
End Sub

Private Sub ClearData()
    Dim i As Long
    

    
    TxtInformation = ""
    Set mLocalRecordset = Nothing
    For i = 0 To 9
        Set GridData(i).DataSource = mLocalRecordset
        If i = 0 Then
            GridData(i).ToolTipText = "Results"
        Else
            GridData(i).ToolTipText = ""
        End If
        GridData(i).Refresh
    Next i
    For i = 1 To 9
        SSTabResults.TabVisible(i) = False
    Next i
    DoEvents
End Sub

Private Sub ExecuteSql(Optional ByVal pIsTrans As Boolean = False)
    Dim i As Long
    Dim Flag As Boolean
    Dim ErrorPosition As Long
    Dim Errorstring As String
    Dim LocalSql As String
    Dim ConnectionIsOK As Boolean
    Dim LocalConnection As ADODB.Connection
    Dim LocalRS As ADODB.Recordset
    Dim SqlArray As Variant
    Dim SqlTempArray As Variant
    Dim BeginTime As Double
    Dim EndTime As Double
    Dim strSQL As String
    Dim k As Long
    Dim Played As Boolean
    Dim AffectedRows As Double
    
    '记录每条SQL的执行情况
    Dim InfoArray() As String
    
    '异常错误次数，累计3次退出恢复
    Dim pErrTimes As Long
    
    On Error GoTo ErrorHandleOK
    
    Me.MousePointer = vbHourglass
    Flag = True
    
    '存储当前的连接设置
    mServer = Trim(txtServer.Text)
    mPort = Trim(txtPort.Text)
    mSID = Trim(TxtSID.Text)
    mUser = Trim(txtUID.Text)
    mPassword = Trim(txtPWD.Text)
    
    
    '初始化为结果区域显示
    SSTabResults.Tab = 10
    
    If TxtSQL.Text = "" Then
        Flag = False
        Errorstring = ""
    End If
    
    If Flag = True Then
        
        ErrorPosition = 0
        
        ConnectionIsOK = False
        CmdGo.Enabled = False
        CmdTranGo.Enabled = False
        CmdClear.Enabled = False
        CmdExcel.Enabled = False
        
        ErrorPosition = 1
        Call ClearData
        

        ErrorPosition = 2
        
        If mDatabasetype <> 2 Then
            If mServer = "" Or mUser = "" Or mDatabase = "" Or Trim(TxtSQL) = "" Then
                Flag = False
                Errorstring = "Please check the integrality of the informations."
            End If
        Else
            If mServer = "" Or mSID = "" Or mUser = "" Or Trim(TxtSQL) = "" Then
                Flag = False
                Errorstring = "Please check the integrality of the informations."
            End If
        End If
    End If
    
    If Flag = True And mPort <> "" Then
        ErrorPosition = 3
        If IsNumeric(mPort) = False Then
            Flag = False
            Errorstring = "数据库服务器侦听的端口号(Server Port)必须是正整数！"
        Else
            If InStr(1, mPort, ".") > 0 Or Val(mPort) <= 0 Then
                    Flag = False
                    Errorstring = "数据库服务器侦听的端口号(Server Port)必须是正整数！"
            End If
        End If
    End If
    
    If Flag = True Then
        ErrorPosition = 4
        If GetADOConnection(mDatabasetype, mServer, mSID, mUser, mPassword, LocalConnection, mDatabase, mPort, Errorstring) = False Then
            Flag = False
        Else
            ConnectionIsOK = True
'            If mDatabasetype = 0 Then
'                LocalConnection.BeginTrans
'            End If
            If pIsTrans = True Then
                ErrorPosition = 5
                LocalConnection.BeginTrans
            End If
        End If
    End If
    
    If Flag = True Then
        If Trim(TxtSQL.SelText) <> "" Then
            strSQL = Trim(TxtSQL.SelText)
        Else
            strSQL = Trim(TxtSQL.Text)
        End If
        
        'V3.5.5--记录SQL提交日志文件头内容
        Call WriteSqlLogHeader
        
        ErrorPosition = 6
        
        
        'V3.5.5--使用 "; & vbcrlf" 作为SQL语句分开提交的分割符
        '替换所有的“，‘”
        LocalSql = Trim(Replace(Replace(strSQL, "，", ","), "‘", "'"))
        
        If Right(LocalSql, 3) = ";" & vbCrLf Then
            LocalSql = Left(LocalSql, Len(LocalSql) - 3)
        End If
        
        SqlTempArray = Split(LocalSql, ";" & vbCrLf)
        Call DealSqlArray(SqlTempArray, SqlArray)
        
        
        
        k = -1
        With PRG
            .Visible = True
            .Min = 0
            .Max = UBound(SqlArray) + 1
            .Value = 0
        End With
        
        ErrorPosition = 7
        
        ReDim InfoArray(UBound(SqlArray))
        Dim AffectedRows1 As Long
        For i = 0 To UBound(SqlArray)
            If i < 10 Then
                If Trim(SqlArray(i)) <> "" Then
                    k = k + 1
                    SSTabResults.TabVisible(k) = True


                    ErrorPosition = 8
                    BeginTime = Timer
                    
                    Set mLocalRecordset = LocalConnection.Execute(SqlArray(i), AffectedRows1)
                    ErrorPosition = 9
                    Set GridData(k).DataSource = mLocalRecordset

                    EndTime = Timer
                    
                    InfoArray(i) = "SQL(" & i + 1 & ")： Status: OK.  Elapsed time: " & Format(EndTime - BeginTime, "0.000") & " s "
                    
                    If mDatabasetype = 0 Or mDatabasetype = 1 Then
                        ErrorPosition = 14
                        LocalSql = " SELECT @@ROWCOUNT as AffectedRows "
                        Set LocalRS = LocalConnection.Execute(LocalSql)
                        ErrorPosition = 15
                        AffectedRows = LocalRS.Fields("AffectedRows")
                        InfoArray(i) = InfoArray(i) & " ,  Affected Rows: " & AffectedRows
                    Else
                        If Not (mLocalRecordset Is Nothing Or mLocalRecordset.State <> adStateOpen) Then
                            InfoArray(i) = InfoArray(i) & " ,  Affected Rows: " & mLocalRecordset.RecordCount
                        End If
                    End If
                    
                    
                    
                    
                    ErrorPosition = 10
                    
ExecuteOK:
                    GridData(k).Refresh
                    GridData(k).ToolTipText = Trim(SqlArray(i))
                    
                    'V3.5.5--记录SQL提交日志
                    Call WriteSqlLog(Trim(SqlArray(i)))
                    
                    ErrorPosition = 12
                    PRG.Value = i + 1
'                    DoEvents
                    ErrorPosition = 13
                End If
            End If
        Next i
        

        
        SSTabResults.Tab = 10
        With PRG
            .Visible = False
            .Min = 0
            .Max = 1
            .Value = 0
        End With

    End If
    
DisplayOk:
    If Flag = False Then
        
        '任务执行发生问题则回滚事务
        If pIsTrans = True And ConnectionIsOK = True Then
            LocalConnection.RollbackTrans
        End If
        
        If Errorstring <> "" Then
'            TxtInformation.ForeColor = vbRed
            TxtInformation.Text = Errorstring
            TxtInformation.SelStart = 0
            TxtInformation.SelLength = Len(Errorstring)
            TxtInformation.SelColor = vbRed
        End If
    Else
        '任务执行成功则回滚事务
        If pIsTrans = True And ConnectionIsOK = True Then
            LocalConnection.CommitTrans
        End If

    End If
    
    If Len(LocalSql) > 0 Then
        TxtInformation.Text = ""
        
        For i = 0 To UBound(InfoArray)
            TxtInformation.Text = TxtInformation.Text & InfoArray(i) & vbCrLf
        Next
        
        For i = 0 To UBound(InfoArray)
'
'            TxtInformation.Text = TxtInformation.Text & InfoArray(i) & vbCrLf
            TxtInformation.SelStart = InStr(1, TxtInformation.Text, InfoArray(i)) - 1
            TxtInformation.SelLength = Len(InfoArray(i))
            If InStr(1, InfoArray(i), "Status: OK.") > 0 Then
                TxtInformation.SelColor = &H404040
            Else
                TxtInformation.SelColor = vbRed
            End If

        Next
        
        TxtInformation.SelStart = 0
        TxtInformation.SelLength = 0
        TxtInformation.Refresh
    End If
    
    Set LocalConnection = Nothing

    CmdGo.Enabled = True
    CmdTranGo.Enabled = True
    CmdClear.Enabled = True
    CmdExcel.Enabled = True
    PRG.Visible = False
    
    If Flag = True Then
        SSTabResults.Tab = 0
    End If
    Me.MousePointer = vbDefault
    
    Exit Sub
    
ErrorHandleOK:
    Flag = False
    
    pErrTimes = pErrTimes + 1
    
    
    If ErrorPosition = 8 Then
        InfoArray(i) = "SQL(" & i + 1 & ")： Status: Failure!  Error" & Err.Number & "   " & Err.Description & "   " & Err.Source
        Resume ExecuteOK
    Else
        Errorstring = "Err.Number:" & Err.Number & "  Err.Description:" & Err.Description
    End If
    
    'V3.5.8--异常错误累计出现3次以上就退出该过程（可能进入的错误捕获死循环）
    If pErrTimes <= 3 Then
        Resume DisplayOk
    End If
    
    TxtInformation.Text = Err.Description
    TxtInformation.SelColor = vbRed
    
    Me.MousePointer = vbDefault
End Sub


'V3.5.5--处理SQL语句数组，剔除无效的元素
Private Sub DealSqlArray(ByRef pSqlArray As Variant, ByRef pReturnArray As Variant)
    Dim i As Long
    Dim pArrayElement As String
    Dim pElements As Long
    Dim pSqlTempArray() As String
    
    pElements = -1
    
    For i = 0 To UBound(pSqlArray)
        pArrayElement = Trim(pSqlArray(i))
        pArrayElement = Trim(Replace(Trim(Replace(Trim(Replace(Trim(Replace(pArrayElement, vbCrLf, " ")), ";" & vbCr, " ")), vbCr, " ")), vbLf, " "))
        '去掉末尾的分号
        pArrayElement = TrimCharacters(pArrayElement, ";")
        
        If pArrayElement <> "" Then
            'pArrayElement = TrimCharacters(pArrayElement, vbCrLf)
            
            
            '如果处理后的SQL串有效则添加到返回数组中
            If Len(pArrayElement) > 0 Then
                pElements = pElements + 1
                ReDim Preserve pSqlTempArray(pElements)
                pSqlTempArray(pElements) = pArrayElement
            End If
        End If
    Next
    
    '返回处理后的SQL数组
    pReturnArray = pSqlTempArray
    
End Sub


'导出查询结果到Excel
Private Sub CmdExcel_Click()

    '导出之前设置其他按钮不可用
    CmdGo.Enabled = False
    CmdTranGo.Enabled = False
    CmdClear.Enabled = False
    CmdExcel.Enabled = False
    
    Call ExportToExcel
    
    '导出之后设置其他按钮可用
    CmdGo.Enabled = True
    CmdTranGo.Enabled = True
    CmdClear.Enabled = True
    CmdExcel.Enabled = True
End Sub

Private Sub CmdGo_Click()
    Call ExecuteSql(False)
End Sub


Private Sub CmdTranGo_Click()
    Call ExecuteSql(True)
End Sub



Private Sub Form_Activate()
    If mFirstLoad = True Then
        mFirstLoad = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    If KeyCode = vbKeyF5 Then
        If ((mDatabasetype = 0 Or mDatabasetype = 1) And CBOdatabases.Text <> "") Or mDatabasetype = 2 Then
            If TxtSQL.Text <> "" Then
                Call ExecuteSql
                KeyCode = 0
            End If
        End If
    End If
    
    If KeyCode = vbKeyF9 Then
        If CBOdatabases.Text <> "" Then
            If mDatabasetype = 0 Or mDatabasetype = 1 Then
                TxtSQL = "exec sp_tables"
            Else
                TxtSQL = "SELECT * FROM ALL_CATALOG WhERE OWNER='" & CBOdatabases.Text & "' ORDER BY TABLE_TYPE,TABLE_NAME"
            End If
            
            Call ExecuteSql
            KeyCode = 0
        End If
    End If
    
    If KeyCode = vbKeyF1 Then
        TxtInformation.Text = "**************************************************************************" & vbCrLf & "快捷键说明：" & vbCrLf & "F2：撤销对SQL命令框中内容的最近一次编辑" & vbCrLf & "F3：取消对SQL命令框中内容执行的撤销操作" & vbCrLf & "F5：执行SQL命令" & vbCrLf & "F9：显示所选数据库的所有表对象" & vbCrLf & "F10：调整SQL命令框的大小至最大，或还原至正常大小" & vbCrLf & "**************************************************************************"
        TxtInformation.SelStart = 0
        TxtInformation.SelLength = Len(TxtInformation.Text)
        TxtInformation.SelColor = vbBlue
        SSTabResults.Tab = 10
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyF10 Then
        
        If mSQLMAX = False Then
            TxtSQL.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
            mSQLMAX = True
        Else
            mSQLMAX = False
            Call Form_Resize
            
        End If
        
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyF2 And mCurrentCXIndex >= 1 And mCurrentCXIndex <= 100 Then
        
        

        Call AddQXCXarray
        

        TxtSQL = mSQLarray(mCurrentCXIndex)

        
        mCurrentCXIndex = mCurrentCXIndex - 1
        If mCurrentCXIndex = 0 Then
            mCurrentCXMaxIndex = 0
        End If
        
        KeyCode = 0
        
    End If
    
    If KeyCode = vbKeyF3 And mCurrentQXCXIndex >= 1 And mCurrentQXCXIndex <= 100 Then
        

        TxtSQL = mSQLQXCXarray(mCurrentQXCXIndex)

        
        mCurrentQXCXIndex = mCurrentQXCXIndex - 1
        If mCurrentQXCXIndex = 0 Then
            mCurrentQXCXMaxIndex = 0
        End If
        
        KeyCode = 0
        
    End If
    
End Sub

'设置窗体caption属性
Private Sub SetCaption(Optional ByVal pCaptionStr As String)
    If mDatabasetype = 0 Or mDatabasetype = 1 Then
        Me.Caption = "DisplayData " & App.Major & "." & App.Minor & "." & App.Revision & Space(5) & "[" & CBOtype.Text & Space(2) & txtServer.Text & "." & CBOdatabases.Text & "]"
    Else
        Me.Caption = "DisplayData " & App.Major & "." & App.Minor & "." & App.Revision & Space(5) & "[" & CBOtype.Text & Space(2) & txtServer.Text & "." & TxtSID.Text & "." & CBOdatabases.Text & "]"
    End If
End Sub

Private Sub Form_Load()
    Dim LocalFont As New StdFont
    Dim X As Object
    Dim i As Long
    Dim DBType As String
    Dim pLocalFso As New Scripting.FileSystemObject
    
    '按钮皮肤初始化
    If gConApplayCmdSkin = 1 Then
        BSE1.SchemeStyle = gConCmdSkinType
        BSE1.EndSubClassing
        BSE1.InitSubClassing
    End If
    
    mFirstLoad = True
    
    mBorderHeight = GetSystemMetrics(SM_CYBORDER) * Screen.TwipsPerPixelX
    mCaptionHeight = GetSystemMetrics(SM_CYCAPTION) * Screen.TwipsPerPixelY
    
    Set mCurrentDBGrid = GridData(0)

    '初始化为结果区域显示
    SSTabResults.Tab = 10
    PRG.Visible = False
    
    '支持结果集列表的鼠标中键滚动功能
    Hook Me.hwnd
    
    '设置窗体caption属性
    Me.Caption = "DisplayData " & App.Major & "." & App.Minor & "." & App.Revision
    CBOdatabases.Text = ""
    
    Image1.Visible = True
    Image1.Stretch = True
    Image1.ToolTipText = "选项设置..."

    
    
    TxtSQL.ToolTipText = "SQL Commands"
'    TxtInformation.ToolTipText = "Informations"
    GridData(0).ToolTipText = "Results"
    
    For i = 1 To 9
        SSTabResults.TabVisible(i) = False
    Next i
    mFrmHeight = Val(GetConfigFileString(App.Path & "\Configure.ini", "Face", "Height"))
    mFrmWidth = Val(GetConfigFileString(App.Path & "\Configure.ini", "Face", "Width"))
    
    If mFrmHeight > 0 And mFrmWidth > 0 Then
        Me.Height = mFrmHeight
        Me.Width = mFrmWidth
    ElseIf mFrmHeight = 0 Or mFrmWidth = 0 Then
        Me.WindowState = vbMaximized
    Else
        mFrmHeight = 7000
        mFrmWidth = 8000
        Me.Height = mFrmHeight
        Me.Width = mFrmWidth
    End If
    
    If pLocalFso.FileExists(App.Path & "\Configure.ini") = True Then
        gFontName = Trim(GetConfigFileString(App.Path & "\Configure.ini", "Face", "FontName"))
        gFontSize = Trim(GetConfigFileString(App.Path & "\Configure.ini", "Face", "FontSize"))
        mServer = Trim(GetConfigFileString(App.Path & "\Configure.ini", "DatabaseServerInfo", "DatabaseServerName"))
        txtServer.Text = mServer
        mSID = Trim(GetConfigFileString(App.Path & "\Configure.ini", "DatabaseServerInfo", "OracleServerSID"))
        TxtSID.Text = mSID
        mPort = Trim(GetConfigFileString(App.Path & "\Configure.ini", "DatabaseServerInfo", "DatabaseServerPort"))
        txtPort.Text = mPort
        mUser = Trim(GetConfigFileString(App.Path & "\Configure.ini", "DatabaseServerInfo", "DatabaseUserName"))
        txtUID.Text = mUser
        mPassword = Trim(GetConfigFileString(App.Path & "\Configure.ini", "DatabaseServerInfo", "DatabasePassword"))
        txtPWD.Text = mPassword
        
        
        
        DBType = Trim(GetConfigFileString(App.Path & "\Configure.ini", "DatabaseServerInfo", "DatabaseServerType"))
        If IsNumeric(DBType) = False Then
            DBType = "0"
        End If
        mDatabasetype = CLng(DBType)
        
        'SQL 日志文件的保存路径，如果不存在则缺省保存在安装目录下SQL.txt
        mSqlPath = Trim(GetConfigFileString(App.Path & "\Configure.ini", "SQL", "SaveFolder"))
        If pLocalFso.FolderExists(mSqlPath) = False Then
            mSqlPath = App.Path
        End If
    Else
        gFontName = "Courier New"
        gFontSize = 9
        txtServer = ""
        txtPort = ""
        txtUID = ""
        txtPWD = ""
        DBType = "0"
    End If
    
    CBOtype.ListIndex = Val(mDatabasetype)
    
    CBOdatabases.Text = Trim(GetConfigFileString(App.Path & "\Configure.ini", "DatabaseServerInfo", "DatabaseName"))
    
    On Error GoTo DefaultFont
    LocalFont.Name = gFontName
    LocalFont.Size = gFontSize
    For Each X In Me.Controls
        If X.Name <> "BSE1" And X.Name <> "Status" And X.Name <> "PRG" And X.Name <> "Image1" And X.Name <> "UD" And X.Name <> "LR" And X.Name <> "TxtInformation" Then
          Set X.Font = LocalFont
        End If
    Next


LoadOk:
    Set LocalFont = Nothing

    Exit Sub
    
DefaultFont:
    MsgBox Err.Description, vbOKOnly + vbInformation, "提示"
    Resume LoadOk
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call WritePrivateProfileString("DatabaseServerInfo", "DatabaseServerType", CStr(CBOtype.ListIndex), App.Path & "\Configure.ini")
    Call WritePrivateProfileString("DatabaseServerInfo", "DatabaseServerName", CStr(Trim(txtServer.Text)), App.Path & "\Configure.ini")
    Call WritePrivateProfileString("DatabaseServerInfo", "OracleServerSID", CStr(Trim(TxtSID.Text)), App.Path & "\Configure.ini")
    Call WritePrivateProfileString("DatabaseServerInfo", "DatabaseServerPort", CStr(Trim(txtPort.Text)), App.Path & "\Configure.ini")
    Call WritePrivateProfileString("DatabaseServerInfo", "DatabaseUserName", CStr(Trim(txtUID.Text)), App.Path & "\Configure.ini")
    Call WritePrivateProfileString("DatabaseServerInfo", "DatabasePassword", CStr(Trim(txtPWD.Text)), App.Path & "\Configure.ini")
    Call WritePrivateProfileString("DatabaseServerInfo", "DatabaseName", CStr(Trim(CBOdatabases.Text)), App.Path & "\Configure.ini")
    
    '最大化情况下写入0
    If Me.WindowState = vbMaximized Then
        Call WritePrivateProfileString("Face", "Width", CStr(0), App.Path & "\Configure.ini")
        Call WritePrivateProfileString("Face", "Height", CStr(0), App.Path & "\Configure.ini")
    Else
        Call WritePrivateProfileString("Face", "Width", CStr(mFrmWidth), App.Path & "\Configure.ini")
        Call WritePrivateProfileString("Face", "Height", CStr(mFrmHeight), App.Path & "\Configure.ini")
    End If
    
    Set mCurrentDBGrid = Nothing
    
    UnHook Me.hwnd
    
End Sub


'V3.5.9  10:48 2009/8/19  --修订Bug：在Vista或Win7系统下主界面部分被遮盖显示不全的问题；
Private Sub Form_Resize()
    Dim i As Long

    If Me.WindowState <> vbMinimized Then
        '如果当前是sql命令框最大化，则resize移动txtsql
        If mSQLMAX = True Then
            TxtSQL.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        Else
            FrameLogin.Move 0, 0
            If Me.ScaleWidth - FrameLogin.Width > 0 Then
                TxtSQL.Move FrameLogin.Width + mBorderHeight, 0, Me.ScaleWidth - FrameLogin.Width - mBorderHeight, FrameLogin.Height
            End If
            
            If Me.ScaleHeight - TxtSQL.Height - mBorderHeight > 0 Then
                SSTabResults.Move 0, TxtSQL.Height + mBorderHeight, Me.ScaleWidth, Me.ScaleHeight - TxtSQL.Height - mBorderHeight
            End If
            
            '移动Grid
            If SSTabResults.Width - mBorderHeight * 6 > 0 And SSTabResults.Height - (SSTabResults.TabHeight + mBorderHeight * 6) > 0 Then
                For i = 0 To 9
        
                    GridData(i).Move mBorderHeight * 3, SSTabResults.TabHeight + mBorderHeight * 3, SSTabResults.Width - mBorderHeight * 6, SSTabResults.Height - (SSTabResults.TabHeight + mBorderHeight * 6)
                Next
            End If
            
            '移动信息提示框
            If SSTabResults.Width - mBorderHeight * 6 > 0 And SSTabResults.Height - (SSTabResults.TabHeight + mBorderHeight * 6) > 0 Then
                TxtInformation.Move mBorderHeight * 3, SSTabResults.TabHeight + mBorderHeight * 3, SSTabResults.Width - mBorderHeight * 6, SSTabResults.Height - (SSTabResults.TabHeight + mBorderHeight * 6)
            End If
            
            '移动进度条
            If SSTabResults.Width - 6500 - mBorderHeight > 0 And SSTabResults.TabHeight - mBorderHeight * 2 > 0 Then
                PRG.Move 6500, mBorderHeight, SSTabResults.Width - 6500 - mBorderHeight, SSTabResults.TabHeight - mBorderHeight * 2
            End If
        End If
        
        If mFirstLoad = False Then
            mFrmHeight = Me.Height
            mFrmWidth = Me.Width
        End If
    End If
    
End Sub


Private Sub Image1_Click()
    FrmOptions.SqlPath = mSqlPath
    FrmOptions.Show vbModal
End Sub

Private Sub SSTabResults_Click(PreviousTab As Integer)
    Dim pIndex As Long
    
    pIndex = SSTabResults.Tab
    If pIndex < SSTabResults.Tabs - 1 Then
        Set mCurrentDBGrid = GridData(pIndex)
        mCurrentDBGrid.ZOrder
    Else
        TxtInformation.ZOrder
    End If
    
End Sub

Private Sub txtPort_Change()
    mPort = Trim(txtPort.Text)
    txtPort.ToolTipText = mPort
End Sub

Private Sub txtPWD_Change()
    mPassword = Trim(txtPWD.Text)
End Sub

Private Sub txtServer_Change()
    mServer = Trim(txtServer.Text)
    txtServer.ToolTipText = mServer
    
    Call SetCaption
End Sub


Private Sub TxtSID_Change()
    mSID = Trim(TxtSID.Text)
    TxtSID.ToolTipText = mSID
    Call SetCaption
End Sub

Private Sub TxtSQL_Change()
'    If Right(TxtSQL.Text, Len(TxtSQL.Text) - InStrRev(TxtSQL.Text, " ")) = "select" Then
'        TxtSQL.SelStart = InStrRev(TxtSQL.Text, " ")
'        TxtSQL.SelLength = 6
'        TxtSQL.SelColor = vbBlue
'    End If
    'mSQLarray(0) = TxtSQL
End Sub

Private Sub TxtSQL_DblClick()
    Dim pBeginPos As Long
    Dim pEndPos As Long
    
    
    '定位当前行的首尾换行字符在文本框中的位置
    Call LocatePosition(TxtSQL.hwnd, pBeginPos, pEndPos)
    
    
    '选择的内容
    If pEndPos - pBeginPos > 0 Then
        TxtSQL.SelStart = pBeginPos
        TxtSQL.SelLength = pEndPos - pBeginPos
    End If
    
End Sub

Private Sub TxtSQL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Or KeyCode = vbKeyReturn Or TxtSQL.SelText <> "" Then
        Call AddToCXarray
    End If
End Sub

Private Sub MoveCXArrayPosition()
    Dim i As Long
    For i = 1 To 99
        mSQLarray(i) = mSQLarray(i + 1)
    Next i
End Sub

Private Sub MoveQXCXArrayPosition()
    Dim i As Long
    For i = 1 To 99
        mSQLQXCXarray(i) = mSQLQXCXarray(i + 1)
    Next i
End Sub

Private Sub AddToCXarray()
    mCurrentCXIndex = mCurrentCXMaxIndex + 1
    If mCurrentCXIndex > 100 Then
        mCurrentCXIndex = 100
        Call MoveCXArrayPosition
    End If
    mSQLarray(mCurrentCXIndex) = TxtSQL
    mCurrentCXMaxIndex = mCurrentCXIndex
End Sub

Private Sub AddQXCXarray()
    mCurrentQXCXIndex = mCurrentQXCXMaxIndex + 1
    If mCurrentQXCXIndex > 100 Then
        mCurrentQXCXIndex = 100
        Call MoveQXCXArrayPosition
    End If
    mSQLQXCXarray(mCurrentQXCXIndex) = TxtSQL
    mCurrentQXCXMaxIndex = mCurrentQXCXIndex
End Sub

Private Sub TxtSQL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 1 Then
        TxtSQL.SelStart = 0
        TxtSQL.SelLength = Len(TxtSQL)
        KeyAscii = 0
    End If
    
    If KeyAscii = 26 Then
        Call AddQXCXarray
    End If
    
End Sub


Private Sub txtUID_Change()
    mUser = Trim(txtUID.Text)
    txtUID.ToolTipText = mUser
End Sub

Private Sub UD_Resize()
    Dim i As Long
    
    '移动Grid
    For i = 0 To 9
        If SSTabResults.Width - 60 > 0 Then
            GridData(i).Width = SSTabResults.Width - 60
        End If
        If SSTabResults.Height - 800 > 0 Then
            GridData(i).Height = SSTabResults.Height - 800
        End If
    Next i
    
    '移动信息提示框
    If SSTabResults.Width - 60 > 0 Then
        TxtInformation.Width = SSTabResults.Width - 60
    End If
    If SSTabResults.Height - 800 > 0 Then
        TxtInformation.Height = SSTabResults.Height - 800
    End If
    
End Sub


'函数名：ExportToExcel
'功能：把记录集对象中的数据导出到Excel表单中
'Errorstring -----导出失败的可能原因描述
'返回值：True----导出成功、 False-----导出失败
Private Function ExportToExcel(Optional ByRef Errorstring As String) As Boolean
'    Dim xclApplication As Excel.Application
'    Dim xclBook As Excel.Workbook
'    Dim xclWorkSheet As Excel.Worksheet
'    Dim xclRange As Excel.Range
    Dim xclApplication As Object
    Dim xclBook As Object
    Dim xclWorkSheet As Object
    Dim xclRange As Object
    Dim RowIndex As Long
    Dim ColIndex As Long
    Dim i As Long
    Dim m As Long

    Dim Flag As Boolean
    Dim ErrorPosition As Long
    Dim pDataSource As Recordset
    Dim pSql As String  'SQL语句源
    Dim pFieldsCount As Long  '字段数
    
    Dim pIsLastOne As Boolean  '是最后一个结果集
    
    'On Error GoTo Errhandle
    Me.MousePointer = vbHourglass
    Flag = True
    
    '导出时从最后一个结果集开始，这样的目的是为了导出的excel文档的worksheet排列和结果集的顺序一致
    
    '导出的第一个worksheet(也就是最后一个结果集)要使用excel app创建时默认创建的worksheet
    pIsLastOne = True
    
    ErrorPosition = 0
    Set xclApplication = CreateObject("Excel.Application")
'    xclApplication.Visible = True

    '添加工作簿
    ErrorPosition = 1
    Set xclBook = xclApplication.Workbooks.Add
    
    While xclBook.Worksheets.Count > 1
        xclBook.Worksheets(1).Delete
    Wend
    
    '初始化为进度条
    SSTabResults.Tab = 10
    PRG.Min = 0
    PRG.Max = 10
    PRG.Value = 0
    PRG.Visible = True
    DoEvents
        
    ErrorPosition = 2
    '总共10个Grid，挨个检查数据源是否存在
    For m = 9 To 0 Step -1
        If Not (GridData(m).DataSource Is Nothing) Then
            Exit For
        End If
    Next

    '如果不存在任何数据源则给予提示
    If m = -1 Then
        Flag = False
        Errorstring = "没有需要导出的记录集." & Space(10)
        GoTo ExportOK
    End If
    
    '总共10个Grid，挨个导出
    For m = 9 To 0 Step -1
        
        '如果标签可见说明有数据结果，需要进行导出
        If SSTabResults.TabVisible(m) = True Then
            ErrorPosition = 3
            '获取数据源
            pSql = GridData(m).ToolTipText
            Set pDataSource = GridData(m).DataSource
            If pDataSource Is Nothing Or pDataSource.State <> adStateOpen Then
                GoTo ExcelNext
            End If
            
            ErrorPosition = 4
            pFieldsCount = pDataSource.Fields.Count
            TxtInformation.Text = TxtInformation.Text & "正在导出Excel(查询结果集" & m + 1 & ")： " & pSql & vbCrLf
            DoEvents
            
            ErrorPosition = 5
            '添加工作表
            If pIsLastOne = True Then
                '已经有个worksheet在视图中，无需添加，就用它作为第一个worksheet
                Set xclWorkSheet = xclBook.Worksheets(1)
                
                '用过之后，就将该标志置为false
                pIsLastOne = False
            Else
                Set xclWorkSheet = xclBook.Worksheets.Add
            End If
        
            '设置打印页面为横像(V3.5.3--容错没有打印机的情况：无打印机执行横向设置会报对象作用于方法失败的错误)
            ErrorPosition = 6
            'xclWorkSheet.PageSetup.Orientation = 2
        
            '设置工作表名称
            xclWorkSheet.Name = "查询结果集" & CStr(m + 1)
                    
            ErrorPosition = 7
            
            '设置工作表标题和其他台头信息
            If Flag = True Then
                With xclWorkSheet
                    ErrorPosition = 8
                    
                    '设置单元格宽度
                    .StandardWidth = 12
                    
                    '设置标题      '合并第一行的1－pFieldsCount格为一个单元格
                    .range(.Cells(1, 1), .Cells(1, pFieldsCount)).Merge (1)
    '                .Cells(1, 1).Font.Color = RGB(0, 128, 0)
    '                .cells(1, 1).Font.Bold = True
                    .Cells(1, 1).Font.Name = "黑体"
                    .Cells(1, 1).Font.Size = 12
                    .Cells(1, 1).VerticalAlignment = 3
                    .Cells(1, 1).HorizontalAlignment = 3
                    .Cells(1, 1).Value = "查询结果集" & CStr(m + 1)
                    
                    '设置记录数，SQL语句源小标题信息
                    
                    .range(.Cells(2, 1), .Cells(2, pFieldsCount)).Merge (1)
                    .Cells(2, 1).Font.Color = RGB(0, 128, 0)
    '                .Cells(2, 1).Font.Bold = True
                    .Cells(2, 1).Font.Name = "宋体"
                    .Cells(2, 1).Font.Size = 10
                    .Cells(2, 1).VerticalAlignment = 3
                    .Cells(2, 1).HorizontalAlignment = 1
                    .Cells(2, 1).Value = "记录数：" & pDataSource.RecordCount
                    
                    .range(.Cells(3, 1), .Cells(3, pFieldsCount)).Merge (1)
                    .Cells(3, 1).Font.Color = RGB(0, 128, 0)
    '                .Cells(3, 1).Font.Bold = True
                    .Cells(3, 1).Font.Name = "宋体"
                    .Cells(3, 1).Font.Size = 10
                    .Cells(3, 1).VerticalAlignment = 3
                    .Cells(3, 1).HorizontalAlignment = 1
                    .Cells(3, 1).Value = "SQL：" & pSql
                                   
                    ErrorPosition = 9
                    '设置列标题
                    For i = 1 To pFieldsCount
                        .Cells(4, i).Font.Color = vbBlue
                        .Cells(4, i).Font.Name = "宋体"
                        .Cells(4, i).Font.Size = 10
                        .Cells(4, i).VerticalAlignment = 3
                        .Cells(4, i).HorizontalAlignment = 3
                        .Cells(4, i).Value = pDataSource.Fields(i - 1).Name
                    Next
                    
                    '设置标题行的线条
                    .range(.Cells(4, 1), .Cells(4, pFieldsCount)).borders.LineStyle = 1
                    
                    ErrorPosition = 10
                    '设置数据区域的字体格式
                    With .range(.Cells(5, 1), .Cells(pDataSource.RecordCount + 4, pFieldsCount))
                        .VerticalAlignment = 1
                        .HorizontalAlignment = 1
                        .borders.LineStyle = 1
    '                    .Borders.Color = vbRed
                        .Font.Name = "宋体"
                        .Font.Size = 10
                    End With
                    
                End With
            End If
    
            '装载数据到数据区域
    
            ErrorPosition = 11
            If Flag = True Then
                With xclWorkSheet
            '        .Columns(1, 1).Width = 1000
            '        xclWorkSheet.Range(.Cells(2, 1), .Cells(2, 1)).Width = 5000
                    ErrorPosition = 8
                    'On Error Resume Next
                    'If pDataSource.RecordCount <= 2000 Then
                        xclWorkSheet.range(.Cells(5, 1), .Cells(5, 1)).CopyFromRecordset pDataSource
                    'Else
                        'Call CopyFromArray(xclWorkSheet, pDataSource)
                    'End If
                    'On Error GoTo 0
                End With
                ErrorPosition = 12
                'xclWorkSheet.PageSetup.CenterHeader = "&""宋体""&12扫描统计报告" & Space(10) & "&""宋体""&8打印日期：" & "&""Ms Sans Serif""&D"
                
            End If
        End If
        
ExcelNext:
        ErrorPosition = 13
        PRG.Value = 10 - m
    Next
    
    
ExportOK:
    ExportToExcel = Flag
    If Flag = False And Errorstring <> "" Then
        MsgBox Errorstring, vbOKOnly + vbInformation, "提示"
    End If
    
    '显示Excel对象
    If Flag = True Then
        xclApplication.Visible = True
    Else
        If ErrorPosition > 0 Then
            xclApplication.DisplayAlerts = False
            xclApplication.quit
        End If
    End If
    
    PRG.Visible = False
    Me.MousePointer = vbDefault
    
    Exit Function
    
Errhandle:
    Flag = False
    If ErrorPosition = 0 Then
        Errorstring = "'导出Excel...' 要求系统已经预先安装了'Microsoft Excel'，请检查系统是否已经正确安装该软件。"
    Else
        Errorstring = "Err.Number:" & Err.Number & Space(5) & "Err.Position:" & ErrorPosition & Space(5) & "Err.Description:" & Err.Description & Space(5) & "Err.Source:" & Err.Source
    End If
    Resume ExportOK
End Function


'从数组导出到Excel
Private Function CopyFromArray(ByRef xclWorkSheet As Object, ByRef pDataSource As Recordset, Optional ByRef Errorstring As String) As Boolean

    Dim recArray As Variant
    Dim recCount As Long
    Dim fldCount As Long
    Dim iCol As Long
    Dim iRow As Long
    
  
    'EXCEL 97 or earlier: Use GetRows then copy array to Excel

    ' Copy recordset to an array
    recArray = pDataSource.GetRows
    'Note: GetRows returns a 0-based array where the first
    'dimension contains fields and the second dimension
    'contains records. We will transpose this array so that
    'the first dimension contains records, allowing the
    'data to appears properly when copied to Excel
    
    ' Determine number of records

    recCount = UBound(recArray, 2) + 1 '+ 1 since 0-based array
    fldCount = pDataSource.Fields.Count

    ' Check the array for contents that are not valid when
    ' copying the array to an Excel worksheet
    For iCol = 0 To fldCount - 1
        For iRow = 0 To fldCount - 1
            ' Take care of Date fields
            If IsDate(recArray(iCol, iRow)) Then
                recArray(iCol, iRow) = Format(recArray(iCol, iRow))
            ' Take care of OLE object fields or array fields
            ElseIf IsArray(recArray(iCol, iRow)) Then
                recArray(iCol, iRow) = "无效值"
            End If
        Next iRow 'next record
    Next iCol 'next field
            
    ' Transpose and Copy the array to the worksheet,
    ' starting in cell A2
    xclWorkSheet.Cells(5, 1).Resize(recCount, fldCount).Value = TransposeDim(recArray)


End Function


Function TransposeDim(v As Variant) As Variant
' Custom Function to Transpose a 0-based array (v)
    
    Dim X As Long, Y As Long, Xupper As Long, Yupper As Long
    Dim tempArray As Variant
    
    Xupper = UBound(v, 2)
    Yupper = UBound(v, 1)
    
    ReDim tempArray(Xupper, Yupper)
    For X = 0 To Xupper
        For Y = 0 To Yupper
            tempArray(X, Y) = v(Y, X)
        Next Y
    Next X
    
    TransposeDim = tempArray

End Function


'功能：向程序运行日志文件中写入程序的出错信息
Private Sub WriteSqlLogHeader()
    Dim ErrFile As TextStream
    Dim LocalFso As Scripting.FileSystemObject
    Dim LocalFile As Scripting.File
    Dim LogFile As String
    Dim LogFileNew As String  '新
    Dim LogFileSize As Double
    Dim pCaption As String
    
    On Error GoTo Errhandle
    
    LogFile = mSqlPath & "\SQL.txt"
    
    Set LocalFso = New Scripting.FileSystemObject
    
    pCaption = vbCrLf & "=================================================================================================================" & vbCrLf
    pCaption = pCaption & "Server  : " & mServer & vbCrLf
    pCaption = pCaption & "Port    : " & mPort & vbCrLf
    pCaption = pCaption & "User ID : " & mUser & vbCrLf
    pCaption = pCaption & "Password: " & mPassword & vbCrLf
    
    pCaption = pCaption & "Time    : " & Now & vbCrLf
    
    If LocalFso.FileExists(LogFile) = True Then
        Set LocalFile = LocalFso.GetFile(LogFile)
        LogFileSize = LocalFile.Size
        If LogFileSize > 10000000 Then
            LogFileNew = "SQL(" & Format(Date, "yyyymmdd") & "." & Timer & ").txt"
            
            LocalFile.Name = LogFileNew
            Set ErrFile = LocalFso.CreateTextFile(LogFile, True, False)
        Else
            Set ErrFile = LocalFso.OpenTextFile(LogFile, ForAppending, , TristateFalse)
        End If
    Else
        Set ErrFile = LocalFso.CreateTextFile(LogFile, True, False)
    End If
    '开始写入信息文件
    ErrFile.WriteLine pCaption
    ErrFile.Close
    Set LocalFile = Nothing
    Set LocalFso = Nothing
    Set ErrFile = Nothing
    
DealOK:
    
    Exit Sub
    
Errhandle:

    
End Sub



'功能：记录SQL提交日志
Private Sub WriteSqlLog(ByVal pSqlCommands As String)
    Dim ErrFile As TextStream
    Dim LocalFso As Scripting.FileSystemObject
    Dim LocalFile As Scripting.File
    Dim LogFile As String
    Dim LogFileNew As String  '新
    Dim LogFileSize As Double
    Dim pCaption As String
    
    On Error GoTo Errhandle
    
    If Trim(pSqlCommands) = "" Then Exit Sub
    
    LogFile = mSqlPath & "\SQL.txt"
    
    Set LocalFso = New Scripting.FileSystemObject
    
    If LocalFso.FileExists(LogFile) = True Then
        Set LocalFile = LocalFso.GetFile(LogFile)
        LogFileSize = LocalFile.Size
        If LogFileSize > 10000000 Then
            LogFileNew = "SQL(" & Format(Date, "yyyymmdd") & "." & Timer & ").txt"
            
            LocalFile.Name = LogFileNew
            Set ErrFile = LocalFso.CreateTextFile(LogFile, True, False)
        Else
            Set ErrFile = LocalFso.OpenTextFile(LogFile, ForAppending, , TristateFalse)
        End If
    Else
        Set ErrFile = LocalFso.CreateTextFile(LogFile, True, False)
    End If
    '开始写入信息文件
    ErrFile.WriteLine Trim(pSqlCommands)
    ErrFile.Close
    Set LocalFile = Nothing
    Set LocalFso = Nothing
    Set ErrFile = Nothing
    
DealOK:
    
    Exit Sub
    
Errhandle:

    
End Sub

'V3.5.6--查找当前光标之前和之后的换行符号或文本结尾在整个文本框中的位置
Public Function LocatePosition(ByVal pTextBoxHwnd As Long, ByRef pBeginPosition As Long, ByRef pEndPosition As Long, Optional ByRef Errorstring As String) As Boolean

    Dim Flag As Boolean
    Dim ErrPos As Long
    Dim i As Long
    Dim pLineNo As Long
    Dim pColNo As Long
    
    Dim pCountLine As Long
    Dim StartPos As Long
    Dim EndPos As Long
    
    On Error GoTo ErrHand
    Flag = True
    ErrPos = 1

    '获取光标当前的位置：行，列
    Call GetCaretPos(pTextBoxHwnd, pLineNo, pColNo)
    
    pCountLine = 1
    
    For i = 1 To Len(TxtSQL.Text)
        If Mid(TxtSQL.Text, i, 1) = vbLf Then
            pCountLine = pCountLine + 1
            
            If pCountLine = pLineNo Then
                StartPos = i
            End If
            
            If pCountLine = pLineNo + 1 Then
                EndPos = i - 2
                Exit For
            End If
        End If
    Next
    
    If EndPos = 0 Then
        EndPos = Len(TxtSQL.Text)
    End If
    
    pBeginPosition = StartPos
    pEndPosition = EndPos
    
DealOK:

    LocatePosition = Flag
    If Flag = False And Trim(Errorstring) <> "" Then
        Errorstring = "CaretPosition.LocatePosition--->" & Errorstring
    End If
    
    '销毁局部对象

    
    Exit Function
    
ErrHand:
    
    Flag = False
    Errorstring = "Err.Number:" & Err.Number & Space(3) & "Err.Description:" & Err.Description & Space(3) & "Err.Source:" & Err.Source & Space(3) & "Err.Position:" & ErrPos
    
    Resume DealOK
End Function


