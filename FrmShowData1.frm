VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{94D275A4-6691-48EC-8588-15928B9BA664}#1.0#0"; "MSIDiffSplitter.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form FrmShowData1 
   Caption         =   "DisplayData"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12015
   Icon            =   "FrmShowData1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   12015
   StartUpPosition =   2  '屏幕中心
   Begin VBUSplitterControl2.vbuSplitter2 UD 
      Height          =   7065
      Left            =   210
      TabIndex        =   11
      Top             =   210
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   12462
      Style           =   0
      SplitterTop     =   3000
      SplitterLeft    =   1418
      Begin VBUSplitterControl2.vbuSplitter2 LR 
         Height          =   2970
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   10845
         _ExtentX        =   19129
         _ExtentY        =   5239
         Style           =   0
         SplitterTop     =   1110
         SplitterLeft    =   3585
         SplitterOrientation=   1
         Begin VB.Frame Frame1 
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
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   3555
            Begin VB.CommandButton CmdTranGo 
               Height          =   500
               Left            =   2010
               Picture         =   "FrmShowData1.frx":030A
               Style           =   1  'Graphical
               TabIndex        =   8
               ToolTipText     =   "清除数据项"
               Top             =   2400
               Width           =   700
            End
            Begin VB.CommandButton CmdGo 
               Height          =   500
               Left            =   2730
               Picture         =   "FrmShowData1.frx":0BD4
               Style           =   1  'Graphical
               TabIndex        =   7
               ToolTipText     =   "执行SQL语句"
               Top             =   2400
               Width           =   700
            End
            Begin VB.CommandButton CmdClear 
               Height          =   500
               Left            =   1290
               Picture         =   "FrmShowData1.frx":149E
               Style           =   1  'Graphical
               TabIndex        =   9
               ToolTipText     =   "清除数据项"
               Top             =   2400
               Width           =   700
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
               Height          =   345
               ItemData        =   "FrmShowData1.frx":1D68
               Left            =   1290
               List            =   "FrmShowData1.frx":1D6A
               TabIndex        =   5
               Top             =   2010
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
               Height          =   330
               Left            =   1290
               TabIndex        =   2
               Top             =   930
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
               Height          =   345
               ItemData        =   "FrmShowData1.frx":1D6C
               Left            =   1290
               List            =   "FrmShowData1.frx":1D79
               Style           =   2  'Dropdown List
               TabIndex        =   0
               Top             =   180
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
               Height          =   330
               Left            =   1290
               TabIndex        =   4
               Top             =   1650
               Width           =   2200
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
               Height          =   330
               Left            =   1290
               TabIndex        =   3
               Top             =   1290
               Width           =   2200
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
               Height          =   330
               Left            =   1290
               TabIndex        =   1
               Top             =   570
               Width           =   2200
            End
            Begin VB.Image Image1 
               Height          =   495
               Left            =   180
               Picture         =   "FrmShowData1.frx":1D9D
               Stretch         =   -1  'True
               Top             =   2400
               Width           =   450
            End
            Begin VB.Label Label6 
               Caption         =   "Server Port:"
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
               Left            =   120
               TabIndex        =   31
               Top             =   990
               Width           =   1650
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
               Left            =   120
               TabIndex        =   30
               Top             =   270
               Width           =   1200
            End
            Begin VB.Label Label4 
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
               Left            =   120
               TabIndex        =   29
               Top             =   2100
               Width           =   1200
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
               Left            =   120
               TabIndex        =   28
               Top             =   1740
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
               Left            =   120
               TabIndex        =   27
               Top             =   1350
               Width           =   1080
            End
            Begin VB.Label Label1 
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
               Left            =   120
               TabIndex        =   26
               Top             =   660
               Width           =   1200
            End
         End
         Begin VB.TextBox TxtSQL 
            Height          =   2970
            HideSelection   =   0   'False
            Left            =   3645
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   6
            Top             =   0
            Width           =   7200
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4005
         Left            =   0
         TabIndex        =   12
         Top             =   3060
         Width           =   10845
         _ExtentX        =   19129
         _ExtentY        =   7064
         _Version        =   393216
         Style           =   1
         Tabs            =   11
         Tab             =   10
         TabsPerRow      =   12
         TabHeight       =   520
         TabCaption(0)   =   " 1  "
         TabPicture(0)   =   "FrmShowData1.frx":2667
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "GridData(0)"
         Tab(0).Control(1)=   "PRG"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   " 2  "
         TabPicture(1)   =   "FrmShowData1.frx":2683
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "GridData(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   " 3  "
         TabPicture(2)   =   "FrmShowData1.frx":269F
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "GridData(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   " 4  "
         TabPicture(3)   =   "FrmShowData1.frx":26BB
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "GridData(3)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   " 5  "
         TabPicture(4)   =   "FrmShowData1.frx":26D7
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "GridData(4)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   " 6  "
         TabPicture(5)   =   "FrmShowData1.frx":26F3
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "GridData(5)"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   " 7  "
         TabPicture(6)   =   "FrmShowData1.frx":270F
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "GridData(6)"
         Tab(6).ControlCount=   1
         TabCaption(7)   =   " 8  "
         TabPicture(7)   =   "FrmShowData1.frx":272B
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "GridData(7)"
         Tab(7).ControlCount=   1
         TabCaption(8)   =   " 9  "
         TabPicture(8)   =   "FrmShowData1.frx":2747
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "GridData(8)"
         Tab(8).ControlCount=   1
         TabCaption(9)   =   " 10 "
         TabPicture(9)   =   "FrmShowData1.frx":2763
         Tab(9).ControlEnabled=   0   'False
         Tab(9).Control(0)=   "GridData(9)"
         Tab(9).ControlCount=   1
         TabCaption(10)  =   " OUT "
         TabPicture(10)  =   "FrmShowData1.frx":277F
         Tab(10).ControlEnabled=   -1  'True
         Tab(10).Control(0)=   "TxtInformation"
         Tab(10).Control(0).Enabled=   0   'False
         Tab(10).ControlCount=   1
         Begin RichTextLib.RichTextBox TxtInformation 
            Height          =   3660
            Left            =   30
            TabIndex        =   10
            Top             =   320
            Width           =   10785
            _ExtentX        =   19024
            _ExtentY        =   6456
            _Version        =   393217
            BackColor       =   -2147483624
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"FrmShowData1.frx":279B
         End
         Begin MSDataGridLib.DataGrid GridData 
            Height          =   3660
            Index           =   0
            Left            =   -74970
            TabIndex        =   13
            Top             =   320
            Width           =   10785
            _ExtentX        =   19024
            _ExtentY        =   6456
            _Version        =   393216
            BackColor       =   -2147483624
            ForeColor       =   16711680
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
            Index           =   1
            Left            =   -74970
            TabIndex        =   14
            ToolTipText     =   "查询结果"
            Top             =   320
            Width           =   10785
            _ExtentX        =   19024
            _ExtentY        =   6456
            _Version        =   393216
            BackColor       =   -2147483624
            ForeColor       =   16711680
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
            TabIndex        =   15
            ToolTipText     =   "查询结果"
            Top             =   320
            Width           =   10785
            _ExtentX        =   19024
            _ExtentY        =   6456
            _Version        =   393216
            BackColor       =   -2147483624
            ForeColor       =   16711680
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
            TabIndex        =   16
            ToolTipText     =   "查询结果"
            Top             =   320
            Width           =   10785
            _ExtentX        =   19024
            _ExtentY        =   6456
            _Version        =   393216
            BackColor       =   -2147483624
            ForeColor       =   16711680
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
            TabIndex        =   17
            ToolTipText     =   "查询结果"
            Top             =   320
            Width           =   10785
            _ExtentX        =   19024
            _ExtentY        =   6456
            _Version        =   393216
            BackColor       =   -2147483624
            ForeColor       =   16711680
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
            TabIndex        =   18
            ToolTipText     =   "查询结果"
            Top             =   320
            Width           =   10785
            _ExtentX        =   19024
            _ExtentY        =   6456
            _Version        =   393216
            BackColor       =   -2147483624
            ForeColor       =   16711680
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
            TabIndex        =   19
            ToolTipText     =   "查询结果"
            Top             =   320
            Width           =   10785
            _ExtentX        =   19024
            _ExtentY        =   6456
            _Version        =   393216
            BackColor       =   -2147483624
            ForeColor       =   16711680
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
            TabIndex        =   20
            ToolTipText     =   "查询结果"
            Top             =   320
            Width           =   10785
            _ExtentX        =   19024
            _ExtentY        =   6456
            _Version        =   393216
            BackColor       =   -2147483624
            ForeColor       =   16711680
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
            TabIndex        =   21
            ToolTipText     =   "查询结果"
            Top             =   320
            Width           =   10785
            _ExtentX        =   19024
            _ExtentY        =   6456
            _Version        =   393216
            BackColor       =   -2147483624
            ForeColor       =   16711680
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
            TabIndex        =   22
            ToolTipText     =   "查询结果"
            Top             =   320
            Width           =   10785
            _ExtentX        =   19024
            _ExtentY        =   6456
            _Version        =   393216
            BackColor       =   -2147483624
            ForeColor       =   16711680
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
            Left            =   -68625
            TabIndex        =   23
            Top             =   60
            Visible         =   0   'False
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   397
            _Version        =   393216
            Appearance      =   1
         End
      End
   End
End
Attribute VB_Name = "FrmShowData1"
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

Private Sub CBOdatabases_dropdown()
    Dim Errorstring As String
    Dim Flag As Boolean
'    Dim LocalSql As String
    Dim LocalConnection As ADODB.Connection
    Dim LocalRS As ADODB.Recordset
    Dim mDatabasetype As Long
    
    On Error GoTo ErrorHandleOK
    Screen.MousePointer = vbHourglass
    CmdGo.Enabled = False
    CmdClear.Enabled = False
    DoEvents
    Flag = True

    If Trim(txtServer) = "" Or Trim(txtUID) = "" Then
        Flag = False
        Errorstring = "信息不全：数据库服务和用户名不能为空！" & Space(5)
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
        If InStr(1, LCase(CBOtype.Text), "sql") > 0 Then
            mDatabasetype = 0
        End If
        If InStr(1, LCase(CBOtype.Text), "sybase") > 0 Then
            mDatabasetype = 1
        End If
        If InStr(1, LCase(CBOtype.Text), "oracle") > 0 Then
            mDatabasetype = 2
        End If
        If GetADOConnection(mDatabasetype, Trim(txtServer), Trim(txtUID), Trim(txtPWD), LocalConnection, "master", Trim(txtPort), Errorstring) = False Then
            Flag = False
        End If
    End If
    
    If Flag = True Then
        CBOdatabases.Clear
        CBOdatabases.Refresh
        If mDatabasetype = 0 Or mDatabasetype = 1 Then
            Set LocalRS = LocalConnection.Execute("sp_databases")
        Else
            Set LocalRS = LocalConnection.Execute("select distinct tablespace_name from dba_free_space")
        End If
        If Not (LocalRS.EOF And LocalRS.BOF) Then
            While LocalRS.EOF = False
                If Trim(LocalRS.Fields(0) & "") <> "" Then
                    CBOdatabases.AddItem Trim(LocalRS.Fields(0) & "")
                    LocalRS.MoveNext
                End If
            Wend
        End If
    End If
    
DatabaseOK:
    Set LocalRS = Nothing
    Set LocalConnection = Nothing
    CmdGo.Enabled = True
    CmdClear.Enabled = True
    Screen.MousePointer = vbDefault
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
        
    If InStr(1, LCase(CBOtype.Text), "oracle") > 0 Then
'        CBOdatabases.Text = ""
'        CBOdatabases.ForeColor = vbWhite
'        CBOdatabases.BackColor = &H80000003
'        CBOdatabases.Enabled = False
        Label4 = "TableSapces:"
    Else
'        CBOdatabases.ForeColor = vbBlack
'        CBOdatabases.BackColor = vbWhite
'        CBOdatabases.Enabled = True
        Label4 = "Databases:"
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
        SSTab1.TabVisible(i) = False
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
    Dim BeginTime As Double
    Dim EndTime As Double
    Dim strSQL As String
    Dim k As Long
    Dim Played As Boolean
    Dim AffectedRows As Double
    
    '记录每条SQL的执行情况
    Dim InfoArray() As String
    
    
    On Error GoTo ErrorHandleOK
    
    Screen.MousePointer = vbHourglass
    Flag = True
    
    
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

        
        ErrorPosition = 1
        Call ClearData
        

        ErrorPosition = 2
        
        If mDatabasetype <> 2 Then
            If Trim(txtServer) = "" Or Trim(txtUID) = "" Or Trim(CBOdatabases.Text) = "" Or Trim(TxtSQL) = "" Then
                Flag = False
                Errorstring = "Please check the integrality of the informations."
            End If
        Else
            If Trim(txtServer) = "" Or Trim(txtUID) = "" Or Trim(TxtSQL) = "" Then
                Flag = False
                Errorstring = "Please check the integrality of the informations."
            End If
        End If
    End If
    
    If Flag = True And Trim(txtPort) <> "" Then
        ErrorPosition = 3
        If IsNumeric(Trim(txtPort)) = False Then
            Flag = False
            Errorstring = "数据库服务器侦听的端口号(Server Port)必须是正整数！"
        Else
            If InStr(1, Trim(txtPort), ".") > 0 Or Val(Trim(txtPort)) <= 0 Then
                    Flag = False
                    Errorstring = "数据库服务器侦听的端口号(Server Port)必须是正整数！"
            End If
        End If
    End If
    
    If Flag = True Then
        ErrorPosition = 4
        If GetADOConnection(mDatabasetype, Trim(txtServer), Trim(txtUID), Trim(txtPWD), LocalConnection, Trim(CBOdatabases.Text), Trim(txtPort), Errorstring) = False Then
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
        
        ErrorPosition = 6
        LocalSql = Trim(Replace(Replace(Replace(Replace(strSQL, Chr(13), " "), Chr(10), " "), "，", ","), "‘", "'"))
        If Right(LocalSql, 1) = ";" Then
            LocalSql = Left(LocalSql, Len(LocalSql) - 1)
        End If
        
        SqlArray = Split(LocalSql, ";")
        k = -1
        With PRG
            .Visible = True
            .Min = 0
            .Max = UBound(SqlArray) + 1
            .Value = 0
        End With
        
        ErrorPosition = 7
        
        ReDim InfoArray(UBound(SqlArray))
        
        For i = 0 To UBound(SqlArray)
            If i < 10 Then
                If Trim(SqlArray(i)) <> "" Then
                    k = k + 1
                    SSTab1.TabVisible(k) = True


                    ErrorPosition = 8
                    BeginTime = Timer
                    
                    Set mLocalRecordset = LocalConnection.Execute(SqlArray(i))
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
                    End If
                    
                    
                    
                    
                    ErrorPosition = 10
                    
ExecuteOK:
                    GridData(k).ToolTipText = Trim(SqlArray(i))
                    ErrorPosition = 11
                    GridData(k).Refresh
                    ErrorPosition = 12
                    PRG.Value = i + 1
                    ErrorPosition = 13
                    
                    
                    DoEvents
                End If
            End If
        Next i
        

        
        SSTab1.Tab = 10
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
                TxtInformation.SelColor = vbBlue
            Else
                TxtInformation.SelColor = vbRed
            End If

        Next
    End If
    
    Set LocalConnection = Nothing

    CmdGo.Enabled = True
    CmdTranGo.Enabled = True
    CmdClear.Enabled = True
    PRG.Visible = False
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
ErrorHandleOK:
    Flag = False
    
    If ErrorPosition = 8 Then
        InfoArray(i) = "SQL(" & i + 1 & ")： Status: Failure!  Error" & Err.Number & "   " & Err.Description & "   " & Err.Source
        Resume ExecuteOK
    Else
        Errorstring = "Err.Number:" & Err.Number & "  Err.Description:" & Err.Description
    End If
    
    Resume DisplayOk
    
End Sub



Private Sub CmdGo_Click()
    Call ExecuteSql(False)
End Sub


Private Sub CmdTranGo_Click()
    Call ExecuteSql(True)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    If KeyCode = vbKeyF5 Then
        If CBOdatabases.Text <> "" Then
            Call ExecuteSql
            KeyCode = 0
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
        TxtInformation.Text = "**************************************************************************" & vbCrLf & "快捷键说明：" & vbCrLf & "F2：撤销对SQL命令框中内容的最近一次编辑" & vbCrLf & "F3：取消对SQL命令框中内容执行的撤销操作" & vbCrLf & "F5：执行SQL命令" & vbCrLf & "F9：显示所选数据库的所有表对象" & vbCrLf & "F10：调整SQL命令框的大小至最大，或还原至正常大小" & vbCrLf & "**************************************************************************" & vbCrLf & "夏妙 2004.09.11 Email:ToTopx@163.com" & vbCrLf & "**************************************************************************"
        TxtInformation.SelStart = 0
        TxtInformation.SelLength = Len(TxtInformation.Text)
        TxtInformation.SelColor = vbBlue
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyF10 Then
        
        
        If mSQL_CurrentIsFullMode = True Then
            TxtSQL.Top = mTop
            TxtSQL.Left = mLeft
            TxtSQL.Height = mHeight
            TxtSQL.Width = mWidth
            
            mSQL_CurrentIsFullMode = False
            
        Else
            mTop = TxtSQL.Top
            mLeft = TxtSQL.Left
            mHeight = TxtSQL.Height
            mWidth = TxtSQL.Width
            TxtSQL.Top = 0
            TxtSQL.Left = 0
            TxtSQL.Height = Me.Height - 400
            TxtSQL.Width = Me.Width - 120
            
            mSQL_CurrentIsFullMode = True
            
        End If
        
        KeyCode = 0
    End If
    
    If KeyCode = vbKeyF2 And mCurrentCXIndex >= 1 And mCurrentCXIndex <= 100 Then
        
        

        Call AddQXCXarray
        
'        If mCurrentCXIndex = 0 Then
'            TxtSQL = mSQLarray(1)
'        Else
            TxtSQL = mSQLarray(mCurrentCXIndex)
'        End If
        
        
        mCurrentCXIndex = mCurrentCXIndex - 1
        If mCurrentCXIndex = 0 Then
            mCurrentCXMaxIndex = 0
        End If
        
        KeyCode = 0
        
    End If
    
    If KeyCode = vbKeyF3 And mCurrentQXCXIndex >= 1 And mCurrentQXCXIndex <= 100 Then
        
        
'        If mCurrentQXCXIndex = 0 Then
'            mCurrentQXCXIndex = 1
'        End If

'        If mCurrentQXCXIndex = 0 Then
'            TxtSQL = mSQLQXCXarray(1)
'        Else
            TxtSQL = mSQLQXCXarray(mCurrentQXCXIndex)
'        End If

        
        mCurrentQXCXIndex = mCurrentQXCXIndex - 1
        If mCurrentQXCXIndex = 0 Then
            mCurrentQXCXMaxIndex = 0
        End If
        
        KeyCode = 0
        
    End If
    
End Sub

Private Sub Form_Load()
    Dim DBType As String
    Dim LocalFont As New StdFont
    Dim LocalFSO As New Scripting.FileSystemObject
    Dim x As Object
    Dim i As Long
    
    Me.Caption = "DisplayData " & App.Major & "." & App.Minor & "." & App.Revision

    Image1.Visible = True
    Image1.Stretch = True

    PRG.Visible = False
    
    TxtSQL.ToolTipText = "SQL Command"
'    TxtInformation.ToolTipText = "Informations"
    GridData(0).ToolTipText = "Results"
    
    For i = 1 To 9
        SSTab1.TabVisible(i) = False
    Next i
    mFrmHeight = Me.Height
    mFrmWidth = Me.Width
    
    If LocalFSO.FileExists(App.Path & "\Configure.ini") = True Then
        gFontName = Trim(GetConfigFileString(App.Path & "\Configure.ini", "Face", "FontName"))
        gFontSize = Trim(GetConfigFileString(App.Path & "\Configure.ini", "Face", "FontSize"))
        txtServer = Trim(GetConfigFileString(App.Path & "\Configure.ini", "DatabaseServerInfo", "DatabaseServerName"))
        txtPort = Trim(GetConfigFileString(App.Path & "\Configure.ini", "DatabaseServerInfo", "DatabaseServerPort"))
        txtUID = Trim(GetConfigFileString(App.Path & "\Configure.ini", "DatabaseServerInfo", "DatabaseUserName"))
        txtPWD = Trim(GetConfigFileString(App.Path & "\Configure.ini", "DatabaseServerInfo", "DatabasePassword"))
        DBType = Trim(GetConfigFileString(App.Path & "\Configure.ini", "DatabaseServerInfo", "DatabaseServerType"))
        If IsNumeric(DBType) = False Then
            DBType = "0"
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
    
    CBOtype.ListIndex = Val(DBType)
    
    On Error GoTo DefaultFont
    LocalFont.Name = gFontName
    LocalFont.Size = gFontSize
    For Each x In Me.Controls
        If x.Name <> "Status" And x.Name <> "PRG" And x.Name <> "Image1" And x.Name <> "UD" And x.Name <> "LR" And x.Name <> "TxtInformation" Then
          Set x.Font = LocalFont
        End If
    Next
'
'    '移动Grid
'    For i = 0 To 9
'        GridData(i).Move 30, 320, 1785, 3660
'    Next i
'
'    TxtInformation.Move 30, 320, 1785, 3660

LoadOk:
    Set LocalFont = Nothing
    Set LocalFSO = Nothing

    Exit Sub
    
DefaultFont:
    MsgBox Err.Description, vbOKOnly + vbInformation, "提示"
    Resume LoadOk
End Sub


Private Sub Form_Resize()
    Dim i As Long
    
    If Me.WindowState <> vbMinimized Then
    
        '移动分栏控件
        UD.Move 0, 0, Me.Width, Me.Height
        UD.SplitterTop = 3000
        LR.SplitterLeft = 3585
        
        '移动进度条
        PRG.Left = 6375
        PRG.Width = Me.Width - PRG.Left - 145
        
        '移动Grid
        For i = 0 To 9
            If SSTab1.Width - 60 > 0 Then
                GridData(i).Width = SSTab1.Width - 60
            End If
            If SSTab1.Height - 800 > 0 Then
                GridData(i).Height = SSTab1.Height - 800
            End If
        Next i
        
        '移动信息提示框
        If SSTab1.Width - 60 > 0 Then
            TxtInformation.Width = SSTab1.Width - 60
        End If
        If SSTab1.Height - 800 > 0 Then
            TxtInformation.Height = SSTab1.Height - 800
        End If
        
    End If
    
End Sub




Private Sub TxtSQL_Change()
'    If Right(TxtSQL.Text, Len(TxtSQL.Text) - InStrRev(TxtSQL.Text, " ")) = "select" Then
'        TxtSQL.SelStart = InStrRev(TxtSQL.Text, " ")
'        TxtSQL.SelLength = 6
'        TxtSQL.SelColor = vbBlue
'    End If
    'mSQLarray(0) = TxtSQL
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

Private Sub vbuSplitter21_Resize()
    MsgBox TxtSQL.Width
End Sub

Private Sub UD_Resize()
    Dim i As Long
    
    '移动Grid
    For i = 0 To 9
        If SSTab1.Width - 60 > 0 Then
            GridData(i).Width = SSTab1.Width - 60
        End If
        If SSTab1.Height - 800 > 0 Then
            GridData(i).Height = SSTab1.Height - 800
        End If
    Next i
    
    '移动信息提示框
    If SSTab1.Width - 60 > 0 Then
        TxtInformation.Width = SSTab1.Width - 60
    End If
    If SSTab1.Height - 800 > 0 Then
        TxtInformation.Height = SSTab1.Height - 800
    End If
    
End Sub
