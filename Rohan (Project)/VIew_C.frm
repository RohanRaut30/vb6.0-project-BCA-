VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "View Complaints"
   ClientHeight    =   8505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14085
   LinkTopic       =   "Form4"
   Picture         =   "VIew_C.frx":0000
   ScaleHeight     =   8505
   ScaleWidth      =   14085
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton del 
      Caption         =   "Delete"
      Height          =   495
      Left            =   10440
      TabIndex        =   6
      Top             =   7560
      Width           =   1095
   End
   Begin VB.CommandButton search 
      Caption         =   "Search"
      Height          =   375
      Left            =   11880
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   8280
      TabIndex        =   4
      Top             =   1560
      Width           =   3255
   End
   Begin VB.CommandButton back 
      Caption         =   "Back"
      Height          =   495
      Left            =   11880
      TabIndex        =   2
      Top             =   7560
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3240
      Top             =   7680
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   688
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\BCA(Project)\database.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\BCA(Project)\database.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Register_C"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "VIew_C.frx":5A13
      Height          =   4695
      Left            =   1320
      TabIndex        =   1
      Top             =   2400
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8281
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ALL Records"
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
            LCID            =   16393
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
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Search By Name:-"
      BeginProperty Font 
         Name            =   "Yu Gothic Medium"
         Size            =   9.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Details of Registered Complaints"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3840
      TabIndex        =   0
      Top             =   480
      Width           =   7095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub back_Click()
Me.Hide
Form1.Show
End Sub

Private Sub del_Click()
Dim conform As Integer
conform = MsgBox("Do you want to delete the record", vbYesNo + vbExclamation, "warning msg")
If conform = vbYes Then
Adodc1.Recordset.Delete
MsgBox "Record Deleted Succesfully", vbInformation, "Delete Record Confrmation"
Else
MsgBox "Record Not Deleted ", vbInformation, "Record Not Deleted"
End If
End Sub

Private Sub search_Click()
cn.Close
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\BCA(Project)\database.mdb;Persist Security Info=False"
rs.CursorLocation = adUseClient
DataGrid1.Refresh
rs.Open "select * from Register_C where name like '%" & Text1.Text & "%'", cn, adOpenDynamic, adLockOptimistic
If rs.EOF Then
MsgBox "No Record Found"
Else
Set DataGrid1.DataSource = rs
End If
End Sub


Private Sub Form_Load()
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\BCA(Project)\database.mdb;Persist Security Info=False"
rs.CursorLocation = adUseClient
rs.Open "select * from Register_C", cn, adOpenKeyset, adLockPessimistic, adCmdText
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
Set rs = Nothing

End Sub


