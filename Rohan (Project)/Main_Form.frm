VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   Caption         =   "Complaint & Renovation Management System"
   ClientHeight    =   6915
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   13830
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   17.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Main_Form.frx":0000
   ScaleHeight     =   6915
   ScaleWidth      =   13830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Total Complaints/Renovation     "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   3000
      Width           =   3135
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   735
      Left            =   5400
      TabIndex        =   1
      Top             =   3000
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1296
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483645
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   26
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
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
         ScrollBars      =   0
         ScrollGroup     =   0
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7680
      Top             =   7200
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   "select count(T_Complaint) from total;"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\BCA(Project)\database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "total"
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   1440
      Top             =   1200
      Width           =   10935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Complaint and Renovation Management System (MES)"
      BeginProperty Font 
         Name            =   "@Microsoft JhengHei UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   1680
      TabIndex        =   0
      Top             =   1320
      Width           =   10695
   End
   Begin VB.Menu menu 
      Caption         =   "Menu"
      Begin VB.Menu rcomplaint 
         Caption         =   "Register For Complaint"
      End
      Begin VB.Menu rrenovation 
         Caption         =   "Register For Renovation"
      End
   End
   Begin VB.Menu view 
      Caption         =   "View "
      Begin VB.Menu vcomplaints 
         Caption         =   "Complaints"
      End
      Begin VB.Menu vrenovation 
         Caption         =   "Renovation"
      End
   End
   Begin VB.Menu reports 
      Caption         =   "Reports"
      Begin VB.Menu creport 
         Caption         =   "Complaints Report"
      End
      Begin VB.Menu rreport 
         Caption         =   "Renovation Report"
      End
      Begin VB.Menu Trecords 
         Caption         =   "Total Records"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
   End
   Begin VB.Menu about 
      Caption         =   "About us"
   End
   Begin VB.Menu Exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub about_Click()
Form8.Show
End Sub



Private Sub Command1_Click()
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\BCA(Project)\database.mdb;Persist Security Info=False"
rs.CursorLocation = adUseClient
rs.Open "select count(*) from total", cn, adOpenKeyset, adLockPessimistic, adCmdText
tot_c = CInt(rs)

End Sub

Private Sub creport_Click()
DataReport1.Show
End Sub



Private Sub Exit_Click()
If MsgBox("Do you want to exit", vbYesNo, "Complaint & Renovation System") = vbYes Then
End
End If
End Sub



Private Sub Form_Load()
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\BCA(Project)\database.mdb;Persist Security Info=False"
rs.CursorLocation = adUseClient
rs.Open "select count(T_Complaint),count(T_Renovation) from total", cn, adOpenKeyset, adLockPessimistic, adCmdText
Set DataGrid1.DataSource = rs
DataGrid1.Refresh
Set rs = Nothing
End Sub

Private Sub Help_Click()
MsgBox "If any Query contact here:- Mobile No:-8888083655 or visit :- www.complaint&renovation_system.com", vbInformation, "Help"
End Sub

Private Sub rcomplaint_Click()
Form2.Show
End Sub


Private Sub rrenovation_Click()
Form3.Show
End Sub

Private Sub rreport_Click()
DataReport2.Show
End Sub

Private Sub Trecords_Click()
DataReport3.Show

End Sub

Private Sub vcomplaints_Click()
Form4.Show
End Sub

Private Sub vrenovation_Click()
Form5.Show
End Sub
