VERSION 5.00
Begin VB.Form Form2 
   Caption         =   " Register Complaint"
   ClientHeight    =   9360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13545
   LinkTopic       =   "Form2"
   Picture         =   "Register_C.frx":0000
   ScaleHeight     =   9360
   ScaleWidth      =   13545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Gadugi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   18
      ToolTipText     =   "Clear Fields"
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C000&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Gadugi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12120
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   17
      ToolTipText     =   "Go Back"
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "Gadugi"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      TabIndex        =   16
      ToolTipText     =   "Submit Record"
      Top             =   7920
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000016&
      Height          =   1095
      Left            =   4680
      MultiLine       =   -1  'True
      TabIndex        =   15
      ToolTipText     =   "Enter Your Problem"
      Top             =   7200
      Width           =   3495
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000016&
      Height          =   975
      Left            =   4680
      MultiLine       =   -1  'True
      TabIndex        =   14
      ToolTipText     =   "Enter Address"
      Top             =   5760
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000016&
      Height          =   495
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   13
      ToolTipText     =   "Enter Contact Number"
      Top             =   4920
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000016&
      Height          =   315
      Left            =   4680
      TabIndex        =   12
      ToolTipText     =   "Enter Rank"
      Top             =   4200
      Width           =   3255
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Female"
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   3360
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Male"
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   3360
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000016&
      Height          =   495
      Left            =   4680
      TabIndex        =   9
      ToolTipText     =   "Enter Age"
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000016&
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      ToolTipText     =   "Enter Name"
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Problem:-"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:-"
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No:-"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation/Rank:-"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender:-"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Age:-"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name:-"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Register 
      BackStyle       =   0  'Transparent
      Caption         =   "Registration For Complaint"
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   0
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db, d As Database
Dim rs, r As Recordset
Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Or Combo1.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
MsgBox "Please provide all information", vbInformation
Exit Sub
End If

Set db = OpenDatabase("C:\BCA(Project)\database.mdb")
Set rs = db.OpenRecordset("select * from Register_c")
rs.AddNew
rs.Fields(0).Value = Text1.Text
rs.Fields(1).Value = CInt(Text2.Text)
If Option1.Value = True Then
rs.Fields(2).Value = "Male"
End If
If Option2.Value = True Then
rs.Fields(2).Value = "Female"
End If
rs.Fields(3).Value = Combo1.Text
rs.Fields(4).Value = CDbl(Text3.Text)
rs.Fields(5).Value = Text4.Text
rs.Fields(6).Value = Text5.Text

MsgBox ("Record Saved Succesfully!!!")
rs.Update
'code to store some values in "table_c";
Set d = OpenDatabase("C:\BCA(Project)\database.mdb")
Set r = db.OpenRecordset("select * from total")
r.AddNew
r.Fields(0).Value = 1
Form1.DataGrid1.Refresh
r.Update
End Sub

Private Sub Command2_Click()
Me.Hide
Form1.Show

End Sub

Private Sub Command3_Click()
Unload Me
Me.Show
End Sub

Private Sub Form_Load()
Combo1.AddItem "Civilian"
Combo1.AddItem "Officer"
Combo1.AddItem "Commodo"
Combo1.AddItem "CEO/XO"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Combo1.Text = "Select"
Option1.Value = False
Option2.Value = False

End Sub





