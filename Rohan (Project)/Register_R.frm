VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Register Renovation"
   ClientHeight    =   10275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14475
   LinkTopic       =   "Form3"
   Picture         =   "Register_R.frx":0000
   ScaleHeight     =   10275
   ScaleWidth      =   14475
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
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
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8160
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
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
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Go Back"
      Top             =   8880
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Renovation For :-"
      Height          =   2895
      Left            =   2280
      TabIndex        =   14
      Top             =   7080
      Width           =   10335
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Submit"
         DisabledPicture =   "Register_R.frx":3B9FC
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
         Left            =   6600
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Submit"
         Top             =   1800
         UseMaskColor    =   -1  'True
         Width           =   1815
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H80000003&
         Caption         =   "New Windows/Doors"
         Height          =   195
         Left            =   6360
         TabIndex        =   21
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000016&
         Height          =   375
         Left            =   3000
         TabIndex        =   20
         Top             =   1920
         Width           =   1455
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H80000003&
         Caption         =   "Lights/Fan"
         Height          =   255
         Left            =   4200
         TabIndex        =   19
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H80000003&
         Caption         =   "Tiles"
         Height          =   255
         Left            =   4200
         TabIndex        =   18
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H80000003&
         Caption         =   "New Furniture"
         Height          =   195
         Left            =   1320
         TabIndex        =   17
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H80000003&
         Caption         =   "Painting"
         Height          =   375
         Left            =   1320
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "150/-"
         Height          =   255
         Left            =   5520
         TabIndex        =   29
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "2700/-"
         Height          =   255
         Left            =   2760
         TabIndex        =   28
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "1000/-"
         Height          =   255
         Left            =   8400
         TabIndex        =   27
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "1050/-"
         Height          =   255
         Left            =   5400
         TabIndex        =   26
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "100/-"
         Height          =   255
         Left            =   2760
         TabIndex        =   25
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000003&
         Caption         =   "Other Charges:-"
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   2040
         Width           =   1455
      End
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000016&
      Height          =   1335
      Left            =   5040
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   5400
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000016&
      Height          =   375
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   12
      ToolTipText     =   "Enter Only Number"
      Top             =   4680
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000016&
      Height          =   315
      Left            =   5040
      TabIndex        =   11
      ToolTipText     =   "Select Rank"
      Top             =   3840
      Width           =   2655
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Female"
      Height          =   375
      Left            =   6600
      TabIndex        =   10
      Top             =   3000
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Male"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5040
      MaskColor       =   &H00C0FFC0&
      TabIndex        =   9
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000016&
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      ToolTipText     =   "Enter your age"
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000016&
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      ToolTipText     =   "Enter Name"
      Top             =   1560
      Width           =   3255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:-"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No:-"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation/Rank:-"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender:-"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Age:-"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:-"
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Registration For Renovation"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   600
      Width           =   5415
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Dim sum As Integer


Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Or Combo1.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
MsgBox "Please provide all information", vbInformation
Exit Sub
End If

Set db = OpenDatabase("C:\BCA(Project)\database.mdb")
Set rs = db.OpenRecordset("select * from Register_R")
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
sum = 0
If Check1.Value = 1 Then
sum = sum + 100
End If
If Check2.Value = 1 Then
sum = sum + 1050
End If
If Check3.Value = 1 Then
sum = sum + 1000
End If
If Check4.Value = 1 Then
sum = sum + 2700
End If
If Check5.Value = 1 Then
sum = sum + 150
End If

sum = sum + CInt(Text5.Text)

rs.Fields(6).Value = sum
MsgBox ("Record Saved Succesfully!!!")
MsgBox ("You have to pay: Rs." & sum & "/-")

rs.Update
'code to add total renovation
Set d = OpenDatabase("C:\BCA(Project)\database.mdb")
Set r = db.OpenRecordset("select * from total")
r.AddNew
r.Fields(0).Value = 1
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
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Combo1.AddItem "Civilian"
Combo1.AddItem "Sailor"
Combo1.AddItem "Officer"
Combo1.AddItem "CEO/XO"
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Check5.Value = 0
Option1.Value = 0
Option2.Value = 0

End Sub

