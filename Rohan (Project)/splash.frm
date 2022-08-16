VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form7 
   BorderStyle     =   0  'None
   Caption         =   "Developed By Rohan Raut"
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8865
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   5535
      Left            =   0
      Picture         =   "splash.frx":0000
      ScaleHeight     =   5475
      ScaleWidth      =   9075
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         ToolTipText     =   "Please Wait..."
         Top             =   3720
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Timer Timer1 
         Interval        =   105
         Left            =   8280
         Top             =   120
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading..."
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   4
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Version 1.0"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   5040
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Management System"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   17.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   495
         Left            =   3840
         TabIndex        =   2
         Top             =   2040
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Complaint And Renovation "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   17.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   1
         Top             =   1560
         Width           =   4815
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Timer1.Enabled = True
End Sub







Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 5
Label4.Caption = "Loading..."
If (ProgressBar1.Value = ProgressBar1.Max) Then
Timer1.Enabled = 0
Unload Me
Form6.Show
End If
End Sub
