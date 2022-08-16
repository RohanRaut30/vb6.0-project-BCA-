VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5460
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Login.frx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Close"
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000000&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Login "
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000000&
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password:-"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Name:-"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  If Text1 = "admin" And Text2 = "admin" Then
   Me.Hide
   Form1.Show
  
   Else
        MsgBox "Invalid Name or Password, try again!", , "Login"
  End If
End Sub

Private Sub Command2_Click()
End
End Sub

