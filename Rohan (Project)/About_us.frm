VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   0  'None
   Caption         =   "About Us"
   ClientHeight    =   5010
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8265
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   18000
      Left            =   0
      Picture         =   "About_us.frx":0000
      ScaleHeight     =   18000
      ScaleWidth      =   28800
      TabIndex        =   0
      Top             =   0
      Width           =   28800
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Version:- 1.0"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6360
         TabIndex        =   7
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyrights "
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   6600
         TabIndex        =   6
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact:- 8888xxxx      |"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   4680
         TabIndex        =   5
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Rohanraut2300@gmail.com        |"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Home        |"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   840
         TabIndex        =   3
         Top             =   4320
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   720
         X2              =   7440
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "This is Project About Complaint And  Renovation Management System. This project can make more simplified way of storeing data."
         BeginProperty Font 
            Name            =   "Ink Free"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   1695
         Left            =   720
         TabIndex        =   2
         Top             =   1440
         Width           =   5295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "About Us"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   720
         TabIndex        =   1
         Top             =   600
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
