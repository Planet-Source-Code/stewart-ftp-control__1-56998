VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "C&ancel"
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Connect"
      Default         =   -1  'True
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Text            =   "21"
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   "Port:"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "User:"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "URL:"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  frmMain.FTP1.Connect App.Title, Text1.Text, Text4.Text, Text2.Text, Text3.Text
  Unload Me
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

