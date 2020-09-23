VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About FTPOCX 1.0"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4530
      TabIndex        =   1
      Top             =   150
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   150
      Picture         =   "frmAbout.frx":0442
      Top             =   120
      Width           =   600
   End
   Begin VB.Label lblCopy 
      Caption         =   "Copyright (C) 2001 AckSoft (You may use freely as long as this copyright remains intact.)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   900
      TabIndex        =   2
      Top             =   510
      Width           =   3465
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      Height          =   975
      Left            =   60
      Top             =   45
      Width           =   5775
   End
   Begin VB.Label lblName 
      Caption         =   "FTP OCX Version ()"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   945
      TabIndex        =   0
      Top             =   180
      Width           =   3435
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  lblName.Caption = "FTP OCX Version (" & App.Major & "." & App.Minor & "." & App.Revision & ")"
  Me.Caption = "FTP OCX Version (" & App.Major & "." & App.Minor & "." & App.Revision & ")"
End Sub
