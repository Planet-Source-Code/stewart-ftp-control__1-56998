VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\..\FTPOCX~1\ftpOCX.vbp"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6150
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "Make Dir"
      Height          =   495
      Left            =   10320
      TabIndex        =   14
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Connect"
      Height          =   615
      Left            =   3000
      TabIndex        =   13
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Disconnect"
      Height          =   495
      Left            =   10320
      TabIndex        =   12
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "About"
      Height          =   495
      Left            =   10320
      TabIndex        =   11
      Top             =   5280
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar pb 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   10
      Top             =   5865
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin ftpOCX.FTP FTP1 
      Left            =   3240
      Top             =   2040
      _extentx        =   1058
      _extenty        =   1058
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Refresh"
      Height          =   495
      Left            =   10320
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
      Height          =   495
      Left            =   10320
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Rename"
      Height          =   495
      Left            =   10320
      TabIndex        =   7
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<<"
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   120
      TabIndex        =   4
      Top             =   2955
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   2775
   End
   Begin MSComctlLib.ImageList img 
      Left            =   7200
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0552
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstMain 
      Height          =   5415
      Left            =   4200
      TabIndex        =   1
      Top             =   360
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   9551
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "img"
      SmallIcons      =   "img"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filename"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Filesize"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date Created"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date Last Modified"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  'FTP1.UploadFile "testfile.txt", , "Hello this is a test"
  FTP1.UploadFile File1.FileName, File1.Path & "\" & File1.FileName
  DoList "*.*"
End Sub

Private Sub Command2_Click()
  Dim Recieve As String
  FTP1.DownloadFile lstMain.SelectedItem.Text, File1.Path & "\" & lstMain.SelectedItem.Text
  'Comment the above line and uncomment the following 2 lines to see it when it doesn't save
  'Recieve = FTP1.DownloadFile(lstMain.SelectedItem.Text, , True)
  'MsgBox Recieve
  File1.Refresh
  DoList "*.*"
End Sub

Private Sub Command3_Click()
  Dim tstStr As String
  tstStr = InputBox("Enter the new name", "New Name", lstMain.SelectedItem.Text)
  FTP1.RenameSelection lstMain.SelectedItem.Text, tstStr
  DoList "*.*"
End Sub

Private Sub Command4_Click()
  FTP1.DeleteSelection lstMain.SelectedItem.Text
  DoList "*.*"
End Sub

Private Sub Command5_Click()
  DoList "*.*"
End Sub

Private Sub Command6_Click()
  FTP1.About
End Sub

Private Sub Command7_Click()
  FTP1.Disconnect
  lstMain.ListItems.Clear
End Sub

Private Sub Command8_Click()
  frmConnect.Show
End Sub

Private Sub Command9_Click()
  Dim NewDir As String
  NewDir = InputBox("Enter the new directory", "New Directory")
  FTP1.MakeDir NewDir
  DoList "*.*"
End Sub

Private Sub Dir1_Change()
  File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
  Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
  FTP1.ConnectionType = CONNECT_PASSIVE
  FTP1.TransferType = TRANSFER_BINARY
End Sub

Private Sub FTP1_GetError(Error As String, Func As String, ErrorNum As Long)
  MsgBox "Error: " & ErrorNum & Chr(10) & "Function: " & Func & Chr(10) & "Generated the following error:" & Chr(10) & Error
End Sub

Private Function DoList(sFilter As String)
  Dim lst As ListItem
  Dim Item As New ftpOCX.cDirItem
  FTP1.GetDirectoryListing sFilter
  lstMain.ListItems.Clear
  For Each Item In FTP1.Directory
    If Item.Directory = True Then
      Set lst = lstMain.ListItems.Add(, , Item.FileName, 1, 1)
      lst.SubItems(1) = "Directory"
    Else
      Set lst = lstMain.ListItems.Add(, , Item.FileName, 2, 2)
      lst.SubItems(1) = Item.FileSize
    End If
    
    lst.SubItems(2) = Item.CreationTime
    lst.SubItems(3) = Item.LastWriteTime
  Next
  Label1.Caption = FTP1.GetFTPDirectory
End Function

Private Sub FTP1_Message(MsgNum As ftpOCX.MessageTypes)
  If MsgNum = MCONNECTED Then
    DoList "*.*"
    MsgBox "Connection succeeded"
  ElseIf MsgNum = MDELETED Then
    MsgBox "File Deleted"
  ElseIf MsgNum = MRENAMED Then
    MsgBox "File Renamed"
  ElseIf MsgNum = MDOWNLOADED Then
    MsgBox "File Downloaded Successfully"
  ElseIf MsgNum = MUPLOADED Then
    MsgBox "File Uploaded Successfully"
  ElseIf MsgNum = MDISCONNECTED Then
    MsgBox "Disconnected from server"
  ElseIf MsgNum = MDIRCREATED Then
    MsgBox "Directory Created"
  End If
End Sub

Private Sub FTP1_Progress(Total As Long, Current As Long)
  On Error Resume Next
  pb.Max = Total
  pb.Value = Current
End Sub

Private Sub lstMain_DblClick()
  If lstMain.SelectedItem.SubItems(1) = "Directory" Then
    FTP1.SetDirectory lstMain.SelectedItem.Text
    DoList "*.*"
  End If
End Sub
