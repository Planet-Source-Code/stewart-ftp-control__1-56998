VERSION 5.00
Begin VB.UserControl FTP 
   CanGetFocus     =   0   'False
   ClientHeight    =   690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   690
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   690
   ScaleWidth      =   690
   ToolboxBitmap   =   "FTP.ctx":0000
   Begin VB.Image imgBack 
      Height          =   600
      Left            =   0
      Picture         =   "FTP.ctx":0312
      Top             =   0
      Width           =   600
   End
End
Attribute VB_Name = "FTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'**********************************************************
'* These are used to make the setting of the variables    *
'* profesional looking on a friends advice. So when you   *
'* type it you get options instead of just having to know *
'**********************************************************
Public Enum Connections
  CONNECT_PASSIVE = INTERNET_FLAG_PASSIVE
  CONNECT_EXISTING = INTERNET_FLAG_EXISTING_CONNECT
End Enum

Public Enum Transfers
  TRANSFER_ASCII = FTP_TRANSFER_TYPE_ASCII
  TRANSFER_BINARY = FTP_TRANSFER_TYPE_BINARY
End Enum
Public Enum MessageTypes
  MCONNECTED = MESSAGE_CONNECT
  MRENAMED = MESSAGE_RENAME
  MDELETED = MESSAGE_DELETE
  MDOWNLOADED = 3
  MUPLOADED = 4
  MDISCONNECTED = 5
  MDIRCREATED = 6
End Enum

'Default Property Values:
Const m_def_ConnectionType = &H8000000
Const m_def_TransferType = &H1
Const m_def_Enabled = 0
'Property Variables:
Dim m_ConnectionType As Connections
Dim m_TransferType As Transfers
Dim m_Enabled As Boolean
Public mCol As New Collection
Public mDirItem As New cDirItem
'Event Declarations:
Public Event GetError(Error As String, Func As String, ErrorNum As Long)
Public Event Message(MsgNum As MessageTypes)  'This will be triggered for msg's. IE
Public Event Progress(Total As Long, Current As Long)



Private Sub UserControl_Resize()
  If UserControl.Width > imgBack.Width Then UserControl.Width = imgBack.Width
  If UserControl.Height > imgBack.Height Then UserControl.Height = imgBack.Height
  If UserControl.Width < imgBack.Width Then UserControl.Width = imgBack.Width
  If UserControl.Height < imgBack.Height Then UserControl.Height = imgBack.Height
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
'    m_User = m_def_User
'    m_Password = m_def_Password
'    m_URL = m_def_URL
'    m_Port = m_def_Port
    m_ConnectionType = m_def_ConnectionType
    m_TransferType = m_def_TransferType
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_ConnectionType = PropBag.ReadProperty("ConnectionType", m_def_ConnectionType)
    m_TransferType = PropBag.ReadProperty("TransferType", m_def_TransferType)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("ConnectionType", m_ConnectionType, m_def_ConnectionType)
    Call PropBag.WriteProperty("TransferType", m_TransferType, m_def_TransferType)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8
Public Function GetFileSize() As Long
  
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
'**********************************************************
'* This is used to connect from within the app using the  *
'* control. It's pretty easy to do. When you call it it   *
'* will connect, or if unable to return an error.         *
'**********************************************************
Public Function Connect(ConnectionName As String, URL As String, Port As String, User As String, Password As String, Optional ProxyName As String, Optional ProxyBypass As String) As Boolean
  On Error GoTo errHandler
  hSession = InternetOpen(ConnectionName, INTERNET_OPEN_TYPE_DIRECT, ProxyName, ProxyBypass, INTERNET_FLAG_NO_CACHE_WRITE)
  hConnect = InternetConnect(hSession, URL, Port, User, Password, INTERNET_SERVICE_FTP, ConnectionType, 0)
  If hSession <> 0 And hConnect <> 0 Then
    'RaiseEvent Connected 'Connection succeeded.
    RaiseEvent Message(MCONNECTED)
    Connect = True
    Exit Function  'Exit before an error can be called because if we got this far
                   'none exist.
  End If
errHandler:
  FTPError Err.LastDllError, "Connect"
  InternetCloseHandle (hSession)
  InternetCloseHandle (hConnect)
  Connect = False
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
'**********************************************************
'* Used to set connection type                            *
'**********************************************************
Public Property Get ConnectionType() As Connections
    ConnectionType = m_ConnectionType
End Property
'**********************************************************
'* Used to set connection type                            *
'**********************************************************
Public Property Let ConnectionType(ByVal New_ConnectionType As Connections)
    m_ConnectionType = New_ConnectionType
    PropertyChanged "ConnectionType"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
'**********************************************************
'* Used to set transfer type                              *
'**********************************************************
Public Property Get TransferType() As Transfers
    TransferType = m_TransferType
End Property
'**********************************************************
'* Used to set transfer type                              *
'**********************************************************
Public Property Let TransferType(ByVal New_TransferType As Transfers)
    m_TransferType = New_TransferType
    PropertyChanged "TransferType"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
'**********************************************************
'* This is called to get the control to load the data of  *
'* the directory into a collection which can be accessed  *
'* from within the application using the control.         *
'**********************************************************
Public Function GetDirectoryListing(sFilter As String)
  On Error GoTo errHandler
  Dim dt As WIN32_FIND_DATA
  Dim hFile As Long, sFile
  Dim x As Long
  Dim sFilename As String
  
  Clear
  hFile = FtpFindFirstFile(hConnect, sFilter, dt, INTERNET_FLAG_RELOAD, INTERNET_FLAG_NO_CACHE_WRITE)
  sFile = 1
  Do Until sFile = 0
    If ((dt.dwFileAttributes And vbDirectory)) Then
      sFilename = Left(dt.cFileName, InStr(1, dt.cFileName, String(1, 0), vbBinaryCompare) - 1)
      Add dt.dwFileAttributes, Win32ToVbTime(dt.ftCreationTime), Win32ToVbTime(dt.ftLastAccessTime), Win32ToVbTime(dt.ftLastWriteTime), dt.nFileSizeLow, sFilename
    End If
    sFile = InternetFindNextFile(hFile, dt)
  Loop
  InternetCloseHandle (hFile)
  InternetCloseHandle (sFile)
  hFile = FtpFindFirstFile(hConnect, sFilter, dt, INTERNET_FLAG_RELOAD, INTERNET_FLAG_NO_CACHE_WRITE)
  sFile = 1
  Do Until sFile = 0
    If ((dt.dwFileAttributes And Not vbDirectory)) Then
      sFilename = Left(dt.cFileName, InStr(1, dt.cFileName, String(1, 0), vbBinaryCompare) - 1)
      Add dt.dwFileAttributes, Win32ToVbTime(dt.ftCreationTime), Win32ToVbTime(dt.ftLastAccessTime), Win32ToVbTime(dt.ftLastWriteTime), dt.nFileSizeLow, sFilename
    End If
    sFile = InternetFindNextFile(hFile, dt)
  Loop
  InternetCloseHandle (hFile)
  InternetCloseHandle (sFile)
  Exit Function
errHandler:
  FTPError Err.LastDllError, "GetDirectoryListing"
  Resume Next
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=23
'**********************************************************
'* This is just sets directory as a function to mCol      *
'* can be accessed from within the program using the con  *
'**********************************************************
Public Function Directory() As Collection
  Set Directory = mCol
End Function

'**********************************************************
'* This function does nothing more than add the data to   *
'* each item in the collection which is used to store the *
'* directory structure.                                   *
'**********************************************************
Private Function Add(lAttrib As Long, dtCreationTime As Date, dtLastAccessTime As Date, dtLastWriteTime As Date, lFileSize As Long, sFilename As String)
   Dim newItem As cDirItem
   Set newItem = New cDirItem
   With newItem
      .Archive = (lAttrib And FILE_ATTRIBUTE_ARCHIVE)
      .Compressed = (lAttrib And FILE_ATTRIBUTE_COMPRESSED)
      .Directory = (lAttrib And FILE_ATTRIBUTE_DIRECTORY)
      .Hidden = (lAttrib And FILE_ATTRIBUTE_HIDDEN)
      .Normal = (lAttrib And FILE_ATTRIBUTE_NORMAL)
      .Offline = (lAttrib And FILE_ATTRIBUTE_OFFLINE)
      .ReadOnly = (lAttrib And FILE_ATTRIBUTE_READONLY)
      .System = (lAttrib And FILE_ATTRIBUTE_SYSTEM)
      .Temporary = (lAttrib And FILE_ATTRIBUTE_TEMPORARY)
      .CreationTime = dtCreationTime
      .LastAccessTime = dtLastAccessTime
      .LastWriteTime = dtLastWriteTime
      .FileSize = lFileSize
      .Filename = sFilename
   End With
   mCol.Add newItem
   Set newItem = Nothing
End Function

Public Sub SetDirectory(Direc As String)
  On Error GoTo errHandler
  Dim hDirec As Boolean
  hDirec = FtpSetCurrentDirectory(hConnect, Direc)
  If hDirec = False Then
    FTPError Err.LastDllError, "SetDirectory"
    Resume Next
  End If
  Exit Sub
errHandler:
  FTPError Err.LastDllError, "SetDirectory"
  Resume Next
End Sub

'**********************************************************
'* This function takes each recieved error and enterprets *
'* them through a simple select case setup and raises the *
'* GetError Event which from within the users app can be  *
'* used to display errors.                                *
'**********************************************************
Private Sub FTPError(ByVal dwError As Long, ByRef szFunc As String)
    Dim dwRet As Long
    Dim dwTemp As Long
    Dim szString As String * 2048, szErrorMessage As String
    dwRet = FormatMessage(FORMAT_MESSAGE_FROM_HMODULE, _
                      GetModuleHandle("wininet.dll"), dwError, 0, _
                      szString, 256, 0)
    szErrorMessage = szFunc & " error code: " & dwError & " Message: " & szString
    If (dwError = 12003) Then
        ' Extended error information was returned
        dwRet = InternetGetLastResponseInfo(dwTemp, szString, 2048)
        szErrorMessage = szString
    End If
    Select Case dwError
      Case 12014
        szErrorMessage = "Login Incorrect"
      Case 12099
        szErrorMessage = "File To Save To Not Set, But Open as String is False"
    End Select
    RaiseEvent GetError(szErrorMessage, szFunc, dwError)
End Sub



'*****************************************************************************
'* This gets the current ftp directory.                                      *
'*****************************************************************************
Public Function GetFTPDirectory() As String
    Dim szDir As String
    szDir = String(1024, Chr$(0))
    If (FtpGetCurrentDirectory(hConnect, szDir, 1024) = False) Then
        Exit Function
    Else
        GetFTPDirectory = Left(szDir, InStr(1, szDir, String(1, 0), vbBinaryCompare) - 1)
    End If
End Function

'*****************************************************************************
'* This function has no purpose to the app using the control so it's just a  *
'* private decleration. When the directory is being checked it needs to have *
'* it's collection cleaned up first.                                         *
'*****************************************************************************
Private Function Clear()
  Dim i As Long
  For i = mCol.Count To 1 Step -1
    mCol.Remove i
  Next
End Function

Public Sub About()
Attribute About.VB_Description = "About the Control and Author"
Attribute About.VB_UserMemId = -552
  frmAbout.Show vbModal, Me
  Unload frmAbout
  Set frmAbout = Nothing
End Sub

Public Sub RenameSelection(OldName As String, NewName As String)
  On Error GoTo errHandler:
  Dim hFile As Long
  hFile = FtpRenameFile(hConnect, OldName, NewName)
  If hFile <> 0 Then
    RaiseEvent Message(MRENAMED)
    Exit Sub
  End If
errHandler:
  FTPError Err.LastDllError, "RenameSelection"
End Sub

Public Sub DeleteSelection(Filename As String)
  'First we have to determine if it's a directory.
  Dim hDel As Long, dt As WIN32_FIND_DATA
  dt = ReturnFile(Filename)
  If (dt.dwFileAttributes And vbDirectory) Then
    hDel = FtpRemoveDirectory(hConnect, dt.cFileName)
  Else
    hDel = FtpDeleteFile(hConnect, dt.cFileName)
  End If
  If hDel = 0 Then
    FTPError Err.LastDllError, "DeleteSelection"
    Exit Sub
  End If
  RaiseEvent Message(MDELETED)
End Sub

Public Function DownloadFile(Filename As String, Optional FileToSaveTo As String, Optional DoNotSave As Boolean) As String
  On Error GoTo errHandler
  'DoNotSave defaults to false. The reason for the optional with the
  'FileToSaveTo, and DoNotSave is that this ocx allows the user of it to download it
  'simply into a string or to save it to a file. This allows someone to use this control
  'to emulate such editors which have built in ftp and are able to open the file directly
  'into the editor.
    
  Dim sBuffer As String, sTotal As String, hFile As Long, fFile As Integer
  Dim totalBytes As Long, currBytes As Long, Ret As Long
  
  If DoNotSave = False And FileToSaveTo = "" Then
    FTPError 12099, "DownloadFile"
    Exit Function
  End If
  If Filename = "" Then
    RaiseEvent GetError("No file specified", "DownloadFile", 12699)
    Exit Function
  End If
  totalBytes = ReturnSize(Filename)
  'If totalBytes = 0 Then
  '  RaiseEvent GetError("File is empty.", "DownloadFile", 12199)
  '  Exit Function
  'End If
  
  hFile = FtpOpenFile(hConnect, Filename, GENERIC_READ, m_TransferType, 0)
  If hFile = 0 Then
    FTPError Err.LastDllError, "DownloadFile"
    Exit Function
  End If
  sBuffer = Space(sReadBuffer)
  
  Do
    If InternetReadFile(hFile, sBuffer, sReadBuffer, Ret) = 0 Then
      FTPError Err.LastDllError, "DownloadFile"
      Exit Function
    End If
    currBytes = currBytes + Ret
    If Ret <> sReadBuffer Then
      sBuffer = Left$(sBuffer, Ret)
    End If
    sTotal = sTotal + sBuffer
    If currBytes > totalBytes Then totalBytes = currBytes
    DoEvents
    RaiseEvent Progress(totalBytes, currBytes)
  Loop Until Ret <> sReadBuffer
  
  If DoNotSave = True Then
    DownloadFile = sTotal
  Else
    fFile = FreeFile()
    If Dir(FileToSaveTo) <> "" Then Kill FileToSaveTo
    Open FileToSaveTo For Binary As #fFile
      Put #fFile, , sTotal
    Close #fFile
  End If
  InternetCloseHandle hFile
  RaiseEvent Message(MDOWNLOADED)
  Exit Function
errHandler:
  FTPError Err.LastDllError, "DownloadFile"
End Function

Public Function UploadFile(FileToWrite As String, Optional FileToSave As String, Optional DataTosave As String)
  'With this one if you use datatosave and filetosave it will default to datatosave
  'Datatosave is just inputting a string. IE "Hello". FileToSave is to input a file.

  If FileToSave = "" And DataTosave = "" Then
    RaiseEvent GetError("Nothing was inputed to save", "UploadFile", 12299)
    Exit Function
  End If
  If DataTosave = "" And FileToSave <> "" Then
    If Dir(FileToSave) = "" Then
      RaiseEvent GetError("File does not exist", "UploadFile", 12399)
      Exit Function
    End If
    UploadAsFile FileToSave, FileToWrite
  Else
    UploadAsString FileToWrite, DataTosave
  End If
    
End Function

Private Sub UploadAsFile(File1 As String, File2 As String)
  Dim nFileLen As Long, nRet As Long, nTotFileLen As Long
  Dim sBuffer As String * 1024, SentBytes As Long, sAllBytes As Long
  Dim hFile As Long, fFile As Integer
  SentBytes = 0
  nFileLen = 0
  hFile = FtpOpenFile(hConnect, File2, GENERIC_WRITE, m_TransferType, 0)
  If hFile = 0 Then
    FTPError Err.LastDllError, "UploadAsFile"
    Exit Sub
  End If
  fFile = FreeFile()
  Open File1 For Binary As #fFile
  nTotFileLen = LOF(fFile)
  Do
    Get #fFile, , sBuffer
    If nFileLen < nTotFileLen - sReadBuffer Then
      If InternetWriteFile(hFile, sBuffer, sReadBuffer, nRet) = 0 Then
        FTPError Err.LastDllError, "UploadAsFile"
        Exit Do
      End If
      SentBytes = SentBytes + sReadBuffer
      sAllBytes = sAllBytes + sReadBuffer
      nFileLen = nFileLen + sReadBuffer
    Else
      If InternetWriteFile(hFile, sBuffer, nTotFileLen - nFileLen, nRet) = 0 Then
        FTPError Err.LastDllError, "UploadAsFile"
        Exit Do
      End If
      SentBytes = SentBytes + (nTotFileLen - nFileLen)
      sAllBytes = sAllBytes + (nTotFileLen - nFileLen)
      nFileLen = nTotFileLen
    End If
  Loop Until nFileLen >= nTotFileLen
  Close #fFile
  InternetCloseHandle (hFile)
  RaiseEvent Message(MUPLOADED)
End Sub

Private Sub UploadAsString(File1 As String, Data As String)
  Dim hFile As Long, sizeLeft As Long, sBuffer As String, Ret As Long
  Dim SaveString As String, currBytes As Long, totalBytes As Long
  SaveString = Data
  currBytes = 0
  totalBytes = Len(Data)
  hFile = FtpOpenFile(hConnect, File1, GENERIC_WRITE, m_TransferType, 0)
  Do
    If hFile = 0 Then
      FTPError Err.LastDllError, "UploadAsFile"
      Exit Sub
    End If
    If Len(SaveString) >= sReadBuffer Then
      sBuffer = Left$(SaveString, sReadBuffer)
      SaveString = Mid(SaveString, sReadBuffer)
    Else
      sBuffer = Left$(SaveString, Len(SaveString))
      SaveString = ""
    End If
    sizeLeft = Len(sBuffer)
    If sizeLeft = sReadBuffer Then
      If InternetWriteFile(hFile, sBuffer, sReadBuffer, Ret) = 0 Then
        RaiseEvent GetError("Error placing file", "Uploadfile", 12599)
        Exit Do
      End If
    Else
      If InternetWriteFile(hFile, sBuffer, sizeLeft, Ret) = 0 Then
        RaiseEvent GetError("Error placing file", "Uploadfile", 12599)
        Exit Do
      End If
    End If
    currBytes = currBytes + Ret
    If currBytes > totalBytes Then totalBytes = currBytes
    DoEvents
    RaiseEvent Progress(totalBytes, currBytes)
  Loop Until currBytes >= totalBytes
  InternetCloseHandle (hFile)
  RaiseEvent Message(MUPLOADED)
End Sub

Public Sub Disconnect()
  InternetCloseHandle hSession
  InternetCloseHandle hConnect
  hSession = 0: hConnect = 0
  RaiseEvent Message(MDISCONNECTED)
End Sub

Public Sub MakeDir(Direc As String)
  Dim hFile As Long
  If Direc = "" Then
    RaiseEvent GetError("No Directory Entered", "MakeDir", 12799)
    Exit Sub
  End If
  hFile = FtpCreateDirectory(hConnect, Direc)
  If hFile = 0 Then
    FTPError Err.LastDllError, "MakeDir"
    InternetCloseHandle hFile
    Exit Sub
  End If
  RaiseEvent Message(MDIRCREATED)
End Sub
