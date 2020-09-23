Attribute VB_Name = "modFunctions"
'*****************************************************************************
'* This module contains some functions to deal with ftp data. Fairly basic   *
'* stuff really.                                                             *
'*****************************************************************************
                                                                
'*****************************************************************************
'* This function will convert the time it recieves from the server to a date *
'* which is nice and easily read by a user. So you see like 10/10/1981       *
'* instead of a bunch of numbers :)                                          *
'*****************************************************************************
Function Win32ToVbTime(ft As Currency) As Date
    Dim ftl As Currency
    ' Call API to convert from UTC time to local time
    If FileTimeToLocalFileTime(ft, ftl) Then
        ' Local time is nanoseconds since 01-01-1601
        ' In Currency that comes out as milliseconds
        ' Divide by milliseconds per day to get days since 1601
        ' Subtract days from 1601 to 1899 to get VB Date equivalent
        Win32ToVbTime = CDate((ftl / rMillisecondPerDay) - rDayZeroBias)
    Else
        MsgBox Err.LastDllError
    End If
End Function
'*****************************************************************************
'* This function will return the filesize of a file online. Used to get the  *
'* filesize when downloading a file :)                                       *
'*****************************************************************************
Public Function ReturnSize(file As String) As Long
  Dim hFile As Long, dt As WIN32_FIND_DATA
  hFile = FtpFindFirstFile(hConnect, "*" & file, dt, INTERNET_FLAG_RELOAD, INTERNET_FLAG_NO_CACHE_WRITE)
  If hFile = 0 Or (dt.dwFileAttributes And vbDirectory) Then
    ReturnSize = 0
    Exit Function
  End If
  ReturnSize = dt.nFileSizeLow
  InternetCloseHandle hFile
End Function

Public Function ReturnFile(file As String) As WIN32_FIND_DATA
  Dim hFile As Long
  hFile = FtpFindFirstFile(hConnect, "*" & file, ReturnFile, INTERNET_FLAG_RELOAD, INTERNET_FLAG_NO_CACHE_WRITE)
  InternetCloseHandle hFile
End Function

Public Function GetSystemFileSize(file As String) As Long
  Dim hFile As Long, dt As WIN32_FIND_DATA
  hFile = FindFirstFile(file, dt)
  GetSystemFileSize = dt.nFileSizeLow
End Function
