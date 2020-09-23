Attribute VB_Name = "wininet"
'*****************************************************************************
'* This module is the standard wininet.bas. I don't know who exactly wrote   *
'* it, but I did not so whoever wrote it deserves the credit. I have made a  *
'* few minor changes to it, to include a couple extra api calls and a few    *
'* new variables.                                                            *
'*****************************************************************************
'* This module contains a lot of api calls. Many of them I never use in the  *
'* control but they can be used for different things. It also contains a few *
'* variables and constants.                                                  *
'*****************************************************************************

Option Explicit
Public Const sReadBuffer = 1024

Public hSession As Long
Public hConnect As Long

Declare Function GetTickCount& Lib "kernel32" ()


Public Const FILE_ATTRIBUTE_OFFLINE = &H1000

Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal pub_lngInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Sub InternetSetStatusCallback Lib "wininet.dll" (ByVal pub_lngInternetSession As Long, ByVal lpfnInternetCallback As Long)
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal pub_lngInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, sOptional As Any, ByVal lOptionalLength As Long) As Integer
Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function InternetWriteFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToWrite As Long, dwNumberOfBytesWritten As Long) As Integer
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Public Declare Function InternetQueryDataAvailable Lib "wininet.dll" (ByVal hInet As Long, dwAvail As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Public Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (ByRef lpdwError As Long, ByVal lpszErrorBuffer As String, ByRef lpdwErrorBufferLength As Long) As Boolean
Public Declare Function InternetTimeToSystemTime Lib "wininet.dll" (ByVal lpszTime As String, ByRef pst As SYSTEMTIME, ByVal dwReserved As Long) As Long
         
Public Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
Public Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String, ByVal fdwAccess As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Long, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Long
Public Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Long
Public Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Long
Public Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Long
Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Public Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Long
                  
' Use registry access settings.
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1

' Number of the TCP/IP port on the server to connect to.
Public Const INTERNET_INVALID_PORT_NUMBER = 0
Public Const INTERNET_DEFAULT_FTP_PORT = 21
Public Const INTERNET_DEFAULT_GOPHER_PORT = 70
Public Const INTERNET_DEFAULT_HTTP_PORT = 80
Public Const INTERNET_DEFAULT_HTTPS_PORT = 443
Public Const INTERNET_DEFAULT_SOCKS_PORT = 1080

' Type of service to access.
Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_SERVICE_GOPHER = 2
Public Const INTERNET_SERVICE_HTTP = 3

' Brings the data across the wire even if it locally cached.
Public Const INTERNET_FLAG_RELOAD = &H80000000
Public Const ERROR_NO_MORE_FILES = 18

Public Const FTP_TRANSFER_TYPE_UNKNOWN As Long = &H0 '0x00000000
Public Const FTP_TRANSFER_TYPE_ASCII As Long = &H1 '0x00000001
Public Const FTP_TRANSFER_TYPE_BINARY  As Long = &H2 '0x00000002

' The possible values for the lInfoLevel parameter include:
Public Const HTTP_QUERY_CONTENT_TYPE = 1
Public Const HTTP_QUERY_CONTENT_LENGTH = 5
Public Const HTTP_QUERY_EXPIRES = 10
Public Const HTTP_QUERY_LAST_MODIFIED = 11
Public Const HTTP_QUERY_PRAGMA = 17
Public Const HTTP_QUERY_VERSION = 18
Public Const HTTP_QUERY_STATUS_CODE = 19
Public Const HTTP_QUERY_STATUS_TEXT = 20
Public Const HTTP_QUERY_RAW_HEADERS = 21
Public Const HTTP_QUERY_RAW_HEADERS_CRLF = 22
Public Const HTTP_QUERY_FORWARDED = 30
Public Const HTTP_QUERY_SERVER = 37
Public Const HTTP_QUERY_USER_AGENT = 39
Public Const HTTP_QUERY_SET_COOKIE = 43
Public Const HTTP_QUERY_REQUEST_METHOD = 45

' Add this flag to the about flags to get request header.
Public Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000

' flags for InternetOpenUrl
Public Const INTERNET_FLAG_RAW_DATA = &H40000000
Public Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000
Public Const INTERNET_FLAG_TRANSFER_ASCII = &H1&
Public Const INTERNET_FLAG_TRANSFER_BINARY = &H2&

' flags for InternetOpen
Public Const INTERNET_FLAG_ASYNC = &H10000000
Public Const INTERNET_FLAG_PASSIVE = &H8000000
Public Const INTERNET_FLAG_DONT_CACHE = &H4000000
Public Const INTERNET_FLAG_MAKE_PERSISTENT = &H2000000
Public Const INTERNET_FLAG_OFFLINE = &H1000000

Public Type INTERNET_ASYNC_RESULT
    dwResult As Long
    dwError As Long
End Type

Public Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000     ' don't write this item to the cache
Public Const INTERNET_STATUS_RESOLVING_NAME = 10
Public Const INTERNET_STATUS_NAME_RESOLVED = 11
Public Const INTERNET_STATUS_CONNECTING_TO_SERVER = 20
Public Const INTERNET_STATUS_CONNECTED_TO_SERVER = 21
Public Const INTERNET_STATUS_SENDING_REQUEST = 30
Public Const INTERNET_STATUS_REQUEST_SENT = 31
Public Const INTERNET_STATUS_RECEIVING_RESPONSE = 40
Public Const INTERNET_STATUS_RESPONSE_RECEIVED = 41
Public Const INTERNET_STATUS_CTL_RESPONSE_RECEIVED = 42
Public Const INTERNET_STATUS_PREFETCH = 43
Public Const INTERNET_STATUS_CLOSING_CONNECTION = 50
Public Const INTERNET_STATUS_CONNECTION_CLOSED = 51
Public Const INTERNET_STATUS_HANDLE_CREATED = 60
Public Const INTERNET_STATUS_HANDLE_CLOSING = 70
Public Const INTERNET_STATUS_REQUEST_COMPLETE = 100
Public Const INTERNET_STATUS_REDIRECT = 110
Public Const INTERNET_STATUS_STATE_CHANGE = 200
Public Const INTERNET_ERROR_BASE = 12000
Public Const ERROR_NOT_CONNECTED = 99

Public Const ERROR_INTERNET_OUT_OF_HANDLES = 12001
Public Const ERROR_INTERNET_TIMEOUT = 12002
Public Const ERROR_INTERNET_EXTENDED_ERROR = 12003
Public Const ERROR_INTERNET_INTERNAL_ERROR = 12004
Public Const ERROR_INTERNET_INVALID_URL = 12005
Public Const ERROR_INTERNET_UNRECOGNIZED_SCHEME = 12006
Public Const ERROR_INTERNET_NAME_NOT_RESOLVED = 12007
Public Const ERROR_INTERNET_PROTOCOL_NOT_FOUND = 12008
Public Const ERROR_INTERNET_INVALID_OPTION = 12009
Public Const ERROR_INTERNET_BAD_OPTION_LENGTH = 12010
Public Const ERROR_INTERNET_OPTION_NOT_SETTABLE = 12011
Public Const ERROR_INTERNET_SHUTDOWN = 12012
Public Const ERROR_INTERNET_INCORRECT_USER_NAME = 12013
Public Const ERROR_INTERNET_INCORRECT_PASSWORD = 12014
Public Const ERROR_INTERNET_LOGIN_FAILURE = 12015
Public Const ERROR_INTERNET_INVALID_OPERATION = 12016
Public Const ERROR_INTERNET_OPERATION_CANCELLED = 12017
Public Const ERROR_INTERNET_INCORRECT_HANDLE_TYPE = 12018
Public Const ERROR_INTERNET_INCORRECT_HANDLE_STATE = 12019
Public Const ERROR_INTERNET_NOT_PROXY_REQUEST = 12020
Public Const ERROR_INTERNET_REGISTRY_VALUE_NOT_FOUND = 12021
Public Const ERROR_INTERNET_BAD_REGISTRY_PARAMETER = 12022
Public Const ERROR_INTERNET_NO_DIRECT_ACCESS = 12023
Public Const ERROR_INTERNET_NO_CONTEXT = 12024
Public Const ERROR_INTERNET_NO_CALLBACK = 12025
Public Const ERROR_INTERNET_REQUEST_PENDING = 12026
Public Const ERROR_INTERNET_INCORRECT_FORMAT = 12027
Public Const ERROR_INTERNET_ITEM_NOT_FOUND = 12028
Public Const ERROR_INTERNET_CANNOT_CONNECT = 12029
Public Const ERROR_INTERNET_CONNECTION_ABORTED = 12030
Public Const ERROR_INTERNET_CONNECTION_RESET = 12031
Public Const ERROR_INTERNET_FORCE_RETRY = 12032
Public Const ERROR_INTERNET_INVALID_PROXY_REQUEST = 12033
Public Const ERROR_INTERNET_NEED_UI = 12034
Public Const NO_ERROR = 0 '  dderror
Public Const ERROR_INTERNET_HANDLE_EXISTS = 12036
Public Const ERROR_INTERNET_SEC_CERT_DATE_INVALID = 12037
Public Const ERROR_INTERNET_SEC_CERT_CN_INVALID = 12038
Public Const ERROR_INTERNET_HTTP_TO_HTTPS_ON_REDIR = 12039
Public Const ERROR_INTERNET_HTTPS_TO_HTTP_ON_REDIR = 12040
Public Const ERROR_INTERNET_MIXED_SECURITY = 12041
Public Const ERROR_INTERNET_CHG_POST_IS_NON_SECURE = 12042
Public Const ERROR_INTERNET_POST_IS_NON_SECURE = 12043
Public Const ERROR_INTERNET_CLIENT_AUTH_CERT_NEEDED = 12044
Public Const ERROR_INTERNET_INVALID_CA = 12045
Public Const ERROR_INTERNET_CLIENT_AUTH_NOT_SETUP = 12046
Public Const ERROR_INTERNET_ASYNC_THREAD_FAILED = 12047
Public Const ERROR_INTERNET_REDIRECT_SCHEME_CHANGE = 12048

'//
'// FTP API errors
'//

Public Const ERROR_FTP_TRANSFER_IN_PROGRESS = 12110
Public Const ERROR_FTP_DROPPED = 12111
Public Const Error_FTP_DIRECTORY_NOT_EMPTY = 550

'//
'// gopher API errors
'//

Public Const ERROR_GOPHER_PROTOCOL_ERROR = 12130
Public Const ERROR_GOPHER_NOT_FILE = 12131
Public Const ERROR_GOPHER_DATA_ERROR = 12132
Public Const ERROR_GOPHER_END_OF_DATA = 12133
Public Const ERROR_GOPHER_INVALID_LOCATOR = 12134
Public Const ERROR_GOPHER_INCORRECT_LOCATOR_TYPE = 12135
Public Const ERROR_GOPHER_NOT_GOPHER_PLUS = 12136
Public Const ERROR_GOPHER_ATTRIBUTE_NOT_FOUND = 12137
Public Const ERROR_GOPHER_UNKNOWN_LOCATOR = 12138

'//
'// HTTP API errors
'//

Public Const ERROR_HTTP_HEADER_NOT_FOUND = 12150
Public Const ERROR_HTTP_DOWNLEVEL_SERVER = 12151
Public Const ERROR_HTTP_INVALID_SERVER_RESPONSE = 12152
Public Const ERROR_HTTP_INVALID_HEADER = 12153
Public Const ERROR_HTTP_INVALID_QUERY_REQUEST = 12154
Public Const ERROR_HTTP_HEADER_ALREADY_EXISTS = 12155
Public Const ERROR_HTTP_REDIRECT_FAILED = 12156
Public Const ERROR_HTTP_NOT_REDIRECTED = 12160               '// BUGBUG

Public Const ERROR_INTERNET_SECURITY_CHANNEL_ERROR = 12157   '// BUGBUG
Public Const ERROR_INTERNET_UNABLE_TO_CACHE_FILE = 12158    ' // BUGBUG
Public Const ERROR_INTERNET_TCPIP_NOT_INSTALLED = 12159      '// BUGBUG

Public Const INTERNET_ERROR_LAST = 12159

Public Const MAX_PATH = 260
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_BEGIN = 0
Public Const FILE_CURRENT = 1
Public Const FILE_END = 2

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const OPEN_ALWAYS = 4

Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
Public Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type

' This originally was this:
'Public Type WIN32_FIND_DATA
'        dwFileAttributes As Long
'        ftCreationTime As FILETIME
'        ftLastAccessTime As FILETIME
'        ftLastWriteTime As FILETIME
'        nFileSizeHigh As Long
'        nFileSizeLow As Long
'        dwReserved0 As Long
'        dwReserved1 As Long
'        cFileName As String * MAX_PATH
'        cAlternate As String * 14
'End Type
'
'You can see the modified version. I modified
'it so I could convert it to date time so it's
'more easily read.

Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As Currency
        ftLastAccessTime As Currency
        ftLastWriteTime As Currency
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type


Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpfilename As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpfilename As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FileTimeToSystemTime& Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME)
Public Declare Function GetTimeFormat Lib "kernel32" Alias "GetTimeFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpTime As SYSTEMTIME, ByVal lpFormat As String, ByVal lpTimeStr As String, ByVal cchTime As Long) As Long
Public Declare Function GetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long


Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
                            (ByVal dwFlags As Long, lpSource As Any, _
                            ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
                            ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Public Const LANG_USER_DEFAULT = &H400&

Public Const SWP_DRAWFRAME = &H20
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As Any, lpLocalFileTime As Any) As Long

Public Const MESSAGE_CONNECT As Long = 0
Public Const MESSAGE_RENAME As Long = 1
Public Const MESSAGE_DELETE As Long = 2

Public Const RAS_MAXENTRYNAME As Integer = 256
Public Const RAS_MAXDEVICETYPE As Integer = 16
Public Const RAS_MAXDEVICENAME As Integer = 128
Public Const RAS_RASCONNSIZE As Integer = 412
 Type RASCONN
    dwSize As Long
    hRasConn As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
    szDeviceType(RAS_MAXDEVICETYPE) As Byte
    szDeviceName(RAS_MAXDEVICENAME) As Byte
End Type

Public Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" (udtRasConn As Any, lpcb As Long, lpcConnections As Long) As Long
Public Declare Function RasHangUp Lib "rasapi32.dll" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long
Public Const rDayZeroBias As Double = 109205#   ' Abs(CDbl(#01-01-1601#))
Public Const rMillisecondPerDay As Double = 10000000# * 60# * 60# * 24# / 10000#

