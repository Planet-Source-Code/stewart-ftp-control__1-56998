VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDirItem"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************************************
'* This class module has one purpose to it. It's used for the collection.    *
'* The class data is set then inserted into the collection and finally       *
'* to clean up memory it's destroyed.                                        *
'*****************************************************************************

Option Explicit

Public ReadOnly As Boolean
Public Hidden As Boolean
Public System As Boolean
Public Directory As Boolean
Public Archive As Boolean
Public Normal As Boolean
Public Temporary As Boolean
Public Compressed As Boolean
Public Offline As Boolean

Public CreationTime As Date
Public LastAccessTime As Date
Public LastWriteTime As Date

Public FileSize As Long
Public Filename As String

