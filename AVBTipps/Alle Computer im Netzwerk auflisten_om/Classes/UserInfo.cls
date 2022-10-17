VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Type UserInfoExt
    Name         As String
    Password     As String
    PasswordAge  As String
    Privilege    As Long
    HomeDir      As String
    Comment      As String
    Flags        As Long
    NoChangePwd  As Boolean
    NoExpirePwd  As Boolean
    AccDisabled  As Boolean
    AccLocked    As Boolean
    ScriptPath   As String
    AuthFlags    As Long
    FullName     As String
    UserComment  As String
    Parms        As String
    Workstations As String
    LastLogon    As Date
    LastLogoff   As Date
    AcctExpires  As Date
    MaxStorage   As Long
    UnitsPerWeek As Long
    LogonHours(0 To 20) As Byte
    BadPwCount   As Long
    NumLogons    As Long
    LogonServer  As String
    CountryCode  As Long
    CodePage     As Long
    UserID       As Long
    PrimaryGroupID As Long
    Profile      As String
    HomeDirDrive As String
    PasswordExpired As Boolean
End Type

Private m_Name As String

Friend Sub New_(UserName As String)
    m_Name = UserName
End Sub

Public Property Get Name() As String
    Name = m_Name
End Property

