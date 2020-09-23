Attribute VB_Name = "Module1"
Option Explicit

'API Declarations & Constants
Public Const NERR_Success = 0&
Public Const NERR_Access_Denied = 5&
Public Const NERR_MoreData = 234&

Public Const SRV_TYPE_SERVER = &H2
Public Const SRV_TYPE_SQLSERVER = &H4
Public Const SRV_TYPE_NT_PDC = &H8
Public Const SRV_TYPE_NT_BDC = &H10
Public Const SRV_TYPE_PRINT = &H200
Public Const SRV_TYPE_NT = &H1000
Public Const SRV_TYPE_ALL = &HFFFF
Public Const SRV_TYPE_RAS = &H400

Public Const SHORT_LEVEL = 10&
Public Const EXTENDED_LEVEL = 3&

Public Const USER_ACC_NOPWD_CHANGE = 577&
Public Const USER_ACC_NOPWD_EXPIRE = 66049
Public Const USER_ACC_DISABLED = 515&
Public Const USER_ACC_LOCKED = 529&

Private Type SERVER_INFO_API
    PlatformId As Long
    ServerName As Long
    Type As Long
    VerMajor As Long
    VerMinor As Long
    Comment As Long
End Type

Private Type WKSTA_INFO_API
    PlatformId As Long
    ComputerName As Long
    LanGroup As Long
    VerMajor As Long
    VerMinor As Long
    LanRoot As Long
End Type

Type ServerInfo
    PlatformId As Long
    ServerName As String
    Type As Long
    VerMajor As Long
    VerMinor As Long
    Comment As String
    Platform As String
    ServerType As Integer
    LanGroup As String
    LanRoot As String
End Type

Type ListOfServer
    Init As Boolean
    LastErr As Long
    List() As ServerInfo
End Type

Private Type USER_INFO_EXT_API
    Name As Long
    Password As Long
    PasswordAge As Long
    Privilege As Long
    HomeDir As Long
    Comment As Long
    Flags As Long
    ScriptPath As Long
    AuthFlags As Long
    FullName As Long
    UserComment As Long
    Parms As Long
    Workstations As Long
    LastLogon As Long
    LastLogoff As Long
    AcctExpires As Long
    MaxStorage As Long
    UnitsPerWeek As Long
    LogonHours As Long
    BadPwCount As Long
    NumLogons As Long
    LogonServer As Long
    CountryCode As Long
    CodePage As Long
    UserID As Long
    PrimaryGroupID As Long
    Profile As Long
    HomeDirDrive As Long
    PasswordExpired As Long
End Type

Type UserInfoExt
    Name As String
    Password As String
    PasswordAge As String
    Privilege As Long
    HomeDir As String
    Comment As String
    Flags As Long
    NoChangePwd As Boolean
    NoExpirePwd As Boolean
    AccDisabled As Boolean
    AccLocked As Boolean
    ScriptPath As String
    AuthFlags As Long
    FullName As String
    UserComment As String
    Parms As String
    Workstations As String
    LastLogon As Date
    LastLogoff As Date
    AcctExpires As Date
    MaxStorage As Long
    UnitsPerWeek As Long
    LogonHours(0 To 20) As Byte
    BadPwCount As Long
    NumLogons As Long
    LogonServer As String
    CountryCode As Long
    CodePage As Long
    UserID As Long
    PrimaryGroupID As Long
    Profile As String
    HomeDirDrive As String
    PasswordExpired As Boolean
End Type

Type ListOfUserExt
    Init As Boolean
    LastErr As Long
    List() As UserInfoExt
End Type

Public Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" _
        (pTo As Any, _
         uFrom As Any, _
         ByVal lSize As Long)

Declare Function lstrlenW Lib "kernel32" _
        (ByVal lpString As Long) As Long

Declare Function NetApiBufferFree Lib "netapi32" _
        (ByVal lBuffer&) As Long

Declare Function NetGetDCName Lib "netapi32" _
        (lpServer As Any, lpDomain As Any, _
         vBuffer As Any) As Long

Declare Function NetServerEnum Lib "netapi32" _
        (lpServer As Any, ByVal lLevel As Long, vBuffer As Any, _
         lPreferedMaxLen As Long, lEntriesRead As Long, lTotalEntries As Long, _
         ByVal lServerType As Long, ByVal sDomain$, vResume As Any) As Long

Declare Function NetUserEnum Lib "netapi32" _
        (lpServer As Any, ByVal Level As Long, _
         ByVal Filter As Long, lpBuffer As Long, _
         ByVal PrefMaxLen As Long, lpEntriesRead As Long, _
         lpTotalEntries As Long, lpResumeHandle As Long) As Long

Public CurrentServer As String

Public Function EnumServer(lServerType As Long) As ListOfServer
    Dim nRet As Long, x As Integer, i As Integer
    Dim lRetCode As Long
    Dim tServerInfo As SERVER_INFO_API
    Dim lServerInfo As Long
    Dim lServerInfoPtr As Long
    Dim ServerInfo As ServerInfo
    Dim lPreferedMaxLen As Long
    Dim lEntriesRead As Long
    Dim lTotalEntries As Long
    Dim sDomain As String
    Dim vResume As Variant
    Dim yServer() As Byte
    Dim SrvList As ListOfServer
    
    yServer = MakeServerName(ByVal "")
    lPreferedMaxLen = 65536
    
    nRet = NERR_MoreData
    Do While (nRet = NERR_MoreData)
        
        'Call NetServerEnum to get a list of Servers
        nRet = NetServerEnum(yServer(0), 101, lServerInfo, _
                             lPreferedMaxLen, lEntriesRead, _
                             lTotalEntries, lServerType, _
                             sDomain, vResume)
        
        If (nRet <> NERR_Success And _
             nRet <> NERR_MoreData) Then
            SrvList.Init = False
            SrvList.LastErr = nRet
            NetError nRet
            Exit Do
        End If
        
        ' NetServerEnum Index is 1 based
        x = 1
        lServerInfoPtr = lServerInfo
        
        Do While x <= lTotalEntries
            
            CopyMem tServerInfo, ByVal lServerInfoPtr, Len(tServerInfo)
            
            ServerInfo.Comment = PointerToStringW(tServerInfo.Comment)
            ServerInfo.ServerName = PointerToStringW(tServerInfo.ServerName)
            ServerInfo.Type = tServerInfo.Type
            ServerInfo.PlatformId = tServerInfo.PlatformId
            ServerInfo.VerMajor = tServerInfo.VerMajor
            ServerInfo.VerMinor = tServerInfo.VerMinor
            
            i = i + 1
            ReDim Preserve SrvList.List(1 To i) As ServerInfo
            SrvList.List(i) = ServerInfo
            
            x = x + 1
            lServerInfoPtr = lServerInfoPtr + Len(tServerInfo)
            
        Loop
        
        lRetCode = NetApiBufferFree(lServerInfo)
        SrvList.Init = (x > 1)
        
    Loop
    
    EnumServer = SrvList
    
End Function

Public Function GetPDCName() As String
    Dim lpBuffer As Long, nRet As Long
    Dim yServer() As Byte
    Dim sLocal As String
    
    yServer = MakeServerName(ByVal "")
    
    nRet = NetGetDCName(yServer(0), yServer(0), lpBuffer)
    
    If nRet = 0 Then
        sLocal = PointerToStringW(lpBuffer)
    End If
    
    If lpBuffer Then Call NetApiBufferFree(lpBuffer)
    
    GetPDCName = sLocal
    
End Function

' Function Read User Information - for future development!
Public Function LongEnumUsers(Server As String) As ListOfUserExt
    Dim yServer() As Byte, lRetCode As Long
    Dim nRead As Long, nTotal As Long
    Dim nRet As Long, nResume As Long
    Dim PrefMaxLen As Long
    Dim i As Long, x As Long
    Dim lUserInfo As Long
    Dim lUserInfoPtr As Long
    Dim UserInfo As UserInfoExt
    Dim UserList As ListOfUserExt
    Dim tUserInfo As USER_INFO_EXT_API
    
    yServer = MakeServerName(ByVal Server)
    PrefMaxLen = 65536
    
    nRet = NERR_MoreData
    Do While (nRet = NERR_MoreData)
        nRet = NetUserEnum(yServer(0), EXTENDED_LEVEL, 2, _
                           lUserInfo, PrefMaxLen, nRead, _
                           nTotal, nResume)
        
        If (nRet <> NERR_Success And _
             nRet <> NERR_MoreData) Then
            UserList.Init = False
            UserList.LastErr = nRet
            NetError nRet
            Exit Do
        End If
        
        lUserInfoPtr = lUserInfo
        
        x = 1
        Do While x <= nRead
            
            CopyMem tUserInfo, ByVal lUserInfoPtr, Len(tUserInfo)
            
            UserInfo.Name = PointerToStringW(tUserInfo.Name)
            UserInfo.Password = PointerToStringW(tUserInfo.Password)
            UserInfo.PasswordAge = Format(tUserInfo.PasswordAge / 86400, "0.0")
            UserInfo.Privilege = tUserInfo.Privilege
            UserInfo.HomeDir = PointerToStringW(tUserInfo.HomeDir)
            UserInfo.Comment = PointerToStringW(tUserInfo.Comment)
            UserInfo.Flags = tUserInfo.Flags
            UserInfo.NoChangePwd = CBool((tUserInfo.Flags Or USER_ACC_NOPWD_CHANGE) = tUserInfo.Flags)
            UserInfo.NoExpirePwd = CBool((tUserInfo.Flags Or USER_ACC_NOPWD_EXPIRE) = tUserInfo.Flags)
            UserInfo.AccDisabled = CBool((tUserInfo.Flags Or USER_ACC_DISABLED) = tUserInfo.Flags)
            UserInfo.AccLocked = CBool((tUserInfo.Flags Or USER_ACC_LOCKED) = tUserInfo.Flags)
            UserInfo.ScriptPath = PointerToStringW(tUserInfo.ScriptPath)
            UserInfo.AuthFlags = tUserInfo.AuthFlags
            UserInfo.FullName = PointerToStringW(tUserInfo.FullName)
            UserInfo.UserComment = PointerToStringW(tUserInfo.UserComment)
            UserInfo.Parms = PointerToStringW(tUserInfo.Parms)
            UserInfo.Workstations = PointerToStringW(tUserInfo.Workstations)
            UserInfo.LastLogon = NetTimeToVbTime(tUserInfo.LastLogon)
            UserInfo.LastLogoff = NetTimeToVbTime(tUserInfo.LastLogoff)
            If tUserInfo.AcctExpires = -1& Then
                UserInfo.AcctExpires = NetTimeToVbTime(0)
            Else
                UserInfo.AcctExpires = NetTimeToVbTime(tUserInfo.AcctExpires)
            End If
            UserInfo.MaxStorage = tUserInfo.MaxStorage
            UserInfo.UnitsPerWeek = tUserInfo.UnitsPerWeek
            CopyMem UserInfo.LogonHours(0), ByVal tUserInfo.LogonHours, 21
            UserInfo.BadPwCount = tUserInfo.BadPwCount
            UserInfo.NumLogons = tUserInfo.NumLogons
            UserInfo.LogonServer = PointerToStringW(tUserInfo.LogonServer)
            UserInfo.CountryCode = tUserInfo.CountryCode
            UserInfo.CodePage = tUserInfo.CodePage
            UserInfo.UserID = tUserInfo.UserID
            UserInfo.PrimaryGroupID = tUserInfo.PrimaryGroupID
            UserInfo.Profile = PointerToStringW(tUserInfo.Profile)
            UserInfo.HomeDirDrive = PointerToStringW(tUserInfo.HomeDirDrive)
            UserInfo.PasswordExpired = CBool(tUserInfo.PasswordExpired)
            
            i = i + 1
            ReDim Preserve UserList.List(1 To i) As UserInfoExt
            UserList.List(i) = UserInfo
            x = x + 1
            
            lUserInfoPtr = lUserInfoPtr + Len(tUserInfo)
            
        Loop
        
        lRetCode = NetApiBufferFree(lUserInfo)
        UserList.Init = (x > 1)
        
    Loop
    
    LongEnumUsers = UserList
    
End Function

Public Function MakeServerName(ByVal ServerName As String)
    Dim yServer() As Byte
    
    If ServerName <> "" Then
        If InStr(1, ServerName, "\\") = 0 Then
            ServerName = "\\" & ServerName
        End If
    End If
    
    yServer = ServerName & vbNullChar
    MakeServerName = yServer
    
End Function

Public Function NetError(nErr As Long, Optional Ret) As String
    Dim Msg As String
    
    If IsMissing(Ret) Then Ret = False
    
    Select Case nErr
        Case 5
            Msg = "Access Denied!"
        Case 1722
            Msg = "Server not accessible!"
        Case 1326
            Msg = " Sie besitzen nicht die Berechtigungen dafür"
        Case Else
            Msg = "Error Nr. (" & nErr & ") !"
    End Select
    
    If Not Ret Then
        Beep
        MsgBox Msg, vbCritical, "Net Error"
    Else
        NetError = Msg
    End If
    
End Function

Public Function NetTimeToVbTime(NetDate As Long) As Double
    Const BaseDate# = 25569   'DateSerial(1970, 1, 1)
    Const SecsPerDay# = 86400
    Dim Tmp As Double
    
    Tmp = BaseDate + (CDbl(NetDate) / SecsPerDay)
    If Tmp <> BaseDate Then
        NetTimeToVbTime = Tmp
    End If
    
End Function

Public Function PointerToStringW(lpStringW As Long) As String
    Dim buffer() As Byte
    Dim nLen As Long
    
    If lpStringW Then
        nLen = lstrlenW(lpStringW) * 2
        If nLen Then
            ReDim buffer(0 To (nLen - 1)) As Byte
            CopyMem buffer(0), ByVal lpStringW, nLen
            PointerToStringW = buffer
        End If
    End If
End Function


