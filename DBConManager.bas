Attribute VB_Name = "DBConManager"
'---------------------------------------------------------
'使用说明：
'    1.建立数据库连接

'    将本VB模块复制到工商企业档案影像管理各模块的源码目录下，
'    在VB工程中添加本模块?使用时调用GetADOConnection函数即可，
'    创建ADO数据连接失败时，返回nothing。
'
'    2.取函数的不同形式
'    对于取字段长度函数的处理，用GetFunctionName函数。
'---------------------------------------------------------


Public gDatabaseType As Long
Public gServerName As String
Public gUID As String
Public gPWD As String
Public gFontName As String
Public gFontSize As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Public Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, _
    ByVal nSize As Long, ByVal lpFileName As String) As Long
    

Public Function GetConfigFileString(ByVal FilePath As String, ByVal Section As String, ByVal KeyName As String) As String
    Dim lpReturnedString As String * 100
    Dim nSize As Long
    Dim lPo As Long
    Dim mFilePath, mSection, mKeyName As String
    mSection = Trim(Section)
    mKeyName = Trim(KeyName)
    mFilePath = Trim(FilePath)
    nSize = Len(lpReturnedString)
    lPo = GetPrivateProfileString(mSection, mKeyName, "", lpReturnedString, nSize, mFilePath)
    GetConfigFileString = Mid(lpReturnedString, 1, lPo)
End Function

'----------------------------------------------------------
'名称：GetADOConnection(创建ADO连接)
'功能：根据指定的参数数据库的ADO连接
'参数：
'       DBType  数据类型：0=SQLServer；1=Sybase；2=Oracle
'       Server  数据库服务器名称
'       DBName  要连接的数据库名称
'       UID     数据库用户名
'       PWD     数据库用户密码
'      pReturnADODBConnection      --待返回的ADODB数据库连接
'      Errorstring----- 数据库连接创建失败的可能原因描述
'返回：数据连接创建成功 True
'      数据连接创建失败 False
'----------------------------------------------------------

Public Function GetADOConnection(ByVal DBType As Integer, ByVal Server As String, ByVal SID As String, ByVal UID As String, ByVal PWD As String, ByRef pReturnADODBConnection As ADODB.Connection, Optional ByVal DBName As String = "", Optional ByVal ServerPort As String = "", Optional ByRef Errorstring As String = "") As Boolean
    Dim ConStr As String
    Dim TempCon As ADODB.Connection
    Dim Flag As Boolean
    Dim ErrorPosition As Long
    
    On Error GoTo ErrorHandleOK
    Flag = True
    ErrorPosition = 1

    Server = Trim(Server)
    ServerPort = Trim(ServerPort)
    DBName = Trim(DBName)
    UID = Trim(UID)
    PWD = Trim(PWD)
    
    If DBType <> 2 Then
        If Server = "" Or DBName = "" Or UID = "" Then
            Flag = False
            Errorstring = "参数非法为空：用于连接数据库的参数中存在非法为空的参数(Server=" & Server & " ,ServerPort=" & ServerPort & " ,UID=" & UID & " ,DBName=" & DBName & ")"
        End If
    Else
        If Server = "" Or UID = "" Then
            Flag = False
            Errorstring = "参数非法为空：用于连接数据库的参数中存在非法为空的参数(Server=" & Server & " ,ServerPort=" & ServerPort & " ,UID=" & UID & ")"
        End If
    End If
    If Flag = True Then
        ErrorPosition = 2
        Select Case DBType
            Case 0
                'ConStr = "Provider=sqloledb;server=" & Server & ";uid=" & UID & ";pwd=" & PWD & ";database=" & DBName
                
                'V3.5.11--连接SQL SERVER的标准OLEDB写法，增加端口号，如果在连接串中添加了缺省档案号1433（曾遇到过奇怪连接不上的现象，去掉就能连接成功）
                If ServerPort = "" Or ServerPort = "1433" Then
                    ConStr = "Provider=sqloledb;Data Source=" & Server & ";Initial Catalog=" & DBName & ";User Id=" & UID & ";Password=" & PWD & ";"
                Else
                    ConStr = "Provider=sqloledb;Data Source=" & Server & "," & ServerPort & ";Initial Catalog=" & DBName & ";User Id=" & UID & ";Password=" & PWD & ";"
                End If

            Case 1
                'ConStr = "Provider=Sybase.ASEOLEDBProvider.2;Initial Catalog=" & DBName & ";Password=" & PWD & ";User ID=" & UID & ";Persist Security Info=True;Server Name=" & Server & "," & ServerPort
                
                'V3.5.8--Persist Security Info表示是否保存密码信息，一般不建议设置为True，据说ADO缺省为true，ADO.NET缺省为false
                ConStr = "Provider=Sybase.ASEOLEDBProvider;Server Name=" & Server & "," & IIf(Len(ServerPort) = 0, "5000", ServerPort) & ";Initial Catalog=" & DBName & ";User Id=" & UID & ";Password=" & PWD & ";"

            Case 2
                
                'ConStr = "Provider=MSDAORA.1;Password=" & PWD & ";User ID=" & UID & ";Data Source=" & Server & ";Persist Security Info=True;Extended Properties=" & DBName
                'ConStr = "Provider=msdaora.1;Data Source=" & Server & ";User Id=" & UID & ";Password=" & PWD & ";"
                
                
                'ConStr = "Provider=OraOLEDB.Oracle.1;Password=" & PWD & ";Persist Security Info=True;User ID=" & UID & ";Data Source=" & Server & ";Extended Properties=" & DBName
                'ConStr = "Provider=OraOLEDB.Oracle;Data Source=" & Server & ";User Id=" & UID & ";Password=" & PWD & ";"

                'V3.5.8--采用这种连接字符串，要求先安装有oracle提供的ODAC组件（包含了oledb驱动），可以单独下载安装ODAC组件或通过oracle客户端管理程序进行安装，并且无需配置net configuration，可以直接连接上指定的服务
                ConStr = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(CID=GTU_APP)(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" & Server & ")(PORT=" & IIf(Len(ServerPort) = 0, "1521", ServerPort) & ")))(CONNECT_DATA=(SID=" & SID & ")(SERVER=DEDICATED)));User Id=" & UID & ";Password=" & PWD & ";"
            
            'V3.6.0--支持DB2
            Case 3
                'Provider=IBMDADB2;Database=myDataBase;Hostname=myServerAddress;Protocol=TCPIP;Port=50000;Uid=myUsername;Pwd=myPassword;
                'ConStr = "Provider=IBMDADB22;Hostname=" & Server & ";Protocol=TCPIP;Port=" & IIf(Len(ServerPort) = 0, "50000", ServerPort) & ";Database=" & DBName & ";Uid=" & UID & ";Pwd=" & PWD & ";"
                
                'V3.6.0--该字符串要安装DB2 Run-Time Client Lite
                ConStr = "Driver={IBM DB2 ODBC DRIVER};Hostname=" & Server & ";Protocol=TCPIP;Port=" & IIf(Len(ServerPort) = 0, "50000", ServerPort) & ";Database=" & DBName & ";Uid=" & UID & ";Pwd=" & PWD & ";"
                
        End Select
    End If
    
    If Flag = True Then
        ErrorPosition = 3
        Set TempCon = New ADODB.Connection
        TempCon.ConnectionTimeout = 3000
        TempCon.CommandTimeout = 3000
        ErrorPosition = 4
        TempCon.CursorLocation = adUseClient
        ErrorPosition = 5
        TempCon.ConnectionString = ConStr
        TempCon.Open
    End If
    
CreateOK:
    Set pReturnADODBConnection = TempCon
    Set TempCon = Nothing
    
    GetADOConnection = Flag
    If Flag = False Then
        Errorstring = "DBConManager.GetADOConnection--->" & Errorstring
    End If
    Exit Function
    
ErrorHandleOK:
    Flag = False
    Errorstring = "Err.Number:" & Err.Number & Space(5) & "Err.Position:" & ErrorPosition & Space(5) & "Err.Description:" & Err.Description & Space(5) & "Err.Source:" & Err.Source
    Resume CreateOK

End Function


'功能说明：从字符串的末尾截去指定的字符串
'参数说明：OperateString--欲进行操作的字符串，MatchStringToBeTrimed--欲截去的字符串
'返回值说明：截去指定字符后的字符串
'例如：12345678900000  --->  TrimCharacters("12345678900000","0")=123456789
'            12345678001001  --->  TrimCharacters("12345678001001","001")=12345678
'            12345678001001  --->  TrimCharacters("12345678001001","1")=1234567800100
Public Function TrimCharacters(ByVal OperateString As String, ByVal MatchStringToBeTrimed As String) As String
    Dim FindPosition As Long
    Dim LenOfOperateStr As Long
    Dim TempStr As String
   
    TempStr = Trim(OperateString)
    LenOfOperateStr = Len(OperateString)
    
    'V3.5.8--如果字符串长度为0则不必进行截取了
    If LenOfOperateStr > 0 Then
        FindPosition = InStrRev(OperateString, MatchStringToBeTrimed, LenOfOperateStr)
        
        If (FindPosition <> 0) And (FindPosition + Len(MatchStringToBeTrimed) - 1 = LenOfOperateStr) Then
            TempStr = TrimCharacters(Left(TempStr, FindPosition - 1), MatchStringToBeTrimed)
        End If
    End If
    
    TrimCharacters = Trim(TempStr)
End Function
