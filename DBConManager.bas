Attribute VB_Name = "DBConManager"
'---------------------------------------------------------
'ʹ��˵����
'    1.�������ݿ�����

'    ����VBģ�鸴�Ƶ�������ҵ����Ӱ������ģ���Դ��Ŀ¼�£�
'    ��VB��������ӱ�ģ��?ʹ��ʱ����GetADOConnection�������ɣ�
'    ����ADO��������ʧ��ʱ������nothing��
'
'    2.ȡ�����Ĳ�ͬ��ʽ
'    ����ȡ�ֶγ��Ⱥ����Ĵ�����GetFunctionName������
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
'���ƣ�GetADOConnection(����ADO����)
'���ܣ�����ָ���Ĳ������ݿ��ADO����
'������
'       DBType  �������ͣ�0=SQLServer��1=Sybase��2=Oracle
'       Server  ���ݿ����������
'       DBName  Ҫ���ӵ����ݿ�����
'       UID     ���ݿ��û���
'       PWD     ���ݿ��û�����
'      pReturnADODBConnection      --�����ص�ADODB���ݿ�����
'      Errorstring----- ���ݿ����Ӵ���ʧ�ܵĿ���ԭ������
'���أ��������Ӵ����ɹ� True
'      �������Ӵ���ʧ�� False
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
            Errorstring = "�����Ƿ�Ϊ�գ������������ݿ�Ĳ����д��ڷǷ�Ϊ�յĲ���(Server=" & Server & " ,ServerPort=" & ServerPort & " ,UID=" & UID & " ,DBName=" & DBName & ")"
        End If
    Else
        If Server = "" Or UID = "" Then
            Flag = False
            Errorstring = "�����Ƿ�Ϊ�գ������������ݿ�Ĳ����д��ڷǷ�Ϊ�յĲ���(Server=" & Server & " ,ServerPort=" & ServerPort & " ,UID=" & UID & ")"
        End If
    End If
    If Flag = True Then
        ErrorPosition = 2
        Select Case DBType
            Case 0
                'ConStr = "Provider=sqloledb;server=" & Server & ";uid=" & UID & ";pwd=" & PWD & ";database=" & DBName
                
                'V3.5.11--����SQL SERVER�ı�׼OLEDBд�������Ӷ˿ںţ���������Ӵ��������ȱʡ������1433����������������Ӳ��ϵ�����ȥ���������ӳɹ���
                If ServerPort = "" Or ServerPort = "1433" Then
                    ConStr = "Provider=sqloledb;Data Source=" & Server & ";Initial Catalog=" & DBName & ";User Id=" & UID & ";Password=" & PWD & ";"
                Else
                    ConStr = "Provider=sqloledb;Data Source=" & Server & "," & ServerPort & ";Initial Catalog=" & DBName & ";User Id=" & UID & ";Password=" & PWD & ";"
                End If

            Case 1
                'ConStr = "Provider=Sybase.ASEOLEDBProvider.2;Initial Catalog=" & DBName & ";Password=" & PWD & ";User ID=" & UID & ";Persist Security Info=True;Server Name=" & Server & "," & ServerPort
                
                'V3.5.8--Persist Security Info��ʾ�Ƿ񱣴�������Ϣ��һ�㲻��������ΪTrue����˵ADOȱʡΪtrue��ADO.NETȱʡΪfalse
                ConStr = "Provider=Sybase.ASEOLEDBProvider;Server Name=" & Server & "," & IIf(Len(ServerPort) = 0, "5000", ServerPort) & ";Initial Catalog=" & DBName & ";User Id=" & UID & ";Password=" & PWD & ";"

            Case 2
                
                'ConStr = "Provider=MSDAORA.1;Password=" & PWD & ";User ID=" & UID & ";Data Source=" & Server & ";Persist Security Info=True;Extended Properties=" & DBName
                'ConStr = "Provider=msdaora.1;Data Source=" & Server & ";User Id=" & UID & ";Password=" & PWD & ";"
                
                
                'ConStr = "Provider=OraOLEDB.Oracle.1;Password=" & PWD & ";Persist Security Info=True;User ID=" & UID & ";Data Source=" & Server & ";Extended Properties=" & DBName
                'ConStr = "Provider=OraOLEDB.Oracle;Data Source=" & Server & ";User Id=" & UID & ";Password=" & PWD & ";"

                'V3.5.8--�������������ַ�����Ҫ���Ȱ�װ��oracle�ṩ��ODAC�����������oledb�����������Ե������ذ�װODAC�����ͨ��oracle�ͻ��˹��������а�װ��������������net configuration������ֱ��������ָ���ķ���
                ConStr = "Provider=OraOLEDB.Oracle;Data Source=(DESCRIPTION=(CID=GTU_APP)(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" & Server & ")(PORT=" & IIf(Len(ServerPort) = 0, "1521", ServerPort) & ")))(CONNECT_DATA=(SID=" & SID & ")(SERVER=DEDICATED)));User Id=" & UID & ";Password=" & PWD & ";"
            
            'V3.6.0--֧��DB2
            Case 3
                'Provider=IBMDADB2;Database=myDataBase;Hostname=myServerAddress;Protocol=TCPIP;Port=50000;Uid=myUsername;Pwd=myPassword;
                'ConStr = "Provider=IBMDADB22;Hostname=" & Server & ";Protocol=TCPIP;Port=" & IIf(Len(ServerPort) = 0, "50000", ServerPort) & ";Database=" & DBName & ";Uid=" & UID & ";Pwd=" & PWD & ";"
                
                'V3.6.0--���ַ���Ҫ��װDB2 Run-Time Client Lite
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


'����˵�������ַ�����ĩβ��ȥָ�����ַ���
'����˵����OperateString--�����в������ַ�����MatchStringToBeTrimed--����ȥ���ַ���
'����ֵ˵������ȥָ���ַ�����ַ���
'���磺12345678900000  --->  TrimCharacters("12345678900000","0")=123456789
'            12345678001001  --->  TrimCharacters("12345678001001","001")=12345678
'            12345678001001  --->  TrimCharacters("12345678001001","1")=1234567800100
Public Function TrimCharacters(ByVal OperateString As String, ByVal MatchStringToBeTrimed As String) As String
    Dim FindPosition As Long
    Dim LenOfOperateStr As Long
    Dim TempStr As String
   
    TempStr = Trim(OperateString)
    LenOfOperateStr = Len(OperateString)
    
    'V3.5.8--����ַ�������Ϊ0�򲻱ؽ��н�ȡ��
    If LenOfOperateStr > 0 Then
        FindPosition = InStrRev(OperateString, MatchStringToBeTrimed, LenOfOperateStr)
        
        If (FindPosition <> 0) And (FindPosition + Len(MatchStringToBeTrimed) - 1 = LenOfOperateStr) Then
            TempStr = TrimCharacters(Left(TempStr, FindPosition - 1), MatchStringToBeTrimed)
        End If
    End If
    
    TrimCharacters = Trim(TempStr)
End Function
