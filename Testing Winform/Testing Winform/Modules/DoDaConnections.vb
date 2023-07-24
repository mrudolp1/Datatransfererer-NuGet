Imports System.Data.SqlClient
Imports System.Security.Principal
Imports Oracle.ManagedDataAccess.Client

Module DoDaSQL
    Public Declare Auto Function LogonUser Lib "advapi32.dll" (ByVal nToken As String, ByVal domain As String, ByVal wToken As String, ByVal lType As Integer, ByVal lProvider As Integer, ByRef Token As IntPtr) As Boolean
    Public Declare Auto Function CloseHandle Lib "kernel32.dll" (ByVal handle As IntPtr) As Boolean
    Public tokenHandle As New IntPtr(0)
    Public impersonatedUser As WindowsImpersonationContext
    Public newId As WindowsIdentity

    '<DebuggerStepThrough()>
    Public Function sqlLoader(ByVal sqlStr As String, ByVal sqlSrc As String, ByVal erNo As Integer) As Boolean
        dtClearer(sqlSrc)
        'Dim newId As New WindowsIdentity(tokenHandle)
        Using edsimpersonatedUser As WindowsImpersonationContext = EDSnewId.Impersonate()
            sqlCon = New SqlConnection(EDSdbActive)
            sqlCon.Open()

            Try
                da = New SqlDataAdapter(sqlStr, sqlCon)
                da.Fill(ds, sqlSrc)
                dt = ds.Tables(sqlSrc)
            Catch ex As Exception
                sqlCon.Close()
                Console.WriteLine(sqlStr)
                sendToast("Failure loading data:" & vbCrLf & ex.Message, "Error " & erNo)
                Return False
            End Try

            sqlCon.Close()
        End Using

        Return True
    End Function

    '<DebuggerStepThrough()>
    Public Function sqlSender(ByVal cmd As String, ByVal erNo As Integer) As Boolean
        Using edsimpersonatedUser As WindowsImpersonationContext = EDSnewId.Impersonate()
            sqlCon = New SqlConnection(EDSdbActive)
            Dim sqlCmd = New SqlCommand(cmd, sqlCon)
            sqlCon.Open()

            Try
                sqlCmd.ExecuteNonQuery()
            Catch ex As Exception
                sqlCon.Close()
                Console.WriteLine(cmd)
                sendToast("Error saving data:" & vbCrLf & ex.Message, "Error " & erNo)
                Return False
            End Try

            sqlCon.Close()
        End Using

        Return True
    End Function

End Module

Module DoDaORACLE

    Private Const ordsDataSource = "(DESCRIPTION =    (ADDRESS = (PROTOCOL = TCP)(HOST = prd-scan)(PORT = 1521))    (CONNECT_DATA =      (SERVER = DEDICATED)      (SERVICE_NAME = ordsprd_batch.crowncastle.com)    )  )"
    Private Const isitDataSource = "(DESCRIPTION =    (ADDRESS = (PROTOCOL = TCP)(HOST = prd-scan)(PORT = 1521))    (CONNECT_DATA =      (SERVICE_NAME = isitprd_utl.crowncastle.com)      (SERVER = DEDICATED)    )  )"
    Private Const odsDataSource = "(DESCRIPTION =    (ADDRESS = (PROTOCOL = TCP)(HOST = prd-scan)(PORT = 1521))    (CONNECT_DATA =      (SERVICE_NAME = odsprd_app.crowncastle.com)      (SERVER = DEDICATED)    )  )"
    'Private Const isitDataSource = "(DESCRIPTION =    (ADDRESS = (PROTOCOL = TCP)(HOST = uat-scan)(PORT = 1521))    (CONNECT_DATA =      (SERVICE_NAME = isituat_batch.crowncastle.com)      (SERVER = DEDICATED)    )  )"

    Private Const ntoken = "270:207:234:213:204:207:258"
    Private Const wtoken = "366:264:339:216:357:159:192:297:171:216"

    Public Function OracleLoader(ByVal sqlStr As String, ByVal sqlSrc As String, ByVal erNo As Integer, ByVal db As String) As Boolean

        Dim oraDatasource As String

        Select Case db
            Case "isit"
                oraDatasource = isitDataSource
            Case "ods"
                oraDatasource = odsDataSource
            Case Else
                'ORDS is the catch-all because it has links to the other DBs
                oraDatasource = ordsDataSource
        End Select

        dtClearer(sqlSrc)

        Dim sb As OracleConnectionStringBuilder = New OracleConnectionStringBuilder()

        sb.DataSource = oraDatasource
        sb.UserID = token(ntoken)
        sb.Password = token(wtoken)
        'By default pooling = true which means that oracle moves connections to an inactive pool when they are closed by the program.
        'This makes reconnecting faster but was causing issues with the connection idle_time being exceeded.
        sb.Pooling = False

        Dim bOraSuccess As Boolean = True

        Using oraCon As New OracleConnection(sb.ToString())
            Try
                Dim oDa = New OracleDataAdapter(sqlStr, oraCon)
                oDa.Fill(ds, sqlSrc)
                dt = ds.Tables(sqlSrc)
            Catch ex As Exception
                bOraSuccess = False
                Console.WriteLine(sqlStr)
                sendToast("Failure loading data:" & vbCrLf & ex.Message, "Error " & erNo)
            End Try
            oraCon.Close()
        End Using

        Return bOraSuccess

    End Function

End Module

'Module DoDaHTTPs
'    Dim cciConnection As Connection

'    Public Function httpPostIt(ByVal payload As String)
'        If (cciConnection Is Nothing) Then
'            cciConnection = New Connection(httpBaseLink)
'        End If

'        Dim myWebRequestResponse As HttpWebResponse
'        Try
'            myWebRequestResponse = cciConnection.ExecuteHttpPostRestApi(httpWOapi, payload)
'            DisplayWebResponse(myWebRequestResponse)
'            Return True
'        Catch ex As System.Net.WebException
'            sendToast(DisplayWebException(ex), "Error:  " & 5001)
'            Return False
'        Finally
'            CloseWebResponse(myWebRequestResponse)
'        End Try
'    End Function

'    Public Function httpGetIt(ByVal fullAPIlink As String)
'        If (cciConnection Is Nothing) Then
'            cciConnection = New Connection(httpBaseLink)
'        End If

'        Dim myWebRequestResponse As HttpWebResponse
'        Try
'            myWebRequestResponse = cciConnection.ExecuteHttpGetRestAPI(fullAPIlink)
'            DisplayWebResponse(myWebRequestResponse)
'            Return True
'        Catch ex As System.Net.WebException
'            DisplayWebException(ex)
'            Return False
'        Finally
'            CloseWebResponse(myWebRequestResponse)
'        End Try
'    End Function


'    Private Sub CloseWebResponse(myWebRequestResponse As HttpWebResponse)
'        If myWebRequestResponse IsNot Nothing Then
'            myWebRequestResponse.Close()
'        End If
'    End Sub

'    Private Function DisplayWebException(ex As WebException)
'        If ex.Status = WebExceptionStatus.ProtocolError Then
'            Dim errorResponseReader As New StreamReader(ex.Response.GetResponseStream())
'            Dim errorResponseContent As String = errorResponseReader.ReadToEnd
'            Return "STATUS: " + DirectCast(ex.Response, HttpWebResponse).StatusCode.ToString & vbCrLf & ex.Message & ": " & errorResponseContent
'        End If
'    End Function

'    Private Function DisplayWebResponse(myWebRequestResponse As HttpWebResponse)
'        Using myWebRequestResponseReader As New StreamReader(myWebRequestResponse.GetResponseStream())
'            Dim myWebRequestResponseContent = myWebRequestResponseReader.ReadToEnd
'            Return "STATUS: " + myWebRequestResponse.StatusCode.ToString & vbCrLf & myWebRequestResponseContent
'        End Using
'    End Function
'End Module