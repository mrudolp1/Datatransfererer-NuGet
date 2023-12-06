Imports System.Data.SqlClient
Imports System.Security.Principal

'------------------------------------------
'------------------------------------------
'Module Name: IDoDeclare
'Purpose: Designed to declare all public variables used throughout the program
'------------------------------------------
'------------------------------------------

Public Module IDoDeclare

    'Current Highest SQL Error codes (Starts with 2000 & 1000)
    ''''''Loader --> 1001
    ''''''Sender --> 2002

    'Current Highest ORACLE Error codes (Starts with 3000)
    ''''''Loader --> 3003

    'Current Highest Custom Error codes (Starts with 4000)
    ''''''Custom --> 4011

    Public testingVersion As String = "1.0.0.8"
    Public betaVersion As String = "BETA " & testingVersion & "b"
    Public currentTestingIteration As Integer = 13

    'Public Objects
    Public xlApp As Object

    'SQL Connection Data
    Public dbDevelopment As String = My.Settings.devSql
    Public userDevelopment As String = My.Settings.ntokendev
    Public userPwDevelopment As String = My.Settings.wtokendev
    Public httpBaseLinkUAT As String = My.Settings.ccisiteshttpurl.Replace("[ENVIRONMENT]", ".uat")

    Public dbProduction As String = My.Settings.prodsql
    Public userProduction As String = My.Settings.ntoken
    Public userPwProduction As String = My.Settings.wtoken
    Public httpBaseLinkProduction As String = My.Settings.ccisiteshttpurl.Replace("[ENVIRONMENT]", "")

    Public dbActive As String
    Public userActive As String
    Public userPwActive As String
    Public fileLocActive As String
    Public httpBaseLink As String
    Public httpWOapi As String = "/eng/rest/ai/wo/v1"

    Public sqlCon As New SqlConnection
    Public ds As New DataSet
    Public da As SqlDataAdapter
    Public dt As New DataTable
    Public sql As String

    'SQL Select All Strings
    Public userSQLstr As String = "SELECT * FROM UserList"

    'SQL List Names
    Public userSQLsrc As String = "UList"

    'SQL Table Names
    Public userTable As String = "UserList"

    'Boolean Checks
    Public isOpening As Boolean = True
    Public isOpenSilent As Boolean = False
    Public isLoading As Boolean = False

    'User vars
    Public userID As String
    Public userFullName As String
    Public userShortName As String
    Public userEmail As String
    Public userSuper As String
    Public userType As String
    Public userDept As String
    Public userJobTitle As String
    Public userDeptPerm As String
    Public userPerm As String
    Public userPermission As Integer = 0
    Public userVer As String

    'EDS Specific connection information
    Public EDSdbDevelopment As String = "Server=DEVCCICSQL3.US.CROWNCASTLE.COM,58061;Database=EngEDSDev;Integrated Security=SSPI"
    Public EDSuserDevelopment As String = "366:204:303:354:207:330:309:207:204:249"
    Public EDSuserPwDevelopment As String = "210:264:258:99:297:303:213:258:246:318:354:111:345:168:300:318:261:219:303:267:246:300:108:165:144:192:324:153:246:300"

    Public EDSdbProduction As String = "Server=CCICSQLCLST2.US.CROWNCASTLE.COM,64540;Database=EDSProd;Integrated Security=SSPI"
    Public EDSuserProduction As String = "366:207:330:309:207:204:249"
    Public EDSuserPwProduction As String = "147:267:297:216:168:297:270:357:234:282:225:156:114:216:147:321:111:144:168:156:168:333:222:258:366:171:126:342:252:147"

    Public EDSdbUserAcceptance As String = "Server=DEVCCICSQL3.US.CROWNCASTLE.COM,58061;Database=EngEDSUat;Integrated Security=SSPI"

    Public EDSdbActive As String
    Public EDSuserActive As String
    Public EDSuserPwActive As String
    Public EDStokenHandle As New IntPtr(0)
    Public EDSimpersonatedUser As WindowsImpersonationContext
    Public EDSnewId As WindowsIdentity
End Module
