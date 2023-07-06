Imports System.Data.SqlClient
Imports Krypton.Navigator

'------------------------------------------
'------------------------------------------
'Module Name: IDoDeclare
'Purpose: Designed to declare all public variables used throughout the program
'------------------------------------------
'------------------------------------------

Module IDoDeclare

    'Current Highest SQL Error codes (Starts with 2000 & 1000)
    ''''''Loader --> 1001
    ''''''Sender --> 2002

    'Current Highest ORACLE Error codes (Starts with 3000)
    ''''''Loader --> 3003

    'Current Highest Custom Error codes (Starts with 4000)
    ''''''Custom --> 4011

    Public betaVersion As String = "BETA 1.0.0.1"

    'Public Objects
    Public xlApp As Object

    'SQL Connection Data
    Public dbDevelopment As String = My.Settings.devSql
    Public userDevelopment As String = My.Settings.ntokenDev
    Public userPwDevelopment As String = My.Settings.wTokenDev
    Public httpBaseLinkUAT As String = My.Settings.ccisitesHttpUrl.Replace("[ENVIRONMENT]", ".uat")

    Public dbProduction As String = My.Settings.prodSql
    Public userProduction As String = My.Settings.nToken
    Public userPwProduction As String = My.Settings.wToken
    Public httpBaseLinkProduction As String = My.Settings.ccisitesHttpUrl.Replace("[ENVIRONMENT]", "")

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



End Module
