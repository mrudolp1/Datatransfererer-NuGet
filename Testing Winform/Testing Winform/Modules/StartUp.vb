'------------------------------------------
'------------------------------------------
'Module Name: StartUp
'Purpose: Houses all procedures used during startup. Checks username against database and adds users automatically. Also tests DB connection.
'------------------------------------------
'------------------------------------------

Imports System.Security.Principal
Namespace UnitTesting


    Module StartUp
        Public deployVer As String

        Public Function StartEverything() As Boolean
            If My.Settings.serverActive = "dbDevelopment" Then
                If My.Settings.dbSelection = "UAT" Then
                    EDSdbActive = EDSdbUserAcceptance
                Else
                    EDSdbActive = EDSdbDevelopment
                End If
                EDSuserActive = EDSuserDevelopment
                EDSuserPwActive = EDSuserPwDevelopment
            Else
                EDSdbActive = EDSdbProduction
                EDSuserActive = EDSuserProduction
                EDSuserPwActive = EDSuserPwProduction
            End If
            'Dim u As String = token(userActive) '"zdevengEDS" '"zEngEDS"
            'Dim p As String = token(userPwActive) '"Ng@kT4oy9qleTZlgloY" '"GYui*AD59@^m$gR"

            AddHandler toaster.UpdateToastContent, AddressOf Toaster_UpdateToastContent

            'Logon the user to be used
            Dim logSuc As Boolean
            Try
                logSuc = LogonUser(token(EDSuserActive), "CCIC", token(EDSuserPwActive), 2, 0, EDStokenHandle) 'token(userActive), token(userPwActive)
                EDSnewId = New WindowsIdentity(EDStokenHandle)
            Catch ex As Exception
                sendToast("Error logging into network.", "Error 4001")
            End Try


            If dbConnectionTest() Then
                If Not CollectUserInfo() Then
                    sendToast("Error collecting user information", "Error 4000")
                    Return False
                End If
            Else
                Return False
            End If

            SetVersion()

            Return True
        End Function

        Public Function CollectUserInfo() As Boolean
            Dim retries As Integer = 0

            Try
                deployVer = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
            Catch ex As Exception
                deployVer = betaVersion
            End Try

            If TableContainsRecords(GetUserInfo) Then
                SetUserInfo()
                If Not SQLMatchesCurrentUserData() Then
                    SetInitialUserInfo()
                    sqlSender("UPDATE gen.UserList SET version= '" & userVer & "', fullName='" & userFullName & "', userName='" & userShortName & "', emailAddress='" & userEmail & "', jobTitle='" & userJobTitle & "', department='" & userDept & "' WHERE ID=" & userID & ";", 2000)
                    GetUserInfo()
                    SetUserInfo()
                End If
            Else
                SetInitialUserInfo()
                sqlSender("INSERT INTO gen.UserList (fullName, userName, emailAddress, jobTitle, department, version) VALUES ('" & userFullName & "', '" & userShortName & "', '" & userEmail & "', '" & userJobTitle & "', '" & userDept & "','" & userVer & "');", 2001)
                If TableContainsRecords(GetUserInfo) Then
                    SetUserInfo()
                Else
                    Return False
                End If
            End If

            Return True
        End Function

        Function GetUserInfo() As DataTable
            sqlLoader("SELECT * FROM gen.UserList WHERE userName='" & Environment.UserName & "'", "UList", 1000)
            Return ds.Tables("UList")
        End Function

        Sub SetUserInfo()
            userID = CType(ds.Tables("UList").Rows(0).Item("ID"), Integer)
            userFullName = ds.Tables("UList").Rows(0).Item("fullName").ToString
            userShortName = ds.Tables("UList").Rows(0).Item("userName").ToString
            userEmail = ds.Tables("UList").Rows(0).Item("emailAddress").ToString
            userJobTitle = ds.Tables("UList").Rows(0).Item("jobTitle").ToString
            userDept = ds.Tables("UList").Rows(0).Item("department").ToString
            userVer = ds.Tables("UList").Rows(0).Item("version").ToString
        End Sub

        Sub SetInitialUserInfo()
            userFullName = FetchUserData("FullName")
            userShortName = FetchUserData("UserName")
            userJobTitle = FetchUserData("Title")
            userDept = FetchUserData("Department")
            userEmail = FetchUserData("EmailAddress")
            userVer = deployVer
        End Sub

        Function SQLMatchesCurrentUserData() As Boolean
            Dim doesit As Boolean = True

            If userFullName <> FetchUserData("FullName") Then doesit = False
            If userShortName <> FetchUserData("UserName") Then doesit = False
            If userJobTitle <> FetchUserData("Title") Then doesit = False
            If userDept <> FetchUserData("Department") Then doesit = False
            If userEmail <> FetchUserData("EmailAddress") Then doesit = False
            If userVer <> deployVer Then doesit = False

            Return doesit
        End Function

        Public Sub SetVersion()
            Try
                UnitTesting.frmMain.Text = My.Application.Info.AssemblyName & " - " & FetchUserData("FullName") & " - " & userID
                UnitTesting.frmMain.verLabel.Text = System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
            Catch ex As Exception
                UnitTesting.frmMain.verLabel.Text = betaVersion
            End Try
        End Sub

        Public Function dbConnectionTest() As Boolean
            If sqlLoader("SELECT TOP(10) * FROM gen.userlist", "UList", 1001) Then
                Try
                    ds.Tables("UList").Clear()
                Catch
                End Try
                Return True
            Else
                Return False
            End If
        End Function

    End Module

End Namespace