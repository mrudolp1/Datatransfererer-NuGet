Attribute VB_Name = "Z_UserData"
' ----------------------------------------------------------------
' Procedure Name: User_Data
' Purpose: Get name of the current user. Not limited to just the name. You can retrieve many user properties.
' Procedure Kind: Function
' Procedure Access: Public
' Return Type: String
' Author: IMiller
' Date: 4/17/2019
' ----------------------------------------------------------------

Public Function UserData(ByVal data As String) ' As String
1     On Error GoTo UserData_Error

2     Dim objAd         As Object: Set objAd = CreateObject("ADSystemInfo")
3     Dim objuser       As Object: Set objuser = GetObject("LDAP://" & objAd.UserName)

4         Select Case data
              Case "FirstName":                       UserData = objuser.FirstName
5             Case "LastName":                        UserData = objuser.LastName
6             Case "FullName":                        UserData = objuser.FullName
7             Case "Description":                     UserData = objuser.Description
8             Case "physicalDeliveryOfficeName":      UserData = objuser.physicalDeliveryOfficeName
9             Case "telephoneNumber":                 UserData = objuser.telephoneNumber
10            Case "EmailAddress":                    UserData = objuser.EmailAddress
11            Case "streetAddress":                   UserData = objuser.streetAddress
12            Case "city":                            UserData = objuser.L
13            Case "state":                           UserData = objuser.st
14            Case "zip":                             UserData = objuser.postalCode
15            Case "UserName":                        UserData = objuser.sAMAccountName
16            Case "Mobile":                          UserData = objuser.Mobile
17            Case "ipPhone":                         UserData = objuser.ipPhone
18            Case "Title":                           UserData = objuser.Title
19            Case "Department":                      UserData = objuser.department
20            Case "Company":                         UserData = objuser.company
21        End Select
          
22    On Error GoTo 0
23    Exit Function

UserData_Error:
          'Call globalerror("UserData", Err.Number, Err.Description, Err.Source, Erl)

24    On Error GoTo 0
End Function


