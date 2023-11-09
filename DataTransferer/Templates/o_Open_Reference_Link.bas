Attribute VB_Name = "o_Open_Reference_Link"
Public Sub OpenQAPDF()
    OpenLinkInDefaultApp ("https://crowncastle.sharepoint.com/:b:/r/sites/TowerAssetsEngineering/Engineering%20Templates/Foundation%20Templates/CCI%20Foundation%20Tools%20Reference%20Manual/CCI%20Foundation%20Tools%20Q%26A.pdf?csf=1&web=1&e=rxzova")
End Sub


Public Sub OpenReferenceManPDF()
    OpenLinkInDefaultApp ("https://crowncastle.sharepoint.com/:b:/r/sites/TowerAssetsEngineering/Engineering%20Templates/Foundation%20Templates/CCI%20Foundation%20Tools%20Reference%20Manual/CCI%20Foundation%20Tools%20Reference%20Manual%20(5-27-2021).pdf?csf=1&web=1&e=BSIwg5")
End Sub


Public Sub OpenCriteriaLink()
    OpenLinkInDefaultApp ("https://crowncastle.sharepoint.com/:b:/r/sites/TowerAssetsEngineering/Engineering%20Templates/Foundation%20Templates/Foundation%20Criteria/CCI_Foundation_Criteria%20v1.2%20(5.5.16).pdf?csf=1&web=1&e=KO50Pp")
End Sub


Sub OpenLinkInDefaultApp(ByVal WebsiteURL As String)
    Dim ShellApp As Object

    ' Create a Shell Application object
    Set ShellApp = CreateObject("Shell.Application")

    ' Open the website URL in the default web browser
    ShellApp.ShellExecute WebsiteURL, "", "", "open", 1

    ' Clean up
    Set ShellApp = Nothing
End Sub

