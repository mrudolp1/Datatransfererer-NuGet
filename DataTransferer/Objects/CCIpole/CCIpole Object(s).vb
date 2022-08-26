Option Strict On

Imports System.ComponentModel
Imports System.Data
Imports DevExpress.Spreadsheet
'Imports Microsoft.Office.Interop

Partial Public Class CCIpole
    Inherits EDSExcelObject

#Region "Inherited"
    Public Overrides ReadOnly Property templatePath As String = IO.Path.Combine(My.Application.Info.DirectoryPath, "Templates", "CCIpole.xlsm")

    Public Overrides ReadOnly Property excelDTParams As List(Of EXCELDTParameter)
        Get
            Return New List(Of EXCELDTParameter) From {New EXCELDTParameter("CCIpole General EXCEL", "A2:K3", "General (SAPI)"),
                                                        New EXCELDTParameter("CCIpole Pole Sections EXCEL", "A2:W20", "Unreinf Pole (SAPI)"),
                                                        New EXCELDTParameter("CCIpole Pole Reinf Sections EXCEL", "A2:W200", "Reinf Pole (SAPI)"),
                                                        New EXCELDTParameter("CCIpole Reinf Groups EXCEL", "A2:J50", "Reinf Groups (SAPI)"),
                                                        New EXCELDTParameter("CCIpole Reinf Details EXCEL", "A2:H200", "Reinf ID (SAPI)"),
                                                        New EXCELDTParameter("CCIpole Int Groups EXCEL", "A2:H50", "Interference Groups (SAPI)"),
                                                        New EXCELDTParameter("CCIpole Int Details EXCEL", "A2:H200", "Interference ID (SAPI)"),
                                                        New EXCELDTParameter("CCIpole Pole Reinf Results EXCEL", "A2:G1000", "Reinf Results (SAPI)"),
                                                        New EXCELDTParameter("CCIpole Reinf Property Details EXCEL", "A2:EA50", "Reinforcements (SAPI)"),
                                                        New EXCELDTParameter("CCIpole Bolt Property Details EXCEL", "A2:R20", "Bolts (SAPI)"),
                                                        New EXCELDTParameter("CCIpole Matl Property Details EXCEL", "A2:F20", "Materials (SAPI)")}
            '***Add additional table references here****
        End Get
    End Property

    Private _Insert As String
    Private _Update As String
    'Private _Delete As String

    Public Overrides ReadOnly Property EDSTableName As String
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Overrides ReadOnly Property EDSObjectName As String
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Overrides Sub workBookFiller(ByRef wb As Workbook)
        Throw New NotImplementedException()
    End Sub

    Public Overrides Function SQLInsertValues() As String
        Throw New NotImplementedException()
    End Function

    Public Overrides Function SQLInsertFields() As String
        Throw New NotImplementedException()
    End Function

    Public Overrides Function SQLUpdateFieldsandValues() As String
        Throw New NotImplementedException()
    End Function

    Public Overrides Function Equals(other As EDSObject, ByRef changes As List(Of AnalysisChange)) As Boolean
        Throw New NotImplementedException()
    End Function

#End Region

End Class
