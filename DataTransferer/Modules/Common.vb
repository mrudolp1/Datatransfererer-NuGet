Imports System.ComponentModel
Imports System.IO
Imports DevExpress.DataAccess.Excel

Module IDoDeclare
    Public ds As New DataSet
    Public queryPath As String = System.Windows.Forms.Application.StartupPath & "\Data Transferer Queries\"
    Public BUNumber As String = "3811932"
    Public STR_ID As String = "A"
    Public CurWO As String = "8794801"
    Public isModelNeeded As Boolean = False 'Update structure model & structure model xref
    Public isfndGroupNeeded As Boolean = False 'Update foundation details, foundation group & structure model
    Public isPileNeeded As Boolean = False 'Update pile details, pile location, pile soil layer & foundation details
    Public isPierAndPadNeeded As Boolean = False 'Update pier and pad details & foundation details

    'if changes were made, we need to ask the user if they want to set this as the ACTIVE model?
    Public overrideActiveModel As Boolean = True 'Structure model xref active (Potentially a boolean column or seperate table)

    Public currentFndGroup As Integer
    Public currentGuyConfig As Integer
    Public currentConnGroup As Integer
    'Lattice
    'Pole
End Module

Public Module Common

    Function GetExistingModelQuery() As String
        Return QueryBuilderFromFile(queryPath & "Existing Model (SELECT).sql").Replace("[BU NUMBER]", BUNumber).Replace("[STRUCTURE_ID]", STR_ID)
    End Function

    Function SaveExistingModelQuery() As String
        Return QueryBuilderFromFile(queryPath & "Existing Model (IN_UP).sql").Replace("[BU NUMBER]", BUNumber).Replace("[STRUCTURE_ID]", STR_ID)
    End Function

    Function SaveFoundationQuery(ByVal fndID As String, ByVal fndType As String) As String
        Return QueryBuilderFromFile(queryPath & "Foundations (IN_UP).sql").Replace("'[Foundation ID]'", IIf(fndID = "", "NULL", "'" & fndID & "'")).Replace("[FOUNDATION TYPE]", fndType)
    End Function

    Function QueryBuilderFromFile(ByVal filetoread As String) As String
        Dim sqlReader As New StreamReader(filetoread)
        Dim temp As String = ""

        Do While sqlReader.Peek <> -1
            temp += sqlReader.ReadLine.Trim + vbNewLine
        Loop

        sqlReader.Dispose()
        sqlReader.Close()

        Return temp
    End Function

    Public Function GetOneExcelRange(ByVal path As String, ByVal rng As String, Optional ByVal ws As String = "") As Object
        Dim exDS As New ExcelDataSource()
        Dim options As New ExcelSourceOptions()
        Dim val As DataTable
        Dim importSettings As Object

        If ws = "" Then
            importSettings = New ExcelDefinedNameSettings()
            importSettings.DefinedName = rng
        Else
            importSettings = New ExcelWorksheetSettings()
            importSettings.WorksheetName = ws
            importSettings.CellRange = rng
        End If

        options.ImportSettings = importSettings
        options.SkipHiddenColumns = False
        options.SkipHiddenRows = False
        options.UseFirstRowAsHeader = False

        exDS.FileName = path
        exDS.SourceOptions = options
        exDS.Fill()

        val = ExcelDatasourceToDataTable(exDS, ws & "|" & rng)

        Return val.Rows(0).ItemArray(0)
        'Return val.Columns(0).ColumnName
    End Function

    Public Function GetExcelDataSource(ByVal path As String, ByVal ws As String, ByVal rng As String) As ExcelDataSource
        'DevExpress specific process to fill an excel data source with information from a range in excel
        Dim exDS As New ExcelDataSource()
        Dim options As New ExcelSourceOptions()
        Dim importSettings = New ExcelWorksheetSettings()
        importSettings.WorksheetName = ws
        importSettings.CellRange = rng
        options.ImportSettings = importSettings
        With exDS
            .FileName = path
            .SourceOptions = options
            .Fill()
        End With

        Return exDS
    End Function

    Public Function ExcelDatasourceToDataTable(ByVal excelDataSource As ExcelDataSource, ByVal datasourcename As String) As DataTable
        'Convert a DevExpress Excel Data Source to a data table 
        Dim list As IList = (CType(excelDataSource, IListSource)).GetList()
        Dim dataView As DevExpress.DataAccess.Native.Excel.DataView = CType(list, DevExpress.DataAccess.Native.Excel.DataView)
        Dim props As List(Of DevExpress.DataAccess.Native.Excel.ViewColumn) = dataView.Columns
        Dim values(props.Count - 1) As Object
        Dim table As DataTable = New DataTable()

        'DevExpress automatically recognizes the top row as the column headers
        'Loop through the properties to add columns to the datatable
        For i As Integer = 0 To props.Count - 1
            Dim prop As PropertyDescriptor = props(i)
            table.Columns.Add(prop.Name, GetType(String)) 'prop.PropertyType)
        Next i

        'Loop through all other items and create new datarows
        'DevExpress automatically reconizes empty rows
        For Each item As DevExpress.DataAccess.Native.Excel.ViewRow In list
            For i As Integer = 0 To values.Length - 1
                values(i) = props(i).GetValue(item)
            Next i
            table.Rows.Add(values)
        Next item

        'Name the datatable to be referenced in the dataset
        table.TableName = datasourcename
        Return table
    End Function

#Region "Get File Properties"
    'Public Function GetFileTitle(ByVal xtdProp As List(Of ShellInfo)) As String
    '    For Each s As ShellInfo In xtdProp
    '        If s.Name = "Title" Then
    '            Return s.Value
    '        End If
    '    Next

    '    Return ""
    'End Function

    'Public Function GetXtdShellInfo(ByVal filepath As String) As List(Of ShellInfo)
    '    ' ToDo: add error checking, maybe Try/Catch and 
    '    ' surely check if the file exists before trying
    '    Dim xtd As New List(Of ShellInfo)

    '    Dim shell As New Shell32.Shell
    '    Dim shFolder As Shell32.Folder
    '    shFolder = shell.NameSpace(Path.GetDirectoryName(filepath))

    '    ' its com so iterate to find what we want -
    '    ' or modify to return a dictionary of lists for all the items
    '    Dim key As String

    '    For Each s In shFolder.Items
    '        ' look for the one we are after
    '        If shFolder.GetDetailsOf(s, 0).ToLowerInvariant = Path.GetFileName(filepath).ToLowerInvariant Then

    '            Dim ndx As Int32 = 0
    '            key = shFolder.GetDetailsOf(shFolder.Items, ndx)

    '            ' there are a varying number of entries depending on the OS
    '            ' 34 min, W7=290, W8=309 with some blanks

    '            ' this should get up to 310 non blank elements

    '            Do Until String.IsNullOrEmpty(key) AndAlso ndx > 310
    '                If String.IsNullOrEmpty(key) = False Then
    '                    xtd.Add(New ShellInfo(key,
    '                                      shFolder.GetDetailsOf(s, ndx)))
    '                End If
    '                ndx += 1
    '                key = shFolder.GetDetailsOf(shFolder.Items, ndx)
    '            Loop

    '            ' we got what we came for
    '            Exit For
    '        End If
    '    Next

    '    Return xtd
    'End Function
#End Region

#Region "Alternate Get One Excel Range"
    'Public Sub GetExcelRanges(ByVal filepath As String, ByRef ExcelRngs As List(Of EXCELRngParameter))
    '    'Dim exDS As New ExcelDataSource()
    '    Dim options As New ExcelSourceOptions()
    '    Dim importSettings As Object
    '    Dim tempdt As DataTable
    '    Dim sourceName As String

    '    'Use new excel datasource
    '    Using exDS As New ExcelDataSource()
    '        'Set the file path of the datasource, passed into method
    '        exDS.FileName = filepath

    '        'Set the variable value for each iteam in the excel range
    '        For Each item As EXCELRngParameter In ExcelRngs
    '            'If a worksheet was not set, it was assume the value is just a named range
    '            'Otherwise you will need a worksheet and a range
    '            If item.xlsSheet = "" Or IsNothing(item.xlsSheet) Then
    '                importSettings = New ExcelDefinedNameSettings()
    '                importSettings.DefinedName = item.rangeName
    '                sourceName = item.rangeName
    '            Else
    '                importSettings = New ExcelWorksheetSettings()
    '                importSettings.WorksheetName = item.xlsSheet
    '                importSettings.CellRange = item.xlsRange
    '                sourceName = item.xlsSheet & "|" & item.rangeName
    '            End If

    '            'Set the import settings and options based on the above inputs
    '            options.ImportSettings = importSettings
    '            exDS.SourceOptions = options
    '            exDS.Fill()

    '            'Convert the range to a datatable
    '            'When looking for 1 value it will set that value as the column name
    '            tempdt = ExcelDatasourceToDataTable(exDS, sourceName)

    '            'Set the variable value to the item 
    '            item.rangeValue = tempdt.Columns(0).ColumnName

    '            'Remove the temporary databale from the dataset
    '            dtClearer(sourceName)
    '        Next
    '    End Using
    'End Sub
#End Region

End Module

Public Class SQLParameter
    Public Property sqlDatatable As String
    Public Property sqlQuery As String

    Sub New()
        'Leave method empty
    End Sub

    Sub New(ByVal DataTableName As String, ByVal QueryFileName As String)
        sqlDatatable = DataTableName
        sqlQuery = QueryFileName
    End Sub
End Class

Public Class EXCELDTParameter
    Public Property xlsRange As String
    Public Property xlsSheet As String
    Public Property xlsDatatable As String = ""
    Sub New()
        'Leave method empty
    End Sub

    Sub New(ByVal ExcelTableName As String, ByVal ExcelRange As String, ByVal ExcelSheet As String)
        xlsDatatable = ExcelTableName
        xlsRange = ExcelRange
        xlsSheet = ExcelSheet
    End Sub
End Class

Public Class EXCELRngParameter
    Public Property xlsRange As String
    Public Property xlsSheet As String
    Public Property rangeName As String
    Public Property rangeValue As String
    Public Property variableName As String

    Sub New()
        'Leave method empty
    End Sub

    Sub New(ByVal ExcelRange As String, ByVal ExcelSheet As String, ByVal EDSName As String)
        xlsRange = ExcelRange
        xlsSheet = ExcelSheet
        variableName = EDSName
    End Sub

    Sub New(ByVal xlNamedRange As String, ByVal EDSName As String)
        rangeName = xlNamedRange
        variableName = EDSName
    End Sub
End Class
