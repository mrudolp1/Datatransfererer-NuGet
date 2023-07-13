'JSON Serializer
Imports System.IO
Imports System.Runtime.Serialization.Json
Imports System.Text

Public Module JsonUtil
    Public Function FromJsonString(Of T)(ByVal jsonString As String) As Tuple(Of T, String)
        Using aMemoryStream As MemoryStream = New MemoryStream(Encoding.UTF8.GetBytes(jsonString))
            Dim ser As DataContractJsonSerializer = New DataContractJsonSerializer(GetType(T))
            Dim myObj As T
            Dim resultTxt As String = "Success"
            Try
                myObj = CType(ser.ReadObject(aMemoryStream), T)
            Catch ex As Exception
                resultTxt = "ERROR DESERIALIZING " & ex.Message
            End Try

            Return New Tuple(Of T, String)(myObj, resultTxt)
        End Using
    End Function

    Public Function FromJsonString(Of T)(ByVal jsonString As String, ByVal serializerInstance As DataContractJsonSerializer) As T
        Using aMemoryStream As MemoryStream = New MemoryStream(Encoding.UTF8.GetBytes(jsonString))
            Dim ser = New DataContractJsonSerializer(GetType(T))
            Return CType(ser.ReadObject(aMemoryStream), T)
        End Using
    End Function

    Public Function ToJsonString(ByVal valueObject As Object, ByVal serializerInstance As DataContractJsonSerializer) As String
        Using aMemoryStream As MemoryStream = New MemoryStream()
            serializerInstance.WriteObject(aMemoryStream, valueObject)
            Return Encoding.[Default].GetString(aMemoryStream.ToArray())
        End Using
    End Function

    Public Function ToJsonString(Of T)(ByVal valueObject As T) As String
        Using aMemoryStream As MemoryStream = New MemoryStream()
            Dim serializer As DataContractJsonSerializer = New DataContractJsonSerializer(GetType(T))
            Try
                serializer.WriteObject(aMemoryStream, valueObject)
            Catch ex As Exception
                Return "ERROR SERIALIZING: " & ex.Message
            End Try

            Return Encoding.[Default].GetString(aMemoryStream.ToArray())
        End Using
    End Function
End Module
