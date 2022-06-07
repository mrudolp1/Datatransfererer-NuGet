Imports System.Reflection

Module Object_Comparison
    'Source: https://www.codeproject.com/Articles/318877/Comparing-the-Properties-of-Two-Objects-via-Reflec

    Public Function AreObjectsEqual(ByVal objectA As Object, ByVal objectB As Object, ParamArray ignoreList As String()) As Boolean
        Dim result As Boolean

        If objectA IsNot Nothing AndAlso objectB IsNot Nothing Then
            Dim objectType As Type
            objectType = objectA.[GetType]()
            result = True

            For Each propertyInfo As PropertyInfo In objectType.GetProperties(BindingFlags.[Public] Or BindingFlags.Instance).Where(Function(p) p.CanRead AndAlso Not ignoreList.Contains(p.Name))
                Dim valueA As Object
                Dim valueB As Object
                valueA = propertyInfo.GetValue(objectA, Nothing)
                valueB = propertyInfo.GetValue(objectB, Nothing)

                If CanDirectlyCompare(propertyInfo.PropertyType) Then

                    If Not AreValuesEqual(valueA, valueB) Then
                        Console.WriteLine("Mismatch with property '{0}.{1}' found.", objectType.FullName, propertyInfo.Name)
                        result = False
                    End If
                ElseIf GetType(IEnumerable).IsAssignableFrom(propertyInfo.PropertyType) Then
                    Dim collectionItems1 As IEnumerable(Of Object)
                    Dim collectionItems2 As IEnumerable(Of Object)
                    Dim collectionItemsCount1 As Integer
                    Dim collectionItemsCount2 As Integer

                    If valueA Is Nothing AndAlso valueB IsNot Nothing OrElse valueA IsNot Nothing AndAlso valueB Is Nothing Then
                        Console.WriteLine("Mismatch with property '{0}.{1}' found.", objectType.FullName, propertyInfo.Name)
                        result = False
                    ElseIf valueA IsNot Nothing AndAlso valueB IsNot Nothing Then
                        collectionItems1 = (CType(valueA, IEnumerable)).Cast(Of Object)()
                        collectionItems2 = (CType(valueB, IEnumerable)).Cast(Of Object)()
                        collectionItemsCount1 = collectionItems1.Count()
                        collectionItemsCount2 = collectionItems2.Count()

                        If collectionItemsCount1 <> collectionItemsCount2 Then
                            Console.WriteLine("Collection counts for property '{0}.{1}' do not match.", objectType.FullName, propertyInfo.Name)
                            result = False
                        Else

                            For i As Integer = 0 To collectionItemsCount1 - 1
                                Dim collectionItem1 As Object
                                Dim collectionItem2 As Object
                                Dim collectionItemType As Type
                                collectionItem1 = collectionItems1.ElementAt(i)
                                collectionItem2 = collectionItems2.ElementAt(i)
                                collectionItemType = collectionItem1.[GetType]()

                                If CanDirectlyCompare(collectionItemType) Then

                                    If Not AreValuesEqual(collectionItem1, collectionItem2) Then
                                        Console.WriteLine("Item {0} in property collection '{1}.{2}' does not match.", i, objectType.FullName, propertyInfo.Name)
                                        result = False
                                    End If
                                ElseIf Not AreObjectsEqual(collectionItem1, collectionItem2, ignoreList) Then
                                    Console.WriteLine("Item {0} in property collection '{1}.{2}' does not match.", i, objectType.FullName, propertyInfo.Name)
                                    result = False
                                End If
                            Next
                        End If
                    End If
                ElseIf propertyInfo.PropertyType.IsClass Then

                    If Not AreObjectsEqual(propertyInfo.GetValue(objectA, Nothing), propertyInfo.GetValue(objectB, Nothing), ignoreList) Then
                        Console.WriteLine("Mismatch with property '{0}.{1}' found.", objectType.FullName, propertyInfo.Name)
                        result = False
                    End If
                Else
                    Console.WriteLine("Cannot compare property '{0}.{1}'.", objectType.FullName, propertyInfo.Name)
                    result = False
                End If
            Next
        Else
            result = Object.Equals(objectA, objectB)
        End If

        Return result
    End Function

    Private Function CanDirectlyCompare(ByVal type As Type) As Boolean
        Return GetType(IComparable).IsAssignableFrom(type) OrElse type.IsPrimitive OrElse type.IsValueType
    End Function

    Private Function AreValuesEqual(ByVal valueA As Object, ByVal valueB As Object) As Boolean
        Dim result As Boolean
        Dim selfValueComparer As IComparable
        selfValueComparer = TryCast(valueA, IComparable)

        If valueA Is Nothing AndAlso valueB IsNot Nothing OrElse valueA IsNot Nothing AndAlso valueB Is Nothing Then
            result = False
        ElseIf selfValueComparer IsNot Nothing AndAlso selfValueComparer.CompareTo(valueB) <> 0 Then
            result = False
        ElseIf Not Object.Equals(valueA, valueB) Then
            result = False
        Else
            result = True
        End If

        Return result
    End Function
End Module
