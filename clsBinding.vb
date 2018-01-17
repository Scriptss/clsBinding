Imports System.ComponentModel
Imports System.Collections.Generic
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json
Imports System.Xml.Serialization

Public Class clsBinding
    Inherits System.Collections.CollectionBase
    Implements System.ComponentModel.IBindingList

#Region " IBindingList "
    Private _AllowSort As Boolean = True
    Private _IsSorted As Boolean
    Private _SortProperty As PropertyDescriptor
    Private _ListSortDirection As ListSortDirection = ListSortDirection.Ascending
    ''' <summary>
    ''' Indicates that one or more items in the collection have changed.
    ''' </summary>

    Friend Event SortEvent(ByVal PropertyToSort As String, ByVal SortOrder As ListSortDirection)
    Public Event ListChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ListChangedEventArgs) Implements System.ComponentModel.IBindingList.ListChanged
    Public Event Sorted(ByVal sender As Object, ByVal e As SortedEventArgs)


    ''' Returns True to indicate that we support change notification by
    ''' raising the
    Public ReadOnly Property SupportsChangeNotification() As Boolean _
        Implements System.ComponentModel.IBindingList.SupportsChangeNotification
        Get
            Return True
        End Get
    End Property

    ''' Returns True to indicate that we support automatic addition
    ''' of new line item objects to the collection.
    Public ReadOnly Property AllowNew() As Boolean Implements System.ComponentModel.IBindingList.AllowNew
        Get
            Return False
        End Get
    End Property

    ''' Returns True to indicate that we allow in-place editing of
    ''' line item objects.
    Public ReadOnly Property AllowEdit() As Boolean Implements System.ComponentModel.IBindingList.AllowEdit
        Get
            Return True
        End Get
    End Property

    ''' Returns True to indicate that we support automatic removal
    ''' of line item objects from the collection.
    Public ReadOnly Property AllowRemove() As Boolean Implements System.ComponentModel.IBindingList.AllowRemove
        Get
            Return True
        End Get
    End Property

    ''' Adds a new item to the collection on request by the UI control.
    Public Function AddNew() As Object Implements System.ComponentModel.IBindingList.AddNew


    End Function

    Protected Overrides Sub OnInsertComplete(ByVal Index As Integer, ByVal Value As Object)

        RaiseEvent ListChanged(Me, New ListChangedEventArgs(ListChangedType.ItemAdded, Index))

    End Sub

    Protected Overrides Sub OnSetComplete(ByVal Index As Integer, ByVal OldValue As Object, ByVal NewValue As Object)

        RaiseEvent ListChanged(Me, New ListChangedEventArgs(ListChangedType.ItemChanged, Index))

    End Sub

    Protected Overrides Sub OnRemoveComplete(ByVal Index As Integer, ByVal Value As Object)

        RaiseEvent ListChanged(Me, New ListChangedEventArgs(ListChangedType.ItemDeleted, Index))

    End Sub

    Protected Overrides Sub OnClearComplete()
        RaiseEvent ListChanged(Me, New ListChangedEventArgs(ListChangedType.Reset, -1))

    End Sub
    ' the remainder of the IBindingList methods are not implemented
    ' we do not support searching or sorting

    Public ReadOnly Property SupportsSearching() As Boolean Implements System.ComponentModel.IBindingList.SupportsSearching
        Get
            Return False
        End Get
    End Property

    Public ReadOnly Property SupportsSorting() As Boolean Implements System.ComponentModel.IBindingList.SupportsSorting
        Get
            Return _AllowSort
        End Get
    End Property

    Public Property AllowSort() As Boolean
        Get
            Return _AllowSort
        End Get
        Set(ByVal Value As Boolean)
            _AllowSort = Value
        End Set
    End Property

    Public ReadOnly Property IsSorted() As Boolean Implements System.ComponentModel.IBindingList.IsSorted
        Get
            Return _IsSorted
        End Get
    End Property

    Public Sub AddIndex(ByVal [property] As System.ComponentModel.PropertyDescriptor) Implements System.ComponentModel.IBindingList.AddIndex
        _IsSorted = True
        _SortProperty = [property]
    End Sub

    Public Sub ApplySort(ByVal [property] As System.ComponentModel.PropertyDescriptor, ByVal direction As System.ComponentModel.ListSortDirection) Implements System.ComponentModel.IBindingList.ApplySort

        _IsSorted = True
        _SortProperty = [property]
        _ListSortDirection = direction

        RaiseEvent SortEvent([property].Name, direction)
        'Sort([property].Name, direction)
        RaiseEvent Sorted(Me, New SortedEventArgs(_ListSortDirection))

    End Sub

    Public Function Find(ByVal [property] As System.ComponentModel.PropertyDescriptor, ByVal key As Object) As Integer Implements System.ComponentModel.IBindingList.Find

    End Function

    Public Sub RemoveIndex(ByVal [property] As System.ComponentModel.PropertyDescriptor) Implements System.ComponentModel.IBindingList.RemoveIndex
        _SortProperty = Nothing
    End Sub

    Public Sub RemoveSort() Implements System.ComponentModel.IBindingList.RemoveSort
        _SortProperty = Nothing
        _IsSorted = False
    End Sub

    Public ReadOnly Property SortDirection() As System.ComponentModel.ListSortDirection _
        Implements System.ComponentModel.IBindingList.SortDirection
        Get

            Return _ListSortDirection

        End Get
    End Property

    Public ReadOnly Property SortProperty() As System.ComponentModel.PropertyDescriptor Implements System.ComponentModel.IBindingList.SortProperty
        Get
            Return _SortProperty
        End Get
    End Property


#End Region


End Class


#Region "Binding"
Public Class clsBindingListBase(Of T)
    Inherits BindingList(Of T)

    Implements IDisposable
    Implements System.ComponentModel.ITypedList
    Implements IRaiseItemChangedEvents
    Implements ICancelAddNew

    'properties and methods of your business object list class 

    'these are used to implement the sorting for the BindingList 
    Private _mSorted As Boolean = False
    Private _mSortDirection As ListSortDirection = ListSortDirection.Ascending
    Private _mSortProperty As PropertyDescriptor = Nothing
    Private failedToAdd As Boolean = False
    Private lastNewIndex As Integer = -1


    '-- This one is used to store Keys
    Private mvarkeys As Collections.ArrayList

    Private Const BaseError As Integer = 513

    '-- These error values arbitrarily start at 1000
    Private Const ERR_ITEM_SET As Integer = BaseError + 1000
    Private Const ERR_ITEM_GET As Integer = BaseError + 1001
    Private Const ERR_ADD As Integer = BaseError + 1002
    Private Const ERR_REMOVE As Integer = BaseError + 1004
    Private Const ERR_ADDRANGE As Integer = BaseError + 1005
    Private Const ERR_SORT As Integer = BaseError + 1006

    Private _mSortDescriptions As ListSortDescriptionCollection

    Private _mSorts As PropertyComparerCollection(Of T)

    Private comparers As List(Of PropertyComparer2(Of T))

    Public Event Sorted(ByVal sender As Object, ByVal e As SortedEventArgs)


    Public Sub New()
        mvarkeys = New Collections.ArrayList
    End Sub

    Public Sub New(list As System.Collections.Generic.IList(Of T))
        MyBase.New(list)
        mvarkeys = New Collections.ArrayList
    End Sub

    Protected Overridable Overloads Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            Me.Clear()
        End If
    End Sub

#Region " IDisposable Support "
    ' Do not change or add Overridable to these methods. 
    ' Put cleanup code in Dispose(ByVal disposing As Boolean). 
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
    Protected Overrides Sub Finalize()
        Dispose(False)
        MyBase.Finalize()
    End Sub
#End Region


    Public Property Keys() As ArrayList
        Get
            Return mvarkeys
        End Get
        Set(ByVal value As ArrayList)
            mvarkeys = value
        End Set
    End Property

    Public Function GetItemProperties(ByVal listAccessors() As PropertyDescriptor) As PropertyDescriptorCollection Implements ITypedList.GetItemProperties
        Return TypeDescriptor.GetProperties(GetType(T), New Attribute() {New BrowsableAttribute(True)})
    End Function 'GetItemProperties

    Public Function GetListName(ByVal listAccessors() As PropertyDescriptor) As String Implements ITypedList.GetListName

        Return GetType(T).Name

    End Function 'GetItemProperties


    Protected Overloads Overrides Sub OnListChanged(ByVal e As System.ComponentModel.ListChangedEventArgs)
        If Not failedToAdd Then
            MyBase.OnListChanged(e)
        End If
        failedToAdd = False
    End Sub

    Protected Overloads Overrides Sub OnAddingNew(ByVal e As System.ComponentModel.AddingNewEventArgs)
        failedToAdd = e.NewObject Is Nothing
        MyBase.OnAddingNew(e)
    End Sub

    Public Overrides Sub CancelNew(ByVal itemIndex As Integer)
        MyBase.CancelNew(itemIndex)

    End Sub

    Public Overrides Sub EndNew(ByVal itemIndex As Integer)
        MyBase.EndNew(itemIndex)
    End Sub

    Public Overloads Function Add(ByVal value As T, Optional ByVal Key As String = "") As T
        Try
            MyBase.Add(value)
            mvarkeys.Add(LCase(Key)) '-- Add the key to mvarkeys - even if it's "". It's critical to keep the same number of elements in each array.
            '-- LCase is for case-insensitivity. 
            Return value
        Catch e As Exception
            Err.Raise(ERR_ADD, "col.Add", e.Message)
        End Try
    End Function

    Public Overloads Function AddNew() As T
        Dim obj
        obj = MyBase.AddNew()
        mvarkeys.Add("")
        MyBase.OnListChanged(New ListChangedEventArgs(ListChangedType.ItemAdded, Me.Count - 1))

        Return obj
    End Function



    Public Overridable ReadOnly Property Index(ByVal Key As String) As Integer
        '-- Returns the index given a key.

        Get
            Dim ThisIndex As Integer = mvarkeys.IndexOf(LCase(Key))

            '-- Return the index.
            Return ThisIndex
        End Get
    End Property

    Default Public Overloads Property Item(ByVal Index As Integer) As T
        Get
            Try
                Return CType(MyBase.Item(Index), T)
            Catch e As Exception
                Err.Raise(ERR_ITEM_GET, "col.Item", e.Message)
            End Try
        End Get

        Set(ByVal Value As T)
            Try
                MyBase.Item(Index) = Value
            Catch e As Exception
                Err.Raise(ERR_ITEM_GET, "col.Item", e.Message)
            End Try

        End Set

    End Property


    Default Public Overloads ReadOnly Property Item(ByVal myValue As T) As T
        Get
            Dim myIndex As Integer
            Try
                myIndex = MyBase.IndexOf(myValue)
                Return CType(MyBase.Item(myIndex), T)
            Catch e As Exception
                Err.Raise(ERR_ITEM_GET, "col.Item", e.Message)
            End Try
        End Get

    End Property

    Default Public Overloads Property Item(ByVal Key As String) As T
        Get
            Try
                Dim ThisIndex As Integer = mvarkeys.IndexOf(LCase(CStr(Key)))

                '-- If valid, return the item.
                If ThisIndex >= 0 Then
                    Return CType(MyBase.Item(ThisIndex), T)
                End If
            Catch e As Exception
                Err.Raise(ERR_ITEM_GET, "col.Item", e.Message)
            End Try
        End Get

        Set(ByVal Value As T)

            Try
                Dim ThisIndex As Integer = mvarkeys.IndexOf(LCase(CStr(Key)))

                MyBase.Item(ThisIndex) = Value

            Catch e As Exception
                Err.Raise(ERR_ITEM_GET, "colAddresses.Item", e.Message)
            End Try
        End Set
    End Property


    Public Overridable Property Key(ByVal Index As Integer) As String
        '-- Returns the Key given an index
        Get
            Dim ThisIndex As Integer = CInt(Index)

            '-- The Keys collection holds the keys

            If Index > -1 Then
                If mvarkeys.Count > Index Then
                    Return CStr(mvarkeys(ThisIndex))
                End If
            End If
        End Get

        Set(ByVal Value As String)
            Dim ThisIndex As Integer = CInt(Index)

            If ThisIndex < mvarkeys.Count AndAlso ThisIndex > -1 Then
                mvarkeys(ThisIndex) = LCase(Value)
            ElseIf ThisIndex = mvarkeys.Count Then
                mvarkeys.Add(LCase(Value))
            End If
        End Set
    End Property

    Public Overridable Overloads Sub Clear()
        '-- Clears the collection

        Try
            For Each item As IDisposable In Me
                If item IsNot Nothing Then
                    item.Dispose()
                End If
            Next

        Catch ex As Exception

        End Try

        MyBase.Clear()
        mvarkeys.Clear()

    End Sub

    Public Overloads Sub Remove(ByVal Value As T)
        Try
            Dim Index As Integer
            Index = MyBase.IndexOf(Value)
            'RaiseEvent ListChanged(Me, New ListChangedEventArgs(ListChangedType.ItemDeleted, Index))

            MyBase.Remove(Value)

            If Index > -1 And Index < mvarkeys.Count Then
                mvarkeys.RemoveAt(Index)
            End If
        Catch e As Exception
            Err.Raise(ERR_REMOVE, "col.Remove", e.Message)
        End Try
    End Sub

    Public Overloads Sub Remove(ByVal Index As Integer)

        Dim value As T
        Try
            value = CType(MyBase.Item(Index), T)
            If Not value Is Nothing Then
                MyBase.Remove(value)
                mvarkeys.RemoveAt(CInt(Index))
            End If
        Catch e As Exception
            Err.Raise(ERR_REMOVE, "col.Remove", e.Message)
        End Try
    End Sub

    Public Overloads Sub Remove(ByVal Key As String)

        Try
            Dim ThisIndex As Integer = mvarkeys.IndexOf(LCase(CStr(Key)))

            If ThisIndex >= 0 Then  '-- Did we find it?
                '-- Yes! Remove the item and its key
                MyBase.RemoveAt(ThisIndex)
                mvarkeys.RemoveAt(ThisIndex)
            End If

        Catch e As Exception
            Err.Raise(ERR_REMOVE, "col.Remove", e.Message)
        End Try
    End Sub

    Public Overloads Sub RemoveAt(ByVal Index As Integer)

        Dim value As T
        Try
            value = CType(MyBase.Item(Index), T)
            If Not value Is Nothing Then
                MyBase.RemoveAt(Index)
                mvarkeys.RemoveAt(CInt(Index))
            End If
        Catch e As Exception
            Err.Raise(ERR_REMOVE, "col.Remove", e.Message)
        End Try
    End Sub

    Public Shadows ReadOnly Property IndexOf(ByVal Item As T) As Integer
        '-- Returns the index of a given Item
        Get
            Return MyBase.IndexOf(Item)
        End Get
    End Property

#Region "sorting"
    ''' Returns True to indicate that we support change notification by
    ''' raising the
    Public Overridable Sub Sort(ByVal PropertyToSort As String)

        Sort(PropertyToSort, ListSortDirection.Ascending)

    End Sub

    Public Overridable Sub Sort(ByVal PropertyToSort As String, ByVal SortOrder As ListSortDirection)

        Dim myProperty As PropertyDescriptor = TypeDescriptor.GetProperties(GetType(T)).Item(PropertyToSort)

        Dim pi() As System.Reflection.PropertyInfo = GetType(T).GetProperties()
        Dim p As System.Reflection.PropertyInfo


        ApplySortCore(myProperty, SortOrder)


    End Sub

    Protected Overloads Overrides ReadOnly Property SupportsSortingCore() As Boolean
        Get
            Return True
        End Get
    End Property

    ''' <summary> 
    ''' Return the value retained locally. Tells if the list is sorted or not. 
    ''' </summary> 
    Protected Overloads Overrides ReadOnly Property IsSortedCore() As Boolean
        Get
            Return _mSorted
        End Get
    End Property

    ''' <summary> 
    ''' Return the value retained locally. Tells which direction the list is sorted. 
    ''' </summary> 
    Protected Overloads Overrides ReadOnly Property SortDirectionCore() As ListSortDirection
        Get
            Return _mSortDirection
        End Get
    End Property

    ''' <summary> 
    ''' Return the value retained locally. Tells which property the list is sorted on. 
    ''' </summary> 
    Protected Overloads Overrides ReadOnly Property SortPropertyCore() As PropertyDescriptor
        Get
            Return _mSortProperty
        End Get
    End Property


    Protected Overloads Overrides Sub ApplySortCore(ByVal prop As PropertyDescriptor, ByVal direction As ListSortDirection)
        _mSortDirection = direction
        _mSortProperty = prop
        Dim comparer As New PropertyComparer(Of T)(prop, direction)
        ApplySortInternal(comparer)
    End Sub

    'Protected Overloads Overrides Sub ApplySortCore(ByVal prop As PropertyDescriptor, ByVal direction As ListSortDirection)
    '    Dim arr As ListSortDescription() = {New ListSortDescription(prop, direction)}
    '    ApplySort(New ListSortDescriptionCollection(arr))
    'End Sub

    Private Sub ApplySortInternal(ByVal comparer As PropertyComparer(Of T))
        'this causes the items in the collection maintained by the base class to be sorted 
        ' according to the criteria provided to the BOSortComparer class. 
        Dim listRef As List(Of T) = TryCast(Me.Items, List(Of T))
        If listRef Is Nothing Then
            Exit Sub
        End If

        'let List<T> do the actual sorting based on your comparer 
        listRef.Sort(comparer)
        _mSorted = True

        'fire an event through a call to the base class OnListChanged method indicating 
        ' that the list has been changed. 
        'Use 'reset' because it's likely that most members have been moved around. 
        OnListChanged(New ListChangedEventArgs(ListChangedType.Reset, -1))


    End Sub


    Public Sub ApplySort(ByVal sortCollection As ListSortDescriptionCollection)
        Dim oldRaise As Boolean = RaiseListChangedEvents
        RaiseListChangedEvents = False
        Try
            Dim tmp As New PropertyComparerCollection(Of T)(sortCollection)
            Dim items As New List(Of T)(Me)
            items.Sort(tmp)
            Dim index As Integer = 0
            For Each item As T In items
                'SetItem(System.Math.Max(System.Threading.Interlocked.Increment(index), index - 1), item)
                SetItem(index, item)
                index += 1
            Next
            _mSorts = tmp
        Finally
            RaiseListChangedEvents = oldRaise
            ResetBindings()
        End Try
    End Sub


#End Region

#Region "Filter"
    Private _mUnfilteredListValue As New List(Of T)()
    Private _mFilterValue As String = Nothing
    Private _mCriteria As String

    Public ReadOnly Property UnfilteredList() As List(Of T)
        Get
            Return _mUnfilteredListValue
        End Get
    End Property
    Public Property Filter() As String
        Get
            Return _mFilterValue
        End Get
        Set(ByVal value As String)
            If _mFilterValue <> value Then


                RaiseListChangedEvents = False

                ' If filter value is null, reset list. 

                If value Is Nothing Then


                    Me.ClearItems()

                    For Each t As T In _mUnfilteredListValue

                        Me.Items.Add(t)
                    Next

                    _mFilterValue = value
                    _mUnfilteredListValue.Clear()




                    ' If the value is empty string, do nothing. 

                    ' This behavior is compatible with DataGridView 

                    ' AutoFilter code. 

                ElseIf value = "" Then

                    'If the value is not null or string, than process normal. 


                ElseIf Regex.Matches(value, "[?[\w ]+]? ?[=<>] ?'?[\w|/: ]+'?", RegexOptions.Singleline).Count = 1 Then

                    ' If the filter is not set. 
                    If _mUnfilteredListValue.Count = 0 Then

                        _mUnfilteredListValue.AddRange(Me.Items)
                        Me.ClearItems()
                    End If

                    _mFilterValue = value

                    GetFilterParts()

                    ApplyFilter()




                ElseIf Regex.Matches(value, "[?[\w ]+]? ?[<]?[>] ?'?[\w|/: ]+'?", RegexOptions.Singleline).Count = 1 Then

                    If _mUnfilteredListValue.Count = 0 Then

                        _mUnfilteredListValue.AddRange(Me.Items)
                        Me.ClearItems()
                    End If

                    _mFilterValue = value

                    GetFilterParts()

                    ApplyFilter()
                ElseIf Regex.Matches(value, "[?[\w ]+]? ?[=>] ?'?[\w|/: ]+'?", RegexOptions.Singleline).Count > 1 Then


                    Throw New ArgumentException("Multi-column" & "filtering is not implemented.")
                Else




                    Throw New ArgumentException("Filter is not" & "in the format: propName = 'value'.")
                End If



                RaiseListChangedEvents = True

                OnListChanged(New ListChangedEventArgs(ListChangedType.Reset, -1))

            End If


        End Set
    End Property

    Private filterPropertyNameValue As String

    Private filterCompareValue As Object

    Public ReadOnly Property FilterPropertyName() As String


        Get
            Return filterPropertyNameValue
        End Get
    End Property


    Public ReadOnly Property FilterCompare() As Object


        Get
            Return filterCompareValue
        End Get
    End Property

    Public Sub GetFilterParts()



        Dim filterParts As String()

        If Filter.Contains("=") Then
            filterParts = Filter.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)
            _mCriteria = "="
        ElseIf Filter.Contains("<>") Then
            Dim strFilter As String = Filter.Replace("<>", "=")
            filterParts = strFilter.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)
            _mCriteria = "<>"
        ElseIf Filter.Contains(">") Then
            filterParts = Filter.Split(New Char() {">"c}, StringSplitOptions.RemoveEmptyEntries)
            _mCriteria = ">"
        ElseIf Filter.Contains("<") Then
            filterParts = Filter.Split(New Char() {"<"c}, StringSplitOptions.RemoveEmptyEntries)
            _mCriteria = "<"

        End If


        filterPropertyNameValue = filterParts(0).Replace("[", "").Replace("]", "").Trim()


        Dim propDesc As PropertyDescriptor = TypeDescriptor.GetProperties(GetType(T))(filterPropertyNameValue.ToString())

        If propDesc IsNot Nothing Then


            Try


                Dim converter As TypeConverter = TypeDescriptor.GetConverter(propDesc.PropertyType)

                filterCompareValue = converter.ConvertFromString(filterParts(1).Replace("'", "").Trim())

            Catch generatedExceptionName As NotSupportedException





                Throw New ArgumentException("Specified filter value " & FilterCompare & " can not be converted from string." & "..Implement a type converter for " & propDesc.PropertyType.ToString())

            End Try

        Else



            Throw New ArgumentException("Specified property '" & FilterPropertyName & "' is not found on type " & GetType(T).Name & ".")
        End If

    End Sub

    Private Sub ApplyFilter()


        '_mUnfilteredListValue.Clear()

        '_mUnfilteredListValue.AddRange(Me.Items)

        Dim results As New List(Of T)()
        Dim col As New List(Of T)

        If Me.Items.Count = 0 Then
            col = _mUnfilteredListValue
        Else
            col = Me.Items
        End If


        If _mCriteria = "<>" Then
            For Each t As T In col

                results.Add(t)
            Next
        End If


        Dim propDesc As PropertyDescriptor = TypeDescriptor.GetProperties(GetType(T))(FilterPropertyName)

        If propDesc IsNot Nothing Then


            Dim tempResults As Integer = -1

            Do


                tempResults = FindCore(tempResults + 1, propDesc, FilterCompare, _mCriteria)
                'tempResults = FindCore(propDesc, FilterCompare)

                If tempResults <> -1 Then

                    If _mCriteria = "<>" Then
                        results.Remove(_mUnfilteredListValue(tempResults))
                    Else
                        results.Add(_mUnfilteredListValue(tempResults))
                    End If

                End If

            Loop While tempResults <> -1

        End If

        'Me.ClearItems()

        If _mCriteria = "<>" Then
            If results IsNot Nothing Then
                Me.ClearItems()
                For Each itemFound As T In results
                    Me.Add(itemFound)
                Next
            End If
        Else
            If results IsNot Nothing AndAlso results.Count > 0 Then


                For Each itemFound As T In results

                    If Me.IndexOf(itemFound) = -1 Then
                        Me.Add(itemFound)
                    End If
                Next
            End If
        End If

    End Sub

    Public Sub RemoveFilter()


        If Filter IsNot Nothing Then
            Filter = Nothing
        End If

    End Sub


#End Region

#Region "Searching"

    Protected Overloads Overrides ReadOnly Property SupportsSearchingCore() As Boolean


        Get

            Return False

        End Get
    End Property

    Protected Overloads Function FindCore(ByVal prop As PropertyDescriptor, ByVal key As Object) As Integer


        MyBase.FindCore(prop, key)



    End Function

    Protected Overloads Function FindCore(ByVal prop As PropertyDescriptor, ByVal key As Object, ByVal criteria As String) As Integer


        Return FindCore(0, prop, key, criteria)



    End Function

    Protected Overloads Function FindCore(ByVal startIndex As Integer, ByVal prop As PropertyDescriptor, ByVal key As Object, ByVal criteria As String) As Integer


        ' Get the property info for the specified property. 

        Dim propInfo As Reflection.PropertyInfo = GetType(T).GetProperty(prop.Name)

        Dim item As T

        If key IsNot Nothing Then


            ' Loop through the items to see if the key 

            ' value matches the property value. 

            For i As Integer = startIndex To _mUnfilteredListValue.Count - 1


                item = DirectCast(_mUnfilteredListValue(i), T)

                If criteria = "=" OrElse criteria = "<>" Then
                    If propInfo.GetValue(item, Nothing).Equals(key) Then

                        Return i
                    End If
                Else
                    If Not propInfo.GetValue(item, Nothing).Equals(key) Then

                        Return i
                    End If
                End If

            Next

        End If

        Return -1

    End Function

    Public Function Find(ByVal startIndex As Integer, ByVal [property] As String, ByVal key As Object, ByVal blnExactMatch As Boolean) As Integer


        ' Check the properties for a property with the specified name. 

        Dim propDesc As PropertyDescriptor = TypeDescriptor.GetProperties(GetType(T))([property])


        If propDesc Is Nothing OrElse key.ToString = "" Then

            Return -1

        Else

            Dim propInfo As Reflection.PropertyInfo = GetType(T).GetProperty(propDesc.Name)

            Dim item As T

            If key IsNot Nothing Then


                ' Loop through the items to see if the key 

                ' value matches the property value. 

                For i As Integer = 0 To Me.Items.Count - 1


                    item = DirectCast(Me.Items(i), T)

                    Dim converter As TypeConverter = TypeDescriptor.GetConverter(propDesc.PropertyType)

                    If blnExactMatch Then
                        If propInfo.GetValue(item, Nothing).Equals(key) Then

                            Return i
                        End If
                    Else
                        Dim DataValue As String = converter.ConvertToString(propInfo.GetValue(item, Nothing))

                        If UCase(DataValue).StartsWith(UCase(converter.ConvertToString(key))) Then
                            Return i
                        End If
                    End If


                Next

            End If


        End If


        Return -1

    End Function



#End Region


End Class



#End Region

#Region "Comparer"
Class PropertyComparer2(Of t)
    Implements IComparer(Of t)
    Private _mPropDesc As PropertyDescriptor = Nothing
    Private _mDirection As ListSortDirection = ListSortDirection.Ascending

    Public Sub New(ByVal propDesc As PropertyDescriptor, ByVal direction As ListSortDirection)
        _mPropDesc = propDesc
        _mDirection = direction
    End Sub

    Private Function Compare(ByVal x As t, ByVal y As t) As Integer Implements IComparer(Of t).Compare
        Dim xValue As Object = _mPropDesc.GetValue(x)
        Dim yValue As Object = _mPropDesc.GetValue(y)
        Return CompareValues(xValue, yValue, _mDirection)
    End Function


    Private Function CompareValues(ByVal xValue As Object, ByVal yValue As Object, ByVal direction As ListSortDirection) As Integer
        Dim retValue As Integer = 0
        If xValue Is Nothing AndAlso yValue Is Nothing Then
            Return 0
        End If
        If TypeOf xValue Is IComparable Then
            'can ask the x value 
            retValue = DirectCast(xValue, IComparable).CompareTo(yValue)
        ElseIf TypeOf yValue Is IComparable Then
            'can ask the y value 
            retValue = DirectCast(yValue, IComparable).CompareTo(xValue)
            'not comparable, compare string representations 
        ElseIf xValue Is Nothing OrElse yValue Is Nothing Then
            Return retValue
        ElseIf Not xValue.Equals(yValue) Then
            retValue = xValue.ToString().CompareTo(yValue.ToString())
        End If
        If direction = ListSortDirection.Ascending Then
            Return retValue
        Else
            Return retValue * -1
        End If

    End Function
End Class

#End Region


Public Class PropertyComparerCollection(Of T)
    Implements IComparer(Of T)
    Private ReadOnly m_sorts As ListSortDescriptionCollection
    Private ReadOnly comparers As PropertyComparer(Of T)()
    Public ReadOnly Property Sorts() As ListSortDescriptionCollection
        Get
            Exit Property
        End Get
    End Property
    Public Sub New(ByVal sorts As ListSortDescriptionCollection)
        If sorts Is Nothing Then
            Throw New ArgumentNullException("sorts")
        End If
        Me.m_sorts = sorts
        Dim list As New List(Of PropertyComparer(Of T))()
        For Each item As ListSortDescription In sorts
            list.Add(New PropertyComparer(Of T)(item.PropertyDescriptor, item.SortDirection = ListSortDirection.Descending))
        Next
        comparers = list.ToArray()
    End Sub
    Public ReadOnly Property PrimaryProperty() As PropertyDescriptor
        Get
            Exit Property
        End Get
    End Property
    Public ReadOnly Property PrimaryDirection() As ListSortDirection
        Get
            Exit Property
        End Get
    End Property

    Private Function Compare(ByVal x As T, ByVal y As T) As Integer Implements IComparer(Of T).Compare
        Dim result As Integer = 0
        For i As Integer = 0 To comparers.Length - 1
            result = comparers(i).Compare(x, y)
            If result <> 0 Then
                Exit For
            End If
        Next
        Return result
    End Function
End Class

Public Class PropertyComparer(Of T)
    Implements IComparer(Of T)
    Private ReadOnly m_descending As Boolean
    Public ReadOnly Property Descending() As Boolean
        Get
            Return m_descending
        End Get
    End Property
    Private ReadOnly m_property As PropertyDescriptor
    Public ReadOnly Property [Property]() As PropertyDescriptor
        Get
            Return m_property
        End Get
    End Property
    Public Sub New(ByVal [property] As PropertyDescriptor, ByVal descending As Boolean)
        If [property] Is Nothing Then
            Throw New ArgumentNullException("property")
        End If
        Me.m_descending = descending
        Me.m_property = [property]
    End Sub
    Public Function Compare(ByVal x As T, ByVal y As T) As Integer Implements IComparer(Of T).Compare
        ' todo; some null cases 
        Dim value As Integer
        If m_property.PropertyType Is GetType(Date) Then
            value = Comparer.[Default].Compare(Format(m_property.GetValue(x), "yyyyMMddHHmmss"), Format(m_property.GetValue(y), "yyyyMMddHHmmss"))
        Else
            value = Comparer.[Default].Compare(m_property.GetValue(x), m_property.GetValue(y))
        End If
        Return IIf(m_descending, -value, value)
    End Function
End Class






Public Class SortedEventArgs
    Inherits EventArgs
    Private _direction As System.ComponentModel.ListSortDirection = ListSortDirection.Ascending

    Public Sub New(direction As System.ComponentModel.ListSortDirection)
        _direction = direction
    End Sub
    Public Property Direction As System.ComponentModel.ListSortDirection
        Get
            Return _direction
        End Get
        Set(value As System.ComponentModel.ListSortDirection)
            _direction = value
        End Set
    End Property

End Class