# clsBinding
Simple .NET class that makes working with collections allot easier
The class is fairly self explanatory but ive included an example anyway 

------------

### Currently Supported
- Add 
- Remove
- Sort
- Filter
- Search
- Compare

------------

### Example Code
simple sample code that creates 2 people, sorts them, removes one and then disposes of the collection
````vb
Public Class clsPerson

    Private _firstName As String = ""
    Private _lastName As String = ""

    Public Property firstName() As String
        Get
            Return _firstName
        End Get
        Set(ByVal value As String)
            _firstName = value
        End Set
    End Property

    Public Property lastName() As String
        Get
            Return _lastName
        End Get
        Set(ByVal value As String)
            _lastName = value
        End Set
    End Property

End Class

Public Class colPeople
    Inherits clsBindingListBase(Of clsPerson)

    Public Sub addNewPerson()
        Dim _person As New clsPerson 'create a new instance of clsPerson
        _person.firstName = "john" 'assign a first name
        _person.lastName = "doe" 'assign a last name
        Me.Add(_person) 'add john doe to the collection

        _person.firstName = "alan" 'assign a first name
        _person.lastName = "doe" 'assign a last name
        Me.Add(_person) 'add alan doe to the collection

        Me.Sort(_person.firstName) 'sort the collection by first name
        Me.Remove(_person) 'remove the last person added
        Me.Dispose() 'dispose of the collection
    End Sub

End Class
````
