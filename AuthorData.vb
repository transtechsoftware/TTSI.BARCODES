Imports System

Public Class AuthorData
    ''This is a test class for ObjectDataSource proof-of-concept

    ' Private Variables & Associated Properties

    ''m_name
    Private m_name As String = String.Empty
    Public Property Name() As String
        Get
            Return m_name
        End Get
        Set(ByVal Value As String)
            m_name = Value
        End Set
    End Property

    ''m_age
    Private m_age As Integer = 0
    Public Property Age() As Integer
        Get
            Return m_age
        End Get
        Set(ByVal Value As Integer)
            m_age = Value
        End Set
    End Property

    ''m_consultant
    Private m_consultant As Boolean = False
    Public Property Consultant() As Boolean
        Get
            Return m_consultant
        End Get
        Set(ByVal Value As Boolean)
            m_consultant = Value
        End Set
    End Property

    Sub New()
        'Constructor does nothing
    End Sub

End Class

Public Class AuthorDataODS

    Public Function GetData() As ICollection

        Dim list As New ArrayList
        Dim row As New AuthorData

        row.Name = "Sammy Nava"
        row.Age() = 41
        row.Consultant = True
        list.Add(row)

        row = New AuthorData

        row.Name = "Alejandra Grijalva"
        row.Age = 21
        row.Consultant = False
        list.Add(row)

        Return list

    End Function
End Class
