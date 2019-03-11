Imports miami2salesforce.salesforce

Public Class DataAccount
    Private _id As String
    Public Property Id() As String
        Get
            Return _id
        End Get
        Set(ByVal value As String)
            _id = value
        End Set
    End Property

    Private _type As String
    Public Property Type() As String
        Get
            Return _type
        End Get
        Set(ByVal value As String)
            _type = value
        End Set
    End Property

    Private _activeInactive As String
    Public Property ActiveInactive() As String
        Get
            Return _activeInactive
        End Get
        Set(ByVal value As String)
            _activeInactive = value
        End Set
    End Property


    Private _ownerId As String
    Public Property OwnerId() As String
        Get
            Return _ownerId
        End Get
        Set(ByVal value As String)
            _ownerId = value
        End Set
    End Property

    Private _status As String
    Public Property Status() As String
        Get
            Return _status
        End Get
        Set(ByVal value As String)
            _status = value
        End Set
    End Property

    Private _director As String
    Public Property Director() As String
        Get
            Return _director
        End Get
        Set(ByVal value As String)
            _director = value
        End Set
    End Property


    Private _vertical As String
    Public Property Vertical() As String
        Get
            Return _vertical
        End Get
        Set(ByVal value As String)
            _vertical = value
        End Set
    End Property

    Private _industry As String
    Public Property Industry() As String
        Get
            Return _industry
        End Get
        Set(ByVal value As String)
            _industry = value
        End Set
    End Property

    Private _ExternalId As String
    Public Property ExternalId() As String
        Get
            Return _ExternalId
        End Get
        Set(ByVal value As String)
            _ExternalId = value
        End Set
    End Property

    Private _Owner As User
    Public Property Owner() As User
        Get
            Return _Owner
        End Get
        Set(ByVal value As User)
            _Owner = value
        End Set
    End Property

End Class
