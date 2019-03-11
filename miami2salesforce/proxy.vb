
Public Class Proxy
    Private _IP As String
    Public Property IP() As String
        Get
            Return _IP
        End Get
        Set(ByVal value As String)
            _IP = value
        End Set
    End Property

    Private _Port As Integer
    Public Property Port() As Integer
        Get
            Return _Port
        End Get
        Set(ByVal value As Integer)
            _Port = value
        End Set
    End Property

    Private _Username As String
    Public Property Username() As String
        Get
            Return _Username
        End Get
        Set(ByVal value As String)
            _Username = value
        End Set
    End Property

    Private _Password As String
    Public Property Password() As String
        Get
            Return _Password
        End Get
        Set(ByVal value As String)
            _Password = value
        End Set
    End Property

    Private _IsUsed As Boolean
    Public Property IsUsed() As Boolean
        Get
            Return _IsUsed
        End Get
        Set(ByVal value As Boolean)
            _IsUsed = value
        End Set
    End Property
End Class
