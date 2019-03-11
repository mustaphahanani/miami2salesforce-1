
Public Class SalesforceConfig
    Private _AuthEndPoint As String
    Public Property AuthEndPoint() As String
        Get
            Return _AuthEndPoint
        End Get
        Set(ByVal value As String)
            _AuthEndPoint = value
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

    Private _Token As String
    Public Property Token() As String
        Get
            Return _Token
        End Get
        Set(ByVal value As String)
            _Token = value
        End Set
    End Property

End Class
