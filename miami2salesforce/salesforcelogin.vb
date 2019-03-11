Imports System.Net
Imports System.Web.Services.Protocols
Imports miami2salesforce.salesforce

Public Class SalesforceLogin

    Private _cfg As Config = Nothing

    Private _binding As SforceService
    Public ReadOnly Property Binding() As SforceService
        Get
            Return _binding
        End Get
    End Property

    Sub New(ByVal cfg As Config)
        _cfg = cfg
        _binding = New SforceService()

    End Sub

    Public Function login(ByRef errors As String) As Boolean

        Console.WriteLine("--- Login")
        Dim sf As SalesforceConfig = _cfg.SalesForceConfig
        Dim proxy As Proxy = _cfg.Proxy

        _binding.Timeout = 600000 'milliseconds
        _binding.Url = sf.AuthEndPoint

        If (proxy.IsUsed) Then
            If Not IsNothing(proxy.IP) Then
                Dim webProxy As New WebProxy(proxy.IP, proxy.Port)
                Dim credentials As New NetworkCredential(proxy.Username, proxy.Password)
                webProxy.Credentials = credentials
                _binding.Proxy = webProxy
            End If
        End If
        Try
            Dim loginResult As LoginResult = _binding.login(sf.Username, sf.Password + sf.Token)

            If loginResult.passwordExpired Then
                errors = "Login failed: password expired"
                Console.WriteLine(errors)
                Return False
            End If

            Console.WriteLine("--- Login OK")

            'Once the client application has logged in successfully, it will use
            'the results of the login call to reset the endpoint of the service
            'to the virtual server instance that is servicing your organization
            Dim authEndPoint As String = _binding.Url
            _binding.Url = loginResult.serverUrl

            'The sample client application now has an instance of the SforceService
            'that is pointing to the correct endpoint. Next, the sample client
            'application sets a persistent SOAP header (to be included on all
            'subsequent calls that are made with SforceService) that contains the
            'valid sessionId for our login credentials. To do this, the sample
            'client application creates a new SessionHeader object and persist it to
            'the SforceService. Add the session ID returned from the login to the session(header)
            _binding.SessionHeaderValue = New SessionHeader()
            _binding.SessionHeaderValue.sessionId = loginResult.sessionId
            Debug.Print("--- Login OK")

            'Dim getUserInfoResult As GetUserInfoResult = _binding.getUserInfo()
            'Debug.Print(getUserInfoResult.userFullName)
            'Debug.Print(getUserInfoResult.userEmail)

        Catch ex As SoapException
            errors = "!!! Login failed"
            Console.WriteLine(errors)
            errors += "<br/>" + ex.Message
            Console.WriteLine(ex.Message)
            Return False
        End Try

        Return True

    End Function

    Public Sub logout()
        Try
            _binding.logout()
            Console.WriteLine("--- Logout OK")
        Catch ex As SoapException
            Dim errors As String = "!!! Logout failed"
            Console.WriteLine(errors)
            errors += "<br/>" + ex.Message
            Console.WriteLine(ex.Message)
        End Try
    End Sub

End Class
