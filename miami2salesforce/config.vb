Imports System.Xml
Imports System.IO

Public Class Config

    Private _xmlPath As String
    Private _environment As String
    Public ReadOnly Property Environment() As String
        Get
            Return _environment
        End Get
    End Property

    Private _salesForceConfig As SalesforceConfig
    Public ReadOnly Property SalesForceConfig() As SalesforceConfig
        Get
            Return _salesForceConfig
        End Get
    End Property

    Private _proxy As Proxy
    Public ReadOnly Property Proxy() As Proxy
        Get
            Return _proxy
        End Get
    End Property

    'Constructors
    Sub New()
        Me._xmlPath = Nothing
        Me._environment = My.Settings.Environment

        'Proxy
        Me._proxy = New Proxy()
        Me._proxy.IsUsed = True
        If String.IsNullOrEmpty(My.Settings.ProxyIP) Then
            Me._proxy.IP = My.Settings.ProxyIP
        End If
        Dim ip As String = My.Settings.ProxyPort
        Me._proxy.Port = Integer.Parse(ip)

        Dim cryptedUserName As String = My.Settings.ProxyUserName
        Dim cryptedPassword As String = My.Settings.ProxyPassword
        Me._proxy.Username = CryptoService.Decrypt(cryptedUserName)
        Me._proxy.Password = CryptoService.Decrypt(cryptedPassword)

        'Salesforce
        Me._salesForceConfig = New SalesforceConfig()
        Me._salesForceConfig.AuthEndPoint = My.Settings.SFAuthEndPoint

        cryptedUserName = My.Settings.SFUserName
        cryptedPassword = My.Settings.SFPassword

        Me._salesForceConfig.Username = CryptoService.Decrypt(cryptedUserName)
        Me._salesForceConfig.Password = CryptoService.Decrypt(cryptedPassword)
        Me._salesForceConfig.Token = My.Settings.SFToken

    End Sub
    Sub New(ByVal xmlPath As String)
        Me._xmlPath = xmlPath
    End Sub

    Public Sub readConfig()
        'check path
        If IsNothing(_xmlPath) Then
            Console.WriteLine("xml file name not set")
            Return
        End If
        If _xmlPath = String.Empty Then
            Console.WriteLine("xml file name is empty")
            Return
        End If
        'check file
        If Not File.Exists(_xmlPath) Then
            Console.WriteLine("xml file name doesn't exists: {0}", _xmlPath)
            Return
        End If

        Dim element As String = String.Empty
        Dim is_miamisf As Boolean = False
        Dim is_env As Boolean = False
        Dim is_sf As Boolean = False
        Dim is_proxy As Boolean = False
        Dim is_miamidwh As Boolean = False
        Dim is_account As Boolean = False

        'read xml file
        Dim reader As New XmlTextReader(_xmlPath)
        While (reader.Read())
            Select Case reader.NodeType
                Case XmlNodeType.Element
                    element = reader.Name
                    Debug.Print(element)
                    Select Case element
                        Case "miamisf"
                            is_miamisf = True
                        Case "env"
                            is_env = True
                        Case "sf"
                            is_sf = True
                            Me._salesForceConfig = New SalesforceConfig()
                        Case "proxy"
                            is_proxy = True
                            Me._proxy = New Proxy()
                        Case "account"
                            is_account = True
                    End Select
                    While (reader.MoveToNextAttribute())
                        Dim value As String = reader.Value
                        Debug.Print(reader.Name + " => " + reader.Value)
                        Select Case reader.Name
                            Case "authEndPoint"
                                If is_sf Then
                                    Me._salesForceConfig.AuthEndPoint = value
                                End If
                            Case "ip"
                                If is_proxy Then
                                    Me._proxy.IP = value
                                End If
                            Case "port"
                                If is_proxy Then
                                    Me._proxy.Port = Integer.Parse(value)
                                End If
                            Case "use"
                                If is_proxy Then
                                    If value.ToLower = "yes" Then
                                        Me._proxy.IsUsed = True
                                    Else
                                        Me._proxy.IsUsed = False
                                    End If
                                End If
                            Case "username"
                                If is_sf AndAlso is_account Then
                                    Me._salesForceConfig.Username = CryptoService.Decrypt(value)
                                End If
                                If is_proxy AndAlso is_account Then
                                    Me._proxy.Username = CryptoService.Decrypt(value)
                                End If
                            Case "password"
                                If is_sf AndAlso is_account Then
                                    Me._salesForceConfig.Password = CryptoService.Decrypt(value)
                                End If
                                If is_proxy AndAlso is_account Then
                                    Me._proxy.Password = CryptoService.Decrypt(value)
                                End If
                            Case "token"
                                If is_sf AndAlso is_account Then
                                    Me._salesForceConfig.Token = value
                                End If
                        End Select
                    End While
                    '<name>
                Case XmlNodeType.Text
                    Dim value As String = reader.Value
                    Debug.Print(value)
                    If is_env Then
                        Me._environment = value
                    End If
                Case XmlNodeType.EndElement
                    element = reader.Name
                    Debug.Print("/" + element)
                    Select Case element
                        Case "miamisf"
                            is_miamisf = False
                        Case "env"
                            is_env = False
                        Case "sf"
                            is_sf = False
                        Case "proxy"
                            is_proxy = False
                        Case "account"
                            is_account = False
                    End Select
                    '</name>
            End Select
        End While
        reader.Close()

        Debug.Print("TEST Objects")
        Debug.Print("_environment:" + Me._environment)
        Debug.Print("_salesForceConfig.AuthEndPoint:" + Me._salesForceConfig.AuthEndPoint)
        Debug.Print("_salesForceConfig.Username:" + Me._salesForceConfig.Username)
        Debug.Print("_salesForceConfig.Password:" + Me._salesForceConfig.Password)
        Debug.Print("_salesForceConfig.Token:" + Me._salesForceConfig.Token)
        Debug.Print("_proxy.IP:" + Me._proxy.IP)
        Debug.Print("_proxy.Port:" + Me._proxy.Port.ToString())
        Debug.Print("_proxy.Username:" + Me._proxy.Username)
        Debug.Print("_proxy.Password:" + Me._proxy.Password)
        Debug.Print("OK")

    End Sub

End Class
