Imports miami2salesforce.salesforce
Imports Methodes.mthconnexion
Imports System.Text
Public Class loadUser
    Private _binding As SforceService
    'Private _miamiods As DataBase
    'Private _miamigate As DataBase

    'Private _oracleConnectionODS As OracleConnection
    'Private _oracleCommandODS As OracleCommand
    'Private _oracleConnectionGATE As OracleConnection
    'Private _oracleCommandGATE As OracleCommand
    'Private _oracleDataReader As OracleDataReader
    Private _oracleConnectionODS As MthConnexion
    Private _oracleConnectionGATE As MthConnexion

    Sub New(ByVal binding As SforceService, ByVal miamigate As String, ByVal miamiods As String)
        _binding = binding
        '_miamiods = New DataBase()
        '_miamiods.ConnectionString = My.Settings.miamiods
        '_miamigate = New DataBase()
        '_miamigate.ConnectionString = My.Settings.miamigate
        _oracleConnectionGATE = New MthConnexion(miamigate)
        _oracleConnectionODS = New MthConnexion(miamiods)
    End Sub

    Function loadUser(ByVal numberOfDays As Integer) As String
        Dim result As String = String.Empty
        Dim err As String = String.Empty

        '_oracleConnectionODS = _miamiods.connect()
        '_oracleConnectionGATE = _miamigate.connect()

        Try
            '*******
            'Open DB
            '*******
            '_oracleConnectionODS.Open()
            '_oracleCommandODS = New OracleCommand()
            '_oracleCommandODS.Connection = _oracleConnectionODS
            '_oracleCommandODS.CommandType = CommandType.Text

            '_oracleConnectionGATE.Open()
            '_oracleCommandGATE = New OracleCommand()
            '_oracleCommandGATE.Connection = _oracleConnectionGATE
            '_oracleCommandGATE.CommandType = CommandType.Text

            '**********
            'References
            '**********
            'err = _references()
            'If Not String.IsNullOrEmpty(err) Then
            '    Return "ERRORS: <br/>" + err
            'End If

            '********************
            'user
            '********************
            Dim lastModifiedDate As Date
            If numberOfDays > 0 Then
                lastModifiedDate = Date.Now.AddDays(-numberOfDays)
            End If

            'truncate tables
            err = _truncateODSTables()
            result += err

            err = _queryUser(lastModifiedDate)
            result += err

            '********
            'Close DB
            '********
            _oracleConnectionODS.Dispose()
            _oracleConnectionGATE.Dispose()

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            result = "** ERROR **<br/><br/>" + ex.Message
        End Try

        Return result
    End Function

    Private Function _queryUser(ByVal lastModifiedDate As Date) As String
        Dim errors As String
        Dim sb As StringBuilder = New StringBuilder()

        Try
            Dim done As Boolean = False
            Dim query As String = "Select Id, Username, Name, CompanyName, Division, Department, Title, Street, City, Country, Email, LastLoginDate, CreatedDate, CreatedById, LastModifiedDate FROM User "
            If lastModifiedDate > Date.MinValue Then
                Dim sLastModifiedDate As String = lastModifiedDate.ToString("yyyy-MM-ddThh:mm:ssZ")
                Dim where As String = String.Format(" WHERE LastModifiedDate >= {0}", sLastModifiedDate)
                query = query + where
            End If
            Dim result As QueryResult = _binding.query(query)
            If result.size > 0 Then
                Console.WriteLine(String.Format("# User: {0}", result.size))
                Dim parameters As TableauParametres = New TableauParametres()

                While Not done
                    Dim objects() As sObject = result.records
                    Dim count = objects.Length
                    For i As Integer = 0 To count - 1
                        Dim User As User = objects(i)


                        'Dim recordType As RecordType = Global_potential_account.RecordType
                        Debug.WriteLine(String.Format("{0} - {1}", User.Name, User.LastModifiedDate))
                        Console.WriteLine(String.Format("{0} - {1}", User.Name, User.LastModifiedDate))

                        If i >= 107 Then
                            i = i
                        End If
                        Dim insert As String = "Insert Into TMPSF_USER" + _
                "(ID_USER, USERNAME, NAME, COMPANYNAME, DIVISION, DEPARTMENT, TITLE, STREET, CITY, COUNTRY, MOBILE_PHONE, EMAIL, LASTLOGINDATE, CREATEDDATE, CREATEDBYID, LASTMODIFIEDDATE)" + _
                " Values(:ID_USER, :USERNAME, :NAME, :COMPANYNAME, :DIVISION, :DEPARTMENT, :TITLE, :STREET, :CITY, :COUNTRY, :MOBILE_PHONE, :EMAIL, :LASTLOGINDATE, :CREATEDDATE, :CREATEDBYID, :LASTMODIFIEDDATE)"



                        parameters.PurgeParametre()
                        parameters.AjouterParametreChaine(":ID_USER", User.Id)
                        If Not IsNothing(User.Username) Then
                            parameters.AjouterParametreChaine(":USERNAME", User.Username)
                        Else
                            parameters.AjouterParametre(":USERNAME", String.Empty)
                        End If

                        If Not IsNothing(User.Name) Then
                            parameters.AjouterParametreChaine(":NAME", User.Name)
                        Else
                            parameters.AjouterParametre(":NAME", String.Empty)
                        End If

                        If Not IsNothing(User.CompanyName) Then
                            parameters.AjouterParametreChaine(":COMPANYNAME", User.CompanyName)
                        Else
                            parameters.AjouterParametre(":COMPANYNAME", String.Empty)
                        End If

                        If Not IsNothing(User.Division) Then
                            parameters.AjouterParametreChaine(":DIVISION", User.Division)
                        Else
                            parameters.AjouterParametre(":DIVISION", String.Empty)
                        End If

                        If Not IsNothing(User.Department) Then
                            parameters.AjouterParametreChaine(":DEPARTMENT", User.Department)
                        Else
                            parameters.AjouterParametre(":DEPARTMENT", String.Empty)
                        End If

                        If Not IsNothing(User.Title) Then
                            parameters.AjouterParametreChaine(":TITLE", User.Title)
                        Else
                            parameters.AjouterParametre(":TITLE", String.Empty)
                        End If

                        If Not IsNothing(User.Street) Then
                            parameters.AjouterParametreChaine(":STREET", User.Street)
                        Else
                            parameters.AjouterParametre(":STREET", String.Empty)
                        End If

                        If Not IsNothing(User.City) Then
                            parameters.AjouterParametreChaine(":CITY", User.City)
                        Else
                            parameters.AjouterParametre(":CITY", String.Empty)
                        End If

                        If Not IsNothing(User.Country) Then
                            parameters.AjouterParametreChaine(":COUNTRY", User.Country)
                        Else
                            parameters.AjouterParametre(":COUNTRY", String.Empty)
                        End If

                        If Not IsNothing(User.MobilePhone) Then
                            parameters.AjouterParametreChaine(":MOBILE_PHONE", User.MobilePhone)
                        Else
                            parameters.AjouterParametre(":MOBILE_PHONE", String.Empty)
                        End If

                        If Not IsNothing(User.Email) Then
                            parameters.AjouterParametreChaine(":EMAIL", User.Email)
                        Else
                            parameters.AjouterParametre(":EMAIL", String.Empty)
                        End If

                        If Not IsNothing(User.LastLoginDate) Then
                            parameters.AjouterParametreChaine(":LASTLOGINDATE", User.LastLoginDate)
                        Else
                            parameters.AjouterParametre(":LASTLOGINDATE", String.Empty)
                        End If

                        If Not IsNothing(User.CreatedDate) Then
                            parameters.AjouterParametreChaine(":CREATEDDATE", User.CreatedDate)
                        Else
                            parameters.AjouterParametre(":CREATEDDATE", String.Empty)
                        End If

                        If Not IsNothing(User.CreatedById) Then
                            parameters.AjouterParametreChaine(":CREATEDBYID", User.CreatedById)
                        Else
                            parameters.AjouterParametre(":CREATEDBYID", String.Empty)
                        End If

                        If Not IsNothing(User.LastModifiedDate) Then
                            parameters.AjouterParametreChaine(":LASTMODIFIEDDATE", User.LastModifiedDate)
                        Else
                            parameters.AjouterParametre(":LASTMODIFIEDDATE", String.Empty)
                        End If

                        Dim sqlError As String = _oracleConnectionODS.Requete(insert, parameters)
                    Next
                    If result.done Then
                        done = True
                    Else

                        result = _binding.queryMore(result.queryLocator)
                    End If
                End While
            Else
                Console.WriteLine("No User found in Salesforce")
                sb.Append("No User found in Salesforce")
            End If

        Catch ex As Exception
            sb.Append(ex.Message)
            Console.WriteLine(ex.Message)
        End Try

        errors = sb.ToString()

        Return errors
    End Function

    'Private Function _references() As String
    '    Dim sql As String
    '    Dim err As New StringBuilder

    '    Return err.ToString
    'End Function
   
    Private Function _truncateODSTables() As String
        Dim errors As String
        Dim sb As StringBuilder = New StringBuilder()
        Dim tableList As New List(Of String)
        'tableList.Add("TMPSF_ATT_POT_RAND")
        'tableList.Add("TMPSF_CONTACTS")
        'tableList.Add("TMPSF_GLO_POT_ACC")
        'tableList.Add("TMPSF_LEADS")
        tableList.Add("TMPSF_USER")
        'tableList.Add("TMPSF_OPCO_CTC_MATX")
        'tableList.Add("TMPSF_OPPORTUNITY")
        'tableList.Add("TMPSF_OPPORTUNITY_HIST")
        'tableList.Add("TMPSF_RSR_OPP_QUALIF")
        'tableList.Add("TMPSF_SRV_OPCO_FRAM")

        errors = Utils.TruncateODSTables(tableList, _oracleConnectionODS)

        'Try
        '    For Each table As String In tableList
        '        'Dim sql As String = String.Format("Truncate Table {0}", table)
        '        Dim sql As String = String.Format("Delete From {0} Where 1 = 1", table)
        '        Dim sqlError As String = _oracleConnectionODS.Requete(sql)
        '    Next

        'Catch ex As Exception
        '    sb.Append(ex.Message)
        '    Console.WriteLine(ex.Message)
        'End Try

        'errors = sb.ToString()

        Return errors
    End Function

End Class

