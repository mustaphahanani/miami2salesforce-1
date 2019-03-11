Imports miami2salesforce.salesforce
Imports Methodes.mthconnexion
Imports System.Text

Public Class loadProduct2
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

    Function loadProduct2(ByVal numberOfDays As Integer) As String
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
            'Return "ERRORS: <br/>" + err
            'End If

            '********************
            'OPCO Contact matrix
            '********************
            Dim lastModifiedDate As Date
            If numberOfDays > 0 Then
                lastModifiedDate = Date.Now.AddDays(-numberOfDays)
            End If

            'truncate tables
            err = _truncateODSTables()
            result += err

            err = _queryloadProduct2(lastModifiedDate)
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

    Private Function _queryloadProduct2(ByVal lastModifiedDate As Date) As String
        Dim errors As String
        Dim sb As StringBuilder = New StringBuilder()

        Try
            Dim done As Boolean = False
            Dim query As String = "Select Id, Name, ProductCode, Description, IsActive, CreatedDate, CreatedById, LastModifiedDate, SystemModstamp, Family, IsDeleted FROM Product2 "
            If lastModifiedDate > Date.MinValue Then
                Dim sLastModifiedDate As String = lastModifiedDate.ToString("yyyy-MM-ddThh:mm:ssZ")
                Dim where As String = String.Format(" WHERE LastModifiedDate >= {0}", sLastModifiedDate)
                query = query + where
            End If
            Dim result As QueryResult = _binding.query(query)
            If result.size > 0 Then
                Console.WriteLine(String.Format("# Product: {0}", result.size))
                Dim parameters As TableauParametres = New TableauParametres()

                While Not done
                    Dim objects() As sObject = result.records
                    Dim count = objects.Length
                    For i As Integer = 0 To count - 1
                        Dim Product As Product2 = objects(i)


                        'Dim recordType As RecordType = Global_potential_account.RecordType
                        Debug.WriteLine(String.Format("{0} - {1}", Product.Name, Product.LastModifiedDate))
                        Console.WriteLine(String.Format("{0} - {1}", Product.Name, Product.LastModifiedDate))

                        If i >= 107 Then
                            i = i
                        End If
                        Dim insert As String = "Insert Into TMPSF_PRODUCT" + _
                "(ID_PRODUCT, Name, CD_PRODUCT, DESCRIPTION, IS_ACTIVE, SYSTEMMODSTAMP, CREATEDDATE, CREATEDBYID, FAMILY, IS_DELETED)" + _
                " Values(:ID_PRODUCT, :Name, :CD_PRODUCT, :DESCRIPTION, :IS_ACTIVE, :SYSTEMMODSTAMP, :CREATEDDATE, :CREATEDBYID, :FAMILY,: IS_DELETED)"

                        parameters.PurgeParametre()
                        parameters.AjouterParametreChaine(":ID_PRODUCT", Product.Id)
                        If Not IsNothing(Product.Name) Then
                            parameters.AjouterParametreChaine(":Name", Product.Name)
                        Else
                            parameters.AjouterParametre(":Name", String.Empty)
                        End If

                        If Not IsNothing(Product.ProductCode) Then
                            parameters.AjouterParametreChaine(":CD_PRODUCT", Product.ProductCode)
                        Else
                            parameters.AjouterParametre(":CD_PRODUCT", String.Empty)
                        End If

                        If Not IsNothing(Product.Description) Then
                            parameters.AjouterParametreChaine(":DESCRIPTION", Product.Description)
                        Else
                            parameters.AjouterParametre(":DESCRIPTION", String.Empty)
                        End If

                        If Not IsNothing(Product.IsActive) Then
                            parameters.AjouterParametreChaine(":IS_ACTIVE", Product.IsActive)
                        Else
                            parameters.AjouterParametre(":IS_ACTIVE", String.Empty)
                        End If

                        If Not IsNothing(Product.SystemModstamp) Then
                            parameters.AjouterParametreChaine(":SYSTEMMODSTAMP", Product.SystemModstamp)
                        Else
                            parameters.AjouterParametre(":SYSTEMMODSTAMP", String.Empty)
                        End If

                        If Not IsNothing(Product.CreatedDate) Then
                            parameters.AjouterParametreChaine(":CREATEDDATE", Product.CreatedDate)
                        Else
                            parameters.AjouterParametre(":CREATEDDATE", String.Empty)
                        End If

                        If Not IsNothing(Product.CreatedById) Then
                            parameters.AjouterParametreChaine(":CREATEDBYID", Product.CreatedById)
                        Else
                            parameters.AjouterParametre(":CREATEDBYID", String.Empty)
                        End If

                        If Not IsNothing(Product.Family) Then
                            parameters.AjouterParametreChaine(":FAMILY(", Product.Family)
                        Else
                            parameters.AjouterParametre(":FAMILY(", String.Empty)
                        End If
                        If Not IsNothing(Product.IsDeleted) Then
                            parameters.AjouterParametreChaine(":IS_DELETED", Product.IsDeleted)
                        Else
                            parameters.AjouterParametre(":IS_DELETED", String.Empty)
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
                Console.WriteLine("No Product found in Salesforce")
                sb.Append("No Product found in Salesforce")
            End If

        Catch ex As Exception
            sb.Append(ex.Message)
            Console.WriteLine(ex.Message)
        End Try

        errors = sb.ToString()

        Return errors
    End Function
    Private Function _truncateODSTables() As String
        Dim errors As String
        Dim sb As StringBuilder = New StringBuilder()
        Dim tableList As New List(Of String)
        'tableList.Add("TMPSF_ATT_POT_RAND")
        'tableList.Add("TMPSF_CONTACTS")
        'tableList.Add("TMPSF_GLO_POT_ACC")
        'tableList.Add("TMPSF_LEADS")
        tableList.Add("TMPSF_PRODUCT")
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

