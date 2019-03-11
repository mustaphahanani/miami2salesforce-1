Imports miami2salesforce.salesforce
Imports Methodes.mthconnexion
Imports System.Text

Public Class LoadOpportunity_history

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

    Function LoadOpportunity_history(ByVal numberOfDays As Integer) As String
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
            'Opportunity History
            '********************
            'Dim lastModifiedDate As Date
            'If numberOfDays > 0 Then
            '    lastModifiedDate = Date.Now.AddDays(-numberOfDays)
            'End If

            'truncate tables
            err = _truncateODSTables()
            result += err

            err = _queryOpportunity_history()
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


    Private Function _queryOpportunity_history() As String
        Dim errors As String
        Dim sb As StringBuilder = New StringBuilder()

        Try
            Dim done As Boolean = False
            'retrieve the first 500 opportunities
            Dim query As String = "Select Id, OpportunityId, CreatedById, CreatedDate, StageName, " + _
            "Amount, ExpectedRevenue, CloseDate, Probability, ForecastCategory, SystemModstamp, IsDeleted FROM OpportunityHistory"
            Dim result As QueryResult = _binding.query(query)
            If result.size > 0 Then
                Console.WriteLine(String.Format("# Opportunity_history: {0}", result.size))
                Dim parameters As TableauParametres = New TableauParametres()

                While Not done
                    Dim objects() As sObject = result.records
                    Dim count = objects.Length

                    For i As Integer = 0 To count - 1
                        Dim Opportunity_History As OpportunityHistory = objects(i)


                        'Dim recordType As RecordType = Attainable_potential_Randstad.RecordType
                        'Debug.WriteLine(String.Format("{0} - {1}", Opportunity_History., Opportunity_History.LastModifiedDate))
                        'Console.WriteLine(String.Format("{0} - {1}", Opportunity_History.Name, Opportunity_History.LastModifiedDate))

                        If i >= 107 Then
                            i = i
                        End If
                        Dim insert As String = "Insert Into TMPSF_OPPORTUNITY_HIST" + _
                        "(ID_OPP_HIT, ID_OPPORTUNITY, CREATEDBYID, CREATEDDATE, STAGENAME," + _
                        "AMOUNT, EXPECTEDREVENUE, CLOSEDATE, PROBABILITY, FORECASTCATEGORY, SYSTEMMODSTAMP, ISDELETED)" + _
                        " Values(:ID, :OPPORTUNITYID, :CREATEDBYID, :CREATEDDATE, :STAGENAME," + _
                        ":AMOUNT, :EXPECTEDREVENUE, :CLOSEDATE, :PROBABILITY, :FORECASTCATEGORY, :SYSTEMMODSTAMP, :ISDELETED)"

                        parameters.PurgeParametre()
                        If Not IsNothing(Opportunity_History.Id) Then
                            parameters.AjouterParametreChaine(":ID_OPP_HIT", Opportunity_History.Id)
                        Else
                            parameters.AjouterParametreChaine(":ID_OPP_HIT", Opportunity_History.Id)
                        End If
                        If Not IsNothing(Opportunity_History.OpportunityId) Then
                            parameters.AjouterParametreChaine(":ID_OPPORTUNITY", Opportunity_History.OpportunityId)
                        Else
                            parameters.AjouterParametre(":ID_OPPORTUNITY", String.Empty)
                        End If

                        If Not IsNothing(Opportunity_History.CreatedById) Then
                            parameters.AjouterParametreChaine(":CREATEDBYID", Opportunity_History.CreatedById)
                        Else
                            parameters.AjouterParametre(":CREATEDBYID", String.Empty)
                        End If

                        If Not IsNothing(Opportunity_History.CreatedDate) Then
                            parameters.AjouterParametreChaine(":CREATEDDATE", Opportunity_History.CreatedDate)
                        Else
                            parameters.AjouterParametre(":CREATEDDATE", String.Empty)
                        End If

                        If Not IsNothing(Opportunity_History.StageName) Then
                            parameters.AjouterParametreChaine(":STAGENAME", Opportunity_History.StageName)
                        Else
                            parameters.AjouterParametre(":STAGENAME", String.Empty)
                        End If

                        If Not IsNothing(Opportunity_History.Amount) Then
                            parameters.AjouterParametreChaine(":AMOUNT", Opportunity_History.Amount)
                        Else
                            parameters.AjouterParametre(":AMOUNT", String.Empty)
                        End If


                        If Not IsNothing(Opportunity_History.ExpectedRevenue) Then
                            parameters.AjouterParametreChaine(":EXPECTEDREVENUE", Opportunity_History.ExpectedRevenue)
                        Else
                            parameters.AjouterParametre(":EXPECTEDREVENUE", String.Empty)
                        End If

                        If Not IsNothing(Opportunity_History.CloseDate) Then
                            parameters.AjouterParametreChaine(":CLOSEDATE", Opportunity_History.CloseDate)
                        Else
                            parameters.AjouterParametre(":CLOSEDATE", String.Empty)
                        End If

                        If Not IsNothing(Opportunity_History.Probability) Then
                            parameters.AjouterParametreChaine(":PROBABILITY", Opportunity_History.Probability)
                        Else
                            parameters.AjouterParametre(":PROBABILITY", String.Empty)
                        End If

                        If Not IsNothing(Opportunity_History.ForecastCategory) Then
                            parameters.AjouterParametreChaine(":FORECASTCATEGORY", Opportunity_History.ForecastCategory)
                        Else
                            parameters.AjouterParametre(":FORECASTCATEGORY", String.Empty)
                        End If

                        If Not IsNothing(Opportunity_History.SystemModstamp) Then
                            parameters.AjouterParametreChaine(":SYSTEMMODSTAMP", Opportunity_History.SystemModstamp)
                        Else
                            parameters.AjouterParametre(":SYSTEMMODSTAMP", String.Empty)
                        End If

                        If Not IsNothing(Opportunity_History.IsDeleted) Then
                            parameters.AjouterParametreChaine(":ISDELETED", Opportunity_History.IsDeleted)
                        Else
                            parameters.AjouterParametre(":ISDELETED", String.Empty)
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
                Console.WriteLine("No Opportunity History found in Salesforce")
                sb.Append("No Opportunity History found in Salesforce")
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
        'tableList.Add("TMPSF_OPCO_CTC_MATX")
        'tableList.Add("TMPSF_OPPORTUNITY")
        tableList.Add("TMPSF_OPPORTUNITY_HIST")
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
