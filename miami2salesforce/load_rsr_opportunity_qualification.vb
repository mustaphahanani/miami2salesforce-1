Imports miami2salesforce.salesforce
Imports Methodes.mthconnexion
Imports System.Text



Public Class load_RSR_Opportunity_Qualification

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



    Private _SF_RSROPP_REC_TYP As New Dictionary(Of String, String)
    Private _SF_RSROPP_ANPLAC As New Dictionary(Of String, String)
    Private _SF_RSROPP_GESCOPP As New Dictionary(Of String, String)
    Private _SF_RSROPP_LEAD_RG As New Dictionary(Of String, String)
    Private _SF_RSROPP_MSP_CMY As New Dictionary(Of String, String)
    Private _SF_RSROPP_MSP_ENG As New Dictionary(Of String, String)
    Private _SF_RSROPP_PRIORIT As New Dictionary(Of String, String)
    Private _SF_RSROPP_QUALIF As New Dictionary(Of String, String)
    Private _SF_RSROPP_RPO As New Dictionary(Of String, String)
    Private _SF_RSROPP_TRANBUS As New Dictionary(Of String, String)
    Private _SF_RSROPP_VMS_ATS As New Dictionary(Of String, String)
    

    Sub New(ByVal binding As SforceService, ByVal miamigate As String, ByVal miamiods As String)
        _binding = binding
        '_miamiods = New DataBase()
        '_miamiods.ConnectionString = My.Settings.miamiods
        '_miamigate = New DataBase()
        '_miamigate.ConnectionString = My.Settings.miamigate
        _oracleConnectionGATE = New MthConnexion(miamigate)
        _oracleConnectionODS = New MthConnexion(miamiods)
    End Sub

    Function loadRSR_Opportunity_Qualification(ByVal numberOfDays As Integer) As String
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
            err = _references()
            If Not String.IsNullOrEmpty(err) Then
                Return "ERRORS: <br/>" + err
            End If

            '********************
            'RSR Opportunity Qualification
            '********************
            Dim lastModifiedDate As Date
            If numberOfDays > 0 Then
                lastModifiedDate = Date.Now.AddDays(-numberOfDays)
            End If

            'truncate tables
            err = _truncateODSTables()
            result += err

            '      err = _queryRSR_Opportunity_Qualification(lastModifiedDate)
            '     result += err

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
    'Private Function _queryRSR_Opportunity_Qualification(ByVal lastModifiedDate As Date) As String
    '    Dim errors As String
    '    Dim sb As StringBuilder = New StringBuilder()

    '    Try
    '        Dim done As Boolean = False
    '        Dim query As String = "SELECT ID, Name, LastModifiedDate, Account_partner__c, An__c, Annual_Placements__c, RecordType.Name, Contact__c, " + _
    '        "Delivery_Date__c, expected_go_live_date_project__c, Expected_start_date_project__c,Fee_Randstad_opco_s__c, " + _
    '        "Fee_suppliers_3rd_party__c, GEO_Scope_OPP__c, Leading_RSR_region__c, Managed_Spend__c, MSP_Company__c, " + _
    '        "MSP_Engagement_type__c, Opportunity_GCS__c, Priority__c, Qualified__c, RPO_Eng__c, RSR_Implementation_Responsible__c, " + _
    '        "Sourceright_delivery_responsible__c, Temp_Population__c, Transition_positions__c, Transition_Business__c,VMS_Software__c, CreatedDate, CreatedById FROM RSR_Opportunity_Qualification__c"

    '        If lastModifiedDate > Date.MinValue Then
    '            Dim sLastModifiedDate As String = lastModifiedDate.ToString("yyyy-MM-ddThh:mm:ssZ")
    '            Dim where As String = String.Format(" WHERE LastModifiedDate >= {0}", sLastModifiedDate)
    '            query = query + where
    '        End If
    '        Dim result As QueryResult = _binding.query(query)
    '        If result.size > 0 Then
    '            Console.WriteLine(String.Format("# RSR_Opportunity_Qualification: {0}", result.size))
    '            Dim parameters As TableauParametres = New TableauParametres()

    '            While Not done
    '                Dim objects() As sObject = result.records
    '                Dim count = objects.Length
    '                For i As Integer = 0 To count - 1
    '                    Dim RSR_Opportunity_Qualification As RSR_Opportunity_Qualification__c = objects(i)


    '                    Dim recordType As RecordType = RSR_Opportunity_Qualification.RecordType
    '                    Debug.WriteLine(String.Format("{0} - {1}", RSR_Opportunity_Qualification.Name, RSR_Opportunity_Qualification.LastModifiedDate))
    '                    Console.WriteLine(String.Format("{0} - {1}", RSR_Opportunity_Qualification.Name, RSR_Opportunity_Qualification.LastModifiedDate))

    '                    If i >= 107 Then
    '                        i = i
    '                    End If
    '                    Dim insert As String = "Insert Into TMPSF_RSR_OPP_QUALIF" + _
    '            "(ID_RSR_OPP, RSROPP_NAME, IA_CD_PARTNER, ANNUAL_HIRES, ANNUAL_PLACEMENTS, RECORD_TYPE_RSR, CONTACT, " + _
    '            "DELIVERY_DATE, GOLIVE_PROJECT_DATE, START_PROJECT_DATE, FEE_RAND_OPCO, " + _
    '            "FEE_SUPPLIERS, GEO_SCOPE_OPP, LEADING_RSR_REGION, MANAGED_SPEND, MSP_COMPANY, " + _
    '            "MSP_ENGAGEMENT_TYPE, OPPORTUNITY_GCS, PRIORITY, QUALIFIED, RPO_ENGAG_TYPE, " + _
    '            "CONTACT_RSRIMP_RES, RSR_DELIVERY_RES, TEMP_POPULATION, TRANS_POSITIONS, TRANS_BUSINESS, VMS_ATS_SOFT,CREATEDDATE, CREATEDBYID)" + _
    '            " Values(:ID_RSR_OPP, :RSROPP_NAME, :IA_CD_PARTNER, :ANNUAL_HIRES, :ANNUAL_PLACEMENTS, :RECORD_TYPE_RSR, :CONTACT," + _
    '            ":DELIVERY_DATE, :GOLIVE_PROJECT_DATE, :START_PROJECT_DATE, :FEE_RAND_OPCO, " + _
    '            ":FEE_SUPPLIERS, :GEO_SCOPE_OPP, :LEADING_RSR_REGION, :MANAGED_SPEND, :MSP_COMPANY, " + _
    '            ":MSP_ENGAGEMENT_TYPE, :OPPORTUNITY_GCS, :PRIORITY, :QUALIFIED, :RPO_ENGAG_TYPE, " + _
    '            ":CONTACT_RSRIMP_RES, :RSR_DELIVERY_RES, :TEMP_POPULATION, :TRANS_POSITIONS, :TRANS_BUSINESS, :VMS_ATS_SOFT, :CREATEDDATE, :CREATEDBYID)"

    '                    parameters.PurgeParametre()
    '                    parameters.AjouterParametreChaine(":ID_RSR_OPP", RSR_Opportunity_Qualification.Id)

    '                    If Not IsNothing(RSR_Opportunity_Qualification.Name) Then
    '                        parameters.AjouterParametreChaine(":RSROPP_NAME", RSR_Opportunity_Qualification.Name)
    '                    Else
    '                        parameters.AjouterParametre(":RSROPP_NAME", String.Empty)
    '                    End If
    '                    If Not IsNothing(RSR_Opportunity_Qualification.Name) Then
    '                        parameters.AjouterParametreChaine(":RSROPP_NAME", RSR_Opportunity_Qualification.Name)
    '                    Else
    '                        parameters.AjouterParametre(":RSROPP_NAME", String.Empty)
    '                    End If
    '                    If Not IsNothing(RSR_Opportunity_Qualification.Account_partner__c) Then
    '                        parameters.AjouterParametreChaine(":IA_CD_PARTNER", RSR_Opportunity_Qualification.Account_partner__c)
    '                    Else
    '                        parameters.AjouterParametre(":IA_CD_PARTNER", String.Empty)
    '                    End If
    '                    If Not IsNothing(RSR_Opportunity_Qualification.An__c) Then
    '                        parameters.AjouterParametreChaine(":ANNUAL_HIRES", RSR_Opportunity_Qualification.An__c)
    '                    Else
    '                        parameters.AjouterParametre(":ANNUAL_HIRES", String.Empty)
    '                    End If
    '                    If Not IsNothing(RSR_Opportunity_Qualification.Annual_Placements__c) AndAlso _SF_RSROPP_ANPLAC.ContainsKey(RSR_Opportunity_Qualification.Annual_Placements__c) Then
    '                        parameters.AjouterParametreChaine(":ANNUAL_PLACEMENTS", _SF_RSROPP_ANPLAC(RSR_Opportunity_Qualification.Annual_Placements__c))
    '                    Else
    '                        parameters.AjouterParametre(":ANNUAL_PLACEMENTS", String.Empty)
    '                    End If

    '                    If Not IsNothing(recordType) Then
    '                        If Not IsNothing(recordType.Name) Then
    '                            parameters.AjouterParametreChaine(":RECORD_TYPE_RSR", _SF_RSROPP_REC_TYP(recordType.Name))
    '                        Else
    '                            parameters.AjouterParametre(":RECORD_TYPE_RSR", String.Empty)
    '                        End If
    '                    Else
    '                        parameters.AjouterParametre(":RECORD_TYPE_RSR", String.Empty)
    '                    End If

    '                    If Not IsNothing(RSR_Opportunity_Qualification.Contact__c) Then
    '                        parameters.AjouterParametreChaine(":CONTACT", RSR_Opportunity_Qualification.Contact__c)
    '                    Else
    '                        parameters.AjouterParametre(":CONTACT", String.Empty)
    '                    End If

    '                    If Not IsNothing(RSR_Opportunity_Qualification.Delivery_Date__c) Then
    '                        parameters.AjouterParametreChaine(":DELIVERY_DATE", RSR_Opportunity_Qualification.Delivery_Date__c)
    '                    Else
    '                        parameters.AjouterParametre(":DELIVERY_DATE", String.Empty)
    '                    End If

    '                    If Not IsNothing(RSR_Opportunity_Qualification.expected_go_live_date_project__c) Then
    '                        parameters.AjouterParametreChaine(":GOLIVE_PROJECT_DATE", RSR_Opportunity_Qualification.expected_go_live_date_project__c)
    '                    Else
    '                        parameters.AjouterParametre(":GOLIVE_PROJECT_DATE", String.Empty)
    '                    End If

    '                    If Not IsNothing(RSR_Opportunity_Qualification.Expected_start_date_project__c) Then
    '                        parameters.AjouterParametreChaine(":START_PROJECT_DATE", RSR_Opportunity_Qualification.Expected_start_date_project__c)
    '                    Else
    '                        parameters.AjouterParametre(":START_PROJECT_DATE", String.Empty)
    '                    End If
    '                    If Not IsNothing(RSR_Opportunity_Qualification.Fee_Randstad_opco_s__c) Then
    '                        parameters.AjouterParametreChaine(":FEE_RAND_OPCO", RSR_Opportunity_Qualification.Fee_Randstad_opco_s__c)
    '                    Else
    '                        parameters.AjouterParametre(":FEE_RAND_OPCO", String.Empty)
    '                    End If
    '                    If Not IsNothing(RSR_Opportunity_Qualification.Fee_suppliers_3rd_party__c) Then
    '                        parameters.AjouterParametreChaine(":FEE_SUPPLIERS", RSR_Opportunity_Qualification.Fee_suppliers_3rd_party__c)
    '                    Else
    '                        parameters.AjouterParametre(":FEE_SUPPLIERS", String.Empty)
    '                    End If
    '                    If Not IsNothing(RSR_Opportunity_Qualification.GEO_Scope_OPP__c) AndAlso _SF_RSROPP_GESCOPP.ContainsKey(RSR_Opportunity_Qualification.GEO_Scope_OPP__c) Then
    '                        parameters.AjouterParametreChaine(":GEO_SCOPE_OPP", _SF_RSROPP_GESCOPP(RSR_Opportunity_Qualification.GEO_Scope_OPP__c))
    '                    Else
    '                        parameters.AjouterParametre(":GEO_SCOPE_OPP", String.Empty)
    '                    End If
    '                    If Not IsNothing(RSR_Opportunity_Qualification.Leading_RSR_region__c) AndAlso _SF_RSROPP_LEAD_RG.ContainsKey(RSR_Opportunity_Qualification.Leading_RSR_region__c) Then
    '                        parameters.AjouterParametreChaine(":LEADING_RSR_REGION", _SF_RSROPP_LEAD_RG(RSR_Opportunity_Qualification.Leading_RSR_region__c))
    '                    Else
    '                        parameters.AjouterParametre(":LEADING_RSR_REGION", String.Empty)
    '                    End If
    '                    If Not IsNothing(RSR_Opportunity_Qualification.Managed_Spend__c) Then
    '                        parameters.AjouterParametreChaine(":MANAGED_SPEND", RSR_Opportunity_Qualification.Managed_Spend__c)
    '                    Else
    '                        parameters.AjouterParametre(":MANAGED_SPEND", String.Empty)
    '                    End If
    '                    If Not IsNothing(RSR_Opportunity_Qualification.MSP_Company__c) AndAlso _SF_RSROPP_MSP_CMY.ContainsKey(RSR_Opportunity_Qualification.MSP_Company__c) Then
    '                        parameters.AjouterParametreChaine(":MSP_COMPANY", _SF_RSROPP_MSP_CMY(RSR_Opportunity_Qualification.MSP_Company__c))
    '                    Else
    '                        parameters.AjouterParametre(":MSP_COMPANY", String.Empty)
    '                    End If
    '                    If Not IsNothing(RSR_Opportunity_Qualification.MSP_Engagement_type__c) AndAlso _SF_RSROPP_MSP_ENG.ContainsKey(RSR_Opportunity_Qualification.MSP_Engagement_type__c) Then
    '                        parameters.AjouterParametreChaine(":MSP_ENGAGEMENT_TYPE", _SF_RSROPP_MSP_ENG(RSR_Opportunity_Qualification.MSP_Engagement_type__c))
    '                    Else
    '                        parameters.AjouterParametre(":MSP_ENGAGEMENT_TYPE", String.Empty)
    '                    End If
    '                    If Not IsNothing(RSR_Opportunity_Qualification.Opportunity_GCS__c) Then
    '                        parameters.AjouterParametreChaine(":OPPORTUNITY_GCS", RSR_Opportunity_Qualification.Opportunity_GCS__c)
    '                    Else
    '                        parameters.AjouterParametre(":OPPORTUNITY_GCS", String.Empty)
    '                    End If
    '                    If Not IsNothing(RSR_Opportunity_Qualification.Priority__c) AndAlso _SF_RSROPP_PRIORIT.ContainsKey(RSR_Opportunity_Qualification.Priority__c) Then
    '                        parameters.AjouterParametreChaine(":PRIORITY", _SF_RSROPP_PRIORIT(RSR_Opportunity_Qualification.Priority__c))
    '                    Else
    '                        parameters.AjouterParametre(":PRIORITY", String.Empty)
    '                    End If
    '                    If Not IsNothing(RSR_Opportunity_Qualification.Qualified__c) AndAlso _SF_RSROPP_QUALIF.ContainsKey(RSR_Opportunity_Qualification.Qualified__c) Then
    '                        parameters.AjouterParametreChaine(":QUALIFIED", _SF_RSROPP_QUALIF(RSR_Opportunity_Qualification.Qualified__c))
    '                    Else
    '                        parameters.AjouterParametre(":QUALIFIED", String.Empty)
    '                    End If
    '                    If Not IsNothing(RSR_Opportunity_Qualification.RPO_Eng__c) AndAlso _SF_RSROPP_RPO.ContainsKey(RSR_Opportunity_Qualification.RPO_Eng__c) Then
    '                        parameters.AjouterParametreChaine(":RPO_ENGAG_TYPE", _SF_RSROPP_RPO(RSR_Opportunity_Qualification.RPO_Eng__c))
    '                    Else
    '                        parameters.AjouterParametre(":RPO_ENGAG_TYPE", String.Empty)
    '                    End If
    '                    If Not IsNothing(RSR_Opportunity_Qualification.RSR_Implementation_Responsible__c) Then
    '                        parameters.AjouterParametreChaine(":CONTACT_RSRIMP_RES", RSR_Opportunity_Qualification.RSR_Implementation_Responsible__c)
    '                    Else
    '                        parameters.AjouterParametre(":CONTACT_RSRIMP_RES", String.Empty)
    '                    End If
    '                    If Not IsNothing(RSR_Opportunity_Qualification.Sourceright_delivery_responsible__c) Then
    '                        parameters.AjouterParametreChaine(":RSR_DELIVERY_RES", RSR_Opportunity_Qualification.Sourceright_delivery_responsible__c)
    '                    Else
    '                        parameters.AjouterParametre(":RSR_DELIVERY_RES", String.Empty)
    '                    End If
    '                    If Not IsNothing(RSR_Opportunity_Qualification.Temp_Population__c) Then
    '                        parameters.AjouterParametreChaine(":TEMP_POPULATION", RSR_Opportunity_Qualification.Temp_Population__c)
    '                    Else
    '                        parameters.AjouterParametre(":TEMP_POPULATION", String.Empty)
    '                    End If
    '                    If Not IsNothing(RSR_Opportunity_Qualification.Transition_positions__c) Then
    '                        parameters.AjouterParametreChaine(":TRANS_POSITIONS", RSR_Opportunity_Qualification.Transition_positions__c)
    '                    Else
    '                        parameters.AjouterParametre(":TRANS_POSITIONS", String.Empty)
    '                    End If
    '                    If Not IsNothing(RSR_Opportunity_Qualification.Transition_Business__c) AndAlso _SF_RSROPP_TRANBUS.ContainsKey(RSR_Opportunity_Qualification.Transition_Business__c) Then
    '                        parameters.AjouterParametreChaine(":TRANS_BUSINESS", _SF_RSROPP_TRANBUS(RSR_Opportunity_Qualification.Transition_Business__c))
    '                    Else
    '                        parameters.AjouterParametre(":TRANS_BUSINESS", String.Empty)
    '                    End If
    '                    If Not IsNothing(RSR_Opportunity_Qualification.VMS_Software__c) AndAlso _SF_RSROPP_VMS_ATS.ContainsKey(RSR_Opportunity_Qualification.VMS_Software__c) Then
    '                        parameters.AjouterParametreChaine(":VMS_ATS_SOFT", _SF_RSROPP_VMS_ATS(RSR_Opportunity_Qualification.VMS_Software__c))
    '                    Else
    '                        parameters.AjouterParametre(":VMS_ATS_SOFT", String.Empty)
    '                    End If
    '                    If Not IsNothing(RSR_Opportunity_Qualification.CreatedDate) Then
    '                        parameters.AjouterParametreChaine(":CREATEDDATE", RSR_Opportunity_Qualification.CreatedDate)
    '                    Else
    '                        parameters.AjouterParametre(":CREATEDDATE", String.Empty)
    '                    End If

    '                    If Not IsNothing(RSR_Opportunity_Qualification.CreatedById) Then
    '                        parameters.AjouterParametreChaine(":CREATEDBYID", RSR_Opportunity_Qualification.CreatedById)
    '                    Else
    '                        parameters.AjouterParametre(":CREATEDBYID", String.Empty)
    '                    End If
    '                    Dim sqlError As String = _oracleConnectionODS.Requete(insert, parameters)
    '                Next
    '                If result.done Then
    '                    done = True
    '                Else

    '                    result = _binding.queryMore(result.queryLocator)
    '                End If
    '            End While
    '        Else
    '            Console.WriteLine("No RSR Opportunity Qualification found in Salesforce")
    '            sb.Append("No RSR Opportunity Qualification found in Salesforce")
    '        End If

    '    Catch ex As Exception
    '        sb.Append(ex.Message)
    '        Console.WriteLine(ex.Message)
    '    End Try

    '    errors = sb.ToString()

    '    Return errors
    'End Function

    Private Function _references() As String
        Dim sql As String
        Dim err As New StringBuilder

        '
        'Load Miamigate parameters (picklist values)
        '
        '
        'RSR Opportunity Qualification Record_type
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_RSROPP_REC_TYP%'"
        Dim dataTable As New DataTable
        Dim sqlError As String = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In DataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_RSROPP_REC_TYP.Add(SFcode, MiamiCode)
        Next 'read

        'RSR Opportunity Qualification Annual Placements
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_RSROPP_ANPLAC%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_RSROPP_ANPLAC.Add(SFcode, MiamiCode)
        Next 'read

        'RSR Opportunity Qualification GEO Scope OPP
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_RSROPP_GESCOPP%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_RSROPP_GESCOPP.Add(SFcode, MiamiCode)
        Next 'read

        'RSR Opportunity Qualification Leading RSR region
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_RSROPP_LEAD_RG%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_RSROPP_LEAD_RG.Add(SFcode, MiamiCode)
        Next 'read

        'RSR Opportunity Qualification MSP Company
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_RSROPP_MSP_CMY%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_RSROPP_MSP_CMY.Add(SFcode, MiamiCode)
        Next 'read

        'RSR Opportunity Qualification MSP Engagement type
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_RSROPP_MSP_ENG%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_RSROPP_MSP_ENG.Add(SFcode, MiamiCode)
        Next 'read

        'RSR Opportunity Qualification Priority
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_RSROPP_PRIORIT%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_RSROPP_PRIORIT.Add(SFcode, MiamiCode)
        Next 'read

        'RSR Opportunity Qualification Qualified
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_RSROPP_QUALIF%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_RSROPP_QUALIF.Add(SFcode, MiamiCode)
        Next 'read

        'RSR Opportunity Qualification RPO Engagement type
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_RSROPP_RPO%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_RSROPP_RPO.Add(SFcode, MiamiCode)
        Next 'read


        'RSR Opportunity Qualification Transition Business
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_RSROPP_TRANBUS%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_RSROPP_TRANBUS.Add(SFcode, MiamiCode)
        Next 'read

        'RSR Opportunity Qualification VMS/ATS Software
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_RSROPP_VMS_ATS%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_RSROPP_VMS_ATS.Add(SFcode, MiamiCode)
        Next 'read
        '
        'SF Picklist check
        ''

        Dim describeSObjectResult As DescribeSObjectResult = _binding.describeSObject("RSR_Opportunity_Qualification__c")
        Dim fields() As Field = describeSObjectResult.fields
        For Each field As Field In fields

            If field.type.Equals(fieldType.picklist) OrElse field.type.Equals(fieldType.multipicklist) Then
                Debug.WriteLine("*** " + field.name + " ***")

                'RSR Opportunity Qualification Record_type
                If field.name.ToLower.Equals("RecordType".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_RSROPP_REC_TYP.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'RSR Opportunity Qualification Annual Placements
                If field.name.ToLower.Equals("Annual_Placements__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_RSROPP_ANPLAC.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'RSR Opportunity Qualification GEO Scope OPP 
                If field.name.ToLower.Equals("GEO_Scope_OPP__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_RSROPP_GESCOPP.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'RSR Opportunity Qualification Leading RSR region 
                If field.name.ToLower.Equals("Leading_RSR_region__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_RSROPP_LEAD_RG.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'RSR Opportunity Qualification MSP Company 

                If field.name.ToLower.Equals("MSP_Company__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_RSROPP_MSP_CMY.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'RSR Opportunity Qualification MSP Engagement type 
                If field.name.ToLower.Equals("MSP_Engagement_type__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_RSROPP_MSP_ENG.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'RSR Opportunity Qualification Priority 
                If field.name.ToLower.Equals("Priority__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_RSROPP_PRIORIT.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'RSR Opportunity Qualification Qualified 
                If field.name.ToLower.Equals("Qualified__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_RSROPP_QUALIF.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'RSR Opportunity Qualification RPO Engagement type 
                If field.name.ToLower.Equals("RPO_Eng__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_RSROPP_RPO.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'RSR Opportunity Qualification Transition Business 
                If field.name.ToLower.Equals("Transition_Business__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_RSROPP_TRANBUS.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'RSR Opportunity Qualification VMS/ATS Software 
                If field.name.ToLower.Equals("VMS_Software__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_RSROPP_VMS_ATS.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
            End If 'pick list

        Next 'field

        Return err.ToString
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
        'tableList.Add("TMPSF_OPPORTUNITY_HIST")
        tableList.Add("TMPSF_RSR_OPP_QUALIF")
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
