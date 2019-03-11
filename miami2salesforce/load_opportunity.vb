Imports miami2salesforce.salesforce
Imports Methodes.mthconnexion
Imports System.Text

Public Class LoadOpportunity
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

    Private _SF_OPP_CATSKILL As New Dictionary(Of String, String)
    Private _SF_OPP_COMCURR As New Dictionary(Of String, String)
    Private _SF_OPP_CONNDISCAG As New Dictionary(Of String, String)
    Private _SF_OPP_EMPCTR_PRO As New Dictionary(Of String, String)
    Private _SF_OPP_FCAT As New Dictionary(Of String, String)
    Private _SF_OPP_GCSVERTI As New Dictionary(Of String, String)
    Private _SF_OPP_GEOPRES As New Dictionary(Of String, String)
    Private _SF_OPP_LEAD_SOURC As New Dictionary(Of String, String)
    Private _SF_OPP_MISS_CTRY As New Dictionary(Of String, String)
    Private _SF_OPP_PRIOCLASS As New Dictionary(Of String, String)
    Private _SF_OPP_PROFILES As New Dictionary(Of String, String)
    Private _SF_OPP_REC_TYPE As New Dictionary(Of String, String)
    Private _SF_OPP_REG As New Dictionary(Of String, String)
    Private _SF_OPP_SRV_TYPE As New Dictionary(Of String, String)
    Private _SF_OPP_STAGE As New Dictionary(Of String, String)
    Private _SF_OPP_SUPPLI_WON As New Dictionary(Of String, String)
    Private _SF_OPP_TYPE As New Dictionary(Of String, String)

    Sub New(ByVal binding As SforceService, ByVal miamigate As String, ByVal miamiods As String)
        _binding = binding
        '_miamiods = New DataBase()
        '_miamiods.ConnectionString = My.Settings.miamiods
        '_miamigate = New DataBase()
        '_miamigate.ConnectionString = My.Settings.miamigate
        _oracleConnectionGATE = New MthConnexion(miamigate)
        _oracleConnectionODS = New MthConnexion(miamiods)
    End Sub

    Function LoadOpportunity(numberOfDays As Integer) As String
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
            'Opportunity
            '********************
            Dim lastModifiedDate As Date
            If numberOfDays > 0 Then
                lastModifiedDate = Date.Now.AddDays(-numberOfDays)
            End If

            'truncate tables
            err = _truncateODSTables()
            result += err

            '    err = _queryOpportunities(lastModifiedDate)
            '    result += err

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

    'Private Function _queryOpportunities(lastModifiedDate As Date) As String
    '    Dim errors As String
    '    Dim sb As StringBuilder = New StringBuilder()

    '    Try
    '        Dim done As Boolean = False
    '        'retrieve the first 500 opportunities
    '        Dim query As String = "SELECT Id, Name, LastModifiedDate, AccountId, Account.MIAMI_account_ID__c, RecordType.Name," + _
    '        "Fiscal, Amount, CloseDate, ForecastCategoryName, LeadSource, NextStep, Probability, StageName," + _
    '        "Type, Attainable_potential_sales__c, Category_Skills__c, competitor_current_relationship__c," + _
    '        "Confidentiality_or_Non_disclosure_agreem__c, Contract_Value__c, Duration__c, empowerment_central_procurement__c," + _
    '        "estimated_loss_no_presence__c, fill_rate__c, GCS_Vertical__c, GEO_presence_opportunity__c, Margin__c, market_information_opportunity__c," + _
    '        "No_OPCO_for_opportunity__c, Opportunity_region__c, Priority_classification__c, Profiles__c, Reason_Lost_closed_cancelled__c," + _
    '        "right_match_in_capabilities__c, right_match_in_coverage__c, Right_price__c, right_quality_of_proposal__c,     right_quality_of_relationship_with_all_s__c," + _
    '        "right_reference_cases__c, sanity__c, Service_Type_Opportunity__c, Start_date_contract__c, Steady_state_calcualtion__c," + _
    '        "Amount_Existing_business_estimated__c, Amount_new_Business_estimated__c, suppliers_won__c, Total_Client_Spend__c, CreatedDate, CreatedById, Market_share_opportunity__c FROM Opportunity "
    '        'WHERE Account.MIAMI_account_ID__c <> NULL
    '        If lastModifiedDate > Date.MinValue Then
    '            Dim sLastModifiedDate As String = lastModifiedDate.ToString("yyyy-MM-ddThh:mm:ssZ")
    '            Dim where As String = String.Format(" WHERE LastModifiedDate >= {0}", sLastModifiedDate)
    '            query = query + where
    '        End If
    '        Dim result As QueryResult = _binding.query(query)
    '        If result.size > 0 Then
    '            Console.WriteLine(String.Format("# opportunities: {0}", result.size))
    '            Dim parameters As TableauParametres = New TableauParametres()

    '            While Not done
    '                Dim objects() As sObject = result.records
    '                Dim count = objects.Length
    '                For i As Integer = 0 To count - 1
    '                    Dim opportunity As Opportunity = objects(i)

    '                    Dim account As Account = opportunity.Account
    '                    Dim recordType As RecordType = opportunity.RecordType
    '                    Debug.WriteLine(String.Format("{0} - {1}", opportunity.Name, opportunity.LastModifiedDate))
    '                    Console.WriteLine(String.Format("{0} - {1}", opportunity.Name, opportunity.LastModifiedDate))

    '                    If i >= 107 Then
    '                        i = i
    '                    End If

    '                    Dim insert As String = "Insert Into TMPSF_OPPORTUNITY" + _
    '                       "(ID_OPP, ID_ACCOUNT, OPP_NAME, FISCAL, " + _
    '                       " AMOUNT, CLOSE_DATE, FORECAST_CATEGORY, LEAD_SOURCE, NEXT_STEP, " + _
    '                        "RECORD_TYPE_OPP, PROBABILITY, OPP_STAGE, BUSINESS_TYPE, AMOUNT_POTSALES, " + _
    '                        "CATEGORY_SKILLS, COMP_CURRENT_RSHIP, CDA_NDA, CONTRACT_VALUE, DURATION, " + _
    '                        "EMP_CTRL_PROC, ESTIMATED_LOSS, FILL_RATE, GCS_VERTICAL, GEO_PRESENCE, " + _
    '                        "MARGIN, MKT_INFO_OPP, OPP_MISS_CTRY, OPP_REGION, PRTY_CLASS, " + _
    '                        "PROFILES, REASON_LOS_OPP, RGTMATCH_CAPABILITIES, RGTMATCH_COVERAGE, RGT_PRICE, " + _
    '                        "RGTQUALITY_PROPOSAL, RGTQUALITY_RSHIP_STHOL, RGTREFERENCE_CASES, SANITY, OPP_SRV_TYPE, " + _
    '                        "START_DATE_CONTRACT, STDYSTATE, STDYSTATE_EXBUSS, STDYSTATE_NEBUSS, SUPPLIERS_WON, " + _
    '                        "TOT_SPEND_CLIENT, CREATEDDATE, CREATEDBYID, MK_SHARE_OPP)" + _
    '                        " Values(:ID_OPP, :ID_ACCOUNT, :OPP_NAME, :FISCAL, " + _
    '                        ":AMOUNT, :CLOSE_DATE, :FORECAST_CATEGORY, :LEAD_SOURCE, :NEXT_STEP, " + _
    '                        ":RECORD_TYPE_OPP, :PROBABILITY, :OPP_STAGE, :BUSINESS_TYPE, :AMOUNT_POTSALES, " + _
    '                        ":CATEGORY_SKILLS, :COMP_CURRENT_RSHIP, :CDA_NDA, :CONTRACT_VALUE, :DURATION, " + _
    '                        ":EMP_CTRL_PROC, :ESTIMATED_LOSS, :FILL_RATE, :GCS_VERTICAL, :GEO_PRESENCE, " + _
    '                        ":MARGIN, :MKT_INFO_OPP, :OPP_MISS_CTRY, :OPP_REGION, :PRTY_CLASS, " + _
    '                        ":PROFILES, :REASON_LOS_OPP, :RGTMATCH_CAPABILITIES, :RGTMATCH_COVERAGE, :RGT_PRICE, " + _
    '                        ":RGTQUALITY_PROPOSAL, :RGTQUALITY_RSHIP_STHOL, :RGTREFERENCE_CASES, :SANITY, :OPP_SRV_TYPE, " + _
    '                        ":START_DATE_CONTRACT, :STDYSTATE, :STDYSTATE_EXBUSS, :STDYSTATE_NEBUSS, :SUPPLIERS_WON, " + _
    '                        ":TOT_SPEND_CLIENT, :CREATEDDATE, :CREATEDBYID, :MK_SHARE_OPP)"

    '                    parameters.PurgeParametre()
    '                    parameters.AjouterParametreChaine(":ID_OPP", opportunity.Id)

    '                    If Not IsNothing(opportunity.AccountId) Then
    '                        parameters.AjouterParametreChaine(":ID_ACCOUNT", opportunity.AccountId)
    '                    Else
    '                        parameters.AjouterParametre(":ID_ACCOUNT", String.Empty)
    '                    End If


    '                    If Not IsNothing(opportunity.Name) Then
    '                        parameters.AjouterParametreChaine(":OPP_NAME", opportunity.Name)
    '                    Else
    '                        parameters.AjouterParametre(":OPP_NAME", String.Empty)
    '                    End If


    '                    If Not IsNothing(opportunity.Fiscal) Then
    '                        parameters.AjouterParametreChaine(":FISCAL", opportunity.Fiscal)
    '                    Else
    '                        parameters.AjouterParametre(":FISCAL", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.Amount) Then
    '                        parameters.AjouterParametreChaine(":AMOUNT", opportunity.Amount)
    '                    Else
    '                        parameters.AjouterParametre(":AMOUNT", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.CloseDate) Then
    '                        parameters.AjouterParametreDate(":CLOSE_DATE", opportunity.CloseDate)
    '                    Else
    '                        parameters.AjouterParametre(":CLOSE_DATE", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.ForecastCategory) AndAlso _SF_OPP_FCAT.ContainsKey(opportunity.ForecastCategory) Then
    '                        parameters.AjouterParametreChaine(":FORECAST_CATEGORY", _SF_OPP_FCAT(opportunity.ForecastCategory))
    '                    Else
    '                        parameters.AjouterParametre(":FORECAST_CATEGORY", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.LeadSource) AndAlso _SF_OPP_LEAD_SOURC.ContainsKey(opportunity.LeadSource) Then
    '                        parameters.AjouterParametreChaine(":LEAD_SOURCE", _SF_OPP_LEAD_SOURC(opportunity.LeadSource))
    '                    Else
    '                        parameters.AjouterParametre(":LEAD_SOURCE", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.NextStep) Then
    '                        parameters.AjouterParametreChaine(":NEXT_STEP", opportunity.NextStep)
    '                    Else
    '                        parameters.AjouterParametre(":NEXT_STEP", String.Empty)
    '                    End If

    '                    If Not IsNothing(recordType) Then
    '                        If Not IsNothing(recordType.Name) Then
    '                            parameters.AjouterParametreChaine(":RECORD_TYPE_OPP", _SF_OPP_REC_TYPE(recordType.Name))
    '                        Else
    '                            parameters.AjouterParametre(":RECORD_TYPE_OPP", String.Empty)
    '                        End If
    '                    Else
    '                        parameters.AjouterParametre(":RECORD_TYPE_OPP", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.Probability) Then
    '                        parameters.AjouterParametreChaine(":PROBABILITY", opportunity.Probability)
    '                    Else
    '                        parameters.AjouterParametre(":PROBABILITY", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.StageName) AndAlso _SF_OPP_STAGE.ContainsKey(opportunity.StageName) Then
    '                        parameters.AjouterParametreChaine(":OPP_STAGE", _SF_OPP_STAGE(opportunity.StageName))
    '                    Else
    '                        parameters.AjouterParametre(":OPP_STAGE", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.Type) AndAlso _SF_OPP_TYPE.ContainsKey(opportunity.Type) Then
    '                        parameters.AjouterParametreChaine(":BUSINESS_TYPE", _SF_OPP_TYPE(opportunity.Type))
    '                    Else
    '                        parameters.AjouterParametre(":BUSINESS_TYPE", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.Attainable_potential_sales__c) Then
    '                        parameters.AjouterParametreDecimal(":AMOUNT_POTSALES", opportunity.Attainable_potential_sales__c)
    '                    Else
    '                        parameters.AjouterParametre(":AMOUNT_POTSALES", String.Empty)
    '                    End If

    '                    Dim err As String = Utils._multiPicklist(opportunity.Category_Skills__c, _SF_OPP_CATSKILL, ":CATEGORY_SKILLS", parameters)
    '                    If err <> String.Empty Then
    '                        sb.Append("<br/>Category_Skills__c: ")
    '                        sb.Append(err)
    '                    End If

    '                    'If Not IsNothing(opportunity.Category_Skills__c) AndAlso _SF_OPP_CATSKILL.ContainsKey(opportunity.Category_Skills__c) Then
    '                    '    parameters.AjouterParametreChaine(":CATEGORY_SKILLS", _SF_OPP_CATSKILL(opportunity.Category_Skills__c))
    '                    'Else
    '                    '    parameters.AjouterParametre(":CATEGORY_SKILLS", String.Empty)
    '                    'End If

    '                    If Not IsNothing(opportunity.competitor_current_relationship__c) AndAlso _SF_OPP_COMCURR.ContainsKey(opportunity.competitor_current_relationship__c) Then
    '                        parameters.AjouterParametreChaine(":COMP_CURRENT_RSHIP", _SF_OPP_COMCURR(opportunity.competitor_current_relationship__c))
    '                    Else
    '                        parameters.AjouterParametre(":COMP_CURRENT_RSHIP", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.Confidentiality_or_Non_disclosure_agreem__c) AndAlso _SF_OPP_CONNDISCAG.ContainsKey(opportunity.Confidentiality_or_Non_disclosure_agreem__c) Then
    '                        parameters.AjouterParametreChaine(":CDA_NDA", _SF_OPP_CONNDISCAG(opportunity.Confidentiality_or_Non_disclosure_agreem__c))
    '                    Else
    '                        parameters.AjouterParametre(":CDA_NDA", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.Contract_Value__c) Then
    '                        parameters.AjouterParametreChaine(":CONTRACT_VALUE", opportunity.Contract_Value__c)
    '                    Else
    '                        parameters.AjouterParametre(":CONTRACT_VALUE", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.Duration__c) Then
    '                        parameters.AjouterParametreChaine(":DURATION", opportunity.Duration__c)
    '                    Else
    '                        parameters.AjouterParametre(":DURATION", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.empowerment_central_procurement__c) AndAlso _SF_OPP_EMPCTR_PRO.ContainsKey(opportunity.empowerment_central_procurement__c) Then
    '                        parameters.AjouterParametreChaine(":EMP_CTRL_PROC", _SF_OPP_EMPCTR_PRO(opportunity.empowerment_central_procurement__c))
    '                    Else
    '                        parameters.AjouterParametre(":EMP_CTRL_PROC", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.estimated_loss_no_presence__c) Then
    '                        parameters.AjouterParametreChaine(":ESTIMATED_LOSS", opportunity.estimated_loss_no_presence__c)
    '                    Else
    '                        parameters.AjouterParametre(":ESTIMATED_LOSS", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.fill_rate__c) Then
    '                        parameters.AjouterParametreChaine(":FILL_RATE", opportunity.fill_rate__c)
    '                    Else
    '                        parameters.AjouterParametre(":FILL_RATE", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.GCS_Vertical__c) AndAlso _SF_OPP_GCSVERTI.ContainsKey(opportunity.GCS_Vertical__c) Then
    '                        parameters.AjouterParametreChaine(":GCS_VERTICAL", _SF_OPP_GCSVERTI(opportunity.GCS_Vertical__c))
    '                    Else
    '                        parameters.AjouterParametre(":GCS_VERTICAL", String.Empty)
    '                    End If

    '                    err = Utils._multiPicklist(opportunity.GEO_presence_opportunity__c, _SF_OPP_GEOPRES, ":GEO_PRESENCE", parameters)
    '                    If err <> String.Empty Then

    '                        sb.Append("<br/>GEO_presence_opportunity__c: ")
    '                        sb.Append(err)
    '                    End If
    '                    'If Not IsNothing(opportunity.GEO_presence_opportunity__c) AndAlso _SF_OPP_GEOPRES.ContainsKey(opportunity.GEO_presence_opportunity__c) Then
    '                    '    parameters.AjouterParametreChaine(":GEO_PRESENCE", _SF_OPP_GEOPRES(opportunity.GEO_presence_opportunity__c))
    '                    'Else
    '                    '    parameters.AjouterParametre(":GEO_PRESENCE", String.Empty)
    '                    'End If

    '                    If Not IsNothing(opportunity.Margin__c) Then
    '                        parameters.AjouterParametreChaine(":MARGIN", opportunity.Margin__c)
    '                    Else
    '                        parameters.AjouterParametre(":MARGIN", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.market_information_opportunity__c) Then
    '                        parameters.AjouterParametreChaine(":MKT_INFO_OPP", opportunity.market_information_opportunity__c)
    '                    Else
    '                        parameters.AjouterParametre(":MKT_INFO_OPP", String.Empty)
    '                    End If


    '                    err = Utils._multiPicklist(opportunity.No_OPCO_for_opportunity__c, _SF_OPP_MISS_CTRY, ":OPP_MISS_CTRY", parameters)
    '                    If err <> String.Empty Then
    '                        sb.Append("<br/>No_OPCO_for_opportunity__c: ")
    '                        sb.Append(err)
    '                    End If
    '                    'If Not IsNothing(opportunity.No_OPCO_for_opportunity__c) AndAlso _SF_OPP_MISS_CTRY.ContainsKey(opportunity.No_OPCO_for_opportunity__c) Then
    '                    '    parameters.AjouterParametreChaine(":OPP_MISS_CTRY", _SF_OPP_MISS_CTRY(opportunity.No_OPCO_for_opportunity__c))
    '                    'Else
    '                    '    parameters.AjouterParametre(":OPP_MISS_CTRY", String.Empty)
    '                    'End If

    '                    If Not IsNothing(opportunity.Opportunity_region__c) AndAlso _SF_OPP_REG.ContainsKey(opportunity.Opportunity_region__c) Then
    '                        parameters.AjouterParametreChaine(":OPP_REGION", _SF_OPP_REG(opportunity.Opportunity_region__c))
    '                    Else
    '                        parameters.AjouterParametre(":OPP_REGION", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.Priority_classification__c) AndAlso _SF_OPP_PRIOCLASS.ContainsKey(opportunity.Priority_classification__c) Then
    '                        parameters.AjouterParametreChaine(":PRTY_CLASS", _SF_OPP_PRIOCLASS(opportunity.Priority_classification__c))
    '                    Else
    '                        parameters.AjouterParametre(":PRTY_CLASS", String.Empty)
    '                    End If

    '                    err = Utils._multiPicklist(opportunity.Profiles__c, _SF_OPP_PROFILES, ":PROFILES", parameters)
    '                    If err <> String.Empty Then
    '                        sb.Append("<br/>Profiles__c: ")
    '                        sb.Append(err)
    '                    End If
    '                    'If Not IsNothing(opportunity.Profiles__c) AndAlso _SF_OPP_PROFILES.ContainsKey(opportunity.Profiles__c) Then
    '                    '    parameters.AjouterParametreChaine(":PROFILES", _SF_OPP_PROFILES(opportunity.Profiles__c))
    '                    'Else
    '                    '    parameters.AjouterParametre(":PROFILES", String.Empty)
    '                    'End If

    '                    If Not IsNothing(opportunity.Reason_Lost_closed_cancelled__c) Then
    '                        parameters.AjouterParametreChaine(":REASON_LOS_OPP", opportunity.Reason_Lost_closed_cancelled__c)
    '                    Else
    '                        parameters.AjouterParametre(":REASON_LOS_OPP", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.right_match_in_capabilities__c) Then
    '                        parameters.AjouterParametreChaine(":RGTMATCH_CAPABILITIES", opportunity.right_match_in_capabilities__c)
    '                    Else
    '                        parameters.AjouterParametre(":RGTMATCH_CAPABILITIES", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.right_match_in_coverage__c) Then
    '                        parameters.AjouterParametreChaine(":RGTMATCH_COVERAGE", opportunity.right_match_in_coverage__c)
    '                    Else
    '                        parameters.AjouterParametre(":RGTMATCH_COVERAGE", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.Right_price__c) Then
    '                        parameters.AjouterParametreChaine(":RGT_PRICE", opportunity.Right_price__c)
    '                    Else
    '                        parameters.AjouterParametre(":RGT_PRICE", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.right_quality_of_proposal__c) Then
    '                        parameters.AjouterParametreChaine(":RGTQUALITY_PROPOSAL", opportunity.right_quality_of_proposal__c)
    '                    Else
    '                        parameters.AjouterParametre(":RGTQUALITY_PROPOSAL", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.right_quality_of_relationship_with_all_s__c) Then
    '                        parameters.AjouterParametreChaine(":RGTQUALITY_RSHIP_STHOL", opportunity.right_quality_of_relationship_with_all_s__c)
    '                    Else
    '                        parameters.AjouterParametre(":RGTQUALITY_RSHIP_STHOL", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.right_reference_cases__c) Then
    '                        parameters.AjouterParametreChaine(":RGTREFERENCE_CASES", opportunity.right_reference_cases__c)
    '                    Else
    '                        parameters.AjouterParametre(":RGTREFERENCE_CASES", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.sanity__c) Then
    '                        parameters.AjouterParametreChaine(":SANITY", opportunity.sanity__c)
    '                    Else
    '                        parameters.AjouterParametre(":SANITY", String.Empty)
    '                    End If

    '                    err = Utils._multiPicklist(opportunity.Service_Type_Opportunity__c, _SF_OPP_SRV_TYPE, ":OPP_SRV_TYPE", parameters)
    '                    If err <> String.Empty Then
    '                        sb.Append("<br/>Service_Type_Opportunity__c: ")
    '                        sb.Append(err)
    '                    End If
    '                    'If Not IsNothing(opportunity.Service_Type_Opportunity__c) AndAlso _SF_OPP_SRV_TYPE.ContainsKey(opportunity.Service_Type_Opportunity__c) Then
    '                    '    parameters.AjouterParametreChaine(":OPP_SRV_TYPE", _SF_OPP_SRV_TYPE(opportunity.Service_Type_Opportunity__c))
    '                    'Else
    '                    '    parameters.AjouterParametre(":OPP_SRV_TYPE", String.Empty)
    '                    'End If

    '                    If Not IsNothing(opportunity.Start_date_contract__c) Then
    '                        parameters.AjouterParametreDate(":START_DATE_CONTRACT", opportunity.Start_date_contract__c)
    '                    Else
    '                        parameters.AjouterParametre(":START_DATE_CONTRACT", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.Steady_state_calcualtion__c) Then
    '                        parameters.AjouterParametreChaine(":STDYSTATE", opportunity.Steady_state_calcualtion__c)
    '                    Else
    '                        parameters.AjouterParametre(":STDYSTATE", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.Amount_Existing_business_estimated__c) Then
    '                        parameters.AjouterParametreChaine(":STDYSTATE_EXBUSS", opportunity.Amount_Existing_business_estimated__c)
    '                    Else
    '                        parameters.AjouterParametre(":STDYSTATE_EXBUSS", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.Amount_new_Business_estimated__c) Then
    '                        parameters.AjouterParametreChaine(":STDYSTATE_NEBUSS", opportunity.Amount_new_Business_estimated__c)
    '                    Else
    '                        parameters.AjouterParametre(":STDYSTATE_NEBUSS", String.Empty)
    '                    End If

    '                    err = Utils._multiPicklist(opportunity.suppliers_won__c, _SF_OPP_SUPPLI_WON, ":SUPPLIERS_WON", parameters)
    '                    If err <> String.Empty Then
    '                        sb.Append("<br/>suppliers_won__c: ")
    '                        sb.Append(err)
    '                    End If
    '                    'If Not IsNothing(opportunity.suppliers_won__c) AndAlso _SF_OPP_SUPPLI_WON.ContainsKey(opportunity.suppliers_won__c) Then
    '                    '    parameters.AjouterParametreChaine(":SUPPLIERS_WON", _SF_OPP_SUPPLI_WON(opportunity.suppliers_won__c))
    '                    'Else
    '                    '    parameters.AjouterParametre(":SUPPLIERS_WON", String.Empty)
    '                    'End If

    '                    If Not IsNothing(opportunity.Total_Client_Spend__c) Then
    '                        parameters.AjouterParametreChaine(":TOT_SPEND_CLIENT", opportunity.Total_Client_Spend__c)
    '                    Else
    '                        parameters.AjouterParametre(":TOT_SPEND_CLIENT", String.Empty)
    '                    End If


    '                    If Not IsNothing(opportunity.CreatedDate) Then
    '                        parameters.AjouterParametreChaine(":CREATEDDATE", opportunity.CreatedDate)
    '                    Else
    '                        parameters.AjouterParametre(":CREATEDDATE", String.Empty)
    '                    End If

    '                    If Not IsNothing(opportunity.CreatedById) Then
    '                        parameters.AjouterParametreChaine(":CREATEDBYID", opportunity.CreatedById)
    '                    Else
    '                        parameters.AjouterParametre(":CREATEDBYID", String.Empty)
    '                    End If
    '                    'MK_SHARE_OPP

    '                    If Not IsNothing(opportunity.Market_share_opportunity__c) Then
    '                        parameters.AjouterParametreChaine(":MK_SHARE_OPP", opportunity.Market_share_opportunity__c)
    '                    Else
    '                        parameters.AjouterParametre(":MK_SHARE_OPP", String.Empty)
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
    '            Console.WriteLine("No opportunity found in Salesforce")
    '            sb.Append("No opportunity found in Salesforce")
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

        'Opportunity Category skills
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_OPP_CATSKILL%'"
        Dim dataTable As New DataTable
        Dim sqlError As String = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_OPP_CATSKILL.Add(SFcode, MiamiCode)
        Next 'read

        'Opportunity Competitor relationship
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_OPP_COMCURR%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_OPP_COMCURR.Add(SFcode, MiamiCode)
        Next 'read

        'Opportunity Confidentiality or non disclosure agreement
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_OPP_CONNDISCAG%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_OPP_CONNDISCAG.Add(SFcode, MiamiCode)
        Next 'read

        'Opportunity Empowerment central procurement
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_OPP_EMPCTR_PRO%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_OPP_EMPCTR_PRO.Add(SFcode, MiamiCode)
        Next 'read

        'Opportunity Forecast category
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_OPP_FCAT%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_OPP_FCAT.Add(SFcode, MiamiCode)
        Next 'read

        'Opportunity GCS vertical
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_OPP_GCSVERTI%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_OPP_GCSVERTI.Add(SFcode, MiamiCode)
        Next 'read

        'Opportunity Geopresence
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_OPP_GEOPRES%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_OPP_GEOPRES.Add(SFcode, MiamiCode)
        Next 'read

        'Opportunity Lead source
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_OPP_LEAD_SOURC%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_OPP_LEAD_SOURC.Add(SFcode, MiamiCode)
        Next 'read

        'Opportunity Missing country
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_OPP_MISS_CTRY%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_OPP_MISS_CTRY.Add(SFcode, MiamiCode)
        Next 'read

        'Opportunity Priority classification
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_OPP_PRIOCLASS%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_OPP_PRIOCLASS.Add(SFcode, MiamiCode)
        Next 'read

        'Opportunity Profiles
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_OPP_PROFILES%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_OPP_PROFILES.Add(SFcode, MiamiCode)
        Next 'read

        'Opportunity Record type
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_OPP_REC_TYPE%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_OPP_REC_TYPE.Add(SFcode, MiamiCode)
        Next 'read

        'Opportunity Region
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_OPP_REG%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_OPP_REG.Add(SFcode, MiamiCode)
        Next 'read

        'Opportunity Service type
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_OPP_SRV_TYPE%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_OPP_SRV_TYPE.Add(SFcode, MiamiCode)
        Next 'read

        'Opportunity Stage
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_OPP_STAGE%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_OPP_STAGE.Add(SFcode, MiamiCode)
        Next 'read

        'Opportunity Suppliers won
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_OPP_SUPPLI_WON%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_OPP_SUPPLI_WON.Add(SFcode, MiamiCode)
        Next 'read

        'Opportunity Type
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_OPP_TYPE%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_OPP_TYPE.Add(SFcode, MiamiCode)
        Next 'read

        '
        'SF Picklist check
        '
        Dim describeSObjectResult As DescribeSObjectResult = _binding.describeSObject("Opportunity")
        Dim fields() As Field = describeSObjectResult.fields
        For Each field As Field In fields

            If field.type.Equals(fieldType.picklist) OrElse field.type.Equals(fieldType.multipicklist) Then
                Debug.WriteLine("*** " + field.name + " ***")

                'Opportunity Category skills
                If field.name.ToLower.Equals("Category_Skills__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_OPP_CATSKILL.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                        'Dim ok As Boolean = False
                        'Dim pair As KeyValuePair(Of String, String)
                        'For Each pair In _SF_OPP_CATSKILL
                        '    If pair.Value.Equals(pickListEntry.value) Then
                        '        ok = True
                        '        Exit For
                        '    End If
                        'Next
                        'If Not ok Then
                        '    err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        'End If
                    Next
                End If

                'Opportunity Competitor relationship
                If field.name.ToLower.Equals("competitor_current_relationship__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_OPP_COMCURR.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Opportunity Confidentiality or non disclosure agreement
                If field.name.ToLower.Equals("Confidentiality_or_Non_disclosure_agreem__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_OPP_CONNDISCAG.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Opportunity Empowerment central procurement
                If field.name.ToLower.Equals("empowerment_central_procurement__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_OPP_EMPCTR_PRO.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Opportunity Forecast category                
                If field.name.ToLower.Equals("ForecastCategoryName".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_OPP_FCAT.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Opportunity GCS vertical
                If field.name.ToLower.Equals("GCS_Vertical__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_OPP_GCSVERTI.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Opportunity Geopresence
                If field.name.ToLower.Equals("GEO_presence_opportunity__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_OPP_GEOPRES.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Opportunity Lead source
                If field.name.ToLower.Equals("LeadSource".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_OPP_LEAD_SOURC.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Opportunity Missing country
                If field.name.ToLower.Equals("No_OPCO_for_opportunity__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_OPP_MISS_CTRY.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Opportunity Priority classification
                If field.name.ToLower.Equals("Priority_classification__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_OPP_PRIOCLASS.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Opportunity Profiles
                If field.name.ToLower.Equals("Profiles__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_OPP_PROFILES.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Opportunity Record type
                If field.name.ToLower.Equals("RecordType".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_OPP_REC_TYPE.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Opportunity Region
                If field.name.ToLower.Equals("Opportunity_region__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_OPP_REG.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Opportunity Service type
                If field.name.ToLower.Equals("Service_Type_Opportunity__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_OPP_SRV_TYPE.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Opportunity Category skills
                If field.name.ToLower.Equals("Category_Skills__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_OPP_CATSKILL.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Opportunity Stage
                If field.name.ToLower.Equals("StageName".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_OPP_STAGE.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Opportunity Suppliers won
                If field.name.ToLower.Equals("suppliers_won__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_OPP_SUPPLI_WON.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Opportunity Type
                If field.name.ToLower.Equals("Type".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_OPP_TYPE.ContainsKey(pickListEntry.value) Then
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
        tableList.Add("TMPSF_OPPORTUNITY")
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

    'Private Function _multiPicklist(ByVal values As String, ByVal _SF As Dictionary(Of String, String), ByVal param_name As String, ByVal params As TableauParametres) As String
    'Return Utils._multiPicklist(values, _SF, param_name, params)
    '    If IsNothing(values) Then
    '        params.AjouterParametre(param_name, String.Empty)
    '        Return ""
    '    End If

    '    Dim param_values As String = String.Empty
    '    Dim err As String = String.Empty
    '    Dim split As String() = values.Split(";")
    '    For Each value In split
    '        If _SF.ContainsKey(value) Then
    '            If Not param_values.Equals(String.Empty) Then
    '                param_values = String.Concat(param_values, ";")
    '            End If
    '            param_values = String.Concat(param_values, _SF(value))
    '        Else
    '            If Not err.Equals(String.Empty) Then
    '                err = String.Concat(err, ",")
    '            End If
    '            err = String.Concat(err, value)
    '        End If
    '    Next
    '    If Not param_values.Equals(String.Empty) Then
    '        params.AjouterParametreChaine(param_name, param_values)
    '    Else
    '        params.AjouterParametre(param_name, String.Empty)
    '    End If
    '    Return err
    'End Function

End Class
