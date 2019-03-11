Imports miami2salesforce.salesforce
Imports Methodes.mthconnexion
Imports System.Text


Public Class loadService_Pricing_framework
    Private _binding As SforceService
    Private _oracleConnectionODS As MthConnexion
    Private _oracleConnectionGATE As MthConnexion

    Private _SF_SOF_INVOICE_PR As New Dictionary(Of String, String)
    Private _SF_SOF_PAY_TERMS As New Dictionary(Of String, String)
    Private _SF_SOF_RES_STATUS As New Dictionary(Of String, String)



    Sub New(ByVal binding As SforceService, ByVal miamigate As String, ByVal miamiods As String)
        _binding = binding
        '_miamiods = New DataBase()
        '_miamiods.ConnectionString = My.Settings.miamiods
        '_miamigate = New DataBase()
        '_miamigate.ConnectionString = My.Settings.miamigate
        _oracleConnectionGATE = New MthConnexion(miamigate)
        _oracleConnectionODS = New MthConnexion(miamiods)
    End Sub


    Function loadService_Pricing_framework(ByVal numberOfDays As Integer) As String
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
            'Service Opco framework
            '********************
            Dim lastModifiedDate As Date
            If numberOfDays > 0 Then
                lastModifiedDate = Date.Now.AddDays(-numberOfDays)
            End If

            'truncate tables
            err = _truncateODSTables()
            result += err

            err = _queryService_Pricing_framework(lastModifiedDate)
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

    Private Function _queryService_Pricing_framework(ByVal lastModifiedDate As Date) As String
        Dim errors As String
        Dim sb As StringBuilder = New StringBuilder()

        Try
            Dim done As Boolean = False
            Dim query As String = "SELECT Id, Name, LastModifiedDate," + _
            "Opportunity__c, Service_Type__c, Amount_opco_opportunity__c, invoice_period__c, Payment_Days__c, date_received__c," + _
            "OPCO_account_name__c, Margin_amount__c, Response_status__c, Campaign__c," + _
            "Duration_Months__c, Steady_State_opco__c, Total_Client_Spend_opco__c, Contract_Value_opco__c," + _
            "Contact_Opco_1__c, Payment_Date__c, Payment_Terms__c, Contact_Opco_2__c, Contact_Opco_3__c," + _
            "Steady_state_New_Business_opco__c, Steady_state_Existing_business_opco__c, CreatedDate, CreatedById FROM Service_Pricing_framework__c"

            'Select Id, Name, ProductCode, Description FROM Product2 
            If lastModifiedDate > Date.MinValue Then
                Dim sLastModifiedDate As String = lastModifiedDate.ToString("yyyy-MM-ddThh:mm:ssZ")
                Dim where As String = String.Format(" WHERE LastModifiedDate >= {0}", sLastModifiedDate)
                query = query + where
            End If
            Dim result As QueryResult = _binding.query(query)
            If result.size > 0 Then
                Console.WriteLine(String.Format("# Service_Pricing_framework: {0}", result.size))
                Dim parameters As TableauParametres = New TableauParametres()

                While Not done
                    Dim objects() As sObject = result.records
                    Dim count = objects.Length
                    For i As Integer = 0 To count - 1
                        Dim Service_Pricing_framework As Service_Pricing_framework__c = objects(i)

                        Debug.WriteLine(String.Format("{0} - {1}", Service_Pricing_framework.Name, Service_Pricing_framework.LastModifiedDate))
                        Console.WriteLine(String.Format("{0} - {1}", Service_Pricing_framework.Name, Service_Pricing_framework.LastModifiedDate))

                        If i >= 107 Then
                            i = i
                        End If

                        Dim insert As String = "Insert Into TMPSF_SRV_OPCO_FRAM" + _
                        "(ID_SRV_OPCO, SPF_NAME, OPP_NAME, SERVICE_TYPE, " + _
                        "OPP_AMOUNT_OPCO, INVOICE_PERIOD, PAYMENT_DAYS, DATE_RECIEVED, OPCO_IA_CD, MARGIN_AMOUNT, " + _
                        "RESPONSE_STATUS, CAMPAIGN, DURATION, STDY_STATE_OPCO, TOTSPEND_CLIENT_OPCO, " + _
                        "CONTRACT_VALUE, CONTACT_OPCO_1, PAYMENT_DATE, PAYMENT_TERMS, CONTACT_OPCO_2, " + _
                        "CONTACT_OPCO_3, STDYSTATE_NEBUSS_OPCO, STDYSTATE_EXBUSS_OPCO, CREATEDDATE, CREATEDBYID)" + _
                        " Values(:ID_SRV_OPCO, :SPF_NAME, :OPP_NAME, :SERVICE_TYPE, " + _
                        ":OPP_AMOUNT_OPCO, :INVOICE_PERIOD, :PAYMENT_DAYS, :DATE_RECIEVED, :OPCO_IA_CD, :MARGIN_AMOUNT, " + _
                        ":RESPONSE_STATUS, :CAMPAIGN, :DURATION, :STDY_STATE_OPCO, :TOTSPEND_CLIENT_OPCO, " + _
                        ":CONTRACT_VALUE, :CONTACT_OPCO_1, :PAYMENT_DATE, :PAYMENT_TERMS, :CONTACT_OPCO_2, " + _
                        ":CONTACT_OPCO_3, :STDYSTATE_NEBUSS_OPCO, :STDYSTATE_EXBUSS_OPCO, :CREATEDDATE, :CREATEDBYID)"

                        parameters.PurgeParametre()
                        parameters.AjouterParametreChaine(":ID_SRV_OPCO", Service_Pricing_framework.Id)

                        If Not IsNothing(Service_Pricing_framework.Name) Then
                            parameters.AjouterParametreChaine(":SPF_NAME", Service_Pricing_framework.Name)
                        Else
                            parameters.AjouterParametre(":SPF_NAME", String.Empty)
                        End If
                        If Not IsNothing(Service_Pricing_framework.Opportunity__c) Then
                            parameters.AjouterParametreChaine(":OPP_NAME", Service_Pricing_framework.Opportunity__c)
                        Else
                            parameters.AjouterParametre(":OPP_NAME", String.Empty)
                        End If
                        If Not IsNothing(Service_Pricing_framework.Service_Type__c) Then
                            parameters.AjouterParametreChaine(":SERVICE_TYPE", Service_Pricing_framework.Service_Type__c)
                        Else
                            parameters.AjouterParametre(":SERVICE_TYPE", String.Empty)
                        End If

                        If Not IsNothing(Service_Pricing_framework.Amount_opco_opportunity__c) Then
                            parameters.AjouterParametreChaine(":OPP_AMOUNT_OPCO", Service_Pricing_framework.Amount_opco_opportunity__c)
                        Else
                            parameters.AjouterParametre(":OPP_AMOUNT_OPCO", String.Empty)
                        End If

                        'If Not IsNothing(Service_Pricing_framework.invoice_period__c) AndAlso _SF_SOF_INVOICE_PR.ContainsKey(Service_Pricing_framework.invoice_period__c) Then
                        '    parameters.AjouterParametreChaine(":INVOICE_PERIOD", _SF_SOF_INVOICE_PR(Service_Pricing_framework.invoice_period__c))
                        'Else
                        '    parameters.AjouterParametre(":INVOICE_PERIOD", String.Empty)
                        'End If

                        If Not IsNothing(Service_Pricing_framework.Payment_Days__c) Then
                            parameters.AjouterParametreChaine(":PAYMENT_DAYS", Service_Pricing_framework.Payment_Days__c)
                        Else
                            parameters.AjouterParametre(":PAYMENT_DAYS", String.Empty)
                        End If

                        If Not IsNothing(Service_Pricing_framework.date_received__c) Then
                            parameters.AjouterParametreChaine(":DATE_RECIEVED", Service_Pricing_framework.date_received__c)
                        Else
                            parameters.AjouterParametre(":DATE_RECIEVED", String.Empty)
                        End If

                        If Not IsNothing(Service_Pricing_framework.OPCO_account_name__c) Then
                            parameters.AjouterParametreChaine(":OPCO_IA_CD", Service_Pricing_framework.OPCO_account_name__c)
                        Else
                            parameters.AjouterParametre(":OPCO_IA_CD", String.Empty)
                        End If

                        If Not IsNothing(Service_Pricing_framework.Margin_amount__c) Then
                            parameters.AjouterParametreChaine(":MARGIN_AMOUNT", Service_Pricing_framework.Margin_amount__c)
                        Else
                            parameters.AjouterParametre(":MARGIN_AMOUNT", String.Empty)
                        End If

                        If Not IsNothing(Service_Pricing_framework.Response_status__c) AndAlso _SF_SOF_RES_STATUS.ContainsKey(Service_Pricing_framework.Response_status__c) Then
                            parameters.AjouterParametreChaine(":RESPONSE_STATUS", _SF_SOF_RES_STATUS(Service_Pricing_framework.Response_status__c))
                        Else
                            parameters.AjouterParametre(":RESPONSE_STATUS", String.Empty)
                        End If


                        If Not IsNothing(Service_Pricing_framework.Campaign__c) Then
                            parameters.AjouterParametreChaine(":CAMPAIGN", Service_Pricing_framework.Campaign__c)
                        Else
                            parameters.AjouterParametre(":CAMPAIGN", String.Empty)
                        End If


                        If Not IsNothing(Service_Pricing_framework.Duration_Months__c) Then
                            parameters.AjouterParametreChaine(":DURATION", Service_Pricing_framework.Duration_Months__c)
                        Else
                            parameters.AjouterParametre(":DURATION", String.Empty)
                        End If

                        If Not IsNothing(Service_Pricing_framework.Steady_State_opco__c) Then
                            parameters.AjouterParametreChaine(":STDY_STATE_OPCO", Service_Pricing_framework.Steady_State_opco__c)
                        Else
                            parameters.AjouterParametre(":STDY_STATE_OPCO", String.Empty)
                        End If

                        If Not IsNothing(Service_Pricing_framework.Total_Client_Spend_opco__c) Then
                            parameters.AjouterParametreChaine(":TOTSPEND_CLIENT_OPCO", Service_Pricing_framework.Total_Client_Spend_opco__c)
                        Else
                            parameters.AjouterParametre(":TOTSPEND_CLIENT_OPCO", String.Empty)
                        End If

                        If Not IsNothing(Service_Pricing_framework.Contract_Value_opco__c) Then
                            parameters.AjouterParametreChaine(":CONTRACT_VALUE", Service_Pricing_framework.Contract_Value_opco__c)
                        Else
                            parameters.AjouterParametre(":CONTRACT_VALUE", String.Empty)
                        End If

                        If Not IsNothing(Service_Pricing_framework.Contact_Opco_1__c) Then
                            parameters.AjouterParametreChaine(":CONTACT_OPCO_1", Service_Pricing_framework.Contact_Opco_1__c)
                        Else
                            parameters.AjouterParametre(":CONTACT_OPCO_1", String.Empty)
                        End If

                        If Not IsNothing(Service_Pricing_framework.Payment_Date__c) Then
                            parameters.AjouterParametreChaine(":PAYMENT_DATE", Service_Pricing_framework.Payment_Date__c)
                        Else
                            parameters.AjouterParametre(":PAYMENT_DATE", String.Empty)
                        End If

                        If Not IsNothing(Service_Pricing_framework.Payment_Terms__c) Then
                            parameters.AjouterParametreChaine(":PAYMENT_TERMS", _SF_SOF_PAY_TERMS(Service_Pricing_framework.Payment_Terms__c))
                        Else
                            parameters.AjouterParametre(":PAYMENT_TERMS", String.Empty)
                        End If

                        If Not IsNothing(Service_Pricing_framework.Contact_Opco_2__c) Then
                            parameters.AjouterParametreChaine(":CONTACT_OPCO_2", Service_Pricing_framework.Contact_Opco_2__c)
                        Else
                            parameters.AjouterParametre(":CONTACT_OPCO_2", String.Empty)
                        End If

                        If Not IsNothing(Service_Pricing_framework.Contact_Opco_3__c) Then
                            parameters.AjouterParametreChaine(":CONTACT_OPCO_3", Service_Pricing_framework.Contact_Opco_3__c)
                        Else
                            parameters.AjouterParametre(":CONTACT_OPCO_3", String.Empty)
                        End If

                        If Not IsNothing(Service_Pricing_framework.Steady_state_New_Business_opco__c) Then
                            parameters.AjouterParametreChaine(":STDYSTATE_NEBUSS_OPCO", Service_Pricing_framework.Steady_state_New_Business_opco__c)
                        Else
                            parameters.AjouterParametre(":STDYSTATE_NEBUSS_OPCO", String.Empty)
                        End If

                        If Not IsNothing(Service_Pricing_framework.Steady_state_Existing_business_opco__c) Then
                            parameters.AjouterParametreChaine(":STDYSTATE_EXBUSS_OPCO", Service_Pricing_framework.Steady_state_Existing_business_opco__c)
                        Else
                            parameters.AjouterParametre(":STDYSTATE_EXBUSS_OPCO", String.Empty)
                        End If

                        If Not IsNothing(Service_Pricing_framework.CreatedDate) Then
                            parameters.AjouterParametreChaine(":CREATEDDATE", Service_Pricing_framework.CreatedDate)
                        Else
                            parameters.AjouterParametre(":CREATEDDATE", String.Empty)
                        End If

                        If Not IsNothing(Service_Pricing_framework.CreatedById) Then
                            parameters.AjouterParametreChaine(":CREATEDBYID", Service_Pricing_framework.CreatedById)
                        Else
                            parameters.AjouterParametre(":CREATEDBYID", String.Empty)
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
                Console.WriteLine("No Service-Opco framework found in Salesforce")
                sb.Append("No Service-Opco framework found in Salesforce")
            End If

        Catch ex As Exception
            sb.Append(ex.Message)
            Console.WriteLine(ex.Message)
        End Try

        errors = sb.ToString()

        Return errors
    End Function

    Private Function _references() As String
        Dim sql As String
        Dim err As New StringBuilder

        '
        'Load Miamigate parameters (picklist values)
        '
        '
        'Service-Opco framework invoice period
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_SOF_INVOICE_PR%'"
        Dim dataTable As New DataTable
        Dim sqlError As String = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_SOF_INVOICE_PR.Add(SFcode, MiamiCode)
        Next 'read


        'Service-Opco framework Payment Terms
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_SOF_PAY_TERMS%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_SOF_PAY_TERMS.Add(SFcode, MiamiCode)
        Next 'read


        'Service-Opco framework Response status
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_SOF_RES_STATUS%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_SOF_RES_STATUS.Add(SFcode, MiamiCode)
        Next 'read


        '
        'SF Picklist check
        '
        Dim describeSObjectResult As DescribeSObjectResult = _binding.describeSObject("Service_Pricing_framework__c")
        Dim fields() As Field = describeSObjectResult.fields()
        For Each field As Field In fields

            If field.type.Equals(fieldType.picklist) OrElse field.type.Equals(fieldType.multipicklist) Then
                Debug.WriteLine("*** " + field.name + " ***")

                'Service-Opco framework invoice period
                If field.name.ToLower.Equals("invoice_period__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_SOF_INVOICE_PR.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Service-Opco framework Payment Terms
                If field.name.ToLower.Equals("Payment_Terms__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_SOF_PAY_TERMS.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Service-Opco framework Response status
                If field.name.ToLower.Equals("Response_status__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_SOF_RES_STATUS.ContainsKey(pickListEntry.value) Then
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
        'tableList.Add("TMPSF_RSR_OPP_QUALIF")
        tableList.Add("TMPSF_SRV_OPCO_FRAM")

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


