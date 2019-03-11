Imports miami2salesforce.salesforce
Imports Methodes.mthconnexion
Imports System.Text

Public Class loadLeads
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

    Private _SF_LEAD_REC_TYPE As New Dictionary(Of String, String)
    Private _SF_LEAD_INDUSTRY As New Dictionary(Of String, String)
    Private _SF_LEAD_SOURCE As New Dictionary(Of String, String)
    Private _SF_LEAD_STATUS As New Dictionary(Of String, String)
    Private _SF_LEAD_RATING As New Dictionary(Of String, String)
    Private _SF_LEAD_CTRY_SCP As New Dictionary(Of String, String)
    Private _SF_LEAD_GCS_CTC As New Dictionary(Of String, String)
    Private _SF_LEAD_GCSVERTI As New Dictionary(Of String, String)
    Private _SF_LEAD_RSR_CONT As New Dictionary(Of String, String)
    Private _SF_LEAD_RSR_QUALI As New Dictionary(Of String, String)
    Private _SF_LEAD_SRV_TYPE As New Dictionary(Of String, String)
    Private _SF_LEAD_REQUEST As New Dictionary(Of String, String)
    Sub New(ByVal binding As SforceService, ByVal miamigate As String, ByVal miamiods As String)
        _binding = binding
        '_miamiods = New DataBase()
        '_miamiods.ConnectionString = My.Settings.miamiods
        '_miamigate = New DataBase()
        '_miamigate.ConnectionString = My.Settings.miamigate
        _oracleConnectionGATE = New MthConnexion(miamigate)
        _oracleConnectionODS = New MthConnexion(miamiods)
    End Sub

    Function loadLeads(ByVal numberOfDays As Integer) As String
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
            'Lead
            '********************
            Dim lastModifiedDate As Date
            If numberOfDays > 0 Then
                lastModifiedDate = Date.Now.AddDays(-numberOfDays)
            End If

            'truncate tables
            err = _truncateODSTables()
            result += err

            err = _queryLeads(lastModifiedDate)
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

    Private Function _queryLeads(ByVal lastModifiedDate As Date) As String
        Dim errors As String
        Dim sb As StringBuilder = New StringBuilder()

        Try
            Dim done As Boolean = False
            Dim query As String = "SELECT Id, LastModifiedDate, Email, Industry, RecordType.Name, " + _
            "LeadSource, Status, Name, Company, NumberOfEmployees, AnnualRevenue, Phone, " + _
            "Rating, Title, Website, Converted_Opportunity__c, GEO_mapping_leads__c, " + _
            "expected_tender_date__c, GCS_RSR_Account_lookup__c, GCS_contact__c, GCS_Vertical__c, Lead_Name__c, Next_Steps__c, " + _
            "RSR_Contacts__c, RSR_Qualification__c, Reason_Unqualified__c,  Service_Type__c, Type_of_request__c, CreatedDate, CreatedById FROM Lead"

            If lastModifiedDate > Date.MinValue Then
                Dim sLastModifiedDate As String = lastModifiedDate.ToString("yyyy-MM-ddThh:mm:ssZ")
                Dim where As String = String.Format(" WHERE LastModifiedDate >= {0}", sLastModifiedDate)
                query = query + where
            End If
            Dim result As QueryResult = _binding.query(query)
            If result.size > 0 Then
                Console.WriteLine(String.Format("# Leads: {0}", result.size))
                Dim parameters As TableauParametres = New TableauParametres()

                While Not done
                    Dim objects() As sObject = result.records
                    Dim count = objects.Length
                    For i As Integer = 0 To count - 1
                        Dim Leads As Lead = objects(i)


                        Dim recordType As RecordType = Leads.RecordType
                        Debug.WriteLine(String.Format("{0} - {1}", Leads.Name, Leads.LastModifiedDate))
                        Console.WriteLine(String.Format("{0} - {1}", Leads.Name, Leads.LastModifiedDate))

                        If i >= 107 Then
                            i = i
                        End If

                        Dim insert As String = "Insert Into TMPSF_LEADS" + _
                      "(ID_LEAD, EMAIL, INDUSTRY, RECORD_TYPE_LEAD, SOURCE,  " + _
                      "STATUS, NAME, COMPANY, NUMBER_EMPLY, ANNUAL_REVENUE, PHONE, RATING, TITLE,  " + _
                      "WEBSITE, CONVERTED_OPP, COUNTRIES_SCOPE, EXP_TENDER_DATE,  " + _
                      "GCS_RSR_ACCOUNT, GCS_CONTACT, GCS_VERTICAL, LEAD_NAME, NEXT_STEPS,  " + _
                      "RSR_CONTACT, RSR_QUALIF, REASON_UNQUALIFIED, SERVICE_TYPE, TYPE_REQUEST, CREATEDDATE, CREATEDBYID)" + _
                      " Values(:ID_LEAD, :EMAIL, :INDUSTRY, :RECORD_TYPE_LEAD, :SOURCE, " + _
                      ":STATUS, :NAME, :COMPANY, :NUMBER_EMPLY, :ANNUAL_REVENUE, :PHONE, :RATING, :TITLE,  " + _
                      ":WEBSITE, :CONVERTED_OPP, :COUNTRIES_SCOPE, :EXP_TENDER_DATE,  " + _
                      ":GCS_RSR_ACCOUNT, :GCS_CONTACT, :GCS_VERTICAL, :LEAD_NAME, :NEXT_STEPS,  " + _
                      ":RSR_CONTACT, :RSR_QUALIF, :REASON_UNQUALIFIED,:SERVICE_TYPE, :TYPE_REQUEST, :CREATEDDATE, :CREATEDBYID)"

                        parameters.PurgeParametre()
                        parameters.AjouterParametreChaine(":ID_LEAD", Leads.Id)
                        If Not IsNothing(Leads.Email) Then
                            parameters.AjouterParametreChaine(":EMAIL", Leads.Email)
                        Else
                            parameters.AjouterParametre(":EMAIL", String.Empty)
                        End If
                        If Not IsNothing(Leads.Industry) AndAlso _SF_LEAD_INDUSTRY.ContainsKey(Leads.Industry) Then
                            parameters.AjouterParametreChaine(":INDUSTRY", _SF_LEAD_INDUSTRY(Leads.Industry))
                        Else
                            parameters.AjouterParametre(":INDUSTRY", String.Empty)
                        End If
                        If Not IsNothing(recordType) Then
                            If Not IsNothing(recordType.Name) Then
                                parameters.AjouterParametreChaine(":RECORD_TYPE_LEAD", _SF_LEAD_REC_TYPE(recordType.Name))
                            Else
                                parameters.AjouterParametre(":RECORD_TYPE_LEAD", String.Empty)
                            End If
                        Else
                            parameters.AjouterParametre(":RECORD_TYPE_LEAD", String.Empty)
                        End If
                        If Not IsNothing(Leads.LeadSource) AndAlso _SF_LEAD_SOURCE.ContainsKey(Leads.Status) Then
                            parameters.AjouterParametreChaine(":SOURCE", _SF_LEAD_SOURCE(Leads.LeadSource))
                        Else
                            parameters.AjouterParametre(":SOURCE", String.Empty)
                        End If
                        If Not IsNothing(Leads.Status) AndAlso _SF_LEAD_STATUS.ContainsKey(Leads.Status) Then
                            parameters.AjouterParametreChaine(":STATUS", _SF_LEAD_STATUS(Leads.Status))
                        Else
                            parameters.AjouterParametre(":STATUS", String.Empty)
                        End If
                        If Not IsNothing(Leads.Name) Then
                            parameters.AjouterParametreChaine(":NAME", Leads.Name)
                        Else
                            parameters.AjouterParametre(":NAME", String.Empty)
                        End If

                        If Not IsNothing(Leads.Company) Then
                            parameters.AjouterParametreChaine(":COMPANY", Leads.Company)
                        Else
                            parameters.AjouterParametre(":COMPANY", String.Empty)
                        End If
                        If Not IsNothing(Leads.NumberOfEmployees) Then
                            parameters.AjouterParametreChaine(":NUMBER_EMPLY", Leads.NumberOfEmployees)
                        Else
                            parameters.AjouterParametre(":NUMBER_EMPLY", String.Empty)
                        End If
                        If Not IsNothing(Leads.AnnualRevenue) Then
                            parameters.AjouterParametreChaine(":ANNUAL_REVENUE", Leads.AnnualRevenue)
                        Else
                            parameters.AjouterParametre(":ANNUAL_REVENUE", String.Empty)
                        End If
                        If Not IsNothing(Leads.Phone) Then
                            parameters.AjouterParametreChaine(":PHONE", Leads.Phone)
                        Else
                            parameters.AjouterParametre(":PHONE", String.Empty)
                        End If
                        If Not IsNothing(Leads.Rating) AndAlso _SF_LEAD_RATING.ContainsKey(Leads.Rating) Then
                            parameters.AjouterParametreChaine(":RATING", _SF_LEAD_RATING(Leads.Rating))
                        Else
                            parameters.AjouterParametre(":RATING", String.Empty)
                        End If
                        If Not IsNothing(Leads.Title) Then
                            parameters.AjouterParametreChaine(":TITLE", Leads.Title)
                        Else
                            parameters.AjouterParametre(":TITLE", String.Empty)
                        End If
                        If Not IsNothing(Leads.Website) Then
                            parameters.AjouterParametreChaine(":WEBSITE", Leads.Website)
                        Else
                            parameters.AjouterParametre(":WEBSITE", String.Empty)
                        End If
                        If Not IsNothing(Leads.Converted_Opportunity__c) Then
                            parameters.AjouterParametreChaine(":CONVERTED_OPP", Leads.Converted_Opportunity__c)
                        Else
                            parameters.AjouterParametre(":CONVERTED_OPP", String.Empty)
                        End If

                        Dim err As String = Utils._multiPicklist(Leads.GEO_mapping_leads__c, _SF_LEAD_CTRY_SCP, ":COUNTRIES_SCOPE", parameters)
                        If err <> String.Empty Then
                            sb.Append("<br/>GEO_mapping_leads__c: ")
                            sb.Append(err)
                        End If

                        'If Not IsNothing(Leads.GEO_mapping_leads__c) Then
                        '    parameters.AjouterParametreChaine(":COUNTRIES_SCOPE", _SF_LEAD_CTRY_SCP(Leads.GEO_mapping_leads__c))
                        'Else
                        '    parameters.AjouterParametre(":COUNTRIES_SCOPE", String.Empty)
                        'End If

                        If Not IsNothing(Leads.expected_tender_date__c) Then
                            parameters.AjouterParametreChaine(":EXP_TENDER_DATE", Leads.expected_tender_date__c)
                        Else
                            parameters.AjouterParametre(":EXP_TENDER_DATE", String.Empty)
                        End If
                        If Not IsNothing(Leads.GCS_RSR_Account_lookup__c) Then
                            parameters.AjouterParametreChaine(":GCS_RSR_ACCOUNT", Leads.GCS_RSR_Account_lookup__c)
                        Else
                            parameters.AjouterParametre(":GCS_RSR_ACCOUNT", String.Empty)
                        End If

                        err = Utils._multiPicklist(Leads.GCS_contact__c, _SF_LEAD_GCS_CTC, ":GCS_CONTACT", parameters)
                        If err <> String.Empty Then
                            sb.Append("<br/>GCS_contact__c: ")
                            sb.Append(err)
                        End If

                        'If Not IsNothing(Leads.GCS_contact__c) Then
                        '    parameters.AjouterParametreChaine(":GCS_CONTACT", _SF_LEAD_GCS_CTC(Leads.GCS_contact__c))
                        'Else
                        '    parameters.AjouterParametre(":GCS_CONTACT", String.Empty)
                        'End If

                        If Not IsNothing(Leads.GCS_Vertical__c) AndAlso _SF_LEAD_GCSVERTI.ContainsKey(Leads.GCS_Vertical__c) Then
                            parameters.AjouterParametreChaine(":GCS_VERTICAL", _SF_LEAD_GCSVERTI(Leads.GCS_Vertical__c))
                        Else
                            parameters.AjouterParametre(":GCS_VERTICAL", String.Empty)
                        End If
                        If Not IsNothing(Leads.Lead_Name__c) Then
                            parameters.AjouterParametreChaine(":LEAD_NAME", Leads.Lead_Name__c)
                        Else
                            parameters.AjouterParametre(":LEAD_NAME", String.Empty)
                        End If
                        If Not IsNothing(Leads.Next_Steps__c) Then
                            parameters.AjouterParametreChaine(":NEXT_STEPS", Leads.Next_Steps__c)
                        Else
                            parameters.AjouterParametre(":NEXT_STEPS", String.Empty)
                        End If
                        'If Not IsNothing(Leads.partner_information__c) Then
                        '    parameters.AjouterParametreChaine(":INFO_PARTNER", Leads.partner_information__c)
                        'Else
                        '    parameters.AjouterParametre(":INFO_PARTNER", String.Empty)
                        'End If

                        err = Utils._multiPicklist(Leads.RSR_Contacts__c, _SF_LEAD_RSR_CONT, ":RSR_CONTACT", parameters)
                        If err <> String.Empty Then
                            sb.Append("<br/>RSR_Contacts__c: ")
                            sb.Append(err)
                        End If
                        'If Not IsNothing(Leads.RSR_Contacts__c) Then
                        '    parameters.AjouterParametreChaine(":RSR_CONTACT", _SF_LEAD_RSR_CONT(Leads.RSR_Contacts__c))
                        'Else
                        '    parameters.AjouterParametre(":RSR_CONTACT", String.Empty)
                        'End If

                        If Not IsNothing(Leads.RSR_Qualification__c) AndAlso _SF_LEAD_RSR_QUALI.ContainsKey(Leads.RSR_Qualification__c) Then
                            parameters.AjouterParametreChaine(":RSR_QUALIF", _SF_LEAD_RSR_QUALI(Leads.RSR_Qualification__c))
                        Else
                            parameters.AjouterParametre(":RSR_QUALIF", String.Empty)
                        End If

                        If Not IsNothing(Leads.Reason_Unqualified__c) Then
                            parameters.AjouterParametreChaine(":REASON_UNQUALIFIED", Leads.Reason_Unqualified__c)
                        Else
                            parameters.AjouterParametre(":REASON_UNQUALIFIED", String.Empty)
                        End If

                        err = Utils._multiPicklist(Leads.Service_Type__c, _SF_LEAD_SRV_TYPE, ":SERVICE_TYPE", parameters)
                        If err <> String.Empty Then
                            sb.Append("<br/>GCS_contact__c: ")
                            sb.Append(err)
                        End If
                        'If Not IsNothing(Leads.Service_Type__c) Then
                        '    parameters.AjouterParametreChaine(":SERVICE_TYPE", _SF_LEAD_SRV_TYPE(Leads.Service_Type__c))
                        'Else
                        '    parameters.AjouterParametre(":SERVICE_TYPE", String.Empty)
                        'End If

                        If Not IsNothing(Leads.Type_of_request__c) AndAlso _SF_LEAD_REQUEST.ContainsKey(Leads.Type_of_request__c) Then
                            parameters.AjouterParametreChaine(":TYPE_REQUEST", _SF_LEAD_REQUEST(Leads.Type_of_request__c))
                        Else
                            parameters.AjouterParametre(":TYPE_REQUEST", String.Empty)
                        End If

                        If Not IsNothing(Leads.CreatedDate) Then
                            parameters.AjouterParametreChaine(":CREATEDDATE", Leads.CreatedDate)
                        Else
                            parameters.AjouterParametre(":CREATEDDATE", String.Empty)
                        End If

                        If Not IsNothing(Leads.CreatedById) Then
                            parameters.AjouterParametreChaine(":CREATEDBYID", Leads.CreatedById)
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
                Console.WriteLine("No Leads found in Salesforce")
                sb.Append("No Leads found in Salesforce")
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
        'Lead RECORD TYPE
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_LEAD_REC_TYPE%'"
        Dim dataTable As New DataTable
        Dim sqlError As String = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_LEAD_REC_TYPE.Add(SFcode, MiamiCode)
        Next 'read

        'Lead INDUSTRY
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_LEAD_INDUSTRY%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_LEAD_INDUSTRY.Add(SFcode, MiamiCode)
        Next 'read

       

        'Lead STATUS
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_LEAD_STATUS%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_LEAD_STATUS.Add(SFcode, MiamiCode)
        Next 'read


        'Lead LEAD SOURCE
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_LEAD_SOURCE%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_LEAD_SOURCE.Add(SFcode, MiamiCode)
        Next 'read

        'Lead RATING
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_LEAD_RATING%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_LEAD_RATING.Add(SFcode, MiamiCode)
        Next 'read

        'Lead COUNTRIES IN SCOPE
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_LEAD_CTRY_SCP%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_LEAD_CTRY_SCP.Add(SFcode, MiamiCode)
        Next 'read

        'Lead GCS CONTACT
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_LEAD_GCS_CTC%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_LEAD_GCS_CTC.Add(SFcode, MiamiCode)
        Next 'read


        'Lead GCS VERTICAL
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_LEAD_GCSVERTI%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_LEAD_GCSVERTI.Add(SFcode, MiamiCode)
        Next 'read

        'Lead RSR CONTACTS
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_LEAD_RSR_CONT%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_LEAD_RSR_CONT.Add(SFcode, MiamiCode)
        Next 'read

        'Lead RSR QUALIFICATION
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_LEAD_RSR_QUALI%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_LEAD_RSR_QUALI.Add(SFcode, MiamiCode)
        Next 'read

        'Lead SERVICE TYPE
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_LEAD_SRV_TYPE%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_LEAD_SRV_TYPE.Add(SFcode, MiamiCode)
        Next 'read

        'Lead TYPE OF REQUEST
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_LEAD_REQUEST%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_LEAD_REQUEST.Add(SFcode, MiamiCode)
        Next 'read

        '
        'SF Picklist check
        ''

        Dim describeSObjectResult As DescribeSObjectResult = _binding.describeSObject("lead")
        Dim fields() As Field = describeSObjectResult.fields
        For Each field As Field In fields

            If field.type.Equals(fieldType.picklist) OrElse field.type.Equals(fieldType.multipicklist) Then
                Debug.WriteLine("*** " + field.name + " ***")

                'Lead RECORD TYPE
                If field.name.ToLower.Equals("RecordType".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_LEAD_REC_TYPE.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'Lead INDUSTRY 
                If field.name.ToLower.Equals("Industry".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_LEAD_INDUSTRY.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'Lead LEAD SOURCE
                If field.name.ToLower.Equals("LeadSource".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_LEAD_SOURCE.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'Lead LEAD STATUS
                If field.name.ToLower.Equals("Status".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_LEAD_STATUS.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Lead RATING
                If field.name.ToLower.Equals("Rating".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_LEAD_RATING.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'Lead COUNTRIES IN SCOPE
                If field.name.ToLower.Equals("GEO_mapping_leads__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_LEAD_CTRY_SCP.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Lead GCS CONTACT
                If field.name.ToLower.Equals("GCS_contact__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_LEAD_GCS_CTC.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'Lead GCS VERTICAL
                If field.name.ToLower.Equals("GCS_Vertical__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_LEAD_GCSVERTI.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Lead RSR CONTACTS
                If field.name.ToLower.Equals("RSR_Contacts__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_LEAD_RSR_CONT.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'Lead RSR QUALIFICATION
                If field.name.ToLower.Equals("RSR_Qualification__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_LEAD_RSR_QUALI.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'Lead SERVICE TYPE
                If field.name.ToLower.Equals("Service_Type__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_LEAD_SRV_TYPE.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'Lead TYPE OF REQUEST
                If field.name.ToLower.Equals("Type_of_request__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_LEAD_REQUEST.ContainsKey(pickListEntry.value) Then
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
        tableList.Add("TMPSF_LEADS")
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
    'Private Function _multiPicklist(ByVal values As String, ByVal _SF As Dictionary(Of String, String), ByVal param_name As String, ByVal params As TableauParametres) As String
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

    '            params.AjouterParametreChaine(param_name, param_values)
    '        End If
    '    Next
    '    Return err
    'End Function

End Class

