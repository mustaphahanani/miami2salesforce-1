Imports miami2salesforce.salesforce
Imports Methodes.mthconnexion
Imports System.Text
Public Class loadContact
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

    Private _SF_CTC_LEADSOURCE As New Dictionary(Of String, String)
    Private _SF_CTC_BDMBR_CTC As New Dictionary(Of String, String)
    Private _SF_CTC_PRIORITY As New Dictionary(Of String, String)
    Private _SF_CTC_DEPT_QUAL As New Dictionary(Of String, String)
    Private _SF_CTC_DMU_RSHIP As New Dictionary(Of String, String)
    Private _SF_CTC_LEVEL_IMP As New Dictionary(Of String, String)
    Private _SF_CTC_SPOC_SPLTY As New Dictionary(Of String, String)
    Private _SF_CTC_STATUS As New Dictionary(Of String, String)


    Sub New(ByVal binding As SforceService, ByVal miamigate As String, ByVal miamiods As String)
        _binding = binding
        '_miamiods = New DataBase()
        '_miamiods.ConnectionString = My.Settings.miamiods
        '_miamigate = New DataBase()
        '_miamigate.ConnectionString = My.Settings.miamigate
        _oracleConnectionGATE = New MthConnexion(miamigate)
        _oracleConnectionODS = New MthConnexion(miamiods)
    End Sub

    Function loadContact(ByVal numberOfDays As Integer) As String
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
            'Contact
            '********************
            Dim lastModifiedDate As Date
            If numberOfDays > 0 Then
                lastModifiedDate = Date.Now.AddDays(-numberOfDays)
            End If

            'truncate tables
            err = _truncateODSTables()
            result += err

            err = _queryContact(lastModifiedDate)
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
    Private Function _queryContact(ByVal lastModifiedDate As Date) As String
        Dim errors As String
        Dim sb As StringBuilder = New StringBuilder()

        Try
            Dim done As Boolean = False
            Dim query As String = "Select ID, LastModifiedDate, AccountId, Name, AssistantName, AssistantPhone, Birthdate, Department, Email, " + _
            "LeadSource, MobilePhone, ReportsToId, Title, Account_name_partner__c, Boardmember_contact__c, " + _
            "Contact_priority__c, Date_non_active__c, Department_qualification__c, DMU_relationship__c, GCS_Vertical_lookup__c, " + _
            "Included_for_Marketing__c, Level__c, partner_function_title__c, Partner_name__c, SPOC_Specialty__c, " + _
            "Start_Date_Relation__c, Status__c, CreatedDate, CreatedById FROM Contact "

            If lastModifiedDate > Date.MinValue Then
                Dim sLastModifiedDate As String = lastModifiedDate.ToString("yyyy-MM-ddThh:mm:ssZ")
                Dim where As String = String.Format(" WHERE LastModifiedDate >= {0}", sLastModifiedDate)
                query = query + where
            End If
            Dim result As QueryResult = _binding.query(query)
            If result.size > 0 Then
                Console.WriteLine(String.Format("# Contact: {0}", result.size))
                Dim parameters As TableauParametres = New TableauParametres()

                While Not done
                    Dim objects() As sObject = result.records
                    Dim count = objects.Length
                    For i As Integer = 0 To count - 1
                        Dim Contact As Contact = objects(i)


                        'Dim recordType As RecordType = Attainable_potential_Randstad.RecordType
                        Debug.WriteLine(String.Format("{0} - {1}", Contact.Name, Contact.LastModifiedDate))
                        Console.WriteLine(String.Format("{0} - {1}", Contact.Name, Contact.LastModifiedDate))

                        If i >= 107 Then
                            i = i
                        End If
                        Dim insert As String = "Insert Into TMPSF_CONTACTS" + _
                        "(ID_CONTACT, ID_ACCOUNT, NAME, ASST_NAME, ASST_PHONE, BIRTH_DATE,DEPARTMENT,EMAIL, " + _
                        "LEAD_SOURCE, MOBILE_PHONE, REPORTS_TO, TITLE, IA_NAME_PARTNER, BRD_MBR_CONTACT," + _
                        "CONTACT_PRIORITY, DATE_NACTIVE, DEP_QUALIF, DMU_RSHIP, GCS_VERTICAL," + _
                        "INCLD_FMRKETN, LEVEL_IMPACT, PRN_FUNC_TIT, PARTNER_NAME, SPOC_SPLTY," + _
                        "START_DATE_REL, STATUS, CREATEDDATE, CREATEDBYID)" + _
                        " Values(:ID_CONTACT, :ID_ACCOUNT, :NAME, :ASST_NAME, :ASST_PHONE, :BIRTH_DATE, :DEPARTMENT, :EMAIL, " + _
                        ":LEAD_SOURCE, :MOBILE_PHONE, :REPORTS_TO, :TITLE, :IA_NAME_PARTNER, :BRD_MBR_CONTACT," + _
                        ":CONTACT_PRIORITY, :DATE_NACTIVE, :DEP_QUALIF, :DMU_RSHIP, :GCS_VERTICAL," + _
                        ":INCLD_FMRKETN, :LEVEL_IMPACT, :PRN_FUNC_TIT, :PARTNER_NAME, :SPOC_SPLTY," + _
                        ":START_DATE_REL, :STATUS, :CREATEDDATE, :CREATEDBYID)"
                        parameters.PurgeParametre()
                        parameters.AjouterParametreChaine(":ID_CONTACT", Contact.Id)
                        If Not IsNothing(Contact.AccountId) Then
                            parameters.AjouterParametreChaine(":ID_ACCOUNT", Contact.AccountId)
                        Else
                            parameters.AjouterParametre(":ID_ACCOUNT", String.Empty)
                        End If
                        If Not IsNothing(Contact.Name) Then
                            parameters.AjouterParametreChaine(":NAME", Contact.Name)
                        Else
                            parameters.AjouterParametre(":NAME", String.Empty)
                        End If
                        If Not IsNothing(Contact.AssistantName) Then
                            parameters.AjouterParametreChaine(":ASST_NAME", Contact.AssistantName)
                        Else
                            parameters.AjouterParametre(":ASST_NAME", String.Empty)
                        End If
                        If Not IsNothing(Contact.AssistantPhone) Then
                            parameters.AjouterParametreChaine(":ASST_PHONE", Contact.AssistantPhone)
                        Else
                            parameters.AjouterParametre(":ASST_PHONE", String.Empty)
                        End If
                        If Not IsNothing(Contact.Birthdate) Then
                            parameters.AjouterParametreChaine(":BIRTH_DATE", Contact.Birthdate)
                        Else
                            parameters.AjouterParametre(":BIRTH_DATE", String.Empty)
                        End If
                        If Not IsNothing(Contact.Department) Then
                            parameters.AjouterParametreChaine(":DEPARTMENT", Contact.Department)
                        Else
                            parameters.AjouterParametre(":DEPARTMENT", String.Empty)
                        End If
                        If Not IsNothing(Contact.Email) Then
                            parameters.AjouterParametreChaine(":EMAIL", Contact.Email)
                        Else
                            parameters.AjouterParametre(":EMAIL", String.Empty)
                        End If
                        If Not IsNothing(Contact.LeadSource) Then
                            parameters.AjouterParametreChaine(":LEAD_SOURCE", _SF_CTC_LEADSOURCE(Contact.LeadSource))
                        Else
                            parameters.AjouterParametre(":LEAD_SOURCE", String.Empty)
                        End If

                        If Not IsNothing(Contact.MobilePhone) Then
                            parameters.AjouterParametreChaine(": MOBILE_PHONE", Contact.MobilePhone)
                        Else
                            parameters.AjouterParametre(":MOBILE_PHONE", String.Empty)
                        End If

                        If Not IsNothing(Contact.ReportsToId) Then
                            parameters.AjouterParametreChaine(":REPORTS_TO", Contact.ReportsToId)
                        Else
                            parameters.AjouterParametre(":REPORTS_TO", String.Empty)
                        End If

                        If Not IsNothing(Contact.Title) Then
                            parameters.AjouterParametreChaine(":TITLE", Contact.Title)
                        Else
                            parameters.AjouterParametre(":TITLE", String.Empty)
                        End If

                        If Not IsNothing(Contact.Account_name_partner__c) Then
                            parameters.AjouterParametreChaine(":IA_NAME_PARTNER", Contact.Account_name_partner__c)
                        Else
                            parameters.AjouterParametre(":IA_NAME_PARTNER", String.Empty)
                        End If


                        Dim err As String = Utils._multiPicklist(Contact.Boardmember_contact__c, _SF_CTC_BDMBR_CTC, ":BRD_MBR_CONTACT", parameters)
                        If err <> String.Empty Then
                            sb.Append("<br/>Boardmember_contact__c: ")
                            sb.Append(err)
                        End If

                        'If Not IsNothing(Contact.Boardmember_contact__c) Then
                        '    parameters.AjouterParametreChaine(":BRD_MBR_CONTACT", _SF_CTC_BDMBR_CTC(Contact.Boardmember_contact__c))
                        'Else
                        '    parameters.AjouterParametre(":BRD_MBR_CONTACT", String.Empty)
                        'End If

                        If Not IsNothing(Contact.Contact_priority__c) Then
                            parameters.AjouterParametreChaine(":CONTACT_PRIORITY", _SF_CTC_PRIORITY(Contact.Contact_priority__c))
                        Else
                            parameters.AjouterParametre(":CONTACT_PRIORITY", String.Empty)
                        End If

                        If Not IsNothing(Contact.Date_non_active__c) Then
                            parameters.AjouterParametreChaine(":DATE_NACTIVE", Contact.Date_non_active__c)
                        Else
                            parameters.AjouterParametre(":DATE_NACTIVE", String.Empty)
                        End If

                        If Not IsNothing(Contact.Department_qualification__c) Then
                            parameters.AjouterParametreChaine(":DEP_QUALIF", _SF_CTC_DEPT_QUAL(Contact.Department_qualification__c))
                        Else
                            parameters.AjouterParametre(":DEP_QUALIF", String.Empty)
                        End If

                        If Not IsNothing(Contact.DMU_relationship__c) Then
                            parameters.AjouterParametreChaine(":DMU_RSHIP", _SF_CTC_DMU_RSHIP(Contact.DMU_relationship__c))
                        Else
                            parameters.AjouterParametre(":DMU_RSHIP", String.Empty)
                        End If

                        If Not IsNothing(Contact.GCS_Vertical_lookup__c) Then
                            parameters.AjouterParametreChaine(":GCS_VERTICAL", Contact.GCS_Vertical_lookup__c)
                        Else
                            parameters.AjouterParametre(":GCS_VERTICAL", String.Empty)
                        End If

                        If Not IsNothing(Contact.Included_for_Marketing__c) Then
                            parameters.AjouterParametreChaine(":INCLD_FMRKETN", Contact.Included_for_Marketing__c)
                        Else
                            parameters.AjouterParametre(":INCLD_FMRKETN", String.Empty)
                        End If

                        If Not IsNothing(Contact.Level__c) AndAlso _SF_CTC_LEVEL_IMP.ContainsKey(Contact.Level__c) Then
                            parameters.AjouterParametreChaine(":LEVEL_IMPACT", _SF_CTC_LEVEL_IMP(Contact.Level__c))
                        Else
                            parameters.AjouterParametre(":LEVEL_IMPACT", String.Empty)
                        End If


                        If Not IsNothing(Contact.partner_function_title__c) Then
                            parameters.AjouterParametreChaine(":PRN_FUNC_TIT", Contact.partner_function_title__c)
                        Else
                            parameters.AjouterParametre(":PRN_FUNC_TIT", String.Empty)
                        End If

                        If Not IsNothing(Contact.Partner_name__c) Then
                            parameters.AjouterParametreChaine(":PARTNER_NAME", Contact.Partner_name__c)
                        Else
                            parameters.AjouterParametre(":PARTNER_NAME", String.Empty)
                        End If


                        err = Utils._multiPicklist(Contact.SPOC_Specialty__c, _SF_CTC_SPOC_SPLTY, ":SPOC_SPLTY", parameters)
                        If err <> String.Empty Then
                            sb.Append("<br/>SPOC_Specialty__c: ")
                            sb.Append(err)
                        End If
                        'If Not IsNothing(Contact.SPOC_Specialty__c) Then
                        '    parameters.AjouterParametreChaine(":SPOC_SPLTY", _SF_CTC_SPOC_SPLTY(Contact.SPOC_Specialty__c))
                        'Else
                        '    parameters.AjouterParametre(":SPOC_SPLTY", String.Empty)
                        'End If

                        If Not IsNothing(Contact.Start_Date_Relation__c) Then
                            parameters.AjouterParametreChaine(":START_DATE_REL", Contact.Start_Date_Relation__c)
                        Else
                            parameters.AjouterParametre(":START_DATE_REL", String.Empty)
                        End If

                        If Not IsNothing(Contact.Status__c) Then
                            parameters.AjouterParametreChaine(":STATUS", _SF_CTC_STATUS(Contact.Status__c))
                        Else
                            parameters.AjouterParametre(":STATUS", String.Empty)
                        End If

                        If Not IsNothing(Contact.CreatedDate) Then
                            parameters.AjouterParametreChaine(":CREATEDDATE", Contact.CreatedDate)
                        Else
                            parameters.AjouterParametre(":CREATEDDATE", String.Empty)
                        End If

                        If Not IsNothing(Contact.CreatedById) Then
                            parameters.AjouterParametreChaine(":CREATEDBYID", Contact.CreatedById)
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
                Console.WriteLine("No Contact found in Salesforce")
                sb.Append("No Contact found in Salesforce")
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
        '
        'Contact LEAD SOURCE
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_CTC_LEADSOURCE%'"
        Dim dataTable As New DataTable
        Dim sqlError As String = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_CTC_LEADSOURCE.Add(SFcode, MiamiCode)
        Next 'read

        'Contact BOARD MEMBER CONTACT
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_CTC_BDMBR_CTC%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_CTC_BDMBR_CTC.Add(SFcode, MiamiCode)
        Next 'read

        'Contact CONTACT PRIORITY
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_CTC_PRIORITY%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_CTC_PRIORITY.Add(SFcode, MiamiCode)
        Next 'read

        'Contact DEPARTMENT QUALIFICATION
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_CTC_DEPT_QUAL%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_CTC_DEPT_QUAL.Add(SFcode, MiamiCode)
        Next 'read

        'Contact DMU RELATIONSHIP
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_CTC_DMU_RSHIP%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_CTC_DMU_RSHIP.Add(SFcode, MiamiCode)
        Next 'read

        'Contact LEVEL - IMPACT
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_CTC_LEVEL_IMP%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_CTC_LEVEL_IMP.Add(SFcode, MiamiCode)
        Next 'read

        'Contact SPOC SPECIALTY 
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_CTC_SPOC_SPLTY%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_CTC_SPOC_SPLTY.Add(SFcode, MiamiCode)
        Next 'read

        'Contact STATUS 
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_CTC_STATUS%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_CTC_STATUS.Add(SFcode, MiamiCode)
        Next 'read

        '
        'SF Picklist check
        '
        Dim describeSObjectResult As DescribeSObjectResult = _binding.describeSObject("Contact")
        Dim fields() As Field = describeSObjectResult.fields
        For Each field As Field In fields

            If field.type.Equals(fieldType.picklist) OrElse field.type.Equals(fieldType.multipicklist) Then
                Debug.WriteLine("*** " + field.name + " ***")

                'Contact LEAD SOURCE LeadSource
                If field.name.ToLower.Equals("LeadSource".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_CTC_LEADSOURCE.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Contact BOARD MEMBER CONTACT 
                If field.name.ToLower.Equals("Boardmember_contact__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_CTC_BDMBR_CTC.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Contact CONTACT PRIORITY 
                If field.name.ToLower.Equals("Contact_priority__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_CTC_PRIORITY.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                'Contact DEPARTMENT QUALIFICATION 

                If field.name.ToLower.Equals("Department_qualification__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_CTC_DEPT_QUAL.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'Contact DMU RELATIONSHIP  
                If field.name.ToLower.Equals("DMU_relationship__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_CTC_DMU_RSHIP.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'Contact LEVEL - IMPACT 

                If field.name.ToLower.Equals("Level__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_CTC_LEVEL_IMP.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'Contact SPOC SPECIALTY 
                If field.name.ToLower.Equals("SPOC_Specialty__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_CTC_SPOC_SPLTY.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'Contact STATUS 
                If field.name.ToLower.Equals("Status__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_CTC_STATUS.ContainsKey(pickListEntry.value) Then
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
        tableList.Add("TMPSF_CONTACTS")
        'tableList.Add("TMPSF_GLO_POT_ACC")
        'tableList.Add("TMPSF_LEADS")
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
