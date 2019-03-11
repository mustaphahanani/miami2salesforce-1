Imports miami2salesforce.salesforce
Imports Methodes.mthconnexion
Imports System.Text

Public Class load_Attainable_potential_Randstad
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

    Private _SF_APR_COUNTRY As New Dictionary(Of String, String)
    Private _SF_APR_REGION As New Dictionary(Of String, String)
    Private _SF_APR_SRV_TYPE As New Dictionary(Of String, String)
    Private _SF_APR_CTL_PROC As New Dictionary(Of String, String)
    Private _SF_APR_MKT_STA As New Dictionary(Of String, String)
    Private _SF_APR_MKT_PRO As New Dictionary(Of String, String)
    Private _SF_APR_SPC_CTRY As New Dictionary(Of String, String)

  

    Sub New(ByVal binding As SforceService, ByVal miamigate As String, ByVal miamiods As String)
        _binding = binding
        '_miamiods = New DataBase()
        '_miamiods.ConnectionString = My.Settings.miamiods
        '_miamigate = New DataBase()
        '_miamigate.ConnectionString = My.Settings.miamigate
        _oracleConnectionGATE = New MthConnexion(miamigate)
        _oracleConnectionODS = New MthConnexion(miamiods)
    End Sub

    Function load_Attainable_potential_Randstad(ByVal numberOfDays As Integer) As String
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
            'Attainable potential Randstad
            '********************
            Dim lastModifiedDate As Date
            If numberOfDays > 0 Then
                lastModifiedDate = Date.Now.AddDays(-numberOfDays)
            End If

            'truncate tables
            err = _truncateODSTables()
            result += err

            err = _queryAttainable_potential_Randstad(lastModifiedDate)
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
    Private Function _queryAttainable_potential_Randstad(ByVal lastModifiedDate As Date) As String
        Dim errors As String
        Dim sb As StringBuilder = New StringBuilder()

        Try
            Dim done As Boolean = False
            Dim query As String = "SELECT ID, Name, LastModifiedDate, Account_lookup__c, Attainable_Potential_Sales__c," + _
            "Central_procurement_mandated__c, SPOC_for_the_country__c, Increase_marketshare_staffing__c, Increase_marketshare_Professionals__c," + _
            "country__c, Region__c, Service_Type__c, CreatedDate, CreatedById FROM Attainable_potential_Randstad__c"
            If lastModifiedDate > Date.MinValue Then
                Dim sLastModifiedDate As String = lastModifiedDate.ToString("yyyy-MM-ddThh:mm:ssZ")
                Dim where As String = String.Format(" WHERE LastModifiedDate >= {0}", sLastModifiedDate)
                query = query + where
            End If
            Dim result As QueryResult = _binding.query(query)
            If result.size > 0 Then
                Console.WriteLine(String.Format("# Attainable_potential_Randstad: {0}", result.size))
                Dim parameters As TableauParametres = New TableauParametres()

                While Not done
                    Dim objects() As sObject = result.records
                    Dim count = objects.Length
                    For i As Integer = 0 To count - 1
                        Dim Attainable_potential_Randstad As Attainable_potential_Randstad__c = objects(i)


                        'Dim recordType As RecordType = Attainable_potential_Randstad.RecordType
                        Debug.WriteLine(String.Format("{0} - {1}", Attainable_potential_Randstad.Name, Attainable_potential_Randstad.LastModifiedDate))
                        Console.WriteLine(String.Format("{0} - {1}", Attainable_potential_Randstad.Name, Attainable_potential_Randstad.LastModifiedDate))

                        If i >= 107 Then
                            i = i
                        End If
                        Dim insert As String = "Insert Into TMPSF_ATT_POT_RAND" + _
                        "(ID_ATT_POT, APR_NAME, ID_ACCOUNT, APR_SALES, " + _
                        "CENTRAL_PROC_RAND, INCREASE_MKT_STAF, INCREASE_MKT_PROF, SPOC_COUNTRY, " + _
                        "COUNTRY, REGION, SERVICE_TYPE, CREATEDDATE, CREATEDBYID)" + _
                        " Values(:ID_ATT_POT, :APR_NAME, :ID_ACCOUNT, :APR_SALES," + _
                        ":CENTRAL_PROC_RAND, :INCREASE_MKT_STAF, :INCREASE_MKT_PROF, :SPOC_COUNTRY, " + _
                        ":COUNTRY, :REGION, :SERVICE_TYPE, :CREATEDDATE, :CREATEDBYID)"

                        parameters.PurgeParametre()
                        parameters.AjouterParametreChaine(":ID_ATT_POT", Attainable_potential_Randstad.Id)
                        If Not IsNothing(Attainable_potential_Randstad.Name) Then
                            parameters.AjouterParametreChaine(":APR_NAME", Attainable_potential_Randstad.Name)
                        Else
                            parameters.AjouterParametre(":APR_NAME", String.Empty)
                        End If

                        If Not IsNothing(Attainable_potential_Randstad.Account_lookup__c) Then
                            parameters.AjouterParametreChaine(":ID_ACCOUNT", Attainable_potential_Randstad.Account_lookup__c)
                        Else
                            parameters.AjouterParametre(":ID_ACCOUNT", String.Empty)

                        End If

                        If Not IsNothing(Attainable_potential_Randstad.Attainable_Potential_Sales__c) Then
                            parameters.AjouterParametreChaine(":APR_SALES", Attainable_potential_Randstad.Attainable_Potential_Sales__c)
                        Else
                            parameters.AjouterParametre(":APR_SALES", String.Empty)

                        End If

                        If Not IsNothing(Attainable_potential_Randstad.Central_procurement_mandated__c) Then
                            parameters.AjouterParametreChaine(":CENTRAL_PROC_RAND", _SF_APR_CTL_PROC(Attainable_potential_Randstad.Central_procurement_mandated__c))
                        Else
                            parameters.AjouterParametre(":CENTRAL_PROC_RAND", String.Empty)
                        End If

                        If Not IsNothing(Attainable_potential_Randstad.Increase_marketshare_staffing__c) Then
                            parameters.AjouterParametreChaine(":INCREASE_MKT_STAF", _SF_APR_MKT_STA(Attainable_potential_Randstad.Increase_marketshare_staffing__c))
                        Else
                            parameters.AjouterParametre(":INCREASE_MKT_STAF", String.Empty)
                        End If

                        If Not IsNothing(Attainable_potential_Randstad.Increase_marketshare_Professionals__c) Then
                            parameters.AjouterParametreChaine(":INCREASE_MKT_PROF", _SF_APR_MKT_PRO(Attainable_potential_Randstad.Increase_marketshare_Professionals__c))
                        Else
                            parameters.AjouterParametre(":INCREASE_MKT_PROF", String.Empty)
                        End If

                        If Not IsNothing(Attainable_potential_Randstad.SPOC_for_the_country__c) Then
                            parameters.AjouterParametreChaine(":SPOC_COUNTRY", _SF_APR_SPC_CTRY(Attainable_potential_Randstad.SPOC_for_the_country__c))
                        Else
                            parameters.AjouterParametre(":SPOC_COUNTRY", String.Empty)
                        End If


                        If Not IsNothing(Attainable_potential_Randstad.country__c) Then
                            parameters.AjouterParametreChaine(":COUNTRY", _SF_APR_COUNTRY(Attainable_potential_Randstad.country__c))
                        Else
                            parameters.AjouterParametre(":COUNTRY", String.Empty)

                        End If

                        If Not IsNothing(Attainable_potential_Randstad.Region__c) Then
                            parameters.AjouterParametreChaine(":REGION", _SF_APR_REGION(Attainable_potential_Randstad.Region__c))
                        Else
                            parameters.AjouterParametre(":REGION", String.Empty)

                        End If
                        If Not IsNothing(Attainable_potential_Randstad.Service_Type__c) Then
                            parameters.AjouterParametreChaine(":SERVICE_TYPE", _SF_APR_SRV_TYPE(Attainable_potential_Randstad.Service_Type__c))
                        Else
                            parameters.AjouterParametre(":SERVICE_TYPE", String.Empty)
                        End If
                        If Not IsNothing(Attainable_potential_Randstad.CreatedDate) Then
                            parameters.AjouterParametreChaine(":CREATEDDATE", Attainable_potential_Randstad.CreatedDate)
                        Else
                            parameters.AjouterParametre(":CREATEDDATE", String.Empty)
                        End If

                        If Not IsNothing(Attainable_potential_Randstad.CreatedById) Then
                            parameters.AjouterParametreChaine(":CREATEDBYID", Attainable_potential_Randstad.CreatedById)
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
                Console.WriteLine("No Attainable potential Randstad found in Salesforce")
                sb.Append("No Attainable potential Randstad found in Salesforce")
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
        'Attainable potential Randstad COUNTRY
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_APR_COUNTRY%'"
        Dim dataTable As New DataTable
        Dim sqlError As String = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_APR_COUNTRY.Add(SFcode, MiamiCode)
        Next 'read
        'Attainable potential Randstad REGION
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_APR_REGION%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In DataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_APR_REGION.Add(SFcode, MiamiCode)
        Next 'read
        'Attainable potential Randstad SERVICE TYPE
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_APR_SRV_TYPE%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In DataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_APR_SRV_TYPE.Add(SFcode, MiamiCode)
        Next 'read

        

        'Central procurement mandated
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_APR_CTL_PROC%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_APR_CTL_PROC.Add(SFcode, MiamiCode)
        Next 'read

        'Increase marketshare Professionals
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_APR_MKT_PRO%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_APR_MKT_PRO.Add(SFcode, MiamiCode)
        Next 'read

        'Increase marketshare staffing
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_APR_MKT_PRO%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_APR_MKT_STA.Add(SFcode, MiamiCode)
        Next 'read

        'SPOC for the country
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_APR_SPC_CTRY%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_APR_SPC_CTRY.Add(SFcode, MiamiCode)
        Next 'read


        '
        'SF Picklist check
        '
        Dim describeSObjectResult As DescribeSObjectResult = _binding.describeSObject("Attainable_potential_Randstad__c")
        Dim fields() As Field = describeSObjectResult.fields
        For Each field As Field In fields

            If field.type.Equals(fieldType.picklist) OrElse field.type.Equals(fieldType.multipicklist) Then
                Debug.WriteLine("*** " + field.name + " ***")

                'Attainable potential Randstad COUNTRY
                If field.name.ToLower.Equals("country__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_APR_COUNTRY.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'Attainable potential Randstad REGION
                If field.name.ToLower.Equals("Region__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_APR_REGION.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'Attainable potential Randstad SERVICE TYPE
                If field.name.ToLower.Equals("Service_Type__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_APR_SRV_TYPE.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                If field.name.ToLower.Equals("Central_procurement_mandated__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_APR_CTL_PROC.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                If field.name.ToLower.Equals("Increase_marketshare_Professionals__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_APR_MKT_PRO.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                If field.name.ToLower.Equals("Increase_marketshare_staffing__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_APR_MKT_STA.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If

                If field.name.ToLower.Equals("SPOC_for_the_country__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_APR_SPC_CTRY.ContainsKey(pickListEntry.value) Then
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
        tableList.Add("TMPSF_ATT_POT_RAND")
        'tableList.Add("TMPSF_CONTACTS")
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



