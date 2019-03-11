Imports miami2salesforce.salesforce
Imports Methodes.mthconnexion
Imports System.Text

Public Class load_Global_potential_account
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

    Private _SF_GPA_ACC_CTRY As New Dictionary(Of String, String)
    Private _SF_GPA_ACC_REG As New Dictionary(Of String, String)
    Private _SF_GPA_SRV_TYP As New Dictionary(Of String, String)

    Sub New(ByVal binding As SforceService, ByVal miamigate As String, ByVal miamiods As String)
        _binding = binding
        '_miamiods = New DataBase()
        '_miamiods.ConnectionString = My.Settings.miamiods
        '_miamigate = New DataBase()
        '_miamigate.ConnectionString = My.Settings.miamigate
        _oracleConnectionGATE = New MthConnexion(miamigate)
        _oracleConnectionODS = New MthConnexion(miamiods)
    End Sub

    Function load_Global_potential_account(ByVal numberOfDays As Integer) As String
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
            'Global potential account
            '********************
            Dim lastModifiedDate As Date
            If numberOfDays > 0 Then
                lastModifiedDate = Date.Now.AddDays(-numberOfDays)
            End If

            'truncate tables
            err = _truncateODSTables()
            result += err

            err = _queryGlobal_potential_account(lastModifiedDate)
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

    Private Function _queryGlobal_potential_account(ByVal lastModifiedDate As Date) As String
        Dim errors As String
        Dim sb As StringBuilder = New StringBuilder()

        Try
            Dim done As Boolean = False
            Dim query As String = "SELECT ID, Name, LastModifiedDate, Account__c, country__c, Global_Potential_Sales__c, " + _
            "Region__c, Service_Type__c, CreatedDate, CreatedById FROM Global_potential_account__c"
            If lastModifiedDate > Date.MinValue Then
                Dim sLastModifiedDate As String = lastModifiedDate.ToString("yyyy-MM-ddThh:mm:ssZ")
                Dim where As String = String.Format(" WHERE LastModifiedDate >= {0}", sLastModifiedDate)
                query = query + where
            End If
            Dim result As QueryResult = _binding.query(query)
            If result.size > 0 Then
                Console.WriteLine(String.Format("# Global_potential_account: {0}", result.size))
                Dim parameters As TableauParametres = New TableauParametres()

                While Not done
                    Dim objects() As sObject = result.records
                    Dim count = objects.Length
                    For i As Integer = 0 To count - 1
                        Dim Global_potential_account As Global_potential_account__c = objects(i)


                        'Dim recordType As RecordType = Global_potential_account.RecordType
                        Debug.WriteLine(String.Format("{0} - {1}", Global_potential_account.Name, Global_potential_account.LastModifiedDate))
                        Console.WriteLine(String.Format("{0} - {1}", Global_potential_account.Name, Global_potential_account.LastModifiedDate))

                        If i >= 107 Then
                            i = i
                        End If
                        Dim insert As String = "Insert Into TMPSF_GLO_POT_ACC" + _
                        "(ID_GLO_POT, GPA_NAME, ID_ACCOUNT, COUNTRY, GPA_SALES," + _
                        "REGION, SERVICE_TYPE, CREATEDDATE, CREATEDBYID)" + _
                        " Values(:ID_GLO_POT, :GPA_NAME, :ID_ACCOUNT, :COUNTRY, :GPA_SALES, " + _
                        ":REGION, :SERVICE_TYPE, :CREATEDDATE, :CREATEDBYID)"

                        parameters.PurgeParametre()
                        parameters.AjouterParametreChaine(":ID_ATT_POT", Global_potential_account.Id)
                        If Not IsNothing(Global_potential_account.Name) Then
                            parameters.AjouterParametreChaine(":GPA_NAME", Global_potential_account.Name)
                        Else
                            parameters.AjouterParametre(":GPA_NAME", String.Empty)
                        End If

                        If Not IsNothing(Global_potential_account.Account__c) Then
                            parameters.AjouterParametreChaine(":ID_ACCOUNT", Global_potential_account.Account__c)
                        Else
                            parameters.AjouterParametre(":ID_ACCOUNT", String.Empty)
                        End If

                        If Not IsNothing(Global_potential_account.country__c) Then
                            parameters.AjouterParametreChaine(":COUNTRY", _SF_GPA_ACC_CTRY(Global_potential_account.country__c))
                        Else
                            parameters.AjouterParametre(":COUNTRY", String.Empty)
                        End If

                        If Not IsNothing(Global_potential_account.Global_Potential_Sales__c) Then
                            parameters.AjouterParametreChaine(":GPA_SALES", Global_potential_account.Global_Potential_Sales__c)
                        Else
                            parameters.AjouterParametre(":GPA_SALES", String.Empty)
                        End If

                        'If Not IsNothing(Global_potential_account.Notes__c) Then
                        '    parameters.AjouterParametreChaine(":NOTES", Global_potential_account.Notes__c)
                        'Else
                        '    parameters.AjouterParametre(":NOTES", String.Empty)
                        'End If

                        If Not IsNothing(Global_potential_account.Region__c) Then
                            parameters.AjouterParametreChaine(":REGION", _SF_GPA_ACC_REG(Global_potential_account.Region__c))
                        Else
                            parameters.AjouterParametre(":REGION", String.Empty)
                        End If

                        If Not IsNothing(Global_potential_account.Service_Type__c) Then
                            parameters.AjouterParametreChaine(":SERVICE_TYPE", _SF_GPA_SRV_TYP(Global_potential_account.Service_Type__c))
                        Else
                            parameters.AjouterParametre(":SERVICE_TYPE", String.Empty)
                        End If
                        If Not IsNothing(Global_potential_account.CreatedDate) Then
                            parameters.AjouterParametreChaine(":CREATEDDATE", Global_potential_account.CreatedDate)
                        Else
                            parameters.AjouterParametre(":CREATEDDATE", String.Empty)
                        End If

                        If Not IsNothing(Global_potential_account.CreatedById) Then
                            parameters.AjouterParametreChaine(":CREATEDBYID", Global_potential_account.CreatedById)
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
                Console.WriteLine("No Global potential account found in Salesforce")
                sb.Append("No Global potential account found in Salesforce")
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
        'Global potential account COUNTRY
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_GPA_ACC_CTRY%'"
        Dim dataTable As New DataTable
        Dim sqlError As String = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_GPA_ACC_CTRY.Add(SFcode, MiamiCode)
        Next 'read

        'Global potential account REGION
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_GPA_ACC_REG%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_GPA_ACC_REG.Add(SFcode, MiamiCode)
        Next 'read

        'Global potential account SERVICE TYPE
        sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_GPA_SRV_TYP%'"
        dataTable.Clear()
        sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            'Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))
            _SF_GPA_SRV_TYP.Add(SFcode, MiamiCode)
        Next 'read

        '
        'SF Picklist check
        '
        Dim describeSObjectResult As DescribeSObjectResult = _binding.describeSObject("Global_potential_account__c")
        Dim fields() As Field = describeSObjectResult.fields
        For Each field As Field In fields

            If field.type.Equals(fieldType.picklist) OrElse field.type.Equals(fieldType.multipicklist) Then
                Debug.WriteLine("*** " + field.name + " ***")

                'Global potential account COUNTRY
                If field.name.ToLower.Equals("country__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_GPA_ACC_CTRY.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'Global potential account REGION
                If field.name.ToLower.Equals("Region__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_GPA_ACC_REG.ContainsKey(pickListEntry.value) Then
                            err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
                        End If
                    Next
                End If
                'Global potential account SERVICE TYPE
                If field.name.ToLower.Equals("Service_Type__c".ToLower) Then
                    Dim pickListEntries() As PicklistEntry = field.picklistValues
                    For Each pickListEntry As PicklistEntry In pickListEntries
                        Debug.WriteLine(pickListEntry.value)
                        'Check Service type
                        If Not _SF_GPA_SRV_TYP.ContainsKey(pickListEntry.value) Then
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
        tableList.Add("TMPSF_GLO_POT_ACC")
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
