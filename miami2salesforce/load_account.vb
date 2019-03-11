Imports System.Text
Imports System.IO
Imports miami2salesforce.salesforce
Imports Methodes.mthconnexion

Public Class LoadAccount

    Private _binding As SforceService
    'Private _miamidwh As MthConnexion
    'Private _miamigate As MthConnexion

    Private _SF_ACC_SRV_TYPE As New Dictionary(Of String, String)
    Private _SF_ACC_SF_TYPE As New Dictionary(Of String, String)
    Private _SF_ACC_SF_STATUS As New Dictionary(Of String, String)
    Private _SF_ACC_VERTICAL As New Dictionary(Of String, String)
    Private _SF_ACC_DIRECTOR As New Dictionary(Of String, String)
    Private _SF_USR_IADOM As New Dictionary(Of String, String)

    Private _oracleConnectionDWH As MthConnexion
    'Private _oracleCommandDWH As OracleCommand
    'Private _oracleDataReader As OracleDataReader
    Private _oracleConnectionGATE As MthConnexion
    'Private _oracleCommandGATE As OracleCommand


    Private _toYearMonth As String
    Private _fromYearMonth As String

    Private _DEBUG As Boolean

    Sub New(ByVal binding As SforceService, ByVal miamigate As String, ByVal miamidwh As String)
        _binding = binding
        _oracleConnectionGATE = New MthConnexion(miamigate)
        _oracleConnectionDWH = New MthConnexion(miamidwh)
        '_miamidwh = New DataBase()
        '_miamidwh.ConnectionString = My.Settings.miamidwh        

        '_miamigate = New DataBase()
        '_miamigate.ConnectionString = My.Settings.miamigate

    End Sub

    Public Function loadAccount(ByVal bDebug As Boolean) As String

        Dim result As String = String.Empty
        Dim err As String = String.Empty

        '_oracleConnectionDWH.GetConnexion()
        '_oracleConnectionGATE = _miamigate.connect()
        _DEBUG = bDebug

        Try
            '*******
            'Open DB
            '*******
            '_oracleConnectionDWH.Open()
            '_oracleCommandDWH = New OracleCommand()
            '_oracleCommandDWH.Connection = _oracleConnectionDWH
            '_oracleCommandDWH.CommandType = CommandType.Text

            '_oracleConnectionGATE.Open()
            '_oracleCommandGATE = New OracleCommand()
            '_oracleCommandGATE.Connection = _oracleConnectionGATE
            '_oracleCommandGATE.CommandType = CommandType.Text

            '*********************************
            'SF Users => Miami gate parameters
            '*********************************
            err = _load_user()

            '**********
            'References
            '**********
            'err = _references()
            'If Not String.IsNullOrEmpty(err) Then
            '    Return "ERRORS: <br/>" + err
            'End If

            '**************
            'Period max/min
            '**************
            '_Periods()

            '_Attachment("f:\\temp\\export.xlsx")

            '********************
            'Account and Sublevel
            '********************

            result = _level1()
            'err = _level1()
            'result += err
            'err = _level2()
            'result += err
            'err = _level3()
            'result += err
            'err = _level4()
            'result += err

            '****************
            'Sales and Margin
            '****************
            'err = _resetActuals()
            'result += err

            'err = _actualsLevel1()
            'result += err
            'err = _actualsLevel2()
            'result += err
            'err = _actualsLevel3()
            'result += err
            'err = _actualsLevel4()
            'result += err

            'result += err
            'err = _actualsLevel3SumLevel4()
            'result += err
            'err = _actualsLevel2SumLevel3()
            'result += err
            'err = _actualsLevel2SumLevel4()
            'result += err

            '********
            'Close DB
            '********
            '_oracleDataReader.Close()
            '_oracleConnectionDWH.Close()
            _oracleConnectionDWH.Dispose()
            _oracleConnectionGATE.Dispose()

        Catch ex As Exception
            Console.WriteLine(ex.Message)
            result = "** ERROR **<br/><br/>" + ex.Message
        End Try

        Return result
    End Function

    Private Function _upsertAccounts(ByVal accounts As IList(Of Account), ByVal dataAccounts As IList(Of DataAccount)) As String
        Dim result As StringBuilder = New StringBuilder()

        Dim objects(accounts.Count - 1) As sObject
        Dim i As Integer = 0
        Dim ok As Boolean

        Dim sdebug As String = ""

        For Each acc As Account In accounts
            ok = False
            Dim dataAccount As DataAccount = Nothing
            If Not IsNothing(dataAccounts) Then
                For Each dataAccount In dataAccounts
                    Dim externalId As String = acc.MIAMI_account_ID__c.Split(".").GetValue(0).ToString()
                    If dataAccount.ExternalId = externalId Then
                        ok = True
                        Exit For
                    End If
                Next
            End If

            If ok Then
                'Dim owner As New User()
                'owner.Id = dataAccount.OwnerId
                acc.Type = dataAccount.Type
                ' acc.GCS_RSR_Opps_account_owner__c = dataAccount.OwnerId
                'acc.Owner = owner
                'acc.Act_Inac_in_MIAMI__c = dataAccount.ActiveInactive
                acc.Account_Status__c = dataAccount.Status
                acc.GCS_industry_sector__c = dataAccount.Industry
                acc.GCS_Vertical__c = dataAccount.Vertical
                acc.GCS_account_director__c = dataAccount.Director
                acc.OwnerId = dataAccount.OwnerId
            End If

            objects(i) = acc

            sdebug = sdebug + acc.Name + " - " + acc.MIAMI_account_ID__c + vbCrLf

            i = i + 1
        Next

        Try
            'i = 0
            Dim results() As UpsertResult = _binding.upsert("MIAMI_account_ID__c", objects)
            For Each upsertResult As UpsertResult In results
                'i = i + 1
                If upsertResult.success Then
                    'Debug.Print("Result: OK " + upsertResult.id)
                Else
                    Debug.Print("Errors:")
                    result.Append("Errors:")
                    Dim errors() As miami2salesforce.salesforce.Error = upsertResult.errors
                    For Each err As miami2salesforce.salesforce.Error In errors
                        Debug.Print(err.message)
                        result.Append("<br/>")
                        result.Append(err.message)
                    Next
                End If 'success
            Next 'upsertResult
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            result.Append(ex.Message)
        End Try

        Return result.ToString()

    End Function

    Private Function _queryAccounts(ByVal accounts As IList(Of Account), ByRef errors As String) As IList(Of DataAccount)
        Dim sb As StringBuilder = New StringBuilder()

        Dim dataAccounts As IList(Of DataAccount) = New List(Of DataAccount)

        Dim externalIds As IList(Of String) = New List(Of String)
        Dim inExternalIds As New StringBuilder()

        For Each acc As Account In accounts
            Dim externalId As String = acc.MIAMI_account_ID__c.Split(".").GetValue(0).ToString()
            'If inExternalIds.Length > 0 Then
            '    inExternalIds.Append(",")
            'End If
            'inExternalIds.Append("'")
            'inExternalIds.Append(externalId)
            'inExternalIds.Append("'")
            If Not externalIds.Contains(externalId) Then
                externalIds.Add(externalId)
            End If
        Next

        Dim externalIdsArray(externalIds.Count - 1) As String
        Dim j As Integer = 0
        For Each id As String In externalIds
            externalIdsArray(j) = id
            j += 1
        Next

        inExternalIds.Append("'")
        inExternalIds.Append(String.Join("','", externalIdsArray))
        inExternalIds.Append("'")
        'Dim inExternalIds As String = "'" + String.Join("','", externalIds) + "'"

        Try
            Dim done As Boolean = False
            'retrieve the first 500 accounts
            Dim result As QueryResult = _binding.query("Select Id, MIAMI_account_ID__c, Type, Act_Inac_in_MIAMI__c, GCS_RSR_Opps_account_owner__c, Account_Status__c, Industry_Sector__c, GCS_account_director__c, GCS_Vertical__c From Account Where MIAMI_account_ID__c In (" + inExternalIds.ToString() + ")")
            If result.size > 0 Then
                While Not done
                    Dim objects() As sObject = result.records
                    Dim count = objects.Length
                    For i As Integer = 0 To count - 1
                        Dim account As Account = objects(i)
                        Dim dataAccount As New DataAccount()
                        dataAccount.Id = account.Id
                        dataAccount.Type = account.Type

                        REM test
                        dataAccount.Director = account.country_1__c
                        dataAccount.Director = account.country_2__c
                        dataAccount.Director = account.country_3__c
                        dataAccount.Director = account.CFS_2_opco_1_pct__c
                        dataAccount.Director = account.CFS_2_opco_2_pct__c
                        dataAccount.Director = account.CFS_2_opco_3_pct__c

                        'dataAccount.ActiveInactive = account.Act_Inac_in_MIAMI__c
                        'dataAccount.OwnerId = account.OwnerId
                        dataAccount.Status = account.Account_Status__c
                        dataAccount.Industry = account.GCS_industry_sector__c
                        dataAccount.ExternalId = account.MIAMI_account_ID__c
                        ' dataAccount.OwnerId = account.GCS_RSR_Opps_account_owner__c
                        'dataAccount.Owner = account.Owner
                        dataAccount.Director = account.GCS_account_director__c
                        dataAccount.Vertical = account.GCS_Vertical__c
                        dataAccounts.Add(dataAccount)
                    Next
                    If result.done Then
                        done = True
                    Else
                        'retrieve the next 500 accounts
                        result = _binding.queryMore(result.queryLocator)
                    End If
                End While
            Else
                Console.WriteLine("No account found in Salesforce with ids: {0}", inExternalIds.ToString())
                sb.Append("No account found in Salesforce with ids: {0}", inExternalIds.ToString())
            End If

        Catch ex As Exception
            sb.Append(ex.Message)
            Console.WriteLine(ex.Message)
        End Try

        errors = sb.ToString()

        Return dataAccounts
    End Function

    'Private Function _upsertActuals(ByVal actuals As IList(Of Account_Actual_Miami__c)) As String
    '    Dim result As StringBuilder = New StringBuilder()

    '    Dim objects(actuals.Count - 1) As sObject
    '    Dim i As Integer = 0
    '    For Each actual As Account_Actual_Miami__c In actuals
    '        objects(i) = actual
    '        i = i + 1
    '    Next

    '    Try
    '        'i = 0
    '        Dim actuals_upserts() As UpsertResult = _binding.upsert("ExternalID__c", objects)
    '        For Each upsertResult As UpsertResult In actuals_upserts
    '            'i = i + 1
    '            If upsertResult.success Then
    '                'Debug.Print("Result: OK " + upsertResult.id)
    '            Else
    '                Debug.Print("Errors:")
    '                result.Append("Errors:")
    '                Dim errors() As miami2salesforce.salesforce.Error = upsertResult.errors
    '                For Each err As miami2salesforce.salesforce.Error In errors
    '                    Debug.Print(err.message)
    '                    result.Append("<br/>")
    '                    result.Append(err.message)
    '                Next
    '            End If 'success
    '        Next 'upsertResult
    '    Catch ex As Exception
    '        Console.WriteLine(ex.Message)
    '        Console.WriteLine(ex.Message)
    '    End Try

    '    Return result.ToString()

    'End Function

    Private Function _level1() As String

        Dim errors As String = String.Empty

        'International accounts (Level 1)
        Console.WriteLine("--- Level 1")
        Dim sb As New StringBuilder()
        sb.Append("Select E_INT_ACCOUNT.IA_CD, E_INT_ACCOUNT.IA_NAME, E_INT_ACCOUNT.INACTIVE_FLAG, R_IA_TYPE.TYPE_DESC, E_INDUSTRY.INDUSTRY_NAME, ")
        'TIC0102954 Replaced values for VP,SVP, Industry name, Executive, Country and Sector if it is Unknown in miami 
        sb.Append("  Replace(D.DIRECTOR_CD, '0', null) as DIRECTOR_CD, ")
        sb.Append("  Replace(E_MANAGER.MANAGER_CD, '0', null) as MANAGER_CD, ")
        sb.Append("  R_EXECUTIVE.EXECUTIVE_DESC, R_COUNTRY.COUNTRY_NAME ")
        sb.Append(" , Replace(E_SECTOR.SECTOR_NAME, '** Unknown **', ' ') as SECTOR_NAME ")
        sb.Append(" , ACCOUNT_PERCENT_TOP1.PERCENTAGE TOP1_ACCOUNT_PERCENTAGE, ACCOUNT_PERCENT_TOP1.COUNTRY_NAME TOP1_ACCOUNT_COUNTRY, ACCOUNT_PERCENT_TOP2.PERCENTAGE TOP2_ACCOUNT_PERCENTAGE, ACCOUNT_PERCENT_TOP2.COUNTRY_NAME TOP2_ACCOUNT_COUNTRY")
        sb.Append(" , ACCOUNT_PERCENT_TOP3.PERCENTAGE TOP3_ACCOUNT_PERCENTAGE, ACCOUNT_PERCENT_TOP3.COUNTRY_NAME TOP3_ACCOUNT_COUNTRY, E_INT_ACCOUNT.NB_ACTIVE_COUNTRIES")
        sb.Append(" , STAFFING.PERCENTAGE STAFFING_PERCENTAGE, INHOUSE.PERCENTAGE INHOUSE_PERCENTAGE, PROFESSIONALS.PERCENTAGE PROFESSIONALS_PERCENTAGE")
        sb.Append(" , RECRUITMENT.PERCENTAGE RECRUITMENT_PERCENTAGE, HR_SOLUTIONS.PERCENTAGE HR_SOLUTIONS_PERCENTAGE, OTHER_CONCEPTS.PERCENTAGE OTHER_CONCEPTS_PERCENTAGE")
        sb.Append("  From")
        sb.Append(" (Select E_ACC_IAD.IA_CD, listagg (E_IAD.DIRECTOR_CD, ';')")
        sb.Append(" Within Group (ORDER BY E_IAD.DIRECTOR_CD) DIRECTOR_CD")
        sb.Append(" From E_ACC_IAD")
        sb.Append(" Inner Join E_IAD On E_IAD.DIRECTOR_CD=E_ACC_IAD.DIRECTOR_CD")
        sb.Append(" Group By E_ACC_IAD.IA_CD) D,")
        sb.Append(" E_INT_ACCOUNT")
        sb.Append(" Inner Join E_INDUSTRY On E_INDUSTRY.INDUSTRY_CD=E_INT_ACCOUNT.INDUSTRY_CD")
        sb.Append(" Inner Join R_IA_TYPE On R_IA_TYPE.TYPE_CD=E_INT_ACCOUNT.IA_TYPE")
        sb.Append(" Inner Join E_MANAGER On E_MANAGER.MANAGER_CD=E_INT_ACCOUNT.MANAGER_CD")
        sb.Append(" Inner Join R_EXECUTIVE On R_EXECUTIVE.EXECUTIVE_CD=E_INT_ACCOUNT.EXECUTIVE_CD")
        sb.Append(" Inner Join R_COUNTRY On R_COUNTRY.COUNTRY_CD=E_INT_ACCOUNT.COUNTRY_HQ_CD")
        sb.Append(" Inner Join E_SECTOR On E_SECTOR.SECTOR_CD=E_INT_ACCOUNT.SECTOR_CD")
        sb.Append(" Left Join (Select IA_CD, COUNTRY_NAME, PERCENTAGE From E_ACC_TOP_COUNTRIES Inner Join R_COUNTRY On R_COUNTRY.COUNTRY_CD=E_ACC_TOP_COUNTRIES.COUNTRY_CD Where RANK=1) ACCOUNT_PERCENT_TOP1 ON ACCOUNT_PERCENT_TOP1.IA_CD=E_INT_ACCOUNT.IA_CD")
        sb.Append(" Left Join (Select IA_CD, COUNTRY_NAME, PERCENTAGE From E_ACC_TOP_COUNTRIES Inner Join R_COUNTRY On R_COUNTRY.COUNTRY_CD=E_ACC_TOP_COUNTRIES.COUNTRY_CD Where RANK=2) ACCOUNT_PERCENT_TOP2 ON ACCOUNT_PERCENT_TOP2.IA_CD=E_INT_ACCOUNT.IA_CD")
        sb.Append(" Left Join (Select IA_CD, COUNTRY_NAME, PERCENTAGE From E_ACC_TOP_COUNTRIES Inner Join R_COUNTRY On R_COUNTRY.COUNTRY_CD=E_ACC_TOP_COUNTRIES.COUNTRY_CD Where RANK=3) ACCOUNT_PERCENT_TOP3 ON ACCOUNT_PERCENT_TOP3.IA_CD=E_INT_ACCOUNT.IA_CD")
        sb.Append(" Left Join (Select IA_CD, PERCENTAGE From E_ACC_PCT_CONCEPT Where SERVICE_CD='SRC') STAFFING ON STAFFING.IA_CD=E_INT_ACCOUNT.IA_CD")
        sb.Append(" Left Join (Select IA_CD, PERCENTAGE From E_ACC_PCT_CONCEPT Where SERVICE_CD='INH') INHOUSE ON INHOUSE.IA_CD=E_INT_ACCOUNT.IA_CD")
        sb.Append(" Left Join (Select IA_CD, PERCENTAGE From E_ACC_PCT_CONCEPT Where SERVICE_CD='INP') PROFESSIONALS ON PROFESSIONALS.IA_CD=E_INT_ACCOUNT.IA_CD")
        sb.Append(" Left Join (Select IA_CD, PERCENTAGE From E_ACC_PCT_CONCEPT Where SERVICE_CD='SSL') RECRUITMENT ON RECRUITMENT.IA_CD=E_INT_ACCOUNT.IA_CD")
        sb.Append(" Left Join (Select IA_CD, PERCENTAGE From E_ACC_PCT_CONCEPT Where SERVICE_CD='HRS') HR_SOLUTIONS ON HR_SOLUTIONS.IA_CD=E_INT_ACCOUNT.IA_CD")
        sb.Append(" Left Join (Select IA_CD, PERCENTAGE From E_ACC_PCT_CONCEPT Where SERVICE_CD not in ('SRC','INH','INP','SSL','HRS')) OTHER_CONCEPTS ON OTHER_CONCEPTS.IA_CD=E_INT_ACCOUNT.IA_CD")
        sb.Append(" Where E_INT_ACCOUNT.IA_CD=D.IA_CD")
        sb.Append(" And E_INT_ACCOUNT.IA_CD>0")
        ''sb.Append(" And E_INT_ACCOUNT.INACTIVE_FLAG=0")
        'sb.Append(" AND E_INT_ACCOUNT.IA_CD = 1 ")
#If Debug Then
        'sb.Append(" Where And E_INT_ACCOUNT.IA_CD Between 100 And 200")
#End If

        Dim sql As String = sb.ToString()
        Dim count As Integer = 0
        Dim accounts As IList(Of Account) = New List(Of Account)
        Dim strnull As String = ""
        Dim dataTable As New DataTable
        Dim sqlError As String = _oracleConnectionDWH.Requete(sql, dataTable)
        'Dim dataTable As New DataTable
        'Dim sqlError As String = _oracleConnectionDWH.Requete(sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            count = count + 1
            If count > 200 Then
                errors = _upsertAccounts(accounts, Nothing)
                accounts.Clear()
                count = 1
            End If

            Dim ia_cd As String = row.Item("IA_CD").ToString
            Dim ia_name As String = row.Item("IA_NAME").ToString
            Dim INACTIVE_FLAG As Integer = row.Item("INACTIVE_FLAG")

            Dim TYPE_DESC As String = row.Item("TYPE_DESC").ToString
            Dim INDUSTRY_NAME As String = row.Item("INDUSTRY_NAME").ToString
            Dim DIRECTOR_NAMES As String = row.Item("DIRECTOR_CD").ToString()
            Dim MANAGER_NAME As String = row.Item("MANAGER_CD").ToString()
            Dim EXECUTIVE_DESC As String = row.Item("EXECUTIVE_DESC").ToString
            Dim COUNTRY_NAME As String = row.Item("COUNTRY_NAME").ToString
            Dim SECTOR_NAME As String = row.Item("SECTOR_NAME").ToString

            REM MANTIS 5729 => Export new fields to SalesForce (.Net part)
            REM the top 3 countries per account with percentual share of sales
            Dim TOP1_ACCOUNT_COUNTRY_NAME As String = row.Item("TOP1_ACCOUNT_COUNTRY").ToString
            Dim TOP1_ACCOUNT_PERCENT As Double : If Not row.Item("TOP1_ACCOUNT_PERCENTAGE") Is System.DBNull.Value Then TOP1_ACCOUNT_PERCENT = row.Item("TOP1_ACCOUNT_PERCENTAGE") Else TOP1_ACCOUNT_PERCENT = 0
            Dim TOP2_ACCOUNT_COUNTRY_NAME As String = row.Item("TOP2_ACCOUNT_COUNTRY").ToString
            Dim TOP2_ACCOUNT_PERCENT As Double : If Not row.Item("TOP2_ACCOUNT_PERCENTAGE") Is System.DBNull.Value Then TOP2_ACCOUNT_PERCENT = row.Item("TOP2_ACCOUNT_PERCENTAGE") Else TOP2_ACCOUNT_PERCENT = 0
            Dim TOP3_ACCOUNT_COUNTRY_NAME As String = row.Item("TOP3_ACCOUNT_COUNTRY").ToString
            Dim TOP3_ACCOUNT_PERCENT As Double : If Not row.Item("TOP3_ACCOUNT_PERCENTAGE") Is System.DBNull.Value Then TOP3_ACCOUNT_PERCENT = row.Item("TOP3_ACCOUNT_PERCENTAGE") Else TOP3_ACCOUNT_PERCENT = 0
            REM the percentual share per account of sales per concepts
            Dim STAFFING As Double : If Not row.Item("STAFFING_PERCENTAGE") Is System.DBNull.Value Then STAFFING = row.Item("STAFFING_PERCENTAGE") Else STAFFING = 0
            Dim INHOUSE As Double : If Not row.Item("INHOUSE_PERCENTAGE") Is System.DBNull.Value Then INHOUSE = row.Item("INHOUSE_PERCENTAGE") Else INHOUSE = 0
            Dim PROFESSIONALS As Double : If Not row.Item("PROFESSIONALS_PERCENTAGE") Is System.DBNull.Value Then PROFESSIONALS = row.Item("PROFESSIONALS_PERCENTAGE") Else PROFESSIONALS = 0
            Dim RECRUITMENT As Double : If Not row.Item("RECRUITMENT_PERCENTAGE") Is System.DBNull.Value Then RECRUITMENT = row.Item("RECRUITMENT_PERCENTAGE") Else RECRUITMENT = 0
            Dim HR_SOLUTIONS As Double : If Not row.Item("HR_SOLUTIONS_PERCENTAGE") Is System.DBNull.Value Then HR_SOLUTIONS = row.Item("HR_SOLUTIONS_PERCENTAGE") Else HR_SOLUTIONS = 0
            Dim OTHER_CONCEPTS_PERCENTAGE As Double : If Not row.Item("OTHER_CONCEPTS_PERCENTAGE") Is System.DBNull.Value Then OTHER_CONCEPTS_PERCENTAGE = row.Item("OTHER_CONCEPTS_PERCENTAGE") Else OTHER_CONCEPTS_PERCENTAGE = 0
            REM total spend last 12 months
            Dim TOTAL_SALE_12MONTHS As Double : If Not row.Item("TOTAL_SALE_12MONTHS") Is System.DBNull.Value Then TOTAL_SALE_12MONTHS = row.Item("TOTAL_SALE_12MONTHS")
            REM total number of active countries last 12 months
            Dim NB_ACTIVE_COUNTRIES As Double : If Not row.Item("NB_ACTIVE_COUNTRIES") Is System.DBNull.Value Then NB_ACTIVE_COUNTRIES = row.Item("NB_ACTIVE_COUNTRIES")

            Dim account As New Account()
            account.Name = ia_name
            account.MIAMI_account_ID__c = ia_cd
            account.Miami_Level__c = "1"

            If INACTIVE_FLAG > 0 Then
                account.Act_Inac_in_MIAMI__c = "Inactive"
            Else
                account.Act_Inac_in_MIAMI__c = "Active"
            End If
            Dim fieldsToNull As New List(Of String)
            account.Account_Status__c = TYPE_DESC
            If DIRECTOR_NAMES <> "" Then
                account.SVP__c = GetSalesForcesID(DIRECTOR_NAMES)
            Else
                fieldsToNull.Add("SVP__c")
            End If
            If MANAGER_NAME <> "" Then
                account.VP__c = GetSalesForcesID(MANAGER_NAME)
            Else
                fieldsToNull.Add("VP__c")
            End If

            If fieldsToNull.Count > 0 Then
                account.fieldsToNull = fieldsToNull.ToArray()
            End If


            account.Executive_involved__c = EXECUTIVE_DESC
            account.GCS_Vertical__c = INDUSTRY_NAME
            account.GCS_industry_sector__c = SECTOR_NAME
            account.Location_global_HQ__c = COUNTRY_NAME

            REM MANTIS 5729 => Export new fields to SalesForce (.Net part)
            account.country_1__c = TOP1_ACCOUNT_COUNTRY_NAME
            account.CFS_2_opco_1_pct__c = TOP1_ACCOUNT_PERCENT
            account.CFS_2_opco_1_pct__cSpecified = True
            account.country_2__c = TOP2_ACCOUNT_COUNTRY_NAME
            account.CFS_2_opco_2_pct__c = TOP2_ACCOUNT_PERCENT
            account.CFS_2_opco_2_pct__cSpecified = True
            account.country_3__c = TOP3_ACCOUNT_COUNTRY_NAME
            account.CFS_2_opco_3_pct__c = TOP3_ACCOUNT_PERCENT
            account.CFS_2_opco_3_pct__cSpecified = True
            account.CFS_2_staffing__c = STAFFING
            account.CFS_2_staffing__cSpecified = True
            account.inhouse__c = INHOUSE
            account.inhouse__cSpecified = True
            account.CFS_2_professionals__c = PROFESSIONALS
            account.CFS_2_professionals__cSpecified = True
            account.recruitment__c = RECRUITMENT
            account.recruitment__cSpecified = True
            account.CFS_2_hr_solutions__c = HR_SOLUTIONS
            account.CFS_2_hr_solutions__cSpecified = True
            If TOTAL_SALE_12MONTHS > 0 Then
                account.CFS_2_global_randstad_spend__c = Math.Round(TOTAL_SALE_12MONTHS / 1000000, 1)
            Else
                account.CFS_2_global_randstad_spend__c = 0
            End If
            account.CFS_2_global_randstad_spend__cSpecified = True
            account.CFS_2_active_serviced_countries__c = NB_ACTIVE_COUNTRIES.ToString

            accounts.Add(account)
        Next

        If count > 0 Then
            errors = _upsertAccounts(accounts, Nothing)
        End If

        Console.WriteLine("--- Level 1 OK")

        errors += "<br/>--- Level 1 OK"
        Return errors

    End Function

    Private Function GetSalesForcesID(ByVal directorCode As String) As String
        Dim sb As New StringBuilder()
        Dim dataTable As New DataTable
        Dim sFID As String = String.Empty
        sb.Append("SELECT * FROM Param_r WHERE param_cd like 'SF_USR_IADOM%' AND param_value LIKE '%" + directorCode + "%'")
        Dim sqlError As String = _oracleConnectionGATE.Requete(sb.ToString, dataTable)

        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            sFID = codes(1)
        Next
        Return sFID
    End Function


    'Private Function _level2() As String
    '    Dim errors As String = String.Empty

    '    'Sublevel (Level 2)
    '    Console.WriteLine("--- Level 2")
    '    Dim sb As New StringBuilder()
    '    sb.Append("Select Distinct E_LEVEL.IA_CD, E_LEVEL.IA_CD_N2, E_LEVEL.IA_NAME_N2, E_LEVEL.INACTIVE_FLAG")
    '    sb.Append(" From E_LEVEL")
    '    sb.Append(" Where E_LEVEL.IA_CD_N2 > 0")
    '    sb.Append(" And Exists (")
    '    sb.Append(" Select * From E_INT_ACCOUNT")
    '    sb.Append(" Where E_INT_ACCOUNT.IA_CD = E_LEVEL.IA_CD")
    '    sb.Append(" And E_INT_ACCOUNT.IA_CD>0")
    '    ''sb.Append(" And E_INT_ACCOUNT.INACTIVE_FLAG=0 ")
    '    sb.Append(" )")
    '    If _DEBUG Then
    '        sb.Append(" And E_LEVEL.IA_CD Between 849 And 849")
    '    End If
    '    sb.Append(" Order By E_LEVEL.IA_CD, E_LEVEL.IA_CD_N2")
    '    Dim sql As String = sb.ToString()

    '    Dim count As Integer = 0
    '    Dim accounts As IList(Of Account) = New List(Of Account)
    '    Dim dataAccounts As IList(Of DataAccount)

    '    Dim dataTable As New DataTable
    '    Dim sqlError As String = _oracleConnectionDWH.Requete(sql, dataTable)
    '    For Each row As System.Data.DataRow In dataTable.Rows
    '        count = count + 1
    '        If count > 200 Then
    '            Dim err As String = String.Empty
    '            dataAccounts = _queryAccounts(accounts, err)
    '            If Not String.IsNullOrEmpty(err) Then
    '                errors += "<br/>" + err
    '            End If
    '            err = _upsertAccounts(accounts, dataAccounts)
    '            If Not String.IsNullOrEmpty(err) Then
    '                errors += "<br/>" + err
    '            End If

    '            accounts.Clear()
    '            count = 1
    '        End If

    '        Dim ia_cd As String = row.Item("IA_CD").ToString()
    '        Dim ia_cd_n2 As String = ia_cd + "." + row.Item("IA_CD_N2").ToString()
    '        Dim ia_name_n2 As String = row.Item("IA_NAME_N2")
    '        Dim INACTIVE_FLAG As Integer = row.Item("INACTIVE_FLAG")

    '        Dim accountRef As Account = New Account()
    '        accountRef.MIAMI_account_ID__c = ia_cd

    '        Dim account As New Account()
    '        account.Name = ia_name_n2
    '        account.MIAMI_account_ID__c = ia_cd_n2
    '        account.Parent = accountRef
    '        account.Miami_Level__c = "2"
    '        'account.Type = "GCS Account"

    '        If INACTIVE_FLAG > 0 Then
    '            account.Act_Inac_in_MIAMI__c = "Inactive"
    '        Else
    '            account.Act_Inac_in_MIAMI__c = "Active"
    '        End If

    '        accounts.Add(account)

    '        'Console.WriteLine(ia_name_n2)
    '    Next 'read

    '    If count > 0 Then
    '        Dim err As String = String.Empty
    '        dataAccounts = _queryAccounts(accounts, err)
    '        If Not String.IsNullOrEmpty(err) Then
    '            errors += "<br/>" + err
    '        End If
    '        err = _upsertAccounts(accounts, dataAccounts)
    '        If Not String.IsNullOrEmpty(err) Then
    '            errors += "<br/>" + err
    '        End If
    '    End If
    '    Console.WriteLine("--- Level 2 OK")

    '    errors += "<br/>--- Level 2 OK"
    '    Return errors
    'End Function

    'Private Function _level3() As String
    '    Dim errors As String = String.Empty

    '    'Sublevel (Level 3)
    '    Console.WriteLine("--- Level 3")
    '    Dim sb As New StringBuilder()
    '    sb.Append("Select Distinct E_LEVEL.IA_CD, E_LEVEL.IA_CD_N2, E_LEVEL.IA_CD_N3, E_LEVEL.IA_NAME_N3, E_LEVEL.INACTIVE_FLAG")
    '    sb.Append(" From E_LEVEL")
    '    sb.Append(" Where E_LEVEL.IA_CD_N3 > 0")
    '    sb.Append(" And Exists (")
    '    sb.Append(" Select * From E_INT_ACCOUNT")
    '    sb.Append(" Where E_INT_ACCOUNT.IA_CD = E_LEVEL.IA_CD")
    '    sb.Append(" And E_INT_ACCOUNT.IA_CD>0")
    '    ''sb.Append(" And E_INT_ACCOUNT.INACTIVE_FLAG=0 ")
    '    sb.Append(" )")
    '    If _DEBUG Then
    '        sb.Append(" And E_LEVEL.IA_CD Between 849 And 849")
    '    End If
    '    sb.Append(" Order By E_LEVEL.IA_CD, E_LEVEL.IA_CD_N2, E_LEVEL.IA_CD_N3")
    '    Dim sql As String = sb.ToString()

    '    Dim count As Integer = 0
    '    Dim accounts As IList(Of Account) = New List(Of Account)
    '    Dim dataAccounts As IList(Of DataAccount)

    '    Dim dataTable As New DataTable
    '    Dim sqlError As String = _oracleConnectionDWH.Requete(sql, dataTable)
    '    For Each row As System.Data.DataRow In dataTable.Rows
    '        count = count + 1
    '        If count > 200 Then
    '            Dim err As String = String.Empty
    '            dataAccounts = _queryAccounts(accounts, err)
    '            If Not String.IsNullOrEmpty(err) Then
    '                errors += "<br/>" + err
    '            End If
    '            err = _upsertAccounts(accounts, dataAccounts)
    '            If Not String.IsNullOrEmpty(err) Then
    '                errors += "<br/>" + err
    '            End If

    '            accounts.Clear()
    '            count = 1
    '        End If

    '        Dim ia_cd As String = row.Item("IA_CD").ToString()
    '        Dim ia_cd_n2 As String = ia_cd + "." + row.Item("IA_CD_N2").ToString()
    '        'Dim ia_cd_n3 As String = ia_cd_n2 + "." + row.Item("IA_CD_N3").ToString()
    '        Dim ia_cd_n3 As String = ia_cd + "." + row.Item("IA_CD_N3").ToString()
    '        Dim ia_name_n3 As String = row.Item("IA_NAME_N3")
    '        Dim INACTIVE_FLAG As Integer = row.Item("INACTIVE_FLAG")

    '        Dim accountRef As Account = New Account()
    '        accountRef.MIAMI_account_ID__c = ia_cd_n2

    '        Dim account As New Account()
    '        account.Name = ia_name_n3
    '        account.MIAMI_account_ID__c = ia_cd_n3
    '        account.Parent = accountRef
    '        account.Miami_Level__c = "3"
    '        'account.Type = "GCS Account"

    '        If INACTIVE_FLAG > 0 Then
    '            account.Act_Inac_in_MIAMI__c = "Inactive"
    '        Else
    '            account.Act_Inac_in_MIAMI__c = "Active"
    '        End If

    '        accounts.Add(account)

    '        'Console.WriteLine(ia_name_n3)
    '    Next 'read

    '    If count > 0 Then
    '        Dim err As String = String.Empty
    '        dataAccounts = _queryAccounts(accounts, err)
    '        If Not String.IsNullOrEmpty(err) Then
    '            errors += "<br/>" + err
    '        End If
    '        err = _upsertAccounts(accounts, dataAccounts)
    '        If Not String.IsNullOrEmpty(err) Then
    '            errors += "<br/>" + err
    '        End If
    '    End If
    '    Console.WriteLine("--- Level 3 OK")

    '    errors += "<br/>--- Level 3 OK"
    '    Return errors
    'End Function

    'Private Function _level4() As String
    '    Dim errors As String = String.Empty

    '    'Sublevel (Level 4)
    '    Dim sb As New StringBuilder()
    '    Console.WriteLine("--- Level 4")
    '    sb.Append("Select Distinct E_LEVEL.IA_CD, E_LEVEL.IA_CD_N2, E_LEVEL.IA_CD_N3, E_LEVEL.IA_CD_N4, E_LEVEL.IA_NAME_N4, E_LEVEL.INACTIVE_FLAG")
    '    sb.Append(" From E_LEVEL")
    '    sb.Append(" Where E_LEVEL.IA_CD_N4 > 0")
    '    sb.Append(" And Exists (")
    '    sb.Append(" Select * From E_INT_ACCOUNT")
    '    sb.Append(" Where E_INT_ACCOUNT.IA_CD = E_LEVEL.IA_CD")
    '    sb.Append(" And E_INT_ACCOUNT.IA_CD>0")
    '    'sb.Append(" And E_INT_ACCOUNT.INACTIVE_FLAG=0 ")
    '    sb.Append(" )")
    '    If _DEBUG Then
    '        sb.Append(" And E_LEVEL.IA_CD Between 849 And 849")
    '    End If
    '    sb.Append(" Order By E_LEVEL.IA_CD, E_LEVEL.IA_CD_N2, E_LEVEL.IA_CD_N3, E_LEVEL.IA_CD_N4")
    '    Dim sql As String = sb.ToString()

    '    Dim count As Integer = 0
    '    Dim accounts As IList(Of Account) = New List(Of Account)
    '    Dim dataAccounts As IList(Of DataAccount)

    '    Dim dataTable As New DataTable
    '    Dim sqlError As String = _oracleConnectionDWH.Requete(sql, dataTable)
    '    For Each row As System.Data.DataRow In dataTable.Rows
    '        count = count + 1
    '        If count > 200 Then
    '            Dim err As String = String.Empty
    '            dataAccounts = _queryAccounts(accounts, err)
    '            If Not String.IsNullOrEmpty(err) Then
    '                errors += "<br/>" + err
    '            End If
    '            err = _upsertAccounts(accounts, dataAccounts)
    '            If Not String.IsNullOrEmpty(err) Then
    '                errors += "<br/>" + err
    '            End If

    '            accounts.Clear()
    '            count = 1
    '        End If

    '        Dim ia_cd As String = row.Item("IA_CD").ToString()
    '        Dim ia_cd_n2 As String = ia_cd + "." + row.Item("IA_CD_N2").ToString()
    '        'Dim ia_cd_n3 As String = ia_cd_n2 + "." + row.Item("IA_CD_N3").ToString()
    '        'Dim ia_cd_n4 As String = ia_cd_n3 + "." + row.Item("IA_CD_N4").ToString()
    '        Dim ia_cd_n3 As String = ia_cd + "." + row.Item("IA_CD_N3").ToString()
    '        Dim ia_cd_n4 As String = ia_cd + "." + row.Item("IA_CD_N4").ToString()
    '        Dim ia_name_n4 As String = row.Item("IA_NAME_N4")
    '        Dim INACTIVE_FLAG As Integer = row.Item("INACTIVE_FLAG")

    '        Dim accountRef As Account = New Account()
    '        accountRef.MIAMI_account_ID__c = ia_cd_n3

    '        Dim account As New Account()
    '        account.Name = ia_name_n4
    '        account.MIAMI_account_ID__c = ia_cd_n4
    '        account.Parent = accountRef
    '        account.Miami_Level__c = "4"
    '        'account.Type = "GCS Account"

    '        If INACTIVE_FLAG > 0 Then
    '            account.Act_Inac_in_MIAMI__c = "Inactive"
    '        Else
    '            account.Act_Inac_in_MIAMI__c = "Active"
    '        End If

    '        accounts.Add(account)

    '        'Console.WriteLine(ia_name_n4)
    '    Next 'read

    '    If count > 0 Then
    '        Dim err As String = String.Empty
    '        dataAccounts = _queryAccounts(accounts, err)
    '        If Not String.IsNullOrEmpty(err) Then
    '            errors += "<br/>" + err
    '        End If
    '        err = _upsertAccounts(accounts, dataAccounts)
    '        If Not String.IsNullOrEmpty(err) Then
    '            errors += "<br/>" + err
    '        End If
    '    End If
    '    Console.WriteLine("--- Level 4 OK")

    '    errors += "<br/>--- Level 4 OK"
    '    Return errors
    'End Function

    'Private Function _actualsLevel1() As String
    '    Dim errors As String = String.Empty

    '    'Sales and Margin per Account and per Service type
    '    Console.WriteLine("--- Actuals Level 1")
    '    Dim sb As New StringBuilder()
    '    sb.Append("Select H_LINK.IA_CD, E_DIVISION.SERVICE_CD, A_SALES.YEAR, sum(Nvl(A_SALES.TOTAL_SALES,0)*H_CURRENCY.CHANGE_RATE) as TOTAL_SALES,")
    '    sb.Append(" sum((Nvl(A_SALES.TOTAL_SALES,0)-Nvl(A_SALES.DIRECT_COST,0))*H_CURRENCY.CHANGE_RATE) as MARGIN")
    '    sb.Append(" From A_SALES")
    '    sb.Append(" Inner Join E_DIVISION On A_SALES.COUNTRY_CD=E_DIVISION.COUNTRY_CD")
    '    sb.Append(" And A_SALES.SENDER_ID=E_DIVISION.SENDER_ID")
    '    sb.Append(" And A_SALES.DIVISION_CD=E_DIVISION.DIVISION_CD")
    '    sb.Append(" Inner Join H_LINK On A_SALES.COUNTRY_CD=E_DIVISION.COUNTRY_CD")
    '    sb.Append(" And A_SALES.SENDER_ID=H_LINK.SENDER_ID")
    '    sb.Append(" And A_SALES.DIVISION_CD=H_LINK.DIVISION_CD")
    '    sb.Append(" And A_SALES.CLIENT_CD=H_LINK.CLIENT_CD")
    '    sb.Append(" Inner Join H_CURRENCY On A_SALES.YEAR=H_CURRENCY.YEAR")
    '    sb.Append(" And A_SALES.MONTH=H_CURRENCY.MONTH")
    '    sb.Append(" And A_SALES.COUNTRY_CD=H_CURRENCY.COUNTRY_CD")
    '    sb.Append(" Inner Join E_INT_ACCOUNT On H_LINK.IA_CD=E_INT_ACCOUNT.IA_CD")
    '    sb.Append(" Where A_SALES.YEAR||'-'||A_SALES.MONTH Between '")
    '    sb.Append(_fromYearMonth)
    '    sb.Append("' And '")
    '    sb.Append(_toYearMonth)
    '    sb.Append("' And H_LINK.IA_CD>0")
    '    sb.Append(" And E_INT_ACCOUNT.INACTIVE_FLAG=0")
    '    If _DEBUG Then
    '        sb.Append(" And H_LINK.IA_CD Between 849 And 849")
    '    End If
    '    sb.Append(" Group By H_LINK.IA_CD, E_DIVISION.SERVICE_CD, A_SALES.YEAR")
    '    Dim sql As String = sb.ToString()

    '    Dim count As Integer = 0
    '    Dim actuals As IList(Of Account_Actual_Miami__c) = New List(Of Account_Actual_Miami__c)

    '    Dim dataTable As New DataTable
    '    Dim sqlError As String = _oracleConnectionDWH.Requete(sql, dataTable)
    '    For Each row As System.Data.DataRow In dataTable.Rows
    '        count = count + 1
    '        If count > 200 Then

    '            Dim err As String = String.Empty
    '            err = _upsertActuals(actuals)
    '            If Not String.IsNullOrEmpty(err) Then
    '                errors += "<br/>" + err
    '            End If

    '            actuals.Clear()
    '            count = 1
    '        End If

    '        Dim ia_cd As String = row.Item("IA_CD")
    '        Dim year As String = row.Item("YEAR")
    '        Dim servicetype As String = row.Item("SERVICE_CD")
    '        Dim margin As Double = 0.0
    '        If Not IsDBNull(row.Item("MARGIN")) Then
    '            margin = row.Item("MARGIN")
    '        End If
    '        Dim total_sales As Double = 0.0
    '        If Not IsDBNull(row.Item("TOTAL_SALES")) Then
    '            total_sales = row.Item("TOTAL_SALES")
    '        End If

    '        If total_sales <> 0.0 AndAlso margin <> 0.0 Then
    '            Dim accountRef As Account = New Account()
    '            accountRef.MIAMI_account_ID__c = ia_cd

    '            Dim actual As New Account_Actual_Miami__c()
    '            actual.ExternalID__c = ia_cd + "-" + year + "-" + servicetype
    '            actual.Account_Name__r = accountRef
    '            actual.Year__c = year
    '            Dim err As String = _getSFPickListValue(_SF_ACC_SRV_TYPE, "_SF_ACC_SRV_TYPE", servicetype, actual.Service_Type__c)
    '            If Not String.IsNullOrEmpty(err) Then
    '                errors += "<br/>" + err
    '            End If
    '            actual.Actual_Sales__c = Math.Round(total_sales, 0)
    '            actual.Actual_Sales__cSpecified = True
    '            actual.Margin_Amount__c = Math.Round(margin, 0)
    '            actual.Margin_Amount__cSpecified = True
    '            actuals.Add(actual)
    '        End If
    '    Next
    '    If count > 0 Then
    '        Dim err As String = String.Empty
    '        err = _upsertActuals(actuals)
    '        If Not String.IsNullOrEmpty(err) Then
    '            errors += "<br/>" + err
    '        End If
    '    End If
    '    Console.WriteLine("--- Actuals Level 1 OK")

    '    errors += "<br/>--- Actuals Level 1 OK"
    '    Return errors
    'End Function

    'Private Function _actualsLevel2() As String
    '    'Sales and Margin per Level 2 and Service type

    '    Dim errors As String = String.Empty

    '    Console.WriteLine("--- Actuals Level 2")
    '    Dim sb As New StringBuilder()
    '    sb.Append(" Select H_LINK.IA_CD, H_LINK.IA_CD_N2, E_DIVISION.SERVICE_CD, A_SALES.YEAR,")
    '    sb.Append(" sum(Nvl(A_SALES.TOTAL_SALES,0)*H_CURRENCY.CHANGE_RATE) as TOTAL_SALES,")
    '    sb.Append(" sum((Nvl(A_SALES.TOTAL_SALES,0)-Nvl(A_SALES.DIRECT_COST,0))*H_CURRENCY.CHANGE_RATE) as MARGIN")
    '    sb.Append(" From A_SALES")
    '    sb.Append(" Inner Join E_DIVISION On A_SALES.COUNTRY_CD=E_DIVISION.COUNTRY_CD")
    '    sb.Append(" And A_SALES.SENDER_ID=E_DIVISION.SENDER_ID")
    '    sb.Append(" And A_SALES.DIVISION_CD=E_DIVISION.DIVISION_CD")
    '    sb.Append(" Inner Join H_LINK On A_SALES.COUNTRY_CD=E_DIVISION.COUNTRY_CD")
    '    sb.Append(" And A_SALES.SENDER_ID=H_LINK.SENDER_ID")
    '    sb.Append(" And A_SALES.DIVISION_CD=H_LINK.DIVISION_CD")
    '    sb.Append(" And A_SALES.CLIENT_CD=H_LINK.CLIENT_CD")
    '    sb.Append(" Inner Join H_CURRENCY On A_SALES.YEAR=H_CURRENCY.YEAR")
    '    sb.Append(" And A_SALES.MONTH=H_CURRENCY.MONTH")
    '    sb.Append(" And A_SALES.COUNTRY_CD=H_CURRENCY.COUNTRY_CD")
    '    sb.Append(" Inner Join E_LEVEL On E_LEVEL.IA_CD = H_LINK.IA_CD")
    '    sb.Append(" And E_LEVEL.IA_CD_N2 = H_LINK.IA_CD_N2")
    '    sb.Append(" And E_LEVEL.IA_CD_N3 = H_LINK.IA_CD_N3")
    '    sb.Append(" And E_LEVEL.IA_CD_N4 = H_LINK.IA_CD_N4")
    '    sb.Append(" Inner Join E_INT_ACCOUNT On H_LINK.IA_CD=E_INT_ACCOUNT.IA_CD")
    '    sb.Append(" Where A_SALES.YEAR||'-'||A_SALES.MONTH Between '")
    '    sb.Append(_fromYearMonth)
    '    sb.Append("' And '")
    '    sb.Append(_toYearMonth)
    '    sb.Append("' And H_LINK.IA_CD>0")
    '    sb.Append(" And H_LINK.IA_CD_N2>0")
    '    sb.Append(" And H_LINK.IA_CD_N3=0")
    '    sb.Append(" And H_LINK.IA_CD_N4=0")
    '    sb.Append(" And E_INT_ACCOUNT.INACTIVE_FLAG=0")
    '    If _DEBUG Then
    '        sb.Append(" And H_LINK.IA_CD Between 849 And 849")
    '        'sb.Append(" And E_DIVISION.SERVICE_CD='SRC'")
    '    End If
    '    sb.Append(" Group By H_LINK.IA_CD, H_LINK.IA_CD_N2, E_DIVISION.SERVICE_CD, A_SALES.YEAR")
    '    sb.Append(" Order By H_LINK.IA_CD, H_LINK.IA_CD_N2, E_DIVISION.SERVICE_CD, A_SALES.YEAR")
    '    Dim sql As String = sb.ToString()

    '    Dim count As Integer = 0
    '    Dim actuals As IList(Of Account_Actual_Miami__c) = New List(Of Account_Actual_Miami__c)

    '    Dim dataTable As New DataTable
    '    Dim sqlError As String = _oracleConnectionDWH.Requete(sql, dataTable)
    '    For Each row As System.Data.DataRow In dataTable.Rows
    '        count = count + 1
    '        If count > 200 Then
    '            Dim err As String = String.Empty
    '            err = _upsertActuals(actuals)
    '            If Not String.IsNullOrEmpty(err) Then
    '                errors += "<br/>" + err
    '            End If

    '            actuals.Clear()
    '            count = 1
    '        End If

    '        Dim ia_cd As String = row.Item("IA_CD")
    '        Dim ia_cd_n2 As String = ia_cd + "." + row.Item("IA_CD_N2").ToString()
    '        Dim year As String = row.Item("YEAR")
    '        Dim servicetype As String = row.Item("SERVICE_CD")
    '        Dim margin As Double = 0.0
    '        If Not IsDBNull(row.Item("MARGIN")) Then
    '            margin = row.Item("MARGIN")
    '        End If
    '        Dim total_sales As Double = 0.0
    '        If Not IsDBNull(row.Item("TOTAL_SALES")) Then
    '            total_sales = row.Item("TOTAL_SALES")
    '        End If

    '        If total_sales <> 0.0 AndAlso margin <> 0.0 Then
    '            Dim accountRef As Account = New Account()
    '            accountRef.MIAMI_account_ID__c = ia_cd_n2

    '            Dim actual As New Account_Actual_Miami__c()
    '            actual.ExternalID__c = ia_cd_n2 + "-" + year + "-" + servicetype
    '            actual.Account_Name__r = accountRef
    '            actual.Year__c = year
    '            Dim err As String = _getSFPickListValue(_SF_ACC_SRV_TYPE, "_SF_ACC_SRV_TYPE", servicetype, actual.Service_Type__c)
    '            If Not String.IsNullOrEmpty(err) Then
    '                errors += "<br/>" + err
    '            End If
    '            actual.Actual_Sales__c = Math.Round(total_sales, 0)
    '            actual.Actual_Sales__cSpecified = True
    '            actual.Margin_Amount__c = Math.Round(margin, 0)
    '            actual.Margin_Amount__cSpecified = True
    '            actuals.Add(actual)
    '        End If
    '    Next
    '    If count > 0 Then
    '        Dim err As String = String.Empty
    '        err = _upsertActuals(actuals)
    '        If Not String.IsNullOrEmpty(err) Then
    '            errors += "<br/>" + err
    '        End If
    '    End If
    '    Console.WriteLine("--- Actuals Level 2 OK")

    '    errors += "<br/>--- Actuals Level 2 OK"
    '    Return errors
    'End Function

    'Private Function _actualsLevel3() As String
    '    'Sales and Margin per Level 3 and Service type

    '    Dim errors As String = String.Empty

    '    Console.WriteLine("--- Actuals Level 3")
    '    Dim sb As New StringBuilder()
    '    sb.Append("Select H_LINK.IA_CD, H_LINK.IA_CD_N2, H_LINK.IA_CD_N3, E_DIVISION.SERVICE_CD, A_SALES.YEAR,")
    '    sb.Append(" sum(Nvl(A_SALES.TOTAL_SALES,0)*H_CURRENCY.CHANGE_RATE) as TOTAL_SALES,")
    '    sb.Append(" sum((Nvl(A_SALES.TOTAL_SALES,0)-Nvl(A_SALES.DIRECT_COST,0))*H_CURRENCY.CHANGE_RATE) as MARGIN")
    '    sb.Append(" From A_SALES")
    '    sb.Append(" Inner Join E_DIVISION On A_SALES.COUNTRY_CD=E_DIVISION.COUNTRY_CD")
    '    sb.Append(" And A_SALES.SENDER_ID=E_DIVISION.SENDER_ID")
    '    sb.Append(" And A_SALES.DIVISION_CD=E_DIVISION.DIVISION_CD")
    '    sb.Append(" Inner Join H_LINK On A_SALES.COUNTRY_CD=E_DIVISION.COUNTRY_CD")
    '    sb.Append(" And A_SALES.SENDER_ID=H_LINK.SENDER_ID")
    '    sb.Append(" And A_SALES.DIVISION_CD=H_LINK.DIVISION_CD")
    '    sb.Append(" And A_SALES.CLIENT_CD=H_LINK.CLIENT_CD")
    '    sb.Append(" Inner Join H_CURRENCY On A_SALES.YEAR=H_CURRENCY.YEAR")
    '    sb.Append(" And A_SALES.MONTH=H_CURRENCY.MONTH")
    '    sb.Append(" And A_SALES.COUNTRY_CD=H_CURRENCY.COUNTRY_CD")
    '    sb.Append(" Inner Join E_LEVEL On E_LEVEL.IA_CD = H_LINK.IA_CD")
    '    sb.Append(" And E_LEVEL.IA_CD_N2 = H_LINK.IA_CD_N2")
    '    sb.Append(" And E_LEVEL.IA_CD_N3 = H_LINK.IA_CD_N3")
    '    sb.Append(" And E_LEVEL.IA_CD_N4 = H_LINK.IA_CD_N4")
    '    sb.Append(" Inner Join E_INT_ACCOUNT On H_LINK.IA_CD=E_INT_ACCOUNT.IA_CD")
    '    sb.Append(" Where A_SALES.YEAR||'-'||A_SALES.MONTH Between '")
    '    sb.Append(_fromYearMonth)
    '    sb.Append("' And '")
    '    sb.Append(_toYearMonth)
    '    sb.Append("' And H_LINK.IA_CD>0")
    '    sb.Append(" And H_LINK.IA_CD_N2>0")
    '    sb.Append(" And H_LINK.IA_CD_N3>0")
    '    sb.Append(" And H_LINK.IA_CD_N4=0")
    '    sb.Append(" And E_INT_ACCOUNT.INACTIVE_FLAG=0")
    '    If _DEBUG Then
    '        sb.Append(" And H_LINK.IA_CD Between 849 And 849")
    '        'sb.Append(" And E_DIVISION.SERVICE_CD='SRC'")
    '    End If
    '    sb.Append(" Group By H_LINK.IA_CD, H_LINK.IA_CD_N2, H_LINK.IA_CD_N3, E_DIVISION.SERVICE_CD, A_SALES.YEAR")
    '    sb.Append(" Order By H_LINK.IA_CD, H_LINK.IA_CD_N2, H_LINK.IA_CD_N3, E_DIVISION.SERVICE_CD, A_SALES.YEAR")
    '    Dim sql As String = sb.ToString()

    '    Dim count As Integer = 0
    '    Dim actuals As IList(Of Account_Actual_Miami__c) = New List(Of Account_Actual_Miami__c)

    '    Dim dataTable As New DataTable
    '    Dim sqlError As String = _oracleConnectionDWH.Requete(sql, dataTable)
    '    For Each row As System.Data.DataRow In dataTable.Rows
    '        count = count + 1
    '        If count > 200 Then
    '            Dim err As String = String.Empty
    '            err = _upsertActuals(actuals)
    '            If Not String.IsNullOrEmpty(err) Then
    '                errors += "<br/>" + err
    '            End If

    '            actuals.Clear()
    '            count = 1
    '        End If

    '        Dim ia_cd As String = row.Item("IA_CD")
    '        Dim ia_cd_n2 As String = ia_cd + "." + row.Item("IA_CD_N2").ToString()
    '        'Dim ia_cd_n3 As String = ia_cd_n2 + "." + row.Item("IA_CD_N3").ToString()
    '        Dim ia_cd_n3 As String = ia_cd + "." + row.Item("IA_CD_N3").ToString()
    '        Dim year As String = row.Item("YEAR")
    '        Dim servicetype As String = row.Item("SERVICE_CD")
    '        Dim margin As Double = 0.0
    '        If Not IsDBNull(row.Item("MARGIN")) Then
    '            margin = row.Item("MARGIN")
    '        End If
    '        Dim total_sales As Double = 0.0
    '        If Not IsDBNull(row.Item("TOTAL_SALES")) Then
    '            total_sales = row.Item("TOTAL_SALES")
    '        End If

    '        If total_sales <> 0.0 AndAlso margin <> 0.0 Then
    '            Dim accountRef As Account = New Account()
    '            accountRef.MIAMI_account_ID__c = ia_cd_n3

    '            Dim actual As New Account_Actual_Miami__c()
    '            actual.ExternalID__c = ia_cd_n3 + "-" + year + "-" + servicetype
    '            actual.Account_Name__r = accountRef
    '            actual.Year__c = year
    '            Dim err As String = _getSFPickListValue(_SF_ACC_SRV_TYPE, "_SF_ACC_SRV_TYPE", servicetype, actual.Service_Type__c)
    '            If Not String.IsNullOrEmpty(err) Then
    '                errors += "<br/>" + err
    '            End If
    '            actual.Actual_Sales__c = Math.Round(total_sales, 0)
    '            actual.Actual_Sales__cSpecified = True
    '            actual.Margin_Amount__c = Math.Round(margin, 0)
    '            actual.Margin_Amount__cSpecified = True
    '            actuals.Add(actual)
    '        End If
    '    Next
    '    If count > 0 Then
    '        Dim err As String = String.Empty
    '        err = _upsertActuals(actuals)
    '        If Not String.IsNullOrEmpty(err) Then
    '            errors += "<br/>" + err
    '        End If
    '    End If
    '    Console.WriteLine("--- Actuals Level 3 OK")

    '    errors += "<br/>--- Actuals Level 3 OK"
    '    Return errors
    'End Function

    'Private Function _actualsLevel4() As String
    '    'Sales and Margin per Level 4 and Service type

    '    Dim errors As String = String.Empty

    '    Console.WriteLine("--- Actuals Level 4")
    '    Dim sb As New StringBuilder()
    '    sb.Append("Select H_LINK.IA_CD, H_LINK.IA_CD_N2, H_LINK.IA_CD_N3, H_LINK.IA_CD_N4, E_DIVISION.SERVICE_CD, A_SALES.YEAR,")
    '    sb.Append(" sum(Nvl(A_SALES.TOTAL_SALES,0)*H_CURRENCY.CHANGE_RATE) as TOTAL_SALES,")
    '    sb.Append(" sum((Nvl(A_SALES.TOTAL_SALES,0)-Nvl(A_SALES.DIRECT_COST,0))*H_CURRENCY.CHANGE_RATE) as MARGIN")
    '    sb.Append(" From A_SALES")
    '    sb.Append(" Inner Join E_DIVISION On A_SALES.COUNTRY_CD=E_DIVISION.COUNTRY_CD")
    '    sb.Append(" And A_SALES.SENDER_ID=E_DIVISION.SENDER_ID")
    '    sb.Append(" And A_SALES.DIVISION_CD=E_DIVISION.DIVISION_CD")
    '    sb.Append(" Inner Join H_LINK On A_SALES.COUNTRY_CD=E_DIVISION.COUNTRY_CD")
    '    sb.Append(" And A_SALES.SENDER_ID=H_LINK.SENDER_ID")
    '    sb.Append(" And A_SALES.DIVISION_CD=H_LINK.DIVISION_CD")
    '    sb.Append(" And A_SALES.CLIENT_CD=H_LINK.CLIENT_CD")
    '    sb.Append(" Inner Join H_CURRENCY On A_SALES.YEAR=H_CURRENCY.YEAR")
    '    sb.Append(" And A_SALES.MONTH=H_CURRENCY.MONTH")
    '    sb.Append(" And A_SALES.COUNTRY_CD=H_CURRENCY.COUNTRY_CD")
    '    sb.Append(" Inner Join E_LEVEL On E_LEVEL.IA_CD = H_LINK.IA_CD")
    '    sb.Append(" And E_LEVEL.IA_CD_N2 = H_LINK.IA_CD_N2")
    '    sb.Append(" And E_LEVEL.IA_CD_N3 = H_LINK.IA_CD_N3")
    '    sb.Append(" And E_LEVEL.IA_CD_N4 = H_LINK.IA_CD_N4")
    '    sb.Append(" Inner Join E_INT_ACCOUNT On H_LINK.IA_CD=E_INT_ACCOUNT.IA_CD")
    '    sb.Append(" Where A_SALES.YEAR||'-'||A_SALES.MONTH Between '")
    '    sb.Append(_fromYearMonth)
    '    sb.Append("' And '")
    '    sb.Append(_toYearMonth)
    '    sb.Append("' And H_LINK.IA_CD>0")
    '    sb.Append(" And H_LINK.IA_CD_N2>0")
    '    sb.Append(" And H_LINK.IA_CD_N3>0")
    '    sb.Append(" And H_LINK.IA_CD_N4>0")
    '    sb.Append(" And E_INT_ACCOUNT.INACTIVE_FLAG=0")
    '    If _DEBUG Then
    '        sb.Append(" And H_LINK.IA_CD Between 849 And 849")
    '        'sb.Append(" And E_DIVISION.SERVICE_CD='SRC'")
    '    End If
    '    sb.Append(" Group By H_LINK.IA_CD, H_LINK.IA_CD_N2, H_LINK.IA_CD_N3, H_LINK.IA_CD_N4, E_DIVISION.SERVICE_CD, A_SALES.YEAR")
    '    sb.Append(" Order By H_LINK.IA_CD, H_LINK.IA_CD_N2, H_LINK.IA_CD_N3, H_LINK.IA_CD_N4, E_DIVISION.SERVICE_CD, A_SALES.YEAR")
    '    Dim sql As String = sb.ToString()

    '    Dim count As Integer = 0
    '    Dim actuals As IList(Of Account_Actual_Miami__c) = New List(Of Account_Actual_Miami__c)

    '    Dim dataTable As New DataTable
    '    Dim sqlError As String = _oracleConnectionDWH.Requete(sql, dataTable)
    '    For Each row As System.Data.DataRow In dataTable.Rows
    '        count = count + 1
    '        If count > 200 Then
    '            Dim err As String = String.Empty
    '            err = _upsertActuals(actuals)
    '            If Not String.IsNullOrEmpty(err) Then
    '                errors += "<br/>" + err
    '            End If

    '            actuals.Clear()
    '            count = 1
    '        End If

    '        Dim ia_cd As String = row.Item("IA_CD")
    '        Dim ia_cd_n2 As String = ia_cd + "." + row.Item("IA_CD_N2").ToString()
    '        'Dim ia_cd_n3 As String = ia_cd_n2 + "." + row.Item("IA_CD_N3").ToString()
    '        'Dim ia_cd_n4 As String = ia_cd_n3 + "." + row.Item("IA_CD_N4").ToString()
    '        Dim ia_cd_n3 As String = ia_cd + "." + row.Item("IA_CD_N3").ToString()
    '        Dim ia_cd_n4 As String = ia_cd + "." + row.Item("IA_CD_N4").ToString()
    '        Dim year As String = row.Item("YEAR")
    '        Dim servicetype As String = row.Item("SERVICE_CD")
    '        Dim margin As Double = 0.0
    '        If Not IsDBNull(row.Item("MARGIN")) Then
    '            margin = row.Item("MARGIN")
    '        End If
    '        Dim total_sales As Double = 0.0
    '        If Not IsDBNull(row.Item("TOTAL_SALES")) Then
    '            total_sales = row.Item("TOTAL_SALES")
    '        End If

    '        If total_sales <> 0.0 AndAlso margin <> 0.0 Then
    '            Dim accountRef As Account = New Account()
    '            accountRef.MIAMI_account_ID__c = ia_cd_n4

    '            Dim actual As New Account_Actual_Miami__c()
    '            actual.ExternalID__c = ia_cd_n4 + "-" + year + "-" + servicetype
    '            actual.Account_Name__r = accountRef
    '            actual.Year__c = year
    '            Dim err As String = _getSFPickListValue(_SF_ACC_SRV_TYPE, "_SF_ACC_SRV_TYPE", servicetype, actual.Service_Type__c)
    '            If Not String.IsNullOrEmpty(err) Then
    '                errors += "<br/>" + err
    '            End If
    '            actual.Actual_Sales__c = Math.Round(total_sales, 0)
    '            actual.Actual_Sales__cSpecified = True
    '            actual.Margin_Amount__c = Math.Round(margin, 0)
    '            actual.Margin_Amount__cSpecified = True
    '            actuals.Add(actual)
    '        End If
    '    Next
    '    If count > 0 Then
    '        Dim err As String = String.Empty
    '        err = _upsertActuals(actuals)
    '        If Not String.IsNullOrEmpty(err) Then
    '            errors += "<br/>" + err
    '        End If
    '    End If
    '    Console.WriteLine("--- Actuals Level 4 OK")


    '    errors += "<br/>--- Actuals Level 4 OK"
    '    Return errors
    'End Function

    'Private Function _actualsLevel3SumLevel4() As String
    '    'Sales and Margin per Level 3 using Level 4 grouped by Level 3 and Service type

    '    Dim errors As String = String.Empty

    '    Console.WriteLine("--- Actuals Level 3 Sum L4")
    '    Dim sb As New StringBuilder()
    '    sb.Append(" Select H_LINK.IA_CD, H_LINK.IA_CD_N2, H_LINK.IA_CD_N3, E_DIVISION.SERVICE_CD, A_SALES.YEAR,")
    '    sb.Append(" sum(Nvl(A_SALES.TOTAL_SALES,0)*H_CURRENCY.CHANGE_RATE) as TOTAL_SALES,")
    '    sb.Append(" sum((Nvl(A_SALES.TOTAL_SALES,0)-Nvl(A_SALES.DIRECT_COST,0))*H_CURRENCY.CHANGE_RATE) as MARGIN")
    '    sb.Append(" From A_SALES")
    '    sb.Append(" Inner Join E_DIVISION On A_SALES.COUNTRY_CD=E_DIVISION.COUNTRY_CD")
    '    sb.Append(" And A_SALES.SENDER_ID=E_DIVISION.SENDER_ID")
    '    sb.Append(" And A_SALES.DIVISION_CD=E_DIVISION.DIVISION_CD")
    '    sb.Append(" Inner Join H_LINK On A_SALES.COUNTRY_CD=E_DIVISION.COUNTRY_CD")
    '    sb.Append(" And A_SALES.SENDER_ID=H_LINK.SENDER_ID")
    '    sb.Append(" And A_SALES.DIVISION_CD=H_LINK.DIVISION_CD")
    '    sb.Append(" And A_SALES.CLIENT_CD=H_LINK.CLIENT_CD")
    '    sb.Append(" Inner Join H_CURRENCY On A_SALES.YEAR=H_CURRENCY.YEAR")
    '    sb.Append(" And A_SALES.MONTH=H_CURRENCY.MONTH")
    '    sb.Append(" And A_SALES.COUNTRY_CD=H_CURRENCY.COUNTRY_CD")
    '    sb.Append(" Inner Join E_LEVEL On E_LEVEL.IA_CD = H_LINK.IA_CD")
    '    sb.Append(" And E_LEVEL.IA_CD_N2 = H_LINK.IA_CD_N2")
    '    sb.Append(" And E_LEVEL.IA_CD_N3 = H_LINK.IA_CD_N3")
    '    sb.Append(" And E_LEVEL.IA_CD_N4 = H_LINK.IA_CD_N4")
    '    sb.Append(" Inner Join E_INT_ACCOUNT On H_LINK.IA_CD=E_INT_ACCOUNT.IA_CD")
    '    sb.Append(" Where A_SALES.YEAR||'-'||A_SALES.MONTH Between '")
    '    sb.Append(_fromYearMonth)
    '    sb.Append("' And '")
    '    sb.Append(_toYearMonth)
    '    sb.Append("' And H_LINK.IA_CD>0")
    '    sb.Append(" And H_LINK.IA_CD_N2>0")
    '    sb.Append(" And H_LINK.IA_CD_N3>0")
    '    sb.Append(" And H_LINK.IA_CD_N4>0")
    '    sb.Append(" And E_INT_ACCOUNT.INACTIVE_FLAG=0")
    '    If _DEBUG Then
    '        sb.Append(" And H_LINK.IA_CD Between 849 And 849")
    '        'sb.Append(" And E_DIVISION.SERVICE_CD='SRC'")
    '    End If
    '    sb.Append(" Group By H_LINK.IA_CD, H_LINK.IA_CD_N2, H_LINK.IA_CD_N3, E_DIVISION.SERVICE_CD, A_SALES.YEAR")
    '    sb.Append(" Order By H_LINK.IA_CD, H_LINK.IA_CD_N2, H_LINK.IA_CD_N3, E_DIVISION.SERVICE_CD, A_SALES.YEAR")
    '    Dim sql As String = sb.ToString()

    '    Dim count As Integer = 0
    '    Dim actuals As IList(Of Account_Actual_Miami__c) = New List(Of Account_Actual_Miami__c)

    '    Dim dataTable As New DataTable
    '    Dim sqlError As String = _oracleConnectionDWH.Requete(sql, dataTable)
    '    For Each row As System.Data.DataRow In dataTable.Rows
    '        count = count + 1
    '        If count > 200 Then
    '            Dim err As String = String.Empty
    '            err = _upsertActuals(actuals)
    '            If Not String.IsNullOrEmpty(err) Then
    '                errors += "<br/>" + err
    '            End If

    '            actuals.Clear()
    '            count = 1
    '        End If

    '        Dim ia_cd As String = row.Item("IA_CD")
    '        Dim ia_cd_n2 As String = ia_cd + "." + row.Item("IA_CD_N2").ToString()
    '        'Dim ia_cd_n3 As String = ia_cd_n2 + "." + row.Item("IA_CD_N3").ToString()
    '        Dim ia_cd_n3 As String = ia_cd + "." + row.Item("IA_CD_N3").ToString()
    '        Dim year As String = row.Item("YEAR")
    '        Dim servicetype As String = row.Item("SERVICE_CD")
    '        Dim margin As Double = 0.0
    '        If Not IsDBNull(row.Item("MARGIN")) Then
    '            margin = row.Item("MARGIN")
    '        End If
    '        Dim total_sales As Double = 0.0
    '        If Not IsDBNull(row.Item("TOTAL_SALES")) Then
    '            total_sales = row.Item("TOTAL_SALES")
    '        End If

    '        If total_sales <> 0.0 AndAlso margin <> 0.0 Then
    '            Dim accountRef As Account = New Account()
    '            accountRef.MIAMI_account_ID__c = ia_cd_n3

    '            Dim actual As New Account_Actual_Miami__c()
    '            actual.ExternalID__c = ia_cd_n3 + "-" + year + "-" + servicetype
    '            actual.Account_Name__r = accountRef
    '            actual.Year__c = year
    '            Dim err As String = _getSFPickListValue(_SF_ACC_SRV_TYPE, "_SF_ACC_SRV_TYPE", servicetype, actual.Service_Type__c)
    '            If Not String.IsNullOrEmpty(err) Then
    '                errors += "<br/>" + err
    '            End If
    '            actual.Actual_Sales__c = Math.Round(total_sales, 0)
    '            actual.Actual_Sales__cSpecified = True
    '            actual.Margin_Amount__c = Math.Round(margin, 0)
    '            actual.Margin_Amount__cSpecified = True
    '            actuals.Add(actual)
    '        End If
    '    Next
    '    If count > 0 Then
    '        Dim err As String = String.Empty
    '        err = _upsertActuals(actuals)
    '        If Not String.IsNullOrEmpty(err) Then
    '            errors += "<br/>" + err
    '        End If
    '    End If
    '    Console.WriteLine("--- Actuals Level 3 Sum L4 OK")

    '    errors += "<br/>--- Actuals Level 3 Sum L4 OK"
    '    Return errors
    'End Function

    'Private Function _actualsLevel2SumLevel3() As String
    '    'Sales and Margin per Level 2 using Level 3 grouped by Level 2 and Service type

    '    Dim errors As String = String.Empty

    '    Console.WriteLine("--- Actuals Level 3 Sum L3")
    '    Dim sb As New StringBuilder()
    '    sb.Append("Select H_LINK.IA_CD, H_LINK.IA_CD_N2, E_DIVISION.SERVICE_CD, A_SALES.YEAR,")
    '    sb.Append(" sum(Nvl(A_SALES.TOTAL_SALES,0)*H_CURRENCY.CHANGE_RATE) as TOTAL_SALES,")
    '    sb.Append(" sum((Nvl(A_SALES.TOTAL_SALES,0)-Nvl(A_SALES.DIRECT_COST,0))*H_CURRENCY.CHANGE_RATE) as MARGIN")
    '    sb.Append(" From A_SALES")
    '    sb.Append(" Inner Join E_DIVISION On A_SALES.COUNTRY_CD=E_DIVISION.COUNTRY_CD")
    '    sb.Append(" And A_SALES.SENDER_ID=E_DIVISION.SENDER_ID")
    '    sb.Append(" And A_SALES.DIVISION_CD=E_DIVISION.DIVISION_CD")
    '    sb.Append(" Inner Join H_LINK On A_SALES.COUNTRY_CD=E_DIVISION.COUNTRY_CD")
    '    sb.Append(" And A_SALES.SENDER_ID=H_LINK.SENDER_ID")
    '    sb.Append(" And A_SALES.DIVISION_CD=H_LINK.DIVISION_CD")
    '    sb.Append(" And A_SALES.CLIENT_CD=H_LINK.CLIENT_CD")
    '    sb.Append(" Inner Join H_CURRENCY On A_SALES.YEAR=H_CURRENCY.YEAR")
    '    sb.Append(" And A_SALES.MONTH=H_CURRENCY.MONTH")
    '    sb.Append(" And A_SALES.COUNTRY_CD=H_CURRENCY.COUNTRY_CD")
    '    sb.Append(" Inner Join E_LEVEL On E_LEVEL.IA_CD = H_LINK.IA_CD")
    '    sb.Append(" And E_LEVEL.IA_CD_N2 = H_LINK.IA_CD_N2")
    '    sb.Append(" And E_LEVEL.IA_CD_N3 = H_LINK.IA_CD_N3")
    '    sb.Append(" And E_LEVEL.IA_CD_N4 = H_LINK.IA_CD_N4")
    '    sb.Append(" Inner Join E_INT_ACCOUNT On H_LINK.IA_CD=E_INT_ACCOUNT.IA_CD")
    '    sb.Append(" Where A_SALES.YEAR||'-'||A_SALES.MONTH Between '")
    '    sb.Append(_fromYearMonth)
    '    sb.Append("' And '")
    '    sb.Append(_toYearMonth)
    '    sb.Append("' And H_LINK.IA_CD>0")
    '    sb.Append(" And H_LINK.IA_CD_N2>0")
    '    sb.Append(" And H_LINK.IA_CD_N3>0")
    '    sb.Append(" And H_LINK.IA_CD_N4=0")
    '    sb.Append(" And E_INT_ACCOUNT.INACTIVE_FLAG=0")
    '    If _DEBUG Then
    '        sb.Append(" And H_LINK.IA_CD Between 849 And 849")
    '        'sb.Append(" And E_DIVISION.SERVICE_CD='SRC'")
    '    End If
    '    sb.Append(" Group By H_LINK.IA_CD, H_LINK.IA_CD_N2, E_DIVISION.SERVICE_CD, A_SALES.YEAR")
    '    sb.Append(" Order By H_LINK.IA_CD, H_LINK.IA_CD_N2, E_DIVISION.SERVICE_CD, A_SALES.YEAR")
    '    Dim sql As String = sb.ToString()

    '    Dim count As Integer = 0
    '    Dim actuals As IList(Of Account_Actual_Miami__c) = New List(Of Account_Actual_Miami__c)

    '    Dim dataTable As New DataTable
    '    Dim sqlError As String = _oracleConnectionDWH.Requete(sql, dataTable)
    '    For Each row As System.Data.DataRow In dataTable.Rows
    '        count = count + 1
    '        If count > 200 Then
    '            Dim err As String = String.Empty
    '            err = _upsertActuals(actuals)
    '            If Not String.IsNullOrEmpty(err) Then
    '                errors += "<br/>" + err
    '            End If

    '            actuals.Clear()
    '            count = 1
    '        End If

    '        Dim ia_cd As String = row.Item("IA_CD")
    '        Dim ia_cd_n2 As String = ia_cd + "." + row.Item("IA_CD_N2").ToString()
    '        Dim year As String = row.Item("YEAR")
    '        Dim servicetype As String = row.Item("SERVICE_CD")
    '        Dim margin As Double = 0.0
    '        If Not IsDBNull(row.Item("MARGIN")) Then
    '            margin = row.Item("MARGIN")
    '        End If
    '        Dim total_sales As Double = 0.0
    '        If Not IsDBNull(row.Item("TOTAL_SALES")) Then
    '            total_sales = row.Item("TOTAL_SALES")
    '        End If

    '        If total_sales <> 0.0 AndAlso margin <> 0.0 Then
    '            Dim accountRef As Account = New Account()
    '            accountRef.MIAMI_account_ID__c = ia_cd_n2

    '            Dim actual As New Account_Actual_Miami__c()
    '            actual.ExternalID__c = ia_cd_n2 + "-" + year + "-" + servicetype
    '            actual.Account_Name__r = accountRef
    '            actual.Year__c = year
    '            Dim err As String = _getSFPickListValue(_SF_ACC_SRV_TYPE, "_SF_ACC_SRV_TYPE", servicetype, actual.Service_Type__c)
    '            If Not String.IsNullOrEmpty(err) Then
    '                errors += "<br/>" + err
    '            End If
    '            actual.Actual_Sales__c = Math.Round(total_sales, 0)
    '            actual.Actual_Sales__cSpecified = True
    '            actual.Margin_Amount__c = Math.Round(margin, 0)
    '            actual.Margin_Amount__cSpecified = True
    '            actuals.Add(actual)
    '        End If
    '    Next
    '    If count > 0 Then
    '        Dim err As String = String.Empty
    '        err = _upsertActuals(actuals)
    '        If Not String.IsNullOrEmpty(err) Then
    '            errors += "<br/>" + err
    '        End If
    '    End If
    '    Console.WriteLine("--- Actuals Level 3 Sum L3 OK")

    '    errors += "<br/>--- Actuals Level 3 Sum L3 OK"
    '    Return errors
    'End Function

    'Private Function _actualsLevel2SumLevel4() As String
    '    'Sales and Margin per Level 2 using Level 4 grouped by Level 2 and Service type

    '    Dim errors As String = String.Empty

    '    Console.WriteLine("--- Actuals Level 2 Sum L4")
    '    Dim sb As New StringBuilder()
    '    sb.Append("Select H_LINK.IA_CD, H_LINK.IA_CD_N2, E_DIVISION.SERVICE_CD, A_SALES.YEAR,")
    '    sb.Append(" sum(Nvl(A_SALES.TOTAL_SALES,0)*H_CURRENCY.CHANGE_RATE) as TOTAL_SALES,")
    '    sb.Append(" sum((Nvl(A_SALES.TOTAL_SALES,0)-Nvl(A_SALES.DIRECT_COST,0))*H_CURRENCY.CHANGE_RATE) as MARGIN")
    '    sb.Append(" From A_SALES")
    '    sb.Append(" Inner Join E_DIVISION On A_SALES.COUNTRY_CD=E_DIVISION.COUNTRY_CD")
    '    sb.Append(" And A_SALES.SENDER_ID=E_DIVISION.SENDER_ID")
    '    sb.Append(" And A_SALES.DIVISION_CD=E_DIVISION.DIVISION_CD")
    '    sb.Append(" Inner Join H_LINK On A_SALES.COUNTRY_CD=E_DIVISION.COUNTRY_CD")
    '    sb.Append(" And A_SALES.SENDER_ID=H_LINK.SENDER_ID")
    '    sb.Append(" And A_SALES.DIVISION_CD=H_LINK.DIVISION_CD")
    '    sb.Append(" And A_SALES.CLIENT_CD=H_LINK.CLIENT_CD")
    '    sb.Append(" Inner Join H_CURRENCY On A_SALES.YEAR=H_CURRENCY.YEAR")
    '    sb.Append(" And A_SALES.MONTH=H_CURRENCY.MONTH")
    '    sb.Append(" And A_SALES.COUNTRY_CD=H_CURRENCY.COUNTRY_CD")
    '    sb.Append(" Inner Join E_LEVEL On E_LEVEL.IA_CD = H_LINK.IA_CD")
    '    sb.Append(" And E_LEVEL.IA_CD_N2 = H_LINK.IA_CD_N2")
    '    sb.Append(" And E_LEVEL.IA_CD_N3 = H_LINK.IA_CD_N3")
    '    sb.Append(" And E_LEVEL.IA_CD_N4 = H_LINK.IA_CD_N4")
    '    sb.Append(" Inner Join E_INT_ACCOUNT On H_LINK.IA_CD=E_INT_ACCOUNT.IA_CD")
    '    sb.Append(" Where A_SALES.YEAR||'-'||A_SALES.MONTH Between '")
    '    sb.Append(_fromYearMonth)
    '    sb.Append("' And '")
    '    sb.Append(_toYearMonth)
    '    sb.Append("' And H_LINK.IA_CD>0")
    '    sb.Append(" And H_LINK.IA_CD_N2>0")
    '    sb.Append(" And H_LINK.IA_CD_N3>0")
    '    sb.Append(" And H_LINK.IA_CD_N4>0")
    '    sb.Append(" And E_INT_ACCOUNT.INACTIVE_FLAG=0")
    '    If _DEBUG Then
    '        sb.Append(" And H_LINK.IA_CD Between 849 And 849")
    '        'sb.Append(" And E_DIVISION.SERVICE_CD='SRC'")
    '    End If
    '    sb.Append(" Group By H_LINK.IA_CD, H_LINK.IA_CD_N2, E_DIVISION.SERVICE_CD, A_SALES.YEAR")
    '    sb.Append(" Order By H_LINK.IA_CD, H_LINK.IA_CD_N2, E_DIVISION.SERVICE_CD, A_SALES.YEAR")
    '    Dim sql As String = sb.ToString()

    '    Dim count As Integer = 0
    '    Dim actuals As IList(Of Account_Actual_Miami__c) = New List(Of Account_Actual_Miami__c)

    '    Dim dataTable As New DataTable
    '    Dim sqlError As String = _oracleConnectionDWH.Requete(sql, dataTable)
    '    For Each row As System.Data.DataRow In dataTable.Rows
    '        count = count + 1
    '        If count > 200 Then
    '            Dim err As String = String.Empty
    '            err = _upsertActuals(actuals)
    '            If Not String.IsNullOrEmpty(err) Then
    '                errors += "<br/>" + err
    '            End If

    '            actuals.Clear()
    '            count = 1
    '        End If

    '        Dim ia_cd As String = row.Item("IA_CD")
    '        Dim ia_cd_n2 As String = ia_cd + "." + row.Item("IA_CD_N2").ToString()
    '        Dim year As String = row.Item("YEAR")
    '        Dim servicetype As String = row.Item("SERVICE_CD")
    '        Dim margin As Double = 0.0
    '        If Not IsDBNull(row.Item("MARGIN")) Then
    '            margin = row.Item("MARGIN")
    '        End If
    '        Dim total_sales As Double = 0.0
    '        If Not IsDBNull(row.Item("TOTAL_SALES")) Then
    '            total_sales = row.Item("TOTAL_SALES")
    '        End If

    '        If total_sales <> 0.0 AndAlso margin <> 0.0 Then
    '            Dim accountRef As Account = New Account()
    '            accountRef.MIAMI_account_ID__c = ia_cd_n2

    '            Dim actual As New Account_Actual_Miami__c()
    '            actual.ExternalID__c = ia_cd_n2 + "-" + year + "-" + servicetype
    '            actual.Account_Name__r = accountRef
    '            actual.Year__c = year
    '            Dim err As String = _getSFPickListValue(_SF_ACC_SRV_TYPE, "_SF_ACC_SRV_TYPE", servicetype, actual.Service_Type__c)
    '            If Not String.IsNullOrEmpty(err) Then
    '                errors += "<br/>" + err
    '            End If
    '            actual.Actual_Sales__c = Math.Round(total_sales, 0)
    '            actual.Actual_Sales__cSpecified = True
    '            actual.Margin_Amount__c = Math.Round(margin, 0)
    '            actual.Margin_Amount__cSpecified = True
    '            actuals.Add(actual)
    '        End If
    '    Next
    '    If count > 0 Then
    '        Dim err As String = String.Empty
    '        err = _upsertActuals(actuals)
    '        If Not String.IsNullOrEmpty(err) Then
    '            errors += "<br/>" + err
    '        End If
    '    End If
    '    Console.WriteLine("--- Actuals Level 2 Sum L4 OK")

    '    errors += "<br/>--- Actuals Level 2 Sum L4 OK"
    '    Return errors
    'End Function

    'Private Sub _Periods()
    '    Dim sql As String = "Select Max(YEAR_MONTH) as YEAR_MONTH From R_TIME Where ACTIVE_FLAG=1"
    '    Dim dataTable As New DataTable
    '    Dim sqlError As String = _oracleConnectionDWH.Requete(sql, dataTable)
    '    '_toYearMonth = _oracleCommandDWH.ExecuteScalar()
    '    _toYearMonth = dataTable.Rows(0).Item(0)
    '    _fromYearMonth = (Integer.Parse(_toYearMonth.Substring(0, 4)) - 1).ToString + "-01"
    '    Console.WriteLine("--- Period from {0} to {1}", _fromYearMonth, _toYearMonth)
    'End Sub

    'Private Function _resetActuals() As String

    '    Dim errors As String = String.Empty

    '    Try
    '        Dim done As Boolean = False
    '        Dim count As Integer = 0
    '        'retrieve the first 500 actuals
    '        Dim sql As String = "Select Id, Actual_Sales__c, Margin_Amount__c From Account_Actual_Miami__c"

    '        If _DEBUG Then
    '            sql += " Where ExternalID__c = '120'"
    '        End If

    '        Dim result As QueryResult = _binding.query(sql)
    '        If result.size > 0 Then
    '            While Not done
    '                Dim objects() As sObject = result.records
    '                Dim numberOfActuals = objects.Length
    '                For i As Integer = 0 To numberOfActuals - 1
    '                    Dim actual As Account_Actual_Miami__c = objects(i)
    '                    'Debug.Print("ID: {0} | Sales:{1} | Margin: {2}", actual.Id, actual.Actual_Sales__c, actual.Margin_Amount__c)
    '                    actual.Actual_Sales__c = 0.0
    '                    actual.Actual_Sales__cSpecified = True
    '                    actual.Margin_Amount__c = 0.0
    '                    actual.Margin_Amount__cSpecified = True
    '                Next

    '                'update
    '                Dim max As Integer = 200
    '                Dim updateActuals(max) As Account_Actual_Miami__c
    '                Dim loopCount As Integer = numberOfActuals \ max 'interger part of division
    '                If loopCount > 0 Then
    '                    For iLoop As Integer = 0 To loopCount - 1
    '                        For i As Integer = 0 To max - 1
    '                            updateActuals(i) = objects(iLoop * max + i)
    '                        Next
    '                        Dim actuals_save() As SaveResult = _binding.update(updateActuals)
    '                        For Each updateResult As SaveResult In actuals_save
    '                            If updateResult.success Then
    '                                'Debug.Print("Result: OK " + upsertResult.id)
    '                            Else
    '                                Debug.Print("Errors:")
    '                                errors += "Errors:<br/>"
    '                                Dim sferrors() As miami2salesforce.salesforce.Error = updateResult.errors
    '                                For Each err As miami2salesforce.salesforce.Error In sferrors
    '                                    Debug.Print(err.message)
    '                                    errors += err.message + "<br/>"
    '                                Next
    '                            End If 'success
    '                        Next 'updateResult

    '                    Next
    '                End If 'loopCount > 0

    '                Dim rest As Integer = numberOfActuals Mod max
    '                If rest > 0 Then
    '                    ReDim updateActuals(rest)
    '                    For i As Integer = 0 To rest - 1
    '                        updateActuals(i) = objects(loopCount * max + i)
    '                    Next
    '                    Dim actuals_save() As SaveResult = _binding.update(updateActuals)
    '                    For Each updateResult As SaveResult In actuals_save
    '                        If updateResult.success Then
    '                            'Debug.Print("Result: OK " + upsertResult.id)
    '                        Else
    '                            Debug.Print("Errors:")
    '                            errors += "Errors:<br/>"
    '                            Dim sferrors() As miami2salesforce.salesforce.Error = updateResult.errors
    '                            For Each err As miami2salesforce.salesforce.Error In sferrors
    '                                Debug.Print(err.message)
    '                                errors += err.message + "<br/>"
    '                            Next
    '                        End If 'success
    '                    Next 'updateResult
    '                End If 'max > 0


    '                If result.done Then
    '                    done = True
    '                Else
    '                    'retrieve the next 500 actuals
    '                    result = _binding.queryMore(result.queryLocator)
    '                End If

    '            End While
    '        Else
    '            Console.WriteLine("No actuals found in Salesforce")
    '            errors += "<br/>No actuals found in Salesforce"
    '        End If

    '    Catch ex As Exception
    '        Console.WriteLine(ex.Message)
    '    End Try

    '    Console.WriteLine("--- Actuals Reset OK")

    '    errors += "<br/>--- Actuals Reset OK"
    '    Return errors
    'End Function

    'Private Sub _Attachment(ByVal filePath As String)
    '    Dim fileInfo As New FileInfo(filePath)
    '    If fileInfo.Length > 5000000 Then
    '        Console.WriteLine("File too big: {0} {1}", filePath, fileInfo.Length)
    '        Return
    '    End If


    '    Dim fileStream As FileStream = File.OpenRead(filePath)
    '    Dim buffer(fileInfo.Length) As Byte
    '    fileStream.Read(buffer, 0, buffer.Length)

    '    'doc.Body = buffer
    '    'Dim doc As New Document
    '    'doc.Name = fileInfo.Name
    '    'doc.Description = fileInfo.FullName

    '    Dim attachment As New Attachment
    '    attachment.Body = buffer
    '    attachment.Name = fileInfo.Name
    '    attachment.Description = fileInfo.FullName
    '    attachment.ParentId = "001c000000Q6l47AAB"
    '    Dim results() As SaveResult = _binding.create(New sObject() {attachment})
    '    For Each createResult As SaveResult In results
    '        If createResult.success Then
    '        Else
    '            Debug.Print("Errors:")
    '            Dim errors() As miami2salesforce.salesforce.Error = createResult.errors
    '            For Each err As miami2salesforce.salesforce.Error In errors
    '                Debug.Print(err.message)
    '            Next

    '        End If
    '    Next


    'End Sub

    Private Function _load_user() As String
        Dim err As New StringBuilder
        Dim sfcodes As List(Of String) = New List(Of String)
        Dim value As New KeyValuePair(Of String, String)
        'Existing SF user in Miami gate parameters
        Dim lastPARAM_CD As Integer = 0
        Dim Sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_USR_IADOM%' Order By PARAM_CD"
        Dim dataTable As New DataTable
        Dim sqlError As String = _oracleConnectionGATE.Requete(Sql, dataTable)
        For Each row As System.Data.DataRow In dataTable.Rows
            Dim param_code As String = row.Item("PARAM_CD")
            param_code = param_code.Replace("SF_USR_IADOM_", "")
            lastPARAM_CD = Integer.Parse(param_code)

            Dim param_value As String = row.Item("PARAM_VALUE")
            Dim codes() As String = param_value.Split("¤")
            Dim MiamiCode As String = codes(0)
            Dim SFcode As String = codes(1)
            Debug.WriteLine(String.Format("User: {0} => {1}", MiamiCode, SFcode))
            If Not String.IsNullOrEmpty(MiamiCode) And (Not (_SF_USR_IADOM.ContainsKey(MiamiCode) Or _SF_USR_IADOM.ContainsValue(SFcode))) Then
                _SF_USR_IADOM.Add(MiamiCode, SFcode)
            End If
            sfcodes.Add(SFcode.ToLower)

        Next 'read

        Try
            Dim parameters As TableauParametres = New TableauParametres()
            Dim done As Boolean = False
            'retrieve the first 500 accounts
            Dim result As QueryResult = _binding.query("SELECT Id, Username, Email, Name, ProfileId, UserRoleId, UserType, Profile.Name FROM User Where Profile.Name Like '%GCS%' Or Profile.Name='System Administrator'")
            If result.size > 0 Then
                While Not done
                    Dim objects() As sObject = result.records
                    Dim count = objects.Length
                    For i As Integer = 0 To count - 1
                        Dim user As User = objects(i)

                        Dim bInsert As Boolean = True
                        bInsert = Not sfcodes.Contains(user.Id.ToLower)
                        'For Each kvp As KeyValuePair(Of String, String) In _SF_USR_IADOM
                        '    Dim SFcode As String = kvp.Value
                        '    If SFcode.ToLower.Equals(user.Id.ToLower) Then
                        '        bInsert = False
                        '        Exit For
                        '    End If
                        'Next

                        If bInsert Then
                            lastPARAM_CD += 1
                            Dim param_code As String = "SF_USR_IADOM_" + lastPARAM_CD.ToString("D4")
                            Dim split = user.Username.Split("@")
                            Dim param_value As String = "¤" & user.Id + "¤" & split(0)

                            Dim insert As StringBuilder = New StringBuilder()
                            insert.Append("Insert Into PARAM_R (PARAM_CD, PARAM_VALUE, CREATION_DT, CREATION_BY)")
                            insert.Append(" VALUES(:PARAM_CD, :PARAM_VALUE, SYSDATE, 'Salesforce')")
                            parameters.PurgeParametre()
                            parameters.AjouterParametreChaine(":PARAM_CD", param_code)
                            parameters.AjouterParametreChaine(":PARAM_VALUE", param_value)
                            sqlError = _oracleConnectionGATE.Requete(insert.ToString(), parameters)
                        End If

                    Next
                    If result.done Then
                        done = True
                    Else
                        'retrieve the next 500 accounts
                        result = _binding.queryMore(result.queryLocator)
                    End If
                End While
            Else
                Console.WriteLine("No user found in Salesforce")
                err.Append("No user found in Salesforce")
            End If

        Catch ex As Exception
            err.Append(ex.Message)
            Console.WriteLine(ex.Message)
        End Try

        Return err.ToString
    End Function

    'Private Function _references() As String
    '    Dim err As New StringBuilder
    '    Dim sql As String

    '    'Service type Miami
    '    sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_ACC_SRV_TYPE%'"
    '    Dim dataTable As New DataTable
    '    Dim sqlError As String = _oracleConnectionGATE.Requete(sql, dataTable)
    '    For Each row As System.Data.DataRow In dataTable.Rows
    '        Dim param_value As String = row.Item("PARAM_VALUE")
    '        Dim codes() As String = param_value.Split("¤")
    '        Dim MiamiCode As String = codes(0)
    '        Dim SFcode As String = codes(1)
    '        Debug.WriteLine(String.Format("Service type: {0} => {1}", MiamiCode, SFcode))

    '        _SF_ACC_SRV_TYPE.Add(MiamiCode, SFcode)
    '    Next 'read

    '    'Miami actual picklists
    '    Dim describeSObjectResult As DescribeSObjectResult = _binding.describeSObject("Account_Actual_Miami__c")
    '    Dim fields() As Field = describeSObjectResult.fields
    '    For Each field As Field In fields
    '        'Service type
    '        If field.type.Equals(fieldType.picklist) AndAlso field.name.ToLower.Equals("Service_Type__c".ToLower) Then
    '            Dim pickListEntries() As PicklistEntry = field.picklistValues
    '            For Each pickListEntry As PicklistEntry In pickListEntries
    '                Debug.WriteLine(pickListEntry.value)
    '                If Not _SF_ACC_SRV_TYPE.ContainsValue(pickListEntry.value) Then
    '                    err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
    '                End If
    '            Next
    '        End If
    '    Next

    '    'Account Type Salesforce
    '    sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_ACC_SF_TYPE%'"
    '    dataTable.Clear()
    '    sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
    '    For Each row As System.Data.DataRow In dataTable.Rows
    '        Dim param_value As String = row.Item("PARAM_VALUE")
    '        Dim codes() As String = param_value.Split("¤")
    '        Dim MiamiCode As String = codes(0)
    '        Dim SFcode As String = codes(1)
    '        Debug.WriteLine(String.Format("Account Type Salesforce: {0} => {1}", MiamiCode, SFcode))

    '        _SF_ACC_SF_TYPE.Add(MiamiCode, SFcode)
    '    Next 'read

    '    'Account Status Salesforce
    '    sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_ACC_SF_STATUS%'"
    '    dataTable.Clear()
    '    sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
    '    For Each row As System.Data.DataRow In dataTable.Rows
    '        Dim param_value As String = row.Item("PARAM_VALUE")
    '        Dim codes() As String = param_value.Split("¤")
    '        Dim MiamiCode As String = codes(0)
    '        Dim SFcode As String = codes(1)
    '        Debug.WriteLine(String.Format("Account Status Salesforce: {0} => {1}", MiamiCode, SFcode))

    '        _SF_ACC_SF_STATUS.Add(MiamiCode, SFcode)
    '    Next 'read

    '    'Account Vertical
    '    sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_ACC_VERTICAL%'"
    '    dataTable.Clear()
    '    sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
    '    For Each row As System.Data.DataRow In dataTable.Rows
    '        Dim param_value As String = row.Item("PARAM_VALUE")
    '        Dim codes() As String = param_value.Split("¤")
    '        Dim MiamiCode As String = codes(0)
    '        Dim SFcode As String = codes(1)
    '        Debug.WriteLine(String.Format("Account Vertical Salesforce: {0} => {1}", MiamiCode, SFcode))

    '        _SF_ACC_VERTICAL.Add(MiamiCode, SFcode)
    '    Next 'read

    '    'Account Director
    '    sql = "Select * From PARAM_R Where PARAM_CD Like 'SF_ACC_DIRECTOR%'"
    '    dataTable.Clear()
    '    sqlError = _oracleConnectionGATE.Requete(sql, dataTable)
    '    For Each row As System.Data.DataRow In dataTable.Rows
    '        Dim param_value As String = row.Item("PARAM_VALUE")
    '        Dim codes() As String = param_value.Split("¤")
    '        Dim MiamiCode As String = codes(0)
    '        Dim SFcode As String = codes(1)
    '        Debug.WriteLine(String.Format("Account Director Salesforce: {0} => {1}", MiamiCode, SFcode))

    '        _SF_ACC_DIRECTOR.Add(MiamiCode, SFcode)
    '    Next 'read

    '    'Account picklists
    '    describeSObjectResult = _binding.describeSObject("Account")
    '    fields = describeSObjectResult.fields
    '    For Each field As Field In fields
    '        'Account Type Salesforce
    '        If field.type.Equals(fieldType.picklist) AndAlso field.name.ToLower.Equals("Type".ToLower) Then
    '            Dim pickListEntries() As PicklistEntry = field.picklistValues
    '            For Each pickListEntry As PicklistEntry In pickListEntries
    '                Debug.WriteLine(pickListEntry.value)
    '                If Not _SF_ACC_SF_TYPE.ContainsValue(pickListEntry.value) Then
    '                    err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
    '                End If
    '            Next
    '        End If

    '        'Account Status Salesforce
    '        If field.type.Equals(fieldType.picklist) AndAlso field.name.ToLower.Equals("Account_Status__c".ToLower) Then
    '            Dim pickListEntries() As PicklistEntry = field.picklistValues
    '            For Each pickListEntry As PicklistEntry In pickListEntries
    '                Debug.WriteLine(pickListEntry.value)
    '                If Not _SF_ACC_SF_STATUS.ContainsValue(pickListEntry.value) Then
    '                    err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
    '                End If
    '            Next
    '        End If

    '        'Account Vertical
    '        If field.type.Equals(fieldType.picklist) AndAlso field.name.ToLower.Equals("GCS_Vertical__c".ToLower) Then
    '            Dim pickListEntries() As PicklistEntry = field.picklistValues
    '            For Each pickListEntry As PicklistEntry In pickListEntries
    '                Debug.WriteLine(pickListEntry.value)
    '                If Not _SF_ACC_VERTICAL.ContainsValue(pickListEntry.value) Then
    '                    err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
    '                End If
    '            Next
    '        End If

    '        'Account Director
    '        If field.type.Equals(fieldType.picklist) AndAlso field.name.ToLower.Equals("GCS_account_director__c".ToLower) Then
    '            Dim pickListEntries() As PicklistEntry = field.picklistValues
    '            For Each pickListEntry As PicklistEntry In pickListEntries
    '                Debug.WriteLine(pickListEntry.value)
    '                If Not _SF_ACC_DIRECTOR.ContainsValue(pickListEntry.value) Then
    '                    err.Append(String.Format("{0} picklist missing value in Miami: {1}<br/>", field.name, pickListEntry.value))
    '                End If
    '            Next
    '        End If
    '    Next

    '    Return err.ToString
    'End Function

    'Private Function _getSFPickListValue(ByVal _SF As Dictionary(Of String, String), ByVal picklistName As String, ByVal miamiValue As String, ByRef SFValue As String)
    '    Dim err As String = String.Empty

    '    If _SF.ContainsKey(miamiValue) Then
    '        SFValue = _SF(miamiValue)
    '    Else
    '        err = String.Format("Missing Miami value for {0}: {1}", picklistName, miamiValue)
    '        SFValue = "Unknown"
    '    End If

    '    Return err
    'End Function

End Class
