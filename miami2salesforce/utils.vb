Imports Methodes.mthconnexion
Imports System.Text

Public Class Utils

    'Add SF multipicklist values into a Mthconnexion parameter
    Public Shared Function _multiPicklist(ByVal values As String, ByVal _SF As Dictionary(Of String, String), ByVal param_name As String, ByVal params As TableauParametres) As String
        If IsNothing(values) Then
            params.AjouterParametre(param_name, String.Empty)
            Return ""
        End If

        Dim param_values As String = String.Empty
        Dim err As String = String.Empty
        Dim split As String() = values.Split(";")
        For Each value In split
            If _SF.ContainsKey(value) Then
                If Not param_values.Equals(String.Empty) Then
                    param_values = String.Concat(param_values, ";")
                End If
                param_values = String.Concat(param_values, _SF(value))
            Else
                If Not err.Equals(String.Empty) Then
                    err = String.Concat(err, ",")
                End If
                err = String.Concat(err, value)
            End If
        Next
        If Not param_values.Equals(String.Empty) Then
            params.AjouterParametreChaine(param_name, param_values)
        Else
            params.AjouterParametre(param_name, String.Empty)
        End If
        Return err
    End Function

    'Truncate a table
    Public Shared Function TruncateODSTables(ByVal table As String, ByVal oracleConnection As MthConnexion) As String
        Dim errors As String
        Dim sb As StringBuilder = New StringBuilder()

        If oracleConnection Is Nothing Then
            Return "oracleConnection is null"
        End If

        If Not String.IsNullOrEmpty(table) Then
            Try
                Dim sql As String = String.Format("Delete From {0} Where 1 = 1", table)
                Dim sqlError As String = oracleConnection.Requete(sql)

            Catch ex As Exception
                sb.Append(ex.Message)
                Console.WriteLine(ex.Message)
            End Try

        Else
            sb.Append("Table name empty")
        End If

        errors = sb.ToString()

        Return errors
    End Function

    Public Shared Function TruncateODSTables(ByVal tableList As List(Of String), ByVal oracleConnection As MthConnexion) As String
        Dim errors As String
        Dim sb As StringBuilder = New StringBuilder()
        'Dim tableList As New List(Of String)
        'tableList.Add("TMPSF_ATT_POT_RAND")
        'tableList.Add("TMPSF_CONTACTS")
        'tableList.Add("TMPSF_GLO_POT_ACC")
        'tableList.Add("TMPSF_LEADS")
        'tableList.Add("TMPSF_OPCO_CTC_MATX")
        'tableList.Add("TMPSF_OPPORTUNITY")
        'tableList.Add("TMPSF_OPPORTUNITY_HIST")
        'tableList.Add("TMPSF_RSR_OPP_QUALIF")
        'tableList.Add("TMPSF_SRV_OPCO_FRAM")

        If oracleConnection Is Nothing Then
            Return "oracleConnection is null"
        End If

        If Not tableList Is Nothing Then
            If tableList.Count > 0 Then
                Try
                    For Each table As String In tableList
                        'Dim sql As String = String.Format("Truncate Table {0}", table)
                        Dim sql As String = String.Format("Delete From {0} Where 1 = 1", table)
                        Dim sqlError As String = oracleConnection.Requete(sql)
                    Next

                Catch ex As Exception
                    sb.Append(ex.Message)
                    Console.WriteLine(ex.Message)
                End Try
            Else
                sb.Append("TableList empty")
            End If
        Else
            sb.Append("TableList is null")
        End If

        errors = sb.ToString()

        Return errors
    End Function

End Class
