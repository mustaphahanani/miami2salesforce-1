Imports System.Net.Mail
Imports System.Net
Imports System.Text

Module Main

    Sub Main()

        'Parameters: My.Settings.xxx

        'Uncomment this to generate a crypted value (username, password)
        'Dim sEncrypt = CryptoService.Encrypt("svc.randstats")
        'Debug.Print(sEncrypt)
        'sEncrypt = CryptoService.Encrypt("t5mGij2Y3ic2ZAviVlHw")
        'Debug.Print(sEncrypt)

        'Dim sDecrypt = CryptoService.Decrypt("OkQF5CetvBLQlBQ/d07OV/KIBHmUJaZD2E+/PvBPAkE=")
        'Debug.Print(sDecrypt)
        'Dim sDecryptP = CryptoService.Decrypt("UaaEAp7PLY43Cg7VSVt65w==")
        'Debug.Print(sDecryptP)

        'Email settings
        Dim smtpHost As String = My.Settings.ServeurSmtp
        Dim smtpPort As Integer = 25
        Dim smtpClient As SmtpClient = New SmtpClient(smtpHost, smtpPort)

        'Database settings
        Dim BaseMiamiGateConnect As String = My.Settings.BaseMiamiGateConnect
        Dim BaseMiamiODSConnect As String = My.Settings.BaseMiamiODSConnect
        Dim BaseMiamiDWHConnect As String = My.Settings.BaseMiamiDWHConnect

        'Processing
        Const Upload_Account As String = "upload_account"
        Const Download_Opportunity As String = "download_opportunity"
        'Const Download_lead As String = "download_lead"

        Dim operation As String = My.Settings.Processing

        'Debug mode
        Dim bDebug As Boolean = My.Settings.Debug

        Dim result As String = String.Empty
        System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls

        smtpClient.EnableSsl = False
        smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network
        smtpClient.UseDefaultCredentials = False
        smtpClient.Credentials = CredentialCache.DefaultNetworkCredentials

        Dim from As String = "noreply@randstad.fr"
        Dim tos As String = getRecipients("SmtpRecipient")

        Dim mailAddressCollection As MailAddressCollection = New MailAddressCollection()

        Dim mailMessage As MailMessage = New MailMessage()
        mailMessage.From = New MailAddress(from)

        Dim toList As String() = tos.Split(",")
        For Each _to As String In toList
            mailMessage.To.Add(New MailAddress(_to))
        Next

        mailMessage.IsBodyHtml = True
        mailMessage.Subject = "Miami-SF "

        'check # argument(s)
        'operation = String.Empty
        If String.IsNullOrEmpty(operation) Then
            result = "Operation name missing"
            Console.WriteLine(result)
            mailMessage.Body = String.Format("Hello, <br/>Here is the result:{0}<br/><br/> Regards", result)
            Try
                smtpClient.Send(mailMessage)
            Catch ex As Exception
                Debug.Print(ex.Message)
                Console.WriteLine(ex.Message)
            End Try
            Return
        End If

        'check argument value        
        Select Case operation
            Case Upload_Account
            Case Download_Opportunity
            Case Else
                result = String.Format("Operation unknown: {0}", operation)
                Console.WriteLine(result)
                mailMessage.Body = String.Format("Hello, <br/>Here is the result:{0}<br/><br/> Regards", result)
                Try
                    smtpClient.Send(mailMessage)
                Catch ex As Exception
                    Debug.Print(ex.Message)
                    Console.WriteLine(ex.Message)
                End Try
                Return
        End Select

        Debug.Print("Start")
        Console.WriteLine("Start")

        Dim config As New Config()
        'config.readConfig()

        Dim sfLogin As New SalesforceLogin(config)

        Dim err As String = String.Empty
        Dim ok As Boolean = sfLogin.login(err)
        result += err
        If Not ok Then
            mailMessage.Body = String.Format("Hello, <br/>Here is the result:{0}<br/><br/> Regards", result)
            Try
                smtpClient.Send(mailMessage)
            Catch ex As Exception
                Debug.Print(ex.Message)
                Console.WriteLine(ex.Message)
            End Try
            Return
        End If

        Select Case operation
            Case Upload_Account
                Dim load As New LoadAccount(sfLogin.Binding, BaseMiamiGateConnect, BaseMiamiDWHConnect)
                result = load.loadAccount(bDebug)

            Case Download_Opportunity
                Dim res As String
                Dim numberOfDays As Integer = My.Settings.NumberOfDays
                Dim loadOpportunities As New LoadOpportunity(sfLogin.Binding, BaseMiamiGateConnect, BaseMiamiODSConnect)
                'Dim sNumberOfDays As String = My.Settings.NumberOfDays
                'If Not String.IsNullOrEmpty(sNumberOfDays) Then
                '    If IsNumeric(sNumberOfDays) Then
                '        numberOfDays = Integer.Parse(sNumberOfDays)
                '    End If
                'End If
                ''res = loadOpportunities.LoadOpportunity(numberOfDays)
                ''result += "** Opportunity **<br/>" + res

                'LoadOpportunity()
                Dim LoadOpportunity As New LoadOpportunity(sfLogin.Binding, BaseMiamiGateConnect, BaseMiamiODSConnect)
                res = LoadOpportunity.LoadOpportunity(numberOfDays)
                result += "** LoadOpportunity **<br/>" + res

                ''loadLeads()
                Dim loadLeads As New loadLeads(sfLogin.Binding, BaseMiamiGateConnect, BaseMiamiODSConnect)
                res = loadLeads.loadLeads(numberOfDays)
                result += "** Leads **<br/>" + res

                ''loadRSR_Opportunity_Qualification()
                Dim loadRSR_Opportunity_Qualification As New miami2salesforce.load_RSR_Opportunity_Qualification(sfLogin.Binding, BaseMiamiGateConnect, BaseMiamiODSConnect)
                res = loadRSR_Opportunity_Qualification.loadRSR_Opportunity_Qualification(numberOfDays)
                result += "** RSR_Opportunity_Qualification **<br/>" + res

                'loadService_Pricing_framework()
                Dim loadService_Pricing_framework As New loadService_Pricing_framework(sfLogin.Binding, BaseMiamiGateConnect, BaseMiamiODSConnect)
                res = loadService_Pricing_framework.loadService_Pricing_framework(numberOfDays)
                result += "** Service_Pricing_framework **<br/>" + res

                ''load_Attainable_potential_Randstad()
                Dim load_Attainable_potential_Randstad As New load_Attainable_potential_Randstad(sfLogin.Binding, BaseMiamiGateConnect, BaseMiamiODSConnect)
                res = load_Attainable_potential_Randstad.load_Attainable_potential_Randstad(numberOfDays)
                result += "** Attainable_potential_Randstad **<br/>" + res

                ''load_Global_potential_account()
                Dim load_Global_potential_account As New load_Global_potential_account(sfLogin.Binding, BaseMiamiGateConnect, BaseMiamiODSConnect)
                res = load_Global_potential_account.load_Global_potential_account(numberOfDays)
                result += "** Attainable_potential_Randstad **<br/>" + res

                ''loadOPCO_Contact_matrix()
                Dim loadOPCO_Contact_matrix As New loadOPCO_Contact_matrix(sfLogin.Binding, BaseMiamiGateConnect, BaseMiamiODSConnect)
                res = loadOPCO_Contact_matrix.loadOPCO_Contact_matrix(numberOfDays)
                result += "** loadOPCO_Contact_matrix **<br/>" + res

                ''loadContact()
                Dim loadContact As New loadContact(sfLogin.Binding, BaseMiamiGateConnect, BaseMiamiODSConnect)
                res = loadContact.loadContact(numberOfDays)
                result += "** Contact **<br/>" + res

                ''LoadOpportunity_history()
                Dim LoadOpportunity_history As New LoadOpportunity_history(sfLogin.Binding, BaseMiamiGateConnect, BaseMiamiODSConnect)
                res = LoadOpportunity_history.LoadOpportunity_history(numberOfDays)
                result += "** LoadOpportunity_history **<br/>" + res

                ''loadProduct2()
                Dim loadProduct2 As New loadProduct2(sfLogin.Binding, BaseMiamiGateConnect, BaseMiamiODSConnect)
                res = loadProduct2.loadProduct2(numberOfDays)
                result += "** loadProduct **<br/>" + res

                ''Load_Name_Account()
                Dim Load_Name_Account As New Load_Name_Account(sfLogin.Binding, BaseMiamiGateConnect, BaseMiamiODSConnect)
                res = Load_Name_Account.Load_Name_Account(numberOfDays)
                result += "** Load_Name_Account **<br/>" + res

                ''loadUser()
                Dim loadUser As New loadUser(sfLogin.Binding, BaseMiamiGateConnect, BaseMiamiODSConnect)
                res = loadUser.loadUser(numberOfDays)
                result += "** loadUser **<br/>" + res

                'Case Download_lead
                '    Dim reslead As String
                '    Dim numberOfDays As Integer = 0
                '    Dim load As New loadLeads(sfLogin.Binding, BaseMiamiGateConnect, BaseMiamiODSConnect)
                '    Dim sNumberOfDays As String = My.Settings.NumberOfDays")
                '    If Not String.IsNullOrEmpty(sNumberOfDays) Then
                '        If IsNumeric(sNumberOfDays) Then
                '            numberOfDays = Integer.Parse(sNumberOfDays)
                '        End If
                '    End If
                '    reslead = load.loadLeads(numberOfDays)
                '    result += "** Leads **<br/>" + reslead
        End Select

        sfLogin.logout()

        mailMessage.Subject += " - " + config.Environment
        If result.Contains("ERROR") Then
            mailMessage.Subject += " ERROR"
        End If

        mailMessage.Body = String.Format("Hello, <br/>Here is the result:{0}<br/><br/> Regards", result)

        Try
            smtpClient.Send(mailMessage)
        Catch ex As Exception
            Debug.Print(ex.Message)
            Console.WriteLine(ex.Message)
        End Try

        Console.WriteLine("The End")
        'Console.ReadLine()
        Debug.Print("End")


    End Sub

    Private Function getRecipients(ByVal parameterPattern As String) As String
        Dim recipients As New StringBuilder
        ' --------------------------------------------------------------------
        ' Loop over parameters with name corresponding to a pattern (wildcard)
        ' --------------------------------------------------------------------

        'For Each parameter As String In Pilote.ListerParametres(parameterPattern)
        '    Dim value As String = Pilote.DonnerParametre(parameter)
        '    value = value.Trim()
        '    value = value.Trim(",")
        '    If Not String.IsNullOrEmpty(value) Then
        '        If recipients.Length > 0 Then
        '            recipients.Append(", ")
        '        End If
        '        recipients.Append(value)
        '    End If
        'Next

        For Each CSPV As System.Configuration.SettingsPropertyValue In My.Settings.PropertyValues
            Dim propertyName = CSPV.Name
            If propertyName.Contains(parameterPattern) Then
                Dim value As String = CSPV.PropertyValue
                value = value.Trim()
                value = value.Trim(",")
                If Not String.IsNullOrEmpty(value) Then
                    If recipients.Length > 0 Then
                        recipients.Append(", ")
                    End If
                    recipients.Append(value)
                End If
            End If
        Next
        Return recipients.ToString

    End Function

End Module
