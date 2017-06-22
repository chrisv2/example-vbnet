Imports RFCCONNECTORLib
Imports System.Runtime.InteropServicesModule server
    ' ---------------
    ' RFC Server data
    ' ---------------
    '
    ' RFC_DESTINATION must be set up in transaction SM59 as
    ' "TCP/IP Connection" with "Registered Server Program"
    Const RFC_DESTINATION = "ZRFCCTEST"
    ' Hostname of the SAP server we are registering with
    Const GATEWAY_HOST = "was-npl"
    ' Gateway service (this Is "sapgw" plus the system number)
    Const GATEWAY_SERVICE = "sapgw42"

    ' ---------------
    ' RFC Client data
    ' ---------------
    '
    ' This is only used for importing the function definition at startup
    Const SAPLOGON_ID = "NPL"
    Const CLIENT = "001"
    Const USER = "DEVELOPER"
    Const PW = "developer1"
    Const LANGUAGE = "EN"
    ' global RFC server
    Dim WithEvents srv As New NWRfcServer
    ' global function call definition
    Dim fn1 As FunctionCall, fn2 As FunctionCall

    Sub ImportFunctions()        Dim session As New NWRfcSession

        ' import the function definitions we need later
        Console.WriteLine("Retrieving function definitions from remote system...")        session.SystemID = SAPLOGON_ID        session.RfcSystemData.ConnectString = String.Format("SAPLOGON_ID={0}", SAPLOGON_ID)        session.LogonData.Client = CLIENT        session.LogonData.User = USER        session.LogonData.Password = PW        session.LogonData.Language = LANGUAGE
        session.Connect()
        If session.Error Then
            Console.WriteLine("Error: " + session.ErrorInfo.Message)
        End If
        fn1 = session.ImportCall("BAPI_FLIGHT_GETLIST")

        session.Disconnect()
    End Sub
    Private Sub srv_IncomingCall(ByVal fn As FunctionCall) Handles srv.IncomingCall
        Select Case fn.Function
            Case "BAPI_FLIGHT_GETLIST"
                Dim airline As String, dest_from As String, dest_to As String
                Console.WriteLine("Received a call to BAPI_FLIGHT_GETLIST!")

                If fn.Importing.HasKey("AIRLINE") Then
                    airline = fn.Importing("AIRLINE").value
                Else
                    airline = "LH"
                End If
                If fn.Importing.HasKey("DEST_FROM") Then
                    dest_from = fn.Importing("DEST_FROM").value
                Else
                    dest_from = "SFO"
                End If
                If fn.Importing.HasKey("DEST_TO") Then
                    dest_to = fn.Importing("DEST_TO").value
                Else
                    dest_to = "JFK"
                End If
                ' add some static rows 
                Dim r As RfcFields, dt As DateTime
                dt = DateTime.Today
                For i = 1 To 10
                    r = fn.Tables("FLIGHT_LIST").Rows.AddRow()
                    r("AIRLINEID").value = airline
                    r("AIRPORTFR").value = dest_from
                    r("AIRPORTTO").value = dest_to
                    r("CONNECTID").value = i
                    r("FLIGHTDATE").value = dt.AddDays(i)
                Next
            Case Else
                ' we don't know about this function, so we raise an exception
                fn.RaiseException("SYSTEM_ERROR")
        End Select
    End Sub    ''' <summary>
    ''' Event handler (callback) which gets called when a user "logs on" to our server. The
    ''' username we get here is already authenticated with SAP, so we do not need to check
    ''' the password. However, we can limit our service to certain usernames if we wish to.
    ''' </summary>
    ''' <param name="li">Logon Information</param>
    Private Sub srv_Logon(ByVal li As RfcServerLogonInfo) Handles srv.Logon
        If li.User = "DEVELOPER" Then
            li.RequestAllowed = True
        Else
            li.RequestAllowed = False
        End If
    End Sub    ''' <summary>
    ''' Event handler (callback) which gets called when there is an error with the server
    ''' </summary>
    ''' <param name="ei"></param>
    Private Sub srv_ServerError(ByVal ei As RfcServerErrorInfo) Handles srv.ServerError
        ' log the error to console
        Console.WriteLine(ei.Message)
        ' restart the server
        ei.Restart = True
    End Sub

    Sub Main()

        ' import the function definition(s)
        ImportFunctions()
        ' connection data for SAP server we are registering with
        srv.ProgramID = RFC_DESTINATION
        srv.GatewayHost = GATEWAY_HOST
        srv.GatewayService = GATEWAY_SERVICE

        ' install the function(s) we want to handle
        srv.InstallFunction(fn1)

        ' Start the server
        Console.WriteLine("Starting server...")
        srv.Serve()
        If srv.Error Then
            Console.WriteLine("Error: Could not start server.")
            Console.WriteLine(srv.ErrorInfo.Message)
            Exit Sub
        End If
        ' wait until key is pressed
        Console.WriteLine("Server running. Press key to stop...")
        Console.ReadLine()

        ' then stop the server
        srv.Shutdown()
        Console.WriteLine("Server exited. Press key to leave...")
        Console.ReadLine()
    End Sub

End Module