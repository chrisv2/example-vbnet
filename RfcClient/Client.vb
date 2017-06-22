Imports RFCCONNECTORLib
Imports System.Runtime.InteropServices

Module Client

    ''' <summary>
    ''' Create a session through NWRFC protocol. This is the preferred method of connecting
    ''' when GUI functionality or Single Sign On (SSO) Is required. This requires SAP GUI 7.50
    ''' or later, or a separate installation of the NWRFC library.
    ''' </summary>
    ''' <param name="saplogon_id">target system (as defined in SAP Logon)</param>
    ''' <returns>session instance</returns>
    Function GetNwRfcSession(saplogon_id As String) As ISession
        Dim session As New NWRfcSession
        session.RfcSystemData.ConnectString = String.Format("SAPLOGON_ID={0}", saplogon_id)
        Return session
    End Function
    ''' <summary>
    ''' Create a session through SOAP. This does not require any additional libraries, but
    ''' you need to enable the service /sap/bc/soap/rfc/ in transaction SICF.
    ''' 
    ''' The port number is usually SAP system number + 8000.
    ''' </summary>
    ''' <param name="host">hostname of the SAP server</param>
    ''' <param name="port">port of the SAP server</param>
    ''' <returns>session instance</returns>
    Function GetSOAPSession(host As String, port As Integer) As ISession
        Dim session As New SoapSession
        session.HttpSystemData.Host = host
        session.HttpSystemData.Port = port
        Return session
    End Function

    ''' <summary>
    ''' Connect through classic RFC. This is not recommended any more, and should be used
    ''' only in legacy situations, where it is not possible to upgrade SAPGUI.
    ''' 
    ''' This mode can only be used with SAP GUI 7.40 or older.
    ''' </summary>
    ''' <param name="saplogon_id">target system (as defined in SAP Logon)</param>
    ''' <returns>session instance</returns>
    Function GetClassicRfcSession(saplogon_id As String) As ISession
        Dim session As New RfcSession
        session.RfcSystemData.ConnectString = String.Format("SAPLOGON_ID={0}", saplogon_id)
        Return session
    End Function

    ''' <summary>
    ''' Log on to the SAP system with the given parameters, and check errors
    ''' </summary>
    ''' <param name="session">session instance</param>
    ''' <param name="client">SAP client, e.g. "001"</param>
    ''' <param name="user">SAP username</param>
    ''' <param name="password">SAP password</param>
    ''' <param name="language">SAP language for connection, e.g. "EN"</param>
    ''' <returns></returns>
    Function Logon(session As ISession, client As String, user As String, password As String, language As String) As Boolean
        ' set up logon data
        session.LogonData.Client = client
        session.LogonData.User = user
        session.LogonData.Password = password
        session.LogonData.Language = language

        ' enable trace file, see http://rfcconnector.com/documentation/kb/0004/
        session.Option("trace.file") = "C:\\tracefile.txt"

        ' connect and check error
        session.Connect()
        If session.Error = True Then
            Console.WriteLine(session.ErrorInfo.Message)
            Return False
        Else
            Return True
        End If
    End Function
    ''' <summary>
    ''' retrieve the signature of BAPI_FLIGHT_GETLIST, set some parameters and call it
    ''' </summary>
    ''' <param name="session">session instance</param>
    Sub callFunction(session As ISession)
        Dim fn As FunctionCall

        ' import the function call definition from backend
        fn = session.ImportCall("BAPI_FLIGHT_GETLIST")

        If session.Error Then
            Console.WriteLine(session.ErrorInfo.Message)
            Return
        End If

        ' find all flights with American Airlines (AA) starting from San Francisco (SFO)
        fn.Importing("AIRLINE").value = "AA"
        ' DESTINATION_FROM Is a structure, so we can access individual fields through the Fields property
        fn.Importing("DESTINATION_FROM").Fields("AIRPORTID").value = "SFO"

        ' call function
        session.CallFunction(fn, True)

        If session.Error Then
            Console.WriteLine(session.ErrorInfo.Message)
            Return
        End If

        ' loop over the rows of the FLIGHT_LIST table
        For Each row As RfcFields In fn.Tables("FLIGHT_LIST").Rows
            Console.WriteLine(row("AIRLINE").value + " " + row("FLIGHTDATE").value)
        Next
    End Sub

    Sub Main()
        Dim session As ISession

        Try
            session = GetNwRfcSession("NPL")
            If Logon(session, "001", "DEVELOPER", "developer1", "EN") Then
                callFunction(session)
            End If
        Catch exc As COMException
            Console.WriteLine(exc.ToString())
        End Try
        Console.WriteLine("press any key to exit")
        Console.ReadLine()
    End Sub

End Module
