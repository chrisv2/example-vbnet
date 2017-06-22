Imports RFCCONNECTORLib
Imports System.Runtime.InteropServices

Module Reader
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
    ''' Use the TableReader to read the SAP table SFLIGHT with RFC_READ_TABLE
    ''' 
    ''' For limitations of RFC_READ_TABLE And possible solutions, please refer to
    ''' http://rfcconnector.com/documentation/kb/0007/
    ''' </summary>
    ''' <param name="session">session instance</param>
    Sub ReadTable(session As ISession)
        Dim tr As TableReader
        ' get a TableReader instance for table SFLIGHT
        tr = session.GetTableReader("SFLIGHT")
        ' add the fields/columns we want to read
        tr.Fields.Add("CARRID")
        tr.Fields.Add("CONNID")
        tr.Fields.Add("FLDATE")
        ' set a query expression/WHERE clause: select all entries where CARRID equals 'LH'
        tr.Query.Add("CARRID EQ 'LH'")
        ' read the table, starting at row 0, with no limit
        tr.Read(0, 0)
        ' process the result
        For Each row As RfcFields In tr.Rows
            Console.WriteLine(row("CONNID").value + " " + row("FLDATE").value + " " + row("CARRID").value)
        Next
    End Sub

    Sub Main()
        Dim session As ISession

        Try
            session = GetNwRfcSession("NPL")
            If Logon(session, "001", "DEVELOPER", "developer1", "EN") Then
                ReadTable(session)
            End If
        Catch exc As COMException
            Console.WriteLine(exc.ToString())
        End Try
        Console.WriteLine("press any key to exit")
        Console.ReadLine()
    End Sub

End Module
