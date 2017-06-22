# VB.NET TableReader Example

This example project shows how to use the `TableReader` component of [RfcConnector](http://rfcconnector.com/)
to read data from a SAP table.

Reading the contents of an SAP table involves the following steps:

1. Connect to the SAP system using a `Session` instance
2. Create a `TableReader` instance
3. Set query parameters (optional) and retrieve the data

```vbnet
' create a session instance
Dim session As New NWRfcSession

' configure the connection (using connection data from SAPLogon entry)
session.RfcSystemData.ConnectString = "SAPLOGON_ID=my_system_id"

' fill in credentials
session.LogonData.Client = "000"
session.LogonData.User = "myuser"
session.LogonData.Password = "***"
session.LogonData.Language = "EN"

' connect to the SAP system
session.Connect()
 
' create `TableReader` instance
TableReader tr = session.GetTableReader("SFLIGHT")

' set up query parameters
tr.Query.Add("CARRID EQ 'LH'");

' read rows
tr.Read(0, 0);

' process the result
For Each row As RfcFields In tr.Rows
	Console.WriteLine(row("CONNID").value + " " + row("FLDATE").value + " " + row("CARRID").value)
Next
```

For more information, please visit [http://rfcconnector.com/](http://rfcconnector.com/)