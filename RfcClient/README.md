# VB.NET RFC Client Example

This example project shows how to call an ABAP function module (or BAPI) with [RfcConnector](http://rfcconnector.com/)

Calling a function module involves three steps:

1. Connect to the SAP system using a `Session` instance
2. Import the function module prototype
3. Call the function with the desired parameters and receive the results

```vbnet
' create a session instance
Dim session as new NWRfcSession

' configure the connection (using connection data from SAPLogon entry)
session.RfcSystemData.ConnectString = "SAPLOGON_ID=my_system_id"

' fill in credentials
session.LogonData.Client = "000"
session.LogonData.User = "myuser"
session.LogonData.Password = "***"
session.LogonData.Language = "EN"

' connect to the SAP system
session.Connect()
 
' import function module's prototype
Dim fn as FunctionCall
fn = Session.ImportCall("BAPI_FLIGHT_GETLIST", True)
 
' set importing parameters
fn.Importing("AIRLINE").value = "LH"
 
' call the function 
session.CallFunction(fn, True)
 
' process the result
For Each row As RfcFields In fn.Tables("FLIGHT_LIST").Rows
	Console.WriteLine(row("AIRLINE").value + " " + row("FLIGHTDATE").value)
Next
```

For more information, please visit [http://rfcconnector.com/](http://rfcconnector.com/)