# VB.NET RFC Server Example

Implementing an RFC server which can receive function calls from SAP requires the following steps:

1. Create a server object and set the necessary parameters: `GatewayHost`, `GatewayService`, `ProgramID`
2. Create or import one or more `FunctionCall` prototypes and install them using the `InstallFunction` method
3. Subscribe to the `IncomingCall` event
4. Run the server and wait for incoming calls. The server will run in another thread, so the main thread 
   can do other work, for example handling GUI updates.

```vbnet
' 
' set up the server and starts it
'
Sub Main
	' import the function definition(s)
    ImportFunctions()
    ' connection data for SAP server we are registering with
    srv.ProgramID = RFC_DESTINATION
    srv.GatewayHost = GATEWAY_HOST
    srv.GatewayService = GATEWAY_SERVICE

    ' install the function(s) we want to handle
    srv.InstallFunction(fn1)

    ' Start the server
    srv.Serve()
    ' wait until key is pressed
    Console.WriteLine("Server running. Press key to stop...")
    Console.ReadLine()

    ' then stop the server
    srv.Shutdown()

End Sub

'
' OnIncoming call is called by the server every time we receive
' a function call. We fill in the required data and send it back
' to the caller
'     
Private Sub srv_IncomingCall(ByVal fn As FunctionCall) Handles srv.IncomingCall
	Select Case fn.Function
		Case "BAPI_FLIGHT_GETLIST"			' add some static rows 
			Dim r As RfcFields, dt As DateTime
            dt = DateTime.Today
			For i = 1 To 10
				r = fn.Tables("FLIGHT_LIST").Rows.AddRow()
				r("AIRLINEID").value = "LH"
                r("AIRPORTFR").value = "LAX"
                r("AIRPORTTO").value = "FRA"
                r("CONNECTID").value = i
                r("FLIGHTDATE").value = dt.AddDays(i)
			Next
		Case Else
			' we don't know about this function, so we raise an exception
			fn.RaiseException("SYSTEM_ERROR")
	End Select
End Sub
```

# Prerequisites

The RFC connection must be created in the SAP server using transaction `SM59`:

![RFC Connection Setting / SM59](./docs/SM59.png)

After defining the connection, start the server and use the "Connection Test"
button to test the connection.

To actually connect to your server and retrieve data, you have to write an 
ABAP report which calls the service. For `DESTINATION` use the name of the 
RFC destination from SM59:

```abap
REPORT  zrfcclient.

DATA: lt_flights TYPE TABLE OF bapisfldat,
      wa TYPE bapisfldat.

CALL FUNCTION 'BAPI_FLIGHT_GETLIST' DESTINATION 'ZRFCCTEST'
  EXPORTING
    airline     = 'LH'
  TABLES
    flight_list = lt_flights
  EXCEPTIONS
    OTHERS      = 1.

WRITE:/ 'RESULT', sy-subrc.

LOOP AT lt_flights INTO wa.
  WRITE:/ wa-airlineid, wa-connectid, wa-flightdate.

ENDLOOP.
```
