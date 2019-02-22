# Office-Addin-Test-Infrastructure

Provides a framework for testing Office Taskpane Add-ins by allowing Add-ins to send results to a test server.  The results can then be consumed and used
by tests to validate that the Add-in is working as expected.

## Command-Line Interface
* [start](#start)

### start
Start the test server. 

Syntax:

`office-addin-test-server start [options]`

Options:

`-p [port]`<br>
`--port [port]`<br>
    Port number must be between 0 - 65535. If no port specified, port defaults to 8080
#
