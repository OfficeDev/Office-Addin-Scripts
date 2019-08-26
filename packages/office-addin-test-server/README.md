# Office-Addin-Test-Server

This package provides a framework for testing Office task pane add-ins by allowing add-ins to send results to a test server. The results can then be consumed and used by tests to validate that the add-in is working as expected.

## Command-Line Interface
* [start](#start)

### start
Start the test server. 

Syntax:

`office-addin-test-server start [options]`

Options:

`-p [port]`<br>
`--port [port]`<br>

Port number must be between 0 - 65535. If no port specified, port defaults to 4201
#
