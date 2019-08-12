# Office-Addin-telemetry
This package allows for sending telemetry event and exception data to the selected telemetry infrastructure (e.g. ApplicationInsights).


## Command-Line Interface
* [List](#List)
* [Off](#Off)
* [On](#On)
* [Privacy](#Privacy)

### List
Display the current telemetry settings.

Syntax:

`list`

#

### Off
Sets the telemetry level to Basic.

Syntax:

`off`

#

### On
Sets the telemetry level to Verbose.

Syntax:

`on`

#

### Privacy
The Office Addin-Telemetery package collects anonymized usage data and sends it to Microsoft. For more details on how we use this data and under what circumstances it may be shared, 
please see the Microsoft privacy statement.

The package collects:
* Usage data about operations performed.
* Exception call stacks to help diagnose issues.