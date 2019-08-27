# Office-Addin-Usage-Data
This package allows for reporting usage data events and exception data to the selected telemetry infrastructure (e.g. ApplicationInsights).

# Privacy
The Office Addin-Usage-Data package collects anonymized usage data and sends it to Microsoft. For more details on how we use this data and under what circumstances it may be shared, 
please see the [Microsoft privacy statement](https://privacy.microsoft.com/en-us/privacystatement).

The package collects:
* Usage data about operations performed.
* Exception call stacks to help diagnose issues.


## Command-Line Interface
* [List](#List)
* [Off](#Off)
* [On](#On)

### List
Display the current usage data settings.

Syntax:

`list`

#

### Off
Sets the usage data level to Off(sending no usage data data).

Syntax:

`off`

#

### On
Sets the usage data level to On(sending usage and exception data).

Syntax:

`on`

#