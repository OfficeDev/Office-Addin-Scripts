# Data usage notice

It is important for us to understand how the Office Add-in CLI tools are used so they can be improved. In order to do so, the tools collect anonymized usage data and send it to Microsoft.

For more details on how we use this data and under what circumstances it may be shared, please see the Microsoft privacy statement. https://privacy.microsoft.com/en-us/privacystatement

## Examples of data collected include:
- Date, time and location of project creation.
- The frameworks used in the created project (e.g. React, Angular, etc.)
- The target Office Application (e.g. Excel, Word, PowerPoint, Outlook, etc.)
- Selected language for the project (e.g. TypeScript, JavaScript)
- Exceptions and errors that users hit
- Whether the project was created successfully
- A random GUID of the device used

## Disable data collection
Microsoft uses this data to provide and improve our products, to troubleshoot problems, and to operate our business. 

To disable data collection, run the following command before you use these tools:
```
npx office-addin-usage-data off
```
