# msg2eml

This tiny tool allows you convert Outlook MSG files to EML files which can be read by standard mail clients.
Usage:

`msg2eml <input.msg> [output.eml]`

If output.eml is not specified, it will be created in the same directory as input.msg with the same name.
Examples:

```
msg2eml Testmail.msg
msg2eml C:\Temp\Testmail.msg C:\Output.eml
```
