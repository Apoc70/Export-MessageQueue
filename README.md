# Get-DailyBackupAlerts.ps1
Export messages from a transport queue to file system for manual replay 

##Description
This script suspends a transport queue, exports the messages to the local file system. After successful export the messages are optionally deleted from the queue.

This script utilizes the GlobalFunctions library https://github.com/Apoc70/GlobalFunctions

##Inputs
.PARAMETER Queue
Full name of the transport queue, e.g. SERVER\354
Use Get-Queue -Server SERVERNAME to identify message queue

.PARAMETER Path
Path to folder for exprted messages

.PARAMETER DeleteAfterExport
Switch to delete per Exchange Server subfolders and creating new folders

##Outputs
None

##Examples
```
.\Export-MessageQueue -Queue MCMEP01\45534 -Path D:\ExportedMessages
```
Export messages from queue MCMEP01\45534 to D:\ExportedMessages and do not delete messages after export

```
.\Export-MessageQueue -Queue MCMEP01\45534 -Path D:\ExportedMessages -DeleteAfterExport
```
Export messages from queue MCMEP01\45534 to D:\ExportedMessages and delete messages after export

##TechNet Gallery
Find the script at TechNet Gallery
* --


##Credits
Written by: Thomas Stensitzki

Find me online:

* My Blog: https://www.granikos.eu/en/justcantgetenough
* Archived Blog:	http://www.sf-tools.net/
* Twitter:	https://twitter.com/apoc70
* LinkedIn:	http://de.linkedin.com/in/thomasstensitzki
* Github:	https://github.com/Apoc70

For more Office 365, Cloud Security and Exchange Server stuff checkout services provided by Granikos

* Blog:     http://blog.granikos.eu/
* Website:	https://www.granikos.eu/en/
* Twitter:	https://twitter.com/granikos_de
