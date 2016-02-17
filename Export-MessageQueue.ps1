<#
    .SYNOPSIS
    Export messages from a transport queue to file system for manual replay 
   
   	Thomas Stensitzki
	
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
	Version 1.0, 2016-02-17

    Ideas, comments and suggestions to support@granikos.eu 
 
    .LINK  
    More information can be found at http://www.granikos.eu/en/scripts 
	
    .DESCRIPTION
	
    This script suspends a transport queue, exports the messages to the local file system. After successful export the messages are optionally deleted from the queue.
    
    .NOTES 
    Requirements 
    - Windows Server 2008 R2 SP1, Windows Server 2012 or Windows Server 2012 R2  
    - Utilizes global functions library

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0     Initial community release 
	
	.PARAMETER Queue
    Full name of the transport queue, e.g. SERVER\354
    Use Get-Queue -Server SERVERNAME to identify message queue

    .PARAMETER Path
    Path to folder for exprted messages

    .PARAMETER DeleteAfterExport
    Switch to delete per Exchange Server subfolders and creating new folders

	.EXAMPLE
    Export messages from queue MCMEP01\45534 to D:\ExportedMessages and do not delete messages after export

    .\Export-MessageQueue -Queue MCMEP01\45534 -Path D:\ExportedMessages

    .EXAMPLE
    Export messages from queue MCMEP01\45534 to D:\ExportedMessages and delete messages after export

    .\Export-MessageQueue -Queue MCMEP01\45534 -Path D:\ExportedMessages -DeleteAfterExport

#>
param(
	[parameter(Mandatory=$true,HelpMessage='Transport queue holding messages to be exported (e.g. SERVER\354)')]
		[string] $Queue,
	[parameter(Mandatory=$true,HelpMessage='File path to local folder for exprted messages (e.g. E:\Export)')]
		[string] $Path,
	[parameter(Mandatory=$false,HelpMessage='Target Exchange server to copy the selected receive connector to')]
		[switch] $DeleteAfterExport
)

Set-StrictMode -Version Latest

# Implementation of global module
Import-Module GlobalFunctions
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name
$logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 14
$logger.Write("Script started")
$logger.Write("Working on message queue $($Queue), Export folder: $($Path), DeleteAfterExport: $($DeleteAfterExport)")

### FUNCTIONS -----------------------------

function Request-Choice {
    param([string]$Caption)
    $choices =  [System.Management.Automation.Host.ChoiceDescription[]]@("&Yes","&No")
    [int]$defaultChoice = 1

    $choiceReturn = $Host.UI.PromptForChoice($Caption, "", $choices, $defaultChoice)

    return $choiceReturn   
}

function Check-Folders {
    # Check, if export folder exists
    if(!(Test-Path $Path)) {
        # Folder does not exist, lets create a new root folder
        New-Item -Path $Path -ItemType Directory | Out-Null
        $logger.Write("Folder $($Path) created")
    }
}


function Check-Queue {
    # Check message queue
    $messageCount = -1
    try {
        $messageQueue = Get-Queue $Queue
        $messageCount = $messageQueue.MessageCount
        $logger.Write("$($messageCount) message(s) found in queue $($Queue)")
    }
    catch {
        $logger.Write("Queue $($Queue) cannot be accessed")
    }
    $messageCount
}

function Export-Messages {
    # Export suspended messages
    try {
        # Suspend messages in queue
        $logger.Write("Suspending queue $($Queue)")
        Get-Queue $Queue | Get-Message -ResultSize Unlimited | Suspend-Message -Confirm:$false

        # Fetch suspended messages
        $logger.Write("Fetching suspended messages from queue $($Queue)")
        
        $messages = @(Get-Message $Queue -ResultSite Unlimited | ?{$_.Status -eq "Suspened"} )

        $logger.Write("$($messages.Count) suspended messages fetched from queue $($Queue)")

        # Export fetched messages
        $messages | ForEach-Object {$m++;Export-Message $_.Identity | AssembleMessage -Path (Join-Path "$($m).eml" -Path $Path)}
    }
    catch {}
}

function Delete-Messages {
    # Delete suspended messages from queue
    $logger.Write("Delete  suspended messages from queue $($Queue)")
    Get-Message -Queue $Queue -ResultSize Unlimited | ?{$_.Status -eq "Suspened"} | Remove-Message -WithNDR $false -Confirm:$false 
}


# MAIN ####################################################

# 1. Check export folder
Check-Folders

# 2. Check queue
if((Check-Queue -gt 0)) {
    if((Request-Choice -Caption "Do you want to suspend and export all messages in queue $($Queue)?") -eq 0) {
        # Yes, we want to suspend and delete messages
        Export-Messages
    }
    else {
        # No, we do not want to delete message
        $logger.Write("User choice: Do not suspend and export messages")
    }
    if($DeleteAfterExport) {
        if((Request-Choice -Caption "Do you want to DELETE all suspended messages from queue $($Queue)?") -eq 0) {
            $logger.Write("User choice: DELETE suspended")
            Write-Output "Suspended messages will be deleted WITHOUT sending a NDR!"
            Delete-Messages 
        } 
        else {
            $logger.Write("User choice: DO NOT DELETE suspended")
            Write-Output "Exported messages have NOT been deleted from queue!"
            Write-Output "Remove messages manually and be sure, if you want to send a NDR!"
        }
    }
}
else {
    Write-Output "Queue $($Queue) does not contain any messages"
    $logger.Write("Queue $($Queue) does not contain any messages")
}

$logger.Write("Script finished")
Write-Host "Script finished"