# Slack text width with the formatting we use maxes out ~80 characters...
$Width = 80
$CommandsToExport = @()

function Get-ComodoHelp {
    <#
    .SYNOPSIS
        Display the last message for a ticket
    .EXAMPLE
        !TicketHelp
    #>
    [cmdletbinding()]
    [PoshBot.BotCommand(
        KeepHistory = $false,
        CommandName = 'TicketHelp',
        Aliases = ('ticket')
    )]

    param (
        [parameter(position = 1)][string]$help
    )
    $commands = @(
        @{Name = "TicketLast";       MSG = "Show the last message for a ticket";         Usage = "!TicketLast TicketID"}
        @{Name = "TicketStatus";     MSG = "Show the Status for a ticket (Open/Closed)"; Usage = "!TicketStatus TicketID"}
        @{Name = "TicketAllReplys";  MSG = "Display all messages for a ticket";          Usage = "!TicketAllReplys TicketID"}
        @{Name = "TicketTime";       MSG = "Show how long ticket was open";              Usage = "!TicketTime TicketID"}
        @{Name = "TicketReply";      MSG = "Send a reply to ticket";                     Usage = "!TicketReply TicketID `"Message`" (Message MUST be in QUOTES)"}
        @{Name = "TicketClose";      MSG = "Send a reply and Close a Ticket";            Usage = "!TicketClose TicketID `"Message`" (Message MUST be in QUOTES)"} 
    )
    $o += "Commands availble in this plugin`n"
    if ($help) {
        if ($commands  | Where-Object {$_.Name -like $help}) {
            $Name  = ($commands  | Where-Object {$_.Name -like $help}).Name
            $MSG   = ($commands  | Where-Object {$_.Name -like $help}).MSG
            $usage = ($commands  | Where-Object {$_.Name -like $help}).Usage
            New-PoshBotCardResponse -Type Warning -Text "$Name - $MSG`n  $usage"
        } else {
            New-PoshBotCardResponse -Type Error -Text "$help is not part of this plugin"
        }
    } else {
        foreach ($item in $Commands) {
            $Name = $item.Name.PadRight(15, " ")
            $MSG = $item.MSG
            $o += "  !$Name - $MSG`n"
        }
        $o += "To see usage try !ticket COMMAND"
        New-PoshBotCardResponse -Type Warning -Text $o
    }
    
}
$CommandsToExport += 'Get-ComodoHelp'

function Get-ComodoLastMSG {
    <#
    .SYNOPSIS
        Display the last message for a ticket
    .EXAMPLE
        !TicketLats TicketID
    #>
    [cmdletbinding()]
    [PoshBot.BotCommand(
        CommandName = 'TicketLast',
        Aliases = ('tl')
    )]

    param (
        [PoshBot.FromConfig('ComodoAPIHost')][parameter(mandatory)][string]$APIHost,
        [PoshBot.FromConfig('ComodoAPIKey')][parameter(mandatory)][string]$APIKey,
        [parameter(position = 1,mandatory)][string]$TicketID
    )
    # HTTPS Workaround
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    try {
        #check if $TicketID is Number
        0 + $x | Out-Null
        $request    = Invoke-WebRequest -UseBasicParsing -Uri "https://$APIHost/clientapi/index.phpWhere-ObjectserviceName=viewticket" -Headers @{"Authorization" = $APIKey; } -ContentType "application/json" -Method Post -Body "{`"ticketId`": `"$TicketID`"}"
	    $ticket     = $($Request.Content | ConvertFrom-Json).data
        $ticket.threads = $ticket.threads | Where-Object {$_.body -notmatch "Ticket closed (.*)|Stage `".*`" completed|Ticket claimed.*"} | Where-Object {$_.title -notmatch "New Collaborator .*|New Ticket|Ticket Marked Overdue"}
        $replyCount = $ticket.threads.count
        $Message = $ticket.threads[$replyCount-1]
        $replyDate = $Message.created
        $replyFrom = $Message.poster
        #Remove some HTML
        $Message = $message.body.Replace("&nbsp;","")
		$Message = $Message.replace("<br>", "`n")  -replace '<[^>]+>',''
		#shorten to 1024 Chars and add ... if it was truncated
		if ($Message.Length -gt 1024) {
			$Message = $Message.subString(0, 1023)
			$Message = "$Message ....."
		}
        $fields = [pscustomobject]@{
            "Subject" = $ticket.subject
            "Date" = $replyDate
            "From" = $replyFrom
            "Message" = $Message
        }
        New-PoshBotCardResponse -Type Normal -Title "Last Reply from #$TicketID" -Text ($fields | Format-List -Property * | Out-String)
    }
    catch {
        New-PoshBotCardResponse -Type Error -Text "Failed to get Ticket Details for #$TicketID`n"
    }
}
$CommandsToExport += 'Get-ComodoLastMSG'

function Get-ComodoStatus {
    <#
    .SYNOPSIS
        Display the Status a ticket
    .EXAMPLE
        !TicketStatus TicketID
    #>
    [cmdletbinding()]
    [PoshBot.BotCommand(
        CommandName = 'TicketStatus',
        Aliases = ('ts')
    )]
    param (
        [PoshBot.FromConfig('ComodoAPIHost')][parameter(mandatory)][string]$APIHost,
        [PoshBot.FromConfig('ComodoAPIKey')][parameter(mandatory)][string]$APIKey,
        [parameter(position = 1,mandatory)][string]$TicketID
    )
    # HTTPS Workaround
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    try {
        #check if $TicketID is Number
        0 + $x | Out-Null
        $request    = Invoke-WebRequest -UseBasicParsing -Uri "https://$APIHost/clientapi/index.phpWhere-ObjectserviceName=viewticket" -Headers @{"Authorization" = $APIKey; } -ContentType "application/json" -Method Post -Body "{`"ticketId`": `"$TicketID`"}"
	    $ticket     = $($Request.Content | ConvertFrom-Json).data
        $status     = $ticket.status
        New-PoshBotCardResponse -Type Normal -Text "Ticket #$TicketID is $status"
    }
    catch {
        New-PoshBotCardResponse -Type Error -Text "Failed to get Ticket Details for #$TicketID`n"
    }
}
$CommandsToExport += 'Get-ComodoStatus'

function Get-ComodoAllMSG {
    <#
    .SYNOPSIS
        Display All message for a ticket
    .EXAMPLE
        !TicketAllReplys TicketID
    #>
    [cmdletbinding()]
    [PoshBot.BotCommand(
        CommandName = 'TicketAllReplys',
        Aliases = ('ta')
    )]

    param (
        [PoshBot.FromConfig('ComodoAPIHost')][parameter(mandatory)][string]$APIHost,
        [PoshBot.FromConfig('ComodoAPIKey')][parameter(mandatory)][string]$APIKey,
        [parameter(position = 1,mandatory)][string]$TicketID
    )
    # HTTPS Workaround
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    try {
        #check if $TicketID is Number
        0 + $x | Out-Null
        $request    = Invoke-WebRequest -UseBasicParsing -Uri "https://$APIHost/clientapi/index.phpWhere-ObjectserviceName=viewticket" -Headers @{"Authorization" = $APIKey; } -ContentType "application/json" -Method Post -Body "{`"ticketId`": `"$TicketID`"}"
	    $ticket     = $($Request.Content | ConvertFrom-Json).data
        $replyCount = $ticket.threads.count
        $Messages = $ticket.threads
        $o = "Showing $replyCount Messages from #$TicketID`n`n--------------------------------------------------------------------`n"
        foreach ($Message in $Messages) {
            $replyDate = $Message.created
            $replyFrom = $Message.poster
            #Remove some HTML
            $Message = $message.body.Replace("&nbsp;","")
		    $Message = $Message.replace("<br>", "`n")  -replace '<[^>]+>',''
		    #shorten to 1024 Chars and add ... if it was truncated
		    if ($Message.Length -gt 1024) {
			    $Message = $Message.subString(0, 1023)
			    $Message = "$Message ....."
		    }
            $o += "$replyFrom - $replyDate `n$Message`n--------------------------------------------------------------------`n"

        }
        $fromName = $global:PoshBotContext.CallingUserInfo.FullName
        New-PoshBotCardResponse -Type Normal -Text "Ticket Replies sent Directly to $fromName" 
        New-PoshBotCardResponse -Type Normal -Text $o -DM
    }
    catch {
        New-PoshBotCardResponse -Type Error -Text "Failed to get Ticket Details for #$TicketID`n"
    }
}
$CommandsToExport += 'Get-ComodoAllMSG'

function New-ComodoReply {
    <#
    .SYNOPSIS
        Send a reply to a ticket
    .EXAMPLE
        !TicketReply TicketID "MASSAGE"
    #>
    [cmdletbinding()]
    [PoshBot.BotCommand(
        CommandName = 'TicketReply',
        Aliases = ('tr')
    )]

    param (
        [PoshBot.FromConfig('ComodoAPIHost')][parameter(mandatory)][string]$APIHost,
        [PoshBot.FromConfig('ComodoAPIKey')][parameter(mandatory)][string]$APIKey,
        [parameter(position = 1,mandatory)][string]$TicketID,
        [parameter(position = 2,mandatory)][string]$Message
    )
    # HTTPS Workaround
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    try {
        #check if $TicketID is Number
        0 + $x | Out-Null
        $email = $global:PoshBotContext.CallingUserInfo.Email
        $request    = Invoke-WebRequest -UseBasicParsing -Uri "https://$APIHost/clientapi/index.phpWhere-ObjectserviceName=ticketpostreply" -Headers @{"Authorization" = $APIKey; } -ContentType "application/json" -Method Post -Body "{`"ticketId`": `"$TicketID`", `"email`": `"$email`", `"message`": `"$message`"}"
        if ($request.statusCode -ne "200") {
            New-PoshBotCardResponse -Type Error -Text "Failed to Send reply to #$ticketID"
        }
    }
    catch {
         New-PoshBotCardResponse -Type Error -Text "Failed to Send reply to #$ticketID"
    }
}
$CommandsToExport += 'New-ComodoReply'

function Set-ComodoClose {
    <#
    .SYNOPSIS
        Close a ticket
    .EXAMPLE
        !TicketClose TicketID Message (Close Ticket with a Message)
    #>
    [cmdletbinding()]
    [PoshBot.BotCommand(
        CommandName = 'TicketClose',
        Aliases = ('tc')
    )]
    param (
        [PoshBot.FromConfig('ComodoAPIHost')][parameter(mandatory)][string]$APIHost,
        [PoshBot.FromConfig('ComodoAPIKey')][parameter(mandatory)][string]$APIKey,
        [parameter(position = 1,mandatory)][string]$TicketID,
        [parameter(position = 2,mandatory)][string]$Message,
        [switch]$Force
    )
    # HTTPS Workaround
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    try {
        #check if $TicketID is Number
        0 + $x | Out-Null
        $email = $global:PoshBotContext.CallingUserInfo.Email
        #get ticket status
        $request    = Invoke-WebRequest -UseBasicParsing -Uri "https://$APIHost/clientapi/index.phpWhere-ObjectserviceName=viewticket" -Headers @{"Authorization" = $APIKey; } -ContentType "application/json" -Method Post -Body "{`"ticketId`": `"$TicketID`"}"
	    $ticket     = $($Request.Content | ConvertFrom-Json).data
        if (($ticket.status -eq "open") -or ($Force)) {
            $request2    = Invoke-WebRequest -UseBasicParsing -Uri "https://$APIHost/clientapi/index.phpWhere-ObjectserviceName=ticketpostreply" -Headers @{"Authorization" = $APIKey; } -ContentType "application/json" -Method Post -Body "{`"ticketId`": `"$TicketID`", `"email`": `"$email`", `"message`": `"$message`"}"
            if ($request2.statusCode -ne "200") {
                New-PoshBotCardResponse -Type Error -Text "Failed to Send reply to #$ticketID"
            } else {
                $request3    = Invoke-WebRequest -UseBasicParsing -Uri "https://$APIHost/clientapi/index.phpWhere-ObjectserviceName=closeTicket" -Headers @{"Authorization" = $APIKey; } -ContentType "application/json" -Method Post -Body "{`"ticketId`": `"$TicketID`"}"
                if ($request3.statusCode -ne "200") {
                    New-PoshBotCardResponse -Type Error -Text "Failed to Close #$ticketID"
                } else {
                    New-PoshBotCardResponse -Type Normal -Text "Ticket #$TicketID Reply sent and Closed"
                }
            }
        } elseif ($ticket.status -eq "closed") {
            New-PoshBotCardResponse -Type Normal -Text "Ticket #$TicketID is already Closed, Reply not posted. you can use -force to post the reply reguardless"
        } else {
            New-PoshBotCardResponse -Type Error -Text "Failed to get #$ticketID Status" 
        }
    }
    catch {
         New-PoshBotCardResponse -Type Error -Text "Failed to Close #$ticketID"
    }
}
$CommandsToExport += 'Set-ComodoClose'

function Get-ComodoTime {
    <#
    .SYNOPSIS
        Find out how long a ticket was open
    .EXAMPLE
        !TicketTime TicketID
    #>
    [cmdletbinding()]
    [PoshBot.BotCommand(
        CommandName = 'TicketTime',
        Aliases = ('tt')
    )]

    param (
        [PoshBot.FromConfig('ComodoAPIHost')][parameter(mandatory)][string]$APIHost,
        [PoshBot.FromConfig('ComodoAPIKey')][parameter(mandatory)][string]$APIKey,
        [parameter(position = 1,mandatory)][string]$TicketID
    )
    # HTTPS Workaround
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    try {
        #check if $TicketID is Number
        0 + $x | Out-Null
        $request    = Invoke-WebRequest -UseBasicParsing -Uri "https://$APIHost/clientapi/index.phpWhere-ObjectserviceName=viewticket" -Headers @{"Authorization" = $APIKey; } -ContentType "application/json" -Method Post -Body "{`"ticketId`": `"$TicketID`"}"
	    $ticket     = $($Request.Content | ConvertFrom-Json).data
        if ($ticket.status -ne "closed") {
            New-PoshBotCardResponse -Type Error -Text "#$ticketID has not been closed yet"
        } else {
            #####
            #get ticket details
            $request    = Invoke-WebRequest -UseBasicParsing -Uri "https://$APIHost/clientapi/index.phpWhere-ObjectserviceName=viewticket" -Headers @{"Authorization" = $APIKey; } -ContentType "application/json" -Method Post -Body "{`"ticketId`": `"$TicketID`"}"
            $ticket     = $($Request.Content | ConvertFrom-Json).data
            $Messages = $ticket.threads
            $Opened = Get-Date($Messages[0].Created)
            $closed = Get-Date($($messages | Where-Object {$_.title -eq "Ticket Closed"}).Created)
            $TotalTime = $Closed - $Opened
            $Days = $TotalTime.Days
            $Hours = $TotalTime.Hours
            $Minutes = $TotalTime.Minutes
            $Time += "$Days Days, $Hours Hours, and $Minutes Minutes"
            New-PoshBotCardResponse -Type Normal -Text "#$TicketID Was open for $time"
        }
    }
    catch {
         New-PoshBotCardResponse -Type Error -Text "Failed get information for #$ticketID"
    }
}
$CommandsToExport += 'Get-ComodoTime'

function Get-ComodoMemUsage {
    <#
    .SYNOPSIS
        Find out how long a ticket was open
    .EXAMPLE
        !TicketTime TicketID
    #>
    [cmdletbinding()]
    [PoshBot.BotCommand(
        CommandName = 'TicketMemory',
        Aliases = ('tm')
    )]

    param (
        [parameter(position = 1)][string]$ComputerName
    )
    if (!$ComputerName) {
        $ComputerName = $env:COMPUTERNAME
    }
    try {
        $powershell = get-process -ComputerName $ComputerName | Where-Object {$_.ProcessName -eq "powershell"}
        $powershell | ForEach-Object {$total += $_.ws}
        $total = [math]::Round($($total/1024/1024),0)
        New-PoshBotCardResponse -Type Normal -Text "Comodo Memory Usage $total MB"
    }
    catch {
        New-PoshBotCardResponse -Type Error -Text "Unable to get memory usage"
    }
}
$CommandsToExport += 'Get-ComodoMemUsage'

Export-ModuleMember -Function $CommandsToExport