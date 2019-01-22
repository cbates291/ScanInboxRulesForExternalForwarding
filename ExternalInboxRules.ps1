#Script Created by: Chris Bates
#Version 1.2
#Date Last Updated: 1-11-18

#Setup O365 Connection
$SetPath = "C:\!temp" #Path Used to Store Files, etc.
$MSOLCred = IMPORT-CLIXML "$($SetPath)\MSOLCred.xml"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $MSOLCred -Authentication Basic -AllowRedirection
Import-PSSession $Session -AllowClobber
Connect-MsolService –Credential $MSOLCred

#Gather the current $WarningPreference setting so we can return later
$prefBackup = $WarningPreference

<#Set the WarningPreference to Silently Continue. The reason for this is that
in this script there can be "broken" inbox rules and they give a warning
we simply don't care about those broken rules and are wanting to omit
them from the output#>
$WarningPreference = 'Silently Continue'

#Get current date
$fulldateandtime = get-date -Format "MM-dd-yyyy  hh-mm tt dddd"

#Creates file that will be made and attached to email
$attachment = "C:\MailboxRules$fulldateandtime.csv"

#Sets Logfile path
$logfilename = "$SetPath\MailboxRuleForward.txt"

#Gather all Mailboxes
$AllMailboxes = Get-Mailbox -ResultSize Unlimited

#Sets Threshold for O365 Reconnect
$reconnectThreshold = 1000

#Sets Initial Count
$processedCount = 0

<#Parse the mailboxes and search their inbox rules for ForwardTo, ForwardAsAttachmentTo, and RedirectTo.
We are ensuring that each option is not NULL and does not include "cn=Recipients"(this is so that we don't
pull rules that include users in our organization)#>
foreach ($mailbox in $AllMailboxes) {

    # Reconnect if threshold is reached
    if($processedCount -ge $reconnectThreshold)
    {

        # Creates New Session
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://outlook.office365.com/powershell-liveid/" -Credential $MSOLCred -Authentication Basic -AllowRedirection

        # Start a new session
        Import-PSSession $Session -AllowClobber

        # Reset processed counter
        $processedCount = 0
    }
	<#Parse the mailboxes and search their inbox rules for ForwardTo, ForwardAsAttachmentTo, and RedirectTo.
    We are ensuring that each option is not NULL and does not include "cn=Recipients"(this is so that we don't
    pull rules that include users in our organization)#>
    Get-InboxRule -Mailbox $mailbox.DistinguishedName | where {$_.ForwardTo -ne $null -and $_.ForwardTo -inotlike "*cn=Recipients*" -or $_.ForwardAsAttachmentTo -ne $null -and $_.ForwardAsAttachmentTo -inotlike "*Recipients*" -or $_.RedirectTo -ne $null -and $_.RedirectTo -inotlike "*Recipients*" } | Select-Object Identity,Description,ForwardTo,ForwardAsAttachmentTo,RedirectTo | Export-Csv $attachment -append
    $processedCount++
	
}

#Set the WarningPreference Back to what it was before the script
$WarningPreference = $prefBackup

#generates email to user using .net smtpclient to notify them who has client Mailbox Rule forwards.
		         $emailFrom = "no-reply@domain.com"
		         $emailTo = "user@domain.com"
		         $subject = "Mailbox Rule Forwards Externally Report"
				 
##########################################################
##########################################################
##### Start of Email #####################################
##########################################################
##########################################################
				 $body = @"
				 <html>

<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=Generator content="Microsoft Word 14 (filtered)">
<style>
<!--
 /* Font Definitions */
 @font-face
	{font-family:Calibri;
	panose-1:2 15 5 2 2 2 4 3 2 4;}
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
	{margin:0in;
	margin-bottom:.0001pt;
	font-size:11.0pt;
	font-family:"Calibri","sans-serif";}
a:link, span.MsoHyperlink
	{color:blue;
	text-decoration:underline;}
a:visited, span.MsoHyperlinkFollowed
	{color:purple;
	text-decoration:underline;}
.MsoChpDefault
	{font-family:"Calibri","sans-serif";}
.MsoPapDefault
	{margin-bottom:10.0pt;
	line-height:115%;}
@page WordSection1
	{size:8.5in 11.0in;
	margin:1.0in 1.0in 1.0in 1.0in;}
div.WordSection1
	{page:WordSection1;}
-->
</style>

</head>

<body lang=EN-US link=blue vlink=purple>

<div class=WordSection1>

<p class=MsoNormal style='margin-left:.5in'>Hello,<br>
<br>
Attached you will find any mailboxes that had a mailbox rule forward setup to go to an external source. Please review and process as needed.</p>
<br>

<p class=MsoNormal style='margin-left:.5in'>&nbsp;</p>


<p class=MsoNormal><span style='color:#1F497D'>&nbsp;</span></p>

<p class=MsoNormal>&nbsp;</p>

</div>

</body>

</html>
"@
##########################################################
##########################################################
##### End of Email #######################################
##########################################################
##########################################################
				 $smtpServer = "smtp.netjets.com"
		         $smtp = new-object Net.Mail.SmtpClient($smtpServer)
		         Send-MailMessage -SmtpServer $smtpServer -To $emailTo -From $emailFrom -Attachments $attachment -Subject $subject -Body $body -BodyAsHtml
				 Add-Content $logfilename $fulldateandtime
