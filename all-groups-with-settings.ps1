<#
	Description:
	Script to pull all the groupnames and associated settings
	Assumes GAM is installed and in the PATH
	
	Author:
	matthew.raven@edmonton.ca
	
	Last Modified:
	Feb 21, 2014
#>

#Set these to 1 for writing new files, set to 0 for testing only to screen
$writefile = 1
#generateNewList -> write out the groups list again. Only has to be done once per day
$generateNewList = 1
#Whether or not to send an email
$sendEmail = 1

#where to write the reports, needs trailing slash
$reportDir = "e:\ScriptOutput\"
#where gam is located, needs trailing slash (default is "" if gam is in PATH)
$gamPath = "E:\gam3\"

if ($generateNewList)
{
	iex "$($gamPath)gam print groups > $($reportDir)groups_$(get-date -f MM-dd).csv"
}
$list = Import-Csv "$($reportDir)groups_$(get-date -f MM-dd).csv"

$fields = @(
	"name",
	"adminCreated",
	"directMembersCount",
	"email",
	"description",
	"allowExternalMembers",
	"whoCanJoin",
	"whoCanViewMembership",
	"defaultMessageDenyNotificationText",
	"includeInGlobalAddressList",
	"archiveOnly",
	"isArchived",
	"membersCanPostAsTheGroup",
	"allowWebPosting",
	"messageModerationLevel",
	"replyTo",
	"customReplyTo",
	"sendMessageDenyNotification",
	"whoCanContactOwner",
	"messageDisplayFont",
	"whoCanLeaveGroup",
	"whoCanPostMessage",
	"whoCanInvite",
	"spamModerationLevel",
	"whoCanViewGroup",
	"showInGroupDirectory",
	"maxMessageBytes",
	"allowGoogleCommunication" )

$fieldsString = ""
$hashTable =@{}

foreach($i in $fields)
	{
		$fieldsString =  $fieldsString + $($i) +","
	}
$fieldsString = $fieldsString.TrimEnd(",")

if ($writefile)
	{
		$filename = "$($reportDir)groups-with-settings-$(get-date -f yyyy-MM-dd_hhmmss).csv"
		$stream = [System.IO.StreamWriter] "$($filename)"
		$stream.WriteLine($($fieldsString))
	}

foreach ($entry in $list)
  {
	"$($entry.Email)"
	
	$email = $($entry.Email) -replace "\'","``'"
	$comp = iex "$($gamPath)gam info group $($email)"
	if ($comp -ne $null) {$comp_branch = $comp.Split("`n")}
	
	$hashTable.Clear()
	
	foreach($i in $comp_branch)
	{
		foreach($x in $fields)
		{	
			$i = $i.TrimStart(" ")
			$a = $i -match $("^$($x)")
			if ($a)
			{
					$i = $i -replace "$($x)",""
					$i = $i -replace "\'","``'"
					$i = $i -replace "\,"," "
					$i = $i.Trim(": ")
					#"$($x) is $($i) for $($entry.Email)"
					$hashTable.Add( $x , $i )
			}
		}
	}

	if ($writefile)
	{
		$hashTableString = ""
		foreach($x in $fields)
		{	
			$hashTableString = $hashTableString + $hashTable.Get_Item($x) +","
		}
		$hashTableString = $hashTableString.TrimEnd(",")
		#"Writing: $($hashTableString)"
		$stream.WriteLine("$($hashTableString)")
		$stream.Flush()
	}
  
 }
if ($writefile)
{
	$stream.close()
}

if ($sendEmail)
{
	#now mail it out
	$smtpServer = "mailserver.testdomain.com"
	$att = new-object Net.Mail.Attachment($filename)
	$msg = new-object Net.Mail.MailMessage
	$smtp = new-object Net.Mail.SmtpClient($smtpServer)
	$msg.From = "noreply@testdomain.com"
	$msg.To.Add("username@testdomain.com")
	$msg.Subject = "List of all groups and associated settings"
	$msg.Body = "Attached is the groups settings report"
	$msg.Attachments.Add($att)
	$smtp.Send($msg)
	$att.Dispose()
}
