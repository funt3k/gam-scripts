#where to write the reports, needs trailing slash
$reportDir = "e:\ScriptOutputTest\"
#where gam is located, needs trailing slash (default is "" if gam is in PATH)
$gamPath = "E:\gam3_edmonton.ca\"

<#
	FLAGS FLAGS FLAGS FLAGS
#>
$sendemailflag = 0
#flag to print the groups in the email (not implemented)
$printgroups = 1
#flag to add the reviewer as co-manager of solely owned groups
$changegroupowner = 0
#flag whether or not to change the ownership
$changecalendarowner = 0
#flag whether or not to print the calendars
$printcalendarowner = 1
#to delegate to reviewer
$delegatemailtoreviewer = 0
#Set Auto-vacation responder?
$setvacationresponder = 0
#Should we reset the password?
$resetpassword = 0 


#Calendar vars
$calendars = @()
$ownscalendars = @()
$allcalendars = @{}
$calendarmessage = ""

#groups vars
$gmember = @()
$groupsowners = @{}
$groupsemailname = @{}

Function SendEmail ($toAddy, $fromAddy, $subject, $messageBody)
{
	"`n`nTO: $toAddy`nFROM: $fromAddy`nSUBJECT: $subject`nBODY:`n$messageBody"
	#now mail it out
	$smtpServer = "unixmail.gov.edmonton.ab.ca"
	#$att = new-object Net.Mail.Attachment($filename)
	$msg = new-object Net.Mail.MailMessage
	$smtp = new-object Net.Mail.SmtpClient($smtpServer)
	$msg.From = $($fromAddy)
	if($toAddy -is [array]){
		foreach($address in $toAddy){
			$msg.To.Add($address)
		}
	} else {
		$msg.To.Add($toAddy)
	}
	$msg.Subject = $subject
	$msg.IsBodyHTML = $true

	$msg.Body += $messageBody
	#$msg.Attachments.Add($att)
	$smtp.Send($msg)
	#$att.Dispose()
}

#GET initial input
$name = Read-Host 'What is the username to be offboarded (first.last)?'
$supername = Read-Host 'Who is the supervisor (first.last)?'
$submitter = Read-Host 'Your Google Apps Admin Username to receive a copy of finished report (first.last)?'
$testrun = Read-Host 'Commit changes? (y/N)'
$sendemail = Read-Host 'Send email to reviewer? (y/N)'

if($testrun.toString() -eq "y")
{
	"Answered YES to making changes"
	$changecalendarowner = 1
	$changegroupowner = 1
	$delegatemailtoreviewer = 1
	$setvacationresponder = 1
	$resetpassword = 1
}
else
{
	"Answered NO to making changes"
}


if($sendemail.toString() -eq "y")
{
	"Answered YES to sending email to reviewer"
	$sendemailflag = 1
}
else
{
	"Answered NO to sending email to reviewer"
}

#CALENDAR parsing
$gamoutput = iex "$($gamPath)gam user $($name) show calendars"
if ($gamoutput -ne $null) 
{
	$gamoutputlines = $gamoutput.Split("`n")
}

$switch = $true
$calendarname = ""

#get a list of all calendars the user owns
foreach($i in $gamoutputlines)
{
	#$owner = $false

	$a = $i -match "^  Name: "
	if ($a)
	{
		$i = $i -replace "^  Name: ",""
		$calendarname = [String] $i
	}
	
	$a = $i -match "^  Summary: "
	if ($a)
	{
		$i = $i -replace "^  Summary: ",""
		$summary = [String] $i
	}
	
	$a = $i -match "^    Access Level: "
	if ($a)
	{
		$i = $i -replace "^    Access Level: ",""
		if ($i -eq "owner")
		{
			"owner found! for $calendarname"
			$calendars += $calendarname
			$allcalendars.Add( $calendarname, $summary )
		} else {
			"not an owner of $calendarname"
		}	
	}
}


#go through each calendar the user owns, and see if they are the sole owner
#also check to see if this calendar can/should be reassigned
#basically, we only want to reassign non-user calendars (so not @edmonton.ca)
#and we *do* want to move CoE calendars (so starting with edmonton.ca_)

$calendarowners = @()
foreach($i in $calendars){
	
	$calendarowners.Clear()
	$gamoutput = ""
	if($i -match "^edmonton.ca_"){
		"`n`nExamining $i"
		$gamoutput = iex "$($gamPath)gam calendar $($i) showacl"
	} else {
		"Skipping $i"
	}
	if ($gamoutput -ne "") 
	{
		$gamoutputlines = $gamoutput.Split("`n")
	}
	
	foreach($x in $gamoutputlines){
		#"examining: $x"
		$a = $x -match "^  Scope user - "
		if ($a)
		{
			$x = $x -replace "^  Scope user - ",""
			$username = [String] $x
		}
		$a = $x -match "^  Role: "
		if ($a)
		{
			$x = $x -replace "^  Role: ",""
			if ($x -eq "owner")
			{
				"Found owner $username of $i"
				$calendarowners += $username
			} else {
				#"not an owner of $username"
			}	
		}
	}
	if ($calendarowners.length -eq 2)
	{
		$ownscalendars += $i
	}
}

#add the user's own calendar. We'll want to do that regardless as the manager will want to review the events
$ownscalendars += $name + "`@edmonton.ca"

"`ncalendars uniquely owned are: $ownscalendars`n"
"all calendars owned: $calendars"
"`n`nAssigning Calendars to Users"

$calendarmessage += "<H3>Calendar Information</H3>"
#now to reassign the calendars to the supervisor
foreach($i in $ownscalendars)
{
	$gamoutput = "NO Change Made - Testing"
	if ($changecalendarowner){
		$gamoutput = iex "$($gamPath)gam calendar $($i) add owner $supername 2>&1"
	}
	"tried to add $supername as owner of $i  - $($allcalendars.Get_Item($i.toString()).toString()) `nresult was: $gamoutput"
	$calendarmessage += "Gave $supername manager access to the calendar: $($allcalendars.Get_Item($i.toString()).toString())<br />"
}

<#
--------------------------------------------------------------
GROUP message start
--------------------------------------------------------------
#>

$gamoutput = iex "$($gamPath)gam info user $($name) 2>&1"

if ($gamoutput -match "Error 404: Resource Not Found: userKey")
{
	"Username not found!"
	break
}

if ($gamoutput -ne $null) 
{
	$gamoutputlines = $gamoutput.Split("`n")
}

foreach($i in $gamoutputlines)
{
	$tokens = $i.split("`<")
	if($tokens[0] -eq "Groups:")
	{
		$gsection = 1
	}
	
	if($gsection -and ($tokens[1] -ne $null))
	{
		$gmember += $tokens[1].trim("`>")
	}
}

foreach($x in $gmember)
{
	Write-Host `nWorking on $($x)
	$comp_branch = @()
	
	$tempgrp = $x -replace "`@edmonton.ca$"
	
	#"Command: $($gamPath)gam info group $($tempgrp)"

	$continue = $true
	$counter = 0
	while($continue)
	{
		Start-Sleep -s $counter
		$process = New-Object System.Diagnostics.Process
		$process.StartInfo.Arguments = "info group $($tempgrp)"
		$process.StartInfo.UseShellExecute = $false
		$process.StartInfo.RedirectStandardOutput = $true
		$process.StartInfo.RedirectStandardError = $true
		$process.StartInfo.CreateNoWindow = $true
		$process.StartInfo.WorkingDirectory = $WORKING_DIRECTORY
		$process.StartInfo.FileName = "$($gamPath)gam"
		$started = $process.Start()
		$comp = $process.StandardOutput.ReadToEnd()
		$errors = $process.StandardError.ReadToEnd()
		$process.WaitForExit()
		#Write-Host Output - $comp
		if ([String] $error -ne ""){
			#Write-Host Error - $errors
		}
		if($errors -notmatch "Error 403: Serving Limit Exceeded - serviceLimit")
		{
			$continue = $false
		}
		if ($counter -eq 20)
		{
			$continue = $false
			Write-Host "Hit max retries, skipping group $x."
		}
		$counter++
		if($continue)
		{
			Write-Host "Error encountered, retrying in $counter seconds (max 20)."
		}
	
	}
	
	if ($comp -ne $null) {
		$comp_branch = $comp.Split("`n")
	}
	
	$groupowners = @()
	$groupname = ""
	
	foreach($i in $comp_branch)
	{
		#Search for groupname
		$a = $i -match "^ name: "
		if ($a)
		{
			$i = $i -replace "^ name: ",""
			$groupname = [String] $i
		}
		
		#Search for owners
		$i = $i -replace " ",""
		
		$a = $i -match "^owner:"
		if ($a)
		{
			$i = $i -replace "^owner:",""
			$i = $i -replace "\(user\)",""
			$i = $i -replace "\(group\)",""
			$groupowners += [String] $i
		}
		
		#Search for managers
		$a = $i -match "^manager:"
		if ($a)
		{
			$i = $i -replace "^manager:",""
			$i = $i -replace "\(user\)",""
			$i = $i -replace "\(group\)",""
			$groupowners += [String] $i
		}
	}
	
	$groupsowners.Add( $x, $groupowners )
	$groupname = $groupname -replace ",","-"
	"$x is called $groupname"
	$groupsemailname.Add( $x, $groupname )
	"Owners/Managers: $($groupsowners.Get_Item($x))"
}
#we now have a list of all groups the user is in, plus the owners of those groups stored in $groupsowners
#$groupsemailname stores the human name and the email address

#send email to all the group owners, commented out because of FOIP concerns
<#
foreach ($h in $groupsowners.GetEnumerator())
{
	$subject = "Group Update: User $name is Being Offboarded"
	$messageBody = "As the owner of a group, you a being notified that the user in the subject line is being terminated. Please either remove the user from your group, and find a possible replacement group member for them. `nUser being removed: $name`nGroup you own: $($h.Name)"
	$toAddy = $h.Value
	SendEmail $toAddy "crsitgoogleoffboarding@edmonton.ca" $subject $messageBody
}
#>

#generate the list of groups that the user is sole owner/manager of
#and generate a list of groups that the user is co-owner of
#and generate a list of non-CoE groups the user is a member of

$solemanager = @{}
$sharedmanager = @{}
$noncoemember = @{}

$fullemail = $name +"`@edmonton.ca"

foreach ($h in $groupsowners.GetEnumerator())
{
	"`nGroup $($h.Name) has $($h.Value.length) owners/managers"
	if ($h.Value.length -gt 1)
	{
		foreach($x in $h.Value)
		{
			if([String] $x -match [String] $fullemail)
			{
				$sharedmanager.Add( $h.Name , $groupsemailname.Get_Item($h.Name))
				"SHARED: $x matches $fullemail"
			} else {
				"SHARED: $x does not match $fullemail"
			}
		}
	}
	
	if ($h.Value.length -eq 1)
	{
		foreach($x in $h.Value)
		{
			if([String] $x -match [String] $fullemail)
			{
				$solemanager.Add( $h.Name , $groupsemailname.Get_Item($h.Name))
				"SOLE: $x matches $fullemail"
			} else {
				"SOLE: $x does not match $fullemail"
			}
		}
	}
	
	foreach($x in $h.Name)
	{
		if($x -notmatch "@edmonton.ca")
		{
			$noncoemember.Add( $h.Name , $groupsemailname.Get_Item($h.Name) )
			"NONCOE: $x does not exist in COE"
		} else {
			"NONCOE: Apparently $x has an @edmonton.ca address"
		}
		
	}
}

#Add the reviewer as co-manager of the groups that the user uniquely owns.
if($changegroupowner)
{	
	"Let's change some groups permissions!"
	foreach ($h in $solemanager.GetEnumerator() )
	{
		$gamoutput = iex "$($gamPath)gam update group $($h.Name) add manager $($supername)"
		"Adding $($supername) as manager of $($h.Name). Gamoutput: $($gamoutput)"
		$gamoutput = iex "$($gamPath)gam update group $($h.Name) update manager $($supername)"
		"Updated $($supername) as manager of $($h.Name). Gamoutput: $($gamoutput)"
	}
}

if($delegatemailtoreviewer)
{	
	"`nStarting delegation step"
	$gamoutput = iex "$($gamPath)gam user $($name) delegate to $($supername)"
	"Adding $($supername) as a delegate of $($h.Name). Gamoutput: $($gamoutput)"
}

if($setvacationresponder)
{
	"`nStarting vacation reminder step"
	$gamoutput = iex "$($gamPath)gam user $($name) vacation on subject `"$($name)`@edmonton.ca Is Unavailable - Please Send Email to Alternate Address`" message `"Thank you for your message. I am no longer montoring this email account. Any inquiries should be forwarded to $($supername)`@edmonton.ca. \n\nThank-you!`""
	"Adding vacation reminder for $($name). Gamoutput: $($gamoutput)"
}

$passwd = $NULL
$passwdMsg = ""
if($resetpassword)
{
	"Starting to generate new password"
	$passwd = iex ".\Get-RandomString.ps1 -Numbers -UpperCase -Length 10"
	"New password will be $passwd"
	$gamoutput = iex "$($gamPath)gam update user $($name) password $($passwd)"
	"Reset password for $($name). Gamoutput: $($gamoutput)"
	$passwdMsg = "<br /><h3>Password for Account</h3>The password for the account is: <b> $passwd </b>"
}

#send email to the supervisor
$subject = "User $name Offboarding"
$allgroupsString = $groupsemailname.GetEnumerator() |Sort-Object Name |ForEach-Object {"<tr><td>{0}</td> <td>{1}</td></tr>" -f $_.Name,$_.Value} | Out-String -Width 200
$solemanagerstring = $solemanager.GetEnumerator() |Sort-Object Name |ForEach-Object {"<tr><td>{0}</td> <td>{1}</td></tr>" -f $_.Name,$_.Value} | Out-String -Width 200
$sharedmanagerstring = $sharedmanager.GetEnumerator() |Sort-Object Name |ForEach-Object {"<tr><td>{0}</td> <td>{1}</td></tr>" -f $_.Name,$_.Value} | Out-String -Width 200
$noncoememberstring = $noncoemember.GetEnumerator() |Sort-Object Name |ForEach-Object {"<tr><td>{0}</td> <td>{1}</td></tr>" -f $_.Name,$_.Value} | Out-String -Width 200
$groupownersstring = ""

foreach($h in $groupsowners.GetEnumerator())
{
	$groupownersstring += '<H4>Owner(s) of ' + $($h.Name) + ':</H4>&emsp;'
	foreach($x in $h.Value.GetEnumerator())
	{
		"x is $($x.toString())"
		if($x.toString() -match "@edmonton.ca")
		{
			$groupownersstring += $x.toString()
			"Adding $($x.toString()) to the string"
		} else {
			"$x.toString is not in the @edmonton.ca domain"
		}
	}
}

$messageBody = "As the supervisor/document reviewer of user $name, access has been setup for you to review the user's Google account (Mail, Drive, Calendar, Sites, Contacts, Groups). "
$messageBody += "The review period is 30 calendar days. Once completed, you must submit the Declaration Form via the link below to authorize the permanent deletion of the user's Google account: <br>"
$messageBody += "<p><a href=`"http://webapps/OffboardingAdmin/declarationform.aspx`">http://accwebapps/TerminationsAdmin/DeclarationForm.aspx</a><br /><br />"
$messageBody += "Please find information below particular to this user that you may want to action. You have been given a 'manager' role to both calendars and groups that of which the user was the sole owner. "
$messageBody += "Their mail has been delegated to you and a Vacation Responder has been set on the account to contact you. You will receive a separate email if there are any Google Sites of which the user was an owner. "
$messageBody += "For more information regarding the role and responsibility of reviewer, and some walkthroughs of what needs to be done, please "
$messageBody += "<a href=`"https://docs.google.com/a/edmonton.ca/document/d/16QYW-LAEkak5r1zgvFXK3wKkQaNMEqt9zz2dOuKhmvQ/edit?usp=sharing`">click here</a> for more information.<br><p>"

#add in the password stuff
$messageBody += $passwdMsg

#add in the Calendar stuff
if($printcalendarowner){
	Write-Host "`nAdding calendar stuff to message"
	$messageBody += $calendarmessage
}


#Groups email message title
if($printgroups){
	$messageBody += "<br /><h3>Groups Information</h3>"

	#The table of groups
	$messageBody += "<table style=`"width:500px`">"
	$messageBody += "<tr><th align=left>All Groups User is a Member of:</th></tr>$allgroupsString"
	if($solemanagerstring -ne "")
	{
		$messageBody += "<tr></tr><tr><th align=left>Sole Ownership Groups</th></tr>$solemanagerstring"
	}
	if($sharedmanagerstring -ne "")
	{
		$messageBody += "<tr></tr><tr><th align=left>Group The User Co-Manages/Owns</th></tr>$sharedmanagerstring"
	}
	if($noncoememberstring -ne "")
	{
		$messageBody += "<tr></tr><tr><th align=left>Non-CoE Groups:</th></tr>$noncoememberstring"
	}

	$messageBody += "</table>"
	if($groupownersstring -ne "")
	{
		$messageBody += '<h3>Group Owners:</h3>' + $groupownersstring
	}
}

$style = "<style>BODY{font-family: Arial; font-size: 12pt;}"
$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; }"
$style = $style + "TD{border: 1px solid black; padding: 5px; }"
$style = $style + "</style>"
$style = "<head><pre>$style</pre></head>"
$messageBody = ConvertTo-HTML -head $style -body $messageBody

#send email
$toAddy = $submitter + "`@edmonton.ca"
if($sendemailflag){
	$toAddy += ',' + $supername + "`@edmonton.ca"
}
	SendEmail $toAddy "crsitgoogleoffboarding@edmonton.ca" $subject $messageBody
