#where to write the reports, needs trailing slash
$reportDir = "e:\ScriptOutputTest\"
#where gam is located, needs trailing slash (default is "" if gam is in PATH)
$gamPath = ""

$sendemailflag = 1

$gmember = @()
$groupsowners = @{}
$groupsemailname = @{}

Function SendEmail ($toAddy, $fromAddy, $subject, $messageBody)
{
	"`n`nTO: $toAddy`nFROM: $fromAddy`nSUBJECT: $subject`nBODY:`n$messageBody"
	if($sendemailflag){
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
}

$name = Read-Host 'What is the username?'
$supername = Read-Host 'Who is the supervisor (username)?'
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
	$owners = @()
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
			$owners += [String] $i
		}
		
		#Search for managers
		$a = $i -match "^manager:"
		if ($a)
		{
			$i = $i -replace "^manager:",""
			$i = $i -replace "\(user\)",""
			$i = $i -replace "\(group\)",""
			$owners += [String] $i
		}
	}
	
	$groupsowners.Add( $x, $owners )
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
	SendEmail $toAddy "testscript@edmonton.ca" $subject $messageBody
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
			if([String] $x -eq [String] $fullemail)
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
			"Apparently $x has an @edmonton.ca address"
		}
		
	}
}

#send email to the supervisor
$subject = "User $name Offboarding"
$allgroupsString = $groupsemailname.GetEnumerator() |Sort-Object Name |ForEach-Object {"<tr><td>{0}</td> <td>{1}</td></tr>" -f $_.Name,$_.Value} | Out-String -Width 200
$solemanagerstring = $solemanager.GetEnumerator() |Sort-Object Name |ForEach-Object {"<tr><td>{0}</td> <td>{1}</td></tr>" -f $_.Name,$_.Value} | Out-String -Width 200
$sharedmanagerstring = $sharedmanager.GetEnumerator() |Sort-Object Name |ForEach-Object {"<tr><td>{0}</td> <td>{1}</td></tr>" -f $_.Name,$_.Value} | Out-String -Width 200
$noncoememberstring = $noncoemember.GetEnumerator() |Sort-Object Name |ForEach-Object {"<tr><td>{0}</td> <td>{1}</td></tr>" -f $_.Name,$_.Value} | Out-String -Width 200

$messageBody = "As the supervisor of user $name, you are being notified of the groups of which the user is a member.<br>"
$messageBody += "Please be advised that this user will be removed from all groups once the id is removed.<br>"
$messageBody += "You will have their email delegated to you. <br><br>User being removed: $name<br><p>"
$messageBody += "<table style=`"width:500px`">"
$messageBody += "<tr><th align=left>All Groups User is a Member of:</th></tr>$allgroupsString"
if($solemanagerstring -ne "")
{
	$messageBody += "<tr></tr><tr><th align=left>sole Ownership Groups</th></tr>$solemanagerstring"
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

$style = "<style>BODY{font-family: Arial; font-size: 12pt;}"
$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; }"
$style = $style + "TD{border: 1px solid black; padding: 5px; }"
$style = $style + "</style>"
$style = "<head><pre>$style</pre></head>"
$messageBody = ConvertTo-HTML -head $style -body $messageBody

$toAddy = $supername + "`@edmonton.ca"
SendEmail $toAddy "testscript@edmonton.ca" $subject $messageBody