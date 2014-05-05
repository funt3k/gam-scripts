<#
	Description:
	Script to pull all the groupnames, email addresses and owners/managers
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

#where to write the reports, needs trailing slash
$reportDir = "e:\ScriptOutput\"
#where gam is located, needs trailing slash (default is "" if gam is in PATH)
$gamPath = "E:\gam3_edmonton.ca\"

if ($generateNewList)
{
	iex "$($gamPath)gam print groups > $($reportDir)groups_$(get-date -f MM-dd).csv"
}
$list = Import-Csv "$($reportDir)groups_$(get-date -f MM-dd).csv"

if ($writefile)
{
	$filename = "$($reportDir)groups-with-owners-$(get-date -f yyyy-MM-dd_hhmmss).csv"
	$stream = [System.IO.StreamWriter] "$($filename)"
	$stream.WriteLine("email,groupname,owner(s)")
}

foreach ($entry in $list)
  {
	"$($entry.Email)"
	
	$email = $($entry.Email) -replace "\'","``'"
	$comp = iex "$($gamPath)gam info group $($email)"
	if ($comp -ne $null) {$comp_branch = $comp.Split("`n")}
	$owners = @()
	$groupname = "Not Found"

	foreach($i in $comp_branch)
	{
		#Search for owners or managers
		$a = $i -match "^ owner: "
		if ($a)
		{
			
			$i = $i -replace "^ owner:",""
			"Owner pre-replace: $i"
			if ($i -match "\(user\)$") 
			{
				$i = $i -replace "\(user\)$",""
			}
			if ($i -match "\(group\)$")
			{ 
				$i = $i -replace "\(group\)$",""
			}
			"Owner after replace: $i"
			$owners += $i
		}
		
		#Search for owners or managers
		$a = $i -match "^ manager: "
		if ($a)
		{
			$i = $i -replace "^ manager:",""
			
			if ($i -match "\(user\)$") 
			{
				$i = $i -replace "\(user\)$",""
			}
			if ($i -match "\(group\)$") 
			{ 
				$i = $i -replace "\(group\)$",""
			}
			
			$owners += $i
		}
		
	
		#Search for groupname
		$a = $i -match "^ name: "
		if ($a)
		{
			$i = $i -replace "^ name: ",""
			$groupname = $i
		}
	}
	
	$ownerprntstr = ""
	foreach($i in $owners)
	{
		$ownerprntstr =  $ownerprntstr + $($i) +","
	}
	
	$ownerprntstr = $ownerprntstr.TrimEnd(",")
	$groupname = $groupname -replace ",","-"
	"Groupname: $($groupname)"
	"Owner(s): $($ownerprntstr)`n"
	
	if ($writefile)
	{
		$stream.WriteLine("$($entry.Email),$($groupname), $($ownerprntstr)")
		$stream.Flush()
	}
  }
  
if ($writefile)
{
	$stream.close()
}

#now mail it out
$smtpServer = "unixmail.gov.edmonton.ab.ca"
$att = new-object Net.Mail.Attachment($filename)
$msg = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$msg.From = "GAMScripter@edmonton.ca"
$msg.To.Add("crsitgooglesupport@edmonton.ca")
$msg.Subject = "List of all groups and all owners/managers"
$msg.Body = "Attached is the groups owners/managers report"
$msg.Attachments.Add($att)
$smtp.Send($msg)
$att.Dispose()