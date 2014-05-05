<#
	Description:
	Takes the newest csv file in $reportDir and attempts to move users into the correct CSV
	
	Last modified: Mach 4, 2014
	Author: Matthew Raven
#>
#where to look for the user report(s), with trailing slash.
$reportDir = "e:\ScriptOutput\UsersWithTitleBranchDeptOu\"
#where gam is located, needs trailing slash (default is "" if gam is in PATH)
$gamPath = "E:\gam3_edmonton.ca\"
#Log Dir
$logDir = "e:\ScriptOutput\"
#writefile -> write report
$writefile = 1
#send the report via email
$sendEmail = 1

$inCoECount = 0
$noBranchDeptCount = 0
$alreadyOKCount = 0
$moveCount = 0
$noOUCount = 0
$noMoveCount = 0
$invalidOU = 0
$moveFailed = 0

#retrieve the newest file from $reportDir
$latest = Get-ChildItem -Path $reportDir |  Sort-Object LastAccessTime -Descending | Select-Object -First 1
"Opening $($reportDir)$($latest.name)"
$list = Import-Csv "$($reportDir)$($latest.name)"

#Create the logfile
if ($writefile)
{
	$filename = "$($logDir)update_ou_logs-$(get-date -f yyyy-MM-dd_hhmmss).txt"
	$stream = [System.IO.StreamWriter] "$($filename)"
	$stream.WriteLine("Opening $($reportDir)$($latest.name)")
	$stream.WriteLine("Starting on $(get-date -f yyyy-MM-dd_hhmmss)")
}

foreach ($entry in $list)
{
	$ou = "/CoE/$($entry.dept)/$($entry.branch)"
	$continue = 0
	
	if ($($entry.branch) -eq "Not Listed")
	{
		#if the Branch isn't listed, then put them in their department only. Doesn't affect continue status
		$ou = "/CoE/$($entry.dept)"
	}
	#Make sure it's at least in the /CoE/ OU structure
	if ($($entry.OU) -match "^/CoE/")
	{
		$inCoECount++
		#"CoE $($inCoECount) : $($entry.email): Not in non-CoE Top Level OU. Current: $($entry.ou))"
		$continue = 1
	}
	
	#If empty OU, probably should move it, too
	if ($($entry.OU) -eq "/")
	{	
		$noOUCount++
		#"NOOU $($noOUCount): No Current OU"
		$continue = 1
	}
	
	#if Dept or Branch is Not Found, then skip
	if ($($entry.dept) -eq "Not Listed")
	{
		$noBranchDeptCount++
		#"BD $($noBranchDeptCount): $($entry.email) : Missing dept, OU would have been $($ou)"
		$continue = 0
	}
	
	#If user is in a sub OU of the correct OU, then skip
	if ($($entry.OU) -match $($ou))
	{
		$alreadyOKCount++
		#"OK $($alreadyOKCount) $($entry.email): in the correct /CoE/ ou $($entry.OU)"
		$continue = 0
	}
	
	if ($continue)
	{

		$moveCount++
		"TOMOV $($moveCount): $($entry.email): Attempting move to $($ou)"
		$email = $($entry.Email) -replace "\'","``'"
		$output = iex "$($gamPath)gam update user $($email) org `"$($ou)`"  2>&1"
		"  Command: $($gamPath)gam update user $($entry.email) org `"$($ou)`""
		"  GAM output is: $($output)"
		if ($($output) -match "INVALID_OU_ID")
		{
			$invalidOU++
		}
		if ($($output) -match "Error 412: Precondition Failed - conditionNotMet")
		{
			$moveFailed++
		}
		
		if ($writefile)
		{
			$stream.WriteLine("TOMOV $($moveCount): $($entry.email): Moving to $($ou)")
			$stream.WriteLine("  Command: $($gamPath)gam update user $([regex]::escape($entry.email)) org `"$($ou)`"")
			$stream.WriteLine("  GAM output is: $($output)")
			$stream.Flush()
		}
	}
	else
	{
		if ($writefile)
		{
			#$stream.WriteLine("NOMOV $($noMoveCount): $($entry.email): NOT moving. Leaving in $($entry.OU)")
			$stream.Flush()
		}	
		$noMoveCount++
		#"NOMOV $($noMoveCount): $($entry.email): NOT moving. Leaving in $($entry.OU)"
	}
	
}

if ($writefile)
{
	$stream.WriteLine("Totals")
	$stream.WriteLine("Total Accounts processed: $($moveCount + $noMoveCount)")
	$stream.WriteLine("Attempted Moved: $($moveCount)")
	$stream.WriteLine("Not Moved: $($noMoveCount)")
	$stream.WriteLine("Department Missing: $($noBranchDeptCount)")
	$stream.WriteLine("Failed Move (Error 412): $($moveFailed)")
	$stream.WriteLine("Failed Move (Invalid OU): $($invalidOU)")
	$stream.WriteLine("Successfully Moved: $($moveCount - $invalidOU - $moveFailed)")
	$stream.close()
}

if ($sendEmail)
{
	#now mail it out
	$smtpServer = "unixmail.gov.edmonton.ab.ca"
	$att = new-object Net.Mail.Attachment($filename)
	$msg = new-object Net.Mail.MailMessage
	$smtp = new-object Net.Mail.SmtpClient($smtpServer)
	$msg.From = "GAMScripter@edmonton.ca"
	$msg.To.Add("crsitgooglesupport@edmonton.ca")
	$msg.Subject = "OU Mover Log"
	$msg.Body = "Attached is OU Mover log"
	$msg.Attachments.Add($att)
	$smtp.Send($msg)
	$att.Dispose()
}