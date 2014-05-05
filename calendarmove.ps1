#where to write the reports, needs trailing slash
$reportDir = "e:\ScriptOutputTest\"
#where gam is located, needs trailing slash (default is "" if gam is in PATH)
$gamPath = ""

$sendemailflag = 1

$calendars = @()
$ownscalendars = @()
$allcalendars = @{}

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
	$owner = $false

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
#and we do want to move CoE calendars (so starting edmonton.ca_)
$owners = @()
foreach($i in $calendars){
	
	$owners.Clear()
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
				$owners += $username
			} else {
				#"not an owner of $username"
			}	
		}
	}
	if ($owners.length -eq 2)
	{
		$ownscalendars += $i
	}
}

"`ncalendars uniquely owned are: $ownscalendars`n"
"all calendars owned: $calendars"
"`n`nAssigning Calendars to Users"
#now to reassign the calendars to the supervisor
foreach($i in $ownscalendars)
{
	$gamoutput = ""
	#$gamoutput = iex "$($gamPath)gam calendar $($i) add owner $supername 2>&1"
	"tried to add $supername as owner of $i`nresult was: $gamoutput"
}