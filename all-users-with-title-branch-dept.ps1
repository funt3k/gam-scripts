<#
	Description:
	Script to pull all the user accounts and suplementary information
	Assumes GAM is installed and in the PATH
	
	Author:
	matthew.raven@edmonton.ca
	
	Last Modified:
	March 3, 2014
#>

#writefile -> write report
$writefile = 1
#generateNewList -> write out the users list again. Only has to be done once per day
$generateNewList = 1

#where to write the reports, with trailing slash.
$reportDir = "e:\ScriptOutput\UsersWithTitleBranchDeptOu\"
#where gam is located, needs trailing slash (default is "" if gam is in PATH)
$gamPath = "E:\gam3_edmonton.ca\"

#Not Found text
$notFound = "Not Listed"

if ($generateNewList) 
{
	iex "$($gamPath)gam print users > $($reportDir)users_$(get-date -f MM-dd).csv"
}

$list = Import-Csv "$($reportDir)users_$(get-date -f MM-dd).csv"

if ($writefile)
{
	$filename = "$($reportDir)users-with-title-branch-dept-$(get-date -f yyyy-MM-dd_hhmmss).csv"
	$stream = [System.IO.StreamWriter] "$($filename)"
	$stream.WriteLine("email,first,last,address,title,branch,dept,OU,suspended")
}
foreach ($entry in $list)
  {
  
	"$($entry.email)"
	$email = $($entry.Email) -replace "\'","``'"
	$comp = iex "$($gamPath)gam info user $($email)"
	if ($comp -ne $null) {$comp_branch = $comp.Split("`n")}
	
	$branch = $notFound
	$dept = $notFound
	$title = $notFound
	$first = $notFound
	$last = $notFound
	$address = $notFound
	
	foreach($i in $comp_branch)
	{
	
		#Search for First Name
		$a = $i -match "^First Name: "
		if ($a)
		{
			$i = $i -replace "^First Name: ",""
			$first = $i
		}
		
		#Search for Last Name
		$a = $i -match "^Last Name: "
		if ($a)
		{
			$i = $i -replace "^Last Name: ",""
			$last = $i
		}
		
		#Search for Street Address
		$a = $i -match "^ streetAddress: "
		if ($a)
		{
			$i = $i -replace "^ streetAddress: ",""
			$address = $i
		}
		
		#Search for job title
		$a = $i -match "^ title: "
		if ($a)
		{
			$i = $i -replace "^ title: ",""
			$title = $i
		}
	
		#Search for location (aka Branch)
		$a = $i -match "^ location: "
		if ($a)
		{
			$i = $i -replace "^ location: ",""
			$branch = $i
		}
		
		#Search for department
		$a = $i -match "^ department: "
		if ($a)
		{
			$i = $i -replace "^ department: ",""
			$dept = $i
		}
		
		#Search for OU
		$a = $i -match "^Google Org Unit Path: "
		if ($a)
		{
			$i = $i -replace "^Google Org Unit Path: ",""
			$OU = $i
		}
		
		#Search for Suspended Status
		$a = $i -match "^Account Suspended: "
		if ($a)
		{
			$i = $i -replace "^Account Suspended: ",""
			$suspended = $i
		}
	}

	#fixing the output so it doesn't have any commas (which would break our CSV)

	$title = $title -replace ","," "
	$first = $first -replace ","," "
	$last = $last -replace ","," "
	$address = $address -replace ",", " -"
	$branch = $branch -replace ","," -"
	$dept = $dept -replace ","," -"
	
	
	"First Name: $($first)"
	"Last Name: $($last)"
	"Address: $($address)"
	"Title: $($title)"
	"Branch: $($branch)"
	"Department: $($dept)"
	"OU: $($OU)"
	"Suspended: $($suspended)`n"
	
	if ($writefile)
	{
		$stream.WriteLine("$($entry.email),$($first),$($last),$($address),$($title),$($branch),$($dept),$($OU),$($suspended)")
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
$msg.Subject = "List of all users and associated information"
$msg.Body = "Attached is the users info report"
$msg.Attachments.Add($att)
$smtp.Send($msg)
$att.Dispose()