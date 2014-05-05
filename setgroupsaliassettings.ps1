$sourcefile = "e:\Scripts\Groupalias.csv" #file to read. needs to include full path
$reportpath = "e:\ScriptOutput\" #path to export logging. needs trailing slash
$gampath = "e:\gam3_edmonton.ca\" #path to gam, needs trailing slash
$logging = 1 #set to write a log
$detailedlogging = 1 #add in extra detail to the logs (namely the direct gam output)
$aliascount = 0 #counter for number of aliases processed.
$duplicate = 0 #count the number of duplicates
$invalidemail = 0
$notfound = 0

$list = Import-Csv $($sourcefile)

#Create the logfile
if ($logging)
{
	$filename = "$($reportpath)update-group-aliases-$(get-date -f yyyy-MM-dd_hhmmss).txt"
	$stream = [System.IO.StreamWriter] "$($filename)"
	$stream.WriteLine("Adding aliases to groups!")
	$stream.WriteLine("Starting on $(get-date -f yyyy-MM-dd_hhmmss)")
}

foreach ($entry in $list)
  {
	"ALIAS $($aliascount): Adding $($entry.aliasname) to account $($entry.Groupname)"
	$stream.WriteLine("ALIAS $($aliascount): Adding $($entry.aliasname) to account $($entry.Groupname)")
	
	$groupemail = $($entry.Groupname) -replace "\'","``'"

	$aliasemail = $($entry.aliasname) -replace "\'","``'"
	
	$cmd = "$($gampath)gam.exe create alias `"$($aliasemail)`" group `"$($groupemail)`" 2>&1"
    $output = Invoke-Expression $cmd
	if ($detailedlogging) { $stream.WriteLine("  GAM output is: $($output)") }
	if ($output -match "Error 409: Entity already exists.") { $duplicate++ }
	if ($output -match "Error 404: Resource Not Found") { $notfound++ }
	if ($output -match "Error 400: Invalid Email") { $invalidemail++ }
	$aliascount++
	$stream.Flush()
  }
  
if ($logging)
{	
	$stream.WriteLine("`nTotals")
	$stream.WriteLine("------")
	$stream.WriteLine("Alias adds attempted: $($aliascount)")
	$stream.WriteLine("Duplicate aliases: $($duplicate)")
	$stream.WriteLine("User accounts not found: $($notfound)")
	$stream.WriteLine("Invalid Emails (likely has ampersand): $($invalidemail)")
	$stream.close()
}