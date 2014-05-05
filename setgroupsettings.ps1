$list = Import-Csv groups.csv
foreach ($entry in $list)
  {
	"Changing $($entry.groupname) to $($entry.size)"
    c:\gam\gam.exe update group $($entry.groupname) max_message_bytes $($entry.size)
  }