#Start-Process Powershell -Verb runAS

$Begone = {

$Program = "quit"
$Program = Read-Host -prompt "What program would you like to uninstall? Type `"quit`" to exit"


if ($Program -match  "quit") {
			Exit}

$uninstall32 = get-childitem "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" | foreach { gp $_.PSPath } | ? { $_ -match "$Program" } | select UninstallString
$uninstall64 = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | foreach { gp $_.PSPath } | ? { $_ -match "$Program" } | select UninstallString

If (!$uninstall32 -and !$uninstall64) {Write-Host "Program not found!" 
							& $Begone}

If ($uninstall64) {
	$uninstall64 = $uninstall64.UninstallString -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X",""
	$uninstall64 = $uninstall64.Trim()
		
		if ($uninstall64 -notlike "*`"*" ) {
				$uninstall64 = "`"" + $uninstall64 + "`""
				$uninstall64 = $uninstall64.Trim()
				Write "Adding Quotes to 64bit Uninstall String"
			} 
		
		
Write "Uninstalling $Program."



				If ($uninstall64 -like "*exe*") {
					Write "Running Uninstall Executable"
					iex  "& $uninstall64"
					#Start-Process cmd -ArgumentList "/C $uninstall64 /X" }
					Start-Process cmd -ArgumentList "`"/C `"$uninstall64`" /X `""}
				
	Else {
		Write "Running MSI Exec Uninstall"
		start-process "msiexec.exe" -arg "/X $uninstall64"}

}









If ($uninstall32) {
	$uninstall32 = $uninstall32.UninstallString -Replace "msiexec.exe","" -Replace "/I","" -Replace "/X",""
	$uninstall32 = $uninstall32.Trim()

									
	if ($uninstall32 -notlike "*`"*" ) {
				$uninstall32 = "`"" + $uninstall32 + "`""
				$uninstall32 = $uninstall32.Trim()
				Write "Adding Quotes to 32 Bit Uninstall String"
			} 



Write "Uninstalling $Program"	

				If ($uninstall32 -like "*exe*") {
					Write "Running Uninstall Executable"
					iex  "& $uninstall32"
					# Not sure if needed: Start-Process cmd -ArgumentList "`"/C `"$uninstall32`" /X`"" 
					}
				
	Else {
		Write "Running MSI Exec Uninstall"
		start-process "msiexec.exe" -arg "/X $uninstall32" }

}}
	

& $begone