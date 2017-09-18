#Changes user default reply address from CSV employee list
#Must be run in exchange 2007 management shell
#
#Cody Dills 9/15/17


#init variables
$NewDomain = "yourNewReplyDomain.com"
$directory = "\\location\csvList\"
$csvPath = $directory + "employeeNames.csv"

$DC = "domainController.yourDomain.com"
$Output = $directory + "conversionLog.txt"
$OutputFails = $directory + "conversionErrors.csv"


#Iterate through employee list 
Import-Csv $csvPath | Foreach-Object { 

    foreach ($property in $_.PSObject.Properties)
    {
        $identity = $property.value
		
		Try {
			#get mailbox object from employee name
			$mb = get-mailbox -identity $identity -resultsize Unlimited -DomainController $DC -errorAction Stop
			
			#captures old domain
			$SMTP = $mb.PrimarySmtpAddress
			[string]$Local = $SMTP.Local
			[string]$OldDomain = $SMTP.Domain
			[string]$CPSMTP = $Local + "@" + $OldDomain
		
			#creates new domain string
			[string]$NPSMTP = $Local + "@" + $NewDomain
		
			#logs old and new addresses to file
			[string]$iobject = $CPSMTP + "`t" + $NPSMTP
			Out-File $Output -InputObject $iobject -Append
		
			#actually does the thing
			Set-Mailbox -identity $identity -PrimarySmtpAddress $NPSMTP -EmailAddressPolicyEnabled $false -DomainController $DC -errorAction Stop
		}
		Catch {
			#error log
			Out-File $Output -InputObject ($identity + " `tnot found") -Append
			#adds failed names to CSV for later use
			Add-Content -path $OutputFails -Value $identity
			
		}
	}
}
