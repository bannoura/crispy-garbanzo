$ScriptRoot = Split-Path $MyInvocation.MyCommand.Definition
$Username = "admin@<company>.onmicrosoft.com"
$Password = Cat ($ScriptRoot + "\Password.txt") | ConvertTo-SecureString
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username, $Password

If (!(Get-Command "Get-Mailbox" -ErrorAction SilentlyContinue)) {
	Try {
		$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection
		Import-PSSession $Session -ErrorAction Stop
	} Catch {
		Write-Host "An error occured while establishing a remote PowerShell session with Exchange Online." -ForegroundColor Red
		Exit
	}
}

$SignatureFileName = ($ScriptRoot + "\Signature.html")
$SignatureHtml = Get-Content $SignatureFileName | Out-String
$AutoAddSignature = $true

$MailBoxMessageConfigurationsWithoutSignature = Get-Mailbox -RecipientTypeDetails UserMailbox | Get-MailboxMessageConfiguration | Where-Object { (!($_.SignatureText)) -or $_.SignatureText -eq "`r`n" }

ForEach ($MailBoxMessageConfiguration in $MailBoxMessageConfigurationsWithoutSignature) {
	$user = Get-User ($MailBoxMessageConfiguration.Identity)
	Set-MailboxMessageConfiguration -Identity ($MailBoxMessageConfiguration.Identity) `
	                                -AutoAddSignature $AutoAddSignature `
	                                -SignatureHtml ($SignatureHtml	-replace "%DisplayName%", $user.DisplayName `
	                                -replace "%Title%", $user.Title `
	                                -replace "%Phone%", $user.Phone `
	                                -replace "%MobilePhone%", $user.MobilePhone `
	                                -replace "%Fax%", $user.Fax `
	                                -replace "%WindowsEmailAddress%", $user.WindowsEmailAddress `
	                                -replace "%Notes%", $user.Notes `
	                                -replace "%Office%", $user.Office `
	                                -replace "%StreetAddress%", $user.StreetAddress `
	                                -replace "%PostalCode%", $user.PostalCode `
	                                -replace "%City%", $user.City `
	                                -replace "%Company%", $user.Company
	                                )
}

Remove-PSSession $Session