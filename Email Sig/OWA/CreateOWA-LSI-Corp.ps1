
$ScriptRoot = Split-Path $MyInvocation.MyCommand.Definition

####################################
#Build credential from hashed password

$User = "svc_rmscript365@liquidityservices.com"
$PasswordFile = "E:\vso\Engineering\Scripts and Utilities\Microsoft\365\CredentialStore\365Password.txt"
$KeyFile = "E:\vso\Engineering\Scripts and Utilities\Microsoft\365\CredentialStore\AES.key"
$key = Get-Content $KeyFile
$Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, (Get-Content $PasswordFile | ConvertTo-SecureString -Key $key)

####################################

Connect-MsolService -Credential $Credential 

If (!(Get-Command "Get-Mailbox" -ErrorAction SilentlyContinue)) {
	Try {
		$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection
		Import-PSSession $Session -ErrorAction Stop
	} Catch {
		Write-Host "An error occured while establishing a remote PowerShell session with Exchange Online." -ForegroundColor Red
		Exit
	}
}

$MailBoxMessageConfigurationsWithoutSignature = $null

#AD Groups
#LSI Signature Injection - CAG Template
#LSI Signature Injection - CAG Template UK
#LSI Signature Injection - Corp Template with IT Footer UK
#LSI Signature Injection - Corp Template with IT Support Footer
#LSI Signature Injection - LSI Corp Template
#LSI Signature Injection - LSI Corp Template UK
#LSI Signature Injection - Marketplaces Template
#LSI Signature Injection - Marketplaces Template UK
#LSI Signature Injection - NI Template
#LSI Signature Injection - NI Template UK
#LSI Signature Injection - RSCG Template
#LSI Signature Injection - RSCG Template UK
#LSI Signature Injection - Transporation Template

$EmailSigGroup = Get-MsolGroup -all | where-object { $_.DisplayName -eq "LSI Signature Injection - LSI Corp Template"}
$EmailSigGroupUK = Get-MsolGroup -all | where-object { $_.DisplayName -eq "LSI Signature Injection - LSI Corp Template UK"}

$MailBoxMessageConfigurationsWithoutSignature = Get-MsolGroupMember -GroupObjectId $EmailSigGroup.ObjectId -All
$MailBoxMessageConfigurationsWithoutSignature += Get-MsolGroupMember -GroupObjectId $EmailSigGroupUK.ObjectId -All


#$MailBoxMessageConfigurationsWithoutSignature |Out-GridView

$AutoAddSignature = $true

ForEach ($MailBoxMessageConfiguration in $MailBoxMessageConfigurationsWithoutSignature) {
	$user = Get-User ($MailBoxMessageConfiguration.EmailAddress)
    Write-Host "Now Proccessing" $MailBoxMessageConfiguration.DisplayName
    $UserMailbox = Get-Mailbox ($MailBoxMessageConfiguration.EmailAddress)

    $EmailHtml= "<!DOCTYPE HTML PUBLIC `"-//W3C//DTD HTML 4.0 Transitional//EN`">
<HTML><HEAD><TITLE>Microsoft Office Outlook Signature</TITLE>
<META http-equiv=Content-Type content=`"text/html; charset=windows-1252`">
<META content=`"MSHTML 6.00.3790.186`" name=GENERATOR></HEAD>
<body link=`"#124C86`">
<TABLE id=emailsignature border=0 cellSpacing=0 cellPadding=0>
<TBODY>
           <TR>
			<TD style=`"VERTICAL-ALIGN: top`"<A href=`"https://www.liquidityservices.com`"><IMG title=`"Liquidity Services`" border=0 name=`"Liquidity Services`" alt=`"Liquidity Services`" src=`"http://media.liquidityservices.com/Corporate/Email-Signature_LS01.png`" width=190 height=113></A> </TD>
			<TD style=`"HEIGHT: auto; COLOR: white; VERTICAL-ALIGN: bottom`" bgColor=#ffffff>__</TD>
			<TD style=`"WIDTH: 1px; HEIGHT: auto; COLOR: white; VERTICAL-ALIGN: bottom`" bgColor=#124C86></TD>
			<TD style=`"HEIGHT: auto; COLOR: white; VERTICAL-ALIGN: bottom`" bgColor=#ffffff>__</TD>
               <TD>
                  <TABLE id=emailsignatureDetails border=0 cellSpacing=0 cellPadding=1>
                       <TBODY>
                           <TR>
                               <TD style=`"TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #124C86`"><b>" + $user.DisplayName + "</b></TD>
                           </TR>
                           <TR>
                               <TD style=`"TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #124C86`"><b>" + $user.Title + "</b></TD>
                           </TR>
                           <TR>
                               <TD style=`"TEXT-ALIGN:  Left; FONT-SIZE: 10px;FONT-FAMILY: Arial, Helvetica; COLOR: #FFFFFF`">__</TD>
                           </TR>"
                           if(($User.Phone).tostring().length -ne 0){ $EmailHtml+=
                            "<TR>
                               <TD style=`"TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #124C86`">" + $user.Phone + " (Office)</SPAN></TD>
							</TR>"}
                            if(($User.MobilePhone).tostring().length -ne 0){ $EmailHtml+=
							"<TR>
								<TD style=`"TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #124C86`">" + $user.MobilePhone + " (Mobile)</SPAN> </TD>
							</TR>"}
                            if(($User.fax).tostring().length -ne 0){ $EmailHtml+=
							"<TR>
								<TD style=`"TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #124C86`">" + $user.fax + " (Fax)</SPAN> </TD>
							</TR>"}
                            if(($UserMailbox.CustomAttribute1).tostring().length -ne 0){ $EmailHtml+=
							"<TR>
								<TD style=`"TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #124C86`">" + $UserMailbox.CustomAttribute1 + " (Skype)</SPAN> </TD>
							</TR>"}
							$EmailHtml+= "<TR>
								<TD style=`"TEXT-DECORATION: none; TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #124C86`" noWrap><SPAN id=email>" + $user.WindowsEmailAddress + "</SPAN></TD>
							</TR>
                           <TR>
                               <TD style=`"TEXT-ALIGN:  Left; FONT-SIZE: 10px;FONT-FAMILY: Arial, Helvetica; COLOR: #FFFFFF`">__</TD>
                           </TR>
                           <TR>
                               <TD style=`"TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #124C86`">Liquidity Services</TD>
                           </TR>
                           <TR>
                               <TD style=`"TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #124C86`">" + $user.StreetAddress + "</TD>
                           </TR>
                           <TR>
                               <TD style=`"TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #124C86`">" + $user.City + ", " + $user.StateOrProvince + " " + $user.PostalCode + " " + $user.CountryOrRegion + "</TD>
                           </TR>
							<TR>
								<TD style=`"TEXT-DECORATION: none; TEXT-ALIGN:  Left;FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #124C86`" href=`"http://www.liquidityservices.com`">www.liquidityservices.com</A></TD>
						</TR>
					</TBODY>
				</TABLE>
		</TR>
	</TBODY>
</TABLE>"

    #Write-Host $user.DisplayName
    #Write-Host $user.Fax
    #Write-Host $EmailHtml
    #$EmailHtml |Out-File C:\Users\admin.taird\Desktop\Test.Txt
	Set-MailboxMessageConfiguration -Identity ($MailBoxMessageConfiguration.EmailAddress) `
	                                -AutoAddSignature $AutoAddSignature `
	                                -SignatureHtml $EmailHtml


}

Remove-PSSession $Session