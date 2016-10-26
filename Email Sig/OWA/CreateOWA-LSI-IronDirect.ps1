
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
#LSI Signature Injection - IronDirect Template

$EmailSigGroup = Get-MsolGroup -all | where-object { $_.DisplayName -eq "LSI Signature Injection - IronDirect Template"}
#$EmailSigGroupUK = Get-MsolGroup -all | where-object { $_.DisplayName -eq "LSI Signature Injection - IronDirect Template UK"}

$MailBoxMessageConfigurationsWithoutSignature = Get-MsolGroupMember -GroupObjectId $EmailSigGroup.ObjectId -All
#$MailBoxMessageConfigurationsWithoutSignature += Get-MsolGroupMember -GroupObjectId $EmailSigGroupUK.ObjectId


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
<body link=`"#000000`">
<TABLE id=emailsignature border=0 cellSpacing=0 cellPadding=0>
<TBODY>
           <TR>
			<TD style=`"VERTICAL-ALIGN: top`"<A href=`"https://www.irondirect.com.com`"><IMG title=`"IronDirect`" border=0 name=`"IronDirect`" alt=`"IronDirect`" src=`"http://media.liquidityservices.com/IronDirect/IronDirect_Logo_wTag_forWhtBkgd-01.png`" width=190 height=50></A> </TD>
			<TD style=`"HEIGHT: auto; COLOR: white; VERTICAL-ALIGN: bottom`" bgColor=#ffffff>__</TD>
			<TD style=`"WIDTH: 1px; HEIGHT: auto; COLOR: white; VERTICAL-ALIGN: bottom`" bgColor=#000000></TD>
			<TD style=`"HEIGHT: auto; COLOR: white; VERTICAL-ALIGN: bottom`" bgColor=#ffffff>__</TD>
               <TD rowspan=`"3`">
                  <TABLE id=emailsignatureDetails border=0 cellSpacing=0 cellPadding=1>
                       <TBODY>
                           <TR>
                               <TD style=`"TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #000000`"><b>" + $user.DisplayName + "</b></TD>
                           </TR>
                           <TR>
                               <TD style=`"TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #000000`"><b>" + $user.Title + "</b></TD>
                           </TR>
                           <TR>
                               <TD style=`"TEXT-ALIGN:  Left; FONT-SIZE: 10px;FONT-FAMILY: Arial, Helvetica; COLOR: #FFFFFF`">__</TD>
                           </TR>"
                           if(($User.Phone).tostring().length -ne 0){ $EmailHtml+=
                            "<TR>
                               <TD style=`"TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #000000`">" + $user.Phone + " (Office)</SPAN></TD>
							</TR>"}
                            if(($User.MobilePhone).tostring().length -ne 0){ $EmailHtml+=
							"<TR>
								<TD style=`"TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #000000`">" + $user.MobilePhone + " (Mobile)</SPAN> </TD>
							</TR>"}
                            if(($User.fax).tostring().length -ne 0){ $EmailHtml+=
							"<TR>
								<TD style=`"TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #000000`">" + $user.fax + " (Fax)</SPAN> </TD>
							</TR>"}
                            if(($UserMailbox.CustomAttribute1).tostring().length -ne 0){ $EmailHtml+=
							"<TR>
								<TD style=`"TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #000000`">" + $UserMailbox.CustomAttribute1 + " (Skype)</SPAN> </TD>
							</TR>"}
							$EmailHtml+= "<TR>
								<TD style=`"TEXT-DECORATION: none; TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #000000`" noWrap><SPAN id=email>" + $user.WindowsEmailAddress + "</SPAN></TD>
							</TR>
                           <TR>
                               <TD style=`"TEXT-ALIGN:  Left; FONT-SIZE: 10px;FONT-FAMILY: Arial, Helvetica; COLOR: #FFFFFF`">__</TD>
                           </TR>
                           <TR>
                               <TD style=`"TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #000000`">IronDirect</TD>
                           </TR>
                           <TR>
                               <TD style=`"TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #000000`">" + $user.StreetAddress + "</TD>
                           </TR>
                           <TR>
                               <TD style=`"TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #000000`">" + $user.City + ", " + $user.StateOrProvince + " " + $user.PostalCode + " " + $user.CountryOrRegion + "</TD>
                           </TR>
							<TR>
								<TD style=`"TEXT-DECORATION: none; TEXT-ALIGN:  Left;FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #000000`" href=`"http://www.IronDirect.com`">www.IronDirect.com</A></TD>
						</TR>
					</TBODY>
				</TABLE>
            </TD>
		</TR>
			<TR>
				<td style=`"Vertical-align: bottom`"; align=`"right`">
					<TABLE>
						<TBODY>
							<TR align=`"right`">
								<TD style=`"VERTICAL-ALIGN: bottom`"<A href=`"http://fast.wistia.net/embed/iframe/meg5jhs5jw?videoFoam=true`"><IMG style=`"float: right;`" title=`"About IronDirect`" border=0 name=`"About IronDirect`" alt=`"About IronDirect`" src=`"http://media.liquidityservices.com/IronDirect/About_Us.jpg`" width=157 height=21></A> </TD>
							</TR>
							<TR align=`"right`">
								<TD style=`"VERTICAL-ALIGN: bottom`"<A href=`"http://players.brightcove.net/63193328001/default_default/index.html?videoId=4897877820001`"><IMG  style=`"float: right;`" title=`"About IronDirect`" border=0 name=`"About IronDirect`" alt=`"About IronDirect`" src=`"http://media.liquidityservices.com/IronDirect/See_us_in_action.jpg`" width=157 height=21></A> </TD>
							</TR>
						</TBODY>
					</TABLE>
				</td>
			<TD style=`"HEIGHT: auto; COLOR: white; VERTICAL-ALIGN: bottom`" bgColor=#ffffff>__</TD>
			<TD style=`"WIDTH: 1px; HEIGHT: auto; COLOR: white; VERTICAL-ALIGN: bottom`" bgColor=#000000></TD>
			<TD style=`"HEIGHT: auto; COLOR: white; VERTICAL-ALIGN: bottom`" bgColor=#ffffff>__</TD>
		</TR>
	</TBODY>
</TABLE>"

    #Write-Host $user.DisplayName
    #Write-Host $user.Fax
    #Write-Host $EmailHtml
    $EmailHtml |Out-File "E:\vso\Engineering\Scripts and Utilities\Microsoft\365\Email Sig\OWA\HTML_Test.Txt"
	Set-MailboxMessageConfiguration -Identity ($MailBoxMessageConfiguration.EmailAddress) `
	                                -AutoAddSignature $AutoAddSignature `
	                                -SignatureHtml $EmailHtml
     #Write-Host $user.DisplayName

}

Remove-PSSession $Session