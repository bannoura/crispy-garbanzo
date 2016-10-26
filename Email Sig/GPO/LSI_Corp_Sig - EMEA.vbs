' VBScript: <Signatures.vbs>
' AUTHOR: Peter Aarts
' Contact Info: peter.aarts@l1.nl
' Version 2.04
' Date: January 20, 2006
' Modified By Brad Marsh Now works with both 2003 and 2007 outlook 
' Modified By Todd Aird Now works with both 2003 and 2007,2010,2013 outlook
' Date 20 May 2010
' Tested on windows 7, Vista, XP, XP64 and office 2003, 2007 and 2010. 
' NOTE will not work that well with various email accounts
'==================== 
 
'Option Explicit
On Error Resume Next
 
Dim qQuery, objSysInfo, objuser, ObjOutlook2013
Dim objShell,grouplistD,ADSPath,userPath,listGroup
Dim FullName, EMail, Title, PhoneNumber, MobileNumber, FaxNumber, OfficeLocation, Department
Dim web_address, FolderLocation, HTMFileString, StreetAddress, Town, State, Company
Dim ZipCode, PostOfficeBox, UserDataPath
Dim AddBanner, BannerPath, BannerLink, BannerGroup
 
' Read LDAP(Active Directory) information to assigns the user's info to variables.
'====================
Set objSysInfo = CreateObject("ADSystemInfo")
objSysInfo.RefreshSchemaCache
qQuery = "LDAP://" & objSysInfo.Username
Set objuser = GetObject(qQuery)
 
FullName = objuser.displayname
FirstName = objuser.givenName
LastName = objuser.sn
EMail = objuser.mail
Company = objuser.Company
Title = objuser.title
PhoneNumber = objuser.TelephoneNumber
FaxNumber = objuser.FaxNumber
OfficeLocation = objuser.physicalDeliveryOfficeName
StreetAddress = objuser.streetaddress
PostofficeBox = objuser.postofficebox
Department = objUser.Department
ZipCode = objuser.postalcode
Town = objuser.l
State = objuser.st
Country = objuser.co
MobileNumber = objuser.TelephoneMobile
SkypeName = objuser.extensionAttribute1
TargetAddress = objuser.targetAddress
web_address = "http://www.liquidityservices.com"

'===================================================================
'Add Banners to email foooter
'=================================================================== 
'Set this value to True to add the EMEA banner footer to emails'
 AddBanner = "True"
 'Set the path to the EMEA Banner here'
 BannerPath = "http://media.liquidityservices.com/Corporate/EMEA/EmailSignature_BuyersGuide_Feb2016.jpg"
'Set Link to for banner image'
 BannerLink = "http://www.go-dove.com/en/c/buyers-guide-to-used-industrial-equipment&utm_source=go-dove&utm_medium=email&utm_campaign=staffemail"
 'Set Group to apply banner to'
 BannerGroup = "Email.Banner-EMEA"

'===================================================================


'This is so see if the user we are dealing with is an office 365 user
'====================
'Convert target address to lowercase
TargetAddress = LCase(TargetAddress)


If TargetAddress ="" Then
	WScript.Quit
ElseIf Right(TargetAddress,(Len(TargetAddress)-InStr(TargetAddress,"liquidityservices.mail.onmicrosoft.com"))+1) <> "liquidityservices.mail.onmicrosoft.com" Then
	WScript.Quit
End IF
'====================

'This section creates the signature files names and locations.
'====================
' Corrects Outlook signature folder location. Just to make sure that
' Outlook is using the purposed folder defined with variable : FolderLocation
' Example is based on Dutch version.
' Changing this in a production environment might create extra work
' all employees are missing their old signatures
'====================
Dim RegKey, RegKey07, RegKey10, RegKey13, RegKeyParm
Set objShell = CreateObject("WScript.Shell")
RegKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\11.0\Common\General"
RegKey07 = "HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Common\General"
RegKey10 = "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Common\General"
RegKey13 = "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\Common\General"
RegKey07 = RegKey07 & "\Signatures"
RegKey10 = RegKey10 & "\Signatures"
RegKey13 = RegKey13 & "\Signatures"
RegKey = RegKey & "\Signatures"
objShell.RegWrite RegKey , "Signatures"
objShell.RegWrite RegKey07 , "Signatures"
objShell.RegWrite RegKey10 , "Signatures"
objShell.RegWrite RegKey13 , "Signatures"
UserDataPath = ObjShell.ExpandEnvironmentStrings("%appdata%")
FolderLocation = UserDataPath &"\Microsoft\Signatures\"
HTMFileString = FolderLocation & "LSI_Full.htm"
 
' This section disables the change of the signature by the user.
'====================
'objShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\11.0\Common\MailSettings\NewSignature" , "L1-Handtekening"
'objShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\11.0\Common\MailSettings\ReplySignature" , "L1-Handtekening"
'objShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\11.0\Outlook\Options\Mail\EnableLogging" , "0", "REG_DWORD" 
 
' This section checks if the signature directory exits and if not creates one.
'====================
Dim objFS1
Set objFS1 = CreateObject("Scripting.FileSystemObject")
If (objFS1.FolderExists(FolderLocation)) Then
Else
Call objFS1.CreateFolder(FolderLocation)
End if
 
' The next section builds the signature file
'====================
Dim objFSO
Dim objFile,afile
Dim aQuote
aQuote = chr(34)
 
' This section builds the HTML file version
'====================
Set objFSO = CreateObject("Scripting.FileSystemObject")
 
' This section deletes to other signatures.
' These signatures are automatically created by Outlook 2003.
'====================
Set AFile = objFSO.GetFile(Folderlocation&"LSI.rtf")
aFile.Delete
Set AFile = objFSO.GetFile(Folderlocation&"LSI.txt")
aFile.Delete
 
Set objFile = objFSO.CreateTextFile(HTMFileString,True)
objFile.Close
Set objFile = objFSO.OpenTextFile(HTMFileString, 2)
 
objfile.write "<!DOCTYPE HTML PUBLIC " & aQuote & "-//W3C//DTD HTML 4.0 Transitional//EN" & aQuote & ">" & vbCrLf
objfile.write "<HTML><HEAD><TITLE>Microsoft Office Outlook Signature</TITLE>" & vbCrLf
objfile.write "<META http-equiv=Content-Type content=" & aQuote & "text/html; charset=windows-1252" & aQuote & ">" & vbCrLf
objfile.write "<META content=" & aQuote & "MSHTML 6.00.3790.186" & aQuote & " name=GENERATOR></HEAD>" & vbCrLf
objfile.write "<body link=" & aQuote & "#124C86" & aQuote & ">" & vbCrLf
objfile.write "<TABLE id=emailsignature border=0 cellSpacing=0 cellPadding=0>"& vbCrLf
objfile.write "<TBODY>"& vbCrLf
objfile.write "           <TR>"& vbCrLf
objfile.write "			<TD style=" & aQuote & "VERTICAL-ALIGN: top" & aQuote & "<A href=" & aQuote & "http://www.liquidityservices.com" & aQuote & "><IMG title=" & aQuote & "Liquidity Services Inc." & aQuote & " border=0 name=" & aQuote & "Liquidity Services Inc." & aQuote & " alt=" & aQuote & "Liquidity Services Inc." & aQuote & " src=" & aQuote & "http://media.liquidityservices.com/Corporate/Email-Signature_LS01.png" & aQuote & " width=190 height=113></A><BR><BR><BR><BR><BR><BR><A href=" & aQuote & "http://fast.wistia.net/embed/iframe/i6qzpy410r?videoFoam=true" & aQuote & "><img align=right src=" & aQuote & "http://media.liquidityservices.com/video/video.png" & aQuote & "></A></TD>"& vbCrLf
objfile.write "			<TD style=" & aQuote & "HEIGHT: auto; COLOR: white; VERTICAL-ALIGN: bottom" & aQuote & " bgColor=#ffffff>__</TD>"& vbCrLf
objfile.write "			<TD style=" & aQuote & "WIDTH: 1px; HEIGHT: auto; COLOR: white; VERTICAL-ALIGN: bottom" & aQuote & " bgColor=#124C86></TD>"& vbCrLf
objfile.write "			<TD style=" & aQuote & "HEIGHT: auto; COLOR: white; VERTICAL-ALIGN: bottom" & aQuote & " bgColor=#ffffff>__</TD>"& vbCrLf

objfile.write "               <TD>"& vbCrLf
objfile.write "                  <TABLE id=emailsignatureDetails border=0 cellSpacing=0 cellPadding=1>"& vbCrLf
objfile.write "                       <TBODY>"& vbCrLf
objfile.write "                           <TR>"& vbCrLf
objfile.write "                               <TD style=" & aQuote & "TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #124C86" & aQuote & "><b>" & FirstName & " " & LastName & "</b></TD>"& vbCrLf
objfile.write "                           </TR>"& vbCrLf
objfile.write "                           <TR>"& vbCrLf
objfile.write "                               <TD style=" & aQuote & "TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #124C86" & aQuote & "><b>" & Title & "</b></TD>"& vbCrLf
objfile.write "                           </TR>"& vbCrLf

objfile.write "                           <TR>"& vbCrLf
objfile.write "                               <TD style=" & aQuote & "TEXT-ALIGN:  Left; FONT-SIZE: 10px;FONT-FAMILY: Arial, Helvetica; COLOR: #FFFFFF" & aQuote & ">__</TD>"& vbCrLf
objfile.write "                           </TR>"& vbCrLf

If PhoneNumber <> "" Then
	objfile.write "                           <TR>"& vbCrLf
	objfile.write "                               <TD style=" & aQuote & "TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #124C86" & aQuote & ">" & PhoneNumber & " (Office)</SPAN></TD>"& vbCrLf
	objfile.write "							</TR>"& vbCrLf
End if

If MobileNumber <> "" Then
	objfile.write "							<TR>"& vbCrLf
	objfile.write "								<TD style=" & aQuote & "TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #124C86" & aQuote & ">" & MobileNumber & " (Mobile)</SPAN> </TD>"& vbCrLf
	objfile.write "							</TR>"& vbCrLf
End If

IF FaxNumber <> "" Then
	objfile.write "							<TR>"& vbCrLf
	objfile.write "								<TD style=" & aQuote & "TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #124C86" & aQuote & ">" & FaxNumber & " (Fax)</SPAN> </TD>"& vbCrLf
	objfile.write "							</TR>"& vbCrLf
End IF

If SkypeName <> "" then
	objfile.write "							<TR>"& vbCrLf
	objfile.write "								<TD style=" & aQuote & "TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #124C86" & aQuote & ">" & SkypeName & " (Skype)</SPAN></TD>"& vbCrLf
	objfile.write "							</TR>"& vbCrLf
End if

objfile.write "							<TR>"& vbCrLf
objfile.write "								<TD style=" & aQuote & "TEXT-DECORATION: none; TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #124C86" & aQuote & " noWrap><SPAN id=email>" & EMail & "</SPAN></TD>"& vbCrLf
objfile.write "							</TR>"& vbCrLf



objfile.write "                           <TR>"& vbCrLf
objfile.write "                               <TD style=" & aQuote & "TEXT-ALIGN:  Left; FONT-SIZE: 10px;FONT-FAMILY: Arial, Helvetica; COLOR: #FFFFFF" & aQuote & ">__</TD>"& vbCrLf
objfile.write "                           </TR>"& vbCrLf

objfile.write "                           <TR>"& vbCrLf
objfile.write "                               <TD style=" & aQuote & "TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #124C86" & aQuote & ">Liquidity Services</TD>"& vbCrLf
objfile.write "                           </TR>"& vbCrLf
objfile.write "                           <TR>"& vbCrLf
objfile.write "                               <TD style=" & aQuote & "TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #124C86" & aQuote & ">" & StreetAddress & "</TD>"& vbCrLf
objfile.write "                           </TR>"& vbCrLf
objfile.write "                           <TR>"& vbCrLf
objfile.write "                               <TD style=" & aQuote & "TEXT-ALIGN:  Left; FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #124C86" & aQuote & ">" & Town & ", " & State & " " & ZipCode & " " & Country & "</TD>"& vbCrLf
objfile.write "                           </TR>"& vbCrLf



objfile.write "							<TR>"& vbCrLf
objfile.write "								<TD style=" & aQuote & "TEXT-DECORATION: none; TEXT-ALIGN:  Left;FONT-SIZE: 13px;FONT-FAMILY: Arial, Helvetica; COLOR: #124C86" & aQuote & " href=" & aQuote & "http://www.liquidityservices.com" & aQuote &">www.liquidityservices.com</A></TD>"& vbCrLf
objfile.write "						</TR>"& vbCrLf
objfile.write "                           <TR>"& vbCrLf
objfile.write "                               <br><table width="& aQuote &"100"& aQuote &" border=0 cellspacing=0 cellpadding=1><TD align="& aQuote &"left"& aQuote &"><a href= " & aQuote & "https://www.linkedin.com/company/37685" & aQuote & "><img src= " & aQuote & "http://media.liquidityservices.com/SocialMedia/LinkedIn.png" & aQuote & "></a></td><TD align="& aQuote &"center"& aQuote &"><a href= " & aQuote & "https://www.facebook.com/liquidityservices" & aQuote & "> <img src= " & aQuote & "http://media.liquidityservices.com/SocialMedia/Facebook.png" & aQuote & "></a></td><TD align="& aQuote &"right"& aQuote &"><a href= " & aQuote & "https://twitter.com/liquidityinc" & aQuote &" > <img src= " & aQuote & "http://media.liquidityservices.com/SocialMedia/Twitter.png" & aQuote & "></a></TD></table>"& vbCrLf
objfile.write "                           </TR>"& vbCrLf
objfile.write "					</TBODY>"& vbCrLf
objfile.write "				</TABLE>"& vbCrLf
objfile.write "		</TR>"& vbCrLf
objfile.write "	</TBODY>"& vbCrLf
objfile.write "</TABLE> "& vbCrLf


'Adding Email Banner for EMEA'

If AddBanner = "True" then
	If isMember(BannerGroup) Then
		objfile.write "<TABLE id=EmailSignatureBanner border=0 cellSpacing=0 cellPadding=0>"& vbCrLf
		objfile.write "	<TBODY>"& vbCrLf
		objfile.write "				<TR>"& vbCrLf
		objfile.write "					<TD style=" & aQuote & "TEXT-ALIGN:  Left; FONT-SIZE: 10px;FONT-FAMILY: Arial, Helvetica; COLOR: #FFFFFF" & aQuote & ">__</TD>"& vbCrLf
		objfile.write "				</TR>"& vbCrLf
		objfile.write "				<TR>"& vbCrLf
		objfile.write "					<TD style=" & aQuote & "VERTICAL-ALIGN: top" & aQuote & "<A href=" & aQuote & BannerLink & aQuote & "><IMG title=" & aQuote & "Liquidity Services Inc." & aQuote & " border=0 name=" & aQuote & "Liquidity Services Inc." & aQuote & " alt=" & aQuote & "Liquidity Services Inc." & aQuote & " src=" & aQuote & BannerPath & aQuote & " width=469 height=60></A> </TD>"& vbCrLf
		objfile.write "				</TR>"& vbCrLf
		objfile.write "	</TBODY>"& vbCrLf
		objfile.write "</Table>"& vbCrLf
	End If
End if



ObjFile.Close
' ===========================
' This section reads out the current Outlook profile and then sets the name of the default Signature
' ===========================
' Use this version to set all accounts
' in the default mail profile
' to use a previously created signature 
 
Call SetDefaultSignature("LSI_Full","")
 
' Use this version (and comment the other) to
' modify a named profile.
'Call SetDefaultSignature _
' ("Signature Name", "Profile Name") 
 
Sub SetDefaultSignature(strSigName, strProfile)
Const HKEY_CURRENT_USER = &H80000001
strComputer = "."
 
If Not IsOutlookRunning Then
		
	Set objreg = GetObject("winmgmts:" & _
	"{impersonationLevel=impersonate}!\\" & _
	strComputer & "\root\default:StdRegProv")
	strKeyPath = "Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\"
	strKeyPath2013 = "Software\Microsoft\Office\15.0\Outlook\Profiles\" '2013 Profile path

	'Get default profile name if none specified
	If strProfile = "" Then
		objreg.GetStringValue HKEY_CURRENT_USER, _
		strKeyPath, "DefaultProfile", strProfile
	End If
	
	'Set default profile name using objects for 2013 Users
	If IsNull(strProfile) Then
		Set ObjOutlook2013  = CreateObject("Outlook.Application")
		strProfile = ObjOutlook2013.DefaultProfileName
		Set ObjOutlook2013 = Nothing
	End If

	' build array from signature name
	myArray = StringToByteArray(strSigName, True)
	strKeyPath = strKeyPath & strProfile & "\9375CFF0413111d3B88A00104B2A6676" 'Outlook 2003/2007/2010 Profile PAth
	strKeyPath2013 = strKeyPath2013 & strProfile & "\9375CFF0413111d3B88A00104B2A6676" '2013 profile path
	objreg.EnumKey HKEY_CURRENT_USER, strKeyPath, arrProfileKeys
	objreg.EnumKey HKEY_CURRENT_USER, strKeyPath2013, arrProfileKeys2013
	
	'Set keys for Outlook 2003/2007/2010
	If Not IsNull(arrProfileKeys) Then
		For Each subkey In arrProfileKeys
			strsubkeypath = strKeyPath & "\" & subkey
			objreg.SetBinaryValue HKEY_CURRENT_USER, strsubkeypath, "New Signature", myArray
			objreg.SetBinaryValue HKEY_CURRENT_USER, strsubkeypath, "Reply-Forward Signature", myArray
		Next
	End If
	
	'Set keys for Outlook 2013
	If Not IsNull(arrProfileKeys2013) Then
		For Each subkey2013 In arrProfileKeys2013
			strsubkeypath2013 = strKeyPath2013 & "\" & subkey2013
			objreg.SetBinaryValue HKEY_CURRENT_USER, strsubkeypath2013, "New Signature", myArray
			objreg.SetBinaryValue HKEY_CURRENT_USER, strsubkeypath2013, "Reply-Forward Signature", myArray
		Next
	End If
Else
	strMsg = "Please shut down Outlook before " & "running this script."
	MsgBox strMsg, vbExclamation, "SetDefaultSignature"
End If
End Sub
 
Function IsOutlookRunning()
strComputer = "."
strQuery = "Select * from Win32_Process " & _
"Where Name = 'Outlook.exe'"
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _
& strComputer & "\root\cimv2")
Set colProcesses = objWMIService.ExecQuery(strQuery)
For Each objProcess In colProcesses
If UCase(objProcess.Name) = "OUTLOOK.EXE" Then
IsOutlookRunning = True
Else
IsOutlookRunning = False
End If
Next
End Function
 
Public Function StringToByteArray _
(Data, NeedNullTerminator)
Dim strAll
strAll = StringToHex4(Data)
If NeedNullTerminator Then
strAll = strAll & "0000"
End If
intLen = Len(strAll) \ 2
ReDim arr(intLen - 1)
For i = 1 To Len(strAll) \ 2
arr(i - 1) = CByte _
("&H" & Mid(strAll, (2 * i) - 1, 2))
Next
StringToByteArray = arr
End Function
 
Public Function StringToHex4(Data)
' Input: normal text
' Output: four-character string for each character,
' e.g. "3204" for lower-case Russian B,
' "6500" for ASCII e
' Output: correct characters
' needs to reverse order of bytes from 0432
Dim strAll
For i = 1 To Len(Data)
' get the four-character hex for each character
strChar = Mid(Data, i, 1)
strTemp = Right("00" & Hex(AscW(strChar)), 4)
strAll = strAll & Right(strTemp, 2) & Left(strTemp, 2)
Next
StringToHex4 = strAll
 
End Function

set objShell = WScript.CreateObject( "WScript.Shell" )
 
' *****************************************************
'This function checks to see if the passed group name contains the current
' user as a member. Returns True or False
Function IsMember(groupName)
    If IsEmpty(groupListD) then
        Set groupListD = CreateObject("Scripting.Dictionary")
        groupListD.CompareMode = 1
        ADSPath = EnvString("userdomain") & "/" & EnvString("username")
        Set userPath = GetObject("WinNT://" & ADSPath & ",user")
        For Each listGroup in userPath.Groups
            groupListD.Add listGroup.Name, "-"
        Next
    End if
    IsMember = CBool(groupListD.Exists(groupName))
End Function
' *****************************************************
 
' *****************************************************
'This function returns a particular environment variable's value.
' for example, if you use EnvString("username"), it would return
' the value of %username%.
Function EnvString(variable)
    variable = "%" & variable & "%"
    EnvString = objShell.ExpandEnvironmentStrings(variable)
End Function
' *****************************************************
 
' Clean up
Set objShell = Nothing