<#
    .SYNOPSIS
    UpdateOutlookSignature.PS1
    A script to Update the Outlook signature on a user computer, based on script found on forum
     - https://community.spiceworks.com/topic/447761-powershell-creating-a-outlook-signature

    Konrad Sagała
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 1.0 July 2021

    .DESCRIPTION
	
    This script creates email signature with data for specific user. 

    .PARAMETER Name
    No parameters used

    .EXAMPLE
    Create user signatures
    .\updateOutlookSig.ps1

#>

#
# Script take user data from environmental variables and Active Directory low level query
#
$strName = $env:username
$strFilter = "(&(objectCategory=User)(samAccountName=$strName))"
$objSearcher = New-Object System.DirectoryServices.DirectorySearcher
$objSearcher.Filter = $strFilter

$objPath = $objSearcher.FindOne()
$objUser = $objPath.GetDirectoryEntry()
#
# when for some reason script cannot locate user object in Active Directory it must be finished immediately
#
if ($null -eq $objUser)
  {
    EXIT
  }

#
# assignment user AD attributes to variables used in signature creation
#
$strWebsite = "www.pepug.org"

$strName = $objUser.FullName
$strFirstName = $objUser.givenName
$strLastName = $objUser.sn
$strTitle = $objUser.Title
$strCompany = $objUser.Company

$strStreet = $objUser.StreetAddress
$strCode = $objUser.postalCode
$strCity =  $objUser.l
$strState = $objUser.st
$strPhone = $objUser.homePhone
$strMainPhone = $objUser.telephonenumber
$strMobile = $objUser.mobile
$strEmail = $objUser.mail
$strteam = $objuser.division
$stroffice = $objuser.physicalDeliveryOfficeName
$strDep = $objUser.department

#
# check Signatures attributes in local registry
#

$UserDataPath = $Env:appdata
if (test-path "HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\General")
  {
    get-item -path HKCU:\\Software\\Microsoft\\Office\\14.0\\Common\\General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
  } 

if (test-path "HKCU:\\Software\\Microsoft\\Office\\15.0\\Common\\General")
  {
    get-item -path HKCU:\\Software\\Microsoft\\Office\\15.0\\Common\\General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
  }
if (test-path "HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\General")
  {
    get-item -path HKCU:\\Software\\Microsoft\\Office\\16.0\\Common\\General | new-Itemproperty -name Signatures -value signatures -propertytype string -force
  }
$FolderLocation = $UserDataPath + '\\Microsoft\\signatures'  
mkdir $FolderLocation -force
$signaturefile = "BMsignature2021"

#
# Creates HtmlSignature
#
$stream = [System.IO.StreamWriter] "$FolderLocation\\$signaturefile.htm"
#    $stream.WriteLine("<!DOCTYPE HTML PUBLIC `"-//W3C//DTD HTML 4.0 Transitional//EN`">")
$stream.WriteLine("<html>")
$stream.WriteLine("<head><title>Signature</title>")
$stream.WriteLine("<meta http-equiv=Content-Type content=`"text/html; charset=UTF-8`">")
$stream.WriteLine("<style type=`"text/css`">")
$stream.WriteLine("<!--")
$stream.WriteLine("A:link { COLOR: #0000A0; TEXT-DECORATION: none; font-weight: normal }")
$stream.WriteLine("A:visited { COLOR: #0000A0; TEXT-DECORATION: none; font-weight: normal }")
$stream.WriteLine("A:active { COLOR: black; TEXT-DECORATION: none }")
$stream.WriteLine("A:hover { COLOR: blue; TEXT-DECORATION: none; font-weight: none }")
$stream.WriteLine("-->")
$stream.WriteLine("</style>")
$stream.WriteLine("</head>")
$stream.WriteLine("<body lang=PL>")
$stream.WriteLine("<div style=`"line-height:normal; font-family: 'Trebuchet MS'; font-size:11px; color:#333333;`">")
$stream.WriteLine("$strFirstName $strLastName $strTitle<br>")
$stream.WriteLine("<a href=`"mailto:$strEmail`">$strEmail</a> | ")
$stream.WriteLine("<a href=`"$strWebsite`">$strWebsite</a><br>")
$stream.WriteLine("<span style=`"font-size:8.0pt;font-family:'Trebuchet MS','sans-serif';color:#CD0067'`">$strCompany</span><br>")
#$stream.WriteLine("$strDep | ")
$stream.WriteLine("$strteam<br>")
$stream.WriteLine("$strStreet, $strCode $strCity<br>")
if ($strMainPhone.length -gt 0)
{
    $stream.WriteLine("tel: $strMainPhone | ")
}
if ($strMobile.length -gt 0)
{
    $stream.WriteLine("mob: $strMobile<br>")
}
$stream.WriteLine("<br>")
$stream.WriteLine("<span><img border=0 width=402 height=133 src=`"https://www.pepug.org/Stopka.png`"></span>")
$stream.WriteLine("<br>")
$stream.WriteLine("</div>")
$stream.close()

#
# Creates RTF Signature
#
$wrd = new-object -com word.application 

# Make Word Visible 
$wrd.visible = $false

# Open a document  
$fullPath = $FolderLocation+"\$signaturefile.htm"
$doc = $wrd.documents.open($fullpath) 

# Save as rtf
$opt = 6
$name = $FolderLocation+"\$signaturefile.rtf"
$wrd.ActiveDocument.Saveas($name,$opt)

#
# Set company signature as default for New messages/Reply Messages
#
$EmailOptions = $wrd.EmailOptions
$EmailSignature = $EmailOptions.EmailSignature
#$EmailSignatureEntries = $EmailSignature.EmailSignatureEntries
$EmailSignature.NewMessageSignature=$signaturefile
$EmailSignature.ReplyMessageSignature=$signaturefile

# Close word
$wrd.Quit()

#
# Create Sigture Text File
#
$stream = [System.IO.StreamWriter] "$FolderLocation\\$signaturefile.txt"
$stream.WriteLine("$strFirstName $strLastName")
$stream.WriteLine("$strEmail $strWebsite")
$stream.WriteLine("$StrCompany")
$stream.WriteLine("$strTitle")
$stream.WriteLine(" ")
$stream.WriteLine("$strStreet, $strCode $strCity")
if ($strMainPhone.length -gt 0)
{
    $stream.WriteLine("tel: $strMainPhone | ")
}
if ($strMobile.length -gt 0)
{
    $stream.WriteLine("mob: $strMobile<br>")
}
$stream.close()

#
# EOF
#
