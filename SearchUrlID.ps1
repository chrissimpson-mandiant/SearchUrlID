<#
.SYNOPSIS
  Script for returning the subject line of messages based on the OWA Url ID
.DESCRIPTION
  This script accepts a text file with OWA URL ID strings and searches a specified mailbox for those strings.
.PARAMETER EmailAddress
  The full primary SMTP email address for the mailbox to be searched.  The parameter is REQUIRED.
.PARAMETER OutputDirectory
  The directory location where the results file should be saved.  The default output directory is the working directory.
.PARAMETER ResultFile
  The name of the results file.  If not specified, the results will be saved as results.csv.
.PARAMETER QueryList
  The text file containing the OWA URL ID strings.  See the INPUTS section of the FULL HELP for text file format instructions.  The parameter is REQUIRED.  
.PARAMETER EWSUrl
  Allows a specific EWS URL to be used.  If this value is not specified, then it will be found via the AutoDiscover process.
  It is recommended that this URL be specified if known (speed increase).
.PARAMETER Credentials
  Prompts the user for credentials for binding to Exchange.  If this flag is not specified, binding will occur using the session credentials.
.INPUTS
  A normal OWA URL string:  https://mail.domain.com/owa/?ae=Item&a=Open&t=IPM.Note&id=RgAAAABGaVq9H6o2QpBZtULPn1i9ByDhZ%2cEaQ80XT5rMVYu4jJVcAAAA283aAABVh9i1TPaUQpCja5K%2fzS1TAABL%2bs%2bJAAAJ&pspid=_1454265733661_990260427
  
  The relevant Query ID for the purposes of this script is the string BETWEEN &id= and &pspid=
  
  In this example, our Query ID is: RgAAAABGaVq9H6o2QpBZtULPn1i9ByDhZ%2cEaQ80XT5rMVYu4jJVcAAAA283aAABVh9i1TPaUQpCja5K%2fzS1TAABL%2bs%2bJAAAJ
  
  ******QueryList File Format******
  The input file (QueryList) should be a plain text document.  The first line MUST begin with ID.  Each subsequent line should contain a Query ID value.
  *********************************
  
  ****EXAMPLE****
  ID
  RgAAAABGaVq9H6o2QpBZtULPn1i9ByDhZ%2cEaQ80XT5rMVYu4jJVcAAAA283aAABVh9i1TPaUQpCja5K%2fzS1TAABL%2bs%2bJAAAJ
  RgAAAABGaDhZ%2cEaQ80zS1TAABL%2bXT5rMVYu4jJVcAAAA283aAABVh9Vq9H6o2QpBZtULPn1i9Byi1TPaUQpCja5K%2fs%2bJAAAJ
  **END EXAMPLE**
.NOTES
  Version:        1.0
  Author:         Chris Simpson
  Creation Date:  2/23/2016
  Purpose/Change: Initial script development
  
  REQUIRED DEPENDENCY INSTALL:  Microsoft Exchange Web Services Managed API 2.2
  https://www.microsoft.com/en-us/download/details.aspx?id=42951
  
.EXAMPLE
  SearchUrlID.ps1 -EmailAddress "bill.gates@contoso.com" -QueryList C:\IDs.txt -Credentials -EWSUrl https://mail.contoso.com/ews/Exchange.asmx
  Searches the mailbox bill.gates@contoso.com for each of the URL ID strings in the C:\IDs.txt file.  The syntax also prompts for credentials and manually specifies the EWS URL (rather than using AutoDiscover).  The results are saved in the working directory as results.csv.
.EXAMPLE
  SearchUrlID.ps1 -EmailAddress "freddy.krueger@elmstreet.com" -QueryList C:\IDs.txt
  Searches the mailbox freddy.krueger@elemstreet.com for each of the URL ID strings in the C:\IDs.txt file.  Uses the session credentials to bind to Exchange and uses AutoDiscovery to find the EWS URL.  The results are saved in the working directory as results.csv.
.EXAMPLE
  SearchUrlID.ps1 -EmailAddress "bill.gates@contoso.com" -QueryList C:\IDs.txt -Credentials -EWSUrl https://mail.contoso.com/ews/Exchange.asmx -OutputDirectory "C:\Output" -ResultFile output.csv
  Searches the mailbox bill.gates@contoso.com for each of the URL ID strings in the C:\IDs.txt file.  The syntax also prompts for credentials and manually specifies the EWS URL (rather than using AutoDiscover).  The results are saved in the C:\Output directory as output.csv.
#>

Param (
    [String]
    [Parameter( Mandatory = $True )]
    $EmailAddress,
    
    [String]
    $OutputDirectory = $PWD,

    [String]
    $ResultFile = 'results.csv',
    
	[String]
	[Parameter( Mandatory = $True )]
	$QueryList,
	
	[String]
	$EWSUrl = '',
	
	[switch]
	$Credentials
)

function Encode-URL($rawURL)
{
	return [System.Web.HttpUtility]::UrlEncode($rawURL)
}

function Search-UrlID {

Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll" #https://www.microsoft.com/en-us/download/details.aspx?id=42951
$exchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2, [System.TimeZoneInfo]::Local)

if($Credentials){
	$creds = (Get-Credential).GetNetworkCredential()
	$exchService.Credentials = New-Object System.Net.NetworkCredential -ArgumentList $creds.UserName, $creds.Password, $creds.Domain
}

if($EWSUrl -eq ''){
	$exchService.AutodiscoverUrl($EmailAddress)
}
else{
	$exchService.Url = $EWSUrl
}

$customProps = new-object Microsoft.Exchange.WebServices.Data.PropertySet

$customProps.add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject)
$customProps.add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived)
$customProps.add([Microsoft.Exchange.WebServices.Data.ItemSchema]::ConversationId)
$customProps.add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Id)
$customProps.add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender)
$customProps.add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::WebClientReadFormQueryString)

$itemView = new-object Microsoft.Exchange.WebServices.Data.ItemView(1000,0,[Microsoft.Exchange.WebServices.Data.OffsetBasePoint]::Beginning)
$itemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Shallow
$itemView.PropertySet=$customProps

$PR_FOLDER_TYPE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$folderidcnt = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$EmailAddress)  

$fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000) #Should not be any larger then 1000 folders due to throttling  
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep #Deep Traversal returns all subfolders in mailbox

$searchFilterCollection1 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::Or)
$searchFilterCollection2 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)

$searchFilter1 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ItemClass, "IPM.Note")
$searchFilter2 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_FOLDER_TYPE,"1") #Excludes Search Folders
$searchFilter3 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass, "IPF.Note")
$searchFilter4 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ItemClass, "IPM.Note.SMIME.MultipartSigned")
$searchFilter5 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ItemClass, "IPM.Note.SMIME")

$searchFilterCollection1.add($searchFilter1)
$searchFilterCollection1.add($searchFilter4)
$searchFilterCollection1.add($searchFilter5)
$searchFilterCollection2.Add($searchFilter2)
$searchFilterCollection2.Add($searchFilter3)

$fiResult = $null  

$return_object =@()

do{  
    $fiResult = $exchService.FindFolders($folderidcnt,$searchFilterCollection2,$fvFolderView)  
    foreach($ffFolder in $fiResult.Folders){
		do{
			$FindItems = $exchService.FindItems($ffFolder.Id,$searchFilterCollection1,$itemView)
			foreach ($eItems in $FindItems.Items){
				$props = @{ DateTimeReceived	= $eItems.DateTimeReceived;
							ID					= $eItems.ID;
							ConversationId		= $eItems.ConversationId;
							Sender				= $eItems.Sender.Name;
							Subject				= $eItems.Subject;
							QueryString			= $eItems.WebClientReadFormQueryString;
							}
				$return_object += New-Object -TypeName PSCustomObject -Property $props
			}
			$itemView.Offset= $itemView.Offset + $itemView.PageSize
		}while($FindItems.MoreAvailable)
        $itemView.Offset = 0
	}
    $fvFolderView.Offset += $fiResult.Folders.Count
}while($fiResult.MoreAvailable -eq $true)

$values = Import-Csv -Path $QueryList
Foreach($v in $values)
{
	$query = $v.id.Substring(0,$v.id.Length-1)
	$result = $return_object | where-object {($_.QueryString -Clike "*$query*")}
	if ($result -eq $null){
		$notfound += $v.id + "`n"
	}
	else{
		$result | Export-Csv -Path (Join-Path $OutputDirectory $ResultFile) -Append -NoTypeInformation -Force
	}
	$result = $null
}

Write-Host "The following ID values were not found.  The messages have likely been moved or deleted." -ForegroundColor Red
$notfound

}

Search-UrlID $EmailAddress $QueryList -OutputDirectory $OutputDirectory -ResultFile $ResultFile -Credential:$Credentials -EWSUrl $EWSUrl
