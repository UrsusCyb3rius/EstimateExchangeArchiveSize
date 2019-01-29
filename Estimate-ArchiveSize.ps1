<#  
.SYNOPSIS  
    Queries mailbox and returns the total item size of items older than the expiration date (configurable through AgeLimit parameter). 
    This should allow to estimate the amount of data will be transferred to the archives once activated.
 
.DESCRIPTION  
    Queries mailbox and returns the total item size of items older than the expiration date (configurable through AgeLimit parameter). 
    This should allow to estimate the amount of data will be transferred to the archives once activated.

    Use the -verbose switch to add more detail in the active window.
    Use the -credential switch to specify a set of credentials to use for impersonation that are different from the logged on ones.

       
.NOTES
    Version                 : 1.0
    Rights Req'd            : Application Impersonation
    Other Requirements      : Turn of throttling for the user executing the script
    Sch Task Req            : No
    Exchange Ver            : tested with 2010/2013
    Author                  : Michael Van Horenbeeck, Exchange Server MVP
    Co-Author               : Michel de Rooij, Exchange Server MVP
    Email/Blog/Twitter      : michael@vanhorenbeeck.be - @mvanhorenbeeck
    Blog                    : http://michaelvh.wordpress.com
    Disclaimer              : Use this script at your own risk!
    Special thanks to       : Michel de Rooij (Eightwone) - http://www.eightwone.com
                              Serkan Varoglu (Get-RetentionExpiration.ps1) - http://www.get-mailbox.org
                              Glen Scales (General EWS Stuff) - http://gsexdev.blogspot.com
 
.LINK  
    http://michaelvh.wordpress.com
 
.EXAMPLE
    To estimate the archive size for a single mailbox, use the following syntax:
        .\Estimate-ArchiveSize.ps1 -UserPrimarySMTPAddress user.a@domain.com -AgeLimit 14

    To estimate the archive size for multiple mailboxes, use:
        $mailboxes = "user.a@domain.com","user.b@domain.com"
        .\Estimate-ArchiveSize.ps1 $mailboxes -AgeLimit 14

    To specify a report file, use the following syntax:
        .\Estimate-ArchiveSize.ps1 -UserPrimarySMTPAddress user.a@domain.com -Report c:\reports\archivesizes.txt

#>
 
 
[CmdletBinding()]
[OutputType([int])]
Param
(
    [Parameter(Mandatory=$true, 
               ValueFromPipeline=$true,
               ValueFromPipelineByPropertyName=$false, 
               ValueFromRemainingArguments=$false, 
               Position=0,
               HelpMessage="The primary SMTP address of the user's mailbox")]
    [ValidateNotNullOrEmpty()]
        [array]$UserPrimarySMTPAddresses,
    [Parameter(Mandatory=$false, 
               ValueFromPipeline=$false,
               ValueFromPipelineByPropertyName=$false, 
               ValueFromRemainingArguments=$false, 
               Position=1,
               HelpMessage="The full file path where to store the report file.")]
    [ValidateNotNullOrEmpty()]
        [string]$Report,
    [Parameter(Mandatory=$false, 
               ValueFromPipeline=$false,
               ValueFromPipelineByPropertyName=$false, 
               ValueFromRemainingArguments=$false, 
               Position=2,
               HelpMessage="The item's age limit for retention to take into account.")]
    [ValidateNotNullOrEmpty()]
        [int]$AgeLimit=0,
	[parameter( Mandatory=$false, HelpMessage="Server FQDN to use for EWS")]
		[string]$Server,
	[parameter( Mandatory=$false, HelpMessage="Credentials for Impersonation")]
	    [System.Net.NetworkCredential]$Credentials
)
   
Function Write-Log($err, $LogText){
  $LogString="$(Get-Date) "
  switch ($err) {
    $logInfo    { $LogString += "Info   "; break }
    $logWarning { $LogString += "Warning"; break }
    $logError   { $LogString += "Error  "; $global:errCount++; break }
    default     { $LogString += "       " }
    }
  $LogString += " - $LogText"
  $LogString | Out-File -FilePath $LogFile -Append
}
 
# After calling this any SSL Warning issues caused by Self Signed Certificates will be ignored
Function set-TrustAllWeb() {
    # Source: http://poshcode.org/624
    Write-Verbose "Set to trust all certificates"
    $Provider=New-Object Microsoft.CSharp.CSharpCodeProvider  
    $Compiler=$Provider.CreateCompiler()  
    $Params=New-Object System.CodeDom.Compiler.CompilerParameters  
    $Params.GenerateExecutable=$False  
    $Params.GenerateInMemory=$True  
    $Params.IncludeDebugInformation=$False  
    $Params.ReferencedAssemblies.Add("System.DLL") | Out-Null
    $TASource= @'
            namespace Local.ToolkitExtensions.Net.CertificatePolicy { 
                public class TrustAll : System.Net.ICertificatePolicy { 
                    public TrustAll() {  
                    }
                    public bool CheckValidationResult(System.Net.ServicePoint sp, System.Security.Cryptography.X509Certificates.X509Certificate cert,   System.Net.WebRequest req, int problem) { 
                        return true; 
                    } 
                } 
            }
'@
    $TAResults=$Provider.CompileAssemblyFromSource($Params, $TASource)  
    $TAAssembly=$TAResults.CompiledAssembly  
    $TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")  
    [System.Net.ServicePointManager]::CertificatePolicy=$TrustAll  
}

Function Load-EWSDLL() {
    Write-Verbose "Loading EWS Managed API DLL"
    $EwsDllFile= "Microsoft.Exchange.WebServices.dll"
	$EwsDllKey= Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKLM\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name
	If( $EwsDllKey -eq $null) {
		# Not installed? Try script folder
	    $EwsDllPath= split-path -Path $MyInvocation.ScriptName
	}
    Else {
        $EwsDllPath= (Get-ItemProperty -ErrorAction SilentlyContinue "HKLM:\Software\Microsoft\Exchange\Web Services\2.0")."Install Directory"
    }
    $EwsFullPath= Join-Path $EwsDllPath $EwsDllFile
	If( ($EwsDllPath -ne $null) -and (Test-Path $EwsFullPath)) {
        Write-Verbose "Loading $EwsFullPath"
		try {
        	[void][Reflection.Assembly]::LoadFile( $EwsFullPath)
		    return $true
		}
		catch {
			Write-Error "Issue loading $($EwsFullPath): $($error[0])"
		    return $false
	    }
    }
    Else {
		Write-Error "Can't locate $EwsDllFile at $EwsDllPath"
        return $false
    }
}

Function get-TotalPRMessageSizeFromFolder ( $Folder) {

    $FolderPRMessageSize= 0
    Write-Verbose "Processing folder $($Folder.DisplayName)"

    #Define the ItemView used, should not be any larger then 1000 folders due to throttling
    $ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000);
    $ivItemView.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)

    #PR_MESSAGE_SIZE
    $PRMessageSize= New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0E08, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
    $ivItemView.PropertySet.Add( $PRMessageSize) 

    #PR_CREATION_TIME
    $PR_CREATION_TIME = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3007,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime);
    $ivItemView.PropertySet.Add( $PR_CREATION_TIME);

    #When AgeLimit provided, define Search Filter using PR_CREATION_TIME
    If( $AgeLimit) {
        $ivItemSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo($PR_CREATION_TIME, $ExpirationDate);
        $fiItemResult = $Service.FindItems($Folder.Id, $ivItemSearchFilter, $ivItemView)
    }
    Else {
        $fiItemResult = $Service.FindItems($Folder.Id, $ivItemView)
    }

    If( $fiItemResult) {
        ForEach( $Item in $fiItemResult) {
            $ItemSize= 0
            If( $Item.tryGetProperty( $PRMessageSize, [ref]$ItemSize) -and ($ItemSize -ne $null)) {
                $FolderPRMessageSize+= $ItemSize
                Write-Debug "$($Item.Subject) PR_MESSAGE_SIZE:$ItemSize Size:$($Item.Size)"
            }
            Else {
                Write-Error "Couldn't retrieve PR_MESSAGE_SIZE of item $($Item.Subject)"
            } 
        }
    } 
    return $FolderPRMessageSize
}

############################################################################
# Main script starts here ..
############################################################################

#Requires -Version 2.0

# Timestamp script start 
$startDTM = (Get-Date)
 
#Determine Log file name
$LogFile = "Estimate-ArchiveSize_"+(get-date -UFormat "%Y-%m-%d")+".log"

# Determine expiration timestamp
$ExpirationDate = (Get-Date).AddDays(-$AgeLimit)

# Process each identity
$UserPrimarySMTPAddresses | ForEach {

    $UserPrimarySMTPAddress= $_

    Write-Log "" "Starting estimation for $UserPrimarySMTPAddresses"
 
    # Load EWSDLL (from installation path or current folder)
    If (!( Load-EWSDLL ) ) {
            Write-Error "Problem loading EWS Managed API DLL"
            Exit 1001
    }

    # Set minimum required Exchange version
    $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2007_SP1

    # Create Exchange Service Object
    $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
 
    # Set Credentials to use two options are availible: use explict credentials (Get-Credential) or use the default (logged On) credentials
    If($Credentials) {
        $Service.Credentials = $creds
    }
    Else {
        $Service.UseDefaultCredentials= $true
    }
    $Service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$UserPrimarySMTPAddress)

    # Set to ignore certificate warnings
    set-TrustAllWeb
 
    # Use autodiscover or provided server FQDN
    If ($Server) {
        $EwsUrl= "https://$Server/EWS/Exchange.asmx"
        $service.Url= "$EwsUrl"
    }
    Else {
        Write-Verbose "Looking up EWS URL using Autodiscover for $UserPrimarySMTPAddress"
        try {
            # Set script to terminate on all errors (autodiscover failure isn't) to make try/catch work
            $ErrorActionPreference= "Stop"
            $Service.autodiscoverUrl( $UserPrimarySMTPAddress, {$true})
        }
        catch {
            Write-Error "Autodiscover failed: " $error[0]
            Exit 1003
        }
        $ErrorActionPreference= "Continue"
    } 

    $Service.keepalive = $False
    Write-Verbose "Using EWS on CAS $($Service.Url)"

    try {
        $RootFolder= [Microsoft.Exchange.WebServices.Data.Folder]::Bind( $Service, [Microsoft.Exchange.WebServices.Data.WellknownFolderName]::MsgFolderRoot)
    }
    catch {
        Write-Error "Can't access mailbox information store"
        Exit 1004
    }

    $TotalPRMessageSize= 0

    $fvFolderView= New-Object Microsoft.Exchange.WebServices.Data.FolderView( 1000)
    $fvFolderView.Traversal= [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
    $fvFolderView.PropertySet= New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)

    #Folder Type (no Search Folders)
    $PR_FOLDER_TYPE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);
    $sfSearchFilter= New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_FOLDER_TYPE,"1")  

    Do {
        $fiResult = $Service.FindFolders( $RootFolder.Id, $sfSearchFilter, $fvFolderView)  
        If( $fiResult) {
            ForEach( $Folder in $fiResult) {
                $FolderPRMessageSize= get-TotalPRMessageSizeFromFolder $Folder
                Write-Debug "Folder $($Folder.DisplayName) PR_MESSAGE_SIZE: $FolderPRMessageSize"
                $TotalPRMessageSize+= $FolderPRMessageSize
            }
        }
        $fvFolderView.Offset += $fiResult.Folders.Count
        $fiAllResult += $fiResult
    } While($fiResult.MoreAvailable)

    $TotalPRMessageSize= [int]($TotalPRMessageSize/1MB)
    Write-Verbose "Total PR_MESSAGE_SIZE $($TotalPRMessageSize)MB"

    $object = New-Object –TypeName PSObject
    $object | Add-Member –MemberType NoteProperty –Name Mailbox –Value $UserPrimarySMTPAddress
    $object | Add-Member –MemberType NoteProperty –Name TotalPRMessageSize –Value $TotalPRMessageSize
    Write-Output $object

    If( $Report) {
#        "$UserPrimarySMTPAddress;$(Object.TotalPRMessageSize)" | Out-File $ReportFile -Append
        Out-File -FilePath $Report -Append -InputObject $object
    }
}
 
# Log time elapsed
Write-Log "" "Elapsed Time: $(((Get-Date)-$startDTM).totalseconds) seconds"