<#
.SYNOPSIS
A customized Exchange Management Shell for SSH connections.

.NOTES
    author: Geoffrey.Duke@uvm.edu
    Based on: "E:\exchangeserver\Bin\RemoteExchange.ps1"

    Changes:
     - Updated with Exchange 2016 RemoteExchange.ps1
     - Remove Console Window widening function
     - Relocated localization to after some global paths are defined
#>

# Copyright (c) Microsoft Corporation. All rights reserved.  

## INCREASE WINDOW WIDTH #####################################################
# >>> Removed - not needed for SSH consoles

## ALIASES ###################################################################

set-alias list       format-list 
set-alias table      format-table 

## Confirmation is enabled by default, uncommenting next line will disable it 
# $ConfirmPreference = "None"

## EXCHANGE VARIABLEs ########################################################

$global:exbin = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath + "bin\"
$global:exinstall = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath
$global:exscripts = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath + "scripts\"

## Relocated from top of script to take advantage of the path variables above 
#load hashtable of localized string
Import-LocalizedData -BindingVariable RemoteExchange_LocalizedStrings -BaseDirectory $global:exbin -FileName RemoteExchange.strings.psd1
# 

## LOAD CONNECTION FUNCTIONS #################################################

# ConnectFunctions.ps1 uses some of the Exchange types. PowerShell does some type binding at the 
# time of loading the scripts, so we'd rather load the scripts before we reference those types.
"Microsoft.Exchange.Data.dll", "Microsoft.Exchange.Configuration.ObjectModel.dll" `
  | ForEach { [System.Reflection.Assembly]::LoadFrom((join-path $global:exbin $_)) } `
  | Out-Null

. $global:exbin"CommonConnectFunctions.ps1"
. $global:exbin"ConnectFunctions.ps1"

## LOAD EXCHANGE EXTENDED TYPE INFORMATION ###################################

$FormatEnumerationLimit = 16

# loads powershell types file, parses out just the type names and returns an array of string
# it skips all template types as template parameter types individually are defined in types file
function GetTypeListFromXmlFile( [string] $typeFileName ) 
{
	$xmldata = [xml](Get-Content $typeFileName)
	$returnList = $xmldata.Types.Type | where { (($_.Name.StartsWith("Microsoft.Exchange") -or $_.Name.StartsWith("Microsoft.Office.CompliancePolicy")) -and !$_.Name.Contains("[[")) } | foreach { $_.Name }
	return $returnList
}

# Check if every single type from from Exchange.Types.ps1xml can be successfully loaded
$typeFilePath = join-path $global:exbin "exchange.types.ps1xml"
$typeListToCheck = GetTypeListFromXmlFile $typeFilePath
# Load all management cmdlet related types.
$assemblyNames = [Microsoft.Exchange.Configuration.Tasks.CmdletAssemblyHelper]::ManagementCmdletAssemblyNames
$typeLoadResult = [Microsoft.Exchange.Configuration.Tasks.CmdletAssemblyHelper]::EnsureTargetTypesLoaded($assemblyNames, $typeListToCheck)
# $typeListToCheck is a big list, release it to free up some memory
$typeListToCheck = $null

$SupportPath = join-path $global:exbin "Microsoft.Exchange.Management.Powershell.Support.dll"
[Microsoft.Exchange.Configuration.Tasks.TaskHelper]::LoadExchangeAssemblyAndReferences($SupportPath) > $null

if (Get-ItemProperty HKLM:\Software\microsoft\ExchangeServer\v15\CentralAdmin -ea silentlycontinue)
{
    $CentralAdminPath = join-path $global:exbin "Microsoft.Exchange.Management.Powershell.CentralAdmin.dll"
    [Microsoft.Exchange.Configuration.Tasks.TaskHelper]::LoadExchangeAssemblyAndReferences($CentralAdminPath) > $null
}

if (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapins\Microsoft.Exchange.Management.AntiSpamTasks -ea silentlycontinue)
{
    $AntiSpamTasksPath = join-path $global:exbin "Microsoft.Exchange.Management.AntiSpamTasks.dll"
    [Microsoft.Exchange.Configuration.Tasks.TaskHelper]::LoadExchangeAssemblyAndReferences($AntiSpamTasksPath) > $null
}

if (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\PowerShell\1\PowerShellSnapins\Microsoft.Exchange.Management.PeopleICommunicateWith -ea silentlycontinue)
{
    $picwTasksPath = join-path $global:exbin "Microsoft.Exchange.Management.PeopleICommunicateWith.dll"
    [Microsoft.Exchange.Configuration.Tasks.TaskHelper]::LoadExchangeAssemblyAndReferences($picwTasksPath) > $null
}

# Register Assembly Resolver to handle generic types
[Microsoft.Exchange.Data.SerializationTypeConverter]::RegisterAssemblyResolver()


# Finally, load the types information
# We will load type information only if every single type from Exchange.Types.ps1xml can be successfully loaded
if ($typeLoadResult)
{
	Update-TypeData -PrependPath $typeFilePath
}
else
{
	write-error $RemoteExchange_LocalizedStrings.res_types_file_not_loaded
}

#load partial types
$partialTypeFile = join-path $global:exbin "Exchange.partial.Types.ps1xml"
Update-TypeData -PrependPath $partialTypeFile 

# If Central Admin cmdlets are installed, it loads the types information for those too
if (Get-ItemProperty HKLM:\Software\microsoft\ExchangeServer\v15\CentralAdmin -ea silentlycontinue)
{
	$typeFile = join-path $global:exbin "Exchange.CentralAdmin.Types.ps1xml"
	Update-TypeData -PrependPath $typeFile
}

# Loads FFO-specific type and formatting xml files.
$ffoTypeData = join-path $global:exbin "Microsoft.Forefront.Management.Powershell.types.ps1xml"
$ffoFormatData = join-path $global:exbin "Microsoft.Forefront.Management.Powershell.format.ps1xml"

if ((Test-Path $ffoTypeData) -and (Test-Path $ffoFormatData))
{
    Update-TypeData -PrependPath $ffoTypeData
    Update-FormatData -PrependPath $ffoFormatData
}

## FUNCTIONs #################################################################

## returns all defined functions 

function functions
{ 
    if ( $args ) 
    { 
        foreach($functionName in $args )
        {
             get-childitem function:$functionName | 
                  foreach { "function " + $_.Name; "{" ; $_.Definition; "}" }
        }
    } 
    else 
    { 
        get-childitem function: | 
             foreach { "function " + $_.Name; "{" ; $_.Definition; "}" }
    } 
}

## only returns exchange commands 

function get-excommand
{
	if ($args[0] -eq $null)
	{
		get-command -module $global:importResults
	}
	else
	{
		get-command $args[0] | where { $_.module -eq $global:importResults }
	}
}


## only returns PowerShell commands 

function get-pscommand
{
	if ($args[0] -eq $null) 
	{
		get-command -pssnapin Microsoft.PowerShell* 
	}
	else 
	{
		get-command $args[0] | where { $_.PsSnapin -ilike 'Microsoft.PowerShell*' }	
	}
}

## prints the Exchange Banner in pretty colors 

function get-exbanner
{
	write-host $RemoteExchange_LocalizedStrings.res_welcome_message

	write-host -no $RemoteExchange_LocalizedStrings.res_full_list
	write-host -no " "
	write-host -fore Yellow $RemoteExchange_LocalizedStrings.res_0003

	write-host -no $RemoteExchange_LocalizedStrings.res_only_exchange_cmdlets
	write-host -no " "
	write-host -fore Yellow $RemoteExchange_LocalizedStrings.res_0005

	write-host -no $RemoteExchange_LocalizedStrings.res_cmdlets_specific_role
	write-host -no " "
	write-host -fore Yellow $RemoteExchange_LocalizedStrings.res_0007

	write-host -no $RemoteExchange_LocalizedStrings.res_general_help
	write-host -no " "
	write-host -fore Yellow $RemoteExchange_LocalizedStrings.res_0009

	write-host -no $RemoteExchange_LocalizedStrings.res_help_for_cmdlet
	write-host -no " "
	write-host -fore Yellow $RemoteExchange_LocalizedStrings.res_0011

	write-host -no $RemoteExchange_LocalizedStrings.res_team_blog
	write-host -no " "
	write-host -fore Yellow $RemoteExchange_LocalizedStrings.res_0015

	write-host -no $RemoteExchange_LocalizedStrings.res_show_full_output
	write-host -no " "
	write-host -fore Yellow $RemoteExchange_LocalizedStrings.res_0017

	write-host -no $RemoteExchange_LocalizedStrings.res_updatable_help
	write-host -no " "
	write-host -fore Yellow $RemoteExchange_LocalizedStrings.res_0018
}

## shows quickref guide

function quickref
{
     invoke-expression 'cmd /c start http://go.microsoft.com/fwlink/p/?LinkId=259608'
}

function get-exblog
{
       invoke-expression 'cmd /c start http://go.microsoft.com/fwlink/?LinkId=35786'
}

## FILTERS #################################################################
## Assembles a message and writes it to file from many sequential BinaryFileDataObject instances 
Filter AssembleMessage ([String] $Path) { Add-Content -Path:"$Path" -Encoding:"Byte" -Value:$_.FileData }

## now actually call the functions 

get-exbanner 
get-tip 

#
# TIP: You can create your own customizations and put them in My Documents\WindowsPowerShell\profile.ps1
# Anything in profile.ps1 will then be run every time you start the shell. 
#

