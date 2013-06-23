<#
.SYNOPSIS
	Gets the applications from COM+
.DESCRIPTION
	Gets the applications from COM+. Applications can be filtered by Name, Key
.PARAMETER Name
	Filter on the application name. Supports Wildcards.
.PARAMETER Key
	Filter on the application Key.
.INPUTS
	[PowerShellComPlus.Application[]]
.EXAMPLE
	c:\PS> Get-COMPlusApplication
	
	This example returns all Application from COM+
#>
function Get-COMPlusApplication{
	[CmdletBinding(DefaultParameterSetName= 'PartsSet')]
	[OutputType('PowerShellComPlus.Application')]
	param(
		[Parameter(Position=0, ParameterSetName='PartsSet')]
		[ValidateNotNullOrEmpty()]
		[string[]]$Name,
		[Parameter(Position=1, ParameterSetName='PartsSet')]
		[ValidateNotNullOrEmpty()]
		[string[]]$Key,
		[Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true,ParameterSetName='ApplicationSet')]
		[PowerShellComPlus.Application[]]$Applications
		
	)
	
	$comAdmin = New-Object -ComObject "COMAdmin.COMAdminCatalog"
	
	$objCollection = $comAdmin.GetCollection("Applications")
	$objCollection.Populate()
	foreach($objapplication in $objCollection){
		$Properties = @{Key=$objapplication.Key;Name=$objapplication.Name}
		$Application = New-Object -TypeName PowerShellComPlus.Application -Property $Properties
		Write-Output($Application.Name)
	}
	
}