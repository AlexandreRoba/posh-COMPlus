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
		[guid[]]$Key,
		[Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true,ParameterSetName='ApplicationSet')]
		[PowerShellComPlus.Application[]]$Application
	)
	
	$comAdmin = New-Object -ComObject "COMAdmin.COMAdminCatalog"
	$objCollection = $comAdmin.GetCollection("Applications")
	$objCollection.Populate()
	
	if ($PsCmdlet.ParameterSetName -eq 'ApplicationSet'){
		foreach($objapplication in $objCollection){
			foreach($a in $Application){
				if (($a.Key -eq $objapplication.Key) -and ($a.Name -eq $objapplication.Name)){
					$Properties = @{Key=$objapplication.Key;Name=$objapplication.Name}
					$ApplicationFound = New-Object -TypeName PowerShellComPlus.Application -Property $Properties
					Write-Output($ApplicationFound)
				}
			}
		}
	}
	else{
		foreach($objapplication in $objCollection){
			$hit = $false
			foreach($n in $Name){
				if($objapplication.Name -like $n){
					$hit = $true
					break
				}
			}
			if($Name -and -not $hit){
				continue
			}
			$hit = $false
			foreach($k in $Key){
				if($objapplication.Key -eq $k){
					$hit = $true
					break
				}
			}
			if($Key -and -not $hit){
				continue
			}
			$Properties = @{Key=$objapplication.Key;Name=$objapplication.Name}
			$ApplicationFound = New-Object -TypeName PowerShellComPlus.Application -Property $Properties
			Write-Output($ApplicationFound)
		}
	}
}
