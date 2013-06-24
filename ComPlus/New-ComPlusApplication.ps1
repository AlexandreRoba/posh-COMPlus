<#
.SYNOPSIS
	Create a new COM+ Application.
.DESCRIPTION
    Create a new COM+ Application.
.PARAMETER Name
    The COM+ plus application name.
.INPUTS
	[string]
.EXAMPLE
    C:\PS> New-ComPlusApplication Application01

    This example creates a new COM+ Application named Application01.
#>

function New-ComPlusApplication{
    [CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'ParameterSet')]
	[OutputType('PowerShellComPlus.Application')]
    param(
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = 'ParameterSet')]
        [ValidateNotNullOrEmpty()]
        [string[]] $Name,
        [Switch] $PassThru
    )
    process{
        if ($PsCmdlet.ParameterSetName -eq 'ParameterSet'){
               
        }
        foreach ($n in $Name){
            if (!$PSCmdLet.ShouldProcess($n))
            {
                continue
            }

            $comAdmin = New-Object -ComObject "COMAdmin.COMAdminCatalog"
	        $objCollection = $comAdmin.GetCollection("Applications")
	        $objCollection.Populate()
            $NewApplication = $objCollection.Add()
            $NewApplication.Value("Name") = $n
            $objCollection.SaveChanges()
            Write-Verbose "Created $n COM+ Application."

            if ($PassThru)
            {
                Get-COMPlusApplication $n
            }
        }  
    }
}