#Requires -version 2.0

Set-StrictMode -Version Latest

. $PSScriptRoot\Get-ComPlusApplication.ps1
. $PSScriptRoot\New-ComPlusApplication.ps1

Export-ModuleMember -Function Get-COMPlusApplication,New-ComPlusApplication