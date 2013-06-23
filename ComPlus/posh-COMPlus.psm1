function get-COMPlusApplication{
	param(
		[string]$Name
	)
	
	New-Object -ComObject "COMAdmin.COMAdminCatalog"
	
	
}

Export-ModuleMember -Cmdlet @('Get-COMPlusApplication')