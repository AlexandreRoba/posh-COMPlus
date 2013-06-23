function get-COMPlusApplication{
	param(
		[string]$Name
	)
	
	$comAdmin = New-Object -ComObject "COMAdmin.COMAdminCatalog"
	
	$objCollection = $comAdmin.GetCollection("Applications")
	$objCollection.Populate()
	foreach($application in $objCollection){
		$Properties = @{Key=$application.Key;Name=$application.Name}
		$New-Object -TypeName PsObject -Property $Properties
		Write-Output($application.Name)
	}
	
}