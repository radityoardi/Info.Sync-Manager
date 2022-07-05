[Cmdletbinding()]
param(
	[Parameter(Mandatory = $true, Position = 0)][Alias("s")][object]$Source
)
Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

#Load CSV data into RawData property
try {
	$_temparray = [System.Collections.ArrayList]@()

	if (Test-Path ($Source.FilePath | Format-String -Replacement $global:c.RuntimeConfiguration.PredefinedFormattings)) {
		$sourceFilePath = ($Source.FilePath | Format-String -Replacement $global:c.RuntimeConfiguration.PredefinedFormattings)

		$importedCsv = $sourceFilePath | Import-Csv -Delimiter $(if($Source.Delimiter) { $Source.Delimiter } else { "," }) -Encoding UTF8
		
		$importedCsv | Where-Object { ($null -eq $Source.FilterExpression) -or ($null -ne $Source.FilterExpression -and (Invoke-Expression $Source.FilterExpression)) } | ForEach {
			$_temparray.Add($_)
		}
	
		Push-Verbose "Use-CSVImport: Successfully imported $($_temparray.Count) records."
	
	} else {
		Push-Warning "Use-CSVImport: File '$($Source.FilePath)' specified in '$($Source.Description)' could not be found."
	}

	$Source | Add-Member RawData -MemberType NoteProperty -Value $_temparray
	Remove-Variable _temparray
} catch {
	Push-Error $_ "Use-CSVImport: Error during record retrieval from '$($Source.FilePath)'."
}

