[Cmdletbinding()]
param(
	[Parameter(Mandatory = $false)][object]$ScrManagerEmployeeID,
	[Parameter(Mandatory = $false)][object]$ScrDistinguishedNameColumn
)
Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

$managerDN = ($global:c.SyncConfiguration.Staging.RawData | Where { $_.EmployeeID -eq $ScrManagerEmployeeID } | Select -First 1)
if ($managerDN) {
	Push-Verbose "Use-SearchManagerDN: found 1 manager '$($managerDN."$($ScrDistinguishedNameColumn)")'."
	return $managerDN."$($ScrDistinguishedNameColumn)"	
	
} else {
	return $null
}

