[Cmdletbinding()]
param(
	[Parameter(Mandatory = $false)][string]$ScrPreferredName,
	[Parameter(Mandatory = $false)][string]$ScrFirstName,
	[Parameter(Mandatory = $false)][string]$ScrLastName,
	[Parameter(Mandatory = $false)][string]$ScrLeftValue,
	[Parameter(Mandatory = $false)][string[]]$ScrRightValues
)

$match = (($ScrRightValues | Where-Object { $_ -eq $ScrLeftValue }).Length -gt 0)
Push-Verbose "SWITCHNAMES: It's a match: $($match)"
if ($match) {
	return $ScrFirstName
} else {
	return $ScrPreferredName
}
