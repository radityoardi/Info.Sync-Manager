using namespace System.DirectoryServices
using namespace System.DirectoryServices.AccountManagement

[Cmdletbinding()]
param(
	[Parameter(Mandatory = $true, Position = 0)][Alias("s")][object]$Source
)
Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
#Load LDAP data into RawData property
try {
	$_temparray = [System.Collections.ArrayList]@()
	
	if ((Test-NetConnection $Source.ConnectionSettings.HostName -Port (Resolve-DefaultIfEmpty $Source.ConnectionSettings.Port 389)).TcpTestSucceeded) {
		Optimize-Object ([PrincipalContext]$LDAP = New-Object PrincipalContext([ContextType]::Domain, $Source.ConnectionSettings.HostName, $Source.ConnectionSettings.UserName, (ConvertTo-PlainPassword $Source.ConnectionSettings.Password))) {
			[UserExtPrincipal]$u = New-Object UserExtPrincipal($LDAP)
			[PrincipalSearcher]$pcs = New-Object PrincipalSearcher($u)
			Write-Host "$u"
			#Set the filter criteria
			if ($Source.GeneralSettings -and $Source.GeneralSettings.FilterCriteria -and $Source.GeneralSettings.FilterCriteria -is [array]) {
				Push-Verbose "Use-LDAPImport: Set the filter criteria."				
				$Source.GeneralSettings.FilterCriteria | Where { (Test-IsEnabled $_.Enabled) } | ForEach {
					if ($_.Value -ne $null) {						
						$u.AttributeSet("$($_.Name)", "$($_.Value)")
						Push-Verbose "Use-LDAPImport: Filter Criteria where '$($_.Name)' is '$($_.Value)'."
					} 
					elseif ($_.IsPresent -ne $null) {
						if ($_.IsPresent) {							
							$u.AdvancedSearchFilter.IsPresent("$($_.Name)")
							Push-Verbose "Use-LDAPImport: Filter Criteria where '$($_.Name)' is present."
						} else {
							$u.AdvancedSearchFilter.IsNotPresent("$($_.Name)")
							Push-Verbose "Use-LDAPImport: Filter Criteria where '$($_.Name)' is not present."
						}
					}		
				}
			}
			#Push-Verbose "Exit 1"
			$results = $pcs.FindAll()
	
			Push-Verbose "Use-LDAPImport: Found $(($results | Measure-Object).Count) users."
	
			if ($results -and ($results | Measure-Object).Count -gt 0) {
				$results | ForEach {
					[UserExtPrincipal]$result = $_ -as [UserExtPrincipal]
					[PSCustomObject]$data = [PSCustomObject]@{}
					if ($Source.Mappings -and $Source.Mappings -is [array]) {
						$Source.Mappings | Where { (Test-IsEnabled $_.Enabled) } | ForEach {
							$data | Add-Member -Type NoteProperty -Name $_.SourceProperty -Value $result.AttributeGet("$($_.SourceProperty)")
						}
					}
					$_temparray.Add($data)
				}
			}
		}
	}
	
	$Source | Add-Member RawData -MemberType NoteProperty -Value $_temparray
	Remove-Variable _temparray
}
catch {
	Push-Error $_ "Use-LDAPImport: Error when importing from LDAP '$($Source.ConnectionSettings.HostName)'."
}


