[Cmdletbinding()]
param(
	[Parameter(Mandatory = $false)][object]$ScrJoinDate,
	[Parameter(Mandatory = $false)][object]$ScrLastWorkingDate,
	[Parameter(Mandatory = $false)][object]$ScrEmployeeStatus
)

[bool]$returnvalueAccountExist = $false
[bool]$returnvalueAccountEnabled = $false
[bool]$returnvalueUpdateproperty = $false

	if ($ScrEmployeeStatus -eq "Active" -and $ScrJoinDate -and (-not $ScrLastWorkingDate -or ($ScrLastWorkingDate -and $ScrLastWorkingDate -ge [DateTime]::Today))) {
		#Status is Active & JoinedDate is exist (don't care whether its past or future) & LastWorkingDay is either empty or future date or today's date
		$returnvalueAccountExist = $true
		$returnvalueAccountEnabled =$true
		$returnvalueUpdateproperty =$true
	}
	elseif ($ScrEmployeeStatus -eq "Terminated" -and $ScrJoinDate -and $ScrJoinDate -ge [DateTime]::Today) {
		$returnvalueAccountExist = $true
		$returnvalueAccountEnabled =$true
		$returnvalueUpdateproperty =$true
	}
	elseif ($ScrEmployeeStatus -eq "Terminated" -and $ScrLastWorkingDate -and $ScrLastWorkingDate -lt [datetime]::Today) {
		$returnvalueAccountExist = $false
		$returnvalueAccountEnabled =$false
		$returnvalueUpdateproperty =$false
	}
return [array] @($returnvalueAccountExist,$returnvalueAccountEnabled,$returnvalueUpdateproperty)