using namespace Management.Automation.Host
using namespace Collections.ObjectModel
using namespace System.Text

$CurrentScriptFile = Split-Path $MyInvocation.MyCommand.Definition -Leaf

if (-not (Get-Module PreferenceVariables)) {
	if (-not (Get-InstalledModule PreferenceVariables)) {
		Install-Module PreferenceVariables -Confirm:$false -Force
	}
	Import-Module PreferenceVariables
}


Function Optimize-Object {
	[Cmdletbinding()]
	Param (
		[Parameter(ValueFromPipeline)][System.IDisposable] $inputObject,
		[ScriptBlock] $scriptBlock
	)
	try {
		Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
		& $scriptBlock
	}
	catch {
		Throw
	}
	finally {
		if ($inputObject -ne $null) {
			if ($inputObject.psbase -eq $null) {				
				$inputObject.Dispose()
			} else {
				$inputObject.psbase.Dispose()
			}
		}
	}
}
Export-ModuleMember -Function Optimize-Object

Function Clear-Object {
	[Cmdletbinding()]
	Param (
		[Parameter(ValueFromPipeline)][object] $inputObjects
	)
	begin {
		Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	}
	process {
		$sb = {
			Param (
				[Parameter(Mandatory = $true, Position = 0)][System.IDisposable] $inputObject
			)
			if ($inputObject -is [System.IDisposable]) {
				if ($inputObject -ne $null) {
					if ($inputObject.psbase -eq $null) {
						$inputObject.Dispose()
					} else {
						$inputObject.psbase.Dispose()
					}
				}
			}
		}
		if ($inputObjects -is [array]) {
			$inputObjects | ForEach {
				if ($_ -is [System.IDisposable]) {
					$sb.Invoke($_)
				}
			}
		} elseif ($inputObjects -is [System.IDisposable]) {
			$sb.Invoke($inputObjects)
		}
	}
}
Export-ModuleMember -Function Clear-Object

Function Write-ConfigError
{
	[Cmdletbinding()]
	Param (
		[string]$Message,
		[Parameter(Mandatory = $false)][string]$CategoryReason = "Invalid configuration",
		[Parameter(Mandatory = $false)][string]$CategoryTargetName,
		[Parameter(Mandatory = $false)][string]$RecommendedAction
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	Write-Error "Error in configuration. $($Message)" `
		-Category InvalidOperation `
		-CategoryReason $CategoryReason `
		-CategoryTargetName $CategoryTargetName `
		-RecommendedAction $RecommendedAction
}
Export-ModuleMember -Function Write-ConfigError

Function Write-Success
{
	[Cmdletbinding()]
	Param (
		[Object]$Object,
		[Object]$Separator
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	Write-Host -Object $Object -Separator $Separator -ForegroundColor Green
}
Export-ModuleMember -Function Write-Success

Function Write-Header
{
	[Cmdletbinding()]
	Param (
		[Parameter(Mandatory = $true)][string]$ApplicationName,
		[Parameter(Mandatory = $false)][string]$Version,
		[Parameter(Mandatory = $true)][string]$Description
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	Write-Host "=============================================START============================================="
	Write-Host "Application: $($ApplicationName)"
	if ($Version) {
		Write-Host "Version: $($Version)"
	}
	Write-Host "$($Description)"
	Write-Host "==============================================================================================="
}
Export-ModuleMember -Function Write-Header

Function Write-Footer
{
	[Cmdletbinding()]
	Param ()
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	Write-Host "==============================================END=============================================="
}
Export-ModuleMember -Function Write-Footer

Function Start-YesNoChoice
{
	[Cmdletbinding()]
	Param (
		[Parameter(Mandatory = $true, Position = 0)][string]$Message,
		[Parameter(Mandatory = $true, Position = 1)][string]$Question,
		[Parameter(Mandatory = $false)][string]$YesText = "&Yes",
		[Parameter(Mandatory = $false)][string]$NoText = "&No"
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
	$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList $YesText))
	$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList $NoText))

	$decision = $host.UI.PromptForChoice($Message, $Question, $choices, 1)
	return $decision
}
Export-ModuleMember -Function Start-YesNoChoice

Function ConvertTo-PlainPassword
{
	[Cmdletbinding()]
	Param (
		[Parameter(Mandatory = $true, Position = 0)][string]$PasswordString
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	if ($null -eq $PasswordString)
	{
		Throw "PasswordString parameter is empty."
	}
	else
	{
		$SecurePwd = ConvertTo-SecureString $PasswordString
		if (-not $SecurePwd) { Throw "Unable to convert PasswordString, possibly wrong key decryptor." }
		$tempcrd = New-Object System.Management.Automation.PSCredential ("dummyusername", $SecurePwd)
		return $tempcrd.GetNetworkCredential().Password
	}
}
Export-ModuleMember -Function ConvertTo-PlainPassword

Function Start-MergeCustomOperations
{
	[Cmdletbinding(DefaultParameterSetName = "Table")]
	Param (
		[Parameter(Mandatory = $true, Position = 1, ParameterSetName = "Table")]
		[Parameter(ParameterSetName = "Row")]
		[Object]$Operation,
		[Parameter(Mandatory = $false, Position = 1, ParameterSetName = "Row")]
		[ref]$CurrentRow,
		[Parameter(Mandatory = $false, Position = 2, ParameterSetName = "Table")]
		[ref]$CurrentTable,
		[Parameter(Mandatory = $false, Position = 3, ParameterSetName = "Table")]
		[Parameter(ParameterSetName = "Row")]
		[Object]$CustomParameters
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	#note that $CurrentRow refers to $csvRow for BeforeCurrentRow, and $stagingRow for AfterCurrentRow
	
	$rtn = @{}
	if ($global:c -and $Operation -and (($PSCmdlet.ParameterSetName -eq "Row" -and $CurrentRow.Value) -or ($PSCmdlet.ParameterSetName -eq "Table" -and $CurrentTable.Value)))
	{
		$rtn.DoNothing = $true
		$rtn.MethodIdentifier = $Operation.MethodIdentifier

		switch ($Operation.MethodIdentifier) {
			"StripOffDashAndBefore"
			{
				if ($Operation.Execution -eq "BeforeCurrentRow" -or $Operation.Execution -eq "AfterCurrentRow")
				{
					$ptrn = "^.*-" #selects dash and any characters before
					$CurrentValue = $CurrentRow.Value.CSV."$($Operation.Parameters.SourceProperty)"
					if ([System.Text.RegularExpressions.Regex]::IsMatch($CurrentValue, $ptrn))
					{
						$NewValue = [System.Text.RegularExpressions.Regex]::Replace($CurrentValue, $ptrn, [System.String]::Empty)
						if ($Operation.Execution -eq "BeforeCurrentRow") { $CurrentRow.Value.CSV."$($Operation.Parameters.DestinationProperty)" = $NewValue }
						elseif ($Operation.Execution -eq "AfterCurrentRow") { $CurrentRow.Value.Staging."$($Operation.Parameters.DestinationProperty)" = $NewValue }
						Write-Verbose "Props '$($Operation.Parameters.SourceProperty)' value $($CurrentValue)' is now changed to '$($NewValue)'"
					}
					else
					{
						Write-Verbose "Props '$($Operation.Parameters.SourceProperty)' value '$($CurrentValue)' is currently not match with pattern '$($ptrn)'."
					}
				}
				else
				{
					Write-Warning "Method '$($Operation.MethodIdentifier)' is currently not implemented for '$($Operation.Execution)'."
				}
				break
			}
			"StripOffDashAndAfter"
			{
				if ($Operation.Execution -eq "BeforeCurrentRow" -or $Operation.Execution -eq "AfterCurrentRow")
				{
					$ptrn = "-.*$" #selects dash and any characters after
					$CurrentValue = $CurrentRow.Value.CSV."$($Operation.Parameters.SourceProperty)"
					if ([System.Text.RegularExpressions.Regex]::IsMatch($CurrentValue, $ptrn))
					{
						$NewValue = [System.Text.RegularExpressions.Regex]::Replace($CurrentValue, $ptrn, [System.String]::Empty)
						if ($Operation.Execution -eq "BeforeCurrentRow") { $CurrentRow.Value.CSV."$($Operation.Parameters.DestinationProperty)" = $NewValue }
						elseif ($Operation.Execution -eq "AfterCurrentRow") { $CurrentRow.Value.Staging."$($Operation.Parameters.DestinationProperty)" = $NewValue }
						Write-Verbose "Props '$($Operation.Parameters.SourceProperty)' value '$($CurrentValue)' is now changed to '$($NewValue)'"
					}
					else
					{
						Write-Verbose "Props '$($Operation.Parameters.SourceProperty)' value '$($CurrentValue)' is currently not match with pattern '$($ptrn)'."
					}
				}
				else
				{
					Write-Warning "Method '$($Operation.MethodIdentifier)' is currently not implemented for '$($Operation.Execution)'."
				}
				break
			}
			"JoinString"
			{
				if ($Operation.Execution -eq "AfterCurrentRow")
				{
					$propsValues = @{}
					ForEach ($props in $Operation.Parameters.StagingPropertiesToJoin)
					{
						$propsValues += @{
							"$($props)" = $CurrentRow.Value.Staging."$($props)"
						}
					}
					$CurrentRow.Value.Staging."$($Operation.Parameters.DestinationProperty)" = ($Operation.Parameters.JoinTemplate | Format-String -Replacement $propsValues)
					Write-Verbose "Property '$($Operation.Parameters.DestinationProperty)' is now modified to '$($CurrentRow.Value.Staging."$($Operation.Parameters.DestinationProperty)")'"
				}
				else
				{
					Write-Warning "Method '$($Operation.MethodIdentifier)' is currently not implemented for '$($Operation.Execution)'."
				}
				break
			}
			"LookupCsv"
			{
				if ($Operation.Execution -eq "AfterCurrentRow")
				{
					$csvFilePath = ($Operation.Parameters.CsvFilePath | Format-String -Replacement $CustomParameters.PredefinedFormattings)
					if (Test-Path $csvFilePath)
					{
						$csvObj = Import-Csv -Path $csvFilePath
						Write-Verbose "Staging value of property '$($Operation.Parameters.StagingLookupProperty)': '$($CurrentRow.Value.Staging."$($Operation.Parameters.StagingLookupProperty)")'"
						if ($CurrentRow.Value.Staging."$($Operation.Parameters.StagingLookupProperty)" -and $null -ne $CurrentRow.Value.Staging."$($Operation.Parameters.StagingLookupProperty)" -and $CurrentRow.Value.Staging."$($Operation.Parameters.StagingLookupProperty)".GetType().Name -eq "String")
						{
							$newValue = $csvObj | Where-Object { $_."$($Operation.Parameters.CsvLookupColumn)" -eq $CurrentRow.Value.Staging."$($Operation.Parameters.StagingLookupProperty)".Trim() } | Select-Object -First 1 -ExpandProperty "$($Operation.Parameters.CsvValueColumn)"
							Write-Verbose "New value for '$($Operation.Parameters.DestinationProperty)': '$($newValue)'"
							$CurrentRow.Value.Staging."$($Operation.Parameters.DestinationProperty)" = $newValue
							Clear-Variable csvObj
						}
						else
						{
							Clear-Variable csvObj
							Throw "Staging value of property '$($Operation.Parameters.StagingLookupProperty)' is null or its type not String."
						}
					}
					else
					{
						Throw "Error executing '$($Operation.MethodIdentifier)', csv file '$($Operation.Parameters.CsvFilePath)' does not exist."
					}
				}
				else
				{
					Write-Warning "Method '$($Operation.MethodIdentifier)' is currently not implemented for '$($Operation.Execution)'."
				}
				break
			}
			Default
			{
				Write-Warning "Method '$($Operation.MethodIdentifier)' is currently not implemented."
			}
		}
	}
	else
	{
		Write-Verbose "It's either 'Configuration' is empty, or 'Operation', or either 'CurrentRow' and 'CurrentTable' that is empty. Skips custom operation."
	}
	return $rtn
}
Export-ModuleMember -Function Start-MergeCustomOperations

Function Start-DeployCustomOperations
{
	[Cmdletbinding(DefaultParameterSetName = "Table")]
	Param (
		[Parameter(Mandatory = $true, Position = 1, ParameterSetName = "Table")]
		[Parameter(ParameterSetName = "Row")]
		[Object]$Operation,
		[Parameter(Mandatory = $false, Position = 2, ParameterSetName = "Row")]
		[ref]$CurrentRow,
		[Parameter(Mandatory = $false, Position = 3, ParameterSetName = "Table")]
		[ref]$CurrentTable,
		[Parameter(Mandatory = $false, Position = 4, ParameterSetName = "Table")]
		[Parameter(ParameterSetName = "Row")]
		[Object]$CustomParameters
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	$rtn = @{}
	if ($global:c -and $Operation -and (($PSCmdlet.ParameterSetName -eq "Row" -and $CurrentRow.Value) -or ($PSCmdlet.ParameterSetName -eq "Table" -and $CurrentTable.Value)))
	{
		$rtn.DoNothing = $true
		$rtn.MethodIdentifier = $Operation.MethodIdentifier

		if ($Operation.MethodIdentifier -eq "AccountStatusLogic")
		{
			if ($Operation.Execution -eq "BeforeCurrentRow")
			{
				$JoinedDate = $CurrentRow.Value."$($Operation.Parameters.JoinedDateStagingProperty)"
				$LastWorkingDay = $CurrentRow.Value."$($Operation.Parameters.LastWorkingDateStagingProperty)"
				$IsPrimary = $CurrentRow.Value."$($Operation.Parameters.IsPrimaryStagingProperty)"
				$Status = $CurrentRow.Value."$($Operation.Parameters.StatusStagingProperty)"
				Write-Verbose "Joined: [$($JoinedDate)] - LastDay: [$($LastWorkingDay)] - Status:[$($Status)]"

				if ($Status -eq "Active" -and $JoinedDate -and (-not $LastWorkingDay -or ($LastWorkingDay -and $LastWorkingDay -ge [DateTime]::Now)))
				{
					#Status is Active & JoinedDate is exist (don't care whether its past or future) & LastWorkingDay is either empty or future date or today's date
					$rtn.DoNothing = $false
					$rtn.PropertiesMustUpdate = $true
					$rtn.AccountMustExist = $true
					$rtn.AccountEnabled = $true
					$rtn.MobileRecordMustExist = ($IsPrimary)
					$rtn.Reason = "Joined: [$($JoinedDate)] - LastDay: [$($LastWorkingDay)] - Status:[$($Status)]. Account belongs to an active employee."
				}
				elseif ($Status -eq "Terminated" -and $LastWorkingDay -and $LastWorkingDay -lt [datetime]::Now)
				{
					$rtn.DoNothing = $false
					$rtn.PropertiesMustUpdate = $false
					$rtn.AccountMustExist = $false
					$rtn.AccountEnabled = $false
					$rtn.MobileRecordMustExist = ($IsPrimary)
					$rtn.Reason = "Joined: [$($JoinedDate)] - LastDay: [$($LastWorkingDay)] - Status:[$($Status)]. Account belongs to an inactive employee."
				}
				else
				{
					Write-Verbose "Account status logic is exhausted, no processing further."
				}
			}
			else
			{
				Write-Warning "Method '$($Operation.MethodIdentifier)' is currently not implemented for '$($Operation.Execution)'."					
			}
		}
		elseif ($Operation.MethodIdentifier -eq "DeleteUnlinkedStaffRecords")
		{
			if ($Operation.Execution -eq "AfterWholeTable")
			{
				if ($Operation.Parameters.DeleteSqlCommandTemplate -and $Operation.Parameters.DeleteSqlCommandTemplate.MainCommand -and $Operation.Parameters.DeleteSqlCommandTemplate.ExclusionCommand)
				{
					$UserNames = $CurrentTable.Value | Select-Object -ExpandProperty "$($CustomParameters.KeyMapping.StagingProperty)"
					$ReplacementObject = @{
						DatabaseName     = "$($CustomParameters.Destination.ConnectionSettings.DatabaseName)"
						SchemaName       = "$($CustomParameters.Destination.ConnectionSettings.SchemaName)"
						TableName        = "$($CustomParameters.Destination.ConnectionSettings.TableName)"
						EmployeeIDColumn = "$($Operation.Parameters.EmployeeIDSqlColumnName)"
						FirstEmployeeID  = "$(Convert-SqlValueScript -Value ($Operation.Parameters.FirstEmployeeID -as "$($Operation.Parameters.EmployeeIDType)") -Type $Operation.Parameters.EmployeeIDType)"
						LastEmployeeID   = "$(Convert-SqlValueScript -Value ($Operation.Parameters.LastEmployeeID -as "$($Operation.Parameters.EmployeeIDType)") -Type $Operation.Parameters.EmployeeIDType)"
						KeyColumn        = "$($CustomParameters.KeyMapping.DestinationProperty)"
						AllKeyValues     = ($UserNames | ForEach-Object { return "$(Convert-SqlValueScript -Value ($_ -as "$($CustomParameters.KeyMapping.Type)") -Type $CustomParameters.KeyMapping.Type)" }) -join ","
					}
					$DeleteSqlCommand = $Operation.Parameters.DeleteSqlCommandTemplate.MainCommand | Format-String -Replacement $ReplacementObject
					if (($UserNames | Measure-Object).Count -gt 0)
					{
						$DeleteSqlCommand += ($Operation.Parameters.DeleteSqlCommandTemplate.ExclusionCommand | Format-String -Replacement $ReplacementObject)
					}
					$rtn.DoNothing = $false
					$rtn.SqlCommand = $DeleteSqlCommand
				}
			}
			else
			{
				Write-Warning "Method '$($Operation.MethodIdentifier)' is currently not implemented for '$($Operation.Execution)'."					
			}
		}
		else
		{
			Write-Warning "Method '$($Operation.MethodIdentifier)' is currently not implemented."
		}
	}
	else
	{
		Write-Warning "It's either 'Configuration' is empty, or 'Operation', or either 'CurrentRow' and 'CurrentTable' that is empty. Skips custom operation."
	}
	return $rtn
}
Export-ModuleMember -Function Start-DeployCustomOperations

Function Convert-SqlValueScript
{
	[Cmdletbinding()]
	Param (
		[Parameter(Mandatory = $true, Position = 0)][object]$Value,
		[Parameter(Mandatory = $true, Position = 1)][string]$Type
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	if ($null -eq $Value)
	{
		return "NULL"
	}
	else
	{
		if ("string|nvarchar|varchar|char|text|nvarchar|ntext".Contains($Type.ToLower()))
		{
			return "'$($Value.Replace("'", "''"))'"
		}
		elseif ($Type.ToLower() -eq "datetime")
		{
			return "'$($Value.ToString("yyyy-MM-dd HH:mm:ss"))'"
		}
		elseif ($Type.ToLower() -eq "bool" -or $Type.ToLower() -eq "boolean")
		{
			if ($Value.GetType().Name.ToLower() -eq "string")
			{
				if ($Value.ToString().ToUpper() -eq "T")
				{
					return "1"
				}
				elseif ($Value.ToString().ToUpper() -eq "F")
				{
					return "0"
				}
				else
				{
					if ([System.Boolean]::Parse($Value.ToString()))
					{
						return "1"
					}
					return "0"
				}

			}
			elseif ($Value.GetType().Name.ToLower() -eq "bool" -or $Value.GetType().Name.ToLower() -eq "boolean")
			{
				if ([bool]$Value)
				{
					return "1"
				}
				else
				{
					return "0"
				}
			}
			else
			{
				return "$($Value)"
			}
		}
		else
		{
			return "$($Value)"
		}
	}
}
Export-ModuleMember -Function Convert-SqlValueScript

Function Convert-SqlUpdateScript
{
	[Cmdletbinding()]
	Param (
		[Parameter(Mandatory = $true, Position = 0)][object]$Value,
		[Parameter(Mandatory = $true, Position = 1)][string]$Name,
		[Parameter(Mandatory = $true, Position = 2)][string]$Type
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	return "$($Name)=$(Convert-SqlValueScript -Value $Value -Type $Type)"
}
Export-ModuleMember -Function Convert-SqlUpdateScript

Function IsPropertyExist
{
	[Cmdletbinding()]
	Param (
		[Parameter(Mandatory = $true, Position = 0)][object]$Value,
		[Parameter(Mandatory = $true, Position = 1)][string]$Name
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	return (($Value | Get-Member -Name $Name | Measure-Object).Count -gt 0)
}
Export-ModuleMember -Function IsPropertyExist

Function Format-String
{
	[Cmdletbinding()]
	Param (
		[Parameter(Mandatory = $true, Position = 0, ValueFromPipeline)][string]$InputValue,
		[Parameter(Mandatory = $true, Position = 1)][Hashtable]$Replacement
	)
	begin {
		Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	}
	process {
		Write-Debug "Input String: $($InputValue)"
		Write-Debug "Replacement Type: $($Replacement.GetType().Name) - Count $($Replacement.Keys.Count)"
		if ($InputValue -and $InputValue.Length -gt 0 -and $Replacement.GetType().Name -eq "Hashtable" -and $Replacement.Keys.Count -gt 0)
		{
			$obj = @()
			$i = 0
			$OutputValue = $InputValue
			ForEach ($pair in $Replacement.GetEnumerator())
			{
				Write-Debug "Index: $($i)"
				Write-Debug "Key: $($pair.Key)"
				$pattern = "(?<=\{)$($pair.Key)(?=(,[^\{\}]*){0,1}(:[^\{\}]*){0,1}\})"
				#$pattern = "$($pair.Key)"
				Write-Debug "pattern: $($pattern)"
				if ($InputValue -imatch $pattern)
				{
					Write-Debug "Before replace: $($OutputValue)"
					$OutputValue = $OutputValue -replace $pattern, "$($i)"
					Write-Debug "After replace: $($OutputValue)"
					if ($pair.Value -and $pair.Value.GetType().Name -eq "ScriptBlock")
					{
						$obj += (Invoke-Command $pair.Value)
					}
					else
					{
						$obj += $pair.Value
					}
					$i += 1
				}
			}

			Write-Debug "$($OutputValue)"
			$OutputValue = $OutputValue -f $obj
			Write-Debug "$($OutputValue)"
			Return $OutputValue
		}
		elseif ($InputValue.Length -gt 0)
		{
			Throw "'Replacement' parameter must be in Hashtable, not $($Replacement.GetType().Name). Please read https://ss64.com/ps/syntax-hash-tables.html."
		}
	}
}
Export-ModuleMember -Function Format-String

Function Test-IsEnabled
{
	[Cmdletbinding()]
	Param (
		[Parameter(Mandatory = $false, Position = 0)][object]$InputValue
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	return ($null -eq $InputValue -or $InputValue -eq $true)
}
Export-ModuleMember -Function Test-IsEnabled

Function Set-AllProperties {
	[Cmdletbinding()]
	Param (
		[Parameter(Mandatory = $true, Position = 0)][ref]$Object,
		[Parameter(Mandatory = $true, Position = 1)][object]$CsvRow,
		[Parameter(Mandatory = $true, Position = 2)][object]$StagingProperty,
		[Parameter(Mandatory = $true, Position = 3)][object]$SourceMapping,
		[Parameter(Mandatory = $false, Position = 4)][switch]$Skip
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	if ($Object -and $Object.Value -and $StagingProperty.Name -and -not $Skip) {
		if ($SourceMapping.SourceProperty) {
			if ($null -ne $csvRow."$($SourceMapping.SourceProperty)" -and $csvRow."$($SourceMapping.SourceProperty)".Length -gt 0) {
				if ($StagingProperty.ParseExact -and $StagingProperty.Type -eq "DateTime" -and $csvRow."$($SourceMapping.SourceProperty)" -isnot [System.DateTime]) {
					Set-HashtableValue $Object "$($StagingProperty.Name)" (([System.DateTime]::ParseExact($csvRow."$($SourceMapping.SourceProperty)", $StagingProperty.ParseExact, [System.Globalization.CultureInfo]::CurrentCulture)) -as "$($StagingProperty.Type)")
				} else {
					Set-HashtableValue $Object "$($StagingProperty.Name)" ($csvRow."$($SourceMapping.SourceProperty)" -as "$($StagingProperty.Type)")
				}
			} else {
				Set-HashtableValue $Object "$($StagingProperty.Name)" ($csvRow."$($SourceMapping.SourceProperty)" -as "$($StagingProperty.Type)") $null
			}
		} else {
			if ($null -ne $SourceMapping.DefaultValue -and (-not ($SourceMapping.DefaultValue.GetType().Name -eq "string" -and $SourceMapping.DefaultValue -eq "{NULL}"))) {
				Set-HashtableValue $Object "$($StagingProperty.Name)" ($SourceMapping.DefaultValue -as "$($StagingProperty.Type)")
			} elseif ($null -ne $SourceMapping.DefaultValue -and $SourceMapping.DefaultValue.Length -gt 0 -and $StagingProperty.ParseExact -and $StagingProperty.Type -eq "DateTime") {
				Set-HashtableValue $Object "$($StagingProperty.Name)" (([System.DateTime]::ParseExact("$($SourceMapping.DefaultValue)", $StagingProperty.ParseExact, [System.Globalization.CultureInfo]::CurrentCulture)) -as "$($StagingProperty.Type)")
			} else {
				Set-HashtableValue $Object "$($StagingProperty.Name)" $null
			}
		}
	}
}
Export-ModuleMember -Function Set-AllProperties

Function Set-HashtableValue {
	[Cmdletbinding()]
	Param (
		[Parameter(Mandatory = $true, Position = 0)][ref]$InputHashtable,
		[Parameter(Mandatory = $true, Position = 1)][string]$Key,
		[Parameter(Mandatory = $false, Position = 2)][object]$Value,
		[Parameter(Mandatory = $false, Position = 3)][switch]$SkipWhenExist
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	if ($InputHashtable -and $InputHashtable.Value) {
		if ($InputHashtable.Value.ContainsKey($Key)) {
			if (-not $SkipWhenExist) {
				$InputHashtable.Value."$($Key)" = $Value
			}
		} else {
			$InputHashtable.Value += @{ "$($Key)" = $Value }
		}
	}
}
Export-ModuleMember -Function Set-HashtableValue

Function Set-PSCustomObjectValue {
	[Cmdletbinding()]
	Param (
		[Parameter(Mandatory = $true, Position = 0)][ref]$InputObject,
		[Parameter(Mandatory = $true, Position = 1)][string]$Key,
		[Parameter(Mandatory = $false, Position = 2)][object]$Value,
		[Parameter(Mandatory = $false, Position = 3)][switch]$SkipWhenExist
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	if ($InputObject -and $InputObject.Value) {
		if (($InputObject.Value | Get-Member -Name "$($Key)")) {
			if (-not $SkipWhenExist) {
				$InputObject.Value."$($Key)" = $Value
			}
		} else {
			$InputObject.Value | Add-Member -MemberType NoteProperty -Name "$($Key)" -Value $Value
		}
	}
}
Export-ModuleMember -Function Set-PSCustomObjectValue

Function Invoke-CustomOperations
{
	[Cmdletbinding(DefaultParameterSetName = "ExecutionFilter")]
	Param (
		[Parameter(Mandatory = $false, Position = 0, ParameterSetName = "ExecutionFilter")]
		[Parameter(ParameterSetName = "BlockFilter")]
		[ref]$Object,
		[Parameter(Mandatory = $true, Position = 1, ParameterSetName = "ExecutionFilter")]
		[Parameter(ParameterSetName = "BlockFilter")]
		[ref]$ConfigContext,
		[Parameter(Mandatory = $true, Position = 2, ParameterSetName = "ExecutionFilter")]
		[string]$ExecutionMode,
		[Parameter(Mandatory = $true, Position = 2, ParameterSetName = "BlockFilter")]
		[scriptblock]$BlockFilter,
		[Parameter(Mandatory = $false, Position = 3, ParameterSetName = "ExecutionFilter")]
		[Parameter(ParameterSetName = "BlockFilter")]
		[object]$CustomParameters,
		[Parameter(Mandatory = $true, Position = 4, ParameterSetName = "ExecutionFilter")]
		[Parameter(ParameterSetName = "BlockFilter")]
		[string]$ExecutionStage,
		[Parameter(Mandatory = $false, Position = 5, ParameterSetName = "ExecutionFilter")]
		[Parameter(ParameterSetName = "BlockFilter")]
		[ref]$StagingRow
	)
	<#
	Usually being called in 4 modes: RawSource, Source, Staging, and Destination

	#>
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	if ($ConfigContext.Value.CustomOperations) {
		return ($ConfigContext.Value.CustomOperations | `
		Where $( if ($PSCmdlet.ParameterSetName -eq "ExecutionFilter"){ { $_.Execution -eq $ExecutionMode -and (Test-IsEnabled $_.Enabled) } } else { $BlockFilter }) | Sort-Object -Property @{ Expression = { if ($_.ExecutionOrder) { $_.ExecutionOrder } else { 100000 }} } | `
		ForEach {
			$CustomOperation = $_

			Push-Verbose "Executing Custom Operation '$($CustomOperation.MethodIdentifier)' of type '$($ExecutionMode)'"

			switch ($CustomOperation.MethodIdentifier) {
				"StringBeforeDash" {
					if ($Object.Value.ContainsKey($CustomOperation.Parameters.SourceProperty) -and $Object.Value."$($CustomOperation.Parameters.SourceProperty)") {
						$FirstDashIndex = $Object.Value."$($CustomOperation.Parameters.SourceProperty)".IndexOf("-")
						if ($FirstDashIndex -ge 0) {
							Set-HashtableValue $Object "$($CustomOperation.Parameters.DestinationProperty)" ($Object.Value."$($CustomOperation.Parameters.SourceProperty)".Substring(0, $FirstDashIndex))
						}
					}
				}
				"StringAfterDash" {
					if ($Object.Value.ContainsKey($CustomOperation.Parameters.SourceProperty) -and $Object.Value."$($CustomOperation.Parameters.SourceProperty)") {
						$FirstDashIndex = $Object.Value."$($CustomOperation.Parameters.SourceProperty)".IndexOf("-")
						if ($FirstDashIndex -ge 0) {
							Set-HashtableValue $Object "$($CustomOperation.Parameters.DestinationProperty)" ($Object.Value."$($CustomOperation.Parameters.SourceProperty)".Substring($FirstDashIndex + 1))
						}
					}
				}
				"FormatString" {
					if ($CustomOperation.Parameters.SourceProperties) {
						$propsValues = @{}
						$CustomOperation.Parameters.SourceProperties | ForEach {
							$propsValues["$($_)"] = $Object.Value."$($_)"
						}
						Set-HashtableValue $Object "$($CustomOperation.Parameters.DestinationProperty)" ($CustomOperation.Parameters.JoinTemplate | Format-String -Replacement $propsValues)
					}
				}
				"CsvLookup" {
					$csvFilePath = ($CustomOperation.Parameters.CsvFilePath | Format-String -Replacement $CustomParameters.PredefinedFormattings)
					if (Test-Path $csvFilePath)
					{
						if (-not $CustomOperation.RawData) {
							$CustomOperation | Add-Member -Name RawData -MemberType NoteProperty -Value (Import-Csv -Path $csvFilePath)
						} elseif ($CustomOperation.RawData -and -not (Test-IsEnabled $CustomOperation.Parameters.LoadCsvOnce)) {
							$CustomOperation.RawData = (Import-Csv -Path $csvFilePath)
						}
						if ($Object.Value."$($CustomOperation.Parameters.StagingLookupProperty)" -and $null -ne $Object.Value."$($CustomOperation.Parameters.StagingLookupProperty)" -and $Object.Value."$($CustomOperation.Parameters.StagingLookupProperty)".GetType().Name -eq "String") {
							$Object.Value."$($CustomOperation.Parameters.DestinationProperty)" = `
								($CustomOperation.RawData | Where-Object { $_."$($CustomOperation.Parameters.CsvLookupColumn)" -eq $Object.Value."$($CustomOperation.Parameters.StagingLookupProperty)".Trim() } | Select-Object -First 1 -ExpandProperty "$($CustomOperation.Parameters.CsvValueColumn)")
						}					
					}
				}
				"DeleteEmployeeRecord" {
					if ($CustomOperation.Parameters.DeleteSqlCommandTemplate -and $CustomOperation.Parameters.DeleteSqlCommandTemplate.MainCommand -and $CustomOperation.Parameters.DeleteSqlCommandTemplate.ExclusionCommand)
					{
						$UserNames = $Object.Value | ForEach { $_."$($CustomParameters.KeyMapping.StagingProperty)" }
						$ReplacementObject = @{
							DatabaseName     = "$($CustomParameters.Destination.ConnectionSettings.DatabaseName)"
							SchemaName       = "$($CustomParameters.Destination.ConnectionSettings.SchemaName)"
							TableName        = "$($CustomParameters.Destination.ConnectionSettings.TableName)"
							EmployeeIDColumn = "$($CustomOperation.Parameters.EmployeeIDSqlColumnName)"
							FirstEmployeeID  = "$(Convert-SqlValueScript -Value ($CustomOperation.Parameters.FirstEmployeeID -as "$($CustomOperation.Parameters.EmployeeIDType)") -Type $CustomOperation.Parameters.EmployeeIDType)"
							LastEmployeeID   = "$(Convert-SqlValueScript -Value ($CustomOperation.Parameters.LastEmployeeID -as "$($CustomOperation.Parameters.EmployeeIDType)") -Type $CustomOperation.Parameters.EmployeeIDType)"
							KeyColumn        = "$($CustomParameters.KeyMapping.DestinationProperty)"
							AllKeyValues     = ($UserNames | ForEach-Object { return "$(Convert-SqlValueScript -Value ($_ -as "$($CustomParameters.KeyMapping.Type)") -Type $CustomParameters.KeyMapping.Type)" }) -join ","
						}
						$DeleteSqlCommand = $CustomOperation.Parameters.DeleteSqlCommandTemplate.MainCommand | Format-String -Replacement $ReplacementObject
						if (($UserNames | Measure-Object).Count -gt 0)
						{
							$DeleteSqlCommand += ($CustomOperation.Parameters.DeleteSqlCommandTemplate.ExclusionCommand | Format-String -Replacement $ReplacementObject)
						}
						$ReplacementObject.Clear()
						#return the command
						[PSCustomObject]@{
							CustomOperation = $CustomOperation
							SqlCommand = $DeleteSqlCommand
						}
					}
				}
				"Script" {
					if ($CustomOperation.Parameters.PowerShellCommand) {
						#if running with PowerShell Command directly.
						try {
							$returnval = Invoke-Expression ($CustomOperation.Parameters.PowerShellCommand | Format-String -Replacement $CustomParameters.PredefinedFormattings)
						}
						catch {
							Push-Error $_
						}
						if ($null -ne $returnval -and $returnval -is [array] -and $CustomOperation.Parameters.ScriptOutputParameters) {
							$lowestCounter = ($CustomOperation.Parameters.ScriptOutputParameters | Measure-Object).Count
							if (($returnval | Measure-Object).Count -lt $lowestCounter) {
								$lowestCounter = ($returnval | Measure-Object).Count
							}
							for ($i = 0; $i -lt $lowestCounter; $i++) {
								$Object.Value."$($CustomOperation.Parameters.ScriptOutputParameters[$i].ObjectProperty)" = $returnval[$i]								
							}
						} elseif ($null -ne $returnval -and $CustomOperation.Parameters.ScriptOutputParameter) {
							$Object.Value."$($CustomOperation.Parameters.ScriptOutputParameter.ObjectProperty)" = $returnval
						}
					} elseif ($CustomOperation.Parameters.ScriptPath) {

						#If running with script path.
						$params = @{}
						#Predefined Parameters
						if (Test-IsEnabled $CustomOperation.Parameters.PredefinedInputParameters) {
							if ($CustomParameters) {
								$params.Add("CustomParameters", $CustomParameters)
							}
							if ($Object -and $Object.Value) {
								$params.Add("Object", ([ref]$Object.Value))
							}
							if ($StagingRow -and $StagingRow.Value) {
								$params.Add("StagingRow", ([ref]$StagingRow.Value))
							}
						}
						#make input parameters
						$CustomOperation.Parameters.ScriptInputParameters | ForEach {
							if (Test-IsEnabled $_.DirectValue) {
								Push-Verbose "Setting script parameter '$($_.Name)' to '$($_.Value)'."
								$params.Add($_.Name, $_.Value)
							} elseif ($_.ObjectProperty) {
								if ($Object -and $Object.Value -and $Object.Value."$($_.ObjectProperty)") {
									Push-Verbose "Setting script parameter '$($_.Name)' value from object property '$($_.ObjectProperty)' = '$($Object.Value."$($_.ObjectProperty)")'."
									$params.Add($_.Name, $Object.Value."$($_.ObjectProperty)")	
								} else {
									Push-Warning "Failed setting script parameter '$($_.Name)' value from object property '$($_.ObjectProperty)', the object itself is empty."
								}
							} elseif ($_.StagingProperty) {
								if ($StagingRow -and $StagingRow.Value -and $StagingRow.Value."$($_.StagingProperty)") {
									Push-Verbose "Setting script parameter '$($_.Name)' value from staging property '$($_.StagingProperty)' = '$($StagingRow.Value."$($_.StagingProperty)")'."
									$params.Add($_.Name, $StagingRow.Value."$($_.StagingProperty)")	
								} else {
									Push-Warning "Failed setting script parameter '$($_.Name)' value from staging property '$($_.StagingProperty)', the staging row itself is empty."
								}
							}
						}
						Push-Verbose "Executing '$(($CustomOperation.Parameters.ScriptPath | Format-String -Replacement $CustomParameters.PredefinedFormattings))'."
						#Execute the script
						try {
							$returnval = (& "$(($CustomOperation.Parameters.ScriptPath | Format-String -Replacement $CustomParameters.PredefinedFormattings))" @params)
						}
						catch {
							Push-Error $_
						}
						Push-Verbose "Processing return value."
						#Processing the return value
						if ($null -ne $returnval -and $returnval -is [array] -and $CustomOperation.Parameters.ScriptOutputParameters) {
							$lowestCounter = ($CustomOperation.Parameters.ScriptOutputParameters | Measure-Object).Count
							if (($returnval | Measure-Object).Count -lt $lowestCounter) {
								$lowestCounter = ($returnval | Measure-Object).Count
							}
							for ($i = 0; $i -lt $lowestCounter; $i++) {
								if ($Object.Value.ContainsKey("$($CustomOperation.Parameters.ScriptOutputParameters[$i].ObjectProperty)")) {
									Push-Verbose "Set return value of '$($CustomOperation.Parameters.ScriptOutputParameters[$i].ObjectProperty)' with the value of '$(($returnval[$i] | Out-String).Trim())'"
									$Object.Value."$($CustomOperation.Parameters.ScriptOutputParameters[$i].ObjectProperty)" = $returnval[$i]
								} else {
									Push-Verbose "The object does not have member with name '$($CustomOperation.Parameters.ScriptOutputParameters[$i].ObjectProperty)'."
								}
							}
						} elseif ($null -ne $returnval -and $CustomOperation.Parameters.ScriptOutputParameter) {
							if ($Object.Value.ContainsKey("$($CustomOperation.Parameters.ScriptOutputParameter.ObjectProperty)")) {
								Push-Verbose "Set return value of '$($CustomOperation.Parameters.ScriptOutputParameter.ObjectProperty)' with the value of '$(($returnval | Out-String).Trim())'"
								$Object.Value."$($CustomOperation.Parameters.ScriptOutputParameter.ObjectProperty)" = $returnval
							} else {
								Push-Verbose "The object does not have member with name '$($CustomOperation.Parameters.ScriptOutputParameter.ObjectProperty)'."
							}
						} else {
							Push-Verbose "Nothing is set as there is no return value."
						}
					}
				}
				"TransferFileBinary" {
					$MaxSize = 100000
					$MinSize = 1
					if ($CustomOperation.Parameters.MaxSizeInBytes -and $CustomOperation.Parameters.MaxSizeInBytes -ne $MaxSize) {
						$MaxSize = $CustomOperation.Parameters.MaxSizeInBytes
					}
					if ($CustomOperation.Parameters.MinSizeInBytes -and $CustomOperation.Parameters.MinSizeInBytes -ne $MinSize) {
						$MinSize = $CustomOperation.Parameters.MinSizeInBytes
					}
					$updateDestination = {
						[Cmdletbinding()]
						Param (
							[Parameter(Mandatory = $true)][System.IO.FileInfo]$fileInfo,
							[Parameter(Mandatory = $true)][object]$operation,
							[Parameter(Mandatory = $true)][object]$destinationObject
						)
						Invoke-Expression $operation.Expression
					}
					$updatePhoto = {
						[Cmdletbinding()]
						Param (
							[Parameter(Mandatory = $true)][string]$AbsoluteFilePath
						)
						#if FilePathProperty is specified
						if (Test-IsValidPath $AbsoluteFilePath) {
							$fi = Get-Item $AbsoluteFilePath
							if ($fi.Length -le $MaxSize -and $fi.Length -ge $MinSize) {
								#less than 100KB
								Push-Verbose "Uploading file '$($fi.FullName)' to '$($CustomOperation.Parameters.DestinationProperty)'."
								$CustomOperation.Parameters.Operations | ForEach {
									if ($Object -and $Object.Value -and $Object.Value -is "$($_.ObjectIs)") {
										$updateDestination.Invoke($fi, $_, $Object.Value)
									}
								}
							} else {
								Push-Warning "File '$($fi.FullName)' size is $(($fi.Length / 1000).ToString("#,##0"))KB, must be within $(($MinSize / 1000).ToString("#,##0"))KB to $(($MaxSize / 1000).ToString("#,##0"))KB."
							}
						} else {
							Push-Warning "File '$($AbsoluteFilePath)' does not exist or invalid."
						}
					}
					if ($CustomOperation.Parameters.DestinationProperty) {
						if ($CustomOperation.Parameters.SourceFilePathProperty) {
							if ($StagingRow -and $StagingRow.Value -and $StagingRow.Value."$($CustomOperation.Parameters.SourceFilePathProperty)") {
								$updatePhoto.Invoke($StagingRow.Value."$($CustomOperation.Parameters.SourceFilePathProperty)")
							} else {
								Push-Warning "Either StagingRow parameter is not provided, or can't find staging row with column name '$($CustomOperation.Parameters.SourceFilePathProperty)'."
							}
						} elseif ($CustomOperation.Parameters.SourceFilePath) {
							$updatePhoto.Invoke(($CustomOperation.Parameters.SourceFilePath | Format-String -Replacement $CustomParameters.PredefinedFormattings))
						} else {
							Push-Warning "Either 'Parameters.SourceFilePath' or 'Parameters.SourceFilePathProperty' is not provided."
						}
					}
				}
				Default {}
			}
		})
	} else { return $null }
}
Export-ModuleMember -Function Invoke-CustomOperations

Function ConvertTo-PsCustomObjectFromHashtable { 
	param ( 
		[Parameter(  
			Position = 0,   
			Mandatory = $false,   
			ValueFromPipeline = $true,  
			ValueFromPipelineByPropertyName = $true  
		)] [object[]]$hashtable 
	)

	
	begin {
		Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
		$i = 0;
	}
	
	process {
		foreach ($myHashtable in $hashtable) { 
			if ($myHashtable -is [System.Collections.Hashtable]) { 
				$output = New-Object -TypeName PsObject; 
				Add-Member -InputObject $output -MemberType ScriptMethod -Name AddNote -Value {  
					Add-Member -InputObject $this -MemberType NoteProperty -Name $args[0] -Value $args[1]; 
				}; 
				$myHashtable.Keys | Sort-Object | % {
					$output.AddNote($_, $myHashtable.$_);  
				} 
				$output;
			} else { 
				Write-Warning "Index $i is not of type [hashtable]"; 
			} 
			$i += 1;  
		} 
	}
}
Export-ModuleMember -Function ConvertTo-PsCustomObjectFromHashtable

Function ConvertTo-HashtableFromPsCustomObject { 
	param ( 
		[Parameter(  
			Position = 0,   
			Mandatory = $true,   
			ValueFromPipeline = $true,  
			ValueFromPipelineByPropertyName = $true  
		)] [object[]]$psCustomObject 
	); 
	begin {
		Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	}
	process { 
		foreach ($myPsObject in $psObject) { 
			$output = @{}; 
			$myPsObject | Get-Member -MemberType *Property | % { 
				$output.($_.name) = $myPsObject.($_.name); 
			} 
			$output;
		} 
	} 
}
Export-ModuleMember -Function ConvertTo-HashtableFromPsCustomObject

Function Limit-EventLogMessage {
	[Cmdletbinding()]
	Param (
		[Parameter(Mandatory = $true, Position = 0)]
		[string]$Message
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	$addMsg = "WARNING:-->Truncated as it reached max characters.`r`n"
	if ($Message.Length -gt 30000) {
		return ("$($addMsg)$($Message)").Substring(0, 30000)
	} else {
		return $Message
	}
}
Export-ModuleMember -Function Limit-EventLogMessage

Function Push-Error {
	[Cmdletbinding()]
	Param (
		[Parameter(Mandatory = $true, Position = 0)]
		[System.Management.Automation.ErrorRecord]$ErrorRecord,
		[Parameter(Mandatory = $false, Position = 1)]
		[string]$Message,
		[switch]$NoEventLog,
		[switch]$NoWriteHost
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	@(
		(-join @(
			"$(if($Message){"$($Message)`r`n"})",
			"$(if($ErrorRecord.Exception){"Exception: $($ErrorRecord.Exception.Message)`r`n"})",
			"$(if($ErrorRecord.Exception -and $ErrorRecord.Exception.InnerException){"InnerException: $($ErrorRecord.Exception.InnerException.Message)`r`n"})",
			"$(if($ErrorRecord.Exception -and $ErrorRecord.Exception.InnerException -and $ErrorRecord.Exception.InnerException.InnerException){"InnerException: $($ErrorRecord.Exception.InnerException.InnerException.Message)`r`n"})",
			"$(if($ErrorRecord.InvocationInfo){"$($ErrorRecord.InvocationInfo.PositionMessage)`r`n`r`n"})",
			"$(if($ErrorRecord.ScriptStackTrace){"Script Stack Trace: $($ErrorRecord.ScriptStackTrace)"})"
		))
	) | ForEach {
		if (-not $NoWriteHost) {
			Write-Error $_
		}
		if (-not $NoEventLog) {
			Write-EventLog `
				-LogName $global:c.SystemConfiguration.WindowsLog.Name `
				-Source $global:c.SystemConfiguration.WindowsLog.Source `
				-EventId $global:c.RuntimeConfiguration.LogRefNumber -EntryType Error `
				-Message (Limit-EventLogMessage $_)
		}
	}
}
Export-ModuleMember -Function Push-Error

Function Push-Warning {
	[Cmdletbinding()]
	Param (
		[Parameter(Mandatory = $false, Position = 0)]
		[string]$Message,
		[switch]$NoEventLog,
		[switch]$NoWriteHost
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	@($Message) | ForEach {
		if (-not $NoWriteHost) {
			Write-Warning $Message
		}
		if (-not $NoEventLog) {
			Write-EventLog `
				-LogName $global:c.SystemConfiguration.WindowsLog.Name `
				-Source $global:c.SystemConfiguration.WindowsLog.Source `
				-EventId $global:c.RuntimeConfiguration.LogRefNumber -EntryType Warning `
				-Message (Limit-EventLogMessage $_)
		}
	}
}
Export-ModuleMember -Function Push-Warning

Function Push-Info {
	[Cmdletbinding()]
	Param (
		[Parameter(Mandatory = $true, Position = 0)]
		[string]$Message,
		[Parameter(Mandatory = $false)][Alias("Fore")]
		[ConsoleColor]$ForegroundColor,
		[switch]$NoEventLog,
		[switch]$NoWriteHost
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	@($Message) | ForEach {
		if (-not $NoWriteHost) {
			if ($ForegroundColor) { Write-Host $Message -Fore $ForegroundColor }
			else { Write-Host $Message }
		}
		if (-not $NoEventLog) {
			Write-EventLog `
				-LogName $global:c.SystemConfiguration.WindowsLog.Name `
				-Source $global:c.SystemConfiguration.WindowsLog.Source `
				-EventId $global:c.RuntimeConfiguration.LogRefNumber -EntryType Information `
				-Message (Limit-EventLogMessage $_)
		}
	}
}
Export-ModuleMember -Function Push-Info

Function Push-Verbose {
	[Cmdletbinding()]
	Param (
		[Parameter(Mandatory = $true, Position = 0)]
		[string]$Message,
		[switch]$NoEventLog,
		[switch]$NoWriteHost
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	if ($VerbosePreference -ne "SilentlyContinue") {
		#not being used as of now, by default will always log
	}
	@($Message) | ForEach {
		if (-not $NoWriteHost) {
			Write-Verbose $Message
		}
		if (-not $NoEventLog) {
			Write-EventLog `
				-LogName $global:c.SystemConfiguration.WindowsLog.Name `
				-Source $global:c.SystemConfiguration.WindowsLog.Source `
				-EventId $global:c.RuntimeConfiguration.LogRefNumber -EntryType Information `
				-Message (Limit-EventLogMessage $_)
		}
	}
}
Export-ModuleMember -Function Push-Verbose

Function GetMessageFromMeasurement {
	[Cmdletbinding()]
	Param (
		[Parameter(Mandatory = $true, Position = 0)]
		[System.TimeSpan]$Measurement
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	return (-join @(
		"$(if($Measurement.Days -gt 0){"$($Measurement.Days) d"})",
		"$(if($Measurement.Days -gt 0 -and $Measurement.Hours -gt 0){", "})",
		"$(if($Measurement.Hours -gt 0){"$($Measurement.Hours) hr"})",
		"$(if($Measurement.Hours -gt 0 -and $Measurement.Minutes -gt 0){", "})",
		"$(if($Measurement.Minutes -gt 0){"$($Measurement.Minutes) min"})",
		"$(if($Measurement.Minutes -gt 0 -and $Measurement.Seconds -gt 0){", "})",
		"$(if($Measurement.Seconds -gt 0){"$($Measurement.Seconds) sec"})",
		"$(if($Measurement.Seconds -gt 0 -and $Measurement.Milliseconds -gt 0){", "})",
		"$(if($Measurement.Milliseconds -gt 0){"$($Measurement.Milliseconds) ms"})"
	))
}
Export-ModuleMember -Function GetMessageFromMeasurement

Function Test-IsValidPath {
	[Cmdletbinding()]
	Param (
		[Parameter(Mandatory = $true, Position = 0)]
		[string]$FilePath
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	return ($FilePath -and $null -ne $FilePath -and $FilePath.Length -gt 0 -and (Test-Path $FilePath))
}
Export-ModuleMember -Function Test-IsValidPath

Function Test-IsAdministrator {
	[Cmdletbinding()]
	Param ()
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	return ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}
Export-ModuleMember -Function Test-IsAdministrator

Function Test-IsSpecified {
	[Cmdletbinding()]
	Param (
		[Parameter(Mandatory = $true, Position = 0)]
		[object]$Value
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	return ($null -ne $Value -or ($Value -is [string] -and -not (string.IsNullOrEmpty($Value))))
}
Export-ModuleMember -Function Test-IsSpecified

Function Import-Assemblies {
	Param (
		[Parameter(Mandatory = $true)][PSCustomObject]$SystemConfig
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	if ($SystemConfig.Modules -and $SystemConfig.Modules -is [array]) {
		Push-Info "Importing assemblies from configuration file."
		$SystemConfig.Assemblies | ForEach {
			{
				param (
					[object]$assemblyObj
				)
				if ($assemblyObj -is [string]) {
					try {
					Push-Verbose "Loading assembly from GAC '$($assemblyObj)'."
					Add-Type -AssemblyName $assemblyObj
					Push-Verbose "Assembly '$($assemblyObj)' was loaded."
					}
					catch {
						Push-Error $_ "Error when loading '$($assemblyObj)' assembly"
					}
	
				} else {
	
					if ($assemblyObj -and $assemblyObj.LiteralPath) {
						try {
							Push-Verbose "Loading assembly from a literal path '$($assemblyObj.LiteralPath)'."
							Add-Type -LiteralPath $assemblyObj.LiteralPath 
							Push-Verbose "Assembly '$($assemblyObj.LiteralPath)' was loaded."
						}
						catch {
							Push-Error $_ "Error when loading '$($assemblyObj.LiteralPath)' assembly"
						}
						 } elseif ($assemblyObj -and $assemblyObj.SourceCodePath) {
	
						try {
							Push-Verbose "Loading assembly from a source code path '$($assemblyObj.SourceCodePath)'."
							Add-Type -ReferencedAssemblies $assemblyObj.ReferencedAssemblies -TypeDefinition (Get-Content $assemblyObj.SourceCodePath -Raw) -Language CSharpVersion3
							Push-Verbose "Assembly '$($assemblyObj.SourceCodePath)' was loaded."
						}
						catch {
							Push-Error $_ "Error when loading '$($assemblyObj.SourceCodePath)' assembly"
						}
					} else {
						Push-Verbose "Not supported loading '$(($assemblyObj | Out-String).Trim())'."
					}
				}
			}.Invoke($_)
		}
	}
}
Export-ModuleMember -Function Import-Assemblies

Function Import-PSModules {
	Param (
		[Parameter(Mandatory = $true)][PSCustomObject]$SystemConfig
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	if ($SystemConfig.Modules -and $SystemConfig.Modules -is [array]) {
		Push-Info "Importing other modules from configuration file."
		$SystemConfig.Modules | ForEach {
			{
				param (
					[object]$moduleObj
				)
				try {
					if ($moduleObj.Path) {
						Push-Verbose "Importing module '$($moduleObj.Name)' from '$($moduleObj.Path)''."
						#for custom modules
						Import-Module $moduleObj.Path -Force
				} else {
						#for common modules
						if (-not (Get-Module $moduleObj.Name)) {
							if (-not (Get-InstalledModule $moduleObj.Name)) {
								if (-not (Test-IsAdministrator)) {
									Push-Error "You're not running as Administrator, module installation was failed."
								} else {
									Push-Verbose "Installing module '$($moduleObj.Name)'."
									#for other modules
									Install-Module $moduleObj.Name -Confirm:$false -Force
								}	
							}
							if ((Get-InstalledModule $moduleObj.Name | Measure-Object).Count -gt 0) {
								Push-Verbose "Importing module '$($moduleObj.Name)'."
								#when the module is installed
								Import-Module $moduleObj.Name -Force
							}
						} else {
							Push-Verbose "Module '$($moduleObj.Name)' was already imported."
						}
					}
				}
				catch {
					Push-Error "Module '$($moduleObj.Name)' was not imported."
				}
			}.Invoke({
					param (
						[object]$module
					)
					switch ($module.GetType().Name) {
						"String" { return [PSCustomObject]@{ Name = $module } }
						Default { return $module }
					}
				}.Invoke($_)
			)
		}
	}

}
Export-ModuleMember -Function Import-PSModules

Function Test-IsModulesLoaded {
	Param (
		[Parameter(Mandatory = $true)][string[]]$moduleNames
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	$retval = $true
	$moduleNames | ForEach { $retval = $retval -and ((Get-Module $_) -ne $null) }
	return $retval
}
Export-ModuleMember -Function Test-IsModulesLoaded

Function Resolve-DefaultIfEmpty {
	Param (
		[Parameter(Mandatory = $false)][object]$value,
		[Parameter(Mandatory = $false)][object]$defaultValue
	)
	Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState
	if ($value -ne $null) {
		$value
	} else {
		$defaultValue
	}
}
Export-ModuleMember -Function Resolve-DefaultIfEmpty

Function Test-If {
	Param (
		[Parameter(Mandatory = $true)][bool]$ifBlock,
		[Parameter(Mandatory = $true)][scriptblock]$then,
		[Parameter(Mandatory = $true)][scriptblock]$else
	)
	if ($ifBlock) {
		return $then.Invoke()
	} else {
		return $else.Invoke()
	}
}
Export-ModuleMember -Function Test-If