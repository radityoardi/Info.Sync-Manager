using namespace System.DirectoryServices
using namespace System.DirectoryServices.AccountManagement

[Cmdletbinding()]
param(
	[Parameter(Mandatory = $true, Position = 0)][object]$Destination,
	[Parameter(Mandatory = $true, Position = 1)][object]$DestMappingKey,
	[Parameter(Mandatory = $true, Position = 2)][object]$StagingKey
)
Get-CallerPreference -Cmdlet $PSCmdlet -SessionState $ExecutionContext.SessionState

#Variables
[System.Text.StringBuilder]$LogMessage = New-Object System.Text.StringBuilder
$NewRecordKeySummary = [System.Collections.ArrayList]@()
$ModifiedRecordKeySummary = [System.Collections.ArrayList]@()
$DisabledRecordKeySummary = [System.Collections.ArrayList]@()
$FailedRecordKeySummary = [System.Collections.ArrayList]@()

$ReturnData = [PSCustomObject]@{
	TotalRecords = 0
	NewRecords = 0
	ModifiedRecords = 0
	DisabledRecords = 0
	FailedRecords = 0
	NewRecordKeySummary = ""
	ModifiedRecordKeySummary = ""
	DisabledRecordKeySummary = ""
	FailedRecordKeySummary = ""
}
[string]$dn


if ((Test-NetConnection $Destination.ConnectionSettings.HostName -Port (Resolve-DefaultIfEmpty $Destination.ConnectionSettings.Port 389)).TcpTestSucceeded) {
	#Opening connection for searches
	Optimize-Object ([PrincipalContext]$LDAP = New-Object PrincipalContext([ContextType]::Domain, $Destination.ConnectionSettings.HostName, $Destination.ConnectionSettings.UserName, (ConvertTo-PlainPassword $Destination.ConnectionSettings.Password))) {
		#Before Table CustomOperation
		Invoke-CustomOperations -ConfigContext ([ref]$Destination) -ExecutionMode $global:c.RuntimeConfiguration.CustomOperationTypes.Destination.PreTable -ExecutionStage Destination -CustomParameters @{ PredefinedFormattings = $global:c.RuntimeConfiguration.PredefinedFormattings }

		$pcRowDestProgress = 0

		#iterating through all staging records
		$allProcessedRows = $global:c.SyncConfiguration.Staging.RawData | Where { ($null -ne $_."$($StagingKey.Name)" -and $_."$($StagingKey.Name)".Length -gt 0) }
		$ReturnData.TotalRecords = ($allProcessedRows | Measure-Object).Count

		$allProcessedRows | ForEach {
			$stagingRow = $_
			[bool]$MultiUsersFound = $false
			[bool]$NewAccount = $false

			$pcRowDestProgress += 1
			@("Processing row $($pcRowDestProgress). '$($stagingRow."$($StagingKey.Name)")'") | ForEach {
				Write-Progress -Activity "Processing Rows" -Status $_ -Id 2 -ParentId 1
				Push-Verbose $_
				$LogMessage.AppendLine($_)
			}

			#Check if an account exist
			try {
				[UserExtPrincipal]$u = New-Object UserExtPrincipal($LDAP)
				$u.AttributeSet("$($DestMappingKey.DestinationProperty)", "$($stagingRow."$($StagingKey.Name)")")
				[PrincipalSearcher]$pcs = New-Object PrincipalSearcher($u)
				$results = $pcs.FindAll()
			}
			catch {
				Push-Error $_ "$($stagingRow."$($StagingKey.Name)"): Error when trying to find AD account '$($stagingRow."$($StagingKey.Name)")'."
			}

			@("Found $(($results | Measure-Object).Count) users.") | ForEach {
				Push-Verbose $_
				$LogMessage.AppendLine($_)
			}

			if (($results | Measure-Object).Count -gt 1) {
				$MultiUsersFound = $true
			}

			if (($results | Measure-Object).Count -eq 0) {
				#IF account does not exist, create it first
				$NewAccount = $true
				try {
					if ($Destination.GeneralSettings.AccountInfoProperties -and $Destination.GeneralSettings.AccountInfoProperties.AccountMustExistProperty -and $stagingRow."$($Destination.GeneralSettings.AccountInfoProperties.AccountMustExistProperty)" -eq $null) {
						Throw New-Object System.InvalidOperationException("The destination settings under GeneralSettings.AccountInfoProperties.AccountMustExistProperty is specified, but there is no staging property named '$($Destination.GeneralSettings.AccountInfoProperties.AccountMustExistProperty)'.")
					}
					if ($Destination.GeneralSettings.AccountInfoProperties -and (($Destination.GeneralSettings.AccountInfoProperties.AccountMustExistProperty -and $stagingRow."$($Destination.GeneralSettings.AccountInfoProperties.AccountMustExistProperty)" -eq $false) -or ($Destination.GeneralSettings.AccountInfoProperties.AccountMustExist -eq $false))) {
						@("The configuration of 'AccountMustExist' was set to false or specifically not to create any new account.") | ForEach {
							Push-Verbose $_
							$LogMessage.AppendLine($_)
						}
					} else {
						$DefaultAccountLocation = Resolve-DefaultIfEmpty $Destination.GeneralSettings.DefaultAccountLocation $stagingRow."$($Destination.GeneralSettings.StagingPropertyForDefaultAccountLocation)"
						#Account must exist, or the configuration is not defined
						@("The account will be created in '$($DefaultAccountLocation)'.") | ForEach {
							Push-Verbose $_
							$LogMessage.AppendLine($_)
						}

						Optimize-Object ([PrincipalContext]$NewUserLDAP = New-Object PrincipalContext([ContextType]::Domain, $Destination.ConnectionSettings.HostName, $DefaultAccountLocation, $Destination.ConnectionSettings.UserName, (ConvertTo-PlainPassword $Destination.ConnectionSettings.Password))) {
							Push-Verbose "Connected to '$($NewUserLDAP.Container)'."
							Optimize-Object ([UserExtPrincipal] $uAcc = New-Object UserExtPrincipal($NewUserLDAP)) {
								try {
									$userName = $null
									#required information to create user
									if ($stagingRow."sAMAccountName" -ne $null) {
										$userName = $stagingRow."sAMAccountName"
									} else {
										$userName = $stagingRow."$($StagingKey.Name)"
									}
									@("Creating a user with sAMAccountName '$($userName)'.") | ForEach {
										Push-Verbose $_
										$LogMessage.AppendLine($_)
									}			
									$uAcc.sAMAccountName = $userName
		
									if ($DestMappingKey.DestinationProperty -ne "sAMAccountName") {
										$_v = $stagingRow."$($DestMappingKey.StagingProperty)"
										if ($DestMappingKey.Format) {
											if ($_v -ne $null) {
												$LogMessage.AppendLine("Updating attribute '$($DestMappingKey.DestinationProperty)' with format '$($DestMappingKey.Format)'.")
												$_r = @{ "$($DestMappingKey.StagingProperty)" = $_v }
												$_f = ($DestMappingKey.Format | Format-String -Replacement $_r)
			
												$uAcc.AttributeSet("$($DestMappingKey.DestinationProperty)", $_v.ToString($_f))
												$LogMessage.AppendLine("Attribute '$($DestMappingKey.DestinationProperty)' was set with format '$($DestMappingKey.Format)', but not saved yet.")
											} else {

											}
										} else {
											$LogMessage.AppendLine("Updating attribute '$($DestMappingKey.DestinationProperty)'.")
											$_v = $stagingRow."$($DestMappingKey.StagingProperty)"
											$uAcc.AttributeSet("$($DestMappingKey.DestinationProperty)", $_v)
											$LogMessage.AppendLine("Attribute '$($DestMappingKey.DestinationProperty)' was updated, but not saved yet.")
										}
									}
		
									#setting the default password
									if ($Destination.GeneralSettings.DefaultPassword) {
										$uAcc.SetPassword("$($Destination.GeneralSettings.DefaultPassword)")
									}
		
									if ($Destination.GeneralSettings.StagingPropertyForCN -and $stagingRow."$($Destination.GeneralSettings.StagingPropertyForCN)") {
										$uAcc.Name = $stagingRow."$($Destination.GeneralSettings.StagingPropertyForCN)"
									}
									else {
										$uAcc.Name = $userName
									}

									if ($stagingRow."UserPrincipalName" -ne $null) {
										$uAcc.UserPrincipalName = $stagingRow."UserPrincipalName"
									} else {
										$uAcc.UserPrincipalName = "$($userName)@$($Destination.GeneralSettings.DefaultPrincipalName)"
									}

									if ($Destination.GeneralSettings.DefaultPasswordExpirationAfterCreate) {
										$uAcc.ExpirePasswordNow()
									}
									$uAcc.Save()
		
									@("$($stagingRow."$($StagingKey.Name)"): Account was successfully created.") | ForEach {
										Push-Verbose $_
										$LogMessage.AppendLine($_)
									}
									$ReturnData.NewRecords += 1
									$NewRecordKeySummary.Add($stagingRow."$($StagingKey.Name)")
								}
								catch {
									Push-Error $_ "$($stagingRow."$($StagingKey.Name)"): Error on creating new AD account '$($stagingRow."$($StagingKey.Name)")'."
								}
							}
							}

						#Check if the created account exists
						try {
							[UserExtPrincipal]$u = New-Object UserExtPrincipal($LDAP)
							$u.AttributeSet("$($DestMappingKey.DestinationProperty)", "$($stagingRow."$($StagingKey.Name)")")
							[PrincipalSearcher]$pcs = New-Object PrincipalSearcher($u)
							$results = $pcs.FindAll()
						}
						catch {
							Push-Error $_ "$($stagingRow."$($StagingKey.Name)"): Error when trying to find AD account '$($stagingRow."$($StagingKey.Name)")'."
						}

						@("Found $(($results | Measure-Object).Count) users after creation.") | ForEach {
							Push-Verbose $_
							$LogMessage.AppendLine($_)
						}
				
						if (($results | Measure-Object).Count -gt 1) {
							$MultiUsersFound = $true
						}

					}
				}
				catch {
					$ex = $_
					@("$($stagingRow."$($StagingKey.Name)"): Error while creating the account.") | ForEach {
						Push-Error $ex $_
						$LogMessage.AppendLine($_)
					}
				}
			}

			$pcRowDestFoundProgress = 0
			if ($results -and ($results | Measure-Object).Count -gt 0) {
				$updateAllPropsSuccess = $true
				$update1PropsSuccess = $false
				#Existing account
				$results | ForEach {
					[UserExtPrincipal]$result = $_ -as [UserExtPrincipal]
					
					if ($MultiUsersFound) {
						$pcRowDestFoundProgress += 1
						@("Processing row $($pcRowDestProgress) result $($pcRowDestFoundProgress). '$($stagingRow."$($StagingKey.Name)")'") | ForEach {
							Write-Progress -Activity "Processing Rows" -Status $_ -Id 3 -ParentId 2
							Push-Verbose $_
							$LogMessage.AppendLine($_)
						}	
					}
				
					#IF account exist
					try {
						@("Updating '$($result.UserPrincipalName)' ($($result.DistinguishedName)).") | ForEach {
							Push-Verbose $_ -NoEventLog
							$LogMessage.AppendLine($_)
						}
						
						#Pre-Row
						Invoke-CustomOperations -Object ([ref]$result) -ConfigContext ([ref]$Destination) -StagingRow ([ref]$stagingRow) -ExecutionMode $global:c.RuntimeConfiguration.CustomOperationTypes.Destination.PreRow -ExecutionStage Destination -CustomParameters @{ PredefinedFormattings = $global:c.RuntimeConfiguration.PredefinedFormattings; StagingRow = $stagingRow }

						#Set account's status Enabled/Disabled
						try {
							if ($Destination.GeneralSettings.AccountInfoProperties -eq $null -or $Destination.GeneralSettings.AccountInfoProperties.AccountEnabledProperty -eq $null) {
								Push-Verbose "The destination settings under GeneralSettings.AccountInfoProperties.AccountEnabledProperty is not specified, account Enabled/Disabled status will be skipped."
							} elseif ($Destination.GeneralSettings.AccountInfoProperties -and $Destination.GeneralSettings.AccountInfoProperties.AccountEnabledProperty -and $null -eq $stagingRow."$($Destination.GeneralSettings.AccountInfoProperties.AccountEnabledProperty)") {
								Throw New-Object System.InvalidOperationException("The destination settings under GeneralSettings.AccountInfoProperties.AccountEnabledProperty is specified, but there is no staging property named '$($Destination.GeneralSettings.AccountInfoProperties.AccountEnabledProperty)'.")
							} else {
								#Updating account Enabled/Disabled
								if ($result.Enabled -ne $stagingRow."$($Destination.GeneralSettings.AccountInfoProperties.AccountEnabledProperty)") {
									$LogMessage.AppendLine("Updating account's Enabled/Disabled status.")
									$result.Enabled = $stagingRow."$($Destination.GeneralSettings.AccountInfoProperties.AccountEnabledProperty)"
									$result.Save()
									$LogMessage.AppendLine("Account's Enabled/Disabled status was updated.")

									if ($stagingRow."$($Destination.GeneralSettings.AccountInfoProperties.AccountEnabledProperty)" -eq $false) {
										$ReturnData.DisabledRecords += 1
									}
									$update1PropsSuccess = $true
								}
							}
						}
						catch {
							$updateAllPropsSuccess = $false
							Throw New-Object System.InvalidOperationException("Error on updating existing AD account Enabled/Disabled status for '$($stagingRow."$($StagingKey.Name)")'.", $_.Exception)
						}

						#Set account's attributes
						try {
							if ($Destination.GeneralSettings.AccountInfoProperties -and $Destination.GeneralSettings.AccountInfoProperties.UpdatePropertiesProperty -and $stagingRow."$($Destination.GeneralSettings.AccountInfoProperties.UpdatePropertiesProperty)" -eq $null) {
								Throw New-Object System.InvalidOperationException("The destination settings under GeneralSettings.AccountInfoProperties.UpdatePropertiesProperty is specified, but there is no staging property named '$($Destination.GeneralSettings.AccountInfoProperties.UpdatePropertiesProperty)'.")
							}
							if ($Destination.GeneralSettings.AccountInfoProperties -and (($Destination.GeneralSettings.AccountInfoProperties.UpdatePropertiesProperty -and $stagingRow."$($Destination.GeneralSettings.AccountInfoProperties.UpdatePropertiesProperty)" -eq $false) -or ($Destination.GeneralSettings.AccountInfoProperties.UpdateProperties -eq $false))) {
								@("The configuration of 'UpdateProperties' was set to false or specifically not to update any attributes.") | ForEach {
									Push-Verbose $_
									$LogMessage.AppendLine($_)
								}
							} else {
								$Destination.Mappings | Where { ($NewAccount -eq $false -and $_.StagingProperty -ne $DestMappingKey.DestinationProperty -and (-not $_.ExcludeWhenUpdating)) -or ($NewAccount -eq $true -and (-not $_.ExcludeWhenCreating)) } | ForEach {
									$destMapping = $_
									try {
										$LogMessage.AppendLine("Processing staging attribute '$($destMapping.StagingProperty)' into destination attribute '$($destMapping.DestinationProperty)'.")
										#If the staging attribute is null
										if ($NewAccount -eq $false -and ($null -eq $stagingRow."$($destMapping.StagingProperty)" -or $stagingRow."$($destMapping.StagingProperty)".Length -eq 0)) {
											$LogMessage.AppendLine("Clearing attribute '$($destMapping.DestinationProperty)'.")
											#clear the value
											$result.AttributeSet("$($destMapping.DestinationProperty)", $null)
											$result.Save()
											$LogMessage.AppendLine("Attribute '$($destMapping.DestinationProperty)' was cleared.")
											$update1PropsSuccess = $true

										} elseif ($null -ne $stagingRow."$($destMapping.StagingProperty)" -and $stagingRow."$($destMapping.StagingProperty)".Length -gt 0) {
											if ($destMapping.Format) {
												$LogMessage.AppendLine("Updating attribute '$($destMapping.DestinationProperty)' with format '$($destMapping.Format)'.")
												$_v = $stagingRow."$($destMapping.StagingProperty)"
												if ($_v -eq $null) {
													$result.AttributeSet("$($destMapping.DestinationProperty)", $null)
													$result.Save()
													$LogMessage.AppendLine("Attribute '$($destMapping.DestinationProperty)' was cleared.")
													$update1PropsSuccess = $true
												} else {
													$_r = @{ "$($destMapping.StagingProperty)" = $_v }
													$_f = ($destMapping.Format | Format-String -Replacement $_r)
	
													#only if the attribute's value is different with staging
													if ($result.AttributeGet("$($destMapping.DestinationProperty)") -ne $_v.ToString($_f)) {
														$result.AttributeSet("$($destMapping.DestinationProperty)", $_v.ToString($_f))
														$result.Save()
														$LogMessage.AppendLine("Attribute '$($destMapping.DestinationProperty)' was updated with format '$($destMapping.Format)'.")
														$update1PropsSuccess = $true

													} else {
														$LogMessage.AppendLine("Attribute '$($destMapping.DestinationProperty)' value is the same as staging.")
													}
												}
											} else {
												$LogMessage.AppendLine("Updating attribute '$($destMapping.DestinationProperty)'.")
												#only if the attribute's value is different with staging
												if ($result.AttributeGet("$($destMapping.DestinationProperty)") -ne $stagingRow."$($destMapping.StagingProperty)") {
													$_v = $stagingRow."$($destMapping.StagingProperty)"
													$result.AttributeSet("$($destMapping.DestinationProperty)", $_v)
													$result.Save()
													$LogMessage.AppendLine("Attribute '$($destMapping.DestinationProperty)' was updated.")
													$update1PropsSuccess = $true

												} else {
													$LogMessage.AppendLine("Attribute '$($destMapping.DestinationProperty)' value is the same as staging.")
												}
											}
										} else {
											$LogMessage.AppendLine("Attribute '$($destMapping.DestinationProperty)' was not updated since it's a new account with empty value.")
										}
									}
									catch {
										$updateAllPropsSuccess = $false
										Throw New-Object System.InvalidOperationException("$($stagingRow."$($StagingKey.Name)"): Error on updating existing AD account '$($stagingRow."$($StagingKey.Name)")' on attribute '$($destMapping.DestinationProperty)'.", $_.Exception)
									}
								}
							}
						}
						catch {
							$updateAllPropsSuccess = $false
							Throw New-Object System.InvalidOperationException("Error on updating existing AD account attributes for '$($stagingRow."$($StagingKey.Name)")'.", $_.Exception)
						}

						#Post-Row
						Invoke-CustomOperations -Object ([ref]$result) -ConfigContext ([ref]$Destination) -StagingRow ([ref]$stagingRow) -ExecutionMode $global:c.RuntimeConfiguration.CustomOperationTypes.Destination.PostRow -ExecutionStage Destination -CustomParameters @{ PredefinedFormattings = $global:c.RuntimeConfiguration.PredefinedFormattings }

						@("$($stagingRow."$($StagingKey.Name)"): Account was successfully updated.") | ForEach {
							Push-Verbose $_
							$LogMessage.AppendLine($_)
						}
					}
					catch {
						$ex = $_
						@("$($stagingRow."$($StagingKey.Name)"): Error while updating the account.") | ForEach {
							Push-Error $ex $_
							$LogMessage.AppendLine($_)
						}
					}

					if ($updateAllPropsSuccess -eq $false) {
						$ReturnData.FailedRecords += 1
						$FailedRecordKeySummary.Add($stagingRow."$($StagingKey.Name)")
					}
					if ($update1PropsSuccess -eq $true) {
						$ReturnData.ModifiedRecords += 1
						$ModifiedRecordKeySummary.Add($stagingRow."$($StagingKey.Name)")
					}
				}
			}

			#Update Log
			Push-Warning $LogMessage.ToString() -NoWriteHost
			$LogMessage.Clear()
		}

		$ReturnData.NewRecordKeySummary = (Test-If ($NewRecordKeySummary.Count -gt 0) {$NewRecordKeySummary -join ", "} {"None"})
		$ReturnData.ModifiedRecordKeySummary = (Test-If ($ModifiedRecordKeySummary.Count -gt 0) {$ModifiedRecordKeySummary -join ", "} {"None"})
		$ReturnData.DisabledRecordKeySummary = (Test-If ($DisabledRecordKeySummary.Count -gt 0) {$DisabledRecordKeySummary -join ", "} {"None"})
		$ReturnData.FailedRecordKeySummary = (Test-If ($FailedRecordKeySummary.Count -gt 0) {$FailedRecordKeySummary -join ", "} {"None"})

		@($u, $pcs) | Clear-Object
		$LogMessage | Clear-Object
		$NewRecordKeySummary.Clear()
		$ModifiedRecordKeySummary.Clear()
		$DisabledRecordKeySummary.Clear()
		$FailedRecordKeySummary.Clear()
		
	}
} else {
	Push-Warning "The server at '$($Destination.ConnectionSettings.HostName)' is offline."
}
Push-Verbose $ReturnData
return $ReturnData
