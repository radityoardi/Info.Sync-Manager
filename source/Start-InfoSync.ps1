#Require -Version 1.0

[Cmdletbinding()]
Param(
	[Parameter(Mandatory = $false, Position = 0)][Alias("c")][string]$ConfigurationFilePath = "",
	[Parameter(Mandatory = $false, Position = 1)][Alias("r")][int]$MaxMainSourceRows = 0
)

#default configuration file path
$ConfigFilePath = (&{ if ($ConfigurationFilePath) { $ConfigurationFilePath } else { "configuration.json" } })
#Load Configuration
if (-not (Test-Path $ConfigFilePath)) { Write-Host "Configuration '$($ConfigFilePath)' does not exist." -Fore Yellow; exit; }
$global:c = Get-Content -Raw -Path $ConfigFilePath -ErrorAction Stop | ConvertFrom-Json 

$global:c | Add-Member -TypeName PSObject -NotePropertyMembers ([ordered]@{
	RuntimeConfiguration = [PSCustomObject]@{
		CurrentScriptFile = (Split-Path $MyInvocation.MyCommand.Definition -Leaf)
		LogRefNumber = (Get-Random -Minimum 0 -Maximum 65535).ToString("00000")
		PredefinedFormattings = [hashtable]@{
			Today         = [System.DateTime]::Today
			Now           = { [System.DateTime]::Now }
			Yesterday     = [System.DateTime]::Today.AddDays(-1)
			Tomorrow      = [System.DateTime]::Today.AddDays(1)
			CurrentFolder = $PSScriptRoot
		}
		Subscripts = [PSCustomObject]@{
			Source = [PSCustomObject[]]@(
				[PSCustomObject]@{
					Type      = "CSV"
					Stage     = "RawImport"
					Subscript = "{CurrentFolder}\nodes\CSV-Import.ps1"
				},
				[PSCustomObject]@{
					Type      = "LDAP"
					Stage     = "RawImport"
					Subscript = "{CurrentFolder}\nodes\LDAP-Import.ps1"
				}
			)
			Destination = [PSCustomObject[]]@(
				[PSCustomObject]@{
					Type      = "LDAP"
					Subscript = "{CurrentFolder}\nodes\LDAP-Export.ps1"
				},
				[PSCustomObject]@{
					Type      = "CSV"
					Subscript = "{CurrentFolder}\nodes\CSV-Export.ps1"
				}
			)
			Dependencies = [PSCustomObject[]]@(
				[PSCustomObject]@{
					ModuleName = "GlobalModules"
					Path       = ".\\library\\GlobalModules.psd1"
				}
			)
		}
		CustomOperationTypes = [PSCustomObject]@{
			RawSource = [PSCustomObject]@{
				PreTable   = "Pre-RawSource"
				PostTable  = "Post-RawSource"
			}
			Source = [PSCustomObject]@{
				PreTable   = "Pre-Source"
				PostRow    = "Post-Row"
			}
			Staging = [PSCustomObject]@{
				PostRow    = "Post-Row"
				PostTable = "Post-Staging"
			}
			Destination = [PSCustomObject]@{
				PreTable  = "Pre-Destination"
				PreRow    = "Pre-Row"
				PostRow   = "Post-Row"
				PostTable = "Post-Destination"
			}
		}
	}
}) | Out-Null

#Read Modules
$ReadModuleMeasurement = Measure-Command {
	if ($global:c.RuntimeConfiguration.Subscripts.Dependencies) {
		@("Importing module dependencies.") | ForEach {
			Write-Host $_
			Write-EventLog -LogName $global:c.SystemConfiguration.WindowsLog.Name -Source $global:c.RuntimeConfiguration.CurrentScriptFile -EventId $global:c.RuntimeConfiguration.LogRefNumber -Message $_
		}
		$global:c.RuntimeConfiguration.Subscripts.Dependencies | ForEach {
			$moduleObj = $_
			try {
				if ($moduleObj.Path) {
					@("Importing module dependency '$($moduleObj.ModuleName)' from '$($moduleObj.Path)'.") | ForEach {
						Write-Verbose $_
						Write-EventLog -LogName $global:c.SystemConfiguration.WindowsLog.Name -Source $global:c.RuntimeConfiguration.CurrentScriptFile -EventId $global:c.RuntimeConfiguration.LogRefNumber -Message $_
					}
					#for custom modules
					Import-Module $moduleObj.Path -Force -DisableNameChecking
				}
			}
			catch {
				$("Module dependency '$($moduleObj.ModuleName)' was not imported.") | ForEach {
					Write-Error $_
					Write-EventLog -LogName $global:c.SystemConfiguration.WindowsLog.Name -Source $global:c.RuntimeConfiguration.CurrentScriptFile -EventId $global:c.RuntimeConfiguration.LogRefNumber -Message $_ -EntryType Error
				}
			}
		}
	}

	Import-PSModules $global:c.SystemConfiguration
}

#Clear EventLog
if (Test-IsEnabled $global:c.SystemConfiguration.WindowsLog.ClearLogsOnStart) {
	@("$($global:c.RuntimeConfiguration.CurrentScriptFile) was started with Process ID $($PID) and Event ID $($global:c.RuntimeConfiguration.LogRefNumber).") | ForEach {
		if (-not (Test-IsAdministrator)) {
			Push-Verbose "$($_) The script does not run as Administrator, hence Clear Event Log is not working."
		} else {
			Clear-EventLog -LogName $global:c.SystemConfiguration.WindowsLog.Name
			Push-Verbose $_
		}	
	}
}

#Read Assemblies
$ReadAssembliesMeasurement = Measure-Command {
	Import-Assemblies $global:c.SystemConfiguration
}



if (Test-IsModulesLoaded ($global:c.RuntimeConfiguration.Subscripts.Dependencies | Select-Object -ExpandProperty ModuleName)) {
	try {
####################################################################################################
# Start of the script
####################################################################################################
		Push-Verbose "Info Sync started with configuration: $($ConfigFilePath)."

		#checks
		if ($global:c.SyncConfiguration.Staging) {
			if (($global:c.SyncConfiguration.Staging.Properties | Where-Object { $_.Key -eq $true } | Measure-Object).Count -ne 1) {
				throw "SyncConfiguration.Staging.Properties must contain a property with its 'Key' to 'True'."
			} else {
				#Retrieving Key Property of Staging
				$KeyStagingProperty = $global:c.SyncConfiguration.Staging.Properties | Where { $_.Key -eq $true } | Select-Object -First 1
				Push-Verbose "The Key property is: '$($KeyStagingProperty.Name)'"
			}
		}

		if ($global:c.SyncConfiguration.Source) {
			$global:c.SyncConfiguration.Source | Where { (Test-IsEnabled $_.Enabled) -and (($_.Mappings | Where-Object { $_.StagingProperty -eq $KeyStagingProperty.Name } | Measure-Object).Count -lt 1) } | ForEach {
				throw "SyncConfiguration.Source.Mappings for '$($_.Description)' does not contain a mapping to the key with its StagingProperty '$($KeyStagingProperty.Name)'."
			}
		}

		#main component
		if ($global:c.SyncConfiguration.Source -and $global:c.SyncConfiguration.Staging) { #Destinations is not required now.
####################################################################################################
# PULLING RAW SOURCE
####################################################################################################
			#first level counter
			$pcSourcesProgress = 0
			$pcSourcesTotal = ($global:c.SyncConfiguration.Source | Measure-Object).Count

			#Load RawData into config variable
			$LoadRawDataMeasurement = Measure-Command {
				#Currently supports only CSV
				$global:c.SyncConfiguration.Source | Where { (Test-IsEnabled $_.Enabled) } | ForEach {
					$pcSourcesProgress += 1
					Write-Progress -Activity "Pulling SOURCES" -Status "Pulling $($pcSourcesProgress). '$($source.Description)' ($($source.Type))" -PercentComplete (($pcSourcesProgress / $pcSourcesTotal) * 100) -Id 1
					
					$source = $_ #as local variable

					switch ($source.Type) {
						{"CSV" -or "LDAP"} {
							$SourceSubscript = ($global:c.RuntimeConfiguration.Subscripts.Source | Where { $_.Type -eq $source.Type -and $_.Stage -eq "RawImport" } | Select -First 1).Subscript
							if ($null -ne $SourceSubscript -and (Test-Path ($SourceSubscript | Format-String -Replacement $global:c.RuntimeConfiguration.PredefinedFormattings))) {
								#Execute PreTable for RawSource
								Invoke-CustomOperations -ConfigContext ([ref]$source) -ExecutionMode $global:c.RuntimeConfiguration.CustomOperationTypes.RawSource.PreTable -ExecutionStage RawSource -CustomParameters @{ PredefinedFormattings = $global:c.RuntimeConfiguration.PredefinedFormattings }

								if ($null -ne $source.FilterExpression) {
									Push-Verbose "Importing data with filter expression: '$($source.FilterExpression)'."
								}
								
								#Execute Subscript Path (in order to make it extendable in future)
								& "$(($SourceSubscript | Format-String -Replacement $global:c.RuntimeConfiguration.PredefinedFormattings))" -s $source

								#Execute PostTable for RawSource
								Invoke-CustomOperations -ConfigContext ([ref]$source) -ExecutionMode $global:c.RuntimeConfiguration.CustomOperationTypes.RawSource.PostTable -ExecutionStage RawSource -CustomParameters @{ PredefinedFormattings = $global:c.RuntimeConfiguration.PredefinedFormattings }

							} else {
								Push-Warning "Script '$($SourceSubscript)' specified in '$($source.Description)' could not be found."
							}
						}
						Default {
							Push-Warning "Source '$($source.Description)' with its type '$($source.Type)' is not supported at the moment during this pulling raw source phase."
						}
					}
					Push-Info "Imported from '$($source.Description)' ($($source.Type)): $(($source.RawData | Measure-Object).Count) rows."
				}
			}
			Push-Info "Pulling SOURCES was completed in $(GetMessageFromMeasurement $LoadRawDataMeasurement)."

####################################################################################################
# PROCESS INTO STAGING
####################################################################################################

			$StagingMeasurement = Measure-Command {
				$pcMainSourceProgress = 0

				#PROCESSING MAIN SOURCE
				$global:c.SyncConfiguration.Source | Where { $_.MainSource -eq $true -and (Test-IsEnabled $_.Enabled) } | ForEach {
					$source = $_ #as local variable
	
					$pcMainSourceProgress += 1
					@("Process STAGING $($pcMainSourceProgress). '$($source.Description)' ($($source.Type))") | ForEach {
						Write-Progress -Activity "Processing Main Source into STAGING" -Status $_ -Id 1
						Push-Info $_
					}
	
					$KeySourceProperty = $source.Mappings | Where-Object { $_.StagingProperty -eq $KeyStagingProperty.Name } | Select-Object -First 1

					#Executing PreTable
					Invoke-CustomOperations -ConfigContext ([ref]$source) -ExecutionMode $global:c.RuntimeConfiguration.CustomOperationTypes.Source.PreTable -ExecutionStage Source -CustomParameters @{ PredefinedFormattings = $global:c.RuntimeConfiguration.PredefinedFormattings }
	
					#Iterating RawData main source and add it as RawData member in Staging. This will be the staging table
					$pcRowProgress = 0

					if ($source.RawData) {
						$_temparray = [System.Collections.ArrayList]@()
						#If MainSource RawData exist
						$source.RawData | ForEach {
							$sourceRow = $_ #as local variable
		
							$pcRowProgress += 1
							@("Processing row $($pcRowProgress). '$($sourceRow."$($KeySourceProperty.SourceProperty)")'") | ForEach {
								Write-Progress -Activity "Processing Rows" -Status $_ -Id 2 -ParentId 1
								Push-Verbose $_
							}

							#create new hashtable
							$data = @{}
							#[PSCustomObject]$data = [PSCustomObject]@{} #no longer being used now.
		
							$source.Mappings | Where { (Test-IsEnabled $_.Enabled) } | ForEach {
								$sourceMapping = $_
								$global:c.SyncConfiguration.Staging.Properties | Where { $_.Name -eq $sourceMapping.StagingProperty } | ForEach {
									$stagingProp = $_
									Set-AllProperties -Object ([ref]$data) -CsvRow $sourceRow -StagingProperty $stagingProp -SourceMapping $sourceMapping
								}
							}
	
							#Execute PostRow for MainSource
							Invoke-CustomOperations -Object ([ref]$data) -ConfigContext ([ref]$source) -ExecutionMode $global:c.RuntimeConfiguration.CustomOperationTypes.Source.PostRow -ExecutionStage Source -CustomParameters @{ PredefinedFormattings = $global:c.RuntimeConfiguration.PredefinedFormattings }

							#PROCESSING SECONDARY SOURCES
							$global:c.SyncConfiguration.Source | Where { $_.MainSource -ne $true -and (Test-IsEnabled $_.Enabled) } | Sort-Object -Property @{ Expression = { if ($_.LoadOrder) { $_.LoadOrder } else { 100 } } } | ForEach {
								$secSource = $_
								Write-Progress -Activity "Processing Secondary Sources" -Status "Processing '$($secSource.Description)' ($($secSource.Type) - $(if ($secSource.LoadOrder) { $secSource.LoadOrder } else { 100 }))" -Id 3 -ParentId 2
								$KeySecSourceProperty = $secSource.Mappings | Where-Object { $_.StagingProperty -eq $KeyStagingProperty.Name } | Select-Object -First 1
	
								#Executing PreTable
								Invoke-CustomOperations -ConfigContext ([ref]$secSource) -ExecutionMode $global:c.RuntimeConfiguration.CustomOperationTypes.Source.PreTable -ExecutionStage Source -CustomParameters @{ PredefinedFormattings = $global:c.RuntimeConfiguration.PredefinedFormattings }
	
								if ($secSource.RawData) {
									$secSource.RawData | Where { $_."$($KeySecSourceProperty.SourceProperty)" -eq $sourceRow."$($KeySourceProperty.SourceProperty)" } | ForEach {
										$secCsvRow = $_
										$secSource.Mappings | Where { (Test-IsEnabled $_.Enabled) } | ForEach {
											$sourceSecMapping = $_
											$global:c.SyncConfiguration.Staging.Properties | Where { $_.Name -eq $sourceSecMapping.StagingProperty } | ForEach {
												$stagingSecProp = $_
												Set-AllProperties -Object ([ref]$data) -CsvRow $secCsvRow -StagingProperty $stagingSecProp -SourceMapping $sourceSecMapping -Skip:($stagingSecProp.Name -eq $KeyStagingProperty.Name)
											}
										}
	
										#Execute PostRow for any other Secondary Sources
										Invoke-CustomOperations -Object ([ref]$data) -ConfigContext ([ref]$secSource) -ExecutionMode $global:c.RuntimeConfiguration.CustomOperationTypes.Source.PostRow -ExecutionStage Source -CustomParameters @{ PredefinedFormattings = $global:c.RuntimeConfiguration.PredefinedFormattings }
									}	
								} else {
									Push-Warning "Secondary Source '$($source.Description)' ($($source.Type)) data is empty."
								}
							}

							#dump hashtable to output
							Push-Verbose "$(($data | ConvertTo-PsCustomObjectFromHashtable) | Out-String)"
							$_temparray.Add($data)
						}						

						$global:c.SyncConfiguration.Staging | Add-Member RawData -MemberType NoteProperty -Value $_temparray
						Remove-Variable _temparray
					} else {
						Push-Warning "Primary Source '$($source.Description)' ($($source.Type)) data is empty."
					}
				}
			}

			Push-Info ((@(
				"Process STAGING was completed in $(GetMessageFromMeasurement $StagingMeasurement).",
				"Process STAGING: $(($global:c.SyncConfiguration.Staging.RawData | Measure-Object).Count) rows."
			) | ForEach {
				Push-Info $_ -NoEventLog
				return $_
			}) -join "`r`n") -NoWriteHost

####################################################################################################
# EXPORT STAGING
####################################################################################################
			if ($global:c.SyncConfiguration.Staging.CustomOperations -and $global:c.SyncConfiguration.Staging.CustomOperations -is [array]) {
				if (($global:c.SyncConfiguration.Staging.CustomOperations | Where { (Test-IsEnabled $_.Enabled) -and $_.Execution -eq "Post-Row" } | Measure-Object).Count -gt 0) {
					$global:c.SyncConfiguration.Staging.RawData | ForEach {
						#Execute PostRow for Staging
						Invoke-CustomOperations -Object ([ref]$_) -ConfigContext ([ref]$global:c.SyncConfiguration.Staging) -ExecutionMode $global:c.RuntimeConfiguration.CustomOperationTypes.Staging.PostRow -ExecutionStage Staging -CustomParameters @{ PredefinedFormattings = $global:c.RuntimeConfiguration.PredefinedFormattings }
						Push-Verbose "$(($_ | ConvertTo-PsCustomObjectFromHashtable) | Out-String)"
					}
				}
			}

			if ($global:c.SyncConfiguration.Staging.ExportCsv -eq $true -and $global:c.SyncConfiguration.Staging.CsvPath) {
				$ActualCsvPath = ($global:c.SyncConfiguration.Staging.CsvPath | Format-String -Replacement $global:c.RuntimeConfiguration.PredefinedFormattings)
				($global:c.SyncConfiguration.Staging.RawData | ConvertTo-PsCustomObjectFromHashtable) | Export-Csv $ActualCsvPath -NoTypeInformation | Out-Null
			}

####################################################################################################
# PUSHING TO DESTINATION
####################################################################################################

			$pcDestinationProgress = 0
			$DestinationMeasurement = Measure-Command {
				$global:c.SyncConfiguration.Destination | Where { (Test-IsEnabled $_.Enabled) } | Sort-Object -Property @{ Expression = { if ($_.RunOrder) { $_.RunOrder } else { 100 } } } | ForEach {
					$dest = $_ #move to variable avoiding confusion

					$pcDestinationProgress += 1
					@("Pushing $($pcDestinationProgress). '$($dest.Description)' ($($dest.Type) - $(if ($dest.RunOrder) { $dest.RunOrder } else { 100 }))") | ForEach {
						Write-Progress -Activity "Pushing DESTINATION" -Status $_ -Id 1
						Push-Verbose $_
					}

					$KeyDestProperty = $dest.Mappings | Where-Object { $_.StagingProperty -eq $KeyStagingProperty.Name } | Select-Object -First 1

					switch ($dest.Type) {
						{"LDAP" -or "SQLDB"} {
							$DestinationSubScript = ($global:c.RuntimeConfiguration.Subscripts.Destination | Where { $_.Type -eq $dest.Type } | Select -First 1).Subscript
							if ($null -ne $DestinationSubScript -and (Test-Path ($DestinationSubScript | Format-String -Replacement $global:c.RuntimeConfiguration.PredefinedFormattings))) {
								#Execute Subscript Path (in order to make it extendable in future)
								#Executing PreTable
								Invoke-CustomOperations -ConfigContext ([ref]$dest) -ExecutionMode $global:c.RuntimeConfiguration.CustomOperationTypes.Destination.PreTable -ExecutionStage Destination -CustomParameters @{
									PredefinedFormattings = $global:c.RuntimeConfiguration.PredefinedFormattings
								}

								$destreturnval = (& "$(($DestinationSubScript | Format-String -Replacement $global:c.RuntimeConfiguration.PredefinedFormattings))" -Destination $dest -StagingKey $KeyStagingProperty -DestMappingKey $KeyDestProperty)

								Invoke-CustomOperations -ConfigContext ([ref]$dest) -ExecutionMode $global:c.RuntimeConfiguration.CustomOperationTypes.Destination.PostTable -ExecutionStage Destination -CustomParameters @{
									ReturnValue = $destreturnval
									PredefinedFormattings = $global:c.RuntimeConfiguration.PredefinedFormattings
								}
							} else {
								Push-Warning "Script '$($DestinationSubScript)' specified in '$($source.Description)' could not be found."
							}
						}
						Default {
							Push-Warning "Destination '$($source.Description)' with its type '$($source.Type)' is not supported at the moment."
						}
					}

				}
			}
			Push-Info "Push DESTINATION was completed in $(GetMessageFromMeasurement $DestinationMeasurement)."
		}	
	}
	catch {
		Push-Error $_	
	}

	Push-Verbose "$($global:c.RuntimeConfiguration.CurrentScriptFile) is finished with Process ID $($PID) and Event ID $($global:c.RuntimeConfiguration.LogRefNumber)."
} 
else {
	Push-Verbose "$($global:c.RuntimeConfiguration.CurrentScriptFile) is finished with Process ID $($PID) and Event ID $($global:c.RuntimeConfiguration.LogRefNumber). However, it did not execute due to the missing module dependencies."
}

Remove-Variable c -Scope Global

