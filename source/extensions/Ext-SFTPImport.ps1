using namespace WinSCP
[Cmdletbinding()]
param(
	[Parameter(Mandatory = $false, Position = 0)][Alias("c")][string]$ConfigurationFilePath = ""
)

#default configuration file path
$ConfigFilePath = (& { if ($ConfigurationFilePath) { $ConfigurationFilePath } else { "configuration.json" } })
$ConfigFilePath = ($ConfigFilePath | Format-String -Replacement $global:c.RuntimeConfiguration.PredefinedFormattings)
#load configuration
if (-not (Test-IsValidPath $ConfigFilePath)) { Write-Host "Configuration '$($ConfigFilePath)' does not exist." -Fore Yellow; exit; }
$decconfig = Get-Content -Raw -Path $ConfigFilePath -ErrorAction Stop | ConvertFrom-Json

try
{
	#Import files
	if ($decconfig.ImportConfiguration -and $decconfig.ImportConfiguration.ImportLocations) {
		$decconfig.ImportConfiguration.ImportLocations | Where { (Test-IsEnabled $_.Enabled) } | ForEach {
			$ImportLocation = $_           
			switch ($ImportLocation.Type) {
				"SFTP" {
					if ((Test-NetConnection -ComputerName "$($ImportLocation.ConnectionSettings.HostName)" -Port $ImportLocation.ConnectionSettings.PortNumber -InformationLevel Quiet) -eq $true) {
						#Adding DownloadLocations to the setting                       
						$ImportLocation.ConnectionSettings | Add-Member -TypeName PSObject -NotePropertyMembers ([ordered]@{
							FormattedDownloadLocation = ($ImportLocation.ConnectionSettings.DownloadLocation | Format-String -Replacement $global:c.RuntimeConfiguration.PredefinedFormattings)
							SessionOptions = (New-Object SessionOptions -Property @{   
								Protocol              = [Protocol]::Sftp
								HostName              = $ImportLocation.ConnectionSettings.HostName
								PortNumber            = $ImportLocation.ConnectionSettings.PortNumber
								UserName              = $ImportLocation.ConnectionSettings.UserName
								Password              = (ConvertTo-PlainPassword $ImportLocation.ConnectionSettings.Password)
								SshHostKeyFingerprint = $ImportLocation.ConnectionSettings.SshHostKeyFingerprint                                
							})
							FtpSession = (New-Object Session)
							FtpTransferOptions = (New-Object TransferOptions -Property @{
								TransferMode = [TransferMode]::Binary                                
							})
						}) | Out-Null
						$ImportLocation | Add-Member -TypeName PSObject -NotePropertyMembers ([ordered]@{
							Results = @()
						})
                       
						Push-Info "Opening connection to '$($ImportLocation.ConnectionSettings.HostName):$($ImportLocation.ConnectionSettings.PortNumber)' with '$($ImportLocation.ConnectionSettings.UserName)' downloading to '$($ImportLocation.ConnectionSettings.FormattedDownloadLocation)' as destination."
						#connect				
						$ImportLocation.ConnectionSettings.FtpSession.Open($ImportLocation.ConnectionSettings.SessionOptions)
                         
						#Creating folder if it's not exist
						if (-not (Test-IsValidPath $ImportLocation.ConnectionSettings.FormattedDownloadLocation)) {
							Push-Verbose "Creating folder at '$($ImportLocation.ConnectionSettings.FormattedDownloadLocation)'."
							New-Item $ImportLocation.ConnectionSettings.FormattedDownloadLocation -ItemType Directory
						}

						#formatting FilePaths
						$ImportLocation.ConnectionSettings.FilePaths | ForEach {
							$FilePath = $_ | Format-String -Replacement $global:c.RuntimeConfiguration.PredefinedFormattings
							Push-Info "Downloading '$($FilePath)' to '$($ImportLocation.ConnectionSettings.FormattedDownloadLocation)'."
							try {
								$transferResult = $ImportLocation.ConnectionSettings.FtpSession.GetFiles($FilePath, $ImportLocation.ConnectionSettings.FormattedDownloadLocation, $false, $ImportLocation.ConnectionSettings.FtpTransferOptions)

								# Throw on any error
								$transferResult.Check()
			
								# Print results
								$transferResult.Transfers | ForEach {
									Push-Info "File '$($_.FileName)' was successfully downloaded to '$($_.Destination)'."
									$ImportLocation.Results += [PSCustomObject]@{
										FtpFilePath = $FilePath
										DownloadedFilePath = $_.Destination
									}
								}
							}
							catch {
								Push-Error $_ "Error when downloading '$($FilePath)', proceeding to the next process (if any)."
							}
						}
						if ($ImportLocation.ConnectionSettings.FtpSession) {
							$ImportLocation.ConnectionSettings.FtpSession.Dispose()
							Push-Info "Closing connection with '$($ImportLocation.ConnectionSettings.HostName):$($ImportLocation.ConnectionSettings.PortNumber)'."
						}

					} else {
						Push-Warning "Connection to '$($ImportLocation.ConnectionSettings.HostName):$($ImportLocation.ConnectionSettings.PortNumber)'($($ImportLocation.Type)) failed."
					}
				}
				Default {}
			}
		}
	} else {
		Push-Warning "System can't find if 'ImportConfiguration' exist in the config file."
	}
}
catch
{
	Push-Error $_ "Error when executing SFTPImport Script."
}
