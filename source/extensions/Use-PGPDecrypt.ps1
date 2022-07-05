
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


try{
#Decrypt part
#Push-Verbose $decconfig | Out-String)
	if (-not (Test-IsValidPath $decconfig.DecryptConfiguration.OpenPGPExecutablePath)) {
		throw "OpenPGP Executable Path is not valid or you might not install OpenPGP yet."
	}
	if ($decconfig.DecryptConfiguration -and (Test-IsEnabled $decconfig.DecryptConfiguration.Enabled)) {
		#get the source of decrypt file
		@(switch ($decconfig.DecryptConfiguration.FilePathsSource) {
			"ImportLocations" {
				$decconfig.ImportConfiguration.ImportLocations | Where { (Test-IsEnabled $_.Enabled) } | ForEach {
					$_.Results | ForEach {
						[PSCustomObject]@{
							Passphrase = (ConvertTo-PlainPassword $decconfig.DecryptConfiguration.GlobalPassphrase)
							SourcePath = $_.DownloadedFilePath
							DestinationPath = (-join @(
								("$($decconfig.DecryptConfiguration.BaseOutputFolder)" | Format-String -Replacement $global:c.RuntimeConfiguration.PredefinedFormattings),
								(Split-Path $_.DownloadedFilePath -Leaf)
							))
						}
					}
				}

			}
			"FilePaths" {
				#logic taken
				$decconfig.DecryptConfiguration.FilePaths | Where { (Test-IsEnabled $_.Enabled) } | ForEach {
					$FilePath = $_
					[PSCustomObject]@{
						Passphrase = $(if ($FilePath.Passphrase -and $FilePath.Passphrase.Length -gt 0) {
							(ConvertTo-PlainPassword $FilePath.Passphrase)
						} else {
							(ConvertTo-PlainPassword $decconfig.DecryptConfiguration.GlobalPassphrase)
						})
						SourcePath = ("$($FilePath.Source)" | Format-String -Replacement $global:c.RuntimeConfiguration.PredefinedFormattings)
						DestinationPath = ("$($FilePath.Output)" | Format-String -Replacement $global:c.RuntimeConfiguration.PredefinedFormattings)
					}
				}
			}
			Default {}
		}) | ForEach {
			$SourceEncryptedFile = $_

			if (Test-IsValidPath $SourceEncryptedFile.SourcePath) {
				& "$($decconfig.DecryptConfiguration.OpenPGPExecutablePath)" `
					--batch --yes --ignore-mdc-error `
					--pinentry-mode=loopback `
					--passphrase $SourceEncryptedFile.Passphrase `
					--output "$($SourceEncryptedFile.DestinationPath)" `
					--decrypt "$($SourceEncryptedFile.SourcePath)"

				if (Test-IsValidPath $SourceEncryptedFile.DestinationPath) {
					Push-Info "File '$($SourceEncryptedFile.DestinationPath)' successfully decrypted."
				} else {
					Push-Warning "File '$($SourceEncryptedFile.SourcePath)' decryption failed."
				}
			} else {
				Push-Warning "Source file '$($SourceEncryptedFile.SourcePath)' does not exist."
			}
		}
	} else {
		Push-Warning "System can't find if 'DecryptConfiguration' exist in the config file."
	}
}
catch
{
	Push-Error $_ "Error when executing Import and Decrypt Script."
}

	