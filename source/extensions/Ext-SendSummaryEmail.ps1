[Cmdletbinding()]
param(
	[Parameter(Mandatory = $false, Position = 0)][Alias("c")][string]$ConfigurationFilePath = "",
	$CustomParameters,
	$StagingRow
)

#default configuration file path
$ConfigFilePath = (& { if ($ConfigurationFilePath) { $ConfigurationFilePath } else { "configuration.json" } })
$ConfigFilePath = ($ConfigFilePath | Format-String -Replacement $global:c.RuntimeConfiguration.PredefinedFormattings)
#load configuration
if (-not (Test-IsValidPath $ConfigFilePath)) { Write-Host "Configuration '$($ConfigFilePath)' does not exist." -Fore Yellow; exit; }
$decconfig = Get-Content -Raw -Path $ConfigFilePath -ErrorAction Stop | ConvertFrom-Json

[pscredential]$credObject = New-Object System.Management.Automation.PSCredential($decconfig.EmailConfiguration.SMTP.Credential.Username, (ConvertTo-SecureString $decconfig.EmailConfiguration.SMTP.Credential.Password))

try
{
	#Send email
	if ($decconfig.EmailConfiguration -and $decconfig.EmailConfiguration.SMTP) {
		$filepath = ($decconfig.EmailConfiguration.SMTP.BodyTemplateFile | Format-String -Replacement $global:c.RuntimeConfiguration.PredefinedFormattings)
		
		$ReplacementObject = @{
			TotalRecords	= $CustomParameters.ReturnValue.TotalRecords
			NewRecords = $CustomParameters.ReturnValue.NewRecords
			ModifiedRecords = $CustomParameters.ReturnValue.ModifiedRecords
			DisabledRecords = $CustomParameters.ReturnValue.DisabledRecords
			FailedRecords = $CustomParameters.ReturnValue.FailedRecords
			NewRecordKeySummary = $CustomParameters.ReturnValue.NewRecordKeySummary
			ModifiedRecordKeySummary = $CustomParameters.ReturnValue.ModifiedRecordKeySummary
			DisabledRecordKeySummary = $CustomParameters.ReturnValue.DisabledRecordKeySummary
			FailedRecordKeySummary = $CustomParameters.ReturnValue.FailedRecordKeySummary
		}
		$filecontent = (Get-Content -Raw -Path $filepath | Format-String -Replacement $ReplacementObject)		

		$mailParams = @{                        
			SmtpServer                 = $decconfig.EmailConfiguration.SMTP.Server
			Port                       = $decconfig.EmailConfiguration.SMTP.Port # or '25' if not using TLS
			UseSSL                     = $decconfig.EmailConfiguration.SMTP.UseSSL ## or not if using non-TLS
			Credential                 = $credObject
			From                       = $decconfig.EmailConfiguration.SMTP.Credential.Username
			To                         = $decconfig.EmailConfiguration.SMTP.ToAddresses
			Subject                    = ($decconfig.EmailConfiguration.SMTP.Subject | Format-String -Replacement $global:c.RuntimeConfiguration.PredefinedFormattings)
			Body                       = $filecontent
			BodyAsHtml				         = $true
			DeliveryNotificationOption = 'OnFailure', 'OnSuccess'

		}                   
		Send-MailMessage @mailParams
	} else {
		Push-Warning "System can't find if 'EmailConfiguration' exist in the config file."
	}
}
catch {
		Push-Error $_	
	}

