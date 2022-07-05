using namespace Management.Automation.Host
using namespace Collections.ObjectModel

[Cmdletbinding()]
param(
	[Parameter(Mandatory = $false, Position = 0)][Alias("i")][string]$InputFile,
	[Parameter(Mandatory = $false, Position = 1)][Alias("o")][string]$OutputFile
)

#includes
Import-Module "$($PSScriptRoot)\Extensions\GlobalModules.psm1" -Force -Verbose:$VerbosePreference

if ($InputFile)
{
	$InputObject = Get-Content -Raw -Path $InputFile -ErrorAction Stop | ConvertFrom-Json
	$OutputObject = [PSCustomObject]@{ Accounts = @() }
	if ($InputObject -and $InputObject.Accounts -and ($InputObject.Accounts | Measure-Object).Count -gt 0)
	{
		ForEach ($account in $InputObject.Accounts)
		{
			[PSCustomObject]$outputAcct = @{
				Password = (ConvertTo-SecureString $account.Password -AsPlainText -Force)
				UserName = $account.UserName
			}
			$OutputObject.Accounts += $outputAcct
		}

		if ($OutputObject -and $OutputObject.Accounts)
		{
			ConvertTo-Json $OutputObject | Out-File $OutputFile
		}
	}
}
else
{
	Write-Host "Type your password below, it will be secured as text (encrypted) with currently logged on user account."
	$Pwd = Read-Host -AsSecureString

	$SecurePwd = ConvertFrom-SecureString $Pwd
	if ((Start-YesNoChoice "Your password is secured." "Do you want to copy to clipboard?" -Verbose:$VerbosePreference) -eq 0)
	{
		Set-Clipboard $SecurePwd
	}
	else
	{
		$SecurePwd
	}	
}

