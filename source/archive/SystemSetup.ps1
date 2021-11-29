$ErrorActionPreference = "Stop"

Write-Host "Installing SqlServer PowerShell module..." -NoNewLine
if (-not (Get-InstalledModule SqlServer)) {
	Install-Module SqlServer
	Write-Host " [Installed]" -ForegroundColor Green
} else {
	Write-Host " [Already Installed]" -ForegroundColor Yellow
}

Write-Host "Installing PreferenceVariables PowerShell module..." -NoNewLine
if (-not (Get-InstalledModule PreferenceVariables)) {
	Install-Module PreferenceVariables
	Write-Host " [Installed]" -ForegroundColor Green
} else {
	Write-Host " [Already Installed]" -ForegroundColor Yellow
}
#Ensure .NET Framework 4.0, 4.5, 4.7.1 is installed
#Ensure GnuPG & WinGPG installed

if (-not (Test-Path "C:\Program Files (x86)\GnuPG\bin\gpg.exe")) {
	Write-Error "GnuPG (OpenPGP) is not installed."
}
Write-Host "Installing Event Log Sources..."
@(
	"Info Sync Manager",
	"Start-InfoSync.ps1",
	"GlobalModules.psm1"
) | ForEach {
	Write-Host "  Installing Sources `"$($_)`"." -NoNewLine
	try {
		New-EventLog -LogName "Info Sync Manager" -Source $_
		Write-EventLog -LogName "Info Sync Manager" -Source $_ -EntryType "Information" -Message "This is the first event log to test the functionality." -EventId 1001
		Write-Host " [Installed]" -ForegroundColor Green
	}
	catch {
		if ($_.Exception -and $_.Exception.Message -match "source is already registered on") {
			Write-Host " [Already Installed]" -ForegroundColor Yellow
		} else {
			Write-Host " [Error] $($_.Exception.Message)" -ForegroundColor Red -BackgroundColor Black
		}
		
	}
}

Write-Host "Finished!" -ForegroundColor Green