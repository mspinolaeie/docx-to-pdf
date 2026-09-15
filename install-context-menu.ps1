# Registers "EIE tools > docx2pdf" in the Explorer right-click menu for .docx files.
# Per-user (HKCU), no admin needed. Run with -Uninstall to remove.
param([switch]$Uninstall)

$group = 'HKCU:\Software\Classes\SystemFileAssociations\.docx\shell\EIETools'

if ($Uninstall) {
    Remove-Item -LiteralPath $group -Recurse -ErrorAction SilentlyContinue
    Write-Host 'Removed EIE tools from .docx context menu.'
    return
}

$launcher = Join-Path $PSScriptRoot 'docx-to-pdf-drop.cmd'
$tool = "$group\shell\docx2pdf"

New-Item -Path "$tool\command" -Force | Out-Null
Set-ItemProperty -LiteralPath $group -Name MUIVerb -Value 'EIE tools'
Set-ItemProperty -LiteralPath $group -Name SubCommands -Value ''
Set-ItemProperty -LiteralPath $tool -Name MUIVerb -Value 'docx2pdf'
Set-ItemProperty -LiteralPath "$tool\command" -Name '(default)' -Value "cmd.exe /c `"`"$launcher`" `"%1`"`""

Write-Host "Installed: right-click .docx > EIE tools > docx2pdf -> $launcher"
