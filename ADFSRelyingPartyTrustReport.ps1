<#
  .SYNOPSIS
  Dumps the details of all Relying Party Trusts to an Excel Spreadsheet.
  .DESCRIPTION
  Dumps the details of all Relying Party Trusts of a given ADFS Farm to a nicely formatted and filtered Excel Spreadsheet.
  .PARAMETER
  None
  .EXAMPLE
  ADFSRelyingPartyTrustReport.ps1
  .INPUTS
  None
  .OUTPUTS
  ADFSReport.xlsx
  .NOTES
  Author:        Patrick Horne
  Creation Date: 22/01/2020
  Requires:      ImportExcel Module
  Change Log:
  V1.0:          Initial Development
#>

#Requires -Modules ImportExcel

$ADFSServer = Read-Host "Enter the name of the primary server of your ADFS farm."

$SB = {
    $RPTrusts = Get-AdfsRelyingPartyTrust
    # $EnabledRPTrsusts = $RPTrusts | Where { $_.Enabled -eq $true }
    # $DisabledRPTrusts = $RPTrusts | Where { $_.Enabled -eq $False }

    $RPTrusts

}

$Session = New-PSSession -ComputerName $ADFSServer

if ($Session) {
    $invokeCommandSplat = @{
        ErrorAction = 'SilentlyContinue'
        Session     = $Session
        ScriptBlock = $SB
    }
}


$exportExcelSplat = @{
        Path            = "ADFSReport.xlsx"
        BoldTopRow      = $true
        AutoSize        = $true
        FreezeTopRow    = $true
        WorkSheetname   = "ADFSReport"
        TableName       = "ADFSTable"
        TableStyle      = "Medium6"
    }

Invoke-Command @invokeCommandSplat | Export-Excel @exportExcelSplat
