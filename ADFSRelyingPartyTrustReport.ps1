#Requires -Modules ImportExcel

$SB = {
    $RPTrusts = Get-AdfsRelyingPartyTrust
    $EnabledRPTrsusts = $RPTrusts | Where { $_.Enabled -eq $true }
    $DisabledRPTrusts = $RPTrusts | Where { $_.Enabled -eq $False }

    $RPTrusts

}

$Session = New-PSSession -ComputerName "uk1-p-adf001"

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
