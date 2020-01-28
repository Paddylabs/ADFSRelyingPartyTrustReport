<#
  .SYNOPSIS
  Dumps the details of all Relying Party Trusts to an Excel Spreadsheet.
  .DESCRIPTION
  Dumps the details of all Relying Party Trusts of a given ADFS Farm to a nicely formatted and filtered Excel Spreadsheet.
  .PARAMETER ADFSServer
  Define the primary ADFS Server in your ADFS Farm as this is the only server in the farm that you can query
  .EXAMPLE
  ADFSRelyingPartyTrustReport.ps1 -ADFSServer myprimaryadfsserver
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
  V2.0:          Input from brettmillerb, swapped read-host for param.
#>

#Requires -Modules ImportExcel

param (
    [String]$ADFSServer
)

$SB = {
    $RPTrusts = Get-AdfsRelyingPartyTrust
    # $EnabledRPTrsusts = $RPTrusts | Where { $_.Enabled -eq $true }
    # $DisabledRPTrusts = $RPTrusts | Where { $_.Enabled -eq $False }

    $RPTrusts | Select-Object -Property @( 
    'Name'
    'Identifier'
    'Enabled'
    'MonitoringEnabled'
    'LastMonitoredTime'
    'AutoUpdateEnabled'
    'LastUpdateTime'
    'MetadataUrl'
    'ProtocolProfile'
    'SignedSamlrequestsRequired'
    'SamlResponseSignature'
    'EncryptedNameIdRequired'
    'EncryptionCertificateRevocationCheck'
    'SigningCertificateRevocationCheck'
    'RequestSigningCertificate'
    'EncryptClaims'
    'TokenLifetime'
    'NotBeforeSkew'
    'EnableJWT'
    'AllowedClientTypes'
    'EncryptionCertificate'
    'PublishedThroughProxy'
    'AllowedAuthenticationClassReferences'
    'AlwaysRequireAuthentication'
    'ConflictWithPublishedPolicy'
    'LastPublishedPolicyCheckSuccessful'
    'WSFedEndpoint'
    'AdditionalWSFedEndpoint'
    'ProxyTrustedEndpoints'
    'SamlEndpoints'
    'ClaimsAccepted'
    'ProxyEndpointMappings'
    'IssueOAuthRefreshTokensTo'
    'SignatureAlgorithm'
    'OrganizationInfo'
    'Notes'
    'IssuanceTransformRules'
    'IssuanceAuthorizationRules'
    'ImpersonationAuthorizationRules'
    'DelegationAuthorizationRules'
    'AdditionalAuthenticationRules'
    'ClaimsProviderName'
    )

}

$Session = New-PSSession -ComputerName $ADFSServer

if ($Session) {
    $invokeCommandSplat = @{
        ErrorAction     = 'SilentlyContinue'
        Session         = $Session
        ScriptBlock     = $SB
    }
}

$SelectObjectSplat = @{
        Property        = "*"
        ExcludeProperty = "PSComputerName","PSShowComputerName","RunSpaceID"
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

Invoke-Command @invokeCommandSplat | Select-Object @SelectObjectSplat | Export-Excel @exportExcelSplat

Remove-PSSession $Session
