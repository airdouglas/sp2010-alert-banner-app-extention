Param
(
    [Parameter(Mandatory=$true,
                ValueFromPipelineByPropertyName=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$ActionType
)

$env:certificateBase64Encode = "YOUR CERT CODE"
$env:GRAPH_APP_ID     = 'YOUR GUID'#'9421c078-dab8-42dc-a00c-434c8ebdb614'
$env:tenant           = "{{YOUR_TENANT_NAME}}"
# TODO Change the folder name of the JS file
$AddSP2010BannerMessage = "C:\SP2010RetirementBanner\AddSP2010BannerMessage.js"
$env:SPFXAPPEXTENTIONAPPID = "DA7C13EC-D4E7-401D-BF68-92716ED46248"

# the following check is only to debu on Power Shell ISE
if ( $env:certificateBase64Encode -ne $null)
{
    $certificateBase64Encode = $env:certificateBase64Encode
}
else
{
    # get the PFX secret from the key vault
    $kvSecret = Get-AzKeyVaultSecret -VaultName $env:KeyVaultName -Name $env:KeyVaultSecretName
    $certificateBase64Encode = '';
    $ssPtr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($kvSecret.SecretValue)
    try {
        $certificateBase64Encode = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ssPtr)
    } finally {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ssPtr)
    }
}

function Add-SP2010WKBanner
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("WebUrl2Process")]
        [string]$WebUrl
    )
    begin
    {
    }
    process
    {
        try
        {

            $HashArguments = @{
                    Url                      = $WebUrl
                    ClientId                 = $env:GRAPH_APP_ID
                    CertificateBase64Encoded = $certificateBase64Encode
                    Tenant                   = $("{0}.onmicrosoft.com" -f  $env:tenant)
            }
            $newSiteConn = Connect-PnPOnline @HashArguments  -ReturnConnection

            #### The following lines are to add a message to the site
            # add new folder to the "SiteAssets"
            Add-PnPfolder -Connection $newSiteConn -Folder "SiteAssets" -Name "SP2010WKBannerFiles" -ErrorAction Ignore
            # upload the JS file to the new folder
            Add-PnPFile -Connection $newSiteConn -Path $AddSP2010BannerMessage -Folder "SiteAssets/SP2010WKBannerFiles"  -ErrorAction Ignore
            # add the link to the JS file for site scope
            Add-PnPJavaScriptLink -Connection $newSiteConn  -Name SP2010WKBannerFiles -Url 'SiteAssets/SP2010WKBannerFiles/AddSP2010BannerMessage.js' -Sequence 9999 -Scope Site  -ErrorAction Ignore


            ## add SPFx Extention for secured for 
            Install-PnPApp -Identity "${env:SPFXAPPEXTENTIONAPPID}" -Connection $newSiteConn -Verbose -ErrorAction Ignore



            Disconnect-PnPOnline -Connection $newSiteConn

        }
        catch
        {
        }

    }
}

function Remove-SP2010WKBanner
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("WebUrl2Process")]
        [string]$WebUrl
    )
    begin
    {
    }
    process
    {
        try
        {

            $HashArguments = @{
                    Url                      = $WebUrl
                    ClientId                 = $env:GRAPH_APP_ID
                    CertificateBase64Encoded = $certificateBase64Encode
                    Tenant                   = $("{0}.onmicrosoft.com" -f  $env:tenant)
            }
            $newSiteConn = Connect-PnPOnline @HashArguments  -ReturnConnection

            $jslFound = Get-PnPJavaScriptLink -Name SP2010WKBannerFiles -Scope All -ErrorAction Ignore -Connection $newSiteConn

            if ( $null -ne $jslFound )
            {
                Remove-PnPJavaScriptLink -Identity $jslFound.Id -Connection $newSiteConn -Force -Scope All
            }
            $file = Get-PnPFile -Connection $newSiteConn -Url 'SiteAssets/SP2010WKBannerFiles/AddSP2010BannerMessage.js' -ErrorAction Ignore
            if ( $file )
            {
                Remove-PnPFile -SiteRelativeUrl 'SiteAssets/SP2010WKBannerFiles/AddSP2010BannerMessage.js' -Force -Connection $newSiteConn
            }
            $folder = Get-PnPFolder -Url "SiteAssets/SP2010WKBannerFiles" -ErrorAction Ignore -Connection $newSiteConn
            if ( $folder )
            {
                Remove-PnPFolder -Folder "SiteAssets" -Name "SP2010WKBannerFiles" -Force -Connection $newSiteConn
            }

            ## add SPFx Extention for secured for 
            Uninstall-PnPApp -Identity "${env:SPFXAPPEXTENTIONAPPID}" -Connection $newSiteConn -Verbose -ErrorAction Ignore

            Get-PnPRecycleBinItem -Connection $newSiteConn | Where-Object Title -like "sp-2010-alert-banner-app-extention.sppkg" |  Clear-PnpRecycleBinItem -Connection $newSiteConn -Force

            Disconnect-PnPOnline -Connection $newSiteConn
        }
        catch
        {
        }

    }
}


function Publish-AddRemoveBanner
{
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [Alias("ActionType")]
        [string]$Action
    )
    begin
    {

    }
    process
    {
        try
        {
            Import-Csv C:\SP2010RetirementBanner\site2process.csv | 
            ForEach-Object {
                try
                {
                    if ( $Action -eq "Add" )
                    {
                        Write-Host $("Site URL {0} for {1}" -f $_.Url, $Action )
                        Add-SP2010WKBanner -WebUrl $_.Url
                    }
                    if ( $Action -eq "Remove" )
                    {
                        Write-Host $("Site URL {0} for {1}" -f $_.Url, $Action )
                        Remove-SP2010WKBanner -WebUrl $_.Url
                    }
                }
                catch
                {
                    $ErrorMessage = $_.Exception | format-list -force
                    Write-Host $("Exception {0}" -f $ErrorMessage);
                }
                finally
                {
                }
            }
        }
        catch
        {
        }

    }
}
# Add or Remove
Publish-AddRemoveBanner -Action $ActionType