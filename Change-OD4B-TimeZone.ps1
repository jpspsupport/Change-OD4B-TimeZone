<#
 This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. 
 THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
 INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  
 We grant you a nonexclusive, royalty-free right to use and modify the sample code and to reproduce and distribute the object 
 code form of the Sample Code, provided that you agree: 
    (i)   to not use our name, logo, or trademarks to market your software product in which the sample code is embedded; 
    (ii)  to include a valid copyright notice on your software product in which the sample code is embedded; and 
    (iii) to indemnify, hold harmless, and defend us and our suppliers from and against any claims or lawsuits, including 
          attorneys' fees, that arise or result from the use or distribution of the sample code.
Please note: None of the conditions outlined in the disclaimer above will supercede the terms and conditions contained within 
             the Premier Customer Services Description.
#>
param(
    [Parameter(Mandatory=$true)]
    [String]$SPOAdminUrl,
    [Parameter(Mandatory=$true)]
    [String]$SPOAdminUser,
    [int]$TimeZoneId = 20
)

# Load the required assemblies. 
# Note that SharePoint Online CSOM 16.1.8361.1200 is the required version of this sample.
Add-Type -Path "C:\csom\lib\net45\Microsoft.SharePoint.Client.dll";
Add-Type -Path "C:\csom\lib\net45\Microsoft.SharePoint.Client.Runtime.dll";

function ExecuteQueryWithIncrementalRetry
{
    param (
        [int]$retryCount,
        [int]$delay = 120
    );

    $RetryAfterHeaderName = "Retry-After";
    $retryAttempts = 0;
    $backoffInterval = $delay
    $retryAfterInterval = 0;
    $retry = $false;

    if ($retryCount -le 0) {
        throw "Provide a retry count greater than zero."
    }
    if ($delay -le 0) {
        throw "Provide a delay greater than zero."
    }

    while ($retryAttempts -lt $retryCount) {
        try {
            if (!$retry)
            {
                $script:context.ExecuteQuery();
                return;
            }
            else
            {
                if (($wrapper -ne $null) -and ($wrapper.Value -ne $null))
                {
                    $script:context.RetryQuery($wrapper.Value);
                    return;
                }
            }
        }
        catch [System.Net.WebException] {
            $response = $_.Exception.Response

            if (($null -ne $response) -and (($response.StatusCode -eq 429) -or ($response.StatusCode -eq 503))) {

                $wrapper = [Microsoft.SharePoint.Client.ClientRequestWrapper]($_.Exception.Data["ClientRequest"]);
                $retry = $true


                $retryAfterHeader = $response.GetResponseHeader($RetryAfterHeaderName);
                $retryAfterInMs = $DefaultRetryAfterInMs;

                if (-not [string]::IsNullOrEmpty($retryAfterHeader)) {
                    if (-not [int]::TryParse($retryAfterHeader, [ref]$retryAfterInterval)) {
                        $retryAfterInterval = $DefaultRetryAfterInMs;
                    }
                }
                else
                {
                    $retryAfterInterval = $backoffInterval;
                }

                Write-Output ("CSOM request exceeded usage limits. Sleeping for {0} seconds before retrying." -F ($retryAfterInterval))
                #Add delay.
                Start-Sleep -m ($retryAfterInterval * 1000)
                #Add to retry count.
                $retryAttempts++;
                $backoffInterval = $backoffInterval * 2;
            }
            else {
                throw;
            }
        }
    }

    throw "Maximum retry attempts {0}, have been attempted." -F $retryCount;
}


function Set-OD4BTimeZone
{
    param (
        [String]$siteUrl
    );

    # Generate the credential for CSOM.
    $creds2 = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($SPOAdminUser, $SPOAdminPassword);

    # Connect SPO site.
    $script:context = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl);
    $script:context.Credentials = $creds2;

    # Set UserAgent.
    $script:context.add_ExecutingWebRequest({
        param ($source, $eventArgs);
        $request = $eventArgs.WebRequestExecutor.WebRequest;
        $request.UserAgent = "NONISV|Contoso|Application/1.0";
    });

    # Get time zone.
    $timeZone = $script:context.Web.RegionalSettings.TimeZone;
    $timeZones = $script:context.Web.RegionalSettings.TimeZones;

    $script:context.Load($timeZone);
    $script:context.Load($timeZones);

    ExecuteQueryWithIncrementalRetry -retryCount 5;

    Write-Host ("Before Time Zone ({0}) : {1}" -F $timeZone.Id, $timeZone.Description); 

    # When the time zone is "(UTC-08:00) Pacific Time (US and Canada)" :
    if ($timeZone.Id -eq 13)
    {
        # Set the specified time zone.
        $script:context.Web.RegionalSettings.TimeZone = $timeZones.GetById($TimeZoneId);
        $script:context.Web.Update();

        ExecuteQueryWithIncrementalRetry -retryCount 5;

        Write-Host "Time Zone has changed."; 
    }
    else
    {
        Write-Host "skipped.";
    }
}


# ----------------------
# Main.
# ----------------------

# Input password via console.
$SPOAdminPassword = Read-Host -Prompt "Please enter your password" -AsSecureString;

# Generate the credential for SPO Management shell.
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SPOAdminUser, $SPOAdminPassword;

# Connect SPO admin center.
Connect-SPOService -Url $SPOAdminUrl -Credential $creds;

# Get all URLs of OD4B sites.
$OD4BSites = Get-SPOSite -IncludePersonalSite $true -Template SPSPERS -Limit All;

foreach ($site in $OD4BSites)
{
    $tempAdminUser = $null;

    Write-Host ("OD4B Site : {0}" -F $site.Url) -ForegroundColor Green;

    # Is "$SPOAdminUser" the site collection administrator?
    $isSiteAdmin = (Get-SPOUser -Site $site.Url -LoginName $SPOAdminUser).IsSiteAdmin;

    try
    {
        # When "$SPOAdminUser" is NOT the site collection administrator :
        if (-not $isSiteAdmin)
        {
            # Add "$SPOAdminUser" to the site collection administrators of OD4B site temporarily.
            $tempAdminUser = Set-SPOUser -Site $site.Url -LoginName $SPOAdminUser -IsSiteCollectionAdmin $true;

            Write-Host "Added site collection administrator."; 
        }

        # Change the time zone.
        Set-OD4BTimeZone -siteUrl $site.Url;
    }
    catch
    {
        Write-Host $Error[0] -ForegroundColor Red;
    }
    finally
    {
        # Remove "$SPOAdminUser" from the site collection administrators of OD4B site.
        if ((-not $isSiteAdmin) -and $tempAdminUser)
        {
            Set-SPOUser -Site $site.Url -LoginName $SPOAdminUser -IsSiteCollectionAdmin $false | Out-Null;

            Write-Host "Removed site collection administrator."; 
        }
    }
}
