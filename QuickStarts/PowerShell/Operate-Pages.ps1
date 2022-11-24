<#
.COPYRIGHT
Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
See LICENSE in the project root for license information.
#>

####################################################

function Get-AuthToken {  
    
    $User = Read-Host -Prompt "Please specify your user principal name for Azure Authentication"

    return @{
        'Content-Type'  = 'application/json'
        'Accept'        = 'application/json; odata.metadata=none'
        'Authorization' = "Bearer " + "eyJ0eXAiOiJKV1QiLCJub25jZSI6IktzNUF6TFFTTkhEWTJDRkF5OHlwR0NDZmQzeU1Qc0xXRE93eVV5NXhTNEEiLCJhbGciOiJSUzI1NiIsIng1dCI6IjJaUXBKM1VwYmpBWVhZR2FYRUpsOGxWMFRPSSIsImtpZCI6IjJaUXBKM1VwYmpBWVhZR2FYRUpsOGxWMFRPSSJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9mNDE3MTgwZS1lZjIwLTRiNjQtOTgyYi1kNWNiOGQ1OTM1YTgvIiwiaWF0IjoxNjY5MjkzNzg1LCJuYmYiOjE2NjkyOTM3ODUsImV4cCI6MTY2OTI5NzcwOSwiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFUUUF5LzhUQUFBQUFuM1NSK1N6SXYwZnJzNjl2VzF0b2RwVCtFMUhjT1FwMlBXVDJBMW1BSnJYeHFoZjFQSHR2RHl6QmxvM2RTT0EiLCJhbXIiOlsicHdkIl0sImFwcF9kaXNwbGF5bmFtZSI6IkdyYXBoIEV4cGxvcmVyIiwiYXBwaWQiOiJkZThiYzhiNS1kOWY5LTQ4YjEtYThhZC1iNzQ4ZGE3MjUwNjQiLCJhcHBpZGFjciI6IjAiLCJmYW1pbHlfbmFtZSI6IkRvZSIsImdpdmVuX25hbWUiOiJKb2huIiwiaWR0eXAiOiJ1c2VyIiwiaXBhZGRyIjoiMTY3LjIyMC4yNTUuMzUiLCJuYW1lIjoiVEVTVF9URVNUX1NQT1Byb3ZIZWFydGJlYXRfRTNfMTVfMjIxMTEwMTYwMF85ODgyIiwib2lkIjoiYzQwYWQxNDYtNmMyZC00MWNiLWE5NjAtNmQwMDU4Y2ZiYjA3IiwicGxhdGYiOiIzIiwicHVpZCI6IjEwMDMyMDAyNDlGQjQ3MjMiLCJyaCI6IjAuQVh3QURoZ1g5Q0R2WkV1WUs5WExqVmsxcUFNQUFBQUFBQUFBd0FBQUFBQUFBQUM3QVBRLiIsInNjcCI6IkFwcGxpY2F0aW9uLlJlYWRXcml0ZS5BbGwgQXBwUm9sZUFzc2lnbm1lbnQuUmVhZFdyaXRlLkFsbCBvcGVuaWQgcHJvZmlsZSBTaXRlcy5SZWFkLkFsbCBTaXRlcy5SZWFkV3JpdGUuQWxsIFVzZXIuUmVhZCBlbWFpbCIsInNpZ25pbl9zdGF0ZSI6WyJrbXNpIl0sInN1YiI6ImRLYmlfbXRNVDBsZ1ZrbjdONU0zSFZTdEMzQnIxbDRwS1NkQlJqaThVUU0iLCJ0ZW5hbnRfcmVnaW9uX3Njb3BlIjoiTkEiLCJ0aWQiOiJmNDE3MTgwZS1lZjIwLTRiNjQtOTgyYi1kNWNiOGQ1OTM1YTgiLCJ1bmlxdWVfbmFtZSI6ImFkbWluQGE4MzBlZGFkOTA1MDg0OTg4MkUyMjExMTAxNi5vbm1pY3Jvc29mdC5jb20iLCJ1cG4iOiJhZG1pbkBhODMwZWRhZDkwNTA4NDk4ODJFMjIxMTEwMTYub25taWNyb3NvZnQuY29tIiwidXRpIjoiQklyemtpemlCa0dWSktWY3lfY21BZyIsInZlciI6IjEuMCIsIndpZHMiOlsiNjJlOTAzOTQtNjlmNS00MjM3LTkxOTAtMDEyMTc3MTQ1ZTEwIiwiYjc5ZmJmNGQtM2VmOS00Njg5LTgxNDMtNzZiMTk0ZTg1NTA5Il0sInhtc19zdCI6eyJzdWIiOiJEbzU1M0hEVHo5UlRUY2E4WW04ZjUtZVVQUlFudDZfVUxLRmp2VnJabTg4In0sInhtc190Y2R0IjoxNjY4MTI1MzI0fQ.rIAKngxvgyxwuGmhs1pQTfHnv6cMzkO0q84Wme77xQpiA5ppHsjNK8KDfgooFGQQgsM-Q2pYhGFLOQi26qRTX699xSHaGixpYVT-vgf1llEl-D6b4mQgK7rcPLADzdQDD5s8wWBaP60koDHp_BXrO1XBlfwvVEAa8YZ2taXpvLtRIgjJv1_WzxDEQx1pXwwbbVM8Fv_xu_vNAG_wYovy6QAUoE7lyBwli-sErdB1DNzRgMAkHbgP7HU1oVGp3gdRzC9kEg4d-eEK6hey-qAYyj-Kb-XpoNY0gECtv9diN7ZcY2omrma-YNs1K4An9tbDWje15f3jzy_VzUy1JdNzFg"
    }

  
    $userUpn = New-Object "System.Net.Mail.MailAddress" -ArgumentList $User
  
    $tenant = $userUpn.Host
  
    Write-Host "Checking for AzureAD module..."
  
    $AadModule = Get-Module -Name "AzureAD" -ListAvailable
  
    if ($AadModule -eq $null) {
        Write-Host "AzureAD PowerShell module not found, looking for AzureADPreview"
        $AadModule = Get-Module -Name "AzureADPreview" -ListAvailable
    }
    if ($AadModule -eq $null) {
        write-host
        write-host "AzureAD Powershell module not installed..." -f Red
        write-host "Install by running 'Install-Module AzureAD' or 'Install-Module AzureADPreview' from an elevated PowerShell prompt" -f Yellow
        write-host "Script can't continue..." -f Red
        write-host
        exit
    }
  
    # Getting path to ActiveDirectory Assemblies
    # If the module count is greater than 1 find the latest version
    if ($AadModule.count -gt 1) {
        $Latest_Version = ($AadModule | select version | Sort-Object)[-1]
        $aadModule = $AadModule | ? { $_.version -eq $Latest_Version.version }
  
        # Checking if there are multiple versions of the same module found
        if ($AadModule.count -gt 1) {
            $aadModule = $AadModule | select -Unique
        }
  
        $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
    }
    else {
        $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"
    }
  
    [System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
  
    [System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null
  
    $clientId = "ENTER_YOUR_APP_ID_HERE"
  
    $redirectUri = "urn:ietf:wg:oauth:2.0:oob"
  
    $resourceAppIdURI = "https://graph.microsoft.com"
  
    $authority = "https://login.microsoftonline.com/$Tenant"
  
    try {
        $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority
  
        # https://msdn.microsoft.com/en-us/library/azure/microsoft.identitymodel.clients.activedirectory.promptbehavior.aspx
        # Change the prompt behaviour to force credentials each time: Auto, Always, Never, RefreshSession
  
        $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Auto"
  
        $userId = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier" -ArgumentList ($User, "OptionalDisplayableId")
  
        $authResult = $authContext.AcquireTokenAsync($resourceAppIdURI, $clientId, $redirectUri, $platformParameters, $userId).Result
  
        # If the accesstoken is valid then create the authentication header
        if ($authResult.AccessToken) {
  
            # Creating header for Authorization token
            $authHeader = @{
                'Content-Type'  = 'application/json'
                'Authorization' = "Bearer " + $authResult.AccessToken
                'ExpiresOn'     = $authResult.ExpiresOn
            }
            return $authHeader
        }
        else {
            Write-Host
            Write-Host "Authorization Access Token is null, please re-run authentication..." -ForegroundColor Red
            Write-Host
            break
        }
    }
    catch {
        write-host $_.Exception.Message -f Red
        write-host $_.Exception.ItemName -f Red
        write-host
        break
    }
}
  
Function Get-Pages($siteId) {
  
    [cmdletbinding()]
  
    $graphApiVersion = "beta"
    $resource = "sites/$siteId/pages"

    $authToken = Get-AuthToken
  
    try {
        $uri = "https://graph.microsoft.com/$graphApiVersion/$($resource)"
        $pages = (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get).value
        $count = $pages.Count
        Write-Host "Get all $count pages."
        return $pages
    } 
    catch {
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-Host "Response content:`n$responseBody" -f Red
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
        write-host
        break
    }
}
  
Function Get-Page([string]$siteId, [string]$pageId) {
    
    [cmdletbinding()]

    $rootUrl = "https://canary.graph.microsoft.com/testprodbetapages-api-df/sites"
    $resource = "$($siteId)/pages/$($pageId)?expand=canvasLayout"
  
    $authToken = Get-AuthToken
    
    try {
        $uri = "$rootUrl/$($resource)"
        $page = (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Get)
        Write-Host "Get page(ID: $pageId) successfully."
        return $page
    } 
    catch {
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-Host "Response content:`n$responseBody" -f Red
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
        write-host
        break
    }
}
  
Function Remove-Page([string]$siteId, [string]$pageId) {
    [cmdletbinding()]
    
    $graphApiVersion = "beta"
    $resource = "sites/$($siteId)/pages/$($pageId)"
  
    $authToken = Get-AuthToken
    
    try {
        $uri = "https://graph.microsoft.com/$graphApiVersion/$($resource)"
        Invoke-RestMethod -Uri $uri -Headers $authToken -Method Delete
        Write-Host "Delete page($pageId) successfully."
    } 
    catch {
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-Host "Response content:`n$responseBody" -f Red
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
        write-host
        break
    }
} 

Function Publish-Page([string]$siteId, [string]$pageId) {
    [cmdletbinding()]
    
    $graphApiVersion = "beta"
    $resource = "sites/$($siteId)/pages/$($pageId)/publish"
  
    $authToken = Get-AuthToken
    
    try {
        $uri = "https://graph.microsoft.com/$graphApiVersion/$($resource)"
        Invoke-RestMethod -Uri $uri -Headers $authToken -Method Post
        Write-Host "Publish page($pageId) successfully."
    } 
    catch {
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-Host "Response content:`n$responseBody" -f Red
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
        write-host
        break
    }
}

Function Update-Page([string]$siteId, [string]$pageId, [object]$payload) {

    [cmdletbinding()]
    $rootUrl = "https://canary.graph.microsoft.com/testprodbetapages-api-df/sites"
    $resource = "$($siteId)/pages/$($pageId)"
  
    $authToken = Get-AuthToken
    
    try {
        $uri = "$rootUrl/$($resource)"
        $page = (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Patch -Body $payload)
        Write-Host "Promote page(ID: $($page.id), URL: $($page.webUrl)) successfully."
        return $page
    } 
    catch {
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-Host "Response content:`n$responseBody" -f Red
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
        write-host
        break
    }
}
Function New-Page([string]$siteId, [object]$payload) {

    [cmdletbinding()]
    $rootUrl = "https://canary.graph.microsoft.com/testprodbetapages-api-df/sites"
    $resource = "$($siteId)/pages"
  
    $authToken = Get-AuthToken
    
    try {
        $uri = "$rootUrl/$($resource)"
        $page = (Invoke-RestMethod -Uri $uri -Headers $authToken -Method Post -Body $payload)
        Write-Host "Create page(ID: $($page.id), URL: $($page.webUrl)) successfully."
        return $page
    } 
    catch {
        $ex = $_.Exception
        $errorResponse = $ex.Response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($errorResponse)
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        Write-Host "Response content:`n$responseBody" -f Red
        Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
        write-host
        break
    }
}


Function ModifyPage([object] $page) {
    # add your logic here to modify the page,
    # e.g. add a section, add a column, modify metadatas etc.
    $page.Name = "sample name"
    return $page
}