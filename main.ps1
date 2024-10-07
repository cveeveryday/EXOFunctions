.\variables.ps1
########################################################################################################################
###########################################FUNCTIONS PART###############################################################
########################################################################################################################

function Get-GraphToken {
  param (
    [Parameter(Mandatory = $true)]
    [String]  $appID,
    [Parameter(Mandatory = $true)]
    [String]   $clientSecret,
    [Parameter(Mandatory = $true)]
    [String]    $tenantID
  )

#Prepare token request
$url = 'https://login.microsoftonline.com/' + $tenantId + '/oauth2/v2.0/token'

$body = @{
    grant_type = "client_credentials"
    client_id = $appID
    client_secret = $clientSecret
    scope = "https://graph.microsoft.com/.default"
}

#Obtain the token
$tokenRequest = Invoke-WebRequest -Method Post -Uri $url -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing -ErrorAction Stop
($tokenRequest.Content | ConvertFrom-Json).access_token
}

function Get-MailMessages {
  param ( 
    [Parameter(Mandatory = $true)] [String]$accessToken,
    [Parameter(Mandatory = $true)] [String]$emailAddress,
    [Parameter(Mandatory = $false)][Int] $limit = 0,
    [Parameter(Mandatory = $false)][Int] $skip = 0,
    [Parameter(Mandatory = $false)][String] $folderid = $null,
    [Parameter(Mandatory = $false)][bool] $isRead  = $false,
    [ Parameter(Mandatory = $false)][string] $url  = $null
  )
If (Test-IsEmailAddressValid $emailAddress) {
  $messages = @()
  $params = @{}
  if (!$url) {
    $url = "https://graph.microsoft.com/v1.0/users/" + $emailAddress + "/mailFolders/" + $folderid +  "/messages"
    $params = @{
      'skip' = $skip
      'filter' = "isRead eq " + $isread.ToString().ToLower()
      }
    if ($limit -gt  0) {
      $params['top'] = $limit
    }
  }
  $headers = @{
    'Authorization' = "Bearer " + $accessToken
    'Content-Type' = 'application/json'
  }
     
  try {
    $response = Invoke-RestMethod -Uri $url -Method Get -Headers $headers -UseBasicParsing -Body $params
    $messages += $response.value
    If ($response.'@odata.nextlink' -and ($limit -eq 0 -or $limit -gt 1000)) {
      $messages += Get-MailMessages -accessToken $accessToken -emailAddress $emailAddress -url $response.'@odata.nextlink'
    }
  } catch {
    Write-Host "Error getting messages: $($error[0])"
  }

}else {
    Write-Error  "Invalid email address"
    Break
}
$messages
}

function Set-MailMessageAsRead {
 param (
    [Parameter(Mandatory = $true)][String] $accessToken,
    [Parameter(Mandatory = $true)][string] $emailAddress,
    [Parameter(Mandatory = $true)][string] $messageId
 )
 $url = "https://graph.microsoft.com/v1.0/users/" + $emailAddress + "/messages/"  + $messageId

 $headers = @{
   'Authorization' = "Bearer  "  + $accessToken
    'Content-Type' =  'application/json'
 }
 $params = @{
   'isRead' = "true"
 }
 try {
   $response = Invoke-RestMethod -Uri $url -Method Patch  -Headers $headers  -UseBasicParsing  -Body $params
   Write-Host "Message set as read"
   } catch  {
    Write-Error "Error setting message as read: $($error[0])"
    Break
    }
}

function Test-IsEmailAddressValid {
  param ( 
    [Parameter(Mandatory = $true)] [String]$emailAddress
  )

$emailAddress -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
}

function Get-MailFolder {

  param(
    [Parameter(Mandatory = $true)] [String]$accessToken,
    [Parameter(Mandatory = $true)][String]$emailAddress,
    [Parameter(Mandatory = $false)] [String] $folderName
  )

  $authHeader = @{
    'Content-Type'='application\json'
    'Authorization'="Bearer $token"
  }
 If ($folderName -and $folderName -ne "")
  {
    $uri = "https://graph.microsoft.com/v1.0/users/$emailAddress/mailFolders/?filter=displayname eq '" + $folderName + "'"
    $folders = Invoke-RestMethod -Uri $uri  -Method Get  -Headers $authHeader  -UseBasicParsing
  }
 else{
    $uri = 'https://graph.microsoft.com/v1.0/users/$emailAddress/mailFolders'
    $folders = Invoke-RestMethod -Uri $uri -Method Get -Headers $authHeader -UseBasicParsing
  }
$folders.value
}

function Move-MailMessage {
  param(
    [Parameter(Mandatory = $true)] [String]$accessToken,
    [Parameter(Mandatory = $true)][String]$emailAddress,
    [Parameter(Mandatory = $true)][String]$messageId,
    [Parameter(Mandatory = $true)] [String] $folderName
    )
    $authHeader = @{
      'Content-Type'='application\json'
      'Authorization'="Bearer $token"
    }
    $uri = "https://graph.microsoft.com/v1.0/users/$emailAddress/messages/$messageId/move/"
    $body = @{
      'destinationid' = $folderName
    }
    try {
      $response = Invoke-RestMethod -Uri $uri   -Method Post  -Headers $authHeader  -UseBasicParsing  -BodyAsJson $body
    }
    catch {
      Write-Error $_.Exception.Message
    }
}

function ConvertFrom-Html {
  param([Parameter(Mandatory=$true)][string]$inputString)
  $outputString = $inputString
  $outputString = $outputString -replace '<.*?>', ''
  $outputString = $outputString -replace '&nbsp;', ' '
  $outputString = $outputString -replace '<br>', " `n"
  $outputString = $outputString -replace '</p>', " `n"
  $outputString = $outputString -replace "&#[0-9A-F]{5};", ''
  $outputString = $outputString -replace '(?s)<img.*?>', ''
  $outputString = $outputString -replace "`u{2018}", ''
  $outputString = $outputString -replace "`u{2019}", ''
  $arrayString = $outputString -split "`n"
  $arrayString | ForEach-Object {  
                      if ($_ -match '<!--' -or $css) 
                      { $css = $true} 
                      else 
                      {$newString += $_; $newString += "`n" }; if ( $_ -match '-->' ) {$css = $false }  
                                } 
  return $newString
}


########################################################################################################################
###########################################EXECUTION PART###############################################################
########################################################################################################################
<#
$token = Get-GraphToken -appID $appID -clientSecret $clientSecret -tenantID $tenantID
$folderId = (Get-MailFolder -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com" -folderName "Inbox").id
$emailsAll = Get-MailMessages -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com"  -folderId $folderId -limit 10
foreach ($email in $emailsAll) {
  $email.subject
  $email.sender.emailAddress.address
  $email.toRecipients.emailAddress.address
  $email.receivedDateTime
  $(ConvertFrom-Html -inputString $email.body.content)

}

$commonMailFolders = @("inbox", "sent items", "drafts","deleted items")
#$folders = Get-MailFolder -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com" -folderName "Inbox"
#$folders = Get-MailFolder -accessToken $token -emailAddress "PattiF@zpzbx.onmicrosoft.com" -folderName ""
$folders = @()
foreach ($folder in $commonMailFolders)
{
  $folders += Get-MailFolder -accessToken $token  -emailAddress "PattiF@zpzbx.onmicrosoft.com"  -folderName $folder
}
 
$message = "About to list messages in this folder " + $($folders.Value.Displayname) -join ","
Sleep -Seconds 4

Write-Host $message
$messages = @()
$folders.value | ForEach-Object {
  $messages += Get-MailMessages -accessToken $token  -emailAddress "PattiF@zpzbx.onmicrosoft.com"  -folderId $_.id -limit 1000 -skip 0 -isRead $true
  $messages += Get-MailMessages -accessToken $token  -emailAddress "PattiF@zpzbx.onmicrosoft.com"  -folderId $_.id -limit 1000 -skip 0 -isRead $false
}

$folders.value[0].id
$messages
$messages.count
#>