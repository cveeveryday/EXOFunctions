.\variables.ps1


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

function Get-Messages {
  param ( 
    [Parameter(Mandatory = $true)] [String]$accessToken,
    [Parameter(Mandatory = $true)] [String]$emailAddress,
    [Parameter(Mandatory = $false)][Int] $limit = 10,
    [Parameter(Mandatory = $false)][Int] $skip = 0,
    [Parameter(Mandatory = $false)][Int] $folder ,
    [Parameter(Mandatory = $false)][Int] $isRead  = $false,
    [Parameter(Mandatory = $false)][Int] $unRead = $false
  )
If (Test-IsEmailValid $emailAddress) {
  $messages = @()
  $url = "https://graph.microsoft.com/v1.0/users/" + $emailAddress + "/messages?$top="  + $limit + "&$skip="
  $url
  $headers = @{
    'Authorization' = "Bearer " + $accessToken
    'Content-Type' = 'application/json'
  }
  $params = @{
    '$top' = $limit
    '$skip' = $skip
  }
  if ($folder) { $params.Add('FolderId', $folder) }
  if ($isRead -eq $true) { $params.Add('IsRead', 'True') }
  elseif ($unRead -eq $true) { $params.Add('IsRead', 'False') }
  
  try {
    $response = Invoke-WebRequest -Uri $url -Method Get -Headers $headers -Body $params -UseBasicParsing
    $messages += ($response.Content | ConvertFrom-Json).value
    $messages
  } catch {
    Write-Host "Error getting messages: $($error[0])"
  }

}else {
    Write-Error  "Invalid email address"
    Break
}

}

function Test-IsEmailValid {
  param ( 
    [Parameter(Mandatory = $true)] [String]$emailAddress
  )

Return $emailAddress -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
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
 If ($folderName)
 {
  $uri = "https://graph.microsoft.com/v1.0/users/$emailAddress/mailFolders/?filter=displayname eq '" + $folderName + "'"
  $folders = Invoke-RestMethod -Uri $uri  -Method Get  -Headers $authHeader  -UseBasicParsing

}
 else{
  $folders = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/users/$emailAddress/mailFolders/?" -Method Get -Headers $authHeader -UseBasicParsing

}
$folders
}




$token = Get-GraphToken -appID $appID -clientSecret $clientSecret -tenantID $tenantID
$folders = Get-MailFolder -accessToken $token -emailAddress "PattiF@itissimple.ca" -folderName "Inbox"
$folders.value


