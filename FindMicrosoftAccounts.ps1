#$emailAddresses = 'tinafey@contoso.one','username@domain.com','ronswanson@contoso.one'
$emailAddresses = Get-EXORecipient -RecipientTypeDetails UserMailbox -PropertySets Minimum | select -ExpandProperty Emailaddresses | where {$_ -match "SMTP:"} | foreach {$_ -replace '^smtp:'} 

$ApplicationID  = 'bd53bb89-0cc1-4eb3-90b7-ba008b1f2a2c'
$scope          = 'user.read'

$results = [System.Collections.ArrayList]::new()
$emailAddresses | foreach -Begin { $i = 1} -Process {
    Write-Progress -Activity 'Checking Email Addresses for Microsoft Accounts' -CurrentOperation ("Checking {0} - {1}/{2}" -f $_,$i,$emailAddresses.Count) -PercentComplete ($i/($emailAddresses.count)*100)
    $UserURL  = 'https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize?client_id={0}&scope={1}&response_type=code&state=23424&login_hint={2}' -f $ApplicationID,$scope,$_
    $response = Invoke-WebRequest -Uri $UserURL 

    $results.Add( 
        [psCustomobject]@{
            EmailAddress = $_
            HasMSAccount = $response -match '"HasPassword":1' 
            Result       = $response.StatusCode
        }
    ) | Out-Null
    $i++
} 
$results
