$first_name = 'First Name'
$last_name = 'Last Name'
$mail = 'Office 365 account of person sending email'
$college_name = 'College or District Name'
$api_key = 'Your API Key'
$api_secret = 'Your API Secret'
$helpdesk_Email = 'Helpdesk Email'
$user = "Email account username"
$PWord = ConvertTo-SecureString -String "Email account password" -AsPlainText -Force
###Do not edit anything below this line
$Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $PWord
function Generate-JWT (
    [Parameter(Mandatory = $True)]
    [ValidateSet("HS256", "HS384", "HS512")]
    $Algorithm = $null,
    $type = $null,
    [Parameter(Mandatory = $True)]
    [string]$Issuer = $null,
    [int]$ValidforSeconds = $null,
    [Parameter(Mandatory = $True)]
    $SecretKey = $null
    ){

    $exp = [int][double]::parse((Get-Date -Date $((Get-Date).addseconds($ValidforSeconds).ToUniversalTime()) -UFormat %s)) # Grab Unix Epoch Timestamp and add desired expiration.

    [hashtable]$header = @{alg = $Algorithm; typ = $type}
    [hashtable]$payload = @{iss = $Issuer; exp = $exp}

    $headerjson = $header | ConvertTo-Json -Compress
    $payloadjson = $payload | ConvertTo-Json -Compress
    
    $headerjsonbase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($headerjson)).Split('=')[0].Replace('+', '-').Replace('/', '_')
    $payloadjsonbase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payloadjson)).Split('=')[0].Replace('+', '-').Replace('/', '_')

    $ToBeSigned = $headerjsonbase64 + "." + $payloadjsonbase64

    $SigningAlgorithm = switch ($Algorithm) {
        "HS256" {New-Object System.Security.Cryptography.HMACSHA256}
        "HS384" {New-Object System.Security.Cryptography.HMACSHA384}
        "HS512" {New-Object System.Security.Cryptography.HMACSHA512}
    }

    $SigningAlgorithm.Key = [System.Text.Encoding]::UTF8.GetBytes($SecretKey)
    $Signature = [Convert]::ToBase64String($SigningAlgorithm.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($ToBeSigned))).Split('=')[0].Replace('+', '-').Replace('/', '_')
    
    $token = "$headerjsonbase64.$payloadjsonbase64.$Signature"
    $token
}
$A = Generate-JWT -Algorithm 'HS256' -type 'JWT' -Issuer $api_key -SecretKey $api_secret -ValidforSeconds 540
$Bkey = 'Bearer ' +  $A    
$headers=@{}
$headers.Add("authorization", $Bkey)
$CURRENTDATE=GET-DATE -Hour 0 -Minute 0 -Second 0
$MonthAgo = $CURRENTDATE.AddMonths(-1)
$FIRSTDAYOFMONTH=GET-DATE $MonthAgo -Day 1
$pastdate = $FIRSTDAYOFMONTH.ToString('yyyy-MM-dd')
$LASTDAYOFMONTH=GET-DATE $FIRSTDAYOFMONTH.AddMonths(1).AddSeconds(-1)
$todaydate =$LASTDAYOFMONTH.ToString('yyyy-MM-dd')
$Webinar_particpant_count = 0
$EmailMonthAgo = ($CURRENTDATE.AddMonths(-1)).ToString('MMMM')
#Start getting meeting ids
$APIURL = 'https://api.zoom.us/v2/metrics/meetings?page_size=300&to=' + $todaydate + '&from=' + $pastdate + '&type=past'
$response = Invoke-RestMethod -Uri $APIURL -Method GET -Headers $headers
$meetings += $response.meetings.id
#if more than one page exists go into while to add them to the $meetings variable
while ($response.next_page_token) {
 $token =  $response.next_page_token
 $APIURL = 'https://api.zoom.us/v2/metrics/meetings?page_size=300&to=' + $todaydate + '&from=' + $pastdate + '&type=past&next_page_token=' + $token
 $response = Invoke-RestMethod -Uri $APIURL -Method GET -Headers $headers  
 $meetings += $response.meetings.id
}
#get meeting participants from each meeting and count them
foreach ($meeting in $meetings) {
$meetingurl = 'https://api.zoom.us/v2/metrics/meetings/' + $meeting + '/participants?page_size=30&type=past'
$meetingresponse = Invoke-RestMethod -Uri $meetingurl -Method GET -Headers $headers
$particpant_count += $meetingresponse.total_records
} 
#get Webinar count
$WebinarURL = 'https://api.zoom.us/v2/metrics/webinars?page_size=300&to=' + $todaydate + '&from=' + $pastdate + '&type=past'
$Webinarresponse = Invoke-RestMethod -Uri $WebinarURL -Method GET -Headers $headers
$Webinars = $Webinarresponse.total_records
#if we actually had a webinar then count the participants
if ($Webinars -ige 1){
#checks for Webinar participants
foreach ($Webinar in $Webinars) {
    $Webinarurl = 'https://api.zoom.us/v2/metrics/webinars/' + $Webinar + '/participants?page_size=30&type=past'
    $Webinarresponse = Invoke-RestMethod -Uri $Webinarurl -Method GET -Headers $headers
    $Webinar_particpant_count += $Webinarresponse.total_records
}
}


$body = 
"
<head>
<style>
table {
  font-family: arial, sans-serif;
  border-collapse: collapse;
  width: 100%;
}

td, th {
  border: 1px solid #dddddd;
  text-align: left;
  padding: 8px;
}

tr:nth-child(even) {
  background-color: #dddddd;
}
</style>
</head>
<body>
<h2>Zoom Subaccount Statistics - $($college_name)</h2>
<table>
  <tr>
    <th>Type</th>
    <th>Description</th>
  </tr>
 <tr>
    <td>College / District</td>
    <td>$($college_name)</td>
  </tr>
   <tr>
    <td>Month Reporting For</td>
    <td>$($EmailMonthAgo)</td>
  </tr>
   <tr>
    <td>First Name</td>
    <td>$($first_name)</td>
  </tr>
   <tr>
    <td>Last Name</td>
    <td>$($last_name)</td>
  </tr>
   <tr>
    <td>Email</td>
    <td>$($mail)</td>
  </tr>
  <tr>
    <td>Zoom Meetings</td>
    <td>$($meetings.count)</td>
  </tr>
  <tr>
    <td>Zoom Meeting Participants</td>
    <td>$($particpant_count)</td>
  </tr>
  <tr>
    <td>Zoom Webinar</td>
    <td>$($Webinars)</td>
  </tr>
  <tr>
    <td>Zoom Webinar Particpants</td>
    <td>$($Webinar_particpant_count)</td>
  </tr>
</table>
<p>Hello CCCTechconnect team this report has been autogenerated per your request. <br>If you need additional information or need the current information changed please feel
free to submit a ticket to $($helpdesk_Email)</p>
</body>"

Send-MailMessage -To "support@ccctechconnect.org" -cc $mail -from $user -Subject "Zoom Subaccount Statistics - $($college_name) - $($EmailMonthAgo)" -Priority Normal -Body $body -BodyAsHtml -smtpserver smtp.office365.com -usessl -Credential $Creds -Port 587
Remove-Variable * -ErrorAction SilentlyContinue