<#
.SYNOPSIS
  Send an email using the Google Cloud API for Gmail

.DESCRIPTION
  This script sends an email from a Gmail account to one or more recipients using the REST API, so avoiding use of
  ports other than 443, which may solve connectivity issues on heavily firewalled networks.

.PARAMETER To
  One or more recipient addresses.

.PARAMETER Cc
  Zero or more copy recipient addresses.

.PARAMETER Bcc
  Zero or more blind copy recipient addresses.

.PARAMETER Subject
  Mandatory message subject.

.PARAMETER Body
  Mandatory body plaintext. Can be taken from the pipeline.

.PARAMETER Attachments
  Zero or more files to be attached to the message. The total size of all attachments must not exceed 15MiB

.PARAMETER Config
  Configuration file holding the Google Cloud client ID, secret key, refresh token, current access code, plus the
  'From' address to be used for the message. This defaults to 'Send-WebGmail-creds.json' in the home directory of the invoking account, and should not be required unless testing.

.PARAMETER Setup
  Used without other options to create the credentials file needed to use the script to send emails.

.INPUTS
  The body text can be passed on the pipeline.
  
.OUTPUTS
  There is no output from this script.

.EXAMPLE
  .\Send-WebGmail.ps1 -To "someone@anydomain.com" -Subject "Test" -Body "A test message"

.NOTES
  This script requires the AE.Net.Mail package to be installed from the Nuget repository. It also requires a valid client
  ID, secret and refresh token for a Google Cloud account. This is the procedure at time of writing:
  
  1. Create a new project on https://console.developers.google.com. In the Library section, select Gmail API and click the 'Manage' button.
    In the Clients section click 'Create credentials' to create a new OAuth2 client ID.
  2. Select credential type 'Web Application'. Include 'https://developers.google.com/oauthplayground' as an authorized Redirect URL. 
    Save the client ID and Secret. 
  3. In the OAuth consent screen section click 'Audience' and set the Publishing Status to Production. This is required to avoid your
    refresh token becoming invalid after a week. It will warn you that your app requires verification, but as long as you don't create
    more than 100 clients, all that happens is that when authenticating, you are warned that the app is unverified.
  4. Go to 'https://developers.google.com/oauthplayground'. Click the cog wheel and tick 'Use your own OAuth credentials', then fill in the ID 
    and Secret from step 2.
  5. Enable the APIs you are interested in accessing in the Scopes box on the left. The minimum requirement is Gmail API v1, 
    item 'https://www.googleapis.com/auth/gmail.send'. Click 'Authorize APIs' button. 
    You then have to authenticate as the user who will be the mail sender.
  6. This opens the 'Exchange Authorization code for tokens' dialogue with the Authorization code pre-filled. 
    Click the blue button to get the refresh and access tokens. 
  7. Run this script with parameters -Setup to create the credentials file. You will be prompted for the required information.
  
#>

[CmdletBinding(DefaultParameterSetName = "Send")]
Param(
    [Parameter(Mandatory, ParameterSetName = "Send")] [string[]] $To,
    [Parameter(ParameterSetName = "Send")] [string[]] $Cc,
    [Parameter(ParameterSetName = "Send")] [string[]] $Bcc,
    [Parameter(Mandatory, ParameterSetName = "Send")] [string] $Subject,
    [Parameter(Mandatory, ValueFromPipeline = $true, ParameterSetName = "Send")] [string] $Body,
    [Parameter(ParameterSetName = "Send")] [string[]] $Attachments,
    [Parameter(ParameterSetName = "Setup")] [switch] $Setup,
    [Parameter(ParameterSetName = "Send")][Parameter(ParameterSetName = "Setup")] [string] $Config = "$HOME\Send-WebGmail-creds.json"
)

Function CheckEmail {
    Param(
        [Parameter(Position = 0)] [string] $type,
        [Parameter(Position = 1)] [string] $address
    )
    if ($address -notmatch "^(?(?=^(?:([a-zA-Z0-9_!#$%&'+-/=?^{|}~]+|[a-zA-Z0-9_!#$%&'*+\-\/=?^{|}~].[a-zA-Z0-9_!#$%&'+-/=?^{|}~][\.a-zA-Z0-9_!#$%&'*+\-\/=?^{|}~]))@[a-zA-Z0-9.-]{1,63}$)[a-zA-Z0-9_.!#$%&'*+-/=?^`{|}~]{1,63}@[a-zA-Z0-9-]+(?:.[a-zA-Z0-9-]{2,})+)$") {
        Write-Error "'$type' address '$address' is not a valid email address"
    }
}

Function Convert-Base64Url([string]$MsgIn) {
    $InputBytes = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($MsgIn))
 
    # "Url-Safe" base64 encoding
 
    $InputBytes = $InputBytes.Replace('+', '-').Replace('/', '_').Replace("=", "")
    return $InputBytes
}

function ConvertFrom-SecureStringToPlainText ([System.Security.SecureString]$SecureString) {
    [System.Runtime.InteropServices.Marshal]::PtrToStringAuto( 
        [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    )      
}

# Function from https://stackoverflow.com/questions/11698525/powershell-possible-to-determine-a-files-mime-type

function Get-MimeType() {
    param($extension = $null);
    
    $mimeType = 'application/octet-stream'
    try {
        if ( $null -ne $extension ) {
            $drive = Get-PSDrive HKCR -ErrorAction SilentlyContinue;
            if ( $null -eq $drive ) {
                $drive = New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT
            }
            $mimeType = (Get-ItemProperty HKCR:$extension)."Content Type";
        }
    } catch {
    }
    $mimeType;
}

$ProgressPreference = "SilentlyContinue"
$ErrorActionPreference = "Stop"
$ErrorView = "CategoryView"

# Ensure this is at least Powershell 5

if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Error "This script requires Powershell 5 or higher"
}

if ($PSCmdlet.ParameterSetName -eq "Setup") {
    if (Test-Path $Config) {
        Write-Error "The credentials file '$Config' already exists"
    }
    $creds = @{
        client     = Read-Host -Prompt "Paste the Client ID here";
        secret     = Read-Host -Prompt "Paste the Secret here";
        refresh    = Read-Host -Prompt "Paste the Refresh Key here";
        access     = "";
        issued     = $((Get-Date).AddDays(-1));
        expires_in = 0;
        from       = Read-Host -Prompt "Enter the 'From' address for the emails";
    }
    CheckEmail $creds.from
    if ($creds.from -notlike "*@gmail.com" -and $creds.from -notlike "*@googlemail.com") {
        Write-Error "The 'From' address must be a Gmail address"
    }
    $creds | ConvertTo-Json -Compress | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Set-Content -Path $Config -NoNewLine
    Write-Host "Credentials file '$Config' created"
    exit 0
}

Add-Type -AssemblyName System.IO
Add-Type -AssemblyName System.Text.Encoding

# This script uses AE.Net.Mail to create the message. See https://github.com/andyedinborough/aenetmail
# We need to find the DLL. This method is suggested at https://stackoverflow.com/questions/75536748/powershell-locate-and-load-package-dll

try {
    $ae = Get-Package AE.Net.Mail
}
catch {
    Write-Error "The package AE.Net.Mail is not installed or cannot be found"
}
$aepath = (Get-ChildItem -Filter *.dll -Recurse (Split-Path ($ae).Source)).FullName | Sort-Object -Desc | Select-Object -First 1
Add-Type -Path $aepath
Write-Verbose "AE.Net.Mail DLL found at '$aepath'"

# Open the credentials file. If it does not already exist, set it up by prompting for the details. It is encrypted using Microsoft DPAPI

if (-not($Config)) {
    $Config = "$HOME\Send-Web-Gmail-creds.json"
}
Write-Verbose "Credentials file is '$Config'"

# Get the credentials from the credentials file. This should be encrypted, but might be plaintext JSON if built by hand

if (-Not(Test-Path $Config -PathType Leaf)) {
    Write-Error "The credentials file '$Config' does not exist. Please run this command with the '-Setup' option to set it up"
}

try { 
    $string = Get-Content $Config -Raw
    if (-not $string.StartsWith("{")) {
        $secure = ConvertTo-SecureString -String $string
        $string = ConvertFrom-SecureStringToPlainText $secure
    }
    $creds = ConvertFrom-Json -InputObject $string
}
catch {
    Write-Error "Credentials file '$Config' is not readable or is corrupt"
}

foreach ($key in "client", "secret", "refresh", "access", "issued", "expires_in", "from") {
    if ($null -eq $creds.$key) {
        Write-Error "The entry for '$key' is missing from the credentials file '$Config'"
    }
}

# If the access token expiry time has been reached, renew it using the refresh token and save the new one, encrypted.

if ($([datetime]$creds.issued).AddSeconds($creds.expires_in) -lt (Get-Date).AddSeconds(10)) {
    $RefreshTokenParams = @{
        client_id     = $creds.client;
        client_secret = $creds.secret;
        refresh_token = $creds.refresh;
        grant_type    = 'refresh_token';
    }
    try {
        $RefreshedToken = Invoke-WebRequest -Uri "https://oauth2.googleapis.com/token" -Method POST -Body $refreshTokenParams -UseBasicParsing | ConvertFrom-Json
        Write-Verbose "Access token renewed"
    }
    catch {
        Write-Error "Cannot renew access token. Refresh token is invalid or expired"
    }
    $creds.access = $RefreshedToken.access_token
    $creds.issued = Get-Date -Format o
    $creds.expires_in = $RefreshedToken.expires_in
    $creds | ConvertTo-Json -Compress | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Set-Content -Path $Config -NoNewLine
}
$AccessToken = $creds.access

# Create a new message object

$Msg = New-Object AE.Net.Mail.MailMessage

# Set the addresses for the email, validating each one

$To.ForEach( {
        CheckEmail 'To' $_
        $Msg.To.Add($_)
    } )
$Cc.ForEach( {
        CheckEmail 'To' $_
        $Msg.Cc.Add($_)
    } )
$Bcc.ForEach( {
        CheckEmail 'To' $_
        $Msg.Bcc.Add($_)
    } )
CheckEmail 'From' $creds.from
Write-Verbose "'From' address is '$($creds.from)'"
$Msg.From = $creds.from
$Msg.ReplyTo.Add($creds.from) # Important so email doesn't bounce 
$Msg.Subject = $Subject
$Msg.Body = $Body
$AttachSize = 0
$Attachments.ForEach( {
        if (-not(Test-Path $_ -PathType Leaf)) {
            Write-Error "Attachment '$_' does not exist or is a directory"
        }
        $item = Get-Item $_
        if ($item.length -eq 0) {
            Write-Warning "File '$_' is zero-length and will not be attached"
            continue
        }
        $AttachSize += $item.length
        if ($AttachSize -gt 15000000) {
            Write-Error "Maximum total size of attachments (15MiB) would be exceeded. Mail not sent"
        }
        $mimetype = Get-MimeType($item.extension)
        try {
            $content = [System.IO.File]::ReadAllBytes($_)
        }
        catch {
            Write-Error "Attachment '$item.versioninfo.filename' is not readable"
        }
        $attachment = New-Object AE.Net.Mail.Attachment($content, $mimetype, (Split-Path $_ -Leaf), $true)
        $Msg.Attachments.Add($attachment) 
    } ) 

$MsgSW = New-Object System.IO.StringWriter
$Msg.Save($MsgSW)

$EncodedEmail = Convert-Base64Url $MsgSW

# Found this hint here: https://github.com/thinkAmi/PowerShell_misc/blob/master/gmail_api/gmail_sender.ps1

$Content = @{ "raw" = $EncodedEmail; } | ConvertTo-Json
Invoke-RestMethod -Uri "https://www.googleapis.com/gmail/v1/users/me/messages/send?access_token=$AccessToken" -Method POST -Body $Content -ContentType "Application/Json" | Out-Null
Write-Host "Mail sent"