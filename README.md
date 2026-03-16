# Send-WebGmail #

Send mail using the Google Cloud API.

## Background ##

I needed to be able to send emails from a Powershell script on a system whose network was not under my control and did not allow connections to ports other than 80 and 443. I considered using MailerSend or MailJet, but the sender was going to be a Gmail address, so messages would likely be bounced due to SPF/DMARC issues. If I used the Google Cloud API for Gmail, this would not be a problem. 

## Sample usage ##

```` Send-WebGmail -To "someone@domain.com" -Subject "Test" -Body "A test message" -Attachment "myfile.pdf" ````

## Installation ##

Copy Send-WebGmail.ps1 to a folder on your PATH. You will also need to install the package [AE.Net.Mail](https://www.nuget.org/packages/AE.Net.Mail/) from Nuget, as this is used to construct the messages themselves. The script will do its best to find the required DLL once that package is installed.

Google Cloud uses OAuth2 to authorise use of its APIs. You will need a Google account, which you use to register a project, in which you can specify the sending account (not necessarily the same).

## Obtaining credentials ##

1. Create a new project on https://console.developers.google.com. Select 'Library' on the navigation pane, then select 'Gmail API' and click the 'Manage' button.
2. Return to the navigation pane, select 'Clients' and click 'Create credentials' to create a new OAuth2 client ID. Select credential type 'Web Application'. Include 'https://developers.google.com/oauthplayground' as an authorized Redirect URL. Save the client ID and secret. 
3. Return to the navigation pane, select 'OAuth consent screen', click 'Audience' and set the Publishing Status to Production. This is required to avoid your refresh token becoming invalid after a week. It will warn you that your app requires verification, but as long as you don't create more than 100 clients, all that happens is that when authenticating, you are warned that the app is unverified.
4. Go to 'https://developers.google.com/oauthplayground'. Click the cog wheel and tick 'Use your own OAuth credentials', then fill in the ID and Secret from step 2.
5. Enable the APIs you are interested in accessing in the 'Scopes' box on the left. The minimum requirement is 'Gmail API v1', item 'https://www.googleapis.com/auth/gmail.send'. Click 'Authorize APIs' button. You then have to authenticate as the user who will be the mail sender.
6. This opens the 'Exchange Authorization code for tokens' dialogue with the Authorization code pre-filled. Click the blue button to get the refresh and access tokens. Copy them for the next step.
7. Run this script with the -Setup parameter to create the credentials file. You will be prompted for the required information.

The credentials are saved as 'Send-WebGmail-creds.json' in your home directory. They are encrypted using Microsoft DPAPI library, so can only be decrypted on the same system using the same account.
