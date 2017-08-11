# MicrosoftgraphAPI-ReadAllUserMail

For Email Set up web app  azure
1) Create app through https://portal.azure.com  [Share the tennant ID]
2) Click on Azure Active Directory -> App Registration
3) Click on + New Application of 
	a) Name - WarpBubble
	b) Type Web app /Api
	c) Reply Url -  http://localhost:3000/callback  
4) Copy and provide the Application ID  []
5) Copy and provide the Object ID  [Client Id]
6) Click on Required Permissions -> click on Add
7) Select permission Microsoft graph -> select all application / delegated permission
8) Select permission Skype for business online -> select all application / delegated permission
9) Select permission office 365 exchange online -> select all application / delegated permission
10) Click on Grant Permission [With Admin Consents]
11) Click on Keys undeer Api Access 
12) Create new Key 
	a) Description - WarpBubble_secret 
	b) Expire 2 yrs  [Make a note of expire date]
	c) click save - note the value of key  [Client secret]


13) Share the usercredential username and password will be used to send notification email [warpbubble@tesla.com / P@ssW0rd]

-- Set up certificate Authentication (to send mail under application permission)
1) Create self signed certificate using following command 
makecert -r -pe -n "CN=warpbubble_notification_certificate" -b 07/09/2017 -e 07/09/2018 -ss warpbubble -len 2048
[for ref
 -r create self signed certificate
-pe mark private key exportable
-n name
-b start date
-ss name of subject
]
Get Base64 encoded certificate value and thumbprint from your self-issued certificate by running the following PS code
2) Open Certificate Snap-in and export your certificate to under warpbubble <warpbubble.cer>
3) Export your certificate with private key to <warpbubble.pfx> and enter password  [Make a note of password and share both certificate] 
4) Run the power shell script

$certPath = Read-Host “Enter certificate path (.cer)”
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
$cert.Import($certPath)
$rawCert = $cert.GetRawCertData()
$base64Cert = [System.Convert]::ToBase64String($rawCert)
$rawCertHash = $cert.GetCertHash()
$base64CertHash = [System.Convert]::ToBase64String($rawCertHash)
$KeyId = [System.Guid]::NewGuid().ToString()
$keyCredentials =
‘“keyCredentials”: [
{
“customKeyIdentifier”: “‘+ $base64CertHash + ‘“,
“keyId”: “‘ + $KeyId + ‘“,
“type”: “AsymmetricX509Cert”,
“usage”: “Verify”,
“value”: “‘ + $base64Cert + ‘“
}
],’
$keyCredentials
Write-Host “Certificate Thumbprint:” $cert.Thumbprint

5) make a note of keyCredentials / and certificate thumbprint
6) Open the app created in azure in sstep (3 - web app/api)
7) Click on manifest copy the KeyCredential under keyCredentials : [{}]
8) Set  "oauth2AllowImplicitFlow": true alson in manifest 
9) Save the changes.

Note -- make a note of cetificate expire and client secret expiry, so that can be replaced on time.

-Set up app for akype notification
1) Sign in to your Azure Management Portal at https://portal.azure.com
2) Select Active Directory -> App registrations -> New application registration
   a) Name: Warp_Skype_Notification 
   b) Application type: Native
   c)Redirect URI: http://localhost:3000/callback   (anything will work)
4) Copy and provide the Application ID  []
5) Copy and provide the Object ID  [Client Id]
6) Click All settings -> Required Permissions
7) Click Add
	a) Select an API -> Skype for Business Online (Microsoft.Lync) -> Select
	b) Select permissions -> Select all Delegated Permissions
	c) save
8) Click on Grant Permission [With Admin Consents]
9) Under setting click on Keys
10) Create new key with 
	a) Description : - Skype_secret
	b) Expire :- 2 yrs note the time
	c) Save [copy the value / share the value as client secret]






