<#
.DESCRIPTION
This script is to get statistics of page load times in SharePoint Online.   

The script opens a specified URL, then collect the SPRequestDuration, SPIisLatency and SPRequestGuid
response header, and puts them into a CSV file.
    
.PARAMETER $SPOPageUrl
Required. The URL of the page to be opened. When using the script to collect data to report SPO performance
issues, please use the default page of an OOB Team Site.

.PARAMETER $UseTLS12
Optional. Specifies if the script should force TLS 1.2 for the web service calls. If omitted, the script is
not using TLS1.2.

.PARAMETER $Username
Required. The name of the user to be used to access the page.

.PARAMETER $Integrated
Optional. Decides if the script should use integrated authentication. If omitted, it will automatically
assume 'False' value.

.PARAMETER $TokenOutputFormat
Optional. Displays the received STS Token in the given format. Allowed parameters are: XML, JSON, KEYVALUE,
NAMEVALUE. If omitted, the tokes is dumped in raw format

.PARAMETER $Iteration
Optional. The number of requests to issue. The script waits one minute between every iteration. If omitted,
the script tries only one time.

.PARAMETER Sleep
Optional. The number of seconds the script should wait between iterations. If omitted, the script waits
60 seconds between retries

.PARAMETER $Outputfile
Optional. The full path of the CSV file where results are to be saved.

.PARAMETER $SilentMode
Optional. Switch parameter to specify if the script should give feedback on every request, as well as show a
progress bar



.EXAMPLE
.\SPOPageLoadStatistics.ps1 -SPOPageUrl "https://contoso.sharepoint.com/sites/test/Home.aspx" -UserName "admin@contoso.onmicrosoft.com" -Iteration 10  -Sleep 5 -Outputfile "C:\Temp\SPOStats.csv"

With these parameters the script opens the https://contoso.sharepoint.com/sites/test/Home.aspx page with the 
admin@contoso.onmicrosoft.com account 10 times, waiting about 5 seconds between each request, put the output
onto the screen and finally put the output into the C:\Temp\SPOStats.csv file.


.SYNOPSIS
This script is to get statistics of page load times in SharePoint Online.
The script opens a specified URL, then collect the SPRequestDuration, SPIisLatency and SPRequestGuid
response header, and puts them into a CSV file.

#>


[cmdletbinding()]
param
(
    [parameter(Mandatory=$false)][string]$SPOPageUrl,
    [parameter(Mandatory=$false)][switch]$UseTLS12,
    [parameter(Mandatory=$true)][string]$Username,
    [Parameter(Mandatory=$false)][switch]$Integrated = $false,
    [Parameter(Mandatory=$false)][ValidateSet('XML','JSON','KEYVALUE','NAMEVALUE','RAW')][string]$TokenOutputFormat='RAW',
    [parameter(Mandatory=$false)][int64]$Iteration=1,
    [parameter(Mandatory=$false)][int64]$Sleep=60,
    [parameter(mandatory=$false)][string]$Outputfile,
    [parameter(Mandatory=$false)][switch]$SilentMode
)

# Just loading the required DLLs
cls
Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction SilentlyContinue

# A simple function to log if necessary
Function ToLog($Info)
{
    If(!$SilentMode)
    {
        If([string]::IsNullOrEmpty($Info))
        {
            $Info = ""
        }

        # This is where we can implement logging if necessary
    }
}

# Checking if the output file variable is valid, and if the file exists
If($Outputfile)
{
    If(Test-Path $Outputfile -PathType Container)
    {
        ToLog 'Invalid output file name.'
        Write-Host 'The output file specified (' -NoNewline -ForegroundColor Red
        Write-Host $Outputfile -NoNewline
        Write-Host ') is a container.' -ForegroundColor Red
        Write-Host 'Please correct the variable and try again.' -ForegroundColor Red
        Write-Host 'The script halted.' -ForegroundColor Yellow
        Break 
    }


    If(Test-Path $Outputfile)
    {
        ToLog 'Output file already exist.'
        Write-Host 'The output file specified (' -NoNewline -ForegroundColor Red
        Write-Host $Outputfile -NoNewline
        Write-Host ') already exists.' -ForegroundColor Red
        Write-Host 'Do you want to append the new results to the file?' -ForegroundColor Red
	    [regex]$Yes = 'y|Y'
	    $Answer = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        [bool]$Continue = ($Yes.Match($Answer.Character)).Success 

        If($Answer -ieq 'n')
        {
            Write-Host 'Please correct the variable and try again.' -ForegroundColor Red
            Write-Host 'The script halted.' -ForegroundColor Yellow
            Break 
        }
    }
    Else
    {
        # Prepare the output file
        $FileHeader = "Url`tTime`tSPRequestDuration`tSPIisLatency`tRequestTotal`tSPHealthScore`tRoundTripTime`tSPRequestGuid`tMSEdgeRef"
        Try
        {
            $FileHeader | Out-File $Outputfile
        }
        Catch
        {
            Write-Host 'Could not write the output file.' -ForegroundColor Red
            Write-Host 'The script halted.' -ForegroundColor Yellow
            Break
        }
    }
}


######################################################
# Preparing the Authenticaition Cookie for SPO       #
# We need this in case ADFS is used and the redirect #
# is not picked up by some reason.                   #
######################################################

# Checking if the URL provided is formatted good.
try
{
    If(![uri]::IsWellFormedUriString($SPOPageUrl, [UriKind]::Absolute))
    {
        # Not good, so throwing an error.
        Throw "Parameter 'url' is not a valid URI."
    }
    Else
    {
        # Looks good, so we're going to use this to get the Tenant name
        $Uri = [uri]::New($SPOPageUrl)
        $Tenant = $Uri.Authority
    }

    # Get the Tenant name from the URL
    If($Tenant.EndsWith("sharepoint.com", [System.StringComparison]::OrdinalIgnoreCase))
    {
        $MSODomain = "sharepoint.com"
    }
    Else
    {
        $MSODomain = $Tenant
    }

    # Check if we need to use integrated authentication
    If($Integrated.ToBool())
    {
        # Yes, so we do not use the UserName/Password provided
        [System.Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices") | out-null
        [System.Reflection.Assembly]::LoadWithPartialName("System.DirectoryServices.AccountManagement") | out-null
        $Username = [System.DirectoryServices.AccountManagement.UserPrincipal]::Current.UserPrincipalName
    }
    Else
    {
        # No, so we need to get the UserName and Password
        
        # No, so we need to get the UserName and Password
        $Credential = Get-Credential -UserName $Username -Message "Enter credentials to connect SPO"
        $Username = $Credential.UserName
        $Password = $Credential.GetNetworkCredential().Password
         
    }

    # Prepare the variables for the STS call
    $ContextInfoUrl = $SPOPageUrl.TrimEnd('/') + "/_api/contextinfo"
    $GetRealmUrl = "https://login.microsoftonline.com/GetUserRealm.srf"
    $Realm = "urn:federation:MicrosoftOnline"
    $msoStsAuthUrl = "https://login.microsoftonline.com/rst2.srf"
    $idcrlEndpoint = "https://$Tenant/_vti_bin/idcrl.svc/"
    $Username = [System.Security.SecurityElement]::Escape($Username)
    $Password = [System.Security.SecurityElement]::Escape($Password)

    # Below are the various envelope formats, depending on what is required by the STS

    # In case custom STS integrated authentication needs to be used, the authentication envelope format parameters are:
    #     0: message id - unique guid
    #     1: custom STS auth url
    #     2: realm
    $CustomStsSamlIntegratedRequestFormat = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><s:Envelope xmlns:s=`"http://www.w3.org/2003/05/soap-envelope`" xmlns:a=`"http://www.w3.org/2005/08/addressing`"><s:Header><a:Action s:mustUnderstand=`"1`">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</a:Action><a:MessageID>urn:uuid:{0}</a:MessageID><a:ReplyTo><a:Address>http://www.w3.org/2005/08/addressing/anonymous</a:Address></a:ReplyTo><a:To s:mustUnderstand=`"1`">{1}</a:To></s:Header><s:Body><t:RequestSecurityToken xmlns:t=`"http://schemas.xmlsoap.org/ws/2005/02/trust`"><wsp:AppliesTo xmlns:wsp=`"http://schemas.xmlsoap.org/ws/2004/09/policy`"><wsa:EndpointReference xmlns:wsa=`"http://www.w3.org/2005/08/addressing`"><wsa:Address>{2}</wsa:Address></wsa:EndpointReference></wsp:AppliesTo><t:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</t:KeyType><t:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</t:RequestType></t:RequestSecurityToken></s:Body></s:Envelope>";


    # In case custom STS authentication with username and password needs to be used, the authentication envelope format parameters are:
    #     {0}: ADFS url, such as https://corp.sts.contoso.com/adfs/services/trust/2005/usernamemixed, its value comes from the response in GetUserRealm request.
    #     {1}: MessageId, it could be an arbitrary guid
    #     {2}: UserLogin, such as someone@contoso.com
    #     {3}: Password
    #     {4}: Created datetime in UTC, such as 2012-11-16T23:24:52Z
    #     {5}: Expires datetime in UTC, such as 2012-11-16T23:34:52Z
    #     {6}: tokenIssuerUri, such as urn:federation:MicrosoftOnline, or urn:federation:MicrosoftOnline-int
    $CustomStsSamlRequestFormat = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><s:Envelope xmlns:s=`"http://www.w3.org/2003/05/soap-envelope`" xmlns:wsse=`"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd`" xmlns:saml=`"urn:oasis:names:tc:SAML:1.0:assertion`" xmlns:wsp=`"http://schemas.xmlsoap.org/ws/2004/09/policy`" xmlns:wsu=`"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd`" xmlns:wsa=`"http://www.w3.org/2005/08/addressing`" xmlns:wssc=`"http://schemas.xmlsoap.org/ws/2005/02/sc`" xmlns:wst=`"http://schemas.xmlsoap.org/ws/2005/02/trust`"><s:Header><wsa:Action s:mustUnderstand=`"1`">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</wsa:Action><wsa:To s:mustUnderstand=`"1`">{0}</wsa:To><wsa:MessageID>{1}</wsa:MessageID><ps:AuthInfo xmlns:ps=`"http://schemas.microsoft.com/Passport/SoapServices/PPCRL`" Id=`"PPAuthInfo`"><ps:HostingApp>Managed IDCRL</ps:HostingApp><ps:BinaryVersion>6</ps:BinaryVersion><ps:UIVersion>1</ps:UIVersion><ps:Cookies></ps:Cookies><ps:RequestParams>AQAAAAIAAABsYwQAAAAxMDMz</ps:RequestParams></ps:AuthInfo><wsse:Security><wsse:UsernameToken wsu:Id=`"user`"><wsse:Username>{2}</wsse:Username><wsse:Password>{3}</wsse:Password></wsse:UsernameToken><wsu:Timestamp Id=`"Timestamp`"><wsu:Created>{4}</wsu:Created><wsu:Expires>{5}</wsu:Expires></wsu:Timestamp></wsse:Security></s:Header><s:Body><wst:RequestSecurityToken Id=`"RST0`"><wst:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</wst:RequestType><wsp:AppliesTo><wsa:EndpointReference>  <wsa:Address>{6}</wsa:Address></wsa:EndpointReference></wsp:AppliesTo><wst:KeyType>http://schemas.xmlsoap.org/ws/2005/05/identity/NoProofKey</wst:KeyType></wst:RequestSecurityToken></s:Body></s:Envelope>"

    # The mso envelope format paramters (Used for custom STS + MSO authentication)
    #     0: custom STS assertion
    #     1: mso endpoint
    $msoSamlRequestFormat = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><S:Envelope xmlns:S=`"http://www.w3.org/2003/05/soap-envelope`" xmlns:wsse=`"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd`" xmlns:wsp=`"http://schemas.xmlsoap.org/ws/2004/09/policy`" xmlns:wsu=`"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd`" xmlns:wsa=`"http://www.w3.org/2005/08/addressing`" xmlns:wst=`"http://schemas.xmlsoap.org/ws/2005/02/trust`"><S:Header><wsa:Action S:mustUnderstand=`"1`">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</wsa:Action><wsa:To S:mustUnderstand=`"1`">https://login.microsoftonline.com/rst2.srf</wsa:To><ps:AuthInfo xmlns:ps=`"http://schemas.microsoft.com/LiveID/SoapServices/v1`" Id=`"PPAuthInfo`"><ps:BinaryVersion>5</ps:BinaryVersion><ps:HostingApp>Managed IDCRL</ps:HostingApp></ps:AuthInfo><wsse:Security>{0}</wsse:Security></S:Header><S:Body><wst:RequestSecurityToken xmlns:wst=`"http://schemas.xmlsoap.org/ws/2005/02/trust`" Id=`"RST0`"><wst:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</wst:RequestType><wsp:AppliesTo><wsa:EndpointReference><wsa:Address>{1}</wsa:Address></wsa:EndpointReference></wsp:AppliesTo><wsp:PolicyReference URI=`"MBI`"></wsp:PolicyReference></wst:RequestSecurityToken></S:Body></S:Envelope>"

    # mso envelope format index info (Used for MSO-only authentication)
    #     0: mso endpoint
    #     1: username
    #     2: password
    $msoSamlRequestFormat2 = "<?xml version=`"1.0`" encoding=`"UTF-8`"?><S:Envelope xmlns:S=`"http://www.w3.org/2003/05/soap-envelope`" xmlns:wsse=`"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd`" xmlns:wsp=`"http://schemas.xmlsoap.org/ws/2004/09/policy`" xmlns:wsu=`"http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd`" xmlns:wsa=`"http://www.w3.org/2005/08/addressing`" xmlns:wst=`"http://schemas.xmlsoap.org/ws/2005/02/trust`"><S:Header><wsa:Action S:mustUnderstand=`"1`">http://schemas.xmlsoap.org/ws/2005/02/trust/RST/Issue</wsa:Action><wsa:To S:mustUnderstand=`"1`">{0}</wsa:To><ps:AuthInfo xmlns:ps=`"http://schemas.microsoft.com/LiveID/SoapServices/v1`" Id=`"PPAuthInfo`"><ps:BinaryVersion>5</ps:BinaryVersion><ps:HostingApp>Managed IDCRL</ps:HostingApp></ps:AuthInfo><wsse:Security><wsse:UsernameToken wsu:Id=`"user`"><wsse:Username>{1}</wsse:Username><wsse:Password>{2}</wsse:Password></wsse:UsernameToken></wsse:Security></S:Header><S:Body><wst:RequestSecurityToken xmlns:wst=`"http://schemas.xmlsoap.org/ws/2005/02/trust`" Id=`"RST0`"><wst:RequestType>http://schemas.xmlsoap.org/ws/2005/02/trust/Issue</wst:RequestType><wsp:AppliesTo><wsa:EndpointReference><wsa:Address>sharepoint.com</wsa:Address></wsa:EndpointReference></wsp:AppliesTo><wsp:PolicyReference URI=`"MBI`"></wsp:PolicyReference></wst:RequestSecurityToken></S:Body></S:Envelope>"

    # This is a standard function for an HTTP Post query
    Function Invoke-HttpPost($Endpoint, $Body, $Headers, $Session)
    {
        If (!$SilentMode)
        {
            ToLog
            ToLog "Invoke-HttpPost"
            ToLog "url: $Endpoint"
            ToLog "post body: $Body"
        }

        $Params = @{}
        $Params.Headers = $Headers
        $Params.Uri = $Endpoint
        $Params.Body = $Body
        $Params.Method = "POST"
        $Params.WebSession = $Session

        # Setting the TLS to 1.2 if necessary
        If ($UseTLS12)
        {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        }
        $WebResponse = Invoke-WebRequest @params -ContentType "application/soap+xml; charset=utf-8" -UseDefaultCredentials -UserAgent ([string]::Empty)
        $ResponseContent = $WebResponse.Content

        Return $ResponseContent
    }

    # Get saml Assertion value from the custom STS
    Function Get-AssertionCustomSts($CustomStsAuthUrl)
    {
        If (!$SilentMode)
        {
            ToLog
            ToLog "Get-AssertionCustomSts"
        }
        
        # Prepare the token variables
        $MessageId = [guid]::NewGuid()
        $Created = [datetime]::UtcNow.ToString("o", [System.Globalization.CultureInfo]::InvariantCulture)
        $Expires = [datetime]::UtcNow.AddMinutes(10).ToString("o", [System.Globalization.CultureInfo]::InvariantCulture)

        If($Integrated.ToBool())
        {
            If (!$SilentMode)
            {
                ToLog "Using Integrated authentication with the URL:"
            }

            # LowerCasing the URL as some STS are case sensitive
            $CustomStsAuthUrl = $CustomStsAuthUrl.ToLowerInvariant().Replace("/usernamemixed","/windowstransport")
            If (!$SilentMode)
            {
                ToLog $CustomStsAuthUrl
            }

            # Formatting the request
            $RequestSecurityToken = [string]::Format($CustomStsSamlIntegratedRequestFormat, $MessageId, $CustomStsAuthUrl, $Realm)
            If (!$SilentMode)
            {
                ToLog $RequestSecurityToken
            }
        }
        Else
        {
            If (!$SilentMode)
            {
                ToLog "Using forms authentication with the URL:"
            }
            
            # LowerCasing the URL as some STS are case sensitive
            $CustomStsAuthUrl = $CustomStsAuthUrl.ToLowerInvariant()

            # Formatting the request
            $RequestSecurityToken = [string]::Format($CustomStsSamlRequestFormat, $CustomStsAuthUrl, $MessageId, $Username, $Password, $Created, $Expires, $Realm)
            If (!$SilentMode)
            {
                ToLog $RequestSecurityToken
            }

        }

        # Calling the Custom STS' authentication URL with the token we put together above
        [xml]$STSResponse = Invoke-HttpPost $CustomStsAuthUrl $RequestSecurityToken

        Return $STSResponse.Envelope.Body.RequestSecurityTokenResponse.RequestedSecurityToken.Assertion.OuterXml
    }

    # This function is to get the Binary Security Token
    Function Get-BinarySecurityToken($CustomSTSAssertion, $msoSamlRequestFormatTemp)
    {
        If (!$SilentMode)
        {
            ToLog
            ToLog "Get-BinarySecurityToken"
        }

        If([string]::IsNullOrWhiteSpace($CustomSTSAssertion))
        {
            If (!$SilentMode)
            {
                ToLog "Using username and password for authentication"            
            }
            $msoPostEnvelope = [string]::Format($msoSamlRequestFormatTemp, $MSODomain, $Username, $Password)
        }
        Else
        {
            If (!$SilentMode)
            {
                ToLog "Using custom STS assertion for authentication"                        
            }
            $msoPostEnvelope = [string]::Format($msoSamlRequestFormatTemp, $CustomSTSAssertion, $MSODomain)
        }

        $msoContent = Invoke-HttpPost $msoStsAuthUrl $msoPostEnvelope
    
        # Get binary security token using regex instead of [xml]
        # Using regex to workaround PowerShell [xml] bug where hidden characters cause failure
        [regex]$regex = "BinarySecurityToken Id=.*>([^<]+)<"
        $match = $regex.Match($msoContent).Groups[1]

        Return $match.Value
    }

    Function Get-SPOIDCRLCookie($msoBinarySecurityToken)
    {
        If (!$SilentMode)
        {
            ToLog
            ToLog "Get-SPOIDCRLCookie"
            ToLog 
            ToLog "BinarySecurityToken: $msoBinarySecurityToken"
        }

        $binarySecurityTokenHeader = [string]::Format("BPOSIDCRL {0}", $msoBinarySecurityToken)
        $Params = @{uri=$idcrlEndpoint
                    Method="GET"
                    Headers = @{}
                   }
        $Params.Headers["Authorization"] = $binarySecurityTokenHeader
        $Params.Headers["X-IDCRL_ACCEPTED"] = "t"

        # Setting the TLS to 1.2 if necessary
        If ($UseTLS12)
        {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        }
        $WebResponse = Invoke-WebRequest @params -UserAgent ([string]::Empty)
        $cookie = $WebResponse.BaseResponse.Cookies["SPOIDCRL"]

        Return $cookie
    }

    # Retrieve the configured STS Auth Url (ADFS, PING, etc.)
    Function Get-UserRealmUrl($GetRealmUrl, $Username)
    {
        If (!$SilentMode)
        {
            ToLog
            ToLog "Get-UserRealmUrl"
            ToLog "url: $GetRealmUrl"
            ToLog "username: $Username"
        }

        $Body = "login=$Username&xml=1"
        $WebResponse = Invoke-WebRequest -Uri $GetRealmUrl -Method POST -Body $Body -UserAgent ([string]::Empty)
    
        Return ([xml]$WebResponse.Content).RealmInfo.STSAuthURL
    }

    [System.Net.ServicePointManager]::Expect100Continue = $true

    #1 Get custom STS auth url
    $CustomStsAuthUrl = Get-UserRealmUrl $GetRealmUrl $Username

    If($CustomStsAuthUrl -eq $null)
    {
        #2 Get binary security token from the MSO STS by passing the SAML <Assertion> xml
        $CustomSTSAssertion = $null
        $msoBinarySecurityToken = Get-BinarySecurityToken $CustomSTSAssertion $msoSamlRequestFormat2
    }
    Else
    {
        #2 Get SAML <Assertion> xml from custom STS
        $CustomSTSAssertion = Get-AssertionCustomSts $CustomStsAuthUrl

        #3 Get binary security token from the MSO STS by passing the SAML <Assertion> xml
        $msoBinarySecurityToken = Get-BinarySecurityToken $CustomSTSAssertion $msoSamlRequestFormat
    }

    #3/4 Get SPOIDRCL cookie from SharePoint site by passing the binary security token
    #  Save cookie and reuse with multiple requests
    $idcrl = $null
    $idcrl = Get-SPOIDCRLCookie $msoBinarySecurityToken
    
    If([string]::IsNullOrEmpty($format))
    {
        $format = [string]::Empty
    }
    Else
    {
        $format = $format.Trim().ToUpperInvariant()
    }
    
    $SPOIDCrl = $null
    $SPOIDCrl = $idcrl

    # Dump the token if required
    If(!$SilentMode)
    {
        If($TokenOutputFormat -eq 'XML')
        {
            Write-Output ([string]::Format("<SPOIDCRL>{0}</SPOIDCRL>", $idcrl.Value))
        }
        ElseIf($TokenOutputFormat -eq 'JSON')
        {
            Write-Output ([string]::Format("{{`"SPOIDCRL`":`"{0}`"}}", $idcrl.Value))
        }
        ElseIf(($TokenOutputFormat -eq 'KEYVALUE') -or ($TokenOutputFormat -eq 'NAMEVALUE'))
        {
            Write-Output ("SPOIDCRL:" + $idcrl.Value)
        }
        ElseIf($TokenOutputFormat -eq 'RAW')
        {
            Write-Output $idcrl.Value
        }
    }
}
catch
{
    ToLog $error[0]
    If ($error[0].Exception.ToString().ToLower().Contains('null'))
    {
        ToLog 'Missing Username or Password'
        Write-Host 'The username or password is missing.' -ForegroundColor Red
        Write-Host 'Please correct the variable and try again.' -ForegroundColor Red
        Write-Host 'The script halted.' -ForegroundColor Yellow
        Break
    }
    If ($error[0].Exception.ToString().ToLower().Contains('forbidden'))
    {
        ToLog 'Invalid credentials provided (forbidden.'
        Write-Host 'The username or password provided could not be validated.' -ForegroundColor Red
        Write-Host 'Please correct the variable and try again.' -ForegroundColor Red
        Write-Host 'The script halted.' -ForegroundColor Yellow
        Break
    }
    If ($error[0].ErrorDetails.Message.Contains('The security token could not be authenticated or authorized'))
    {
        ToLog 'Invalid credentials provided.'
        Write-Host 'The username or password provided could not be validated.' -ForegroundColor Red
        Write-Host 'Please correct the variable and try again.' -ForegroundColor Red
        Write-Host 'The script halted.' -ForegroundColor Yellow
        Break
    }
}

#######################################################################

##################################
#                                #
# From here we just do the calls #
#                                #
##################################


# Just a little chit-chat
$Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, 11
Write-Host 'The data collection started on:' -NoNewline
$StartTime = Get-Date
$StartTimeStr = '{0:yyyy.MM.dd HH:mm:ss}' -f $StartTime
Write-Host $StartTimeStr -ForegroundColor Green

If($Iteration -gt 0)
{
    Write-Host 'The expected end time is: ' -NoNewline
    $ExpEndtime = $StartTime.AddSeconds($Iteration*$Sleep)
    $ExpEndtimeStr = '{0:yyyy.MM.dd HH:mm:ss}' -f $ExpEndtime
    Write-Host $ExpEndtimeStr -ForegroundColor Cyan
}

# This is where we start running the requests
$ProgressCounter = 0
For ($i=1; $i -le $Iteration; $i++)
{
    # Getting the current time
    $IterationTime = Get-Date
    $IterationTimeStr = '{0:dd.MM.yyyy HH:mm:ss}' -f $IterationTime

    If((!$SilentMode) -and $Iteration -gt 1)
    {
        Write-Progress -Activity 'Running page queries...' -PercentComplete (($i / $Iteration)*100)
    }

    # Prepare the Web Client for the queries
    $WebClient = New-Object System.Net.WebClient 

    $WebClient.Headers.Add("Cookie", $spoidcrl)
    $WebClient.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f")
    $WebClient.Headers.Add("Accept", "text/html, application/xhtml+xml, image/jxr, */*")
    #$WebClient.Headers.Add("Accept-Encoding", "gzip, deflate")
    $WebClient.Headers.Add("Accept-Language", "en-US,en;q=0.8,hu;q=0.6,de-CH;q=0.4,de;q=0.2")
    $WebClient.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; Touch; rv:11.0) like Gecko")
    $WebClient.Headers.Add("Cache-Control", "no-cache");

    # And run the query
    $stopwatch =  [system.diagnostics.stopwatch]::StartNew()
    $PageData = $WebClient.DownloadData($SPOPageUrl)
    $stopwatch.Stop()
    $pageDownloadTime = $stopwatch.ElapsedMilliseconds.tostring()
    $PageDataASCII = [System.Text.Encoding]::ASCII.GetString($PageData)
    $StartPerData = $PageDataASCII.IndexOf('"perf" : {') + 10
    $EndOfPerfData = $PageDataASCII.IndexOf('},',$StartPerData) 
    $StatString = $PageDataASCII.Substring($StartPerData,$EndOfPerfData - $StartPerData)
    Write-Host 'Perfstring:'
    Write-Host $StatString
    

    If($WebClient.ResponseHeaders["SPRequestDuration"])
    {
        $SPRequestDuration = $WebClient.ResponseHeaders["SPRequestDuration"].ToString()
    }
    Else
    {
        $DurationStart = $StatString.IndexOf('spRequestDuration')
        If($DurationStart -gt 0)
        {
            $DurationStart = $StatString.IndexOf(':"',$DurationStart) + 2
            $DurationEnd = $StatString.IndexOf('",',$DurationStart)
            $Duration = $StatString.Substring($DurationStart, $DurationEnd - $DurationStart)
            [int]$SPRequestDuration = $Duration
        }
        Else
        {
            [int]$SPRequestDuration = 0
        }
    }

    If($WebClient.ResponseHeaders["SPIisLatency"])
    {
        $SPIisLatency = $WebClient.ResponseHeaders["SPIisLatency"].ToString()
    }
    Else
    {
        $LatencyStart = $StatString.IndexOf('IisLatency')
        If($LatencyStart -gt 0)
        {
            $LatencyStart = $StatString.IndexOf(':"',$LatencyStart) + 2
            $LatencyEnd = $StatString.IndexOf('",',$LatencyStart) -1
            $Latency = $StatString.Substring($LatencyStart, $LatencyEnd - $LatencyStart)
            [int]$SPIisLatency = $Latency
        }
        Else
        {
            [int]$SPIisLatency = 0
            $PageDataASCII
        }
    }
    
    If($WebClient.Headers["Statuscode"])
    {
        	
        $HTTPStatus = $WebClient.Headers["Statuscode"].ToString()
    }

    If($WebClient.ResponseHeaders["X-MSEdge-Ref"])
    {
        $MSEdgeRef = $WebClient.ResponseHeaders["X-MSEdge-Ref"].ToString()
    }

    If($WebClient.ResponseHeaders["SPRequestGuid"])
    {
        $SPRequestGuid = $WebClient.ResponseHeaders["SPRequestGuid"].ToString()
    }
    Else
    {
        $SPRequestGuid = 'N/A'
    }

    if($WebClient.ResponseHeaders["x-sharepointhealthscore"])
    {
        $SPHealthScore = $WebClient.ResponseHeaders["x-sharepointhealthscore"].ToString()
    }
    Else
    {
        $SPHealthScore = 'N/A'
    }

    $RequestTotal = [convert]::ToInt32($SPRequestDuration) + [convert]::ToInt32($SPIisLatency)

    If((!$SilentMode) -or ($Iteration -eq 1))
    {
        Write-Host 'Time: ' -NoNewline 
        Write-Host $IterationTime -ForegroundColor Green -NoNewline
        Write-Host ' MSEdge-Ref: ' -NoNewline
        Write-Host $MSEdgeRef -ForegroundColor Green -NoNewline
        Write-Host ' SPRequestDuration: ' -NoNewline
        Write-Host $SPRequestDuration -ForegroundColor Green -NoNewline
        Write-Host ' SPIisLatency: ' -NoNewline
        Write-Host $SPIisLatency -ForegroundColor Green -NoNewline
        Write-Host ' Total Request Duration: ' -NoNewline
        Write-Host $RequestTotal -ForegroundColor Green -NoNewline
        Write-Host ' SPRequestGuid: ' -NoNewline
        Write-Host $SPRequestGuid -ForegroundColor Green -NoNewline
        Write-Host ' BrowserRTT: ' -NoNewline
        Write-Host $pageDownloadTime -ForegroundColor Green -NoNewline
        Write-Host ' SPHealthScore: ' -NoNewline
        Write-Host $SPHealthScore -ForegroundColor Green
        
    }

    # Writing the output to a file
    If($Outputfile)
    {
        $Row = "$SPOPageUrl`t$IterationTimeStr`t$SPRequestDuration`t$SPIisLatency`t$RequestTotal`t$SPHealthScore`t$pageDownloadTime`t$SPRequestGuid`t$MSEdgeRef"
        Try
        {
            $Row | Out-File $Outputfile -Append
        }
        Catch
        {
            Write-Host 'Could not write the output file.' -ForegroundColor Red
            Write-Host 'The script halted.' -ForegroundColor Yellow
            Break
        }
    }

    # Then wait a minute
    Start-Sleep -Seconds $Sleep
}
