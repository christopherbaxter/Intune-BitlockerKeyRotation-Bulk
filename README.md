# Intune-BitlockerKeyRotation-Bulk

If you are migrating to Intune Bitlocker management, with Bitlocker Recovery Keys escrowed to AzureAD, this script will allow you to rotate the keys for all Windows 10 devices in AzureAD. 

The reason this script exists is that (as of 15/02/2022), there is no other way to request the devices to rotate their Bitlocker Recovery keys into AzureAD (or escrow the key if none exists in AzureAD) in bulk.

## What is needed for this script to function?

You will need a Service Principal in AzureAD with sufficient rights. I have a Service Principal that I use for multiple processes, I would not advise copying my permissions. I suggest following the guide from <https://msendpointmgr.com/2021/01/18/get-intune-managed-devices-without-an-escrowed-bitlocker-recovery-key-using-powershell/>. My permissions are set as in the image below. Please do not copy my permissions, this Service Principal is used for numerous tasks. I really should correct this, unfortunately, time has not been on my side, so I just work with what work for now. 

![](https://github.com/christopherbaxter/StaleComputerAccounts/blob/main/Images/ServicePrincipal%20-%20API%20Permissions.jpg)

I also elevate my AzureAD account to 'Intune Administrator', 'Cloud Device Administrator' and 'Security Reader'. These permissions also feel more than needed. Understand that I work in a very large environment, that is very fast paced, so I elevate these as I need them for other tasks as well.

You will need to make sure that you have the following PowerShell modules installed. There is a lot to consider with these modules as some cannot run with others. This was a bit of a learning curve. 

ActiveDirectory\
AzureAD\
ImportExcel\
JoinModule\
MSAL.PS\
PSReadline (May not be needed, not tested without this)

Ultimately, I built a VM on-prem in one of our data centres to run this script, including others. My machine has 4 procs and 16Gb RAM, the reason for an on-prem VM is because most of our workforce is working from home (me included), and running this script is a little slow through the VPN. Our ExpressRoute also makes this data collection significantly more efficient. In a small environment, you will not need this.

# Disclaimer

Ok, so my code may not be very pretty, or efficient in terms of coding. I have only been scripting with PowerShell since September 2020, have had very little (if any), formal PowerShell training and have no previous scripting experience to speak of, apart from the '1 liners' that AD engineers normally create, so please, go easy. I have found that I LOVE PowerShell and finding strange solutions like this have become a passion for me.

## Christopher, enough ramble, How does this thing work?

This script will extract all IntuneDeviceIDs from the MS Graph API. Once extracted, the script splits the IntuneDeviceID array into 30 smaller arrays, then will 'post' a command to rotate the Bitlocker Recovery Keys. This method uses the same URLs that would be called when using the Endpoint manager console.

### Parameters and Functions

The first section is where we supply the TenantID (Of the AzureAD tenant) and the ClientID of the Service Principal you have created. If you populate these (hard code), then the script will not ask for these and will immediately go to the Authentication process.

![](https://github.com/christopherbaxter/Intune-BitlockerKeyRotation-Bulk/blob/348f157a46d96611fcff401e25d5fe79ce47c3b1/Images/01-ParamsandFunctions.jpg)

The functions needed by the script are included in the script. I have modified the 'Invoke-MSGraphOperation' function significantly. I was running into issues with the token and renewing it. I also noted some of the errors went away with a retry or 2, so I built this into the function. Sorry @JankeSkanke @NickolajA for hacking at your work. :-)

### The Variables

The variable section also has a section to use the system proxy. I was having trouble with the proxy, intermittently. Adding these lines solved the problem

![](https://github.com/christopherbaxter/Intune-BitlockerKeyRotation-Bulk/blob/348f157a46d96611fcff401e25d5fe79ce47c3b1/Images/02-Variables.jpg)

### The initial Authentication and Token Acquisition

Ok, so now the 'fun' starts.

The authentication and token acquisition will allow for auth with MFA. You will notice in the script that I have these commands running a few times in the script. This allows for token renewal without requiring MFA again. I also ran into some strange issues with different MS Graph API resources, where a token used for one resource, could not be used on the next resource, this corrects this issue, no idea why, never dug too deep into it because I needed it to work, not be pretty. :-)

![](https://github.com/christopherbaxter/Intune-BitlockerKeyRotation-Bulk/blob/348f157a46d96611fcff401e25d5fe79ce47c3b1/Images/03-GenerateAuthtokenandAuthHeader.jpg)

### Intune Device Data Extraction

This section also requires an authentication process and will allow for MFA. The reason why I added this in here is that the script takes a long time to run in my environment, and so, if I perform this extraction first, without the initial auth\token process, the script will complete this process, then sit waiting for auth and MFA, and in essence, not run. Same if this was moved to after the MS Graph extractions. 

Having the 'authy' bits in this order, the script will ask for auth and MFA for MS Graph, then auth and MFA for AzureAD, one after the other with no delay, allowing the script to run without manual intervention. 

    $IntuneDevIDs = [System.Collections.ArrayList]@(Invoke-MSGraphOperation -Get -APIVersion "Beta" -Resource "deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'" -Headers $AuthenticationHeader -Verbose:$VerbosePreference | Where-Object { $_.azureADDeviceId -ne "00000000-0000-0000-0000-000000000000" } | Select-Object id | Sort-Object id)    

I extract the data into an ArrayList. This was needed for a previous 'join' function, I left it like this because I noted strange errors in other scripts. I never had the time to validate that this is the case here, so I simply left it in place. At some point, I would like to test other array types and test processing time between them, not now, this works exactly as needed.

![](https://github.com/christopherbaxter/Intune-BitlockerKeyRotation-Bulk/blob/348f157a46d96611fcff401e25d5fe79ce47c3b1/Images/04-ExtractIntuneDeviceIDs.jpg)

### Splitting the IntuneDeviceID array.

Now things get interesting, splitting the array into multiple equal parts. This was done as I was having trouble with the tokens expiring while the script was running, even with this process being managed in the function called "Invoke-MSGraphOperation" in the script.

![](https://github.com/christopherbaxter/Intune-BitlockerKeyRotation-Bulk/blob/348f157a46d96611fcff401e25d5fe79ce47c3b1/Images/05-SplitTheIntuneDeviceIDArray.jpg)

### The Script block.

The Script block includes a reporting ability. This is in order to be able to get the failures and retry them again. I was struggling with failures, either due to the authentication token expiring, or timeouts. I created the process to retry the command again. Keep in mind that i am running 30 simultaneous threads, and it is probable that I am overloading something along the way, hence the numerous retries.

![](https://github.com/christopherbaxter/Intune-BitlockerKeyRotation-Bulk/blob/348f157a46d96611fcff401e25d5fe79ce47c3b1/Images/06-ScriptBlock-1.jpg)
![](https://github.com/christopherbaxter/Intune-BitlockerKeyRotation-Bulk/blob/348f157a46d96611fcff401e25d5fe79ce47c3b1/Images/07-ScriptBlock-2.jpg)

### The Foreach Loop.

This is the key to the ability to run the commands in parallel. I made use of a module called PoshRSJob. This allowed me to run simultaneous web requests.

The foreach loop is setup to collect the reporting data created in the relevant Script block.

![](https://github.com/christopherbaxter/Intune-BitlockerKeyRotation-Bulk/blob/348f157a46d96611fcff401e25d5fe79ce47c3b1/Images/08-ForeachLoop.jpg)

### Extracting the failures for the retry process.

The code for extracting the failures, works by extracting the successful (and failed requests. This is not the same type of failure mode. This is an actual failure returned, so a successful request, with a failed result, not a failed request) rotation requests. These successful extractions are then 'blended' with the complete list of IntuneDeviceIDs, then extracting the devices without a 'successful\failed' rotation result, and creating an array with these device IDs. Then splitting the array again like above.

![](https://github.com/christopherbaxter/Intune-BitlockerKeyRotation-Bulk/blob/348f157a46d96611fcff401e25d5fe79ce47c3b1/Images/09-FailedRotationExtractionArrayCreationandSplit.jpg)

### Retry the extract for the failures

There is not much more to see here, only that I again run the extract using the PoshRSJob module.

![](https://github.com/christopherbaxter/Intune-BitlockerKeyRotation-Bulk/blob/348f157a46d96611fcff401e25d5fe79ce47c3b1/Images/10-RetryFailures.jpg)

### The extracted data is 'blended' then exported

The code will test for the retries based on the count of objects in the array for the retries. If this array is empty, the script will export the extracted data as is. If there count in the retried array is more than zero, the script will 'blend' the arrays (the successful extract array, and the retried extract array), then spit out a report.

![](https://github.com/christopherbaxter/Intune-BitlockerKeyRotation-Bulk/blob/348f157a46d96611fcff401e25d5fe79ce47c3b1/Images/11-GenerateandExportReport.jpg)

Visitor Counter (Excuse the font, I have no idea what I'm doing)\
![Visitor Count](https://profile-counter.glitch.me/Intune-BitlockerKeyRotation-Bulk/count.svg)