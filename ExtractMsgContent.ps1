$msgPath = 'c:\path\to\msg\files'
$filenameHash = @(); #an array that will hold objects; info for each row to be outputted to the CSV file

$outlook = New-Object -ComObject Outlook.Application #setup interface with outlook

#create a csv file with a random file name in the source's path
$randoFileName = [System.Guid]::NewGuid().ToString() + '.csv'
$pathToOutputFile = Join-Path $msgPath $randoFileName
Add-Content -Path $pathToOutputFile -Value 'FilePath,Class,Title,Employee,Date,Note'

#get more information from active directory
function queryActiveDirectory($mailTo) {

    $account = Get-ADUser -LdapFilter "(&(objectClass=user)(mail=$mailTo))" -Properties samAccountName, cn, surname, GivenName
   
    return $account

}

#iterate through the messages, and extract data
#send the info to an array of objects so that it can be outputted to a CSV file
Get-ChildItem $msgPath -Filter *.msg |
ForEach-Object {

    #open the message; $_.FullName is the current message's full path name (current child item)
    $msg = $outlook.Session.OpenSharedItem($_.FullName)

    #Recipients contains the email address the message was sent to; however,
    #Recipients is System.__ComObject that I can't read without this trick
    #https://stackoverflow.com/questions/47264561/how-to-get-email-address-from-the-emails-inside-an-oulook-folder-via-powershell
    $mailTo = $msg.Recipients | ForEach-Object {
        $_.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
    }

    #get the message subject line and date sent
    $subject = $msg | Select-Object -ExpandProperty Subject
    $sent = $msg | select-object -ExpandProperty SentOn

    #call the Active Directory query function to get more information about the user that isn't available in the *.msg
    $account = queryActiveDirectory $mailTo

    #define some variables with some user info
    $surname = $account | select-object -ExpandProperty surname
    $GivenName = $account | select-object -ExpandProperty GivenName


    #close and clear the message
    $msg.Close(1)
    $msg = $null

    #send a message to the console indicated the current file being processed
    write-host 'Processing: ' $_

    #append an object to the array
    #each object represents a row that will be written to the csv file with content defined above
    $filenameHash +=
    [pscustomobject]@{
        FilePath = $_.FullName
        Class    = 'The Document Class'
        Title    = $subject
        Employee = $surname + ", " + $GivenName + "_" + $mailTo
        Date     = get-date $sent -Format 'MM-dd-yyyy'
        Note     = 'Internal Note for Staff'
    }



}

# close outlook
$outlook.quit()

# I SAID CLOSE OUTLOOK!
Stop-Process -Name Outlook -Force
Start-Sleep -m 500


# for each object (row) in the hash, send to CSV file
# use a | (pipe) as the delimiter, and remove quotes from output
$filenameHash | convertto-csv -NoTypeInformation -Delimiter "|" | % { $_ -replace '"', '' } | out-file $pathToOutputFile