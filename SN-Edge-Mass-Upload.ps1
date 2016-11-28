<#
SN-Edge-Mass-Upload.ps1
Edge Encryption Bulk Record Upload - version 1.0
Jarod Mundt (@j4rodm - j@jarodm.com) - November 2016

Instructions for use:
* CSV source file must have a header row
* Update the script setup variables
* Update the field mapping ($newrecord area in body of script
* Run and monitor for errors

Future Wish List
* OAuth authentication

ServiceNow Permissions/Roles Needed (one of the following)
* Permission to read/write to the destination table (e.g. itil role)
* See Also Role: rest_service (Can use the REST API to perform REST web service operations such as querying or inserting records.)
* See Also Role: web_service_admin	(Can access the System Web Services application menu and the REST modules.)

Version History
* 1.0 - 28NOV2016 - Initial Release
#>

# ***********************************************************
# *                Variable/Script Setup                    *
# ***********************************************************
$sourcecsv  = ".\sn_si_incident.csv"
$startonrow = 2  #Adjust this if the script gets interrupted (Row 1 is the CSV header)

$SNcertignore = $false #Ignore Edge Encryption Self-Signed Certificates
$SNinstance   = "dev10306.service-now.com"
$SNtable      = "sn_si_incident"
$SNuser       = “justadmin”
$SNpass       = "**replace**with**password**"

$headers = @{
“Accept” = “application/json”
"Content-Type”=”application/json"
}
$i = 1
$row = 2

# ***********************************************************
# *                  Main Script Body                       *
# ***********************************************************
$SNPassEnc = ConvertTo-SecureString –String $SNpass –AsPlainText -Force
$SNowCreds = New-Object –TypeName System.Management.Automation.PSCredential –ArgumentList $SNuser, $SNpassenc
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$SNcertignore}
$importlist = Import-Csv -Path $sourcecsv
$rowstotal = $importlist.count
$timestart = Get-Date


foreach ($record in $importlist)
{
    if ($row -lt $startonrow)
    {
        Write-Warning "Row $row is less than the defined starting row. Skipping import."
    }else{
        Write-Host -ForegroundColor Green "Processing row $row of $rowstotal - " -NoNewLine

        #Variables from CSV Cleanup
        $priorityraw = $record.'Priority'.Substring(0,1)
        $priority = [int]$priorityraw

        $newrecord = @{
            #service_now_field_name=$variable
            short_description=$record.'short description'
            description=$record.'description'
            category=$record.'category'
            assignment_group='5f721d93c0a8010e015533746de18bf9'
            state='18'
            caller_id='Abel Tuter'
            priority=$priority
            contact_type='email'
        }
        $json = $newrecord | ConvertTo-Json
        
        try{
            $response = Invoke-RestMethod -Credential $SNowCreds -Headers $headers -Body $json -Method Post -Uri “https://$SNinstance/api/now/table/$SNtable”
            $incidents = $response.result 
            Write-Host -ForegroundColor DarkCyan $incidents.number "("$incidents.short_description") created in SN"

        }catch{
            $err=$_.Exception
            Write-Host -ForegroundColor DarkYellow "REST API ERROR"
            Write-Host -ForegroundColor Yellow $err.Message           
        }

    }

    #Display a status update every 10 records  
    $updatevar = $i % 10
    if ($updatevar -eq 0){
        $timeelapsed = New-TimeSpan ($timestart) (Get-Date)
        $timeelapsedms = $timeelapsed.TotalMilliseconds

        $timeeachrow = New-TimeSpan -Seconds ($timeelapsed.TotalSeconds / $i)
        $rowsleft = $rowstotal - $row
    
        $rowsleftpct = 100 - (($row / $rowstotal)*100)
        $rowsleftpct1 = [math]::Round($rowsleftpct,1)
    
        $rowsdone = $row / $rowstotal
        $rowsdonepct = $rowsdone * 100
        $rowsdonepct1 = [math]::Round($rowsdonepct,1)
    
        $timetotalms = $timeelapsedms / $rowsdone
        $timeremainms = $timetotalms - $timeelapsedms
        $timeremainminraw = $timeremainms/1000/60
        $timeremainmin = [math]::Round($timeremainminraw,1)

        write-host -ForegroundColor Cyan "Elapsed: $rowsdonepct1% in $timeelapsed for $i rows averaging $timeeachrow. Remaining: $rowsleftpct1% and $timeremainmin minutes  "
    }
    
    $i++
    $row++
}


# ***********************************************************
# *                      Cleanup                            *
# ***********************************************************

# Nothing here at the moment



<#

                            ,===     -Help us Obi-Wan Kenobi...
                           (@o o@
                          / \_-/       ___
                         /| |) )      /() \
                        |  \ \/__   _|_____|_
                        |   \____@=| | === | |
                        |   |      |_|  O  |_|
                        | | |       ||  O  ||
                        | | |       ||__*__||
                       /  |  \     |~ \___/ ~|
                       ~~~~~~~     /=\     /=\
_______________________(_)(__\_____[_]_____[_]_____________________
#>