# Version 1.1
# 5-1-23
# API data still needs to be sorted through and hopefully converted into PSHELL object
# Will also have to change the user from entering game name to possibly game ID

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Clear-Variable results,regkeythere,html -ErrorAction SilentlyContinue
Remove-Variable -Name html -ErrorAction SilentlyContinue
## Manage list of games to watch for, string
$inputList = Read-Host -Prompt "Enter in the games you want to search for, seperate each item with a comma"

#Turn input into array
[array]$list = (($inputList -split ",").trim())
[array]$results = @()

## Gain number of pages to sort through
## Outputs just an int
function numberOfPages
{
    $apiCaptureFirstPage = Invoke-RestMethod -uri "https://api.geekdo.com/api/listitems?page=1&listid=208913"
    if ($($apiCaptureFirstPage.pagination.total)%$($apiCaptureFirstPage.pagination.perPage) -ne 0) 
        {[int]$pageCount = [Math]::Truncate($($apiCaptureFirstPage.pagination.total)/$($apiCaptureFirstPage.pagination.perPage)+1) }
    Else
        {[int]$pageCount = [Math]::Truncate($($apiCaptureFirstPage.pagination.total)/$($apiCaptureFirstPage.pagination.perPage)) }

    return($pageCount)
}

[int]$maxPages = numberOfPages
write-host "max pages are $maxpages"
## Start loop of searching each page
    ## For each match tag the page number it was on and save it in the results variable
    ## In addition try to regex price
    for ($i = 1; $i -le $maxPages; $i++)
        {
            Add-Type -AssemblyName "Microsoft.mshtml" -ErrorAction SilentlyContinue
            Add-Type -Path "C:\Program Files (x86)\Microsoft.NET\Primary Interop Assemblies\Microsoft.mshtml.dll"
            $html = New-Object -ComObject "HTMLFile"
            $apiPagePull = (Invoke-RestMethod -uri "https://api.geekdo.com/api/listitems?page=$i&listid=208913")
            # If office exists, write the html differently.
            $officeregkeythere = Get-Item -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office -ErrorAction SilentlyContinue
            if ($officeregkeythere -eq $null)
                {
                   $html.IHTMLDocument2_write($content)
                }
            Else
                {
                    $src = [System.Text.Encoding]::Unicode.GetBytes($content)
                    $html.write([ref]$src)
                }


            Write-Host "Reading page $i for the following games"
            Write-host $list

            ## For each game in list

            Foreach ($game in $list)
            {

                ## Regex for game
                    $hits = $html.body.outerText | Select-String -Pattern ".*($game).*" -AllMatches | % {$_.matches}

                    foreach ($hit in $hits)
                        {
                            Clear-Variable matches -ErrorAction SilentlyContinue
                            ## If game found regex for price
                            $price = $hit.value -match "$game.*?( \d+|\$\d+)|( \d+|\$\d+).*?$game"
                            ## Write output to $results
                            $results += "$game found on page $i - Price $($matches[1])`r`n"
                        }


             }
            # Cls
        }


## If results has value display hits and give links to the pages
Write-host "==== Results ====" -ForegroundColor Red
write-host $results
write-host "==== For the following games ====" -ForegroundColor Red
foreach ($item in $list) {write-host $item}
Write-host "URL to the forum page has been copied to clipboard"
"https://boardgamegeek.com/geeklist/208913/twin-cities-games-sale-notifications-and-kickstart/page/1" | Set-Clipboard
pause