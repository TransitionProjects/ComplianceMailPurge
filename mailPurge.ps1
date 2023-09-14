#by Jelly Rinne for TPI September 2023

#Self-elevate the script if required
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
        $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
        Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
        Exit
    }
}

#Check for the ExchangeOnlineManagement module. If not installed, install it
if(-not (Get-Module ExchangeOnlineManagement)) {
    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
}

#import the module for use
Import-Module ExchangeOnlineManagement

#intialize variables
$queryName = $null
$mainMenuInput = $null
$sessionMenuInput = $null
$purgeTypeMenuInput = $null
$query = $null
$queryStatus = $null
$queryActionStatus = $null
$queryActionName = $null

################################################# Begin Session functions #################################################

#[1] - Disconnect existing sessions and create a fresh one
function Disconnect-AndStartNew {
    Disconnect-ExchangeOnline -confirm:$false
    Write-Host "Sign in with a profile that has Exchange Admin/Compliance & Security Role Access" -ForegroundColor Green
    Connect-IPPSSession
    Invoke-MainMenu
}

#[2] - Create a new session without disconnecting existing ones
function Connect-ESOOvertop {
    Write-Host "Sign in with a profile that has Exchange Admin/Compliance & Security Role Access" -ForegroundColor Green
    Connect-IPPSSession
    Invoke-MainMenu
}

#[3] - Disconnect existing sessions and exit
function Exit-Clean {
    Disconnect-ExchangeOnline -confirm:$false
    Clear-Host
    Exit
}

#[4] - Exit without doing anything
function Exit-Dirty {
    Clear-Host
    Exit
}

#Session Menu
function Connect-Session {
    Clear-Host
    if (-not (Get-ConnectionInformation)) {
        Connect-IPPSSession
        Invoke-MainMenu
    } else {
        Clear-Host
        Write-Host "A Powershell session connected to Exchange Online or other service already exists!" -ForegroundColor Red
        Write-Host "[.............Exchange Search & Purge.............]" -ForegroundColor Green
        Write-Host "[1] Terminate the existing session, connect with a new one, and continue" -ForegroundColor Yellow
        Write-Host "[2] Create a new session without disconnecting the existing session" -ForegroundColor Yellow
        Write-Host "[3] Disconnect the existing session and exit" -ForegroundColor Yellow
        Write-Host "[4] Exit without disconnecting the existing session" -ForegroundColor Yellow

        $sessionMenuInput = Read-Host -Prompt "Enter a menu option to continue"

        if ($sessionMenuInput -eq "1") {
            Disconnect-AndStartNew
        } elseif ($sessionMenuInput -eq "2") {
            Connect-ESOOvertop
        } elseif ($sessionMenuInput -eq "3") {
            Exit-Clean
        } elseif ($sessionMenuInput -eq "4") {
            Exit-Dirty
        } else {
            $sessionMenuInput = Read-Host -Prompt "Invalid Input. Enter a menu option to continue" 
        }
    }
}

################################################# End Session Functions #################################################

################################################# Begin Tool Functions #################################################

#[1] - Construct a search query
function Set-Query {
    $queryName = Read-Host -Prompt "Give your query a name"
    $query = Read-Host -Prompt "Enter your query. Examples can be found at https://learn.microsoft.com/en-us/powershell/module/exchange/new-compliancesearch?view=exchange-ps"
    New-ComplianceSearch -Name $queryName -ExchangeLocation "All" -ContentMatchQuery $query
    Write-Host "Query set to: " + $query
    Pause
    Invoke-MainMenu
}

#[2] - Modify your constructed query
function Repair-Query {
    $query = Read-Host -Prompt "Enter your modified query"
    Set-ComplianceSearch -Identity $queryName -ContentMatchQuery $query
    Write-Host "Query set to: " + $query
    Pause
    Invoke-MainMenu

}

#[3] - Run search with constructed query
function Test-Query {
    #Start the Compliance Search
    Start-ComplianceSearch -Identity $queryName 

    #Wait for search to complete before continuing
    $queryStatus = (Get-ComplianceSearch -Identity $queryName).Status
    Write-Host "Searching, this can take a few minutes..." -ForegroundColor Green
    while ($queryStatus -ne "Completed") {
        $queryStatus = (Get-ComplianceSearch -Identity $queryName).Status
    }

    #Display results
    Write-Host "Found results in:" -ForegroundColor Yellow

    #Load horrible search results object into a variable
    $complianceSearch = Get-ComplianceSearch -Identity $queryName
    $queryResults = $complianceSearch.SuccessResults

    #Split results object into a usable array, because "successresults" is an insane property
    $resultArray = $queryResults.Split([Environment]::NewLine,
                                       [System.StringSplitOptions]::RemoveEmptyEntries) | Where-Object {
                                        $_ -notlike "*Item count: 0*"
                                       }

    #Make it readable, every location on a new line, display, pause for user input
    $resultArray | ForEach-Object {
        Write-Host $_ -ForegroundColor Green
    }
    Pause
    Invoke-MainMenu
}

#[4] - Purge all items found with constructed query
function Push-Purge {
    Clear-Host
    Write-Host "Push Purge!" -ForegroundColor Blue
    Write-Host "[.............Exchange Search & Purge.............]" -ForegroundColor Green
    Write-Host "[1] Hard Delete (removes from Exchange)"
    Write-Host "[2] Soft Delete (removes from User Inbox)"
    Write-Host "[Q] Cancel"

    $purgeTypeMenuInput = Read-Host -Prompt "Choose a PurgeType to continue" 

    if ($purgeTypeMenuInput -eq "1") {
        New-ComplianceSearchAction -SearchName $queryName -Purge -PurgeType HardDelete
    } elseif ($purgeTypeMenuInput -eq "2") {
        New-ComplianceSearchAction -SearchName $queryName -Purge -PurgeType SoftDelete
    } elseif ($purgeTypeMenuInput -eq "Q") {
        Invoke-MainMenu
    } else {
        $purgeTypeMenuInput = Read-Host -Prompt "Invalid Input. Enter a menu option to continue" 
    }

    $queryActionName = $queryName + "_Purge"
    $queryActionStatus = (Get-ComplianceSearchAction -Identity $queryActionName).Status
    
    #Wait for purge to complete before displaying results
    Write-Host "Purging! This shouldn't take long..." -ForegroundColor Green
    while ($queryActionStatus -ne "Completed") {
        $queryActionStatus = (Get-ComplianceSearchAction -Identity $queryActionName).Status
    }

    #Display Results, pause for user input
    Get-ComplianceSearchAction -Identity $queryActionName | Format-List Results
    Pause
    Invoke-MainMenu
}

#[5] - Clear constructed query
function Reset-Query {
    Remove-ComplianceSearch -Identity $queryName
    Invoke-MainMenu
    Write-Host "Query removed."
    Pause
    Invoke-MainMenu
}

#[6] - Switch between existing queries
function Select-NewQuery {
    $queryName = Read-Host -Prompt "Enter the name of an existing Query. You can list all existing Queries with 'Get-ComplianceSearch'"
    if (Get-ComplianceSearch -Identity $queryName) {
        Invoke-MainMenu
    } else {
        Write-Host "Invalid query name or input, try again." -ForegroundColor Red
        Select-NewQuery
    }
    Write-Host "Selected " + $queryName + "."
    Pause
    Invoke-MainMenu
}

#[Q] - Clear constructed query and exit
function Reset-Shell {
    Reset-Query
    Disconnect-ExchangeOnline -Confirm:$false
    Clear-Host
    Exit
}

#Main Menu
function Invoke-MainMenu {
    Clear-Host
    Write-Host "This script is for quick & dirty, organization-wide removals of malicious e-mails & spam. It does not provide the full featureset of the ComplianceSearch cmdlets. If you are trying to build a complex query, you are likely better off using the cmdlets manually. Documentation for each cmdlet can be found at this link: https://learn.microsoft.com/en-us/powershell/module/exchange/?view=exchange-ps#policy-and-compliance-content-search" -ForegroundColor Red
    Write-Host "[.............Exchange Search & Purge.............]" -ForegroundColor Green
    Write-Host "[1] Construct a search query" -ForegroundColor Yellow
    Write-Host "[2] Modify the current query" -ForegroundColor Yellow
    Write-Host "[3] Run search with constructed query and show itemized results" -ForegroundColor Yellow
    Write-Host "[4] Purge all items found with constructed query" -ForegroundColor Yellow
    Write-Host "[5] Clear & Remove constructed query" -ForegroundColor Yellow
    Write-Host "[6] Select an existing query or change selection to a different existing query" -ForegroundColor Yellow
    Write-Host "[Q] Exit" -ForegroundColor Yellow

    if ($queryName) {
        Write-Host "Current Query Name: $queryName" -ForegroundColor Green
    }
    
    $mainMenuInput = Read-Host -Prompt "Enter a menu option to continue"

    if ($mainMenuInput -eq "1") {
        Set-Query
    } elseif ($mainMenuInput -eq "2") {
        Repair-Query
    } elseif ($mainMenuInput -eq "3") {
        Test-Query
    } elseif ($mainMenuInput -eq "4") {
        Push-Purge
    } elseif ($mainMenuInput -eq "5") {
        Reset-Query
    } elseif ($mainMenuInput -eq "6") {
        Select-NewQuery
    } elseif ($mainMenuInput -eq "Q") {
        Reset-Shell
    } else {
        $mainMenuInput = Read-Host -Prompt "Invalid Input. Enter a menu option to continue" 
    }
    Invoke-MainMenu
}

################################################# End Tool Functions #################################################

#Go!
Connect-Session



