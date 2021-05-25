#List of os
# OperatingSystem -ScriptBlock { "Windows Server 2019", "Windows Server 2016", "Windows 10", "Windows 8.1", "Windows Server 2012 R2", "Windows 8", "Windows Server 2012", "Windows Server 2012 Hyper-V", "Windows 7", "Windows Server 2008 R2", "Windows Vista", "Windows Server 2008", "Windows Small Business Server (SBS) 2008", "Windows Server 2003", "Windows Small Business Server (SBS) 2003", "Windows XP", "Windows XP Media Center Edition (MCE)", "Windows XP Tablet PC Edition", "Windows 2000", "Small Business Server (SBS) 2000", "Windows NT 4.0", "Windows Millennium Edition (ME)", "Windows 98 Second Edition (SE)", "Windows 98", "Windows 95", "Microsoft Windows Update", "Windows Embedded Compact 2013", "Windows Embedded Compact 7", "Windows Embedded CE 6.0", "Windows CE 5.0", "Windows CE .NET 4.2", "Windows CE .NET 4.1" }
# Register-PSFTeppScriptblock -Name Product -ScriptBlock { "Exchange Server 2019", "Exchange Server 2016", "Exchange Server 2013", "Exchange Server 2010", "Exchange Server 2007", "Exchange Server 2003", "Exchange Server 2000", "Exchange Server 5.5", "Exchange Server 5.0", "Exchange Server 4.0", "Microsoft Office 365", "Outlook 2019", "Excel 2019", "Word 2019", "Access 2019", "Outlook 2016", "Excel 2016", "Word 2016", "Access 2016", "Outlook 2013", "Excel 2013", "Word 2013", "Access 2013", "Outlook 2010", "Excel 2010", "Word 2010", "Access 2010", "Outlook 2007", "Excel 2007", "Word 2007", "Access 2007", "PowerPoint 2007", "Visio 2007", "Publisher 2007", "Project 2007", "OneNote 2007", "InfoPath 2007", "Microsoft Office Groove 2007", "Outlook 2003", "Excel 2003", "Word 2003", "Access 2003", "PowerPoint 2003", "FrontPage 2003", "Visio 2003", "Publisher 2003", "Project 2003", "OneNote 2003", "InfoPath 2003", "Outlook 2002 (Outlook XP)", "Excel 2002 (Excel XP)", "Word 2002 (Word XP)", "Access 2002 (Access XP)", "PowerPoint 2002 (PowerPoint XP)", "FrontPage 2002 (FrontPage XP)", "Visio 2002 (Visio XP)", "Publisher 2002 (Publisher XP)", "Project 2002 (Project XP)", "Outlook 2000", "Excel 2000", "Word 2000", "Access 2000", "PowerPoint 2000", "FrontPage 2000", "Visio 2000", "Publisher 2000", "Project 2000", "Microsoft Office Live Meeting 2005", "Microsoft Works Suite 2003", "Microsoft Works Suite 2002", "Microsoft Works Suite 2001", "Microsoft Works 2000", "SharePoint Server 2019", "SharePoint Server 2016", "SharePoint Server 2013", "SharePoint Server 2010", "SharePoint Server 2007", "SharePoint Portal Server 2003", "SharePoint Portal Server 2001", "BizTalk Server 2006", "BizTalk Server 2004", "BizTalk Server 2002", "BizTalk Server 2000", "Internet Security and Acceleration (ISA) Server 2006", "Internet Security and Acceleration (ISA) Server 2004", "Internet Security and Acceleration (ISA) Server 2000", "System Center Essentials (SCE) 2010", "System Center Essentials (SCE) 2007", "System Center Operations Manager (SCOM) 2012", "System Center Virtual Machine Manager (SCVMM) 2012", "System Center Orchestrator (SCO) 2012", "System Center Service Manager (SCSM) 2012", "System Center Configuration Manager (SCCM) 2012", "System Center Configuration Manager (SCCM) 2007", "Systems Management Server (SMS) 2003", "Systems Management Server (SMS) 2.0", "Systems Management Server (SMS) 1.2", "Systems Management Server (SMS) 1.1", "Systems Management Server (SMS) 1.0", "SNA Server 4.0", "SNA Server 3.0", "System Center Operations Manager (SCOM) 2007", "Operations Manager (MOM) 2005", "Operations Manager (MOM) 2000", "Host Integration Server (HIS) 2004", "Host Integration Server (HIS) 2000", "Commerce Server 2007", "Commerce Server 2002", "Commerce Server 2000", "Dynamics CRM 3.0", "Zune", "Xbox 360", "Internet Explorer 11", "Internet Explorer 10", "Internet Explorer 9", "Internet Explorer 8", "Internet Explorer 7", "Internet Explorer 6", "Internet Explorer 5.5", "Internet Explorer 5.0", "SQL Server 2017", "SQL Server 2016", "SQL Server 2014", "SQL Server 2012", "SQL Server 2008 R2", "SQL Server 2008", "SQL Server 2005", "SQL Server 2000", "SQL Server 7.0", "Microsoft Data Access Components (MDAC) 2.8", "Microsoft Data Access Components (MDAC) 2.7", "Microsoft Data Access Components (MDAC) 2.6", "Microsoft Data Access Components (MDAC) 2.5", "Microsoft Data Access Components (MDAC) 2.1", "Visual FoxPro 9.0", "Visual FoxPro 8.0", "Visual FoxPro 7.0", "Visual FoxPro 6.0", ".NET Framework 4.7", ".NET Framework 4.6", ".NET Framework 4.5", ".NET Framework 4", ".NET Framework 3.5", ".NET Framework 3.0", ".NET Framework 2.0", ".NET Framework 1.1", ".NET Framework 1.0", "ASP.NET 2.0", "ASP.NET 1.1", "ASP.NET 1.0", "Visual Studio 2008", "Visual Studio 2005", "Visual C++ 2005", "Visual C# 2005", "Visual Basic 2005", "Visual Studio .NET 2003", "Visual C++ .NET 2003", "Visual C# .NET 2003", "Visual Basic .NET 2003", "Visual Studio .NET 2002", "Visual C++ .NET 2002", "Visual C# .NET 2002", "Visual Basic .NET 2002", "Visual Studio 6.0", "Visual C++ 6.0", "Visual Basic 6.0", "Windows Media Player 11", "Windows Media Player 10", "Windows Media Player 9", "Internet Information Services (IIS) 7.0", "Internet Information Services (IIS) 6.0", "Internet Information Services (IIS) 5.1", "Internet Information Services (IIS) 5.0", "Office Accounting 2007", "Small Business Accounting 2006", "Money 2007", "Money 2006", "Money 2005", "Money 2004", "Money 2003", "Money 2002", "Money 2001", "Visual SourceSafe 6.0", "Microsoft Encarta Encyclopedia 2000", "Age of Empires III (AoE3)", "Age of Empires II (AoE2)", "Age of Mythology", "Zoo Tycoon 2", "Zoo Tycoon", "Microsoft Mail for Appletalk Networks 3.1", "Microsoft Mail for Appletalk Networks 3.0" }


#Update Catalog:
# https://www.catalog.update.microsoft.com/Search.aspx?q=activex

# Project Stole most of this code from:
# https://github.com/potatoqualitee/kbupdate

$kb = "KB4570334"
#$kb = "KB4586864"
# $OperatingSystem = "Windows Server 2016"


#Write-Host "$kb"
#Write-Host  "Searching catalog for $kb"
#if ($OperatingSystem) {
#    $url = "https://www.catalog.update.microsoft.com/Search.aspx?q=$kb+$OperatingSystem"
#    Write-PSFMessage -Level Verbose -Message "Accessing $url"
#    $results = Invoke-TlsWebRequest -Uri $url
#    $kbids = $results.InputFields |
#        Where-Object { $_.type -eq 'Button' -and $_.Value -eq 'Download' } |
#        Select-Object -ExpandProperty ID
#          }




#Get GUIDs
$url = "https://www.catalog.update.microsoft.com/Search.aspx?q=$kb"
#$results = Invoke-TlsWebRequest -Uri $url
$results = Invoke-WebRequest -Uri $url
$kbids = $results.InputFields |
    Where-Object { $_.type -eq 'Button' -and $_.Value -eq 'Download' } |
    Select-Object -ExpandProperty ID



#Get Links???
$resultlinks = $results.Links |
    Where-Object ID -match '_link' |
    Where-Object { $_.OuterHTML -match ( "(?=.*" + ( $Filter -join ")(?=.*" ) + ")" ) } 

#Get Titles And Guids (This is probably what I want)
$guids = @()
foreach ($resultlink in $resultlinks) {
    $itemguid = $resultlink.id.replace('_link', '')
    $itemtitle = ($resultlink.outerHTML -replace '<[^>]+>', '').Trim()
    if ($itemguid -in $kbids) {
        $guids += [pscustomobject]@{
            Guid  = $itemguid
            Title = $itemtitle
        }
    }
}

#$guids

#Used!!!
function Get-Info ($Text, $Pattern) {
    if ($Pattern -match "labelTitle") {
        # this should work... not accounting for multiple divs however?
        [regex]::Match($Text, $Pattern + '[\s\S]*?\s*(.*?)\s*<\/div>').Groups[1].Value
    } elseif ($Pattern -match "span ") {
        [regex]::Match($Text, $Pattern + '(.*?)<\/span>').Groups[1].Value
    } else {
        [regex]::Match($Text, $Pattern + "\s?'?(.*?)'?;").Groups[1].Value
    }
}

# Not used (But Don't Delete!!!)
function Get-SuperInfo ($Text, $Pattern) {
    # this works, but may also summon cthulhu
    $span = [regex]::match($Text, $pattern + '[\s\S]*?<div id')

    switch -Wildcard ($span.Value) {
        "*div style*" { $regex = '">\s*(.*?)\s*<\/div>' }
        "*a href*" { $regex = "<div[\s\S]*?'>(.*?)<\/a" }
        default { $regex = '"\s?>\s*(\S+?)\s*<\/div>' }
    }

    $spanMatches = [regex]::Matches($span, $regex).ForEach( { $_.Groups[1].Value })
    if ($spanMatches -eq 'n/a') { $spanMatches = $null }

    if ($spanMatches) {
        foreach ($superMatch in $spanMatches) {
            $detailedMatches = [regex]::Matches($superMatch, '\b[kK][bB]([0-9]{6,})\b')
            # $null -ne $detailedMatches can throw cant index null errors, get more detailed
            if ($null -ne $detailedMatches.Groups) {
                [PSCustomObject] @{
                    'KB'          = $detailedMatches.Groups[1].Value
                    'Description' = $superMatch
                } | Add-Member -MemberType ScriptMethod -Name ToString -Value { $this.Description } -PassThru -Force
            }
        }
    }
}

$downloaddialogs = @()
foreach ($item in $guids) {
    $guid = $item.Guid
    $itemtitle = $item.Title
    $post = @{ size = 0; updateID = $guid; uidInfo = $guid } | ConvertTo-Json -Compress
    $body = @{ updateIDs = "[$post]" }
    $downloaddialogs += Invoke-WebRequest -Uri 'https://www.catalog.update.microsoft.com/DownloadDialog.aspx' -Method Post -Body $body | Select-Object -ExpandProperty Content
    
}

$links = @()
foreach ($downloaddialog in $downloaddialogs) {
    $title = Get-Info -Text $downloaddialog -Pattern 'enTitle ='
    $arch = Get-Info -Text $downloaddialog -Pattern 'architectures ='
    $longlang = Get-Info -Text $downloaddialog -Pattern 'longLanguages ='
    $updateid = Get-Info -Text $downloaddialog -Pattern 'updateID ='
    $ishotfix = Get-Info -Text $downloaddialog -Pattern 'isHotFix ='
    
    if ($ishotfix) {
        $ishotfix = "True"
    } else {
        $ishotfix = "False"
    }
    if ($longlang -eq "all") {
        $longlang = "All"
    }
    if ($arch -eq "") {
        $arch = $null
    }
    if ($arch -eq "AMD64") {
        $arch = "x64"
    }
    if ($title -match '64-Bit' -and $title -notmatch '32-Bit' -and -not $arch) {
        $arch = "x64"
    }
    if ($title -notmatch '64-Bit' -and $title -match '32-Bit' -and -not $arch) {
        $arch = "x86"
    }

    $downloaddialog = $downloaddialog.Replace('www.download.windowsupdate', 'download.windowsupdate')
    #$links = $downloaddialog | Select-String  -Pattern "(http[s]?\://download\.windowsupdate\.com\/[^\'\""]*)" | Select-Object -Unique
    $links += ($downloaddialog | Select-String -AllMatches -Pattern "(http[s]?\://download\.windowsupdate\.com\/[^\'\""]*)").Matches.Value

    
    
 }


 $links | select -Unique


 #Todo, this script will download the msu installers based off of of KB, it will remove duplicate links, thus
 
 #Create alink list to and download
 #Create a kb list to download
 #Figour out how to install msu, I think it might go like this: Start-Process -FilePath "wusa.exe" -ArgumentList "$SourceFolder$KB.msu /quiet /norestart" -Wait }
 # or this: WUSA C:\Games\Temp\KBDownload\windows.msu /quiet /norestart (It looks like I can also output a log file)


