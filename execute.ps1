#--------------------------------------------------------------------------------------------------------
#Purpose: Basic script that downloads Updates from Microsoft Catalog from a list of KBs
#Couldn't use potatoqualitee's project because of restrictions on my system


#Update Catalog:
# https://www.catalog.update.microsoft.com/Search.aspx?q=activex

# Project I stole most of this code from:
# https://github.com/potatoqualitee/kbupdate

#--------------------------------------------------------------------------------------------------------

##VARS
##---------------------------------------------------------------------------
$filePath = './kbs.txt'
$guids = @()
##---------------------------------------------------------------------------

##Functions!!!!!
##---------------------------------------------------------------------------
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
##---------------------------------------------------------------------------

##Get your Guids based off of your KB numbers, these will allow you to pull the actual
##Download Links
##---------------------------------------------------------------------------
#Get GUIDs

foreach ($kb in $(Get-Content -Path $filePath)) {

        $kb
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
}
##---------------------------------------------------------------------------

##This next section of code will take your guids and create download links from them!!!
##---------------------------------------------------------------------------
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

    $downloaddialog = $downloaddialog.Replace('www.download.windowsupdate', 'download.windowsupdate')
    $links += ($downloaddialog | Select-String -AllMatches -Pattern "(http[s]?\://download\.windowsupdate\.com\/[^\'\""]*)").Matches.Value
    
 }
##---------------------------------------------------------------------------


##Now Download Links, with minor sanitation
##---------------------------------------------------------------------------

#Get your links, excluding arm and x86 arch!!!
$links_Clean = $links | where { $_ -cnotmatch '-x86' } | where { $_ -cnotmatch '-arm64' } | select -Unique

Write-Host ""
Write-Host "READY TO DOWNLOAD, PLEASE PROCEED!!!"
PAUSE
Write-Host ""

foreach ( $i in $links_Clean) {

    $i
    Invoke-WebRequest -Uri $i -OutFile $($i.split("/")[-1])

}

##---------------------------------------------------------------------------

