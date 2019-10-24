Import-Module "$PSScriptRoot/powershell-yaml"

function Out-TocInfo
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $TocDirPath
    )

    $tocInfo = Get-TocInfo -FilePath $TocDirPath

    $output = @()
    foreach ($obj in $tocInfo)
    {
        $output += New-Object PSObject -Property $obj
    }
    
    $outputPath = "$(Split-Path -Path $PSScriptRoot -Parent)/Data/Reports/Final"
    $output | Export-Csv -LiteralPath "$OutputPath/TocInfo.csv" -NoTypeInformation
}

function Get-TocInfo
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $FilePath
    )

    $return = [System.Collections.ArrayList]::new()
    $tocFile =  Get-ChildItem -Path $FilePath -Filter *toc.yml
    [string[]]$toc = Get-Content $tocFile.FullName
    foreach ($line in $toc) 
    {
        $content = $content + "`n" + $line
    }

    $yaml = ConvertFrom-YAML $content -Ordered
    $dataPath = "$(Split-Path -Path $PSScriptRoot -Parent)/Data/Reports/Filtered"
    $csvFiles = Get-ChildItem -Path $dataPath -Filter '*9-19.csv'
    $csvInfo = [System.Collections.ArrayList]::new()
    foreach ($file in $csvFiles)
    {
        $csvObjects = Import-Csv -Path $file.FullName
        foreach ($object in $csvObjects)
        {
            $null = $csvInfo.Add($object)
        }
    }

    foreach ($object in $yaml)
    {
        if ($object.items)
        {
            $area = $object.Name

            $add = Get-ItemInfo -Items $object.Items -Area $area -CsvData $csvInfo
        }
        else
        {
            if ($FilePath -match ".*toc\.json")
            {
                $tocPath = Get-TocPath -Href $href
            }

            $add = Get-HrefInfo -FilePath $FilePath -Name $object.Name -Href $object.Href -Area $area -CsvData $csvInfo
        }

        foreach ($addItem in $add)
        {
            $null = $return.add($addItem)
        }
    }

    $null = return $return
}

function Get-ItemInfo
{
    param
    (
        [Parameter(Mandatory = $true)]
        $Items,

        [Parameter()]
        [AllowNull()]
        [string]
        $Area,

        [Parameter()]
        $CsvData
    )

    $return = [System.Collections.ArrayList]::new()

    foreach($item in $Items)
    {
        if ($item.items)
        {
            $add = Get-ItemInfo -Items $item.Items -Area $Area -CsvData $CsvData
        }
        else
        {
            $add = Get-HrefInfo -FilePath $FilePath -Name $item.Name -Href $item.Href -Area $Area -CsvData $CsvData
        }

        foreach ($addItem in $add)
        {
            $null = $return.add($addItem)
        }
    }

    $null = return $return
}

function Get-HrefInfo
{
    param
    (
        [Parameter(Mandatory = $true)]
        [string]
        $FilePath,

        [Parameter(Mandatory = $true)]
        [string]
        $Name,

        [Parameter(Mandatory = $true)]
        [string]
        $Href,

        [Parameter()]
        [AllowNull()]
        [string]
        $Area,

        [Parameter()]
        $CsvData
    )

    if ($Href -notmatch ".*\.md" -or $href -match ".*\.md#.*")
    {
        return
    }

    $fullPath = "$FilePath/$HRef"
    $uri = $Href.Trim('.md')
    $csvObject = $CsvData | Where-Object -FilterScript {$_.LiveUrl -match ".*$uri"}

    $msAuthor = (Select-String -Path $fullPath -Pattern "(?<=.*ms\.author:\s+).*").Matches.Value
    $author = (Select-String -Path $fullpath -Pattern "(?<=^author:\s+).*").Matches.Value
    $manager = (Select-String -Path $fullpath -Pattern "(?<=manager:\s+).*").Matches.Value
    $msTopic = (Select-String -Path $fullpath -Pattern "(?<=.*ms\.topic:\s+).*").Matches.Value
    $msProd = $csvObject.MsProd | Select-Object -Unique
    $msTechnology = $csvObject.MsTechnology | Select-Object -Unique
    $views = [int32] $csvObject.'PageViews '

    $return = [ordered]@{
        Name         = $Name
        Area         = $Area
        Href         = $Href
        MsAuthor     = $msAuthor
        Author       = $author
        Manager      = $manager
        MsTopic      = $msTopic
        MsProd       = $msProd
        MsTechnology = $msTechnology
        PageViews    = $views
    }

    return $return
}

function Find-ReportData
{
    $filePath = "$(Split-Path -Path $PSScriptRoot -Parent)/Data/"
    $csvFiles = Get-ChildItem -Path $filePath -Filter "*.csv"

    foreach ($file in $csvFiles)
    {
        $csvObject = Import-Csv -Path $file.FullName
        $visualStudio = $csvObject | Where-Object -FilterScript {$_.LiveUrl -match ".*\/VisualStudio\/.*" -and $_.LiveUrl -notMatch ".*release-notes.*"}
        $date = $((Get-Date $visualStudio[0].Date).ToString("MM-yy"))
        #$devOpsPath = "$(Split-Path -Path $PSScriptRoot -Parent)/Data/Reports/Filtered/$date.csv"
        $visualStudioPath = "$(Split-Path -Path $PSScriptRoot -Parent)/Data/Reports/Filtered/$date.csv"

        #$devOps | Export-Csv -Path $devOpsPath -Append -Force
        $visualStudio | Export-Csv -Path $visualStudioPath -Append -Force

        Move-Item -Path $file.FullName -Destination "$(Split-Path -Path $PSScriptRoot -Parent)/Data/Reports/Unfiltered" -Force
    }
}

function Import-DocData
{
    param
    (
        [Parameter()]
        $Area
    )

    if ($Area)
    {
        $files = Get-ChildItem -Path "$PSScriptRoot/../Data/Reports/Filtered/$Area"
    }
    else
    {
        $files = Get-ChildItem -Path "$PSScriptRoot/../Data/Reports/Filtered" -Recurse
    }

    $returnObject = [System.Collections.ArrayList]::new()
    foreach ($file in $files)
    {
        $csv = Import-Csv -Path $file.Fullname
        foreach ($object in $csv)
        {
            $null = $returnObject.add($object)
        }
    }

    return $returnObject
}
