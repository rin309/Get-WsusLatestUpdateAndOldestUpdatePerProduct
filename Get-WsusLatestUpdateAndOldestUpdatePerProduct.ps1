Function Get-WsusLatestUpdateAndOldestUpdatePerProduct($WsusServerName, $WsusServerPortNumber, $CsvPath, $UpdateCategoryIds){
    class UpdateCategoryItem{
        [String] $ParentId
        [String] $Path

        $ApplicationsCount
        $ApplicationsLatestUpdate
        $ApplicationsOldestUpdate
        $CriticalUpdatesCount
        $CriticalUpdatesLatestUpdate
        $CriticalUpdatesOldestUpdate
        $DefinitionUpdatesCount
        $DefinitionUpdatesLatestUpdate
        $DefinitionUpdatesOldestUpdate
        $DriverSetsCount
        $DriverSetsLatestUpdate
        $DriverSetsOldestUpdate
        $DriversCount
        $DriversLatestUpdate
        $DriversOldestUpdate
        $FeaturePacksCount
        $FeaturePacksLatestUpdate
        $FeaturePacksOldestUpdate
        $SecurityUpdatesCount
        $SecurityUpdatesLatestUpdate
        $SecurityUpdatesOldestUpdate
        $ServicePacksCount
        $ServicePacksLatestUpdate
        $ServicePacksOldestUpdate
        $ToolsCount
        $ToolsLatestUpdate
        $ToolsOldestUpdate
        $UpdateRollupsCount
        $UpdateRollupsLatestUpdate
        $UpdateRollupsOldestUpdate
        $UpdatesCount
        $UpdatesLatestUpdate
        $UpdatesOldestUpdate
        $UpgradesCount
        $UpgradesLatestUpdate
        $UpgradesOldestUpdate

        hidden $Category
        hidden $EnglishCategory

        UpdateCategoryItem($Category, $EnglishCategory,[String] $ParentId,[String] $Path){
            $this.Category = $Category
            $this.EnglishCategory = $EnglishCategory
            $this.ParentId = $ParentId
            $this.Path = $Path

            $this.PSObject.Properties.Add((New-Object PSScriptProperty 'Title', {$this.Category.Title}))
            $this.PSObject.Properties.Add((New-Object PSScriptProperty 'EnglishTitle', {$this.EnglishCategory.Title}))
            $this.PSObject.Properties.Add((New-Object PSScriptProperty 'Id', {[String]$this.Category.Id}))
        }
    }
    class UpdateClassificationItem{
        $Title
        $UpdateClassification
        UpdateClassificationItem($Title, $UpdateClassification){
            $this.Title = $Title
            $this.UpdateClassification = $UpdateClassification
        }
    }

    Function Get-WsusProductItemOfWsusCategory($Category, $Classification, $Item, $RetryCount){
        $RetryCount++
        If ($RetryCount -eq 1){
            Write-Host "$(Get-Date -F F): $($Item.Title) ($($Classification.UpdateClassification.Title))"
        }
        ElseIf ($RetryCount -gt $MaximumRetry){
            Write-Host "$(Get-Date -F F): $($Item.Title) ($($Classification.UpdateClassification.Title)) の更新プログラム取得はスキップされました`n"
            Return $Item
        }
        Else{
            Write-Host "$(Get-Date -F F): $($Item.Title) ($($Classification.UpdateClassification.Title)) の更新プログラム取得を再試行します ($RetryCount 回目)`n"
        }

        Try {
            $CurrentUpdateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
            $CurrentUpdateScope.Categories.Add($Category) | Out-Null
            $CurrentUpdateScope.Classifications.Add($Classification.UpdateClassification) | Out-Null
            $Item."$($Classification.Title)Count" = $WsusServer.GetUpdateCount($CurrentUpdateScope)

            If ($Item."$($Classification.Title)Count" -ne 0){
                ForEach($Month in $Months){
                    $CurrentUpdateScope.FromCreationDate = $Month
                    $CurrentUpdateScope.ToCreationDate = $Month.AddMonths(1).AddSeconds(-1)

                    If ($WsusServer.GetUpdateCount($CurrentUpdateScope) -ne 0){
                        $Item."$($Classification.Title)OldestUpdate" = (($WsusServer.GetUpdates($CurrentUpdateScope)) | Sort-Object CreationDate)[0]
                        Break
                    }
                }
                ForEach($Month in ($Months | Sort-Object -Descending)){
                    $CurrentUpdateScope.FromCreationDate = $Month
                    $CurrentUpdateScope.ToCreationDate = $Month.AddMonths(1).AddSeconds(-1)

                    If ($WsusServer.GetUpdateCount($CurrentUpdateScope) -ne 0){
                        $Item."$($Classification.Title)LatestUpdate" = (($WsusServer.GetUpdates($CurrentUpdateScope)) | Sort-Object CreationDate -Descending)[0]
                        Break
                    }
                }
            }
            Return $Item
        }
        Catch [System.Exception]
        {
            Write-Host "$(Get-Date -F F): $($Item.Title) ($($Classification.UpdateClassification.Title)) の更新プログラムを取得できませんでした"
            Write-Host $_.Exception
            Write-Host ""

            Get-WsusProductItemOfWsusCategory $Category $Classification $Item $RetryCount
        }
    }
    
    $WsusServer = Get-WsusServer $WsusServerName -PortNumber $WsusServerPortNumber
    $WsusServerEnglish = Get-WsusServer $WsusServerName -port $WsusServerPortNumber
    $WsusServerEnglish.PreferredCulture = "en"

    $MicrosoftRootUpdateCategory = $WsusServer.GetRootUpdateCategories() | Where-Object {$_.Id -eq "56309036-4c77-4dd9-951a-99ee9c246a94"}
    $MicrosoftRootUpdateCategoryTitle = $MicrosoftRootUpdateCategory.Title

    $UpdateCategoryList = New-Object 'System.Collections.Generic.List[UpdateCategoryItem]'

    $MicrosoftRootUpdateCategory.GetSubcategories() | ForEach-Object {
        $SubcategoryTitle = $_.Title
        $SubcategoryId = $_.Id
        $_.GetSubcategories() | ForEach-Object {
            $UpdateCategoryList.Add((New-Object UpdateCategoryItem -ArgumentList $_, ($WsusServerEnglish.GetUpdateCategory($_.Id)), $SubcategoryId, "Microsoft\$SubcategoryTitle"))
        }
    }

    If ($UpdateCategoryIds -ne $Null){
        $UpdateCategoryList = $UpdateCategoryList | Where-Object {$_.ParentId -in $UpdateCategoryIds}
    }

    $UpdateClassificationList = New-Object 'System.Collections.Generic.List[UpdateClassificationItem]'
    $UpdateClassificationList.Add((New-Object UpdateClassificationItem -ArgumentList "Applications", $WsusServer.GetUpdateClassification("5c9376ab-8ce6-464a-b136-22113dd69801")))
    $UpdateClassificationList.Add((New-Object UpdateClassificationItem -ArgumentList "CriticalUpdates", $WsusServer.GetUpdateClassification("e6cf1350-c01b-414d-a61f-263d14d133b4")))
    $UpdateClassificationList.Add((New-Object UpdateClassificationItem -ArgumentList "DefinitionUpdates", $WsusServer.GetUpdateClassification("e0789628-ce08-4437-be74-2495b842f43b")))
    $UpdateClassificationList.Add((New-Object UpdateClassificationItem -ArgumentList "DriverSets", $WsusServer.GetUpdateClassification("77835c8d-62a7-41f5-82ad-f28d1af1e3b1")))
    $UpdateClassificationList.Add((New-Object UpdateClassificationItem -ArgumentList "Drivers", $WsusServer.GetUpdateClassification("ebfc1fc5-71a4-4f7b-9aca-3b9a503104a0")))
    $UpdateClassificationList.Add((New-Object UpdateClassificationItem -ArgumentList "FeaturePacks", $WsusServer.GetUpdateClassification("b54e7d24-7add-428f-8b75-90a396fa584f")))
    $UpdateClassificationList.Add((New-Object UpdateClassificationItem -ArgumentList "SecurityUpdates", $WsusServer.GetUpdateClassification("0fa1201d-4330-4fa8-8ae9-b877473b6441")))
    $UpdateClassificationList.Add((New-Object UpdateClassificationItem -ArgumentList "ServicePacks", $WsusServer.GetUpdateClassification("68c5b0a3-d1a6-4553-ae49-01d3a7827828")))
    $UpdateClassificationList.Add((New-Object UpdateClassificationItem -ArgumentList "Tools", $WsusServer.GetUpdateClassification("b4832bd8-e735-4761-8daf-37f882276dab")))
    $UpdateClassificationList.Add((New-Object UpdateClassificationItem -ArgumentList "UpdateRollups", $WsusServer.GetUpdateClassification("28bc880e-0592-4cbf-8f95-c79b17911d5f")))
    $UpdateClassificationList.Add((New-Object UpdateClassificationItem -ArgumentList "Updates", $WsusServer.GetUpdateClassification("cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83")))
    $UpdateClassificationList.Add((New-Object UpdateClassificationItem -ArgumentList "Upgrades", $WsusServer.GetUpdateClassification("3689bdc8-b205-4af4-8d4a-a63924c5e9d5")))

    # 一番古い更新プログラムの日付を指定する
    $MinimumDate = [datetime]"2003/1/1"
    $MaximumRetry = 2

    $CurrentDate = $MinimumDate
    $Months = @()
    Do{
        $Months += $CurrentDate
        $CurrentDate = $CurrentDate.AddMonths(1)
    }
    While($CurrentDate -lt (Get-Date -Day 1))

    Write-Host "WSUSサーバー上の更新プログラム ($($WsusServer.GetUpdateCount()) 件) の製品一覧を取得します"

    $UpdateCategoryList | ForEach-Object {
        $Script:Item = $_
        $Category = $_.Category
        $UpdateClassificationList | ForEach-Object {
            $Script:Item = Get-WsusProductItemOfWsusCategory -Category $Category -Classification $_ -Item $Item -RetryCount 0
        }
        $Script:Item | Select-Object Title, EnglishTitle, Path, Count, Id, ParentId, 
            @{Name="Applications.Count";Expression={$Item.ApplicationsCount}}, @{Name="ApplicationsLatestUpdate.Title";Expression={$Item.ApplicationsLatestUpdate.Title}}, @{Name="ApplicationsLatestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.ApplicationsLatestUpdate.Id).Title}}, @{Name="ApplicationsLatestUpdate.CreationDate";Expression={$Item.ApplicationsLatestUpdate.CreationDate}}, @{Name="ApplicationsLatestUpdate.LegacyName";Expression={$Item.ApplicationsLatestUpdate.LegacyName}}, @{Name="ApplicationsOldestUpdate.Title";Expression={$Item.ApplicationsOldestUpdate.Title}}, @{Name="ApplicationsOldestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.ApplicationsOldestUpdate.Id).Title}}, @{Name="ApplicationsOldestUpdate.CreationDate";Expression={$Item.ApplicationsOldestUpdate.CreationDate}}, @{Name="ApplicationsOldestUpdate.LegacyName";Expression={$Item.ApplicationsOldestUpdate.LegacyName}},
            @{Name="CriticalUpdates.Count";Expression={$Item.CriticalUpdatesCount}}, @{Name="CriticalUpdatesLatestUpdate.Title";Expression={$Item.CriticalUpdatesLatestUpdate.Title}}, @{Name="CriticalUpdatesLatestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.CriticalUpdatesLatestUpdate.Id).Title}}, @{Name="CriticalUpdatesLatestUpdate.CreationDate";Expression={$Item.CriticalUpdatesLatestUpdate.CreationDate}}, @{Name="CriticalUpdatesLatestUpdate.LegacyName";Expression={$Item.CriticalUpdatesLatestUpdate.LegacyName}}, @{Name="CriticalUpdatesOldestUpdate.Title";Expression={$Item.CriticalUpdatesOldestUpdate.Title}}, @{Name="CriticalUpdatesOldestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.CriticalUpdatesOldestUpdate.Id).Title}}, @{Name="CriticalUpdatesOldestUpdate.CreationDate";Expression={$Item.CriticalUpdatesOldestUpdate.CreationDate}}, @{Name="CriticalUpdatesOldestUpdate.LegacyName";Expression={$Item.CriticalUpdatesOldestUpdate.LegacyName}},
            @{Name="DefinitionUpdates.Count";Expression={$Item.DefinitionUpdatesCount}}, @{Name="DefinitionUpdatesLatestUpdate.Title";Expression={$Item.DefinitionUpdatesLatestUpdate.Title}}, @{Name="DefinitionUpdatesLatestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.DefinitionUpdatesLatestUpdate.Id).Title}}, @{Name="DefinitionUpdatesLatestUpdate.CreationDate";Expression={$Item.DefinitionUpdatesLatestUpdate.CreationDate}}, @{Name="DefinitionUpdatesLatestUpdate.LegacyName";Expression={$Item.DefinitionUpdatesLatestUpdate.LegacyName}}, @{Name="DefinitionUpdatesOldestUpdate.Title";Expression={$Item.DefinitionUpdatesOldestUpdate.Title}}, @{Name="DefinitionUpdatesOldestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.DefinitionUpdatesOldestUpdate.Id).Title}}, @{Name="DefinitionUpdatesOldestUpdate.CreationDate";Expression={$Item.DefinitionUpdatesOldestUpdate.CreationDate}}, @{Name="DefinitionUpdatesOldestUpdate.LegacyName";Expression={$Item.DefinitionUpdatesOldestUpdate.LegacyName}},
            @{Name="DriverSets.Count";Expression={$Item.DriverSetsCount}}, @{Name="DriverSetsLatestUpdate.Title";Expression={$Item.DriverSetsLatestUpdate.Title}}, @{Name="DriverSetsLatestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.DriverSetsLatestUpdate.Id).Title}}, @{Name="DriverSetsLatestUpdate.CreationDate";Expression={$Item.DriverSetsLatestUpdate.CreationDate}}, @{Name="DriverSetsLatestUpdate.LegacyName";Expression={$Item.DriverSetsLatestUpdate.LegacyName}}, @{Name="DriverSetsOldestUpdate.Title";Expression={$Item.DriverSetsOldestUpdate.Title}}, @{Name="DriverSetsOldestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.DriverSetsOldestUpdate.Id).Title}}, @{Name="DriverSetsOldestUpdate.CreationDate";Expression={$Item.DriverSetsOldestUpdate.CreationDate}}, @{Name="DriverSetsOldestUpdate.LegacyName";Expression={$Item.DriverSetsOldestUpdate.LegacyName}},
            @{Name="Drivers.Count";Expression={$Item.DriversCount}}, @{Name="DriversLatestUpdate.Title";Expression={$Item.DriversLatestUpdate.Title}}, @{Name="DriversLatestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.DriversLatestUpdate.Id).Title}}, @{Name="DriversLatestUpdate.CreationDate";Expression={$Item.DriversLatestUpdate.CreationDate}}, @{Name="DriversLatestUpdate.LegacyName";Expression={$Item.DriversLatestUpdate.LegacyName}}, @{Name="DriversOldestUpdate.Title";Expression={$Item.DriversOldestUpdate.Title}}, @{Name="DriversOldestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.DriversOldestUpdate.Id).Title}}, @{Name="DriversOldestUpdate.CreationDate";Expression={$Item.DriversOldestUpdate.CreationDate}}, @{Name="DriversOldestUpdate.LegacyName";Expression={$Item.DriversOldestUpdate.LegacyName}},
            @{Name="FeaturePacks.Count";Expression={$Item.FeaturePacksCount}}, @{Name="FeaturePacksLatestUpdate.Title";Expression={$Item.FeaturePacksLatestUpdate.Title}}, @{Name="FeaturePacksLatestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.FeaturePacksLatestUpdate.Id).Title}}, @{Name="FeaturePacksLatestUpdate.CreationDate";Expression={$Item.FeaturePacksLatestUpdate.CreationDate}}, @{Name="FeaturePacksLatestUpdate.LegacyName";Expression={$Item.FeaturePacksLatestUpdate.LegacyName}}, @{Name="FeaturePacksOldestUpdate.Title";Expression={$Item.FeaturePacksOldestUpdate.Title}}, @{Name="FeaturePacksOldestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.FeaturePacksOldestUpdate.Id).Title}}, @{Name="FeaturePacksOldestUpdate.CreationDate";Expression={$Item.FeaturePacksOldestUpdate.CreationDate}}, @{Name="FeaturePacksOldestUpdate.LegacyName";Expression={$Item.FeaturePacksOldestUpdate.LegacyName}},
            @{Name="SecurityUpdates.Count";Expression={$Item.SecurityUpdatesCount}}, @{Name="SecurityUpdatesLatestUpdate.Title";Expression={$Item.SecurityUpdatesLatestUpdate.Title}}, @{Name="SecurityUpdatesLatestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.SecurityUpdatesLatestUpdate.Id).Title}}, @{Name="SecurityUpdatesLatestUpdate.CreationDate";Expression={$Item.SecurityUpdatesLatestUpdate.CreationDate}}, @{Name="SecurityUpdatesLatestUpdate.LegacyName";Expression={$Item.SecurityUpdatesLatestUpdate.LegacyName}}, @{Name="SecurityUpdatesOldestUpdate.Title";Expression={$Item.SecurityUpdatesOldestUpdate.Title}}, @{Name="SecurityUpdatesOldestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.SecurityUpdatesOldestUpdate.Id).Title}}, @{Name="SecurityUpdatesOldestUpdate.CreationDate";Expression={$Item.SecurityUpdatesOldestUpdate.CreationDate}}, @{Name="SecurityUpdatesOldestUpdate.LegacyName";Expression={$Item.SecurityUpdatesOldestUpdate.LegacyName}},
            @{Name="ServicePacks.Count";Expression={$Item.ServicePacksCount}}, @{Name="ServicePacksLatestUpdate.Title";Expression={$Item.ServicePacksLatestUpdate.Title}}, @{Name="ServicePacksLatestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.ServicePacksLatestUpdate.Id).Title}}, @{Name="ServicePacksLatestUpdate.CreationDate";Expression={$Item.ServicePacksLatestUpdate.CreationDate}}, @{Name="ServicePacksLatestUpdate.LegacyName";Expression={$Item.ServicePacksLatestUpdate.LegacyName}}, @{Name="ServicePacksOldestUpdate.Title";Expression={$Item.ServicePacksOldestUpdate.Title}}, @{Name="ServicePacksOldestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.ServicePacksOldestUpdate.Id).Title}}, @{Name="ServicePacksOldestUpdate.CreationDate";Expression={$Item.ServicePacksOldestUpdate.CreationDate}}, @{Name="ServicePacksOldestUpdate.LegacyName";Expression={$Item.ServicePacksOldestUpdate.LegacyName}},
            @{Name="Tools.Count";Expression={$Item.ToolsCount}}, @{Name="ToolsLatestUpdate.Title";Expression={$Item.ToolsLatestUpdate.Title}}, @{Name="ToolsLatestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.ToolsLatestUpdate.Id).Title}}, @{Name="ToolsLatestUpdate.CreationDate";Expression={$Item.ToolsLatestUpdate.CreationDate}}, @{Name="ToolsLatestUpdate.LegacyName";Expression={$Item.ToolsLatestUpdate.LegacyName}}, @{Name="ToolsOldestUpdate.Title";Expression={$Item.ToolsOldestUpdate.Title}}, @{Name="ToolsOldestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.ToolsOldestUpdate.Id).Title}}, @{Name="ToolsOldestUpdate.CreationDate";Expression={$Item.ToolsOldestUpdate.CreationDate}}, @{Name="ToolsOldestUpdate.LegacyName";Expression={$Item.ToolsOldestUpdate.LegacyName}},
            @{Name="UpdateRollups.Count";Expression={$Item.UpdateRollupsCount}}, @{Name="UpdateRollupsLatestUpdate.Title";Expression={$Item.UpdateRollupsLatestUpdate.Title}}, @{Name="UpdateRollupsLatestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.UpdateRollupsLatestUpdate.Id).Title}}, @{Name="UpdateRollupsLatestUpdate.CreationDate";Expression={$Item.UpdateRollupsLatestUpdate.CreationDate}}, @{Name="UpdateRollupsLatestUpdate.LegacyName";Expression={$Item.UpdateRollupsLatestUpdate.LegacyName}}, @{Name="UpdateRollupsOldestUpdate.Title";Expression={$Item.UpdateRollupsOldestUpdate.Title}}, @{Name="UpdateRollupsOldestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.UpdateRollupsOldestUpdate.Id).Title}}, @{Name="UpdateRollupsOldestUpdate.CreationDate";Expression={$Item.UpdateRollupsOldestUpdate.CreationDate}}, @{Name="UpdateRollupsOldestUpdate.LegacyName";Expression={$Item.UpdateRollupsOldestUpdate.LegacyName}},
            @{Name="Updates.Count";Expression={$Item.UpdatesCount}}, @{Name="UpdatesLatestUpdate.Title";Expression={$Item.UpdatesLatestUpdate.Title}}, @{Name="UpdatesLatestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.UpdatesLatestUpdate.Id).Title}}, @{Name="UpdatesLatestUpdate.CreationDate";Expression={$Item.UpdatesLatestUpdate.CreationDate}}, @{Name="UpdatesLatestUpdate.LegacyName";Expression={$Item.UpdatesLatestUpdate.LegacyName}}, @{Name="UpdatesOldestUpdate.Title";Expression={$Item.UpdatesOldestUpdate.Title}}, @{Name="UpdatesOldestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.UpdatesOldestUpdate.Id).Title}}, @{Name="UpdatesOldestUpdate.CreationDate";Expression={$Item.UpdatesOldestUpdate.CreationDate}}, @{Name="UpdatesOldestUpdate.LegacyName";Expression={$Item.UpdatesOldestUpdate.LegacyName}},
            @{Name="Upgrades.Count";Expression={$Item.UpgradesCount}}, @{Name="UpgradesLatestUpdate.Title";Expression={$Item.UpgradesLatestUpdate.Title}}, @{Name="UpgradesLatestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.UpgradesLatestUpdate.Id).Title}}, @{Name="UpgradesLatestUpdate.CreationDate";Expression={$Item.UpgradesLatestUpdate.CreationDate}}, @{Name="UpgradesLatestUpdate.LegacyName";Expression={$Item.UpgradesLatestUpdate.LegacyName}}, @{Name="UpgradesOldestUpdate.Title";Expression={$Item.UpgradesOldestUpdate.Title}}, @{Name="UpgradesOldestUpdate.EnglishTitle";Expression={$WsusServerEnglish.GetUpdate($Item.UpgradesOldestUpdate.Id).Title}}, @{Name="UpgradesOldestUpdate.CreationDate";Expression={$Item.UpgradesOldestUpdate.CreationDate}}, @{Name="UpgradesOldestUpdate.LegacyName";Expression={$Item.UpgradesOldestUpdate.LegacyName}} |
        Export-Csv -Path $CsvPath -Encoding UTF8 -NoTypeInformation -Append
    }
}

$CsvPath = "$env:USERPROFILE\Desktop\Get-WsusLatestUpdateAndOldestUpdatePerProduct.csv"

Get-WsusLatestUpdateAndOldestUpdatePerProduct -WsusServerName "localhost" -WsusServerPortNumber 8530 -CsvPath $CsvPath

# 下記のカテゴリー配下の更新プログラムのみを抽出する
# - Windows
# - Office
# - SQL Server
# - Developer Tools, Runtimes, and Redistributables
# - PowerShell
# Get-WsusLatestUpdateAndOldestUpdatePerProduct -WsusServerName "localhost" -WsusServerPortNumber 8530 -CsvPath $CsvPath -UpdateCategoryIds @("6964aab4-c5b5-43bd-a17d-ffb4346a8e1d", "477b856e-65c4-4473-b621-a8b230bb70d9", "48ce8c86-6850-4f68-8e9d-7dc8535ced60", "0a4c6c73-8887-4d7f-9cbe-d08fa8fa9d1e", "fca452fc-9a11-ca4a-bf3e-0a0cef82a80e")
